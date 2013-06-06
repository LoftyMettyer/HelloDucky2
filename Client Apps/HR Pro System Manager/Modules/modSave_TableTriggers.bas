Attribute VB_Name = "modSave_TableTriggers"
Option Explicit

Public Const VARCHARTHRESHOLD = 500

Private miAccordDefaultStatus As Integer
Private miAccordStatusForUtilities As Integer
Private mbAccordAllowDelete As Boolean

Private mstrGetFieldAutoUpdateCode_INSERT As String
Private mstrGetFieldAutoUpdateCode_UPDATE As String
Private mstrGetFieldAutoUpdateCode_DELETE As String

Private sDeclareInsCols As SystemMgr.cStringBuilder
Private sDeclareDelCols As SystemMgr.cStringBuilder
Private sFetchInsCols As SystemMgr.cStringBuilder
Private sFetchDelCols As SystemMgr.cStringBuilder
Private sSelectInsCols As SystemMgr.cStringBuilder
Private sSelectDelCols As SystemMgr.cStringBuilder

Private sConvertInsCols As String
Private sConvertDelCols As String
Private alngAuditColumns() As Long
Private sExprDeclarationCode As SystemMgr.cStringBuilder

Private sInsertSpecialFunctionsCode As String
Private sUpdateSpecialFunctionsCode1 As String
Private sUpdateSpecialFunctionsCode2 As String
Private sDeleteSpecialFunctionsCode As String
Private sCalcDfltCode As SystemMgr.cStringBuilder

Private sDfltExprDeclarationCode As SystemMgr.cStringBuilder

Private sInsertAuditCode As SystemMgr.cStringBuilder
Private sUpdateAuditCode As SystemMgr.cStringBuilder
Private sDeleteAuditCode As SystemMgr.cStringBuilder

Private sInsertWorkflowCode As SystemMgr.cStringBuilder
Private sUpdateWorkflowCode As SystemMgr.cStringBuilder
Private sDeleteWorkflowCode As SystemMgr.cStringBuilder

Private sDateDependentUpdateCode As SystemMgr.cStringBuilder
Private sRelationshipCode As SystemMgr.cStringBuilder

Public Function SetTriggers(palngExpressions() As Long, pfRefreshDatabase As Boolean) As Boolean
  ' Set the triggers that are required for calculated columns, audited columns
  ' and relationships.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim lngTableID As Long
  Dim lngPersonnelTableID As Long
  Dim lngAbsenceTableID As Long
  Dim strDependantsTableName As String
  Dim rsExistingTriggers As New ADODB.Recordset
  Dim strTableName As String
  Dim lngRecordCount As Long
  
  fOK = True
  
  lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
  lngAbsenceTableID = GetModuleSetting(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE, 0)
  strDependantsTableName = GetTableName(GetModuleSetting(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE, 0))
  
  ' Create the triggers for each table.
  With recTabEdit
    .Index = "idxTableID"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      lngRecordCount = .RecordCount
    End If
    
    OutputCurrentProcess2 vbNullString, lngRecordCount
    
    Do While fOK And Not .EOF
      lngTableID = !TableID
     
      If Not !Deleted Then
        ' Create the triggers.
        strTableName = "tbuser_" + .Fields("TableName").value

        OutputCurrentProcess2 .Fields("TableName").value, lngRecordCount
        gobjProgress.UpdateProgress2

        fOK = SetTableTriggers_GetStrings(lngTableID, _
          strTableName, _
          IIf((IsNull(!RecordDescExprID)) Or (!RecordDescExprID < 0), 0, !RecordDescExprID), _
          palngExpressions, pfRefreshDatabase)

        ' Johnny told me to do this... Was getting a bit confused with the record selectors in workflow.
        .Index = "idxTableID"
        .Seek "=", lngTableID

        If fOK Then
          fOK = SetTableTriggers_CreateTriggers(lngTableID, _
            strTableName, _
            IIf((IsNull(!RecordDescExprID)) Or (!RecordDescExprID < 0), 0, !RecordDescExprID), _
            lngPersonnelTableID, (lngTableID = lngAbsenceTableID), strDependantsTableName)
        End If

      End If
           
      If fOK Then
        ' Reposition the recTabEdit pointer as it may have been
        ' moved in other methods.
        .Index = "idxTableID"
        .Seek "=", lngTableID
        .MoveNext
      Else
        Exit Do
      End If


      sSQL = "delete from asrsysemailqueue " & _
             "where linkid not in (select linkid from asrsysemaillinks)" & _
             " and not(columnid is null)"
      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

      fOK = fOK And Not gobjProgress.Cancelled

    Loop
  End With
    
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsExistingTriggers = Nothing
  SetTriggers = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating triggers"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function SetTableTriggers_GetStrings(pLngCurrentTableID As Long, _
  psTableName As String, _
  plngRecDescExprID As Long, _
  ByRef palngExpressions() As Long, _
  pfRefreshDatabase As Boolean) As Boolean
  
  ' Create the triggers for the given table (pLngCurrentTableID).
  ' The DELETE triggers handle Relationships, Calculated Columns and Audit trails.
  ' The INSERT and UPDATE triggers handle just Calculated Columns and Audit trails.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim fExprDone As Boolean
  Dim fExprIsDateDependent As Boolean
  Dim iLoop As Long
  Dim iLoop2 As Long
  Dim iLoop3 As Long
  Dim iIndex As Long
  Dim iIndex2 As Long
  Dim iArrayIndex As Long
  Dim iCalcRelationship As Long
  Dim strColumnID As String
  Dim strColumnName As String
  Dim iControlType As Long
  Dim lngExprID As Long
  Dim lngCalcTableID As Long
  Dim lngCalcColumnID As Long
  'Dim lngChildTableID As Long
  Dim lngLastExprID As Long
  Dim dblMaxValue As Double
  Dim sSQL As String
  Dim sExprName As String
  Dim sDfltValue As String
  Dim sIndent As String
  Dim sAuditCode As String
  Dim sCalcTable As String
  Dim sCalcColumn As String
  Dim iCalcDataType As DataTypes
  Dim iCalcSize As Long
  Dim sIfNullCode As String
  Dim sConvertCode As String
  Dim sCursorName As String
  Dim sExtraSetCode As String
  Dim rsParents As ADODB.Recordset
'  Dim rsChildren As ADODB.Recordset
  Dim rsCalcColumns As ADODB.Recordset
  Dim rsExpressions As ADODB.Recordset
  Dim avExpressions() As Long
  Dim iNextIndex As Long
  Dim sDfltColumn As String
  Dim sDfltColumnID As String
  Dim sDfltOldVar As String
  Dim sDefaultDeclareCode As String
  Dim sDefaultConvertCode As String
  Dim sDefaultIfNullCode As String
  Dim rsDfltColumns As ADODB.Recordset

  Dim alngParents() As Long
  Dim blnBuildDiarySP As Boolean


  Dim alngColumns() As Long
  ReDim alngColumns(0)
  
  'Dim sTemp As String
    
  Dim lngColumnID As Long
    
  ' JPD20020913 - instead of making multiple queries to the triggered table, and
  ' the 'inserted' and 'deleted' tables, we now get all of the required information in
  ' the cursor that we used to loop through to get just the id of each record being
  ' inserted/updated/deleted.
  ' This change was driven by the performance degradation reported by
  ' Islington.
  Dim fColFound As Boolean

  'TM20061010 - Fault 11516
  Dim sAUSQL As String
  Dim rsAULookupColumns As New ADODB.Recordset
  
  'NPG20080415 Sugg S000441
  Dim blnCalculateIfEmpty As Boolean
  
  Dim bIsMaxSize As Boolean
  
  Set sDateDependentUpdateCode = New SystemMgr.cStringBuilder
  Set sExprDeclarationCode = New SystemMgr.cStringBuilder
  
  Set rsParents = New ADODB.Recordset
  'Set rsChildren = New ADODB.Recordset
  Set rsCalcColumns = New ADODB.Recordset
  Set rsExpressions = New ADODB.Recordset
  Set sCalcDfltCode = New SystemMgr.cStringBuilder
  Set sDfltExprDeclarationCode = New SystemMgr.cStringBuilder
  Set rsDfltColumns = New ADODB.Recordset
  
  Set sDeclareInsCols = New SystemMgr.cStringBuilder
  Set sDeclareDelCols = New SystemMgr.cStringBuilder
  Set sFetchInsCols = New SystemMgr.cStringBuilder
  Set sFetchDelCols = New SystemMgr.cStringBuilder
  Set sSelectInsCols = New SystemMgr.cStringBuilder
'  Set sSelectInsCols2 = New SystemMgr.cStringBuilder
  Set sSelectDelCols = New SystemMgr.cStringBuilder
  
  sDeclareInsCols.TheString = "    DECLARE @sTempInsCol varchar(MAX)"
  sDeclareDelCols.TheString = "    DECLARE @sTempDelCol varchar(MAX)"
  sSelectInsCols.TheString = vbNullString
 ' sSelectInsCols2.TheString = vbNullString
  sSelectDelCols.TheString = vbNullString
  sFetchInsCols.TheString = vbNullString
  sFetchDelCols.TheString = vbNullString
 
  'Reset the modular level GetField - AutoUpdate trigger strings
  mstrGetFieldAutoUpdateCode_DELETE = vbNullString
  mstrGetFieldAutoUpdateCode_UPDATE = vbNullString
  mstrGetFieldAutoUpdateCode_INSERT = vbNullString
  
  ' Calculation Relationship constants.
  Const giCALCULATE_UNKNOWN = 0
  Const giCALCULATE_PARENT = 1
  Const giCALCULATE_SELF = 2
  Const giCALCULATE_CHILD = 3

  fOK = True
  ReDim alngAuditColumns(0)
    
  '
  ' Create the trigger code required for handling the Auditing of columns in the current table.
  '
  Set sInsertAuditCode = New SystemMgr.cStringBuilder
  Set sUpdateAuditCode = New SystemMgr.cStringBuilder
  Set sDeleteAuditCode = New SystemMgr.cStringBuilder
  
  Set sInsertWorkflowCode = New SystemMgr.cStringBuilder
  Set sUpdateWorkflowCode = New SystemMgr.cStringBuilder
  Set sDeleteWorkflowCode = New SystemMgr.cStringBuilder
  
  sInsertWorkflowCode.TheString = vbNullString
  sUpdateWorkflowCode.TheString = vbNullString
  sDeleteWorkflowCode.TheString = vbNullString
   
  sInsertAuditCode.TheString = vbNullString
  sUpdateAuditCode.TheString = vbNullString
  sDeleteAuditCode.TheString = vbNullString
  
  ' Pointer will already be on the correct table (unless someone starts messing around in here...)
  With recTabEdit
    
    ' Table level audit insert stuff
    If .Fields("AuditInsert").value = True Then
      sInsertAuditCode.Append vbNewLine & vbTab & "/* Table level audit */" & _
        vbNewLine & vbTab & "EXECUTE dbo.spstat_audittable " & pLngCurrentTableID & ", @recordID, @recordDesc, '* New Record *'" & vbNewLine
    Else
      sInsertAuditCode.TheString = vbNullString
    End If
     
    ' Table level audit delete stuff
    If .Fields("AuditDelete").value = True Then
      sDeleteAuditCode.Append vbNewLine & vbTab & "/* Table level audit */" & _
        vbNewLine & vbTab & "EXECUTE dbo.spstat_audittable " & pLngCurrentTableID & ", @recordID, @recordDesc, '* Deleted Record *'" & vbNewLine
    Else
      sDeleteAuditCode.TheString = vbNullString
    End If
     
    ' Record based workflow links
    sInsertWorkflowCode.Append WorkflowTableTriggerCode(pLngCurrentTableID, WFRELATEDRECORD_INSERT)
    sUpdateWorkflowCode.Append WorkflowTableTriggerCode(pLngCurrentTableID, WFRELATEDRECORD_UPDATE)
    sDeleteWorkflowCode.Append WorkflowTableTriggerCode(pLngCurrentTableID, WFRELATEDRECORD_DELETE)
  End With
    
  
  
  
  ' Payroll Transfer Triggers
  ' --------------------------------
  If gbAccordPayrollModule Then
    
    Set sUpdateAccordCode = New SystemMgr.cStringBuilder
    Set sDeleteAccordCode = New SystemMgr.cStringBuilder

    ' Is in a separate sub routine because this one is getting too big for VB to compile.
    ' All parameters passed by reference!
    SetTableTriggers_AccordTransfer sUpdateAccordCode, sDeleteAccordCode, alngAuditColumns(), _
      sSelectInsCols, sSelectDelCols, _
      sFetchInsCols, sFetchDelCols, _
      sDeclareInsCols, sDeclareDelCols, _
      pLngCurrentTableID
      
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", pLngCurrentTableID
  End If

  'JPD 20050131 Fault 8820
  ' Special Function Triggers
  ' --------------------------------
  ' Is in a separate sub routine because this one is getting too big for VB to compile.
  ' All parameters passed by reference!
  
  SetTableTriggers_SpecialFunctions _
    alngAuditColumns(), _
    sInsertSpecialFunctionsCode, _
    sUpdateSpecialFunctionsCode1, _
    sUpdateSpecialFunctionsCode2, _
    sDeleteSpecialFunctionsCode, _
    pLngCurrentTableID

  recTabEdit.Index = "idxTableID"
  recTabEdit.Seek "=", pLngCurrentTableID



  GetTriggerRelationshipCode pLngCurrentTableID

 
  'MH19991110
  'This creates one stored procedure which can be called
  'for any/all of the three triggers
  If fOK Then
    
    'MH20020213 Need to build diary SP for all new tables...
    blnBuildDiarySP = (Application.ChangedDiaryLink _
                    Or Application.ChangedTableName _
                    Or Application.ChangedColumnName _
                    Or pfRefreshDatabase _
                    Or gfRefreshStoredProcedures)
    If Not blnBuildDiarySP Then
      With recTabEdit
        .Index = "idxTableID"
        If Not (.BOF And .EOF) Then
          .MoveFirst
        End If
        Do While Not .EOF
          
          If !TableID = pLngCurrentTableID Then
            blnBuildDiarySP = !New
            Exit Do
          End If
          
          .MoveNext
        Loop
      End With
    End If
    
    'MH20010718 Might speed things up a little if we check if these have changed
    'NOTE: CAN'T DO THIS FOR EMAIL AS IT WORKS SLIGHTLY DIFFERENTLY !!!
    'CreateDiaryProcsForTable pLngCurrentTableID, psTableName, plngRecDescExprID
    'If Application.ChangedDiaryLink Or pfRefreshDatabase Then
    'If Application.ChangedDiaryLink Or Application.ChangedTableName Or pfRefreshDatabase Then
    If blnBuildDiarySP Then
      CreateDiaryProcsForTable pLngCurrentTableID, psTableName, plngRecDescExprID
    End If


    'MH20040331
    If Not CreateOutlookEventsForTable(pLngCurrentTableID, psTableName, plngRecDescExprID) Then
      SetTableTriggers_GetStrings = False
      Exit Function
    End If


    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are passing a number of variables and an array to the email trigger creation
    ' code so that the email columns can be added to the SELECT statement that is used
    ' to create the cursor, the FETCH statement that used to loop through the cursor,
    ' and the DECLARE statements that are needed.
    ' The email check code is modified for the new implementation.
    ' NB. an array of columns that have been added to the SELECT statement is used
    ' to ensure that columns aren't added more than once. Audit columns, email columns
    ' and calculated columns all use this method.
    ' This change was driven by the performance degradation reported by
    ' Islington.
    'CreateEmailProcsForTable pLngCurrentTableID, psTableName, plngRecDescExprID
    'CreateEmailProcsForTable pLngCurrentTableID, psTableName, plngRecDescExprID, alngAuditColumns, sDeclareInsCols, sDeclareDelCols, sSelectInsCols, sSelectDelCols, sFetchInsCols, sFetchDelCols
    CreateEmailProcsForTable pLngCurrentTableID, psTableName, plngRecDescExprID, alngAuditColumns, _
      sDeclareInsCols, sDeclareDelCols, _
      sSelectInsCols, sSelectDelCols, _
      sFetchInsCols, sFetchDelCols
  
    ' Column/date based workflow links
    CreateWorkflowProcsForTable pLngCurrentTableID, psTableName, plngRecDescExprID, alngAuditColumns, _
      sDeclareInsCols, sDeclareDelCols, _
      sSelectInsCols, sSelectDelCols, _
      sFetchInsCols, sFetchDelCols

    sInsertWorkflowCode.Append WorkflowLinkTriggerCode_Insert
    sUpdateWorkflowCode.Append WorkflowLinkTriggerCode_Update
  End If
  
  
  If fOK Then
    ' Create the default value calculation code.
    ' NB. Straight default values are embedded right in the column definitions in SQL so
    ' we only need worry about the calculated defaults here.
    
    ' Index 1 is the calc code itself.
    ' Index 2 is the update code itself.
    sCalcDfltCode.Append vbNullString
    
    sDfltExprDeclarationCode.Append "        /* --------------------------------------------------------- */" & vbNewLine & _
      "        /* Default expression declaration code. */" & vbNewLine & _
      "        /* --------------------------------------------------------- */" & vbNewLine
          
    ReDim alngParents(0)
    sSQL = "SELECT parentID" & _
      " FROM ASRSysRelations" & _
      " WHERE childID = " & Trim$(Str$(pLngCurrentTableID)) & _
      " ORDER BY parentID"
        
    rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    Do While Not rsParents.EOF
      sDfltExprDeclarationCode.Append _
        "        DECLARE @id_" & Trim(Str(rsParents.Fields(0).value)) & " integer" & vbNewLine
      iNextIndex = UBound(alngParents) + 1
      ReDim Preserve alngParents(iNextIndex)
      alngParents(iNextIndex) = rsParents.Fields(0).value
           
      sSelectInsCols.Append ", isnull(inserted.ID_" & Trim(Str(rsParents.Fields(0).value)) & ",0)"
      sSelectDelCols.Append ", isnull(deleted.ID_" & Trim(Str(rsParents.Fields(0).value)) & ",0)"
      
      sFetchInsCols.Append ", @insParentID_" & Trim(Str(rsParents.Fields(0).value))
      sFetchDelCols.Append ", @delParentID_" & Trim(Str(rsParents.Fields(0).value))
      
      sDeclareInsCols.Append ", @insParentID_" & Trim(Str(rsParents.Fields(0).value)) & " integer"
      sDeclareDelCols.Append ", @delParentID_" & Trim(Str(rsParents.Fields(0).value)) & " integer"

      rsParents.MoveNext
    Loop
    rsParents.Close
    
    sSQL = "SELECT DISTINCT ASRSysExpressions.exprID, ASRSysExpressions.returnType, " & _
      " ASRSysColumns.columnID, ASRSysColumns.columnName, ASRSysColumns.dataType, ASRSysColumns.size, " & _
      " ASRSysColumns.decimals, ASRSysColumns.MultiLine" & _
      " FROM ASRSysExpressions" & _
      " INNER JOIN ASRSysColumns ON ASRSysExpressions.exprID = ASRSysColumns.dfltValueExprID" & _
      " WHERE ASRSysColumns.tableID = " & Trim$(Str$(pLngCurrentTableID)) & _
      " ORDER BY ASRSysExpressions.exprID"
      
    rsDfltColumns.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    lngLastExprID = 0
    With rsDfltColumns
      Do While Not .EOF
        lngExprID = rsDfltColumns(0).value
        sExprName = "dfltexpr" & Trim$(Str$(lngExprID))
        sDfltColumn = rsDfltColumns(3).value
        sDfltColumnID = Trim(Str(rsDfltColumns(2).value))
    
        ' Create the data type size conversion code.
        Select Case !DataType
          Case dtVARCHAR
            If !MultiLine Then
              sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " varchar(max)" & vbNewLine
            Else
              sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " varchar(" & !Size & ")" & vbNewLine
            End If
            sDefaultIfNullCode = "SET @" & sExprName & "_" & sDfltColumnID & " = ''"
            sDefaultConvertCode = "CONVERT(varchar(max), "
            sDfltOldVar = "@oldCharValue"
            
          Case dtLONGVARCHAR
            sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " varchar(max)" & vbNewLine
            sDefaultIfNullCode = "SET @" & sExprName & "_" & sDfltColumnID & " = ''"
            sDefaultConvertCode = "CONVERT(varchar(14), "
            sDfltOldVar = "@oldCharValue"
          
          Case dtINTEGER
            sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " float" & vbNewLine
            sDefaultIfNullCode = "SET @" & sExprName & "_" & sDfltColumnID & " = 0"
            sDefaultConvertCode = "CONVERT(int, "
            sDfltOldVar = "@oldNumValue"
          
          Case dtNUMERIC
            sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " float" & vbNewLine
            sDefaultIfNullCode = "SET @" & sExprName & "_" & sDfltColumnID & " = 0"
            sDefaultConvertCode = "CONVERT(numeric(" & Trim(Str(!Size)) & ", " & Trim(Str(!Decimals)) & "), "
            sDfltOldVar = "@oldNumValue"
        
          Case dtTIMESTAMP
            sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " datetime" & vbNewLine
            sDefaultIfNullCode = "SET @" & sExprName & "_" & sDfltColumnID & " = null"
            sDefaultConvertCode = vbNullString
            sDfltOldVar = "@oldDateValue"
        
          Case dtBIT
            sDefaultDeclareCode = "        DECLARE @" & sExprName & "_" & sDfltColumnID & " bit" & vbNewLine
            sDefaultIfNullCode = "SET @" & sExprName & "_" & sDfltColumnID & " = 0"
            sDefaultConvertCode = vbNullString
            sDfltOldVar = "@oldLogicValue"
        End Select
    
        sDfltExprDeclarationCode.Append sDefaultDeclareCode
        
        sCalcDfltCode.Append vbNewLine & _
          "        SELECT " & sDfltOldVar & " = " & sDfltColumn
        
        For iNextIndex = 1 To UBound(alngParents)
            sCalcDfltCode.Append "," & vbNewLine & _
              "            @id_" & Trim$(Str$(alngParents(iNextIndex))) & " = id_" & Trim$(Str$(alngParents(iNextIndex)))
        Next iNextIndex
        
        sCalcDfltCode.Append vbNewLine & _
          "        FROM " & psTableName & vbNewLine & _
          "        WHERE id = @recordID" & vbNewLine & vbNewLine & _
          "        IF (" & sDfltOldVar & " IS NULL)" & IIf((!DataType = dtVARCHAR) Or (!DataType = dtLONGVARCHAR), " OR (len(ltrim(rtrim(" & sDfltOldVar & "))) = 0)", vbNullString) & vbNewLine & _
          "        BEGIN" & vbNewLine & _
          "            IF (EXISTS(SELECT Name FROM sysobjects WHERE type = 'P' AND name = 'sp_ASRDfltExpr_" & Trim$(Str$(lngExprID)) & "'))" & vbNewLine & _
          "            BEGIN" & vbNewLine & _
          "                EXEC @hResult = dbo.sp_ASRDfltExpr_" & Trim$(Str$(lngExprID)) & " @" & sExprName & "_" & sDfltColumnID & " OUTPUT"
          
        For iNextIndex = 1 To UBound(alngParents)
          sCalcDfltCode.Append ", @id_" & Trim$(Str$(alngParents(iNextIndex)))
        Next iNextIndex
        

        sCalcDfltCode.Append vbNewLine & _
          "                IF @hResult <> 0 " & sDefaultIfNullCode & vbNewLine & _
          "            END" & vbNewLine & _
          "            ELSE " & sDefaultIfNullCode & vbNewLine
    
        If !DataType = dtNUMERIC Then
          dblMaxValue = 10 ^ (!Size - !Decimals)
          sCalcDfltCode.Append _
            "            IF @" & sExprName & "_" & sDfltColumnID & " >= " & Trim$(Str$(dblMaxValue)) & " SET @" & sExprName & "_" & sDfltColumnID & " = 0" & vbNewLine & _
            "            IF @" & sExprName & "_" & sDfltColumnID & " <= -" & Trim$(Str$(dblMaxValue)) & " SET @" & sExprName & "_" & sDfltColumnID & " = 0" & vbNewLine
        End If
    
        sCalcDfltCode.Append _
          "            /* Update the record with the calculated default values. */" & vbNewLine & _
          "            UPDATE " & psTableName & vbNewLine & _
          "                SET " & sDfltColumn & " = " & sDefaultConvertCode & "@" & sExprName & "_" & sDfltColumnID & IIf(LenB(sDefaultConvertCode) <> 0, ")", vbNullString) & vbNewLine & _
          "                WHERE " & psTableName & ".ID = @recordID" & vbNewLine & _
          "        END" & vbNewLine

        .MoveNext
      Loop
      
      .Close
    End With
  
    ' Now do the stored procedure code for doing the straight default values.
    sSQL = "SELECT columnID, columnName, dataType, size, decimals, defaultValue, MultiLine" & _
      " FROM ASRSysColumns" & _
      " WHERE tableID = " & Trim$(Str$(pLngCurrentTableID)) & _
      " AND len(defaultValue) > 0" & _
      " AND (dataType <> " & dtTIMESTAMP & " OR defaultValue <> '__/__/____')" & _
      " ORDER BY columnID"
        
    rsDfltColumns.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    With rsDfltColumns
      Do While Not .EOF
        sDfltValue = "dfltvalue"
        sDfltColumn = rsDfltColumns(1).value
        sDfltColumnID = Trim(Str(rsDfltColumns(0).value))
    
        ' Create the data type size conversion code.
        Select Case !DataType
          Case dtVARCHAR
            If !MultiLine Then
              sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " varchar(max)" & vbNewLine
            Else
              sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " varchar(" & !Size & ")" & vbNewLine
            End If
            sDefaultConvertCode = "CONVERT(varchar(MAX), "
            sDfltOldVar = "@oldCharValue"
    
          Case dtLONGVARCHAR
            sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " varchar(max)" & vbNewLine
            sDefaultConvertCode = "CONVERT(varchar(14), "
            sDfltOldVar = "@oldCharValue"
    
          Case dtINTEGER
            sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " float" & vbNewLine
            sDefaultConvertCode = "CONVERT(int, "
            sDfltOldVar = "@oldNumValue"
    
          Case dtNUMERIC
            sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " float" & vbNewLine
            sDefaultConvertCode = "CONVERT(numeric(" & Trim(Str(!Size)) & ", " & Trim(Str(!Decimals)) & "), "
            sDfltOldVar = "@oldNumValue"
    
          Case dtTIMESTAMP
            sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " datetime" & vbNewLine
            sDefaultConvertCode = vbNullString
            sDfltOldVar = "@oldDateValue"
    
          Case dtBIT
            sDefaultDeclareCode = "        DECLARE @" & sDfltValue & "_" & sDfltColumnID & " bit" & vbNewLine
            sDefaultConvertCode = vbNullString
            sDfltOldVar = "@oldLogicValue"
        End Select
    
        sDfltExprDeclarationCode.Append sDefaultDeclareCode
    
        sCalcDfltCode.Append vbNewLine & _
          "        SELECT " & sDfltOldVar & " = " & sDfltColumn & vbNewLine & _
          "        FROM " & psTableName & vbNewLine & _
          "        WHERE id = @recordID" & vbNewLine & vbNewLine & _
          "        IF (" & sDfltOldVar & " IS NULL)" & IIf((!DataType = dtVARCHAR) Or (!DataType = dtLONGVARCHAR), " OR (len(ltrim(rtrim(" & sDfltOldVar & "))) = 0)", vbNullString) & vbNewLine & _
          "        BEGIN" & vbNewLine
    
        Select Case !DataType
          Case dtVARCHAR, dtLONGVARCHAR
            sCalcDfltCode.Append _
              "            SET @" & sDfltValue & "_" & sDfltColumnID & " = '" & Replace(!DefaultValue, "'", "''") & "'" & vbNewLine
          Case dtTIMESTAMP
            sCalcDfltCode.Append _
              "            SET @" & sDfltValue & "_" & sDfltColumnID & " = '" & Replace(!DefaultValue, "'", "''") & "'" & vbNewLine
          Case dtINTEGER
            sCalcDfltCode.Append _
              "            SET @" & sDfltValue & "_" & sDfltColumnID & " = " & Trim(Str(val(!DefaultValue))) & vbNewLine
          Case dtNUMERIC
            dblMaxValue = 10 ^ (!Size - !Decimals)
            sCalcDfltCode.Append _
              "            SET @" & sDfltValue & "_" & sDfltColumnID & " = " & Trim(Str(val(!DefaultValue))) & vbNewLine & _
              "            IF @" & sDfltValue & "_" & sDfltColumnID & " >= " & Trim$(Str$(dblMaxValue)) & " SET @" & sDfltValue & "_" & sDfltColumnID & " = 0" & vbNewLine & _
              "            IF @" & sDfltValue & "_" & sDfltColumnID & " <= -" & Trim$(Str$(dblMaxValue)) & " SET @" & sDfltValue & "_" & sDfltColumnID & " = 0" & vbNewLine
          Case dtBIT
            sCalcDfltCode.Append _
              "            SET @" & sDfltValue & "_" & sDfltColumnID & " = " & Trim(Str(val(!DefaultValue))) & vbNewLine
        End Select
    
        sCalcDfltCode.Append _
          "            /* Update the record with the calculated default values. */" & vbNewLine & _
          "            UPDATE " & psTableName & vbNewLine & _
          "                SET " & sDfltColumn & " = " & sDefaultConvertCode & "@" & sDfltValue & "_" & sDfltColumnID & IIf(LenB(sDefaultConvertCode) <> 0, ")", vbNullString) & vbNewLine & _
          "                WHERE " & psTableName & ".ID = @recordID" & vbNewLine & _
          "        END" & vbNewLine
    
        .MoveNext
      Loop
    
      .Close
    End With
    
  End If


TidyUpAndExit:
  ' Disassociate object variables.
  Set rsDfltColumns = Nothing
  'Set rsChildren = Nothing
  'Set rsParents = Nothing
  Set rsExpressions = Nothing
  Set rsCalcColumns = Nothing
  Set rsAULookupColumns = Nothing
  
  SetTableTriggers_GetStrings = fOK
  
  Exit Function

ErrorTrap:
  fOK = False
  gobjProgress.Visible = False
  OutputError "Error creating table trigger code"
  Err = False
  Resume TidyUpAndExit

End Function


Private Function SetTableTriggers_CreateTriggers(pLngCurrentTableID As Long, _
  psTableName As String, _
  plngRecDescExprID As Long, _
  lngPersonnelTableID As Long, pfIsAbsenceTable As Boolean, strDependantsTableName As String) As Boolean

  On Error GoTo ErrorTrap

  Dim objTable As SystemFramework.Table
  Dim fOK As Boolean
  Dim sSQL As String
  'Dim sGetRecordDesc As String
  Dim sCursorName As String
  Dim objExpr As CExpression
  Dim iLoop As Integer
  
  Dim miTriggerRecursionLevel As Integer
  Dim fSelfCalcs As Boolean
  Dim fParentCalcs As Boolean
  Dim fChildCalcs As Boolean

  Dim sAccordProhibitFields As String
  Dim rsAccordDetails As DAO.Recordset
  Dim iAccordTypeID As Integer
  Dim mbAccordAllowDelete As Boolean

  Dim strDiaryProcName As String
  Dim sInsertTriggerSQL As SystemMgr.cStringBuilder
  Dim sUpdateTriggerSQL As SystemMgr.cStringBuilder
  Dim sDeleteTriggerSQL As SystemMgr.cStringBuilder
     
  Set sInsertTriggerSQL = New SystemMgr.cStringBuilder
  Set sUpdateTriggerSQL = New SystemMgr.cStringBuilder
  Set sDeleteTriggerSQL = New SystemMgr.cStringBuilder

  miTriggerRecursionLevel = IIf(gbManualRecursionLevel, giManualRecursionLevel, giDefaultRecursionLevel)
  mbAccordAllowDelete = GetModuleSetting(gsMODULEKEY_ACCORD, gsPARAMETERKEY_ALLOWDELETE, False)
    

  fOK = True
    
  Set objTable = gobjHRProEngine.GetTable(pLngCurrentTableID)

  ' We've created the code for auditing, relationships, calculations and the diary.
  ' Now put them all together to make the trigger creation string.
  '
  If fOK Then

    'Run this function that creates 3 trigger strings (insert, update & delete)
    SetTableTriggers_AutoUpdateGetField pLngCurrentTableID, psTableName
    
    '
    ' Create the INSERT trigger creation string if required.
    '
    
    strDiaryProcName = "spASRDiary_" & CStr(pLngCurrentTableID)
      
    ' Create the trigger header.
    sInsertTriggerSQL.Append _
      "    DECLARE @recordID int," & vbNewLine & _
      "        @TStamp int," & vbNewLine & _
      "        @id int," & vbNewLine & _
      "        @hResult int," & vbNewLine & _
      "        @changesMade bit," & vbNewLine & _
      "        @comparisonResult bit," & vbNewLine & _
      "        @oldCharValue varchar(max)," & vbNewLine & _
      "        @oldNumValue float," & vbNewLine & _
      "        @oldDateValue datetime," & vbNewLine & _
      "        @oldLogicValue bit," & vbNewLine & _
      "        @newCharValue varchar(max)," & vbNewLine & _
      "        @newNumValue float," & vbNewLine & _
      "        @newDateValue datetime," & vbNewLine & _
      "        @newLogicValue bit," & vbNewLine & _
      "        @iAccordDefaultStatus integer," & vbNewLine & _
      "        @iAccordBatchID integer," & vbNewLine & _
      "        @iAccordManualSendType smallint," & vbNewLine & _
      "        @bAccordResend bit," & vbNewLine & _
      "        @bAccordBypassFilter bit," & vbNewLine & _
      "        @fUpdatingDateDependentColumns bit," & vbNewLine & _
      "        @fValidRecord bit," & vbNewLine & _
      "        @sInvalidityMessage varchar(max)," & vbNewLine & _
      "        @iValidationSeverity integer," & vbNewLine

    sInsertTriggerSQL.Append _
      "        @cursInsertedRecords cursor," & vbNewLine & _
      "        @parent1TableID integer," & vbNewLine & _
      "        @parent1RecordID integer," & vbNewLine & _
      "        @parent2TableID integer," & vbNewLine & _
      "        @parent2RecordID integer," & vbNewLine & _
      "        @childRecordID integer," & vbNewLine & _
      "        @parentRecordID integer," & vbNewLine & _
      "        @recordDesc varchar(255)," & vbNewLine & _
      "        @RecalculateRecordDesc bit," & vbNewLine & _
      "        @strTemp varchar(max)," & vbNewLine & _
      "        @fResult bit," & vbNewLine & _
      "        @iTemp int" & vbNewLine & vbNewLine
      '"        @login_time datetime" & vbNewLine & vbNewLine

    sInsertTriggerSQL.Append _
      "    SET @fUpdatingDateDependentColumns = 0" & vbNewLine & _
      "    SET @RecalculateRecordDesc = 1" & vbNewLine & _
      "    SET @iAccordManualSendType = -1" & vbNewLine & _
      "    SET @fValidRecord = 1" & vbNewLine & vbNewLine

    sInsertTriggerSQL.Append _
      "    IF EXISTS(SELECT [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordRunningInBatch' AND [SettingKey] = @@SPID)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iAccordDefaultStatus = " & miAccordStatusForUtilities & vbNewLine & _
      "        SET @iAccordBatchID = (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordBatchID' AND [SettingKey] = @@SPID)" & vbNewLine & _
      "    END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iAccordDefaultStatus = " & miAccordDefaultStatus & vbNewLine & _
      "        SET @iAccordBatchID = 0" & vbNewLine & _
      "    END" & vbNewLine

      '"        @oldValue varchar(max)," & vbNewLine & _
      "        @newValue varchar(max)" & vbNewLine & vbNewLine & _

    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the required DECLARE statements to the INSERT trigger.
    sInsertTriggerSQL.Append sDeclareInsCols.ToString & vbNewLine & vbNewLine & _
      sDeclareDelCols.ToString & vbNewLine & vbNewLine
  
    sInsertTriggerSQL.Append _
      "    /* Loop through the virtual 'inserted' table, getting the record ID of each inserted record. */" & vbNewLine & _
      "    SET @cursInsertedRecords = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR SELECT inserted.id, convert(int,inserted.timestamp), inserted.[_description]" & sSelectInsCols.ToString & sSelectDelCols.ToString & " FROM inserted" & vbNewLine & _
      "    LEFT OUTER JOIN deleted ON inserted.id = deleted.id" & vbNewLine & vbNewLine & _
      "    OPEN @cursInsertedRecords" & vbNewLine & _
      "    FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @TStamp, @recordDesc" & sFetchInsCols.ToString & sFetchDelCols.ToString & vbNewLine & vbNewLine & _
      "    WHILE (@@fetch_status = 0) AND (@fValidRecord = 1)" & vbNewLine & _
      "    BEGIN" & vbNewLine

    ' Insert the expression variable declaration code.
    sInsertTriggerSQL.Append sExprDeclarationCode.ToString & vbNewLine

    If pfIsAbsenceTable Then
      sInsertTriggerSQL.Append _
        "        /* -------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
        "        /* Absence module - run the SSP calculation for all related absence records. */" & vbNewLine & _
        "        /* -------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
        "        --IF EXISTS(SELECT Name FROM sysobjects WHERE id = object_id('" & gsSSP_PROCEDURENAME & "') AND sysstat & 0xf = 4)" & vbNewLine & _
        "        --BEGIN" & vbNewLine & _
        "        --    EXEC " & gsSSP_PROCEDURENAME & " @recordID" & vbNewLine & _
        "        --END" & vbNewLine

    End If

    'MH20040331
    sInsertTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Outlook Triggers. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        IF @fValidRecord = 1" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "        IF EXISTS (SELECT Name FROM sysobjects WHERE type = 'P' AND name = 'spASROutlook_" & CStr(pLngCurrentTableID) & "')" & vbNewLine & _
      "          EXEC dbo.spASROutlook_" & CStr(pLngCurrentTableID) & " @recordID" & vbNewLine & _
      "        END" & vbNewLine


    'JPD 20070516 Fault 12231
    If sInsertWorkflowCode.Length = 0 Then
      If Application.WorkflowModule Then
        sInsertTriggerSQL.Append _
          "        /* ------------------------------ */" & vbNewLine & _
          "        /* No Workflow triggers required. */" & vbNewLine & _
          "        /* ------------------------------ */" & vbNewLine & vbNewLine
      End If
    Else
      sInsertTriggerSQL.Append vbNewLine & _
        "        /* ------------------ */" & vbNewLine & _
        "        /* Workflow Triggers. */" & vbNewLine & _
        "        /* ------------------ */" & vbNewLine & _
        "        IF @fValidRecord = 1" & vbNewLine & _
        "        BEGIN" & vbNewLine & _
        sInsertWorkflowCode.ToString & vbNewLine & _
        "        END" & vbNewLine
    End If
       

   ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the required FETCH statements to the INSERT trigger.
    'Get next record which has been inserted
    sInsertTriggerSQL.Append vbNewLine & _
      "        IF @fValidRecord = 1 FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @TStamp, @recorddesc" & sFetchInsCols.ToString & sFetchDelCols.ToString & vbNewLine & _
      "    END" & vbNewLine & _
      "    IF @fValidRecord = 1 CLOSE @cursInsertedRecords;" & vbNewLine & _
      "    DEALLOCATE @cursInsertedRecords;" & vbNewLine & vbNewLine
'      "    --PRINT CONVERT(nvarchar(28), GETDATE(),121) + ' End ([" & psTableName & "].[INS_" & psTableName & "]';" & vbNewLine & vbNewLine & _
'      "END" & vbNewLine

    ' Add special functions
    'sInsertTriggerSQL.Append sInsertSpecialFunctionsCode

    '************  DEBUG CODE  *****************
    If GetSystemSetting("development", "debug triggers", "0") = 1 Then
      Open App.Path & "\trigger_" & psTableName & "_insert.txt" For Append As #1
      Print #1, sInsertTriggerSQL.ToString
      Close #1
    End If
    '*******************************************
    
    
    sUpdateTriggerSQL.Append _
      "    DECLARE @recordID int," & vbNewLine & _
      "        @TStamp int," & vbNewLine & _
      "        @id int," & vbNewLine & _
      "        @hResult int," & vbNewLine & _
      "        @changesMade bit," & vbNewLine & "        @PreventUpdate bit," & vbNewLine & _
      "        @comparisonResult bit," & vbNewLine & _
      "        @oldCharValue varchar(max)," & vbNewLine & _
      "        @oldNumValue float," & vbNewLine & _
      "        @oldDateValue datetime," & vbNewLine & _
      "        @oldLogicValue bit," & vbNewLine & _
      "        @newCharValue varchar(max)," & vbNewLine & _
      "        @newNumValue float," & vbNewLine & _
      "        @newDateValue datetime," & vbNewLine & _
      "        @newLogicValue bit," & vbNewLine & _
      "        @fUpdatingDateDependentColumns bit," & vbNewLine & _
      "        @iAccordBatchID integer," & vbNewLine & _
      "        @iAccordDefaultStatus integer," & vbNewLine & _
      "        @iAccordManualSendType smallint," & vbNewLine & _
      "        @bAccordResend bit," & vbNewLine & _
      "        @bAccordBypassFilter bit," & vbNewLine & _
      "        @lngUpdateLoginColumnSPID float," & vbNewLine & _
      "        @lngByPassTrigger float," & vbNewLine & _
      "        @fValidRecord bit," & vbNewLine

      
    sUpdateTriggerSQL.Append _
      "        @sInvalidityMessage varchar(max)," & vbNewLine & _
      "        @iValidationSeverity integer," & vbNewLine & _
      "        @cursInsertedRecords cursor," & vbNewLine & _
      "        @parent1TableID integer," & vbNewLine & _
      "        @parent1RecordID integer," & vbNewLine & _
      "        @parent2TableID integer," & vbNewLine & _
      "        @parent2RecordID integer," & vbNewLine & _
      "        @childRecordID integer," & vbNewLine & _
      "        @parentRecordID integer," & vbNewLine & _
      "        @oldParentRecordID integer," & vbNewLine & _
      "        @recordDesc varchar(255)," & vbNewLine & _
      "        @RecalculateRecordDesc bit," & vbNewLine & _
      "        @strTemp varchar(max)," & vbNewLine & _
      "        @iTemp int," & vbNewLine & _
      "        @fResult bit" & vbNewLine & vbNewLine
      '"        @login_time datetime" & vbNewLine & vbNewLine
    
'    sUpdateTriggerSQL.Append _
'      "       -- Only fire this trigger when called from the _u02" & vbNewLine & _
'      "       IF UPDATE([updflag]) RETURN;" & vbNewLine & vbNewLine
    
    sUpdateTriggerSQL.Append _
      "    SET @RecalculateRecordDesc = 1" & vbNewLine & vbNewLine

      '"        @oldValue varchar(max)," & vbNewLine & _
      "        @newValue varchar(max)" & vbNewLine & vbNewLine

    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the required DECLARE statements to the UPDATE trigger.
    sUpdateTriggerSQL.Append sDeclareInsCols.ToString & vbNewLine & _
      sDeclareDelCols.ToString & vbNewLine & vbNewLine
    
    'JDM - 18/12/01 - Fault 3197 - Bypass validation/calcs when updating login field in Security Manager.
    sUpdateTriggerSQL.Append _
      "    /* ---------------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
      "    /* Bypass trigger if we are updating the login field through Security Manager. */" & vbNewLine & _
      "    /* ---------------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
      "    SELECT @lngUpdateLoginColumnSPID = SettingValue FROM ASRSysSystemSettings" & vbNewLine & _
      "    WHERE [Section] = 'database' AND [SettingKey] = 'UpdateLoginColumnSPID' " & vbNewLine & _
      "    IF @lngUpdateLoginColumnSPID = @@SPID RETURN" & vbNewLine & vbNewLine & vbNewLine
    
    sUpdateTriggerSQL.Append _
      "    SELECT @fUpdatingDateDependentColumns = SettingValue FROM ASRSysSystemSettings " & vbNewLine & _
      "        WHERE [Section] = 'database' AND [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
      "    SET @fUpdatingDateDependentColumns = ISNULL(@fUpdatingDateDependentColumns, 0)" & vbNewLine & vbNewLine & _
      "    SET @fValidRecord = 1" & vbNewLine & vbNewLine

    sUpdateTriggerSQL.Append _
      "    -- Payroll Export being manually run through Data Manager." & vbNewLine & _
      "    SET @iAccordManualSendType = (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordTransferType' AND [SettingKey] = @@SPID);" & vbNewLine & _
      "    SET @bAccordBypassFilter = (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordBypassFilter' AND [SettingKey] = @@SPID);" & vbNewLine & _
      "    IF EXISTS(SELECT [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordRunningInBatch' AND [SettingKey] = @@SPID)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iAccordDefaultStatus = " & miAccordStatusForUtilities & vbNewLine & _
      "        SET @iAccordBatchID = (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordBatchID' AND [SettingKey] = @@SPID)" & vbNewLine & _
      "    END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iAccordDefaultStatus = " & miAccordDefaultStatus & vbNewLine & _
      "        SET @iAccordBatchID = 0" & vbNewLine & _
      "    END" & vbNewLine & _
      "    IF @iAccordManualSendType IS NULL" & vbNewLine & "    BEGIN" & vbNewLine & _
      "      SET @iAccordManualSendType = -1" & vbNewLine & "      SET @bAccordBypassFilter = 0" & vbNewLine & "    END" & vbNewLine & _
      "    ELSE SET @fUpdatingDateDependentColumns = 1" & vbNewLine & vbNewLine

    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the required FETCH statements to the UPDATE trigger.
    sUpdateTriggerSQL.Append _
      "    /* Loop through the virtual 'inserted' table, getting the record ID of each updated record. */" & vbNewLine & _
      "    SET @cursInsertedRecords = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR SELECT inserted.id, convert(int,inserted.timestamp), inserted.[_description]" & sSelectInsCols.ToString & sSelectDelCols.ToString & vbNewLine & _
      "        FROM deleted" & vbNewLine & _
      "        LEFT OUTER JOIN dbo.[" & psTableName & "] inserted ON inserted.id = deleted.id" & vbNewLine & _
      "    OPEN @cursInsertedRecords" & vbNewLine & _
      "    FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @TStamp, @recorddesc" & sFetchInsCols.ToString & sFetchDelCols.ToString & vbNewLine

    sUpdateTriggerSQL.Append _
      "    WHILE (@@fetch_status = 0) AND (@fValidRecord = 1)" & vbNewLine & _
      "    BEGIN" & vbNewLine
    
    ' Add special functions
    sUpdateTriggerSQL.Append sUpdateSpecialFunctionsCode1
   
    ' Insert the expression variable declaration code.
    sUpdateTriggerSQL.Append sExprDeclarationCode.ToString & vbNewLine
       
    If pfIsAbsenceTable Then
      sUpdateTriggerSQL.Append _
        "        /* -------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
        "        /* Absence module - run the SSP calculation for all related absence records. */" & vbNewLine & _
        "        /* -------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
        "        --IF (@fUpdatingDateDependentColumns = 0)" & vbNewLine & _
        "        --BEGIN" & vbNewLine & _
        "        --    IF EXISTS(SELECT Name FROM sysobjects WHERE id = object_id('" & gsSSP_PROCEDURENAME & "') AND sysstat & 0xf = 4)" & vbNewLine & _
        "        --        EXEC " & gsSSP_PROCEDURENAME & " @recordID" & vbNewLine & _
        "        --END" & vbNewLine
    
    End If
    
    'Auto Update for Destination Tables for Lookup Column Type Values
    Dim sAULookupCode As String

    'A date is required to pass to the diary subroutine.  This is used for the rebuild function.
    'This date indicates not to create diary entries prior to 1980
    ' JPD20020913 Only do the diary if the record passed validation.
    sUpdateTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Diary Triggers. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        IF EXISTS(SELECT Name FROM sysobjects WHERE id = object_id('" & strDiaryProcName & "') AND sysstat & 0xf = 4)" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "            EXEC " & strDiaryProcName & " @recordID;" & vbNewLine & _
      "        END" & vbNewLine
   
    
    If Len(gstrUpdateEmailCode) <> 0 Then
      sUpdateTriggerSQL.Append vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        "        /* Email Triggers. */" & vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        "        IF @fValidRecord = 1 AND @startingtriggertable = " & pLngCurrentTableID & vbNewLine & _
        "        BEGIN" & vbNewLine & _
        gstrUpdateEmailCode & vbNewLine & _
        "        END" & vbNewLine & vbNewLine
    End If
    
    
    'MH20040331
    sUpdateTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Outlook Triggers. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        IF @fValidRecord = 1" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "          IF EXISTS (SELECT Name FROM sysobjects WHERE type = 'P' AND name = 'spASROutlook_" & CStr(pLngCurrentTableID) & "')" & vbNewLine & _
      "            EXEC dbo.spASROutlook_" & CStr(pLngCurrentTableID) & " @recordID" & vbNewLine & _
      "        END" & vbNewLine & vbNewLine
    
    
    'JPD 20070516 Fault 12231
    If sUpdateWorkflowCode.Length = 0 Then
      If Application.WorkflowModule Then
        sUpdateTriggerSQL.Append _
          "        /* ------------------------------ */" & vbNewLine & _
          "        /* No Workflow triggers required. */" & vbNewLine & _
          "        /* ------------------------------ */" & vbNewLine & vbNewLine
      End If
    Else
      sUpdateTriggerSQL.Append vbNewLine & _
        "        /* ------------------ */" & vbNewLine & _
        "        /* Workflow Triggers. */" & vbNewLine & _
        "        /* ------------------ */" & vbNewLine & _
        "        IF @fValidRecord = 1" & vbNewLine & _
        "        BEGIN" & vbNewLine & _
        sUpdateWorkflowCode.ToString & vbNewLine & _
        "        END" & vbNewLine & vbNewLine
    End If
    
    ' Insert the Payroll trigger code.
    If sUpdateAccordCode.Length = 0 Then
      sUpdateTriggerSQL.Append vbNewLine & vbNewLine & _
        "        /* ----------------------------------------- */" & vbNewLine & _
        "        /* No Payroll triggers required. */" & vbNewLine & _
        "        /* ----------------------------------------- */" & vbNewLine & vbNewLine
    Else
      sUpdateTriggerSQL.Append vbNewLine & vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        "        /* Payroll Triggers. */" & vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        "        IF @fValidRecord = 1" & vbNewLine & _
        "        BEGIN" & vbNewLine & _
        sUpdateAccordCode.ToString & _
        "        END" & vbNewLine & vbNewLine
    End If

    sUpdateTriggerSQL.Append vbNewLine & _
      "        IF @fValidRecord = 1 FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @TStamp, @recorddesc" & sFetchInsCols.ToString & sFetchDelCols.ToString & vbNewLine & _
      "    END" & vbNewLine & _
      "    IF @fValidRecord = 1 CLOSE @cursInsertedRecords;" & vbNewLine & _
      "    DEALLOCATE @cursInsertedRecords;" & vbNewLine

    ' Add special functions
    sUpdateTriggerSQL.Append sUpdateSpecialFunctionsCode2

    '************  DEBUG CODE  *****************
    If GetSystemSetting("development", "debug triggers", "0") = 1 Then
      Open App.Path & "\trigger_" & psTableName & "_update.txt" For Append As #1
      Print #1, sUpdateTriggerSQL.ToString
      Close #1
    End If
    '*******************************************

    sDeleteTriggerSQL.Append _
      "    DECLARE @recordID int," & vbNewLine & _
      "        @id int," & vbNewLine & _
      "        @hResult int," & vbNewLine & _
      "        @changesMade bit," & vbNewLine & _
      "        @comparisonResult bit," & vbNewLine & _
      "        @oldCharValue varchar(max)," & vbNewLine & _
      "        @oldNumValue float," & vbNewLine & _
      "        @oldDateValue datetime," & vbNewLine & _
      "        @oldLogicValue bit," & vbNewLine & _
      "        @newCharValue varchar(max)," & vbNewLine & _
      "        @newNumValue float," & vbNewLine & _
      "        @newDateValue datetime," & vbNewLine & _
      "        @newLogicValue bit," & vbNewLine & _
      "        @iAccordDefaultStatus integer," & vbNewLine & _
      "        @iAccordBatchID integer," & vbNewLine & _
      "        @fUpdatingDateDependentColumns bit," & vbNewLine & _
      "        @cursDeletedRecords cursor," & vbNewLine & _
      "        @parent1TableID integer," & vbNewLine & _
      "        @parent1RecordID integer," & vbNewLine & _
      "        @parent2TableID integer," & vbNewLine & _
      "        @parent2RecordID integer," & vbNewLine

    sDeleteTriggerSQL.Append _
      "        @parentRecordID integer," & vbNewLine & _
      "        @childRecordID integer," & vbNewLine & _
      "        @recordDesc varchar(255)," & vbNewLine & _
      "        @RecalculateRecordDesc bit," & vbNewLine & _
      "        @strTemp varchar(max)," & vbNewLine & _
      "        @fResult bit," & vbNewLine & _
      "        @iTemp int" & vbNewLine & vbNewLine
      
'    sDeleteTriggerSQL.Append _
'      "       -- Only fire this trigger when called from the _u02" & vbNewLine & _
'      "       IF UPDATE([updflag]) RETURN;" & vbNewLine & vbNewLine
      
    sDeleteTriggerSQL.Append _
      "    SET @RecalculateRecordDesc = 0" & vbNewLine & _
      "    SET @fUpdatingDateDependentColumns = 0" & vbNewLine & vbNewLine
    
      '"        @oldValue varchar(max)," & vbNewLine & _
      "        @newValue varchar(max)" & vbNewLine & vbNewLine & _

    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the required DECLARE statements to the DELETE trigger.
    'sDeleteTriggerSQL.Append  sDeclareDelCols & vbNewLine & vbNewLine
    sDeleteTriggerSQL.Append sDeclareInsCols.ToString & vbNewLine & vbNewLine & _
      sDeclareDelCols.ToString & vbNewLine & vbNewLine
    
    sDeleteTriggerSQL.Append _
      "    SELECT @fUpdatingDateDependentColumns = SettingValue FROM ASRSysSystemSettings " & vbNewLine & _
      "        WHERE [Section] = 'database' AND [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine
    
    'NPG20080715 Fault 13266
    sDeleteTriggerSQL.Append _
      "    SET @fUpdatingDateDependentColumns = ISNULL(@fUpdatingDateDependentColumns, 0)" & vbNewLine & vbNewLine

    sDeleteTriggerSQL.Append _
      "    IF EXISTS(SELECT [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordRunningInBatch' AND [SettingKey] = @@SPID)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iAccordDefaultStatus = " & miAccordStatusForUtilities & vbNewLine & _
      "        SET @iAccordBatchID = (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordBatchID' AND [SettingKey] = @@SPID)" & vbNewLine & _
      "    END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iAccordDefaultStatus = " & miAccordDefaultStatus & vbNewLine & _
      "        SET @iAccordBatchID = 0" & vbNewLine & _
      "    END" & vbNewLine
    
    
    ' JPD20020913 - instead of making multiple queries to the triggered table, and
    ' the 'inserted' and 'deleted' tables, we now get all of the required information in
    ' the cursor that we used to loop through to get just the id of each record being
    ' inserted/updated/deleted.
    ' Here we are adding the required FETCH statements to the DELETE trigger.
    sDeleteTriggerSQL.Append _
      "    /* Loop through the virtual 'deleted' table, getting the record ID of each deleted record. */" & vbNewLine & _
      "    SET @cursDeletedRecords = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR SELECT deleted.id" & sSelectDelCols.ToString & " FROM deleted" & vbNewLine & _
      "    OPEN @cursDeletedRecords" & vbNewLine & _
      "    FETCH NEXT FROM @cursDeletedRecords INTO @recordID" & sFetchDelCols.ToString & vbNewLine

  
    sDeleteTriggerSQL.Append _
      "    WHILE (@@fetch_status = 0)" & vbNewLine & _
      "    BEGIN" & vbNewLine
     
    If gbAccordlModule Then
    
      sSQL = "SELECT TransferTypeID FROM tmpAccordTransferTypes" _
          & " WHERE ASRBaseTableID = " & CStr(pLngCurrentTableID)
      Set rsAccordDetails = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
      If Not (rsAccordDetails.EOF And rsAccordDetails.BOF) Then
        If mbAccordAllowDelete Then
          sDeleteTriggerSQL.Append vbNewLine & _
            "        -- Prohibit delete if record has been transferred to Payroll" & vbNewLine & _
            "        EXEC dbo.spASRAccordIsRecordInPayroll @recordID, " & rsAccordDetails.Fields("TransferTypeID").value & ", @hResult OUTPUT" & vbNewLine & _
            "        IF @hResult <> 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          EXEC dbo.spasRAccordDeleteTransactionsForRecord @recordID, " & rsAccordDetails.Fields("TransferTypeID").value & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        Else
          sDeleteTriggerSQL.Append vbNewLine & _
            "        -- Prohibit delete if record has been transferred to Payroll" & vbNewLine & _
            "        EXEC dbo.spASRAccordIsRecordInPayroll @recordID, " & rsAccordDetails.Fields("TransferTypeID").value & ", @hResult OUTPUT" & vbNewLine & _
            "        IF @hResult = 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          RAISERROR ('You cannot delete a record that has been transferred to payroll.',16,@hResult)" & vbNewLine & _
            "          ROLLBACK TRANSACTION" & vbNewLine & _
            "          RETURN" & vbNewLine & _
            "        END" & vbNewLine & _
            "        ELSE EXEC dbo.spASRAccordDeleteTransactionsForRecord @recordID, " & rsAccordDetails.Fields("TransferTypeID").value & vbNewLine & vbNewLine
        End If
      End If
    End If
    
     
    If (sDeleteAuditCode.Length > 0) _
      Or (LenB(gstrDeleteEmailCode) > 0) _
      Or (sDeleteWorkflowCode.Length > 0) Then
      'MH20001109 Fault 568 Record Description not appearing
      'in Audit Log for deleted records.
      Set objExpr = New CExpression
      With objExpr
        .ExpressionID = plngRecDescExprID
        If .ConstructExpression Then

          sDeleteTriggerSQL.Append _
            "        /* ------------------------------------- */" & vbNewLine & _
            "        /* Get Record Description */" & vbNewLine & _
            "        /* ------------------------------------- */" & vbNewLine & _
            "        IF @fUpdatingDateDependentColumns = 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @id = @recordID" & vbNewLine & _
            .StoredProcedureCode("@recordDesc", "deleted") & vbNewLine & _
            "          SET @recordDesc = CONVERT(varchar(255), @recordDesc)" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine

        End If
      End With
    End If
                  
    ' Email stuff
    sDeleteTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Email Queue. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        DELETE FROM ASRSysEmailQueue" & vbNewLine & _
      "        WHERE RecordID = @recordID And ASRSysEmailQueue.TableID = " & CStr(pLngCurrentTableID) & vbNewLine & vbNewLine '& _
      "        (SELECT TableID FROM ASRSysColumns WHERE ASRSysEmailQueue.ColumnID = ASRSysColumns.ColumnID)" & vbNewLine & vbNewLine

    If LenB(gstrDeleteEmailCode) = 0 Then
      sDeleteTriggerSQL.Append _
        "        /* ----------------------------------------- */" & vbNewLine & _
        "        /* No Email triggers required. */" & vbNewLine & _
        "        /* ----------------------------------------- */" & vbNewLine & vbNewLine
    Else
      sDeleteTriggerSQL.Append vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        "        /* Email Triggers. */" & vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        gstrDeleteEmailCode & vbNewLine
    End If

    'MH20040331
    sDeleteTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Outlook Triggers. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        UPDATE dbo.[ASRSysOutlookEvents] SET Deleted = 1 " & _
               "WHERE TableID = " & CStr(pLngCurrentTableID) & _
               " AND RecordID = @recordID" & vbNewLine & vbNewLine & vbNewLine

    If sDeleteWorkflowCode.Length = 0 Then
      If Application.WorkflowModule Then
        sDeleteTriggerSQL.Append _
          "        /* ------------------------------ */" & vbNewLine & _
          "        /* No Workflow triggers required. */" & vbNewLine & _
          "        /* ------------------------------ */" & vbNewLine & vbNewLine
      End If
    Else
      sDeleteTriggerSQL.Append vbNewLine & _
        "        /* ------------------ */" & vbNewLine & _
        "        /* Workflow Triggers. */" & vbNewLine & _
        "        /* ------------------ */" & vbNewLine & _
        sDeleteWorkflowCode.ToString & vbNewLine
    End If
    
    ' Insert the Relationship trigger code.
    If LenB(sRelationshipCode.ToString) = 0 Then
      sDeleteTriggerSQL.Append _
        "        /* ---------------------------------------------------- */" & vbNewLine & _
        "        /* No Relationship triggers required. */" & vbNewLine & _
        "        /* ---------------------------------------------------- */" & vbNewLine & vbNewLine
    Else
      sDeleteTriggerSQL.Append _
        "        /* ---------------------------------- */" & vbNewLine & _
        "        /* Relationship Triggers. */" & vbNewLine & _
        "        /* ---------------------------------- */" & vbNewLine & _
        sRelationshipCode.ToString
    End If
        
    ' Insert the expression variable declaration code.
    sDeleteTriggerSQL.Append sExprDeclarationCode.ToString & vbNewLine

    sDeleteTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Diary Events. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        DELETE FROM ASRSysDiaryEvents WHERE RowID = @recordID AND TableID = " & CStr(pLngCurrentTableID)

    sDeleteTriggerSQL.Append vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        /* Workflow Queue. */" & vbNewLine & _
      "        /* ----------------------- */" & vbNewLine & _
      "        DELETE FROM ASRSysWorkflowQueue" & vbNewLine & _
      "        WHERE recordID = @recordID" & vbNewLine & _
      "           AND dateInitiated IS null" & vbNewLine & _
      "           AND linkID IN (" & vbNewLine & _
      "               SELECT WFTL.linkid" & vbNewLine & _
      "               FROM ASRSysWorkflowTriggeredLinks WFTL" & vbNewLine & _
      "               INNER JOIN ASRSysWorkflows WF ON WFTL.workflowID = WF.id" & vbNewLine & _
      "               WHERE WF.baseTable = " & CStr(pLngCurrentTableID) & vbNewLine & _
      "                   AND WFTL.recordDelete = 0)"


    If pfIsAbsenceTable Then
      sDeleteTriggerSQL.Append vbNewLine & _
        "        /* -------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
        "        /* Absence module - run the SSP calculation for all related absence records. */" & vbNewLine & _
        "        /* -------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & _
        "        --IF EXISTS(SELECT Name FROM sysobjects WHERE id = object_id('" & gsSSP_PROCEDURENAME & "') AND sysstat & 0xf = 4)" & vbNewLine & _
        "        --BEGIN" & vbNewLine & _
        "        --    EXEC " & gsSSP_PROCEDURENAME & " @recordID" & vbNewLine & _
        "        --END" & vbNewLine

    End If

    ' Insert the Payroll trigger code.
    If sDeleteAccordCode.Length = 0 Then
      sDeleteTriggerSQL.Append vbNewLine & vbNewLine & _
        "        /* ----------------------------------------- */" & vbNewLine & _
        "        /* No Payroll triggers required. */" & vbNewLine & _
        "        /* ----------------------------------------- */" & vbNewLine & vbNewLine
    Else
      sDeleteTriggerSQL.Append vbNewLine & vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        "        /* Payroll Triggers. */" & vbNewLine & _
        "        /* ----------------------- */" & vbNewLine & _
        sDeleteAccordCode.ToString & vbNewLine & vbNewLine
    End If


    sDeleteTriggerSQL.Append vbNewLine & _
      "        FETCH NEXT FROM @cursDeletedRecords INTO @recordID" & sFetchDelCols.ToString & vbNewLine & _
      "    END" & vbNewLine & _
      "    CLOSE @cursDeletedRecords" & vbNewLine & _
      "    DEALLOCATE @cursDeletedRecords" & vbNewLine

    ' Add special functions
    sDeleteTriggerSQL.Append sDeleteSpecialFunctionsCode

    'MH20090630
    sDeleteTriggerSQL.TheString = RemoveDuplicateDeclares(sDeleteTriggerSQL.ToString)
    
    
    '************  DEBUG CODE  *****************
    If GetSystemSetting("development", "debug triggers", "0") = 1 Then
      Open App.Path & "\trigger_" & psTableName & "_delete.txt" For Append As #1
      Print #1, sDeleteTriggerSQL.ToString
      Close #1
    End If
    '*******************************************
    
    ' Pass this trigegr code to the engine
    objTable.SysMgrInsertTrigger = sInsertTriggerSQL.ToString
    objTable.SysMgrUpdateTrigger = sUpdateTriggerSQL.ToString
    objTable.SysMgrDeleteTrigger = sDeleteTriggerSQL.ToString
    
  End If

TidyUpAndExit:
  
  SetTableTriggers_CreateTriggers = fOK
  
  Exit Function

ErrorTrap:
  fOK = False
  gobjProgress.Visible = False
  OutputError "Error creating table trigger"
  Err = False
  Resume TidyUpAndExit

End Function


A
Private Function SetTableTriggers_AutoUpdate(pLngCurrentTableID As Long, psTableName As String) As String

    'Get any columns that use the current table for an Auto-Update Lookup Column.
    'NB. AU = Auto-Update
    
    Dim sAUSQL As String
    Dim rsAULookupColumns As New ADODB.Recordset
    Dim sAULookupCode As New SystemMgr.cStringBuilder
    
    sAULookupCode.TheString = vbNullString
    
    sAUSQL = "SELECT ASRSysTables.TableName, " & _
      "       ASRSysColumns.ColumnName, " & _
      "       ASRSysColumns.ColumnID, " & _
      "       ASRSysColumns.DataType, " & _
      "       ASRSysColumns.Size, " & _
      "       ASRSysColumns.LookupColumnID, " & _
      "       L_Column.ColumnName AS [LookupColumnName] " & _
      "FROM ASRSysColumns " & _
      "       INNER JOIN ASRSysTables " & _
      "       ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
      "       INNER JOIN ASRSysColumns L_Column " & _
      "       ON L_Column.ColumnID = ASRSysColumns.LookupColumnID " & _
      "WHERE ASRSysColumns.LookupTableID = " & pLngCurrentTableID & " " & _
      "  AND ASRSysColumns.AutoUpdateLookupValues = 1 " & _
      "ORDER BY ASRSysTables.TableName ASC, ASRSysColumns.ColumnName "

    rsAULookupColumns.Open sAUSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        
    With rsAULookupColumns
      'Loop through the tables/columns that reference this lookup table.
      Do While Not .EOF
        Select Case !DataType
        Case dtVARCHAR, dtLONGVARCHAR
          sAULookupCode.Append "        IF (@fUpdatingDateDependentColumns = 0)" & vbNewLine & _
           "       BEGIN" & vbNewLine & _
           "           SELECT @oldCharValue = [" & !LookupColumnName & "] " & vbNewLine & _
           "           FROM Deleted" & vbNewLine & _
           "           WHERE id = @recordID" & vbNewLine & _
           "           SET @newCharValue = CONVERT(varchar(max), @col" & Trim(Str(!LookupColumnID)) & ")" & vbNewLine & _
           "           EXEC dbo.sp_ASRCaseSensitiveCompare @comparisonResult OUTPUT, @oldCharValue, @newCharValue" & vbNewLine & _
           "           IF @comparisonResult = 0 and isnull(@oldCharValue,'') <> ''" & vbNewLine & _
           "           BEGIN" & vbNewLine & _
           "             UPDATE [" & !TableName & "] " & vbNewLine & _
           "             SET [" & !TableName & "].[" & !ColumnName & "] = @col" & Trim(Str(!LookupColumnID)) & vbNewLine & _
           "             WHERE [" & !TableName & "].[" & !ColumnName & "] = @oldCharValue " & vbNewLine & _
           "           END " & vbNewLine & _
           "        END " & vbNewLine & vbNewLine
          
        Case dtINTEGER, dtNUMERIC
           sAULookupCode.Append "        IF (@fUpdatingDateDependentColumns = 0)" & vbNewLine & _
           "        BEGIN" & vbNewLine & _
           "           SELECT @col" & Trim(Str(!LookupColumnID)) & " = " & !LookupColumnName & " FROM " & psTableName & " WHERE id = @recordID" & vbNewLine & _
           "           SELECT @oldNumValue = [" & !LookupColumnName & "] " & vbNewLine & _
           "           FROM Deleted" & vbNewLine & _
           "           WHERE id = @recordID" & vbNewLine & _
           "           SET @newNumValue = CONVERT(float, @col" & Trim(Str(!LookupColumnID)) & ")" & vbNewLine & _
           "           IF (@oldNumValue <> @newNumValue) " & vbNewLine & _
           "             OR ((@oldNumValue IS NULL) AND (NOT @newNumValue IS NULL)) " & vbNewLine & _
           "             OR ((NOT @oldNumValue IS NULL) AND (@newNumValue IS NULL)) " & vbNewLine & _
           "           BEGIN" & vbNewLine & _
           "             UPDATE [" & !TableName & "] " & vbNewLine & _
           "             SET [" & !TableName & "].[" & !ColumnName & "] = @col" & Trim(Str(!LookupColumnID)) & vbNewLine & _
           "             WHERE [" & !TableName & "].[" & !ColumnName & "] = @oldNumValue " & vbNewLine & _
           "           END " & vbNewLine & _
           "        END " & vbNewLine & vbNewLine

        Case dtBIT
          sAULookupCode.Append "        IF (@fUpdatingDateDependentColumns = 0)" & vbNewLine & _
           "        BEGIN" & vbNewLine & _
           "           SELECT @col" & Trim(Str(!LookupColumnID)) & " = " & !LookupColumnName & " FROM " & psTableName & " WHERE id = @recordID" & vbNewLine & _
           "           SELECT @oldLogicValue = [" & !LookupColumnName & "] " & vbNewLine & _
           "           FROM Deleted" & vbNewLine & _
           "           WHERE id = @recordID" & vbNewLine & _
           "           SET @newLogicValue = @col" & Trim(Str(!LookupColumnID)) & vbNewLine & _
           "           IF (@oldLogicValue <> @newLogicValue) " & vbNewLine & _
           "             OR ((@oldLogicValue IS NULL) AND (NOT @newLogicValue IS NULL)) " & vbNewLine & _
           "             OR ((NOT @oldLogicValue IS NULL) AND (@newLogicValue IS NULL)) " & vbNewLine & _
           "           BEGIN" & vbNewLine & _
           "             UPDATE [" & !TableName & "] " & vbNewLine & _
           "             SET [" & !TableName & "].[" & !ColumnName & "] = @col" & Trim(Str(!LookupColumnID)) & vbNewLine & _
           "             WHERE [" & !TableName & "].[" & !ColumnName & "] = @oldLogicValue " & vbNewLine & _
           "           END " & vbNewLine & _
           "        END " & vbNewLine & vbNewLine

        Case dtTIMESTAMP
          sAULookupCode.Append "        IF (@fUpdatingDateDependentColumns = 0)" & vbNewLine & _
           "        BEGIN" & vbNewLine & _
           "           SELECT @col" & Trim(Str(!LookupColumnID)) & " = [" & !LookupColumnName & "] FROM [" & psTableName & "] WHERE id = @recordID" & vbNewLine & _
           "           SELECT @oldDateValue = [" & !LookupColumnName & "] " & vbNewLine & _
           "           FROM Deleted" & vbNewLine & _
           "           WHERE id = @recordID" & vbNewLine & _
           "           SET @newDateValue = CONVERT(datetime, convert(varchar(20), @col" & Trim(Str(!LookupColumnID)) & ", 101))" & vbNewLine & _
           "           IF (@oldDateValue <> @newDateValue) " & vbNewLine & _
           "             OR ((@oldDateValue IS NULL) AND (NOT @newDateValue IS NULL)) " & vbNewLine & _
           "             OR ((NOT @oldDateValue IS NULL) AND (@newDateValue IS NULL)) " & vbNewLine & _
           "           BEGIN" & vbNewLine & _
           "             UPDATE [" & !TableName & "] " & vbNewLine & _
           "             SET [" & !TableName & "].[" & !ColumnName & "] = @col" & Trim(Str(!LookupColumnID)) & vbNewLine & _
           "             WHERE [" & !TableName & "].[" & !ColumnName & "] = @oldDateValue " & vbNewLine & _
           "           END " & vbNewLine & _
           "        END " & vbNewLine & vbNewLine
        
        End Select

        .MoveNext
      Loop
      .Close
    End With
  
  SetTableTriggers_AutoUpdate = sAULookupCode.ToString
  
TidyUpAndExit:
  Set sAULookupCode = Nothing
  Set rsAULookupColumns = Nothing
  Exit Function

ErrorTrap:
  gobjProgress.Visible = False
  OutputError "Error creating table trigger (Auto Update)"
  Err = False
  Resume TidyUpAndExit

End Function

Private Function SetTableTriggers_AutoUpdateGetField(pLngCurrentTableID As Long, psTableName As String) As Boolean

    'TM14072004 It has been decide to remove the GetFieldFromDatabaseRecord - AutoUpdate funcionality
    'due to it not being optional, this code should still be valid for a further solution to the
    'problem.
    Exit Function
    
    'Get any columns that use the current table for an Get Field From Database Record function.
    'NB. AU = Auto-Update
    
    ' iTriggerType
    ' DELETE = 0
    ' INSERT = 1
    ' UPDATE = 2
    
    Dim sTemp As String
    
    Dim rsAUGetField As New ADODB.Recordset
    Dim sAUSQL As String
    Dim rsParentExpr As New ADODB.Recordset
    'Dim sParentExprList As String
    Dim rsParentComp As New ADODB.Recordset
    'Dim sParentCompList As String
   
    Dim sSQL As String
    Dim rsExpr As New ADODB.Recordset
    Dim iCompCount As Long
    Dim sSearchFieldColumnName As String
    Dim lngSearchFieldColumnID As Long
    Dim SearchFieldColumnDataType As SQLDataType
    Dim sSearchFieldTableName As String
    Dim lngSearchFieldTableID As Long
    Dim sSearchExpressionColumnName As String
    Dim lngSearchExpressionColumnID As Long
    Dim SearchExpressionColumnDataType As SQLDataType
    Dim sSearchExpressionTableName As String
    Dim lngSearchExpressionTableID As Long
    Dim sReturnFieldColumnName As String
    Dim lngReturnFieldColumnID As Long
    Dim ReturnFieldColumnDataType As SQLDataType
    Dim sReturnFieldTableName As String
    Dim lngReturnFieldTableID As Long
    
    Dim alngCompletedColumns() As Long
    Dim iIndex As Long
    Dim bSearchAndReturnColumnDone As Boolean
    
    ReDim alngCompletedColumns(2, 0)
    
    bSearchAndReturnColumnDone = False
    
    sAUSQL = vbNullString
    sAUSQL = sAUSQL & "/* Get all the table.columns that use expressions that have a GFFDR function that reference the current column */" & vbNewLine
    sAUSQL = sAUSQL & "SELECT DISTINCT C.columnID, C.columnName, C.columntype, C.dataType, C.calcExprID, C.tableID, T.tableName" & vbNewLine
    sAUSQL = sAUSQL & "FROM ASRSysColumns C" & vbNewLine
    sAUSQL = sAUSQL & "      INNER JOIN ASRSysTables T" & vbNewLine
    sAUSQL = sAUSQL & "      ON C.tableID = T.tableID" & vbNewLine
    sAUSQL = sAUSQL & "      Inner Join" & vbNewLine
    sAUSQL = sAUSQL & "           /* Root Expression that is directly above the GFFDR function */" & vbNewLine
    sAUSQL = sAUSQL & "           (SELECT X.exprID" & vbNewLine
    sAUSQL = sAUSQL & "            FROM ASRSysExpressions X" & vbNewLine
    sAUSQL = sAUSQL & "                  Inner Join" & vbNewLine
    sAUSQL = sAUSQL & "                      /* Get Field From Database Record function component */" & vbNewLine
    sAUSQL = sAUSQL & "                      (SELECT exprID /*, componentID, functionID*/" & vbNewLine
    sAUSQL = sAUSQL & "                      FROM ASRSysExprComponents comp" & vbNewLine
    sAUSQL = sAUSQL & "                                  Inner Join" & vbNewLine
    sAUSQL = sAUSQL & "                                      /* Field Column expression object */" & vbNewLine
    sAUSQL = sAUSQL & "                                      (SELECT R.parentComponentID /*, exprID, Name */" & vbNewLine
    sAUSQL = sAUSQL & "                                      FROM ASRSysExpressions R" & vbNewLine
    sAUSQL = sAUSQL & "                                            Inner Join" & vbNewLine
    sAUSQL = sAUSQL & "                                                /* Field Column ID sub-query */" & vbNewLine
    sAUSQL = sAUSQL & "                                                (SELECT exprID /*, componentID, fieldColumnID */" & vbNewLine
    sAUSQL = sAUSQL & "                                                From ASRSysExprComponents" & vbNewLine
    sAUSQL = sAUSQL & "                                                WHERE fieldColumnID IN (SELECT columnID FROM ASRSysColumns WHERE tableID = " & pLngCurrentTableID & ")) V" & vbNewLine
    sAUSQL = sAUSQL & "                                            ON R.exprID = V.exprID" & vbNewLine
    sAUSQL = sAUSQL & "                                      Where ParentComponentID > 0" & vbNewLine
    sAUSQL = sAUSQL & "                                          ) Y" & vbNewLine
    sAUSQL = sAUSQL & "                                  ON comp.componentID = Y.parentComponentID" & vbNewLine
    sAUSQL = sAUSQL & "                      Where FunctionID = 42" & vbNewLine
    sAUSQL = sAUSQL & "                          ) Z" & vbNewLine
    sAUSQL = sAUSQL & "                  ON X.exprID = Z.exprID" & vbNewLine
    sAUSQL = sAUSQL & "            Where ParentComponentID = 0" & vbNewLine
    sAUSQL = sAUSQL & "              ) B" & vbNewLine
    sAUSQL = sAUSQL & "      ON C.calcExprID = B.exprID" & vbNewLine

    rsAUGetField.Open sAUSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        
    With rsAUGetField
      'Loop through the tables/columns.
      Do While Not .EOF
        sSearchFieldColumnName = vbNullString
        sSearchFieldTableName = vbNullString
        lngSearchFieldColumnID = 0
        lngSearchFieldTableID = 0
        sSearchExpressionColumnName = vbNullString
        sSearchExpressionTableName = vbNullString
        lngSearchExpressionColumnID = 0
        lngSearchExpressionTableID = 0
        sReturnFieldColumnName = vbNullString
        sReturnFieldTableName = vbNullString
        lngReturnFieldColumnID = 0
        lngReturnFieldTableID = 0
        
        sSQL = vbNullString
        sSQL = sSQL & " SELECT COMP.componentID, COMP.exprID, COMP.fieldTableID, TAB.tableName, COMP.fieldColumnID, COL.columnName, COL.dataType " & vbNewLine
        sSQL = sSQL & " FROM ASRSysExprComponents COMP " & vbNewLine
        sSQL = sSQL & "       LEFT OUTER JOIN ASRSysTables TAB " & vbNewLine
        sSQL = sSQL & "       ON COMP.fieldTableID = TAB.tableID " & vbNewLine
        sSQL = sSQL & "       LEFT OUTER JOIN ASRSysColumns COL " & vbNewLine
        sSQL = sSQL & "       ON COMP.fieldColumnID = COL.columnID " & vbNewLine
        sSQL = sSQL & " WHERE COMP.exprID IN " & vbNewLine
        sSQL = sSQL & "            (SELECT exprID FROM ASRSysExpressions WHERE parentComponentID IN " & vbNewLine
        sSQL = sSQL & "                (SELECT componentID FROM ASRSysExprComponents WHERE exprID = " & !CalcExprID & ")) " & vbNewLine
        sSQL = sSQL & " ORDER BY COMP.componentID ASC " & vbNewLine
        
        rsExpr.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
            
        iCompCount = 0
        Do While Not rsExpr.EOF
          iCompCount = iCompCount + 1
          
          Select Case iCompCount
          Case 1
            sSearchFieldColumnName = rsExpr!ColumnName
            sSearchFieldTableName = rsExpr!TableName
            SearchFieldColumnDataType = rsExpr!DataType
            lngSearchFieldColumnID = rsExpr!fieldColumnID
            lngSearchFieldTableID = rsExpr!fieldTableID
          Case 2
            sSearchExpressionColumnName = rsExpr!ColumnName
            sSearchExpressionTableName = rsExpr!TableName
            SearchExpressionColumnDataType = rsExpr!DataType
            lngSearchExpressionColumnID = rsExpr!fieldColumnID
            lngSearchExpressionTableID = rsExpr!fieldTableID
          Case 3
            sReturnFieldColumnName = rsExpr!ColumnName
            sReturnFieldTableName = rsExpr!TableName
            ReturnFieldColumnDataType = rsExpr!DataType
            lngReturnFieldColumnID = rsExpr!fieldColumnID
            lngReturnFieldTableID = rsExpr!fieldTableID
          End Select
          
          rsExpr.MoveNext
        Loop
        rsExpr.Close
        
        If (lngSearchFieldTableID = pLngCurrentTableID) Or (lngReturnFieldTableID = pLngCurrentTableID) Then
          
          bSearchAndReturnColumnDone = False
          For iIndex = 0 To UBound(alngCompletedColumns, 2) Step 1
            If (alngCompletedColumns(1, iIndex) = lngSearchFieldColumnID) And (alngCompletedColumns(2, iIndex) = lngReturnFieldColumnID) Then
              bSearchAndReturnColumnDone = True
            End If
          Next iIndex
          
          If Not bSearchAndReturnColumnDone Then
            iIndex = UBound(alngCompletedColumns, 2) + 1
            ReDim Preserve alngCompletedColumns(2, iIndex)
            alngCompletedColumns(1, iIndex) = lngSearchFieldColumnID
            alngCompletedColumns(2, iIndex) = lngReturnFieldColumnID
            
            sTemp = "        IF (@fUpdatingDateDependentColumns = 0) " & vbNewLine
            sTemp = sTemp & "        BEGIN " & vbNewLine
            mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
            mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
            
            Select Case SearchFieldColumnDataType
            Case dtVARCHAR, dtLONGVARCHAR
                sTemp = "           SELECT @oldCharValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
                
                sTemp = "           SET @newCharValue = CONVERT(varchar(max), @col" & lngSearchFieldColumnID & ") " & vbNewLine
                mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
                sTemp = "           EXEC dbo.sp_ASRCaseSensitiveCompare @comparisonResult OUTPUT, @oldCharValue, @newCharValue " & vbNewLine
                sTemp = sTemp & "           IF @comparisonResult = 0 " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtINTEGER, dtNUMERIC
                sTemp = "           SELECT @oldNumValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
                
                sTemp = "           SET @newNumValue = CONVERT(float, @col" & lngSearchFieldColumnID & ") " & vbNewLine
                mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           IF (@oldNumValue <> @newNumValue) " & vbNewLine
                sTemp = sTemp & "             OR ((@oldNumValue IS NULL) AND (NOT @newNumValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "             OR ((NOT @oldNumValue IS NULL) AND (@newNumValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtBIT
                sTemp = "           SELECT @oldLogicValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
                
                sTemp = "           SET @newLogicValue = @col" & lngSearchFieldColumnID & " " & vbNewLine
                mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           IF (@oldLogicValue <> @newLogicValue) " & vbNewLine
                sTemp = sTemp & "             OR ((@oldLogicValue IS NULL) AND (NOT @newLogicValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "             OR ((NOT @oldLogicValue IS NULL) AND (@newLogicValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtTIMESTAMP
                sTemp = "           SELECT @oldCharValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
                
                sTemp = "           SET @newDateValue = CONVERT(datetime, CONVERT(varchar(20), @col" & lngSearchFieldColumnID & ", 101)) " & vbNewLine
                mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           IF (@oldDateValue <> @newDateValue) " & vbNewLine
                sTemp = sTemp & "             OR ((@oldDateValue IS NULL) AND (NOT @newDateValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "             OR ((NOT @oldDateValue IS NULL) AND (@newDateValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
              
            End Select
            
            sTemp = "             UPDATE [" & sSearchExpressionTableName & "] " & vbNewLine
            sTemp = sTemp & "             SET [" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = [" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] " & vbNewLine
            mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
            mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
            
            'Delete Trigger
            sTemp = "             WHERE ([" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = " & GetSPVariable(SearchFieldColumnDataType, False) & ") " & vbNewLine
            mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
            'Insert Trigger
            sTemp = "             WHERE ([" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = " & GetSPVariable(SearchFieldColumnDataType, True) & ") " & vbNewLine
            mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
            'Update Trigger
            sTemp = "             WHERE ([" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = " & GetSPVariable(SearchFieldColumnDataType, False) & ") OR  ([" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = " & GetSPVariable(SearchFieldColumnDataType, True) & ") " & vbNewLine
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
            
            sTemp = "           END " & vbNewLine
            sTemp = sTemp & "           ELSE " & vbNewLine
            sTemp = sTemp & "           BEGIN " & vbNewLine
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp

            Select Case ReturnFieldColumnDataType
            Case dtVARCHAR, dtLONGVARCHAR
                sTemp = "           SELECT @oldCharValue = [" & sReturnFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           SET @newCharValue = CONVERT(varchar(max), @col" & lngReturnFieldColumnID & ") " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           EXEC dbo.sp_ASRCaseSensitiveCompare @comparisonResult OUTPUT, @oldCharValue, @newCharValue " & vbNewLine
                sTemp = sTemp & "           IF @comparisonResult = 0 " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtINTEGER, dtNUMERIC
                sTemp = "           SELECT @oldNumValue = [" & sReturnFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           SET @newNumValue = CONVERT(float, @col" & lngReturnFieldColumnID & ") " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           IF (@oldNumValue <> @newNumValue) " & vbNewLine
                sTemp = sTemp & "             OR ((@oldNumValue IS NULL) AND (NOT @newNumValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "             OR ((NOT @oldNumValue IS NULL) AND (@newNumValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtBIT
                sTemp = "           SELECT @oldLogicValue = [" & sReturnFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           SET @newLogicValue = @col" & lngReturnFieldColumnID & " " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           IF (@oldLogicValue <> @newLogicValue) " & vbNewLine
                sTemp = sTemp & "             OR ((@oldLogicValue IS NULL) AND (NOT @newLogicValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "             OR ((NOT @oldLogicValue IS NULL) AND (@newLogicValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtTIMESTAMP
                sTemp = "           SELECT @oldCharValue = [" & sReturnFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "           FROM Deleted " & vbNewLine
                sTemp = sTemp & "           WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           SET @newDateValue = CONVERT(datetime, CONVERT(varchar(20), @col" & lngReturnFieldColumnID & ", 101)) " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "           IF (@oldDateValue <> @newDateValue) " & vbNewLine
                sTemp = sTemp & "             OR ((@oldDateValue IS NULL) AND (NOT @newDateValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "             OR ((NOT @oldDateValue IS NULL) AND (@newDateValue IS NULL)) " & vbNewLine
                sTemp = sTemp & "           BEGIN " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            End Select

            Select Case SearchFieldColumnDataType
            Case dtVARCHAR, dtLONGVARCHAR
                sTemp = "             SELECT @oldCharValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "             FROM Deleted " & vbNewLine
                sTemp = sTemp & "             WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "             SET @newCharValue = CONVERT(varchar(max), @col" & lngSearchFieldColumnID & ") " & vbNewLine & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtINTEGER, dtNUMERIC
                sTemp = "             SELECT @oldNumValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "             FROM Deleted " & vbNewLine
                sTemp = sTemp & "             WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "             SET @newNumValue = CONVERT(float, @col" & lngSearchFieldColumnID & ") " & vbNewLine & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtBIT
                sTemp = "             SELECT @oldLogicValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "             FROM Deleted " & vbNewLine
                sTemp = sTemp & "             WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "             SET @newLogicValue = @col" & lngSearchFieldColumnID & " " & vbNewLine & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
              
            Case dtTIMESTAMP
                sTemp = "             SELECT @oldCharValue = [" & sSearchFieldColumnName & "] " & vbNewLine
                sTemp = sTemp & "             FROM Deleted " & vbNewLine
                sTemp = sTemp & "             WHERE id = @recordID " & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
                
                sTemp = "             SET @newDateValue = CONVERT(datetime, CONVERT(varchar(20), @col" & lngSearchFieldColumnID & ", 101)) " & vbNewLine & vbNewLine
                mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
            
            End Select

            sTemp = "               UPDATE [" & sSearchExpressionTableName & "] " & vbNewLine
            sTemp = sTemp & "               SET [" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = [" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] " & vbNewLine
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
            
            sTemp = "               WHERE ([" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = " & GetSPVariable(SearchFieldColumnDataType, False) & ") OR  ([" & sSearchExpressionTableName & "].[" & sSearchExpressionColumnName & "] = " & GetSPVariable(SearchFieldColumnDataType, True) & ") " & vbNewLine
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
          
            sTemp = "             END " & vbNewLine
            sTemp = sTemp & "           END " & vbNewLine
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
          
            
            sTemp = "        END " & vbNewLine & vbNewLine
            mstrGetFieldAutoUpdateCode_INSERT = mstrGetFieldAutoUpdateCode_INSERT & sTemp
            mstrGetFieldAutoUpdateCode_UPDATE = mstrGetFieldAutoUpdateCode_UPDATE & sTemp
            mstrGetFieldAutoUpdateCode_DELETE = mstrGetFieldAutoUpdateCode_DELETE & sTemp
            
          End If
        End If
        .MoveNext
      Loop
      .Close
    End With
    Set rsAUGetField = Nothing
  
  SetTableTriggers_AutoUpdateGetField = True
  
TidyUpAndExit:
  Set rsExpr = Nothing
  Set rsAUGetField = Nothing
  Set rsParentExpr = Nothing
  Set rsParentComp = Nothing
  Exit Function

ErrorTrap:
  SetTableTriggers_AutoUpdateGetField = False
  gobjProgress.Visible = False
  OutputError "Error creating table trigger (Auto Update - Get Field From Database Record)"
  Err = False
  Resume TidyUpAndExit

End Function


Private Function SetTableTriggers_SpecialFunctions( _
  ByRef alngAuditColumns() As Long, _
  ByRef sInsertSpecialFunctionsCode As String, _
  ByRef sUpdateSpecialFunctionsCode1 As String, _
  ByRef sUpdateSpecialFunctionsCode2 As String, _
  ByRef sDeleteSpecialFunctionsCode As String, _
  ByVal pLngCurrentTableID As Long) As Boolean

  On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim rsTemp As DAO.Recordset
  Dim iLoop As Long
  Dim sTableName As String
  'Dim sSubString As String
  Dim iCount As Long
  Dim iCount2 As Long
  Dim fTableIsUsedIn_AbsenceDuration As Boolean
  Dim alngTables_AbsenceDuration() As Long
  Dim fTableIsUsedIn_AbsenceBetween2Dates As Boolean
  Dim alngTables_AbsenceBetween2Dates() As Long
  Dim fTableIsUsedIn_WorkingDaysBetween2Dates As Boolean
  Dim alngTables_WorkingDaysBetween2Dates() As Long
  Dim fIsAbsenceTable As Boolean
  Dim fIsBankHolRegionTable As Boolean
  Dim fIsBankHolTable As Boolean
  Dim fIsPersonnelTable As Boolean
  Dim fIsRegionTable As Boolean
  Dim fIsWorkingPatternTable As Boolean
  Dim lngAbsenceTable As Long
  Dim lngBankHolRegionTable As Long
  Dim sBankHolRegionTable As String
  Dim lngBankHolTable As Long
  Dim lngPersonnelTable As Long
  Dim lngRegionTable As Long
  Dim lngWorkingPatternTable As Long
  Dim alngTables_Done() As Long
  Dim fFound As Boolean
  Dim lngTableID As Long
  Dim sTemp As String
  
  Dim lngAbsenceStartDate As Long
  Dim lngAbsenceStartSession As Long
  Dim lngAbsenceEndDate As Long
  Dim lngAbsenceEndSession As Long
  Dim lngAbsenceType As Long
  Dim sAbsenceStartDate As String
  Dim sAbsenceStartSession As String
  Dim sAbsenceEndDate As String
  Dim sAbsenceEndSession As String
  Dim sAbsenceType As String
  
  Dim lngBHolRegion As Long
  Dim sBHolRegion As String
  Dim lngBHolDate As Long
  Dim sBHolDate As String
  
  Dim lngStaticRegion As Long
  Dim sStaticRegion As String
  Dim lngStaticWP As Long
  Dim sStaticWP As String
  
  Dim lngHistRegion As Long
  Dim sHistRegion As String
  Dim lngHistRegionDate As Long
  Dim sHistRegionDate As String
  
  Dim lngHistWP As Long
  Dim sHistWP As String
  Dim lngHistWPDate As Long
  Dim sHistWPDate As String
  
  Dim fTableDone As Boolean
  Dim sColumnName As String
  Dim sSQL As String
  Dim sInsertUpdate As SystemMgr.cStringBuilder
  Dim sUpdateSelect As SystemMgr.cStringBuilder
  Dim sUpdateUpdate As SystemMgr.cStringBuilder
  Dim sDeleteUpdate As SystemMgr.cStringBuilder
  Dim alngTempArray() As Long
  
  Dim sSSPSwitch1 As String
  Dim sSSPSwitch2 As String
  Dim fDoneAbsenceTable As Boolean

  bOK = True
  Set sInsertUpdate = New SystemMgr.cStringBuilder
  Set sUpdateSelect = New SystemMgr.cStringBuilder
  Set sUpdateUpdate = New SystemMgr.cStringBuilder
  Set sDeleteUpdate = New SystemMgr.cStringBuilder
  
  sInsertSpecialFunctionsCode = vbNullString
  sUpdateSpecialFunctionsCode1 = vbNullString
  sUpdateSpecialFunctionsCode2 = vbNullString
  sDeleteSpecialFunctionsCode = vbNullString

  If gbDisableSpecialFunctionAutoUpdate Then
    sInsertSpecialFunctionsCode = _
      "        /* ------------------------------------------*/" & vbNewLine & _
      "        /* Special Functions Auto Update Disabled    */" & vbNewLine & _
      "        /* ------------------------------------------*/" & vbNewLine & vbNewLine
    sUpdateSpecialFunctionsCode1 = vbNullString
    sUpdateSpecialFunctionsCode2 = _
      "        /* ------------------------------------------*/" & vbNewLine & _
      "        /* Special Functions Auto Update Disabled    */" & vbNewLine & _
      "        /* ------------------------------------------*/" & vbNewLine & vbNewLine
    sDeleteSpecialFunctionsCode = _
      "        /* ------------------------------------------*/" & vbNewLine & _
      "        /* Special Functions Auto Update Disabled    */" & vbNewLine & _
      "        /* ------------------------------------------*/" & vbNewLine & vbNewLine
    
    SetTableTriggers_SpecialFunctions = bOK
    Exit Function
  End If

  ' Check if the current table is used in the following expression function
  ' by virtue of module setup:
  '   Absence Duration
  '   Absence Between Two Dates
  '   Working Days Between Two Dates
  ' If it is then we may need to add trigger code to update any tables that use these functions.
  ' NOTE - AbsenceDuration, AbsenceBetween2Dates and WorkingDaysBetween2Dates functions
  ' are only used in column calcs in the Personnel Table, or children of the Personnel table.
  ReDim alngTables_AbsenceDuration(0)
  fTableIsUsedIn_AbsenceDuration = TableIsUsedInAbsenceDuration(pLngCurrentTableID)
  If fTableIsUsedIn_AbsenceDuration Then
    TablesThatUseFunction alngTables_AbsenceDuration, 30
  End If

  ReDim alngTables_AbsenceBetween2Dates(0)
  fTableIsUsedIn_AbsenceBetween2Dates = TableIsUsedInAbsenceBetween2Dates(pLngCurrentTableID)
  If fTableIsUsedIn_AbsenceBetween2Dates Then
    TablesThatUseFunction alngTables_AbsenceBetween2Dates, 47
    TablesThatUseFunction alngTables_AbsenceBetween2Dates, 73
  End If

  ReDim alngTables_WorkingDaysBetween2Dates(0)
  fTableIsUsedIn_WorkingDaysBetween2Dates = TableIsUsedInWorkingDaysBetween2Dates(pLngCurrentTableID)
  If fTableIsUsedIn_WorkingDaysBetween2Dates Then
    TablesThatUseFunction alngTables_WorkingDaysBetween2Dates, 46
  End If

  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE)
  If sTemp = vbNullString Then sTemp = "0"
  lngAbsenceTable = CLng(sTemp)
  fIsAbsenceTable = (pLngCurrentTableID = lngAbsenceTable)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE)
  If sTemp = vbNullString Then sTemp = "0"
  lngBankHolRegionTable = CLng(sTemp)
  fIsBankHolRegionTable = (pLngCurrentTableID = lngBankHolRegionTable)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE)
  If sTemp = vbNullString Then sTemp = "0"
  lngBankHolTable = CLng(sTemp)
  fIsBankHolTable = (pLngCurrentTableID = lngBankHolTable)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE)
  If sTemp = vbNullString Then sTemp = "0"
  lngPersonnelTable = CLng(sTemp)
  fIsPersonnelTable = (pLngCurrentTableID = lngPersonnelTable)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE)
  If sTemp = vbNullString Then sTemp = "0"
  lngRegionTable = CLng(sTemp)
  fIsRegionTable = (pLngCurrentTableID = lngRegionTable)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE)
  If sTemp = vbNullString Then sTemp = "0"
  lngWorkingPatternTable = CLng(sTemp)
  fIsWorkingPatternTable = (pLngCurrentTableID = lngWorkingPatternTable)

  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE)
  If sTemp = vbNullString Then sTemp = "0"
  lngAbsenceStartDate = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION)
  If sTemp = vbNullString Then sTemp = "0"
  lngAbsenceStartSession = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE)
  If sTemp = vbNullString Then sTemp = "0"
  lngAbsenceEndDate = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION)
  If sTemp = vbNullString Then sTemp = "0"
  lngAbsenceEndSession = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE)
  If sTemp = vbNullString Then sTemp = "0"
  lngAbsenceType = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE)
  If sTemp = vbNullString Then sTemp = "0"
  lngBHolDate = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION)
  If sTemp = vbNullString Then sTemp = "0"
  lngBHolRegion = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_REGION)
  If sTemp = vbNullString Then sTemp = "0"
  lngStaticRegion = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN)
  If sTemp = vbNullString Then sTemp = "0"
  lngStaticWP = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONFIELD)
  If sTemp = vbNullString Then sTemp = "0"
  lngHistRegion = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONDATE)
  If sTemp = vbNullString Then sTemp = "0"
  lngHistRegionDate = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD)
  If sTemp = vbNullString Then sTemp = "0"
  lngHistWP = CLng(sTemp)
  
  sTemp = ReadModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE)
  If sTemp = vbNullString Then sTemp = "0"
  lngHistWPDate = CLng(sTemp)
  
  If lngAbsenceStartDate > 0 Then
    sAbsenceStartDate = GetColumnName(lngAbsenceStartDate, True)
  End If
  If lngAbsenceStartSession > 0 Then
    sAbsenceStartSession = GetColumnName(lngAbsenceStartSession, True)
  End If
  If lngAbsenceEndDate > 0 Then
    sAbsenceEndDate = GetColumnName(lngAbsenceEndDate, True)
  End If
  If lngAbsenceEndSession > 0 Then
    sAbsenceEndSession = GetColumnName(lngAbsenceEndSession, True)
  End If
  If lngAbsenceType > 0 Then
    sAbsenceType = GetColumnName(lngAbsenceType, True)
  End If
  If lngBankHolRegionTable > 0 Then
    sBankHolRegionTable = GetTableName(lngBankHolRegionTable)
  End If
  If lngBHolDate > 0 Then
    sBHolDate = GetColumnName(lngBHolDate, True)
  End If
  If lngBHolRegion > 0 Then
    sBHolRegion = GetColumnName(lngBHolRegion, True)
  End If
  If lngStaticRegion > 0 Then
    sStaticRegion = GetColumnName(lngStaticRegion, True)
  End If
  If lngStaticWP > 0 Then
    sStaticWP = GetColumnName(lngStaticWP, True)
  End If
  If lngHistRegion > 0 Then
    sHistRegion = GetColumnName(lngHistRegion, True)
  End If
  If lngHistRegionDate > 0 Then
    sHistRegionDate = GetColumnName(lngHistRegionDate, True)
  End If
  If lngHistWP > 0 Then
    sHistWP = GetColumnName(lngHistWP, True)
  End If
  If lngHistWPDate > 0 Then
    sHistWPDate = GetColumnName(lngHistWPDate, True)
  End If

  sUpdateSelect.TheString = vbNullString
  
  If fIsAbsenceTable Then
    ' Need to update the associated parent record, only if the startDate, startSession, endDate, endSession
    ' or type values have changed.
    sInsertUpdate.TheString = vbNullString
    sUpdateUpdate.TheString = vbNullString
    sDeleteUpdate.TheString = vbNullString

    ReDim alngTables_Done(0)

    If sUpdateUpdate.Length <> 0 Then
      sUpdateUpdate.Insert 0, "SET @changesMade = 0" & vbNewLine & vbNewLine
      sUpdateUpdate.Append vbNewLine & _
        "            IF @changesMade = 1" & vbNewLine & _
        "            BEGIN" & vbNewLine

      fTableDone = False
      For iCount = 1 To UBound(alngTables_AbsenceBetween2Dates)
        fFound = False
        For iLoop = 1 To UBound(alngTables_Done)
          If alngTables_Done(iLoop) = alngTables_AbsenceBetween2Dates(iCount) Then
            fFound = True
            Exit For
          End If
        Next iLoop
        
        If (Not fFound) And (alngTables_AbsenceBetween2Dates(iCount) <> lngAbsenceTable) Then
          ' Get the first non-system column in the table.
          sColumnName = vbNullString
          
          sSQL = "SELECT columnName" & _
            " FROM tmpColumns" & _
            " WHERE tableID = " & Trim$(Str$(alngTables_AbsenceBetween2Dates(iCount))) & _
            " AND columnType <>" & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
            " AND deleted = FALSE"
          Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          If Not rsTemp.EOF Then
            sColumnName = rsTemp!ColumnName
          End If
          rsTemp.Close
          Set rsTemp = Nothing

          If LenB(sColumnName) <> 0 Then
            fTableDone = True
            sTableName = GetTableName(alngTables_AbsenceBetween2Dates(iCount))
            sUpdateUpdate.Append _
              "    IF NOT EXISTS(SELECT [spid] FROM [tbsys_intransactiontrigger] WHERE [spid] = @@spid AND [tablefromid] = " & CStr(alngTables_AbsenceBetween2Dates(iCount)) & ")" & vbNewLine & _
              "        UPDATE dbo." & sTableName & vbNewLine & _
              "            SET " & sTableName & "." & sColumnName & " = " & sTableName & "." & sColumnName & vbNewLine & _
              "            WHERE " & sTableName & ".ID" & IIf(alngTables_AbsenceBetween2Dates(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
              "                (SELECT inserted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
              "                FROM inserted)" & vbNewLine & _
              "            OR " & sTableName & ".ID" & IIf(alngTables_AbsenceBetween2Dates(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
              "                (SELECT deleted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
              "                FROM deleted)" & vbNewLine
                
            sInsertUpdate.Append _
              "    IF NOT EXISTS(SELECT [spid] FROM [tbsys_intransactiontrigger] WHERE [spid] = @@spid AND [tablefromid] = " & CStr(alngTables_AbsenceBetween2Dates(iCount)) & ")" & vbNewLine & _
              "        UPDATE dbo." & sTableName & vbNewLine & _
              "            SET " & sTableName & "." & sColumnName & " = " & sTableName & "." & sColumnName & vbNewLine & _
              "            WHERE " & sTableName & ".ID" & IIf(alngTables_AbsenceBetween2Dates(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
              "                (SELECT inserted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
              "                FROM inserted)"
            
            sDeleteUpdate.Append _
              "    IF NOT EXISTS(SELECT [spid] FROM [tbsys_intransactiontrigger] WHERE [spid] = @@spid AND [tablefromid] = " & CStr(alngTables_AbsenceBetween2Dates(iCount)) & ")" & vbNewLine & _
              "        UPDATE dbo." & sTableName & vbNewLine & _
              "            SET " & sTableName & "." & sColumnName & " = " & sTableName & "." & sColumnName & vbNewLine & _
              "            WHERE " & sTableName & ".ID" & IIf(alngTables_AbsenceBetween2Dates(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
              "                (SELECT deleted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
              "                FROM deleted)"
            
            ReDim Preserve alngTables_Done(UBound(alngTables_Done) + 1)
            alngTables_Done(UBound(alngTables_Done)) = alngTables_AbsenceBetween2Dates(iCount)
          End If
        End If
      Next iCount
        
  '    sUpdateUpdate.Append "            END"
        
      If Not fTableDone Then
        sUpdateUpdate.TheString = vbNullString
      End If
    End If
    
    sInsertSpecialFunctionsCode = sInsertSpecialFunctionsCode & _
      IIf(LenB(sInsertSpecialFunctionsCode) <> 0, vbNewLine & vbNewLine, vbNullString) & _
      sInsertUpdate.ToString
    sUpdateSpecialFunctionsCode2 = sUpdateSpecialFunctionsCode2 & _
      IIf(LenB(sUpdateSpecialFunctionsCode2) <> 0, vbNewLine & vbNewLine, vbNullString) & _
      sUpdateUpdate.ToString
    sDeleteSpecialFunctionsCode = sDeleteSpecialFunctionsCode & _
      IIf(LenB(sDeleteSpecialFunctionsCode) <> 0, vbNewLine & vbNewLine, vbNullString) & _
      sDeleteUpdate.ToString
  End If
  
  
  'JPD 20050920 Fault 10366
  If fIsRegionTable Or fIsWorkingPatternTable Then
    sInsertUpdate.TheString = vbNullString
    sUpdateUpdate.TheString = vbNullString
    sDeleteUpdate.TheString = vbNullString

    ReDim alngTables_Done(0)

      fTableDone = False
      For iCount2 = 1 To 3
        Select Case iCount2
          Case 1
            alngTempArray = alngTables_AbsenceDuration
          Case 2
            alngTempArray = alngTables_AbsenceBetween2Dates
          Case Else
            alngTempArray = alngTables_WorkingDaysBetween2Dates
        End Select

        For iCount = 1 To UBound(alngTempArray)
          fFound = False
          For iLoop = 1 To UBound(alngTables_Done)
            If alngTables_Done(iLoop) = alngTempArray(iCount) Then
              fFound = True
              Exit For
            End If
          Next iLoop

          If (Not fFound) And _
            (alngTempArray(iCount) <> pLngCurrentTableID) Then
            
            ' Get the first non-system column in the table.
            sColumnName = vbNullString

            sSQL = "SELECT columnName" & _
              " FROM tmpColumns" & _
              " WHERE tableID = " & Trim$(Str$(alngTempArray(iCount))) & _
              " AND columnType <>" & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
              " AND deleted = FALSE"
            Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            If Not rsTemp.EOF Then
              sColumnName = rsTemp!ColumnName
            End If
            rsTemp.Close
            Set rsTemp = Nothing

            If LenB(sColumnName) <> 0 Then
              fTableDone = True
              sTableName = GetTableName(alngTempArray(iCount))

              sUpdateUpdate.Append _
                "    IF NOT EXISTS(SELECT [spid] FROM [tbsys_intransactiontrigger] WHERE [spid] = @@spid AND [tablefromid] = " & CStr(alngTempArray(iCount)) & ")" & vbNewLine & _
                "        UPDATE dbo." & sTableName & vbNewLine & _
                "            SET " & sTableName & "." & sColumnName & " = " & sTableName & "." & sColumnName & vbNewLine & _
                "            WHERE " & sTableName & ".ID" & IIf(alngTempArray(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
                "                (SELECT inserted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
                "                FROM inserted)" & vbNewLine & _
                "            OR " & sTableName & ".ID" & IIf(alngTempArray(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
                "                (SELECT deleted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
                "                FROM deleted)" & vbNewLine & vbNewLine

              sInsertUpdate.Append _
                "    IF NOT EXISTS(SELECT [spid] FROM [tbsys_intransactiontrigger] WHERE [spid] = @@spid AND [tablefromid] = " & CStr(alngTempArray(iCount)) & ")" & vbNewLine & _
                "    UPDATE dbo." & sTableName & vbNewLine & _
                "        SET " & sTableName & "." & sColumnName & " = " & sTableName & "." & sColumnName & vbNewLine & _
                "        WHERE " & sTableName & ".ID" & IIf(alngTempArray(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
                "            (SELECT inserted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
                "            FROM inserted)" & vbNewLine & vbNewLine

              sDeleteUpdate.Append _
                "    IF NOT EXISTS(SELECT [spid] FROM [tbsys_intransactiontrigger] WHERE [spid] = @@spid AND [tablefromid] = " & CStr(alngTempArray(iCount)) & ")" & vbNewLine & _
                "    UPDATE dbo." & sTableName & vbNewLine & _
                "        SET " & sTableName & "." & sColumnName & " = " & sTableName & "." & sColumnName & vbNewLine & _
                "        WHERE " & sTableName & ".ID" & IIf(alngTempArray(iCount) = lngPersonnelTable, vbNullString, "_" & CStr(lngPersonnelTable)) & " IN " & vbNewLine & _
                "            (SELECT deleted.ID_" & CStr(lngPersonnelTable) & vbNewLine & _
                "            FROM deleted)" & vbNewLine & vbNewLine

              ReDim Preserve alngTables_Done(UBound(alngTables_Done) + 1)
              alngTables_Done(UBound(alngTables_Done)) = alngTempArray(iCount)
            End If
          End If
        Next iCount
      Next iCount2

'      sUpdateUpdate.Append "            END"

    If Not fTableDone Then
      sUpdateUpdate.TheString = vbNullString
    End If


    sInsertSpecialFunctionsCode = sInsertSpecialFunctionsCode & _
      IIf(LenB(sInsertSpecialFunctionsCode) <> 0, vbNewLine & vbNewLine, vbNullString) & _
      sInsertUpdate.ToString
    sUpdateSpecialFunctionsCode2 = sUpdateSpecialFunctionsCode2 & _
      IIf(LenB(sUpdateSpecialFunctionsCode2) <> 0, vbNewLine & vbNewLine, vbNullString) & _
      sUpdateUpdate.ToString
    sDeleteSpecialFunctionsCode = sDeleteSpecialFunctionsCode & _
      IIf(LenB(sDeleteSpecialFunctionsCode) <> 0, vbNewLine & vbNewLine, vbNullString) & _
      sDeleteUpdate.ToString
  End If

  If LenB(sUpdateSpecialFunctionsCode2) <> 0 Then
  
    fDoneAbsenceTable = False
    For iLoop = 1 To UBound(alngTables_Done)
      If alngTables_Done(iLoop) = lngAbsenceTable Then
        fDoneAbsenceTable = True
        Exit For
      End If
    Next iLoop

    sInsertSpecialFunctionsCode = vbNewLine & vbNewLine & _
      "    /* ------------------------------------------*/" & vbNewLine & _
      "    /* Special Functions                         */" & vbNewLine & _
      "    /* ------------------------------------------*/" & vbNewLine & _
      sSSPSwitch1 & _
      sInsertSpecialFunctionsCode & vbNewLine & _
      sSSPSwitch2 & vbNewLine & vbNewLine

    sUpdateSpecialFunctionsCode2 = vbNewLine & vbNewLine & _
      "    /* ------------------------------------------*/" & vbNewLine & _
      "    /* Special Functions                         */" & vbNewLine & _
      "    /* ------------------------------------------*/" & vbNewLine & _
      sSSPSwitch1 & _
      sUpdateSpecialFunctionsCode2 & vbNewLine & _
      sSSPSwitch2 & vbNewLine & vbNewLine
  
    sDeleteSpecialFunctionsCode = vbNewLine & vbNewLine & _
      "    /* ------------------------------------------*/" & vbNewLine & _
      "    /* Special Functions                         */" & vbNewLine & _
      "    /* ------------------------------------------*/" & vbNewLine & _
      sSSPSwitch1 & _
      sDeleteSpecialFunctionsCode & vbNewLine & _
      sSSPSwitch2 & vbNewLine & vbNewLine
      
  End If
  
TidyUpAndExit:
  SetTableTriggers_SpecialFunctions = bOK
  Exit Function

ErrorTrap:
  bOK = False
  gobjProgress.Visible = False
  OutputError "Error creating Special Functions table trigger"
  Err = False
  Resume TidyUpAndExit

End Function

Private Function SetTableTriggers_SpecialFunctions_AddColumn( _
  ByRef alngAuditColumns() As Long, _
  plngASRColumnID As Long) As Boolean

On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim bColFound As Boolean
  Dim sColumnName As String
  Dim iLoop As Integer
  
  If plngASRColumnID <= 0 Then Exit Function
  
  bOK = True
  
  ' Check if the column has already been declared and added to the select and fetch strings
  For iLoop = 1 To UBound(alngAuditColumns)
    If alngAuditColumns(iLoop) = plngASRColumnID Then
      bColFound = True
      Exit For
    End If
  Next iLoop
        
  If Not bColFound Then
    ReDim Preserve alngAuditColumns(UBound(alngAuditColumns) + 1)
    alngAuditColumns(UBound(alngAuditColumns)) = plngASRColumnID
    
    sColumnName = GetColumnName(plngASRColumnID, True)

    sSelectInsCols.Append ", inserted." & sColumnName
    sSelectDelCols.Append ", deleted." & sColumnName

    sFetchInsCols.Append ", @insCol_" & Trim(Str(plngASRColumnID))
    sFetchDelCols.Append ", @delCol_" & Trim(Str(plngASRColumnID))

    sDeclareInsCols.Append ", @insCol_" & Trim(Str(plngASRColumnID))
    sDeclareDelCols.Append ", @delCol_" & Trim(Str(plngASRColumnID))
  
    Select Case GetColumnDataType(plngASRColumnID)
    Case dtVARCHAR
      If Not bColFound Then
        sDeclareInsCols.Append " varchar(MAX)"
        sDeclareDelCols.Append " varchar(MAX)"
      End If
  
    Case dtLONGVARCHAR
      If Not bColFound Then
        sDeclareInsCols.Append " varchar(14)"
        sDeclareDelCols.Append " varchar(14)"
      End If
  
    Case dtINTEGER
      If Not bColFound Then
        sDeclareInsCols.Append " integer"
        sDeclareDelCols.Append " integer"
      End If
  
    Case dtNUMERIC
      If Not bColFound Then
        sDeclareInsCols.Append " numeric(" & Trim$(Str$(GetColumnSize(plngASRColumnID, False))) & "," & Trim$(Str$(GetColumnSize(plngASRColumnID, True))) & ")"
        sDeclareDelCols.Append " numeric(" & Trim$(Str$(GetColumnSize(plngASRColumnID, False))) & "," & Trim$(Str$(GetColumnSize(plngASRColumnID, True))) & ")"
      End If
  
    ' For Payroll date formats are converted to YYYYMMDD
    Case dtTIMESTAMP
      If Not bColFound Then
        sDeclareInsCols.Append " datetime"
        sDeclareDelCols.Append " datetime"
      End If
      
    Case dtBIT
      If Not bColFound Then
        sDeclareInsCols.Append " bit"
        sDeclareDelCols.Append " bit"
      End If
  
    Case dtVARBINARY, dtLONGVARBINARY
      If Not bColFound Then
        sDeclareInsCols.Append " varchar(255)"
        sDeclareDelCols.Append " varchar(255)"
      End If
  
    Case Else
      If Not bColFound Then
        sDeclareInsCols.Append " varchar(max)"
        sDeclareDelCols.Append " varchar(max)"
      End If
    
    End Select
  End If

TidyUpAndExit:
  SetTableTriggers_SpecialFunctions_AddColumn = bOK
  Exit Function

ErrorTrap:
  bOK = False
  gobjProgress.Visible = False
  OutputError "Error creating Special Functions (Add Absence Columns) table trigger"
  Err = False
  Resume TidyUpAndExit

End Function

Private Function GetSPVariable(dt As SQLDataType, bNew As Boolean) As String

  Select Case dt
  Case dtVARCHAR, dtLONGVARCHAR
    GetSPVariable = IIf(bNew, "@newCharValue", "@oldCharValue")
  Case dtINTEGER, dtNUMERIC
    GetSPVariable = IIf(bNew, "@newNumValue", "@oldNumValue")
  Case dtBIT
    GetSPVariable = IIf(bNew, "@newLogicValue", "@oldLogicValue")
  Case dtTIMESTAMP
    GetSPVariable = IIf(bNew, "@newDateValue", "@oldDateValue")
  End Select

End Function



Private Function GetTriggerRelationshipCode(pLngCurrentTableID)

  Dim sSQL As String
  Dim lngChildTableID As Long
  Dim sChildTable As String
  Dim iParentCalc As Long

  Dim rsParents As ADODB.Recordset
  Dim rsChildren As ADODB.Recordset
  
  Set rsParents = New ADODB.Recordset
  Set rsChildren = New ADODB.Recordset
  
  '
  ' Create the trigger code required to handle the relationships.
  ' ie. the code to delete any child records when a record in the given table is deleted.
  '
  Set sRelationshipCode = New SystemMgr.cStringBuilder
  
  
  sRelationshipCode.TheString = vbNullString

  ' Get the given table's children.
  sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
    " FROM ASRSysRelations " & _
    " INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID" & _
    " WHERE ASRSysRelations.parentID = " & Trim$(Str$(pLngCurrentTableID))
  rsChildren.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly, adCmdText
  
  ' Loop through the current table's children.
  Do While Not rsChildren.EOF
    lngChildTableID = rsChildren.Fields(1).value
    sChildTable = rsChildren.Fields(0).value
    
    ' Create the code for deleting all records in the child table that
    ' are related to the record that has just been deleted in the given table.
    ' NB. We only want to delete the related records from the child table
    ' if the child recordsthey have no other related parents.
    sRelationshipCode.Append _
      "        DELETE FROM " & sChildTable & _
      " WHERE " & sChildTable & ".ID_" & Trim$(Str$(pLngCurrentTableID)) & " = @recordID" & vbNewLine

    ' Get the list of other parents of the current child table.
    sSQL = "SELECT ParentID FROM ASRSysRelations" & _
      " WHERE childID = " & Trim$(Str$(lngChildTableID)) & _
      " AND parentID <> " & Trim$(Str$(pLngCurrentTableID))

    rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    ' Loop through the other parents of the child table.
    iParentCalc = 0
    Do While Not rsParents.EOF
      ' Ensure that rows are only deleted when all parents are deleted.
      iParentCalc = iParentCalc + 1
      sRelationshipCode.Append "            AND " & sChildTable & ".ID_" & Trim(Str(rsParents(0).value)) & " IS NULL" & vbNewLine
      rsParents.MoveNext
    Loop
    
    If iParentCalc > 0 Then
      sRelationshipCode.Append "        UPDATE " & sChildTable & _
        " SET " & sChildTable & ".ID_" & Trim$(Str$(pLngCurrentTableID)) & " = null" & _
        " WHERE ID_" & Trim$(Str$(pLngCurrentTableID)) & " = @recordID" & vbNewLine
    End If
    
    rsParents.Close
   
    sRelationshipCode.Append vbNewLine

    rsChildren.MoveNext
  Loop
  
  rsChildren.Close

  Set rsChildren = Nothing
  Set rsParents = Nothing

End Function

' Drops the specified trigger
Public Function DropTrigger(ByVal psTriggerName As String) As Boolean

  Dim sSQL As String
  Dim bOK As Boolean

  bOK = True
  sSQL = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('[" & psTriggerName & "]')" & _
    "     AND objectproperty(id, N'IsTrigger') = 1)" & _
    " DROP TRIGGER [" & psTriggerName & "]"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  DropTrigger = bOK

End Function

