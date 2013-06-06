VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprWorkflowField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngTableID As Long
Private mlngColumnID As Long

Private miRecordSelection As WorkflowRecordSelectorTypes
Private msElementIdentifier As String
Private msItemIdentifier As String
Private mlngRecordTableID As Long

Private miSelectionType As FieldSelectionTypes
Private mlngSelectionLine As Long
Private mlngSelOrderID As Long
Private mlngSelFilterID As Long

' Class handling variables.
Private mobjBaseComponent As CExprComponent



Public Property Get WorkflowElement() As String
  ' Return the WorkflowElement property.
  WorkflowElement = msElementIdentifier
  
End Property

Public Property Get WorkflowItem() As String
  ' Return the Workflow Item property.
  WorkflowItem = msItemIdentifier
  
End Property


Public Property Let WorkflowElement(ByVal psNewValue As String)
  ' Set the Workflow Element property.
  msElementIdentifier = psNewValue

End Property
Public Property Let WorkflowItem(ByVal psNewValue As String)
  ' Set the Workflow Item property.
  msItemIdentifier = psNewValue

End Property
Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap

  ContainsExpression = False

  If mlngSelFilterID > 0 Then
    ' Check if the calc component IS the one we're checking for.
    ContainsExpression = (plngExprID = mlngSelFilterID)

    If Not ContainsExpression Then
      ' The calc component IS NOT the one we're checking for.
      ' Check if it contains the one we're looking for.
      ContainsExpression = HasExpressionComponent(mlngSelFilterID, plngExprID)
    End If
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function







Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  glngBottom = CalculateBottomOfPage
  
  ' Position the printing.
  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  Printer.Print ComponentDescription
  
  CheckEndOfPage2 glngBottom
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function StoredProcedureCode(Optional strBaseTable As String) As String
    ' Return a string describing the component in a stored procedure.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim fFound As Boolean
  Dim lngColumnSize As Long
  Dim iDataType As DataTypes
  Dim iColumnDataType As Integer
  Dim sSQL As String
  Dim sVarName As String
  Dim sProcCode As New HRProSystemMgr.cStringBuilder
  Dim sColumnName As String
  Dim sTableName As String
  Dim sOrderCode As String
  Dim sIfNullCode As String
  Dim sLocalResult As String
  Dim sConvertCode As String
  Dim sFetchNextCode As String
  Dim rsInfo As dao.Recordset
  Dim avOrderJoinTables() As Variant
  Dim objColumn As Column
  Dim iPersTableID As Integer
  Dim rsTemp As dao.Recordset
  Dim lngElementID As Long
  Dim iElementType As ElementType

  fOK = True
  sVarName = "component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  sLocalResult = "@" & sVarName & "_localResult "

  ' Get the field component's table name.
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", mlngTableID

    If Not .NoMatch Then
      sTableName = !TableName
    Else
      fOK = False
    End If
  End With

  ' Get the column name and data type.
  If fOK And (miSelectionType <> giSELECT_RECORDCOUNT) Then
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", mlngColumnID

      If Not .NoMatch Then
        sColumnName = !ColumnName
        iColumnDataType = !DataType

        Set objColumn = New Column
        With objColumn
          .ColumnID = mlngColumnID
          .TableID = mlngTableID
          .ReadColumn
          iDataType = .Properties("dataType")
          lngColumnSize = .Properties("size")
        End With
        Set objColumn = Nothing
      Else
        fOK = False
      End If
    End With
  End If

  ' Compose the stored procedure code for returning the required field value.
  If fOK Then
    ' Construct the 'by value' field component strored procedure code.
    sProcCode.TheString = _
      "    -- Evaluate workflow identified field component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & vbNewLine
    
    If miSelectionType = giSELECT_RECORDCOUNT Then
      sProcCode.Append _
        "    DECLARE @" & sVarName & " float" & vbNewLine & _
        "    SET @" & sVarName & " = 0" & vbNewLine & vbNewLine
      sIfNullCode = "SET @" & sVarName & " = 0"
    Else
      Select Case iDataType
        Case dtVARCHAR, dtLONGVARCHAR
          sProcCode.Append _
            "    DECLARE @" & sVarName & " varchar(MAX)" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult varChar(MAX)" & vbNewLine & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = ''"
          sConvertCode = "SET @" & sVarName & " = CONVERT(varchar(MAX), " & sLocalResult & ")"
          
        Case dtNUMERIC, dtINTEGER
          sProcCode.Append _
            "    DECLARE @" & sVarName & " float" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult float" & vbNewLine & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = 0"
          sConvertCode = "SET @" & sVarName & " = CONVERT(float, " & sLocalResult & ")"
          
        Case dtBIT
          sProcCode.Append _
            "    DECLARE @" & sVarName & " bit" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult bit" & vbNewLine & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = 0"
          sConvertCode = "SET @" & sVarName & " = CONVERT(bit, " & sLocalResult & ")"
          
        Case dtTIMESTAMP
          sProcCode.Append _
            "    DECLARE @" & sVarName & " datetime" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult datetime" & vbNewLine & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = null"
          sConvertCode = "SET @" & sVarName & " = CONVERT(datetime, " & sLocalResult & ")"
          
        Case Else
          fOK = False
      End Select
    End If
    
    If fOK Then
      If miSelectionType = giSELECT_RECORDTOTAL Then
        sProcCode.Append vbNewLine & _
          "    SET @" & sVarName & " = 0" & vbNewLine
      End If

      sProcCode.Append _
        "    SET @fDeletedValue = 0" & vbNewLine & vbNewLine

      Select Case miRecordSelection
        Case giWFRECSEL_INITIATOR ' 0 - Initiator's personnel table record
          ' Get the personnel table
          iPersTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
          
          sProcCode.Append _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
            "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
            "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
            "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
            "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
            "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
            "       @iBaseTableID = " & CStr(iPersTableID) & vbNewLine & _
            "    FROM ASRSysWorkflowInstances" & vbNewLine & _
            "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
      
        Case giWFRECSEL_IDENTIFIEDRECORD ' 1 = Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
          sSQL = "SELECT TOP 1 tmpWorkflowElements.ID," & _
            "   tmpWorkflowElements.type" & _
            " FROM tmpWorkflowElements" & _
            " WHERE tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & _
            "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'"
      
          Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          With rsTemp
            If Not (.BOF And .EOF) Then
              lngElementID = !id
              iElementType = !Type
            End If
      
            .Close
          End With
          Set rsTemp = Nothing
      
          If iElementType = elem_WebForm Then
            ' WebForm
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @iBaseRecordID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN convert(integer, ISNULL(IV.tempValue, '0'))" & _
              "            ELSE convert(integer, ISNULL(IV.value, '0'))" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iWorkflowID = Es.workflowID," & vbNewLine & _
              "        @iParent1TableID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent1TableID" & _
              "            ELSE IV.parent1TableID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iParent1RecordID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent1RecordID" & _
              "            ELSE IV.parent1RecordID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iParent2TableID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent2TableID" & _
              "            ELSE IV.parent2TableID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iParent2RecordID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent2RecordID" & _
              "            ELSE IV.parent2RecordID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iBaseTableID = EI.tableID" & vbNewLine
            
            sProcCode.Append _
              "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID" & vbNewLine & _
              "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = '" & Replace(UCase(Trim(msItemIdentifier)), "'", "''") & "'" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'" & vbNewLine & _
              "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "        AND Es.ID = IV.elementID" & vbNewLine & vbNewLine
          Else
            ' StoredData
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @iBaseRecordID = convert(integer, ISNULL(IV.value, '0'))," & vbNewLine & _
              "       @iWorkflowID = Es.workflowID," & vbNewLine & _
              "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
              "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
              "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
              "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
              "       @iBaseTableID = isnull(Es.dataTableID, 0)" & vbNewLine & _
              "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = UPPER(LTRIM(RTRIM(Es.identifier)))" & vbNewLine & _
              "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'" & vbNewLine & _
              "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & vbNewLine
          End If
      
        Case giWFRECSEL_TRIGGEREDRECORD ' 4 = Triggered Base table record
          sProcCode.Append _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
            "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
            "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
            "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
            "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
            "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
            "       @iBaseTableID = isnull(ASRSysWorkflows.baseTable, 0)" & vbNewLine & _
            "    FROM ASRSysWorkflowInstances" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
      End Select

      If (miRecordSelection = giWFRECSEL_INITIATOR) _
        Or (miRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
        Or (miRecordSelection = giWFRECSEL_TRIGGEREDRECORD) Then
      
        sProcCode.Append _
          "    EXEC spASRWorkflowAscendantRecordID" & vbNewLine & _
          "        @iBaseTableID," & vbNewLine & _
          "        @iBaseRecordID," & vbNewLine & _
          "        @iParent1TableID," & vbNewLine & _
          "        @iParent1RecordID," & vbNewLine & _
          "        @iParent2TableID," & vbNewLine & _
          "        @iParent2RecordID," & vbNewLine & _
          "        " & CStr(mlngRecordTableID) & "," & vbNewLine & _
          "        @ID  OUTPUT" & vbNewLine & vbNewLine & _
          "    SET @fValidRecordID = 0" & vbNewLine & vbNewLine & _
          "    IF @ID > 0" & vbNewLine & _
          "    BEGIN" & vbNewLine & _
          "        EXEC spASRWorkflowValidTableRecord" & vbNewLine & _
          "            " & CStr(mlngRecordTableID) & "," & vbNewLine & _
          "            @ID," & vbNewLine & _
          "            @fValidRecordID OUTPUT" & vbNewLine & _
          "    END" & vbNewLine & vbNewLine
      End If

      If mlngTableID = mlngRecordTableID Then
        ' Get field from the record table itself.
        sProcCode.Append _
          "    IF @fValidRecordID = 0" & vbNewLine & _
          "    BEGIN" & vbNewLine & _
          "        SET @sValue = null" & vbNewLine & vbNewLine
          
        If miRecordSelection = giWFRECSEL_TRIGGEREDRECORD Then
          sProcCode.Append _
            "        SELECT @iCount = COUNT(*)" & vbNewLine & _
            "        FROM ASRSysWorkflowQueueColumns QC" & vbNewLine & _
            "        INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID" & vbNewLine & _
            "        WHERE WFQ.instanceID = @piInstanceID" & vbNewLine & _
            "            AND QC.columnID = " & CStr(mlngColumnID) & vbNewLine & vbNewLine & _
            "        IF @iCount = 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '')))" & vbNewLine & _
            "            FROM ASRSysWorkflowQueueColumns QC" & vbNewLine & _
            "            INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID" & vbNewLine & _
            "            WHERE WFQ.instanceID = @piInstanceID" & vbNewLine & _
            "                AND QC.columnID = " & CStr(mlngColumnID) & vbNewLine & vbNewLine & _
            "            SET @fValidRecordID = 1" & vbNewLine & _
            "            SET @fDeletedValue = 1" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        ElseIf miRecordSelection = giWFRECSEL_IDENTIFIEDRECORD Then
          sProcCode.Append _
            "        SELECT @iCount = COUNT(*)" & vbNewLine & _
            "        FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
            "        WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
            "            AND IV.columnID = " & CStr(mlngColumnID) & vbNewLine & _
            "            AND IV.elementID = " & CStr(lngElementID) & vbNewLine & vbNewLine & _
            "        IF @iCount = 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SELECT @sValue = rtrim(ltrim(isnull(IV.value , '')))" & vbNewLine & _
            "            FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
            "            WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
            "                AND IV.columnID = " & CStr(mlngColumnID) & vbNewLine & _
            "                AND IV.elementID = " & CStr(lngElementID) & vbNewLine & vbNewLine & _
            "            SET @fValidRecordID = 1" & vbNewLine & _
            "            SET @fDeletedValue = 1" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        End If
      
        Select Case iDataType
          Case dtVARCHAR, dtLONGVARCHAR
            sProcCode.Append _
              "        SET @" & sVarName & " = isnull(@sValue, '')" & vbNewLine
      
          Case dtNUMERIC, dtINTEGER
            sProcCode.Append _
              "        SET @" & sVarName & " = convert(float, isnull(@sValue, 0))" & vbNewLine
      
          Case dtBIT
            sProcCode.Append _
              "        SET @" & sVarName & " = convert(bit, isnull(@sValue, 0))" & vbNewLine
      
          Case dtTIMESTAMP
            sProcCode.Append _
              "        SET @" & sVarName & " = convert(datetime, @sValue)" & vbNewLine
        End Select
        
        sProcCode.Append _
          "    END" & vbNewLine & _
          "    ELSE" & vbNewLine & _
          "    BEGIN" & vbNewLine & _
          "        SELECT @" & sVarName & " = " & sTableName & "." & sColumnName & vbNewLine & _
          "        FROM " & sTableName & vbNewLine & _
          "        WHERE " & sTableName & ".ID = @ID" & vbNewLine & _
          "    END" & vbNewLine & vbNewLine
      Else
        ' Getting field (or aggregate) from child table
        
        If miSelectionType = giSELECT_RECORDCOUNT Then
          sProcCode.Append _
            "    DECLARE " & sVarName & "_cursor CURSOR LOCAL FAST_FORWARD READ_ONLY" & vbNewLine & _
            "        FOR SELECT " & sTableName & ".id" & vbNewLine
      
          sFetchNextCode = "FETCH NEXT FROM " & sVarName & "_cursor INTO @filterRecordID"
        Else
          sProcCode.Append _
            "    DECLARE " & sVarName & "_cursor CURSOR LOCAL FAST_FORWARD READ_ONLY" & vbNewLine & _
            "        FOR SELECT " & sTableName & ".id, " & sTableName & "." & sColumnName & vbNewLine
            
          sFetchNextCode = "FETCH NEXT FROM " & sVarName & "_cursor INTO @filterRecordID, " & sLocalResult
        End If
      
        sProcCode.Append _
          "        FROM " & sTableName & vbNewLine
      
        ' Create the order code, and an array of parent tables that need to be joined for the order.
        ' Array column 1 = parent table id
        ' Array column 2 = parent table name
        sOrderCode = ""
        ReDim avOrderJoinTables(2, 0)
      
        If mlngSelOrderID > 0 Then
          sSQL = "SELECT tmpColumns.columnName, tmpColumns.tableID, tmpTables.tableName, tmpOrderItems.ascending" & _
            " FROM tmpOrderItems, tmpColumns, tmpTables" & _
            " WHERE orderID = " & CStr(mlngSelOrderID) & _
            " AND type = 'O'" & _
            " AND tmpColumns.columnID = tmpOrderItems.columnID" & _
            " AND tmpColumns.tableID = tmpTables.tableID" & _
            " ORDER BY sequence"
      
          Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
          With rsInfo
            Do While Not .EOF
              sOrderCode = sOrderCode & IIf(LenB(sOrderCode) <> 0, ", ", "") & !TableName & "." & !ColumnName & _
                IIf(!Ascending, "", " DESC")
      
              ' Add the table definition to the array of parent tables if required.
              If mlngTableID <> !TableID Then
                fFound = False
      
                For iLoop = 1 To UBound(avOrderJoinTables, 2)
                  If avOrderJoinTables(1, iLoop) = !TableID Then
                    fFound = True
                    Exit For
                  End If
                Next iLoop
      
                If Not fFound Then
                  iNextIndex = UBound(avOrderJoinTables, 2) + 1
                  ReDim Preserve avOrderJoinTables(2, iNextIndex)
                  avOrderJoinTables(1, iNextIndex) = !TableID
                  avOrderJoinTables(2, iNextIndex) = !TableName
                End If
              End If
      
              .MoveNext
            Loop
      
            .Close
          End With
          Set rsInfo = Nothing
        End If
      
        If UBound(avOrderJoinTables, 2) > 0 Then
          For iLoop = 1 To UBound(avOrderJoinTables, 2)
            sProcCode.Append _
              "        LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & sTableName & ".ID_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".ID" & vbNewLine
          Next iLoop
        End If
      
        sProcCode.Append _
          "        WHERE " & sTableName & ".ID_" & mlngRecordTableID & " = @id" & vbNewLine
      
        If LenB(sOrderCode) <> 0 Then
          sProcCode.Append _
            "        ORDER BY " & sOrderCode & vbNewLine
        End If
      
        If fOK Then
          ' Add the code for opening the cursor.
          sProcCode.Append _
            "    OPEN " & sVarName & "_cursor" & vbNewLine & _
            "    SET @loopCounter = 0 " & vbNewLine & _
            "    " & sFetchNextCode & vbNewLine & _
            "    WHILE (@@fetch_status = 0)" & vbNewLine & _
            "    BEGIN" & vbNewLine
      
          ' Add the code to handle filters.
          If mlngSelFilterID > 0 Then
            sProcCode.Append _
              "        EXEC @hResult = dbo.sp_ASRExpr_" & Trim$(Str$(mlngSelFilterID)) & " @filterResult OUTPUT, @filterRecordID, @piInstanceID, @piTempElement" & vbNewLine & _
              "        IF @hResult <> 0 SET @filterResult = 0" & vbNewLine
          Else
            sProcCode.Append _
              "        SET @filterResult = 1" & vbNewLine
          End If
      
          sProcCode.Append _
            "        IF @filterResult = 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SET @loopCounter = @loopCounter + 1" & vbNewLine
      
          Select Case miSelectionType
            Case giSELECT_RECORDCOUNT
              sProcCode.Append _
                "            SET @" & sVarName & " = @loopCounter" & vbNewLine
                
            Case giSELECT_RECORDTOTAL
              sProcCode.Append _
                "            IF NOT " & sLocalResult & " IS NULL SET @" & sVarName & " = CONVERT(float, @" & sVarName & " + " & sLocalResult & ")" & vbNewLine
                
            Case giSELECT_FIRSTRECORD
              sProcCode.Append _
                "            " & sConvertCode & vbNewLine & _
                "            BREAK" & vbNewLine
                
            Case giSELECT_LASTRECORD
              sProcCode.Append _
                "            " & sConvertCode & vbNewLine
                
            Case giSELECT_SPECIFICRECORD
              sProcCode.Append _
                "            " & sConvertCode & vbNewLine & _
                "            IF @loopCounter = " & Trim$(Str$(mlngSelectionLine)) & " BREAK" & vbNewLine & _
                "            ELSE " & sIfNullCode & vbNewLine
          End Select
      
          ' Add the code for closing the cursor.
          sProcCode.Append _
            "        END" & vbNewLine & _
            "        " & sFetchNextCode & vbNewLine & _
            "    END" & vbNewLine & _
            "    CLOSE " & sVarName & "_cursor" & vbNewLine & _
            "    DEALLOCATE " & sVarName & "_cursor" & vbNewLine
        End If
      End If
    End If
  End If

TidyUpAndExit:
  If fOK Then
    StoredProcedureCode = sProcCode.ToString
  Else
    StoredProcedureCode = ""
  End If

  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated value of the Field component.
End Function



Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the calculation's return type.
  On Error GoTo ErrorTrap

  Dim iType As ExpressionValueTypes
  Dim objColumn As Column
  Dim fValidIdentification As Boolean
  Dim asItems() As String
  Dim aWFPrecedingElements() As VB.Control
  Dim aWFAllElements() As VB.Control
  Dim wfTemp As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim lngLoop3 As Long
  Dim sSQL As String
  Dim rsInfo As dao.Recordset
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  
  ' Initialize the default return type.
  iType = giEXPRVALUE_UNDEFINED
  fValidIdentification = True
  
  If miRecordSelection = giWFRECSEL_IDENTIFIEDRECORD Then
    
    ' Check that the identifiying element/item is still valid.
    aWFPrecedingElements = mobjBaseComponent.ParentExpression.PrecedingWorkflowElements
    aWFAllElements = mobjBaseComponent.ParentExpression.AllWorkflowElements

    If UBound(aWFAllElements) >= 0 Then
      fValidIdentification = False
      
      For iLoop = 2 To UBound(aWFPrecedingElements)
        If UCase(Trim(aWFPrecedingElements(iLoop).Identifier)) = UCase(Trim(msElementIdentifier)) Then
          Set wfTemp = aWFPrecedingElements(iLoop)
          
          ReDim alngValidTables(0)
          
          If wfTemp.ElementType = elem_StoredData Then
            TableAscendants wfTemp.DataTableID, alngValidTables
            
          ElseIf wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items

            For iLoop2 = 1 To UBound(asItems, 2)
              If UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(msItemIdentifier)) _
                And asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then

                TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                Exit For
              End If
            Next iLoop2
          End If

          fFound = False
          For lngLoop3 = 1 To UBound(alngValidTables)
            If (IsChildOfTable(alngValidTables(lngLoop3), mlngTableID)) _
              Or (alngValidTables(lngLoop3) = mlngTableID) Then
              
              fFound = True
              Exit For
            End If
          Next lngLoop3

          If fFound Then
            fValidIdentification = True
          End If

          Exit For
        End If
      Next iLoop
    End If
  End If
  
  If fValidIdentification Then
    ' If the component returns the record count then
    ' the return type must be numeric; otherwise the
    ' return type is determined by the field type.
    If miSelectionType = giSELECT_RECORDCOUNT Then
      iType = giEXPRVALUE_NUMERIC
    Else
      ' Determine the field's type by creating an
      ' instance of the column class, and instructing
      ' it to read its own details (including type).
      Set objColumn = New HRProSystemMgr.Column
      objColumn.ColumnID = mlngColumnID
      objColumn.TableID = mlngTableID
  
      If objColumn.ReadColumn Then
        iType = objColumn.ColumnDataType
      End If
    End If
  End If

TidyUpAndExit:
  Set objColumn = Nothing
  ReturnType = iType
  Exit Property

ErrorTrap:
  iType = giEXPRVALUE_UNDEFINED
  Err = False
  Resume TidyUpAndExit
    
End Property






Public Function ConstructComponent() As Boolean
  ' Read the Field component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  With recCompEdit
    ' Read the Field component record.
    miSelectionType = !fieldSelectionRecord
    mlngTableID = !fieldTableID
    mlngColumnID = !fieldColumnID

    miRecordSelection = IIf(IsNull(!WorkflowRecord), giWFRECSEL_UNKNOWN, !WorkflowRecord)
    msElementIdentifier = IIf(IsNull(!WorkflowElement), "", !WorkflowElement)
    msItemIdentifier = IIf(IsNull(!WorkflowItem), "", !WorkflowItem)
    mlngRecordTableID = IIf(IsNull(!WorkflowRecordTableID), 0, !WorkflowRecordTableID)

    mlngSelectionLine = !fieldSelectionLine
    mlngSelOrderID = !fieldSelectionOrderID
    mlngSelFilterID = !FieldSelectionFilter
  End With
   
TidyUpAndExit:
  ConstructComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function
Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
    
  fOK = True
  
  With recCompEdit
    'Add a record for the new component.
    .AddNew
    
    !ComponentID = mobjBaseComponent.ComponentID
    !ExprID = mobjBaseComponent.ParentExpression.ExpressionID
    !Type = giCOMPONENT_WORKFLOWFIELD
    
    !fieldSelectionRecord = miSelectionType
    !fieldTableID = mlngTableID
    !fieldColumnID = mlngColumnID
    
    !WorkflowRecord = miRecordSelection
    !WorkflowElement = msElementIdentifier
    !WorkflowItem = msItemIdentifier
    !WorkflowRecordTableID = mlngRecordTableID
    
    !fieldSelectionLine = mlngSelectionLine
    !fieldSelectionOrderID = mlngSelOrderID
    !FieldSelectionFilter = mlngSelFilterID

    .Update
  End With

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function

Public Property Get TableID() As Long
  ' Return the table id property.
  TableID = mlngTableID
  
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  ' Set the table id property.
  mlngTableID = plngNewValue
  
End Property

Public Property Get ColumnID() As Long
  ' Return the column id property.
  ColumnID = mlngColumnID

End Property

Public Property Get BaseComponent() As CExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property


Public Property Set BaseComponent(ByVal pobjNewValue As CExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  ' Set the column id property.
  mlngColumnID = plngNewValue

End Property

Public Property Get SelectionType() As FieldSelectionTypes
  ' Return the selection type.
  SelectionType = miSelectionType
  
End Property

Public Property Let SelectionType(ByVal piNewValue As FieldSelectionTypes)
  ' Set the selection type.
  miSelectionType = piNewValue

End Property

Public Property Get SelectionLine() As Long
  ' Return the record slection line property.
  SelectionLine = mlngSelectionLine
  
End Property

Public Property Let SelectionLine(ByVal plngNewValue As Long)
  ' Set the record slection line property.
  mlngSelectionLine = plngNewValue

End Property


Public Property Get SelectionOrderID() As Long
  ' Return the Selection Order property value.
  SelectionOrderID = mlngSelOrderID

End Property

Public Property Let SelectionOrderID(ByVal plngNewValue As Long)
  ' Set the Selection Order property value.
  mlngSelOrderID = plngNewValue

End Property




Public Property Get SelectionFilter() As Long
  ' Return the Selection Filter property value.
  SelectionFilter = mlngSelFilterID

End Property

Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
  ' Return the SQL code for the component.
  ' Used for View Filter and Filter expressions.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim fParentField As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sCode As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim sParentTableName As String
  Dim sOrderCode As String
  Dim sFilterCode As String
  Dim rsInfo As dao.Recordset
  Dim avOrderJoinTables() As Variant
  Dim objFilterExpr As CExpression
  Dim sVarName As String
  Dim lngPersonnelTableID As Long
  Dim lngElementID As Long
  Dim iElementType As ElementType
  Dim rsTemp As dao.Recordset
  Dim lngColumnSize As Long
  Dim iDataType As DataTypes
  Dim iColumnDataType As Integer
  Dim objColumn As Column

    ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If

  sCode = ""
  fOK = True
  sVarName = "@component" & Trim$(Str$(mobjBaseComponent.ComponentID))

  psPrefix = psPrefix & _
    "    DECLARE " & sVarName & "_ID integer" & vbNewLine
  
  lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  
  Select Case miRecordSelection
    Case giWFRECSEL_INITIATOR ' 0 - Initiator's personnel table record
      psPrefix = psPrefix & _
        "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
        "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
        "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
        "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
        "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
        "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
        "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
        "       @iBaseTableID = " & CStr(lngPersonnelTableID) & vbNewLine & _
        "    FROM ASRSysWorkflowInstances" & vbNewLine & _
        "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
    
    Case giWFRECSEL_IDENTIFIEDRECORD ' 1 = Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
    
      sSQL = "SELECT TOP 1 tmpWorkflowElements.ID," & _
        "   tmpWorkflowElements.type" & _
        " FROM tmpWorkflowElements" & _
        " WHERE tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & _
        "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'"
    
      Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      With rsTemp
        If Not (.BOF And .EOF) Then
          lngElementID = !id
          iElementType = !Type
        End If
        
        .Close
      End With
      Set rsTemp = Nothing

      If iElementType = elem_WebForm Then
        ' WebForm
        psPrefix = psPrefix & _
          "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
          "    SELECT @iBaseRecordID = convert(integer, ISNULL(IV.value, '0'))," & vbNewLine & _
          "       @iWorkflowID = Es.workflowID," & vbNewLine & _
          "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
          "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
          "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
          "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
          "       @iBaseTableID = EI.tableID" & vbNewLine & _
          "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
          "    INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier" & vbNewLine & _
          "    INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID" & vbNewLine & _
          "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
          "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = '" & Replace(UCase(Trim(msItemIdentifier)), "'", "''") & "'" & vbNewLine & _
          "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'" & vbNewLine & _
          "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
          "        AND Es.ID = IV.elementID" & vbNewLine & vbNewLine
      Else
        ' StoredData
        psPrefix = psPrefix & _
          "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
          "    SELECT @iBaseRecordID = convert(integer, ISNULL(IV.value, '0'))," & vbNewLine & _
          "       @iWorkflowID = Es.workflowID," & vbNewLine & _
          "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
          "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
          "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
          "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
          "       @iBaseTableID = isnull(Es.dataTableID, 0)" & vbNewLine & _
          "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
          "    INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID" & vbNewLine & _
          "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = UPPER(LTRIM(RTRIM(Es.identifier)))" & vbNewLine & _
          "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
          "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msElementIdentifier)), "'", "''") & "'" & vbNewLine & _
          "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & vbNewLine
      End If
    
    Case giWFRECSEL_TRIGGEREDRECORD ' 4 = Triggered Base table record
      psPrefix = psPrefix & _
        "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
        "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
        "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
        "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
        "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
        "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
        "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
        "       @iBaseTableID = isnull(ASRSysWorkflows.baseTable, 0)" & vbNewLine & _
        "    FROM ASRSysWorkflowInstances" & vbNewLine & _
        "    INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
        "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
  End Select

  If (miRecordSelection = giWFRECSEL_INITIATOR) _
    Or (miRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
    Or (miRecordSelection = giWFRECSEL_TRIGGEREDRECORD) Then

    psPrefix = psPrefix & vbNewLine & vbNewLine & _
      "    SET " & sVarName & "_ID = dbo.udf_ASRWorkflowAscendantRecordID(" & vbNewLine & _
      "        @iBaseTableID," & vbNewLine & _
      "        @iBaseRecordID," & vbNewLine & _
      "        @iParent1TableID," & vbNewLine & _
      "        @iParent1RecordID," & vbNewLine & _
      "        @iParent2TableID," & vbNewLine & _
      "        @iParent2RecordID," & vbNewLine & _
      "        " & CStr(mlngRecordTableID) & ")" & vbNewLine & vbNewLine & _
      "    SET @fValidRecordID = 0" & vbNewLine & vbNewLine & _
      "    IF " & sVarName & "_ID > 0" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @fValidRecordID = dbo.udf_ASRWorkflowValidTableRecord(" & vbNewLine & _
      "            " & CStr(mlngRecordTableID) & "," & vbNewLine & _
      "            " & sVarName & "_ID)" & vbNewLine & _
      "    END" & vbNewLine & vbNewLine
  End If

  ' Get the column and table names.
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", mlngTableID

    fOK = Not .NoMatch
    If fOK Then
      sTableName = !TableName

      ' Get the name of the expressions parent table.
      ' ie. the table of the column this is being calculated.
      .Seek "=", mobjBaseComponent.ParentExpression.BaseTableID

      fOK = Not .NoMatch
      If fOK Then
        sParentTableName = !TableName
      End If
    End If
  End With

  ' Get the column name and data type.
  If fOK And (miSelectionType <> giSELECT_RECORDCOUNT) Then
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", mlngColumnID

      If Not .NoMatch Then
        sColumnName = !ColumnName
        iColumnDataType = !DataType

        Set objColumn = New Column
        With objColumn
          .ColumnID = mlngColumnID
          .TableID = mlngTableID
          .ReadColumn
          iDataType = .Properties("dataType")
          lngColumnSize = .Properties("size")
        End With
        Set objColumn = Nothing
      Else
        fOK = False
      End If
    End With
  End If
  
  If fOK Then
    If mlngRecordTableID = mlngTableID Then
      ' The field is in the expression's base table.

      'JPD20010903 Fault 1159
      ' Use dummy values instead of the field names when validating
      ' as the field may not yet be saved to the SQL Server.
      If pfValidating Then
        Select Case ReturnType
          Case giEXPRVALUE_NUMERIC, giEXPRVALUE_LOGIC
            sCode = "1"
          Case giEXPRVALUE_DATE
            sCode = "'01/01/2000'"
          Case Else
            sCode = "'abc'"
        End Select
      Else
        Select Case iDataType
          Case dtVARCHAR, dtLONGVARCHAR
            psPrefix = psPrefix & _
              "    DECLARE " & sVarName & " varchar(MAX)" & vbNewLine & vbNewLine
            
          Case dtNUMERIC, dtINTEGER
            psPrefix = psPrefix & _
              "    DECLARE " & sVarName & " float" & vbNewLine & vbNewLine
            
          Case dtBIT
            psPrefix = psPrefix & _
              "    DECLARE " & sVarName & " bit" & vbNewLine & vbNewLine
            
          Case dtTIMESTAMP
            psPrefix = psPrefix & _
              "    DECLARE " & sVarName & " datetime" & vbNewLine & vbNewLine
            
          Case Else
            fOK = False
        End Select
        
        psPrefix = psPrefix & _
          "    IF @fValidRecordID = 0" & vbNewLine & _
          "    BEGIN" & vbNewLine & _
          "        SET @sValue = null" & vbNewLine & vbNewLine
        
        If miRecordSelection = giWFRECSEL_TRIGGEREDRECORD Then
          psPrefix = psPrefix & _
            "        SELECT @iCount = COUNT(*)" & vbNewLine & _
            "        FROM ASRSysWorkflowQueueColumns QC" & vbNewLine & _
            "        INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID" & vbNewLine & _
            "        WHERE WFQ.instanceID = @piInstanceID" & vbNewLine & _
            "            AND QC.columnID = " & CStr(mlngColumnID) & vbNewLine & vbNewLine & _
            "        IF @iCount = 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '')))" & vbNewLine & _
            "            FROM ASRSysWorkflowQueueColumns QC" & vbNewLine & _
            "            INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID" & vbNewLine & _
            "            WHERE WFQ.instanceID = @piInstanceID" & vbNewLine & _
            "                AND QC.columnID = " & CStr(mlngColumnID) & vbNewLine & vbNewLine & _
            "            SET @fValidRecordID = 1" & vbNewLine & _
            "            --SET @fDeletedValue = 1" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        ElseIf miRecordSelection = giWFRECSEL_IDENTIFIEDRECORD Then
          psPrefix = psPrefix & _
            "        SELECT @iCount = COUNT(*)" & vbNewLine & _
            "        FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
            "        WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
            "            AND IV.columnID = " & CStr(mlngColumnID) & vbNewLine & _
            "            AND IV.elementID = " & CStr(lngElementID) & vbNewLine & vbNewLine & _
            "        IF @iCount = 1" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SELECT @sValue = rtrim(ltrim(isnull(IV.value , '')))" & vbNewLine & _
            "            FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
            "            WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
            "                AND IV.columnID = " & CStr(mlngColumnID) & vbNewLine & _
            "                AND IV.elementID = " & CStr(lngElementID) & vbNewLine & vbNewLine & _
            "            SET @fValidRecordID = 1" & vbNewLine & _
            "            --SET @fDeletedValue = 1" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        End If
        
        Select Case iDataType
          Case dtVARCHAR, dtLONGVARCHAR
            psPrefix = psPrefix & _
              "        SET " & sVarName & " = isnull(@sValue, '')" & vbNewLine

          Case dtNUMERIC, dtINTEGER
            psPrefix = psPrefix & _
              "        SET " & sVarName & " = convert(float, isnull(@sValue, 0))" & vbNewLine

          Case dtBIT
            psPrefix = psPrefix & _
              "        SET " & sVarName & " = convert(bit, isnull(@sValue, 0))" & vbNewLine

          Case dtTIMESTAMP
            psPrefix = psPrefix & _
              "        SET " & sVarName & " = convert(datetime, @sValue)" & vbNewLine
        End Select
        
        psPrefix = psPrefix & _
          "    END" & vbNewLine & _
          "    ELSE" & vbNewLine & _
          "    BEGIN" & vbNewLine
        
        Select Case iDataType
          Case dtVARCHAR, dtLONGVARCHAR
            psPrefix = psPrefix & _
              "        SELECT " & sVarName & " = IsNull(" & sTableName & "." & sColumnName & ",'')" & vbNewLine & _
              "        FROM " & sTableName & vbNewLine & _
              "        WHERE id = " & sVarName & "_ID" & vbNewLine
            
          Case dtNUMERIC, dtINTEGER
            psPrefix = psPrefix & _
              "        SELECT " & sVarName & " = convert(float," & sTableName & "." & sColumnName & ")" & vbNewLine & _
              "        FROM " & sTableName & vbNewLine & _
              "        WHERE id = " & sVarName & "_ID" & vbNewLine
            
          Case dtBIT
            psPrefix = psPrefix & _
              "        SELECT " & sVarName & " = " & sTableName & "." & sColumnName & vbNewLine & _
              "        FROM " & sTableName & vbNewLine & _
              "        WHERE id = " & sVarName & "_ID" & vbNewLine
            
          Case dtTIMESTAMP
            psPrefix = psPrefix & _
              "        SELECT " & sVarName & " = " & "convert(datetime,convert(varchar(20)," & sTableName & "." & sColumnName & ",101))" & vbNewLine & _
              "        FROM " & sTableName & vbNewLine & _
              "        WHERE id = " & sVarName & "_ID" & vbNewLine
            
          Case Else
            fOK = False
        End Select
          
        psPrefix = psPrefix & _
          "    END" & vbNewLine & vbNewLine
          
        sCode = sVarName
      End If
    Else
      ' Check if the table is a child or parent of the expression's base table.
      With recRelEdit
        .Index = "idxParentID"
        .Seek "=", mlngTableID, mlngRecordTableID

        fParentField = Not .NoMatch
      End With

      If fParentField Then
        ' The field is from a parent table of the expression's base table.

        'JPD20010903 Fault 1159
        ' Use dummy values instead of the field names when validating
        ' as the field may not yet be saved to the SQL Server.
        If pfValidating Then
          Select Case ReturnType
            Case giEXPRVALUE_NUMERIC, giEXPRVALUE_LOGIC
              sCode = "1"
            Case giEXPRVALUE_DATE
              sCode = "'01/01/2000'"
            Case Else
              sCode = "'abc'"
          End Select
        Else
          Select Case iDataType
            Case dtVARCHAR, dtLONGVARCHAR
              psPrefix = psPrefix & _
                "DECLARE " & sVarName & " varchar(MAX)" & vbNewLine & vbNewLine & _
                "SELECT " & sVarName & " = " & _
                "IsNull(" & sTableName & "." & sColumnName & ",'')" & vbNewLine
              
            Case dtNUMERIC, dtINTEGER
              psPrefix = psPrefix & _
                "DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
                "SELECT " & sVarName & " = " & _
                "convert(float," & sTableName & "." & sColumnName & ")" & vbNewLine
              
            Case dtBIT
              psPrefix = psPrefix & _
                "DECLARE " & sVarName & " bit" & vbNewLine & vbNewLine & _
                "SELECT " & sVarName & " = " & _
                sTableName & "." & sColumnName & vbNewLine
              
            Case dtTIMESTAMP
              psPrefix = psPrefix & _
                "DECLARE " & sVarName & " datetime" & vbNewLine & vbNewLine & _
                "SELECT " & sVarName & " = " & _
                "convert(datetime,convert(varchar(20)," & sTableName & "." & sColumnName & ",101))" & vbNewLine
              
            Case Else
              fOK = False
          End Select
            
          psPrefix = psPrefix & _
            "FROM " & sTableName & vbNewLine & _
            "WHERE id = " & sVarName & "_ID" & vbNewLine & vbNewLine
          sCode = sVarName
        End If
      Else
        ' The field is from a child table of the expression's base table.

        'JPD20010903 Fault 1159
        ' Use dummy values instead of the field names when validating
        ' as the field may not yet be saved to the SQL Server.
        If pfValidating Then
          Select Case ReturnType
            Case giEXPRVALUE_NUMERIC, giEXPRVALUE_LOGIC
              sCode = "1"
            Case giEXPRVALUE_DATE
              sCode = "'01/01/2000'"
            Case Else
              sCode = "'abc'"
          End Select
        Else
          ' Construct the order code if required.
          ' Create an array of tables that need to be joined to make the order valid.
          ' Column 1 = table ID.
          ' Column 2 = table name.
          sOrderCode = ""
          ReDim avOrderJoinTables(2, 0)
          If mlngSelOrderID > 0 Then
            sSQL = "SELECT tmpColumns.columnName, tmpColumns.columnID, tmpColumns.tableID, tmpTables.tableName, tmpOrderItems.ascending" & _
              " FROM tmpOrderItems, tmpColumns, tmpTables" & _
              " WHERE orderID = " & Trim$(Str(mlngSelOrderID)) & _
              " AND type = 'O'" & _
              " AND tmpColumns.columnID = tmpOrderItems.columnID" & _
              " AND tmpColumns.tableID = tmpTables.tableID" & _
              " ORDER BY sequence"

            Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
            With rsInfo
              Do While Not .EOF
                ' Construct the order code. Remember that if we are selecting the last record,
                ' we must reverse the ASC/DESC options.
                sOrderCode = sOrderCode & IIf(LenB(sOrderCode) <> 0, ", ", "") & _
                  !TableName & "." & !ColumnName & _
                  IIf(miSelectionType = giSELECT_LASTRECORD, IIf(!Ascending, " DESC", ""), IIf(!Ascending, "", " DESC"))

                If (!TableID <> mlngTableID) Then

                  ' Check if the table has already been added to the array of tables used in the order.
                  fFound = False
                  For iNextIndex = 1 To UBound(avOrderJoinTables, 2)
                    If avOrderJoinTables(1, iNextIndex) = !TableID And _
                      (avOrderJoinTables(2, iNextIndex) = !TableName) Then

                      fFound = True
                      Exit For
                    End If
                  Next iNextIndex

                  If Not fFound Then
                    iNextIndex = UBound(avOrderJoinTables, 2) + 1
                    ReDim Preserve avOrderJoinTables(2, iNextIndex)
                    avOrderJoinTables(1, iNextIndex) = !TableID
                    avOrderJoinTables(2, iNextIndex) = !TableName
                  End If
                End If

                .MoveNext
              Loop

              .Close
            End With
            Set rsInfo = Nothing
          End If

          If fOK Then
            Select Case miSelectionType
              Case giSELECT_FIRSTRECORD, giSELECT_LASTRECORD
                ' First and Last record selection uses the same code here.
                ' The difference is made when creating the 'order by' code above.
                Select Case iDataType
                  Case dtVARCHAR, dtLONGVARCHAR
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " varchar(MAX)" & vbNewLine & vbNewLine & _
                      "SELECT TOP 1 " & sVarName & " = " & _
                      "IsNull(" & sTableName & "." & sColumnName & ",'')" & vbNewLine
                    
                  Case dtNUMERIC, dtINTEGER
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
                      "SELECT TOP 1 " & sVarName & " = " & _
                      "convert(float," & sTableName & "." & sColumnName & ")" & vbNewLine
                    
                  Case dtBIT
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " bit" & vbNewLine & vbNewLine & _
                      "SELECT TOP 1 " & sVarName & " = " & _
                      sTableName & "." & sColumnName & vbNewLine
                    
                  Case dtTIMESTAMP
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " datetime" & vbNewLine & vbNewLine & _
                      "SELECT TOP 1 " & sVarName & " = " & _
                      "convert(datetime,convert(varchar(20)," & sTableName & "." & sColumnName & ",101))" & vbNewLine
                    
                  Case Else
                    fOK = False
                End Select
                  
                psPrefix = psPrefix & _
                  "FROM " & sTableName & vbNewLine
                
                sCode = sVarName
                    
                ' Add the JOIN code for the order.
                For iLoop = 1 To UBound(avOrderJoinTables, 2)
                  psPrefix = psPrefix & _
                    "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & sTableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
                Next iLoop

                psPrefix = psPrefix & _
                  "WHERE " & sTableName & ".id_" & CStr(mlngRecordTableID) & " = " & sVarName & "_ID" & vbNewLine

                ' Add the filter code as required.
                If mlngSelFilterID > 0 Then
                  psPrefix = psPrefix & _
                    "AND " & sTableName & ".id IN" & vbNewLine & _
                    "(SELECT id FROM dbo.udf_ASRWFExpr_" & CStr(mlngSelFilterID) & " (@piInstanceID))" & vbNewLine
                End If

                ' Add the order code as required.
                If mlngSelOrderID > 0 Then
                  psPrefix = psPrefix & _
                    "ORDER BY " & sOrderCode & vbNewLine
                End If

              Case giSELECT_RECORDCOUNT
                psPrefix = psPrefix & _
                  "DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
                  "SELECT " & sVarName & " = COUNT(" & sTableName & ".id)" & vbNewLine & _
                  "FROM " & sTableName & vbNewLine & _
                  "WHERE " & sTableName & ".id_" & CStr(mlngRecordTableID) & " = " & sVarName & "_ID" & vbNewLine

                ' Add the filter code as required.
                If mlngSelFilterID > 0 Then
                  psPrefix = psPrefix & _
                    "AND " & sTableName & ".id IN" & vbNewLine & _
                    "(SELECT id FROM dbo.udf_ASRWFExpr_" & CStr(mlngSelFilterID) & " (@piInstanceID))" & vbNewLine
                End If
                
                sCode = sVarName
              
              Case giSELECT_RECORDTOTAL
                  ' No need to add the order code as it makes no differnt when selecting the record total.
                  psPrefix = psPrefix & _
                    "DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
                    "SELECT " & sVarName & " = isnull(SUM(" & sTableName & "." & sColumnName & "), 0)" & vbNewLine & _
                    "FROM " & sTableName & vbNewLine & _
                    "WHERE " & sTableName & ".id_" & CStr(mlngRecordTableID) & " = " & sVarName & "_ID" & vbNewLine

                  ' Add the filter code as required.
                  If mlngSelFilterID > 0 Then
                    psPrefix = psPrefix & _
                      "AND " & sTableName & ".id IN" & vbNewLine & _
                      "(SELECT id FROM dbo.udf_ASRWFExpr_" & CStr(mlngSelFilterID) & " (@piInstanceID))" & vbNewLine
                  End If
                  
                  sCode = sVarName

              Case giSELECT_SPECIFICRECORD
                Select Case iDataType
                  Case dtVARCHAR, dtLONGVARCHAR
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " varchar(MAX)" & vbNewLine & vbNewLine & _
                      "DECLARE GetRecord CURSOR SCROLL FOR" & vbNewLine & _
                      "SELECT IsNull(" & sTableName & "." & sColumnName & ",'')" & vbNewLine
                    
                  Case dtNUMERIC, dtINTEGER
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " float" & vbNewLine & vbNewLine & _
                      "DECLARE GetRecord CURSOR SCROLL FOR" & vbNewLine & _
                      "SELECT convert(float," & sTableName & "." & sColumnName & ")" & vbNewLine
                    
                  Case dtBIT
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " bit" & vbNewLine & vbNewLine & _
                      "DECLARE GetRecord CURSOR SCROLL FOR" & vbNewLine & _
                      "SELECT " & sTableName & "." & sColumnName & vbNewLine
                    
                  Case dtTIMESTAMP
                    psPrefix = psPrefix & _
                      "DECLARE " & sVarName & " datetime" & vbNewLine & vbNewLine & _
                      "DECLARE GetRecord CURSOR SCROLL FOR" & vbNewLine & _
                      "SELECT convert(datetime,convert(varchar(20)," & sTableName & "." & sColumnName & ",101))" & vbNewLine
                    
                  Case Else
                    fOK = False
                End Select
                
                psPrefix = psPrefix & _
                  "FROM " & sTableName & vbNewLine
                
                sCode = sVarName
                
                ' Add the JOIN code for the order.
                For iLoop = 1 To UBound(avOrderJoinTables, 2)
                  psPrefix = psPrefix & _
                    "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & sTableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
                Next iLoop
                
                psPrefix = psPrefix & _
                  "WHERE " & sTableName & ".id_" & CStr(mlngRecordTableID) & " = " & sVarName & "_ID" & vbNewLine
                
                ' Add the filter code as required.
                If mlngSelFilterID > 0 Then
                  psPrefix = psPrefix & _
                    "AND " & sTableName & ".id IN" & vbNewLine & _
                    "(SELECT id FROM dbo.udf_ASRWFExpr_" & CStr(mlngSelFilterID) & " (@piInstanceID))" & vbNewLine
                End If
                
                ' Add the order code as required.
                If mlngSelOrderID > 0 Then
                  psPrefix = psPrefix & _
                    "ORDER BY " & sOrderCode & vbNewLine
                End If
                
                
                psPrefix = psPrefix & _
                  "OPEN GetRecord" & vbNewLine & _
                  "FETCH ABSOLUTE " & Trim(Str(mlngSelectionLine)) & " FROM GetRecord INTO " & sVarName & vbNewLine & _
                  "CLOSE GetRecord" & vbNewLine & _
                  "DEALLOCATE GetRecord" & vbNewLine & vbNewLine

              Case Else
                ' Unrecognised child record selection option.
                fOK = False
            End Select
          End If
        End If
      End If
    End If
  End If

TidyUpAndExit:
  If fOK Then
    psRuntimeCode = sCode
  Else
    psRuntimeCode = ""
  End If

  RuntimeCode = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Property Let SelectionFilter(ByVal plngNewValue As Long)
  ' Set the Selection Filter property value.
  mlngSelFilterID = plngNewValue

End Property

Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the component type.
  ComponentType = giCOMPONENT_WORKFLOWFIELD
  
End Property

Public Property Get ComponentDescription() As String
  ' Return a description of the field component.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sTableName As String
  Dim sColumnName As String
  Dim sSelectionType As String

  fOK = True

  ' Get the table name.
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", mlngTableID

    If Not .NoMatch Then
      sTableName = !TableName
    Else
      sTableName = "<unknown>"
    End If
  End With

  ' Get the column name.
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", mlngColumnID

    If Not .NoMatch Then
      sColumnName = .Fields("columnName")
    Else
      sColumnName = "<unknown>"
    End If
  End With

  If fOK Then
    ' Only give the full description if the field is in a child table of the
    ' expression's parent table.
    With recRelEdit
      .Index = "idxParentID"
      .Seek "=", mlngRecordTableID, mlngTableID

      If Not .NoMatch Then

        Select Case miSelectionType
          Case giSELECT_FIRSTRECORD
            sSelectionType = " (first record"
          Case giSELECT_LASTRECORD
            sSelectionType = " (last record"
          Case giSELECT_SPECIFICRECORD
            sSelectionType = " (line " & Trim(Str(mlngSelectionLine))
          Case giSELECT_RECORDTOTAL
            sSelectionType = " (total"
          Case giSELECT_RECORDCOUNT
            sSelectionType = " (record count"
          Case Else
            sSelectionType = " <unknown>"
        End Select

        If mlngSelOrderID > 0 Then
          ' Get the order name.
          With recOrdEdit
            .Index = "idxID"
            .Seek "=", mlngSelOrderID

            If Not .NoMatch Then
              sSelectionType = sSelectionType & ", order by '" & !Name & "'"
            End If
          End With
        End If

        If mlngSelFilterID > 0 Then
          ' Get the filter name.
          With recExprEdit
            .Index = "idxExprID"
            .Seek "=", mlngSelFilterID, False

            If Not .NoMatch Then
              sSelectionType = sSelectionType & ", filter by '" & !Name & "'"
            End If
          End With
        End If

        sSelectionType = sSelectionType & ")"
      End If
    End With
  End If

TidyUpAndExit:
  ' Return the component description (to be displayed in the expression treeview).
  If miSelectionType = giSELECT_RECORDCOUNT Then
    ComponentDescription = sTableName & " " & sSelectionType
  Else
    ComponentDescription = sTableName & " : " & sColumnName & " " & sSelectionType
  End If
  Exit Property

ErrorTrap:
  sTableName = "<unknown>"
  sColumnName = "<unknown>"
  sSelectionType = "<unknown>"
  fOK = False
  Resume TidyUpAndExit

End Property


Public Function ComponentContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the component makes any reference to the given column.
  ' Used to check for recursion.
  Dim objFilterExpr As CExpression

  ' Check if the column is the component's field.
  ComponentContainsColumn = (mlngColumnID = plngColumnID)

  If Not ComponentContainsColumn Then
    ' Check if the column is used in the component's child filter.
    If mlngSelFilterID > 0 Then
      Set objFilterExpr = New CExpression
      objFilterExpr.ExpressionID = mlngSelFilterID
      objFilterExpr.ConstructExpression
      ComponentContainsColumn = objFilterExpr.ExpressionContainsColumn(plngColumnID)
      Set objFilterExpr = Nothing
    End If
  End If
  
End Function

Public Sub ColumnsUsedInThisComponent(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  Dim objFilterExpr As CExpression
  Dim iLoop As Integer
  Dim fFound As Boolean

  fFound = False
  For iLoop = 1 To UBound(pvColumns)
    If pvColumns(iLoop) = mlngColumnID Then
      fFound = True
      Exit For
    End If
  Next iLoop

  If Not fFound Then
    ReDim Preserve pvColumns(UBound(pvColumns) + 1)
    pvColumns(UBound(pvColumns)) = mlngColumnID
  End If

  If mlngSelFilterID > 0 Then
    Set objFilterExpr = New CExpression
    objFilterExpr.ExpressionID = mlngSelFilterID
    objFilterExpr.ConstructExpression
    objFilterExpr.ColumnsUsedInThisExpression pvColumns
    Set objFilterExpr = Nothing
  End If

End Sub



Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  Dim objFilterExpr As CExpression
  Dim lngLoop As Long
  Dim fFound As Boolean
  
  If mlngSelFilterID > 0 Then
    fFound = False
    For lngLoop = 1 To UBound(palngExpressions)
      If palngExpressions(lngLoop) = mlngSelFilterID Then
        fFound = True
        Exit For
      End If
    Next lngLoop
    
    If Not fFound Then
      ReDim Preserve palngExpressions(UBound(palngExpressions) + 1)
      palngExpressions(UBound(palngExpressions)) = mlngSelFilterID
      
      ' Get the expressions used in the component expression.
      Set objFilterExpr = New CExpression
      With objFilterExpr
        .ExpressionID = mlngSelFilterID
        .ConstructExpression
        .ExpressionsUsedInThisExpression palngExpressions
      End With
      Set objFilterExpr = Nothing
    End If
  End If

End Sub




Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objFieldCopy As New CExprWorkflowField
  
  ' Copy the component's basic properties.
  With objFieldCopy
    .SelectionType = miSelectionType
    .TableID = mlngTableID
    .ColumnID = mlngColumnID

    .RecordSelectionType = miRecordSelection
    .WorkflowElement = msElementIdentifier
    .WorkflowItem = msItemIdentifier
    .RecordTableID = mlngRecordTableID
    
    .SelectionLine = mlngSelectionLine
    .SelectionOrderID = mlngSelOrderID
    .SelectionFilter = mlngSelFilterID
  End With
  
  Set CopyComponent = objFieldCopy
   
  ' Disassociate object variables.
  Set objFieldCopy = Nothing
   
End Function

Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngIndex As Long
  Dim lngNewTableID As Long
  Dim lngNewOrderID As Long
  Dim lngNewColumnID As Long
  Dim lngNewRecordTableID As Long
  Dim objFieldClone As New CExprWorkflowField

  lngNewTableID = 0
  lngNewColumnID = 0
  lngNewOrderID = 0
  lngNewRecordTableID = 0

  ' Get the new table, column, and order IDs from the clone register.
  For lngIndex = 1 To UBound(pavCloneRegister, 2)
    If (pavCloneRegister(1, lngIndex) = "TABLE") And _
      (pavCloneRegister(2, lngIndex) = mlngTableID) And _
      (lngNewTableID = 0) Then
      lngNewTableID = pavCloneRegister(3, lngIndex)
    End If

    If (pavCloneRegister(1, lngIndex) = "COLUMN") And _
      (pavCloneRegister(2, lngIndex) = mlngColumnID) And _
      (lngNewColumnID = 0) Then
      lngNewColumnID = pavCloneRegister(3, lngIndex)
    End If

    If (pavCloneRegister(1, lngIndex) = "ORDER") And _
      (pavCloneRegister(2, lngIndex) = mlngSelOrderID) And _
      (lngNewOrderID = 0) Then
      lngNewOrderID = pavCloneRegister(3, lngIndex)
    End If

    If (pavCloneRegister(1, lngIndex)) = "TABLE" And _
      (pavCloneRegister(2, lngIndex) = mlngRecordTableID) And _
      (lngNewRecordTableID = 0) Then
      lngNewRecordTableID = pavCloneRegister(3, lngIndex)
    End If

    If ((mlngTableID = 0) Or (lngNewTableID > 0)) And _
      ((mlngSelOrderID = 0) Or (lngNewOrderID > 0)) And _
      ((mlngColumnID = 0) Or (lngNewColumnID > 0)) And _
      ((mlngRecordTableID = 0) Or (lngNewRecordTableID > 0)) Then
      Exit For
    End If
  Next lngIndex

  ' If the cloned versions of the required table/column/order
  ' have not been found then restore the original references.
  If ((mlngTableID > 0) And (lngNewTableID = 0)) Then
    lngNewTableID = mlngTableID
  End If
  If ((mlngColumnID > 0) And (lngNewColumnID = 0)) Then
    lngNewColumnID = mlngColumnID
  End If
  If ((mlngSelOrderID > 0) And (lngNewOrderID = 0)) Then
    lngNewOrderID = mlngSelOrderID
  End If
  If ((mlngRecordTableID > 0) And (lngNewRecordTableID = 0)) Then
    lngNewRecordTableID = mlngRecordTableID
  End If

  ' Copy the component's basic properties.
  With objFieldClone
    .SelectionType = miSelectionType
    .TableID = lngNewTableID
    .ColumnID = lngNewColumnID

    .RecordSelectionType = miRecordSelection
    .WorkflowElement = msElementIdentifier
    .WorkflowItem = msItemIdentifier
    .RecordTableID = lngNewRecordTableID
    
    .SelectionLine = mlngSelectionLine
    .SelectionOrderID = lngNewOrderID
    .SelectionFilter = mlngSelFilterID
  End With

  fOK = True
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objFieldClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objFieldClone = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String(piLevel, vbTab) & ComponentDescription & vbNewLine
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function





Private Sub Class_Initialize()
  ' Initialize properties.
  msElementIdentifier = ""
  msItemIdentifier = ""

End Sub



Public Property Get RecordSelectionType() As WorkflowRecordSelectorTypes
  RecordSelectionType = miRecordSelection
  
End Property

Public Property Let RecordSelectionType(ByVal piNewValue As WorkflowRecordSelectorTypes)
  miRecordSelection = piNewValue
  
End Property

Public Property Get RecordTableID() As Long
  RecordTableID = mlngRecordTableID
  
End Property

Public Property Let RecordTableID(ByVal plngNewValue As Long)
  mlngRecordTableID = plngNewValue
  
End Property

' Code stub for .NET migration
Public Function GetSQLElementCode() As String
  GetSQLElementCode = vbNullString
End Function
