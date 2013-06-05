VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprField"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngTableID As Long
Private mlngColumnID As Long
Private miFieldPassType As Integer
Private miSelectionType As FieldSelectionTypes
Private mlngSelectionLine As Long
Private mlngSelOrderID As Long
Private mlngSelFilterID As Long

' Class handling variables.
Private mobjBaseComponent As CExprComponent



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
  Dim sProcCode As HRProSystemMgr.cStringBuilder
  Dim sColumnName As String
  Dim sTableName As String
  Dim sParentTableName As String
  Dim sOrderCode As String
  Dim sIfNullCode As String
  Dim sLocalResult As String
  Dim sConvertCode As String
  Dim sFetchNextCode As String
  Dim rsInfo As dao.Recordset
  Dim avOrderJoinTables() As Variant
  Dim objColumn As Column
 
  fOK = True
  sVarName = "component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  sLocalResult = "@" & sVarName & "_localResult "
  Set sProcCode = New HRProSystemMgr.cStringBuilder
  
  ' Get the field component's table name.
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", mlngTableID

    If Not .NoMatch Then
      sTableName = !TableName
      If mobjBaseComponent.ParentExpression.BaseTableID = !TableID And Trim(strBaseTable) <> vbNullString Then
        sTableName = Trim(strBaseTable)
      End If

      If miFieldPassType = giPASSBY_VALUE Then
        ' Get the name of the expressions parent table.
        ' ie. the table of the column this is being calculated.
        .Seek "=", mobjBaseComponent.ParentExpression.BaseTableID
  
        If Not .NoMatch Then
          sParentTableName = !TableName
          If mobjBaseComponent.ParentExpression.BaseTableID = !TableID And Trim$(strBaseTable) <> vbNullString Then
            sParentTableName = Trim$(strBaseTable)
          End If
        Else
          fOK = False
        End If
      End If
    Else
      fOK = False
    End If
  End With
        
  ' Get the column name and data type.
  If fOK Then
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
    If miFieldPassType = giPASSBY_REFERENCE Then
      ' Create the stored procedure code for FIELDS PASSED BY REFERENCE.
      sProcCode.TheString = "    /* Evaluate field (by reference) component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ".*/" & vbNewLine & _
        "    DECLARE @" & sVarName & " integer" & vbNewLine & _
        "    SET @" & sVarName & " = " & Trim$(Str$(mlngColumnID)) & vbNewLine
    Else
      ' Construct the 'by value' field component strored procedure code.
      Select Case iDataType
        Case dtVARCHAR, dtLONGVARCHAR
          sProcCode.TheString = "    /* Evaluate field (by value) component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ".*/" & vbNewLine & _
            "    DECLARE @" & sVarName & " nvarchar(MAX)" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult nvarchar(MAX)" & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = ''"
          sConvertCode = "SET @" & sVarName & " = CONVERT(nvarchar(MAX), " & sLocalResult & ")"
        Case dtNUMERIC, dtINTEGER
          sProcCode.TheString = "    /* Evaluate field (by value) component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ".*/" & vbNewLine & _
            "    DECLARE @" & sVarName & " float" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult float" & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = 0"
          sConvertCode = "SET @" & sVarName & " = CONVERT(float, " & sLocalResult & ")"
        Case dtBIT
          sProcCode.TheString = "    /* Evaluate field (by value) component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ".*/" & vbNewLine & _
            "    DECLARE @" & sVarName & " bit" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult bit" & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = 0"
          sConvertCode = "SET @" & sVarName & " = CONVERT(bit, " & sLocalResult & ")"
        Case dtTIMESTAMP
          sProcCode.TheString = "    /* Evaluate field (by value) component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ".*/" & vbNewLine & _
            "    DECLARE @" & sVarName & " datetime" & vbNewLine & _
            "    DECLARE @" & sVarName & "_localResult datetime" & vbNewLine
          sIfNullCode = "SET @" & sVarName & " = null"
          sConvertCode = "SET @" & sVarName & " = CONVERT(datetime, " & sLocalResult & ")"
        Case Else
          fOK = False
      End Select

      If fOK Then
        If miSelectionType = giSELECT_RECORDTOTAL Then
          sProcCode.Append vbNewLine & _
            "    SET @" & sVarName & " = 0" & vbNewLine
        End If

        ' Determine if the field is in a parent table of the expression's root table,
        ' or in a child table of the root table, or in the root table itself.
        If mlngTableID = mobjBaseComponent.ParentExpression.BaseTableID Then
          sProcCode.Append vbNewLine & _
            "    SELECT @" & sVarName & " = " & sColumnName & vbNewLine & _
            "        FROM " & sTableName & vbNewLine & _
            "        WHERE id = @id" & vbNewLine
        Else
          With recRelEdit
            .Index = "idxParentID"
            .Seek "=", mlngTableID, mobjBaseComponent.ParentExpression.BaseTableID
            If Not .NoMatch Then  ' ie. the field table is the parent of the root table.
              If mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_DEFAULTVALUE Then
                sProcCode.Append vbNewLine & _
                  "    SELECT @" & sVarName & " = " & sTableName & "." & sColumnName & vbNewLine & _
                  "        FROM " & sTableName & vbNewLine & _
                  "        WHERE " & sTableName & ".id = @id_" & Trim$(Str$(mlngTableID)) & vbNewLine
              Else
                sProcCode.Append vbNewLine & _
                  "    SELECT @" & sVarName & " = " & sTableName & "." & sColumnName & vbNewLine & _
                  "        FROM " & sTableName & vbNewLine & _
                  "        JOIN " & sParentTableName & " ON " & sParentTableName & ".id_" & Trim$(Str$(mlngTableID)) & " = " & sTableName & ".id" & vbNewLine & _
                  "        WHERE " & sParentTableName & ".id = @id" & vbNewLine
              End If
            Else
              .Seek "=", mobjBaseComponent.ParentExpression.BaseTableID, mlngTableID
              If Not .NoMatch Then  ' ie. the field table is the child of the root table.

                If miSelectionType = giSELECT_RECORDCOUNT Then
                  sProcCode.TheString = "    /* Evaluate field (by value, count) component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ".*/" & vbNewLine & _
                    "    DECLARE @" & sVarName & " float" & vbNewLine & _
                    "        SET @" & sVarName & " = 0" & vbNewLine & vbNewLine & _
                    "    DECLARE " & sVarName & "_cursor CURSOR LOCAL FAST_FORWARD READ_ONLY" & vbNewLine & _
                    "        FOR SELECT " & sTableName & ".id"
                  
                  sIfNullCode = "SET @" & sVarName & " = 0"
                  sFetchNextCode = "FETCH NEXT FROM " & sVarName & "_cursor INTO @filterRecordID"
                Else
                  sProcCode.Append vbNewLine & _
                    "    DECLARE " & sVarName & "_cursor CURSOR LOCAL FAST_FORWARD READ_ONLY" & vbNewLine & _
                    "        FOR SELECT " & sTableName & ".id, " & sTableName & "." & sColumnName
                  sFetchNextCode = "FETCH NEXT FROM " & sVarName & "_cursor INTO @filterRecordID, " & sLocalResult
                End If

                sProcCode.Append vbNewLine & _
                  "        FROM " & sTableName & vbNewLine
                  
                ' Create the order code, and an array of parent tables that need to be joined for the
                ' order.
                ' Array column 1 = parent table id
                ' Array column 2 = parent table name
                sOrderCode = ""
                ReDim avOrderJoinTables(2, 0)

                If mlngSelOrderID > 0 Then
                  sSQL = "SELECT tmpColumns.columnName, tmpColumns.tableID, tmpTables.tableName, tmpOrderItems.ascending" & _
                    " FROM tmpOrderItems, tmpColumns, tmpTables" & _
                    " WHERE orderID = " & Trim$(Str$(mlngSelOrderID)) & _
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
                    ' JPD 29/8/00 Use LEFT OUTER JOIN now, as it should be.
                    sProcCode.Append "        LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & sTableName & ".ID_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".ID" & vbNewLine
                  Next iLoop
                End If
                
                sProcCode.Append _
                  "        WHERE " & sTableName & ".ID_" & mobjBaseComponent.ParentExpression.BaseTableID & " = @id" & vbNewLine
                
                If LenB(sOrderCode) <> 0 Then
                  sProcCode.Append "        ORDER BY " & sOrderCode & vbNewLine
                End If
                
                If fOK Then
                  ' Add the code for opening the cursor.
                  sProcCode.Append "    OPEN " & sVarName & "_cursor" & vbNewLine & _
                    "    SET @loopCounter = 0 " & vbNewLine & _
                    "    " & sFetchNextCode & vbNewLine & _
                    "    WHILE (@@fetch_status = 0)" & vbNewLine & _
                    "    BEGIN" & vbNewLine
  
                  ' Add the code to handle filters (workgroups).
                  If mlngSelFilterID > 0 Then
                    If mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER Then
                      sProcCode.Append _
                        "        EXEC @hResult = dbo.sp_ASRExpr_" & Trim$(Str$(mlngSelFilterID)) & " @filterResult OUTPUT, @filterRecordID, @piInstanceID, @piTempElement" & vbNewLine & _
                        "        IF @hResult <> 0 SET @filterResult = 0" & vbNewLine
                    Else
                      sProcCode.Append _
                        "        EXEC @hResult = dbo.sp_ASRExpr_" & Trim$(Str$(mlngSelFilterID)) & " @filterResult OUTPUT, @filterRecordID" & vbNewLine & _
                        "        IF @hResult <> 0 SET @filterResult = 0" & vbNewLine
                    End If
                  Else
                    sProcCode.Append "        SET @filterResult = 1" & vbNewLine
                  End If

                  sProcCode.Append _
                    "        IF @filterResult = 1" & vbNewLine & _
                    "        BEGIN" & vbNewLine & _
                    "            SET @loopCounter = @loopCounter + 1" & vbNewLine

                  Select Case miSelectionType
                    Case giSELECT_RECORDCOUNT
                      sProcCode.Append "            SET @" & sVarName & " = @loopCounter" & vbNewLine
                    Case giSELECT_RECORDTOTAL
                      sProcCode.Append "            IF NOT " & sLocalResult & " IS NULL SET @" & sVarName & " = CONVERT(float, @" & sVarName & " + " & sLocalResult & ")" & vbNewLine
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
                  sProcCode.Append "        END" & vbNewLine & _
                    "        " & sFetchNextCode & vbNewLine & _
                    "    END" & vbNewLine & _
                    "    CLOSE " & sVarName & "_cursor" & vbNewLine & _
                    "    DEALLOCATE " & sVarName & "_cursor" & vbNewLine
                End If
              Else
                fOK = False
              End If
            End If
          End With
        End If
        
        ' Add the code for handling the times when no field values are found.
        If LenB(sIfNullCode) <> 0 Then
          sProcCode.Append "    IF @" & sVarName & " IS NULL" & vbNewLine & _
            "    BEGIN" & vbNewLine & _
            "        " & sIfNullCode & vbNewLine & _
            "    END"
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
  
  'Set sProcCode = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated value of the Field component.
'  On Error GoTo ErrorTrap
'
'  Dim fFound As Boolean
'  Dim iLoop As Integer
'  Dim iNextIndex As Integer
'  Dim vResult As Variant
'  Dim vFilterResult As Variant
'  Dim sSQL As String
'  Dim sColumnName As String
'  Dim sTableName As String
'  Dim sParentTableName As String
'  Dim rsEvaluationSet As rdoResultset
'  Dim iColumnDataType As Integer
'  Dim sFromCode As String
'  Dim sWhereCode As String
'  Dim sOrderCode As String
'  Dim sOrderJoinCode As String
'  Dim vDefaultValue As Variant
'  Dim objFilterExpr As CExpression
'  Dim fFilterOK As Boolean
'  Dim iFilterLoop As Integer
'  Dim dtDateResult As Date
'  Dim avOrderJoinTables() As Variant
'  Dim rsInfo As dao.Recordset
'
'  vDefaultValue = vbNullString
'
'  ' Get the table and columns names.
'  With recTabEdit
'    .Index = "idxTableID"
'
'    ' Get the field's table name.
'    .Seek "=", mlngTableID
'
'    If Not .NoMatch Then
'      sTableName = .Fields("tableName")
'    Else
'      sTableName = vbNullString
'    End If
'
'    ' Get the expression's parent table name.
'    .Seek "=", mobjbasecomponent.ParentExpression.BaseTableID
'
'    If Not .NoMatch Then
'      sParentTableName = .Fields("tableName")
'    Else
'      sParentTableName = vbNullString
'    End If
'  End With
'
'  ' Get the column name and type.
'  With recColEdit
'    .Index = "idxColumnID"
'    .Seek "=", mlngColumnID
'
'    If Not .NoMatch Then
'      sColumnName = .Fields("columnName")
'      iColumnDataType = .Fields("dataType")
'    Else
'      sColumnName = vbNullString
'      iColumnDataType = dtVARCHAR
'    End If
'  End With
'
'  ' If the component returns the field by reference then
'  ' return a tab delimited string representing the fields
'  ' values.
'  If miFieldPassType = giPASSBY_REFERENCE Then
'
'    ' Query the required table to get the values of the given column.
'    sSQL = "SELECT " & sTableName & "." & sColumnName & " AS recResult " & _
'       "FROM " & sTableName
'
'    Set rsEvaluationSet = rdoCon.OpenResultset(sSQL, _
'      rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'
'    ' Construct a tab delimited string with the data type as the
'    ' first item, followed by the column values.
'    vResult = iColumnDataType
'    With rsEvaluationSet
'
'      Do While Not .EOF
'        If IsNull(!recResult) Then
'          vResult = vResult & vbTab
'        Else
'          If iColumnDataType = dtTIMESTAMP Then
'            dtDateResult = !recResult
'            vResult = vResult & vbTab & Format(dtDateResult, "MM/dd/yyyy")
'          Else
'            vResult = vResult & vbTab & !recResult
'          End If
'        End If
'
'       .MoveNext
'      Loop
'    End With
'
'  Else
'    ' Determine the default column value.
'    Select Case iColumnDataType
'      Case dtVARCHAR
'        vDefaultValue = ""
'      Case dtTIMESTAMP
'        vDefaultValue = Null
'      Case dtINTEGER, rdTypeNUMERIC
'        vDefaultValue = 0
'      Case dtBIT
'        vDefaultValue = False
'    End Select
'
'    ' Determine if the field is in a parent table of the expression's root table,
'    ' or in a child table of the root table, or in the root table itself.
'    If mlngTableID = mobjbasecomponent.ParentExpression.BaseTableID Then
'      sFromCode = " FROM " & sTableName
'      sWhereCode = " WHERE " & sTableName & ".ID = " & Trim(Str(pLngRecordID))
'    Else
'      With recRelEdit
'        .Index = "idxParentID"
'        .Seek "=", mlngTableID, mobjbasecomponent.ParentExpression.BaseTableID
'
'        If Not .NoMatch Then  ' ie. the field table is the parent of the root table.
'          sFromCode = " FROM " & sTableName & ", " & sParentTableName
'          sWhereCode = " WHERE " & sTableName & ".ID = " & sParentTableName & ".ID_" & mlngTableID & _
'            " AND " & sParentTableName & ".ID = " & Trim(Str(pLngRecordID))
'        Else
'          .Seek "=", mobjbasecomponent.ParentExpression.BaseTableID, mlngTableID
'
'          If Not .NoMatch Then  ' ie. the field table is the child of the root table.
'            sFromCode = " FROM " & sTableName
'            sWhereCode = " WHERE " & sTableName & ".ID_" & mobjbasecomponent.ParentExpression.BaseTableID & " = " & Trim(Str(pLngRecordID))
'          Else
'            sFromCode = ""
'            sWhereCode = ""
'          End If
'        End If
'      End With
'    End If
'
'    ' Evaluate any prompted values in the filter (workgroup) expression.
'    If mlngSelFilterID > 0 Then
'      Set objFilterExpr = New CExpression
'      objFilterExpr.ExpressionID = mlngSelFilterID
'
'      If objFilterExpr.ConstructExpression Then
'        vFilterResult = objFilterExpr.EvaluateExpression(pLngRecordID, pfDoPrompts)
'      End If
'    End If
'
'    ' Create the order code, and an array of parent tables that need to be joined for the
'    ' order.
'    ' Array column 1 = parent table id
'    ' Array column 2 = parent table name
'    sOrderCode = ""
'    ReDim avOrderJoinTables(2, 0)
'
'    If mlngSelOrderID > 0 Then
'      sSQL = "SELECT tmpColumns.columnName, tmpColumns.tableID, tmpTables.tableName, tmpOrderItems.ascending" & _
'        " FROM tmpOrderItems, tmpColumns, tmpTables" & _
'        " WHERE orderID = " & Trim(Str(mlngSelOrderID)) & _
'        " AND type = 'O'" & _
'        " AND tmpColumns.columnID = tmpOrderItems.columnID" & _
'        " AND tmpColumns.tableID = tmpTables.tableID" & _
'        " ORDER BY sequence"
'
'      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
'      With rsInfo
'        Do While Not .EOF
'
'          sOrderCode = sOrderCode & IIf(Len(sOrderCode) > 0, ", ", "") & !TableName & "." & !ColumnName & _
'            IIf(!Ascending, "", " DESC")
'
'          ' Add the table definition to the array of parent tables if required.
'          If mlngTableID <> !TableId Then
'            fFound = False
'
'            For iLoop = 1 To UBound(avOrderJoinTables, 2)
'              If avOrderJoinTables(1, iLoop) = !TableId Then
'                fFound = True
'                Exit For
'              End If
'            Next iLoop
'
'            If Not fFound Then
'              iNextIndex = UBound(avOrderJoinTables, 2) + 1
'              ReDim Preserve avOrderJoinTables(2, iNextIndex)
'              avOrderJoinTables(1, iNextIndex) = !TableId
'              avOrderJoinTables(2, iNextIndex) = !TableName
'            End If
'          End If
'
'          .MoveNext
'        Loop
'
'        .Close
'      End With
'      Set rsInfo = Nothing
'    End If
'
'    sOrderJoinCode = ""
'    If UBound(avOrderJoinTables, 2) > 0 Then
'      For iLoop = 1 To UBound(avOrderJoinTables, 2)
'        sOrderJoinCode = sOrderJoinCode & "        JOIN " & avOrderJoinTables(2, iLoop) & " ON " & avOrderJoinTables(2, iLoop) & ".ID = " & sTableName & ".ID_" & avOrderJoinTables(1, iLoop) & vbnewline
'      Next iLoop
'    End If
'
'    ' Construct the SQL query to get the required unfiltered resultset.
'    sSQL = "SELECT " & sTableName & "." & sColumnName & " AS recResult, " & _
'      sTableName & "." & "ID AS recID " & _
'      sFromCode & _
'      sOrderJoinCode & _
'      sWhereCode & _
'      IIf(sOrderCode = "", "", "  ORDER BY " & sOrderCode)
'
'    Set rsEvaluationSet = rdoCon.OpenResultset(sSQL, _
'      rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'
'    iFilterLoop = 0
'
'    With rsEvaluationSet
'      Do While Not .EOF
'        fFilterOK = True
'
'        If mlngSelFilterID > 0 Then
'          fFilterOK = objFilterExpr.EvaluateExpression(!recID, False)
'        End If
'
'        If fFilterOK Then
'          iFilterLoop = iFilterLoop + 1
'
'          Select Case miSelectionType
'            Case giSELECT_RECORDCOUNT
'              vResult = iFilterLoop
'            Case giSELECT_RECORDTOTAL
'              vResult = vResult + IIf(IsNull(!recResult), vDefaultValue, !recResult)
'            Case giSELECT_FIRSTRECORD
'              vResult = IIf(IsNull(!recResult), vDefaultValue, !recResult)
'              Exit Do
'            Case giSELECT_LASTRECORD
'              vResult = IIf(IsNull(!recResult), vDefaultValue, !recResult)
'            Case giSELECT_SPECIFICRECORD
'              vResult = IIf(IsNull(!recResult), vDefaultValue, !recResult)
'              If iFilterLoop = mlngSelectionLine Then
'                Exit Do
'              End If
'          End Select
'        End If
'
'        .MoveNext
'      Loop
'    End With
'  End If
'
'TidyUpAndExit:
'  Set rsEvaluationSet = Nothing
'  EvaluateComponent = vResult
'  Exit Function
'
'ErrorTrap:
'  vResult = vDefaultValue
'  Resume TidyUpAndExit
'
End Function



Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the calculation's return type.
  On Error GoTo ErrorTrap
  
  Dim iType As ExpressionValueTypes
  Dim objColumn As Column
  
  ' Initialize the default return type.
  iType = giEXPRVALUE_UNDEFINED
  
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

      If miFieldPassType = giPASSBY_REFERENCE Then
        iType = iType + 100
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
    mlngTableID = !fieldTableID
    mlngColumnID = !fieldColumnID
    miFieldPassType = !fieldPassBy
    miSelectionType = !fieldSelectionRecord
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
    !Type = giCOMPONENT_FIELD
    !fieldTableID = mlngTableID
    !fieldColumnID = mlngColumnID
    !fieldPassBy = miFieldPassType
    !fieldSelectionRecord = miSelectionType
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

Public Property Get FieldPassType() As FieldPassTypes
  ' Return the field pass type property.
  FieldPassType = miFieldPassType
  
End Property

Public Property Let FieldPassType(ByVal piNewValue As FieldPassTypes)
  ' Set the field pass type property.
  miFieldPassType = piNewValue

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
  Dim fNewSourceTable As Boolean
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
  Dim fWorkflowExpression As Boolean
  Dim strUDFReturnType As String
  Dim strUDFName As String
  Dim strUDFRuntimeCode As String
  
  fWorkflowExpression = (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)
  sCode = ""
  fOK = True
  
  ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If
  
  If (miFieldPassType = giPASSBY_REFERENCE) Then
    sCode = Trim(Str(mlngColumnID))
  Else
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
  
    ' Get the column name.
    If fOK Then
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", mlngColumnID
      
        fOK = Not .NoMatch
        If fOK Then
          sColumnName = !ColumnName
        End If
      End With
    End If
    
    If fOK Then
      If mobjBaseComponent.ParentExpression.BaseTableID = mlngTableID Then
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
          sCode = sTableName & "." & sColumnName
        End If
      Else
        ' Check if the table is a child or parent of the expression's base table.
        With recRelEdit
          .Index = "idxParentID"
          .Seek "=", mlngTableID, mobjBaseComponent.ParentExpression.BaseTableID
          
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
            sCode = sTableName & "." & sColumnName
          End If
          
          fNewSourceTable = True
          For iLoop = 1 To UBound(palngSourceTables, 2)
            If (palngSourceTables(1, iLoop) = 0) And _
              (palngSourceTables(2, iLoop) = mlngTableID) Then
              fNewSourceTable = False
              Exit For
            End If
          Next iLoop
  
          If fNewSourceTable Then
            iNextIndex = UBound(palngSourceTables, 2) + 1
            ReDim Preserve palngSourceTables(2, iNextIndex)
            palngSourceTables(1, iNextIndex) = 0
            palngSourceTables(2, iNextIndex) = mlngTableID
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
            sCode = "(" & vbNewLine
            
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
                  
                  If (!TableID <> mlngTableID) And _
                    (!TableID <> mobjBaseComponent.ParentExpression.BaseTableID) Then
                      
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
              ' Create the filter code if required.
              sFilterCode = ""
              If mlngSelFilterID > 0 Then
                Set objFilterExpr = New CExpression
                objFilterExpr.ExpressionID = mlngSelFilterID
                objFilterExpr.ConstructExpression
                fOK = objFilterExpr.RuntimeFilterCode(sFilterCode, pfValidating)
                Set objFilterExpr = Nothing
              End If
            End If
            
            If fOK Then
              Select Case miSelectionType
                Case giSELECT_FIRSTRECORD, giSELECT_LASTRECORD
                  ' First and Last record selection uses the same code here.
                  ' The difference is made when creating the 'order by' code above.
                    sCode = sCode & _
                      "SELECT TOP 1 " & sTableName & "." & sColumnName & vbNewLine & _
                      "FROM " & sTableName & vbNewLine
    
                    ' Add the JOIN code for the order.
                    For iLoop = 1 To UBound(avOrderJoinTables, 2)
                      sCode = sCode & _
                        "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & sTableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
                    Next iLoop
                  
                    sCode = sCode & _
                      "WHERE " & sParentTableName & ".id = " & sTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine
                  
                    ' Add the filter code as required.
                    If LenB(sFilterCode) <> 0 Then
                      sCode = sCode & _
                        "AND " & sTableName & ".id IN" & vbNewLine & _
                        "(" & vbNewLine & _
                        sFilterCode & _
                        ")" & vbNewLine
                    End If
    
                    ' Add the order code as required.
                    If mlngSelOrderID > 0 Then
                      sCode = sCode & _
                        "ORDER BY " & sOrderCode & vbNewLine
                    End If
                  
                Case giSELECT_RECORDCOUNT
                  ' No need to add the order code as it makes no differnt when selecting the record count.
                    sCode = sCode & _
                      "SELECT COUNT(" & sTableName & ".id)" & vbNewLine & _
                      "FROM " & sTableName & vbNewLine & _
                      "WHERE " & sParentTableName & ".id = " & sTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine
                  
                    ' Add the filter code as required.
                    If LenB(sFilterCode) <> 0 Then
                      sCode = sCode & _
                        "AND " & sTableName & ".id IN" & vbNewLine & _
                        "(" & vbNewLine & _
                        sFilterCode & _
                        ")" & vbNewLine
                    End If
    
                Case giSELECT_RECORDTOTAL
                  ' No need to add the order code as it makes no differnt when selecting the record total.
                    sCode = sCode & _
                      "SELECT SUM(" & sTableName & "." & sColumnName & ")" & vbNewLine & _
                      "FROM " & sTableName & vbNewLine & _
                      "WHERE " & sParentTableName & ".id = " & sTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbNewLine
                  
                    ' Add the filter code as required.
                    If LenB(sFilterCode) <> 0 Then
                      sCode = sCode & _
                        "AND " & sTableName & ".id IN" & vbNewLine & _
                        "(" & vbNewLine & _
                        sFilterCode & _
                        ")" & vbNewLine
                    End If
                  
                Case giSELECT_SPECIFICRECORD
                  ' Specific cannot be selected for runtime filters.
                  If fWorkflowExpression Then
                    Select Case ReturnType
                      Case giEXPRVALUE_DATE
                        strUDFReturnType = "datetime"
                      
                      Case giEXPRVALUE_CHARACTER
                        strUDFReturnType = "nvarchar(MAX)"
                      
                      Case giEXPRVALUE_NUMERIC
                        strUDFReturnType = "float"
                      
                      Case giEXPRVALUE_LOGIC
                        strUDFReturnType = "bit"
                    End Select
              
                    strUDFName = "udf_ASRWFExpr_" & mobjBaseComponent.ParentExpression.ExpressionID & "_" & CStr(mobjBaseComponent.ComponentID)

                    strUDFRuntimeCode = _
                      "CREATE FUNCTION [dbo].[" & strUDFName & "](@ID integer)" & vbCrLf & _
                      "RETURNS " & strUDFReturnType & vbCrLf & _
                      "AS" & vbCrLf & "BEGIN" & vbCrLf & _
                      "   DECLARE @Result " & strUDFReturnType & vbCrLf & _
                      "   DECLARE GetRecord CURSOR SCROLL FOR "

                    strUDFRuntimeCode = strUDFRuntimeCode & _
                      "SELECT " & sTableName & "." & sColumnName & vbCrLf & _
                      "FROM " & sTableName & vbCrLf
                    
                    ' Add the JOIN code for the order.
                    For iLoop = 1 To UBound(avOrderJoinTables, 2)
                      strUDFRuntimeCode = strUDFRuntimeCode & _
                        "LEFT OUTER JOIN " & avOrderJoinTables(2, iLoop) & " ON " & sTableName & ".id_" & avOrderJoinTables(1, iLoop) & " = " & avOrderJoinTables(2, iLoop) & ".id" & vbNewLine
                    Next iLoop
                    
                    strUDFRuntimeCode = strUDFRuntimeCode & _
                      "WHERE @id = " & sTableName & ".id_" & Trim(Str(mobjBaseComponent.ParentExpression.BaseTableID)) & vbCrLf
    
                    ' Add the filter code as required.
                    If LenB(sFilterCode) <> 0 Then
                      strUDFRuntimeCode = strUDFRuntimeCode & _
                        "AND " & sTableName & ".id IN" & vbNewLine & _
                        "(" & vbNewLine & _
                        sFilterCode & _
                        ")" & vbNewLine
                    End If

                    ' Add the order code as required.
                    If mlngSelOrderID > 0 Then
                      strUDFRuntimeCode = strUDFRuntimeCode & _
                        "ORDER BY " & sOrderCode & vbNewLine
                    End If

                    ' Finish off udf code
                    strUDFRuntimeCode = strUDFRuntimeCode & _
                      "OPEN GetRecord" & vbCrLf & _
                      "FETCH ABSOLUTE " & Trim(Str(mlngSelectionLine)) & " FROM GetRecord INTO @Result" & vbCrLf & _
                      "CLOSE GetRecord" & vbCrLf & _
                      "DEALLOCATE GetRecord" & vbCrLf & _
                      "RETURN @Result" & vbCrLf & _
                      "END"
                    
                    DropFunction strUDFName
                    gADOCon.Execute strUDFRuntimeCode, , adExecuteNoRecords
    
                    sCode = sCode & " [dbo].[" & strUDFName & "](" & sParentTableName & ".id)"
                  Else
                    fOK = False
                  End If
    
               Case Else
                  ' Unrecognised child record selection option.
                  fOK = False
              End Select
    
              sCode = sCode & ")"
    
              ' Add the table name to the list of source tables if it is not already there.
              fNewSourceTable = True
              For iLoop = 1 To UBound(palngSourceTables, 2)
                If (palngSourceTables(1, iLoop) = 0) And _
                  (palngSourceTables(2, iLoop) = mlngTableID) Then
                  fNewSourceTable = False
                  Exit For
                End If
              Next iLoop
              
              If fNewSourceTable Then
                iNextIndex = UBound(palngSourceTables, 2) + 1
                ReDim Preserve palngSourceTables(2, iNextIndex)
                palngSourceTables(1, iNextIndex) = 0
                palngSourceTables(2, iNextIndex) = mlngTableID
              End If
            End If
          End If
        End If
      End If
    
      If ReturnType = giEXPRVALUE_DATE Then
        sCode = "convert(" & vbNewLine & _
          "datetime, " & vbNewLine & _
          "convert(" & vbNewLine & _
          "varchar(20), " & vbNewLine & _
          sCode & "," & vbNewLine & _
          "101)" & vbNewLine & _
          ")"
      End If
    
      If ReturnType = giEXPRVALUE_NUMERIC Then
        sCode = "convert(" & vbNewLine & _
          "float, " & vbNewLine & _
          sCode & vbNewLine & _
          ")"
      End If
    End If
    
      ' JDM - 19/12/01 - Fault 3299 - Problems concatenating strings
    If ReturnType = giEXPRVALUE_CHARACTER Then
      sCode = "IsNull((" & sCode & "),'')"
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
  ComponentType = giCOMPONENT_FIELD
  
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
    ' Add the selection type description if required.
    If (miFieldPassType = giPASSBY_VALUE) Then
      ' Only give the full description if the field is in a child table of the
      ' expression's parent table.
      With recRelEdit
        .Index = "idxParentID"
        .Seek "=", mobjBaseComponent.ParentExpression.BaseTableID, mlngTableID
        
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
  Dim objFieldCopy As New CExprField
  
  ' Copy the component's basic properties.
  With objFieldCopy
    .ColumnID = mlngColumnID
    .FieldPassType = miFieldPassType
    .SelectionLine = mlngSelectionLine
    .SelectionOrderID = mlngSelOrderID
    .SelectionType = miSelectionType
    .SelectionFilter = mlngSelFilterID
    .TableID = mlngTableID
  End With
  
  Set CopyComponent = objFieldCopy
   
  ' Disassociate object variables.
  Set objFieldCopy = Nothing
   
End Function

Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewTableID As Long
  Dim lngNewOrderID As Long
  Dim lngNewColumnID As Long
  Dim lngNewParentTableID As Long
  Dim objFieldClone As New CExprField
  
  If miFieldPassType = giPASSBY_REFERENCE Then
    lngNewTableID = mlngTableID
    lngNewColumnID = mlngColumnID
  Else
    lngNewTableID = 0
    lngNewColumnID = 0
  End If
  lngNewOrderID = 0
  lngNewParentTableID = 0
  
  ' Get the new table, column, and order IDs from the clone register.
  For iIndex = 1 To UBound(pavCloneRegister, 2)
    If (pavCloneRegister(1, iIndex) = "TABLE") And _
      (pavCloneRegister(2, iIndex) = mlngTableID) And _
      (lngNewTableID = 0) Then
      lngNewTableID = pavCloneRegister(3, iIndex)
    End If
    
    If (pavCloneRegister(1, iIndex) = "COLUMN") And _
      (pavCloneRegister(2, iIndex) = mlngColumnID) And _
      (lngNewColumnID = 0) Then
      lngNewColumnID = pavCloneRegister(3, iIndex)
    End If
    
    If (pavCloneRegister(1, iIndex) = "ORDER") And _
      (pavCloneRegister(2, iIndex) = mlngSelOrderID) And _
      (lngNewOrderID = 0) Then
      lngNewOrderID = pavCloneRegister(3, iIndex)
    End If
    
    If (pavCloneRegister(1, iIndex)) = "TABLE" And _
      (pavCloneRegister(2, iIndex) = mobjBaseComponent.ParentExpression.BaseTableID) And _
      (lngNewParentTableID = 0) Then
      lngNewParentTableID = pavCloneRegister(3, iIndex)
    End If
    
    If ((mlngTableID = 0) Or (lngNewTableID > 0)) And _
      ((mlngSelOrderID = 0) Or (lngNewOrderID > 0)) And _
      ((mlngColumnID = 0) Or (lngNewColumnID > 0)) And _
      ((mobjBaseComponent.ParentExpression.BaseTableID = 0) Or (lngNewParentTableID > 0)) Then
      Exit For
    End If
  Next iIndex
  
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
  If ((mobjBaseComponent.ParentExpression.BaseTableID > 0) And (lngNewParentTableID = 0) And (lngNewTableID = 0)) Then
    lngNewTableID = mobjBaseComponent.ParentExpression.BaseTableID
  End If
  
  ' Copy the component's basic properties.
  With objFieldClone
    .ColumnID = lngNewColumnID
    .FieldPassType = miFieldPassType
    .SelectionLine = mlngSelectionLine
    .SelectionOrderID = lngNewOrderID
    .SelectionType = miSelectionType
    .SelectionFilter = mlngSelFilterID
    .TableID = lngNewTableID
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


Private Sub Class_Initialize()
  ' Initialise properties.
  miFieldPassType = giPASSBY_VALUE
  
End Sub
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


'MH20060727
'Private Sub Class_Terminate()
'  Terminate
'End Sub
'
'Public Sub Terminate()
'  Set mobjBaseComponent = Nothing
'End Sub

