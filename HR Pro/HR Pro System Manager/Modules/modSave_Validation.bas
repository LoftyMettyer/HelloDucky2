Attribute VB_Name = "modSave_Validation"
Option Explicit


Public Function CreateValidationStoredProcedures(pfRefreshDatabase As Boolean) As Boolean
  ' Create the record validation stored procedures.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean

  fOK = True
  
  ' Create the validation stored procedures for each table.
  With recTabEdit
    .Index = "idxTableID"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    OutputCurrentProcess2 vbNullString, .RecordCount
    
    Do While Not .EOF
      If Not !Deleted Then
        
        OutputCurrentProcess2 !TableName
        gobjProgress.UpdateProgress2
        
        ' Create the view validation stored procedure.
        fOK = CreateViewValidationStoredProcedure(!TableID, !TableName, !TableType, !New, pfRefreshDatabase)
        
        'MH20011030 Fault 3002
        'If fOK And (!New Or !Changed Or pfRefreshDatabase Or (!TableName <> !OriginalTableName)) Then
        'If fOK And (!New Or !Changed Or pfRefreshDatabase Or (Application.ChangedTableName)) Then
        'If fOK And (!New Or !Changed Or pfRefreshDatabase Or Application.ChangedTableName Or Application.ChangedViewName) Then
          ' Create the validation stored procedure.
        fOK = CreateValidationStoredProcedure(!TableID, !TableName, !TableType, pfRefreshDatabase)
        'End If
      End If
      
      If Not fOK Then
        Exit Do
      End If
    
      .MoveNext
    Loop
  End With
    
TidyUpAndExit:
  CreateValidationStoredProcedures = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating validation stored procedure"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function CreateValidationStoredProcedure(pLngCurrentTableID As Long, _
  psCurrentTableName As String, _
  piTableType As Integer, _
  pfRefreshDatabase As Boolean) As Boolean
  ' Create the validation stored procedure for the given table (plngCurrentTableID).
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fCreateSP As Boolean
  Dim fIsLiteral As Boolean
  Dim fIsTopLevel As Boolean
  Dim iLoop As Integer
  Dim iCounter As Integer
  Dim iCharIndex As Integer
  Dim sSQL As String
  Dim sUDFName As String
  Dim sSPCode As String
  Dim sDuplicateCheckCode As String
  Dim sDuplicateColumns As String
  Dim sParentTableName As String
  Dim rsTableName As DAO.Recordset
  Dim aryOverlapColumns() As String
  Dim aryOverlapParentJoins() As String

  fOK = True
  sUDFName = gsVALIDATIONSPPREFIX & psCurrentTableName
  sDuplicateCheckCode = vbNullString
  sDuplicateColumns = vbNullString
  
  fIsTopLevel = (piTableType = iTabParent)

  ' Drop any existing stored procedure.
  DropFunction sUDFName
   
  '
  ' Create the stored procedure creation string.
  '

  '******************************************************************************
  ' TM20010719 Fault 2242 - @ParentIDCount added to SP.                         *
  ' Counts how many parents are linked to the new record if zero parents are    *
  ' linked then returns message to the calling code.                            *
  '******************************************************************************

  sSPCode = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* validation user defined function.         */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE FUNCTION dbo." & sUDFName & "(" & vbNewLine & _
    "    @piRecordID     integer," & vbNewLine & _
    "    @psDescription  nvarchar(255))" & vbNewLine & _
    "RETURNS varchar(MAX)" & vbNewLine & _
    "AS" & vbNewLine
    
  sSPCode = sSPCode & _
    "BEGIN" & vbNewLine & _
    "    DECLARE @iRecCount integer," & vbNewLine & _
    "        @pfResult bit," & vbNewLine & _
    "        @piSeverity integer," & vbNewLine & _
    "        @psInvalidityMessage varchar(MAX)," & vbNewLine & _
    "        @fCustomResult bit," & vbNewLine & _
    "        @fItemOK bit," & vbNewLine & _
    "        @fEmptyMask bit," & vbNewLine & _
    "        @sTmpChar varchar(MAX)," & vbNewLine & _
    "        @sOneChar varchar(1)," & vbNewLine & _
    "        @dblTmpNum float," & vbNewLine & _
    "        @fTmpLogic bit," & vbNewLine & _
    "        @dtTmpDate datetime," & vbNewLine & _
    "        @sGroupName sysname," & vbNewLine & _
    "        @sCommandString nvarchar(MAX)," & vbNewLine & _
    "        @sParamDefinition nvarchar(500)," & vbNewLine & _
    "        @iParentID integer," & vbNewLine & _
    "        @iParentIDCount integer;" & vbNewLine & vbNewLine & _
    "    SET @pfResult = 1;" & vbNewLine & _
    "    SET @piSeverity = 0;" & vbNewLine & _
    "    SET @psInvalidityMessage = '';" & vbNewLine & _
    "    SET @iParentIDCount = 0;" & vbNewLine

  ' Add checks to ensure that the parent records still exist.
  ' eg. A quick-entry screen record is linked to a parent table record. Before the quick-entry
  ' record is saved, the parent record is deleted by someone else. This check will stop the
  ' quick-entry record being saved.
  With recRelEdit
    .Index = "idxChildID"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Seek "=", pLngCurrentTableID
    
      If Not .NoMatch Then
        Do While fOK And (Not .EOF)
          If !childID <> pLngCurrentTableID Then
            Exit Do
          End If
        
          ' Get the parent table name.
          sSQL = "SELECT tableName" & _
            " FROM tmpTables" & _
            " WHERE tableID = " & Trim(Str(!parentID))
          Set rsTableName = daoDb.OpenRecordset(sSQL, _
            dbOpenForwardOnly, dbReadOnly)
          If Not (rsTableName.BOF And rsTableName.EOF) Then
            sParentTableName = rsTableName.Fields("tableName").value
          Else
            fOK = False
          End If
        
          If fOK Then
            sSPCode = sSPCode & vbNewLine & _
              "    /* Check related record in '" & sParentTableName & "' still exists. */" & vbNewLine & _
              "    SELECT @iParentID = id_" & Trim(Str(!parentID)) & vbNewLine & _
              "    FROM " & psCurrentTableName & vbNewLine & _
              "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
              "    IF ISNULL(@iParentID, 0) > 0" & vbNewLine & _
              "    BEGIN" & vbNewLine & _
              "        SET @iParentIDCount = @iParentIDCount + 1" & vbNewLine & _
              "        SELECT @iRecCount = COUNT(" & sParentTableName & ".id)" & vbNewLine & _
              "        FROM " & sParentTableName & vbNewLine & _
              "        WHERE id = @iParentID" & vbNewLine & _
              "        IF @iRecCount = 0" & vbNewLine & _
              "        BEGIN" & vbNewLine & _
              "            SET @pfResult = 0" & vbNewLine & _
              "            SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The linked record in the ''" & sParentTableName & "'' table no longer exists.'" & vbNewLine & _
              "        END" & vbNewLine & _
              "    END" & vbNewLine
              
          End If
          
          .MoveNext
        Loop

        'MH20010828
        'sSPCode = sSPCode & vbNewLine & " IF @iParentIDCount = 0 SET @psInvalidityMessage = 'No link made with the Parent table.'" & vbNewLine
        sSPCode = sSPCode & vbNewLine & _
          "  IF @iParentIDCount = 0" & vbNewLine & _
          "  BEGIN" & vbNewLine & _
          "    SET @psInvalidityMessage = 'No link made with the Parent table.'" & vbNewLine & _
          "    SET @pfResult = 0" & vbNewLine & _
          "  END" & vbNewLine

      End If
    End If
  End With
  
    
  If fOK Then
    ' Loop through the current table's columns, checking if any of them require validating.
    With recColEdit
      .Index = "idxName"
      .Seek ">=", pLngCurrentTableID
  
      If Not .NoMatch Then
        Do While Not .EOF
  
          If !TableID <> pLngCurrentTableID Then
            Exit Do
          End If
  
          If Not !Deleted Then
            If !uniqueCheckType = -1 Then
              ' Add the unique check (within the entire table) code for the current column if required.
              sSPCode = sSPCode & vbNewLine & _
                "    /* '" & !ColumnName & "' - unique check (entire table). */" & vbNewLine & _
                "    SELECT @iRecCount = COUNT(tmpA.id)" & vbNewLine & _
                "    FROM " & psCurrentTableName & " tmpA," & vbNewLine & _
                "        " & psCurrentTableName & " tmpB" & vbNewLine & _
                "    WHERE tmpB.id = @piRecordID" & vbNewLine & _
                "        AND tmpA.id <> @piRecordID" & vbNewLine & _
                "        AND tmpA." & !ColumnName & " = tmpB." & !ColumnName & vbNewLine
              
              'TM05052004 - Unique Not Mandatory
              If Not !Mandatory Then
                Select Case !DataType
                  Case dtVARCHAR, dtLONGVARCHAR, dtLONGVARBINARY, dtVARBINARY
                    sSPCode = sSPCode & _
                      "        AND (ISNULL(tmpA." & !ColumnName & ",'') <> '')" & vbNewLine
                  
                  Case dtINTEGER, dtNUMERIC
                    sSPCode = sSPCode & _
                      "        AND (ISNULL(tmpA." & !ColumnName & ",0) <> 0)" & vbNewLine
                  
                  Case dtBIT
                    sSPCode = sSPCode & _
                      "        /* Logic columns cannot have mandatory checks. */" & vbNewLine
                        
                  Case dtTIMESTAMP
                    sSPCode = sSPCode & _
                      "        AND (tmpA." & !ColumnName & " IS NOT NULL)" & vbNewLine
                End Select
              End If
              
              sSPCode = sSPCode & vbNewLine & _
                "    IF @iRecCount > 0" & vbNewLine & _
                "    BEGIN" & vbNewLine & _
                "        SET @pfResult = 0" & vbNewLine & _
                "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' field is not unique within the entire table.'" & vbNewLine & _
                "    END" & vbNewLine
            Else
              If (!uniqueCheckType = -2) Or (!uniqueCheckType > 0) Then
                ' Add the unique check (within the sibling records) code for the current column if required.
                sSPCode = sSPCode & vbNewLine & _
                  "    /* '" & !ColumnName & "' - unique check (sibling records). */" & vbNewLine & _
                  "    SELECT @iRecCount = COUNT(tmpA.id)" & vbNewLine & _
                  "    FROM " & psCurrentTableName & " tmpA," & vbNewLine & _
                  "        " & psCurrentTableName & " tmpB" & vbNewLine & _
                  "    WHERE tmpB.id = @piRecordID" & vbNewLine & _
                  "        AND tmpA.id <> @piRecordID" & vbNewLine & _
                  "        AND tmpA." & !ColumnName & " = tmpB." & !ColumnName & vbNewLine
                
                'TM05052004 - Unique Not Mandatory
                If Not !Mandatory Then
                  Select Case !DataType
                    Case dtVARCHAR, dtLONGVARCHAR, dtLONGVARBINARY, dtVARBINARY
                      sSPCode = sSPCode & _
                        "        AND (ISNULL(tmpA." & !ColumnName & ",'') <> '')" & vbNewLine
                    
                    Case dtINTEGER, dtNUMERIC
                      sSPCode = sSPCode & _
                        "        AND (ISNULL(tmpA." & !ColumnName & ",0) <> 0)" & vbNewLine
                    
                    Case dtBIT
                      sSPCode = sSPCode & _
                        "        /* Logic columns cannot have mandatory checks. */" & vbNewLine
                          
                    Case dtTIMESTAMP
                      sSPCode = sSPCode & _
                        "        AND (tmpA." & !ColumnName & " IS NOT NULL)" & vbNewLine
                  End Select
                End If
                
                iCounter = 0
                With recRelEdit
                  .Index = "idxChildID"
                  .MoveFirst
                  .Seek "=", pLngCurrentTableID
                
                  If Not .NoMatch Then
                    Do While (Not .EOF)
                      If !childID <> pLngCurrentTableID Then
                        Exit Do
                      End If
                    
                      If (recColEdit!uniqueCheckType = -2) Or (recColEdit!uniqueCheckType = !parentID) Then
                        sSPCode = sSPCode & IIf(iCounter > 0, vbNewLine & "        OR ", "        AND (") & _
                          "(tmpA.ID_" & Trim(Str(!parentID)) & " = tmpB.id_" & Trim(Str(!parentID)) & " AND tmpA.ID_" & Trim(Str(!parentID)) & " > 0)"
                        iCounter = iCounter + 1
                      End If

                      .MoveNext
                    Loop
                  
                    If iCounter > 0 Then
                      sSPCode = sSPCode & ")" & vbNewLine
                    End If
                  End If
                End With
    
                sSPCode = sSPCode & vbNewLine & _
                  "    IF @iRecCount > 0" & vbNewLine & _
                  "    BEGIN" & vbNewLine & _
                  "        SET @pfResult = 0" & vbNewLine & _
                  "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' field is not unique within sibling records.'" & vbNewLine & _
                  "    END" & vbNewLine
              End If
            End If
            
  '******************************************************************************
  ' TM20010719 Fault 2242 - !ColumnType <> 4 clause added to ignore all linked  *
  ' columns. (It doesn't need to validate the linked columns because this is    *
  ' done using the @ParentIDCount.                                              *
  '******************************************************************************
            
            If !Mandatory And !columntype <> 4 Then
              ' Add the mandatory check code for the current column if required.
              sSPCode = sSPCode & vbNewLine & _
                "    /* '" & !ColumnName & "' - mandatory check. */"
                
              Select Case !DataType
                Case dtVARCHAR, dtLONGVARCHAR, dtLONGVARBINARY, dtVARBINARY
                  sSPCode = sSPCode & vbNewLine & _
                    "    SELECT @sTmpChar = " & !ColumnName & vbNewLine & _
                    "    FROM " & psCurrentTableName & vbNewLine & _
                    "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                    "    IF @sTmpChar IS null" & vbNewLine & _
                    "    BEGIN" & vbNewLine & _
                    "        SET @pfResult = 0" & vbNewLine & _
                    "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' field is mandatory.'" & vbNewLine & _
                    "    END" & vbNewLine
                  sSPCode = sSPCode & _
                    "    ELSE" & vbNewLine & _
                    "    BEGIN" & vbNewLine & _
                    "        IF len(ltrim(rtrim(@sTmpChar))) = 0" & vbNewLine & _
                    "        BEGIN" & vbNewLine & _
                    "            SET @pfResult = 0" & vbNewLine & _
                    "            SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' field is mandatory.'" & vbNewLine & _
                    "        END" & vbNewLine & _
                    "    END" & vbNewLine
                Case dtINTEGER, dtNUMERIC
                  'JPD 20060105 Fault 10655
                  sSPCode = sSPCode & vbNewLine & _
                    "    SELECT @dblTmpNum = " & !ColumnName & vbNewLine & _
                    "    FROM " & psCurrentTableName & vbNewLine & _
                    "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                    "    IF @dblTmpNum IS null" & vbNewLine & _
                    "        OR @dblTmpNum = 0" & vbNewLine & _
                    "    BEGIN" & vbNewLine & _
                    "        SET @pfResult = 0" & vbNewLine & _
                    "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' field is mandatory.'" & vbNewLine & _
                    "    END" & vbNewLine
                Case dtBIT
                  sSPCode = sSPCode & vbNewLine & _
                    "    /* Logic columns cannot have mandatory checks. */" & vbNewLine
                Case dtTIMESTAMP
                  sSPCode = sSPCode & vbNewLine & _
                    "    SELECT @dtTmpDate = " & !ColumnName & vbNewLine & _
                    "    FROM " & psCurrentTableName & vbNewLine & _
                    "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                    "    IF @dtTmpDate IS null" & vbNewLine & _
                    "    BEGIN" & vbNewLine & _
                    "        SET @pfResult = 0" & vbNewLine & _
                    "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' field is mandatory.'" & vbNewLine & _
                    "    END" & vbNewLine
              End Select
            End If
  
            If !Duplicate Then
              ' Add the duplicate check code for the current column if required.
              If LenB(sDuplicateCheckCode) = 0 Then
                sDuplicateCheckCode = "    /* Duplicate check. */" & vbNewLine & _
                  "    SELECT @iRecCount = COUNT(tmpA.id)" & vbNewLine & _
                  "    FROM " & psCurrentTableName & " tmpA," & vbNewLine & _
                  "        " & psCurrentTableName & " tmpB" & vbNewLine & _
                  "    WHERE tmpB.id = @piRecordID" & vbNewLine & _
                  "        AND tmpA.id <> @piRecordID" & vbNewLine & _
                  "        AND tmpA." & !ColumnName & " = tmpB." & !ColumnName & vbNewLine
              Else
                sDuplicateCheckCode = sDuplicateCheckCode & _
                  "        AND tmpA." & !ColumnName & " = tmpB." & !ColumnName & vbNewLine
              End If
              
              sDuplicateColumns = sDuplicateColumns & _
                IIf(LenB(sDuplicateColumns) <> 0, " + " & vbNewLine, vbNullString) & _
                "                char(13) + '    ''" & !ColumnName & "'''"
            End If
  
            If !lostFocusExprID > 0 Then
              ' Add the column validation check code for the current column if required.
              sSPCode = sSPCode & vbNewLine & _
                "    -- " & !ColumnName & " - custom validation check." & vbNewLine & _
                "    SELECT @fCustomResult = dbo.[udfmask_" & Trim(Str(!lostFocusExprID)) + "](@piRecordID);" & vbNewLine & _
                "    IF @fCustomResult = 0" & vbNewLine & _
                "    BEGIN" & vbNewLine & _
                "        SET @pfResult = 0;" & vbNewLine & _
                "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + '''" & !ColumnName & "'' - " & IIf(IsNull(!ErrorMessage), "Validation failure", IIf(Len(LTrim(RTrim(!ErrorMessage))) = 0, "Validation failure", Replace(Replace(!ErrorMessage, "'", "''"), "%", "%%"))) & ".';" & vbNewLine & _
                "    END" & vbNewLine
            End If
            
            If !ControlType = giCTRL_SPINNER Then
              ' Add the spinner min/max check code for the current column if required.
              sSPCode = sSPCode & vbNewLine & _
                "    /* '" & !ColumnName & "' - spinner min/max check. */" & vbNewLine & _
                "    SELECT @dblTmpNum = " & !ColumnName & vbNewLine & _
                "    FROM " & psCurrentTableName & vbNewLine & _
                "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                "    IF @dblTmpNum IS null SET @dblTmpNum = 0" & vbNewLine & _
                "    IF @dblTmpNum < " & Trim(Str(!spinnerMinimum)) & vbNewLine & _
                "    BEGIN" & vbNewLine & _
                "        SET @pfResult = 0" & vbNewLine & _
                "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' value is less than the defined minimum of " & Trim(Str(!spinnerMinimum)) & ".'" & vbNewLine & _
                "    END" & vbNewLine
              sSPCode = sSPCode & _
                "    ELSE" & vbNewLine & _
                "    BEGIN" & vbNewLine & _
                "        IF @dblTmpNum > " & Trim(Str(!spinnerMaximum)) & vbNewLine & _
                "        BEGIN" & vbNewLine & _
                "            SET @pfResult = 0" & vbNewLine & _
                "            SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' value is greater than the defined maximum of " & Trim(Str(!spinnerMaximum)) & ".'" & vbNewLine & _
                "        END" & vbNewLine & _
                "    END" & vbNewLine
            End If
            
            If (!columntype = giCOLUMNTYPE_DATA) And _
              ((!ControlType = giCTRL_OPTIONGROUP) Or (!ControlType = giCTRL_COMBOBOX)) Then
              ' Add the optionGroup/dropdownList check code for the current column if required.
              sSPCode = sSPCode & vbNewLine & _
                "    /* '" & !ColumnName & "' - optionGroup/dropdownList check. */" & vbNewLine & _
                "    SELECT @sTmpChar = " & !ColumnName & vbNewLine & _
                "    FROM " & psCurrentTableName & vbNewLine & _
                "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                "    IF @sTmpChar IS null SET @sTmpChar = ''" & vbNewLine & _
                "    SET @fItemOK = 0" & vbNewLine
              
              If (!ControlType = giCTRL_COMBOBOX) Then
                sSPCode = sSPCode & _
                  "    IF len(ltrim(rtrim(@sTmpChar))) = 0 SET @fItemOK = 1" & vbNewLine
              End If
              
              recContValEdit.Index = "idxColumnID"
              recContValEdit.Seek ">=", !ColumnID
    
              If Not recContValEdit.NoMatch Then
                Do While Not recContValEdit.EOF
                  If recContValEdit!ColumnID <> !ColumnID Then
                    Exit Do
                  End If
          
                  sSPCode = sSPCode & _
                    "    IF @sTmpChar = '" & Replace(recContValEdit!value, "'", "''") & "' SET @fItemOK = 1" & vbNewLine
  
                  recContValEdit.MoveNext
                Loop
              End If
  
              sSPCode = sSPCode & _
                "    IF @fItemOK = 0" & vbNewLine & _
                "    BEGIN" & vbNewLine & _
                "        SET @pfResult = 0" & vbNewLine & _
                "        SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' value is not in the list of valid values.'" & vbNewLine & _
                "    END" & vbNewLine
            End If
  
            If (!ControlType = giCTRL_TEXTBOX) And _
              (!DataType = dtVARCHAR) And _
              (Not !MultiLine) And _
              Len(!Mask) > 0 Then
              ' Add the mask check code for the current column if required.
              sSPCode = sSPCode & vbNewLine & _
                "    /* '" & !ColumnName & "' - mask check. */" & vbNewLine & _
                "    SELECT @sTmpChar = " & !ColumnName & vbNewLine & _
                "    FROM " & psCurrentTableName & vbNewLine & _
                "    WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                "    IF (NOT @sTmpChar IS null) AND (len(@sTmpChar) > 0)" & vbNewLine & _
                "    BEGIN" & vbNewLine & _
                "        SET @fItemOK = 1" & vbNewLine & _
                "        SET @fEmptyMask = 1" & vbNewLine
                
              fIsLiteral = False
              iCharIndex = 1
              
              For iLoop = 1 To Len(!Mask)
                If fIsLiteral Then
                  fIsLiteral = False
                  
                  sSPCode = sSPCode & _
                    "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                    "        IF ascii(@sOneChar) <> ascii('" & Mid(!Mask, iLoop, 1) & "') SET @fItemOK = 0" & vbNewLine
                Else
                  Select Case Mid(!Mask, iLoop, 1)
                    Case "A" ' Character must be uppercase alphabetic.
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF (ascii(@sOneChar) < ascii('A')) OR (ascii(@sOneChar) > ascii('Z')) SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine

                    Case "a" ' Character must be lowercase alphabetic.
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF (ascii(@sOneChar) < ascii('a')) OR (ascii(@sOneChar) > ascii('z')) SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine
                    
                    Case "S" ' Character must be uppercase alphabetic or SPACE
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF ((ascii(@sOneChar) < ascii('A')) OR (ascii(@sOneChar) > ascii('Z'))) AND @sOneChar <> ' ' SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine

                    Case "s" ' Character must be lowercase alphabetic or SPACE
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF ((ascii(@sOneChar) < ascii('a')) OR (ascii(@sOneChar) > ascii('z'))) AND @sOneChar <> ' ' SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine
                    
                    Case "\" ' Next character is a literal.
                      fIsLiteral = True
                      iCharIndex = iCharIndex - 1
                      
                    Case "9" ' Character must be numeric.
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF (ascii(@sOneChar) < ascii('0')) OR (ascii(@sOneChar) > ascii('9')) SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine
                    
                    Case "#" ' Character must be numeric or symbolic.
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF ((ascii(@sOneChar) < ascii('0')) OR (ascii(@sOneChar) > ascii('9')))" & vbNewLine & _
                        "            AND (ascii(@sOneChar) <> ascii('$'))" & vbNewLine & _
                        "            AND (ascii(@sOneChar) <> ascii('%'))" & vbNewLine & _
                        "            AND (ascii(@sOneChar) <> ascii('+'))" & vbNewLine & _
                        "            AND (ascii(@sOneChar) <> ascii('-'))" & vbNewLine & _
                        "            AND (ascii(@sOneChar) <> ascii('.'))" & vbNewLine & _
                        "            AND (ascii(@sOneChar) <> ascii(',')) SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine
                    
                    Case "b" ' Character must be boolean (0 or 1).
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF (@sOneChar <> '0') AND (@sOneChar <> '1') SET @fItemOK = 0" & vbNewLine & _
                        "        IF (ascii(@sOneChar) <> ascii('_')) SET @fEmptyMask = 0" & vbNewLine
                    
                    Case Else ' Literal.
                      sSPCode = sSPCode & _
                        "        SELECT @sOneChar = substring(@sTmpChar, " & Trim$(Str$(iCharIndex)) & ", 1)" & vbNewLine & _
                        "        IF ascii(@sOneChar) <> ascii('" & Replace(Mid(!Mask, iLoop, 1), "'", "''") & "') SET @fItemOK = 0" & vbNewLine
                    
                  End Select
                End If
                
                iCharIndex = iCharIndex + 1
              Next iLoop
                
              sSPCode = sSPCode & _
                "        IF (@fItemOK = 0) AND (@fEmptyMask = 0)" & vbNewLine & _
                "        BEGIN" & vbNewLine & _
                "            SET @pfResult = 0" & vbNewLine & _
                "            SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'The ''" & !ColumnName & "'' value does not fit the defined mask.'" & vbNewLine & _
                "        END" & vbNewLine & _
                "    END" & vbNewLine
            End If
            
          End If
          
          .MoveNext
        Loop
      End If
    End With
    
    ' Add the duplicate check if there is one.
    If LenB(sDuplicateCheckCode) <> 0 Then
      sSPCode = sSPCode & vbNewLine & _
        sDuplicateCheckCode & vbNewLine & _
        "    IF @iRecCount > 0" & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        SET @pfResult = 0" & vbNewLine & _
        "        BEGIN" & vbNewLine & _
        "            SET @psInvalidityMessage = @psInvalidityMessage + char(13) + 'Duplicate record found. Duplicate columns :' + " & vbNewLine & _
        sDuplicateColumns & vbNewLine & _
        "        END" & vbNewLine & _
        "    END" & vbNewLine
    End If
    
    
  ' Add overlap checks
  aryOverlapColumns = GetOverlapColumnsArray(pLngCurrentTableID)
  aryOverlapParentJoins = GetOverlapParentJoins(pLngCurrentTableID)
  For iCounter = LBound(aryOverlapColumns, 2) To UBound(aryOverlapColumns, 2) - 1
    sSPCode = sSPCode & vbNewLine & _
        "    SELECT TOP 1 @psInvalidityMessage = @psInvalidityMessage + char(13) + '" & aryOverlapColumns(5, iCounter) & "'" & vbNewLine & _
        "    FROM dbo.[" & psCurrentTableName & "]" & vbNewLine
            
    If UBound(aryOverlapParentJoins) > 0 Then
      sSPCode = sSPCode & vbNewLine & _
          Join(aryOverlapParentJoins, vbNewLine)
    Else
      sSPCode = sSPCode & vbNewLine & _
          "INNER JOIN dbo.[" & psCurrentTableName & "] inserted ON inserted.[ID] = @piRecordID"
    End If
        
    sSPCode = sSPCode & vbNewLine & _
        "    WHERE dbo.udfASRDateOverlap(" & _
        IIf(Not aryOverlapColumns(0, iCounter) = vbNullString, "inserted.[" & aryOverlapColumns(0, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(1, iCounter) = vbNullString, "inserted.[" & aryOverlapColumns(1, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(2, iCounter) = vbNullString, "inserted.[" & aryOverlapColumns(2, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(3, iCounter) = vbNullString, "inserted.[" & aryOverlapColumns(3, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(4, iCounter) = vbNullString, "inserted.[" & aryOverlapColumns(4, iCounter) & "], ", "NULL, ") & vbNewLine & Space(12) & _
        IIf(Not aryOverlapColumns(0, iCounter) = vbNullString, "dbo.[" & psCurrentTableName & "].[" & aryOverlapColumns(0, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(1, iCounter) = vbNullString, "dbo.[" & psCurrentTableName & "].[" & aryOverlapColumns(1, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(2, iCounter) = vbNullString, "dbo.[" & psCurrentTableName & "].[" & aryOverlapColumns(2, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(3, iCounter) = vbNullString, "dbo.[" & psCurrentTableName & "].[" & aryOverlapColumns(3, iCounter) & "], ", "NULL, ") & _
        IIf(Not aryOverlapColumns(4, iCounter) = vbNullString, "dbo.[" & psCurrentTableName & "].[" & aryOverlapColumns(4, iCounter) & "]", "NULL") & ")=1 " & vbNewLine & _
        "        AND NOT dbo.[" & psCurrentTableName & "].[ID] = @piRecordID;" & vbNewLine & _
        "    IF LEN(@psInvalidityMessage) > 0 SET @pfResult = 0;" & vbNewLine
    Next
    
  End If

  sSPCode = sSPCode & vbNewLine & vbNewLine & _
    "    IF LEN(@psInvalidityMessage) > 0" & vbNewLine & _
    "        SET @psInvalidityMessage = ISNULL(@psDescription,'') + CHAR(13)" & vbNewLine & _
    "           + '------------------------------------------'" & vbNewLine & _
    "           + @psInvalidityMessage + CHAR(13) + CHAR(13);" & vbNewLine & _
    "    RETURN @psInvalidityMessage;" & vbNewLine & _
    "END"


  ' Create the stored procedure.
  gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateValidationStoredProcedure = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error creating validation stored procedure"
  Err = False
 
 Resume TidyUpAndExit

End Function


Private Function GetOverlapColumnsArray(ByVal plngTableID As Long) As String()

  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim iDefaultItem As Integer
  Dim sTableName As String
  Dim arOverlaps() As String
  Dim lngCount As Long

  lngCount = 0
  ReDim arOverlaps(7, 10000)

  sSQL = "SELECT *" & _
    " FROM tmpTableValidations" & _
    " WHERE (deleted = FALSE)" & _
    " AND (tableID = " & CStr(plngTableID) & ")"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  While Not rsTemp.EOF
    
    arOverlaps(0, lngCount) = GetColumnName(rsTemp.Fields("EventStartDateColumnID").value, True)
    arOverlaps(1, lngCount) = GetColumnName(rsTemp.Fields("EventStartSessionColumnID").value, True)
    arOverlaps(2, lngCount) = GetColumnName(rsTemp.Fields("EventEndDateColumnID").value, True)
    arOverlaps(3, lngCount) = GetColumnName(rsTemp.Fields("EventEndSessionColumnID").value, True)
    arOverlaps(4, lngCount) = GetColumnName(rsTemp.Fields("EventTypeColumnID").value, True)
    arOverlaps(5, lngCount) = rsTemp.Fields("Message").value
    arOverlaps(6, lngCount) = rsTemp.Fields("FilterID").value
        
    lngCount = lngCount + 1
    rsTemp.MoveNext
  Wend
  rsTemp.Close
  
  Set rsTemp = Nothing
  
  ReDim Preserve arOverlaps(7, lngCount)
  GetOverlapColumnsArray = arOverlaps


End Function

Private Function GetOverlapParentJoins(ByVal plngTableID As Long) As String()

  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim iDefaultItem As Integer
  Dim sTableName As String
  Dim arOverlaps() As String
  Dim lngCount As Long

  lngCount = 0
  ReDim arOverlaps(10000)
  
  sTableName = GetTableName(plngTableID)
  
  sSQL = "SELECT *" & _
    " FROM tmpRelations" & _
    " " & _
    " WHERE (ChildID = " & CStr(plngTableID) & ")"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  While Not rsTemp.EOF
'    arOverlaps(lngCount) = "        INNER JOIN inserted ON inserted.[ID_" & rsTemp.Fields("ParentID").Value & "] = dbo.[" & sTableName & "].[id_" & rsTemp.Fields("ParentID").Value & "]"
    arOverlaps(lngCount) = "    INNER JOIN dbo.[" & sTableName & "] inserted ON inserted.[ID_" & rsTemp.Fields("ParentID").value & "] = dbo.[" & sTableName & "].[id_" & rsTemp.Fields("ParentID").value & "]" & _
          " AND inserted.[ID] = @piRecordID"
    lngCount = lngCount + 1
    rsTemp.MoveNext
  Wend
  rsTemp.Close
  
  Set rsTemp = Nothing
  
  ReDim Preserve arOverlaps(lngCount)
  GetOverlapParentJoins = arOverlaps

End Function

