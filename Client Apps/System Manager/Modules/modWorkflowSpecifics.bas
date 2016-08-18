Attribute VB_Name = "modWorkflowSpecifics"
Option Explicit

Private Const msGetEmailAddresses_PROCEDURENAME = "spASRGetWorkflowEmailAddresses"
Private Const msCheckPendingSteps_PROCEDURENAME = "spASRCheckPendingWorkflowSteps"
Private Const msIntCheckPendingSteps_PROCEDURENAME = "spASRIntCheckPendingWorkflowSteps"
Private Const msWorkspaceCheckPendingSteps_PROCEDURENAME = "spASRWorkspaceCheckPendingWorkflowSteps"
Private Const msGetDelegatedRecords_PROCEDURENAME = "spASRGetWorkflowDelegatedRecords"
Private Const msGetLoginName_PROCEDURENAME = "spASRSysGetLoginName"
Private Const msOutOfOfficeCheck_PROCEDURENAME = "spASRWorkflowOutOfOfficeCheck"
Private Const msOutOfOfficeSet_PROCEDURENAME = "spASRWorkflowOutOfOfficeSet"

Private Const msAscendantRecordID_FUNCTIONNAME = "udf_ASRWorkflowAscendantRecordID"
Private Const msValidTableRecordID_FUNCTIONNAME = "udf_ASRWorkflowValidTableRecord"
Private Const msGetLoginName_FUNCTIONNAME = "udf_ASRGetLoginName"
Private Const msGetDelegatedRecords_FUNCTIONNAME = "udfASRGetWorkflowDelegatedRecords"

Private mvar_fGeneralOK As Boolean
Private mvar_sGeneralMsg As String

Private mvar_sURL As String
Private mvar_sLoginColumn As String
Private mvar_sSecondLoginColumn As String
Private mvar_sLoginTable As String
Private mvar_lngActivateDelegationColumn As Long

Private mvar_sActivateDelegationColumn As String
Private mvar_lngDelegationEmail As Long
Private malngEmailColumns() As Long

Private msInsertLinkCode As String
Private msUpdateLinkCode As String
Private msRebuildLinkCode As String

Private msInsertLinkTemp As String
Private msUpdateLinkTemp As String

Private mfInitTrue As Boolean
Private mabytArray() As Byte
Private mlngHiByte As Long
Private mlngHiBound As Long
Private mabytAddTable(255, 255) As Byte
Private mabytXTable(255, 255) As Byte

Public Sub ParseWebFormMessage(psMessage As String, _
  ByRef psPart1 As String, _
  ByRef psPart2 As String, _
  ByRef psPart3 As String)
  ' Parse the given string, picking out the section tagged as the hypertext bit
  ' (\ul <text>\ulnone )
  
  Dim sSourceText As String
  Dim asText() As String
  Dim sChar As String
  Dim sNextChar As String
  Dim fDoingSlash As Boolean
'''  Dim iBracketLevel As Integer
  Dim fIgnoreChar As Boolean
  Dim sRTFCode As String
  Dim sRTFCodeToDo As String
  Dim iTextIndex As Integer
  Dim fLiteral As Boolean
'''  Dim iSelStart As Integer
'''  Dim iSelLength As Integer
'''  Dim sTemp As String
'''  Dim asDeniedCharacters() As String
'''  Dim iLoop As Integer
'''  Dim fFound As Boolean
'''  Dim sDeniedChar As String
'''  Dim iDeniedCharCount As Integer
'''  Dim sNewRTFText As String
  
  iTextIndex = 0
  ReDim asText(2)
  sSourceText = psMessage
  fDoingSlash = False
  sRTFCode = ""
  
  Do While Len(sSourceText) > 0
    sChar = Mid(sSourceText, 1, 1)
    sNextChar = Mid(sSourceText, 2, 1)

    fLiteral = sChar = "\" _
      And ((sNextChar = "\") _
        Or (sNextChar = "{") _
        Or (sNextChar = "}"))

    fIgnoreChar = fDoingSlash

    If fDoingSlash Then
      If sChar = " " Then
        fDoingSlash = False
        sRTFCodeToDo = sRTFCode
        sRTFCode = ""
      ElseIf sChar = "\" Then
        sRTFCodeToDo = sRTFCode
        sRTFCode = sChar
      Else
        sRTFCode = sRTFCode & Trim(Replace(Replace(sChar, vbCr, ""), vbLf, ""))
      End If
    End If

'''    If (iBracketLevel > 0) And sChar = "}" Then
'''      iBracketLevel = iBracketLevel - 1
'''      sRTFCodeToDo = sRTFCode
'''      sRTFCode = ""
'''    End If

    If (Not fLiteral) Then
      If (sChar = "\") Then
        fDoingSlash = True
        sRTFCode = sChar
'''      ElseIf Not fIgnoreChar Then
''''''        asText(iTextIndex) = asText(iTextIndex) & sChar
      End If
    Else
''''''      asText(iTextIndex) = asText(iTextIndex) & sNextChar
    End If

    ' See if we need to interpret the RTF control code.
    If Len(sRTFCodeToDo) > 0 Then
      If ((sRTFCodeToDo = "\ul") And (iTextIndex = 0)) _
        Or ((sRTFCodeToDo = "\ulnone") And (iTextIndex = 1)) Then
        iTextIndex = iTextIndex + 1
'''      ElseIf (sRTFCodeToDo = "\tab") Or (sRTFCodeToDo = "\cell") Then
'''        asText(iTextIndex) = asText(iTextIndex) & vbTab
'''      ElseIf (sRTFCodeToDo = "\row") Then
'''        asText(iTextIndex) = asText(iTextIndex) & vbNewLine
'''      ElseIf (Mid(sRTFCodeToDo, 1, 2) = "\'") Then
'''        fFound = False
'''        sDeniedChar = Chr(Val("&H" & Mid(sRTFCodeToDo, 3)))
'''        For iLoop = 1 To UBound(asDeniedCharacters)
'''          If sDeniedChar = asDeniedCharacters(iLoop) Then
'''            fFound = True
'''            Exit For
'''          End If
'''        Next iLoop
'''        If Not fFound Then
'''          ReDim Preserve asDeniedCharacters(UBound(asDeniedCharacters) + 1)
'''          asDeniedCharacters(UBound(asDeniedCharacters)) = sDeniedChar
'''        End If
      End If

      sRTFCodeToDo = ""
    End If

    If (Not fLiteral) Then
      If (sChar = "\") Then
''''''        fDoingSlash = True
''''''        sRTFCode = sChar
      ElseIf sChar = "{" Then
''''''        iBracketLevel = iBracketLevel + 1
''''''        sRTFCodeToDo = sRTFCode
''''''        sRTFCode = ""
      ElseIf Not fIgnoreChar Then
        asText(iTextIndex) = asText(iTextIndex) & sChar
      End If
    Else
      asText(iTextIndex) = asText(iTextIndex) & sNextChar
fDoingSlash = False
    End If

    ' Move forward through the text (jump an extra character if we've just processed a literal.
    If fLiteral Then
      sSourceText = Mid(sSourceText, 3)
    Else
      sSourceText = Mid(sSourceText, 2)
    End If
  Loop

  psPart1 = asText(0)
  psPart2 = asText(1)
  psPart3 = asText(2)
  
End Sub


Public Function ExpressionUsedInWorkflow(plngExprID As Long) As Boolean
  ' Return TRUE if the given Expression is actually used in the associated Workflow.
  ' ie. If it is used as a ...
  '   1) Decision - True flow calculation
  '   2) Web Form - Description calculation
  '   3) Web Form - Record Selector filter
  '   4) Web Form - Label calculation
  '   5) Web Form - Validation calculation
  '   6) Web Form - Default value calculation
  '   7) Stored Data - Calculation
  '   8) Email - Calculation
  '   9) Calc/filter in any other calc/filter that is used in the Workflow.
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim fUsed As Boolean
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  
  fUsed = False
  
  '   1) Decision - true flow calculation
  '   2) Web Form - Description calculation
  sSQL = "SELECT ID" & _
    " FROM tmpWorkflowElements" & _
    " WHERE trueFlowExprID = " & CStr(plngExprID) & _
    "   OR descriptionExprID = " & CStr(plngExprID)
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsTemp
    fUsed = Not (.BOF And .EOF)
  End With
  Set rsTemp = Nothing
  
  If Not fUsed Then
    '   3) Web Form - Record Selector filter (recordFilterID)
    '   4) Web Form - Label calculation (calcID)
    '   6) Web Form - Default value calculation (calcID)
    '   8) Email - Calculation (calcID)
    sSQL = "SELECT ID" & _
      " FROM tmpWorkflowElementItems" & _
      " WHERE recordFilterID = " & CStr(plngExprID) & _
      "   OR calcID = " & CStr(plngExprID)
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsTemp
      fUsed = Not (.BOF And .EOF)
    End With
    Set rsTemp = Nothing
  End If
  
  If Not fUsed Then
    '   5) Web Form - Validation calculation
    sSQL = "SELECT ID" & _
      " FROM tmpWorkflowElementValidations" & _
      " WHERE exprID = " & CStr(plngExprID)
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsTemp
      fUsed = Not (.BOF And .EOF)
    End With
    Set rsTemp = Nothing
  End If
  
  If Not fUsed Then
    '   7) Stored Data - Calculation
    sSQL = "SELECT ID" & _
      " FROM tmpWorkflowElementColumns" & _
      " WHERE calcID = " & CStr(plngExprID)
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsTemp
      fUsed = Not (.BOF And .EOF)
    End With
    Set rsTemp = Nothing
  End If
    
  If Not fUsed Then
    '   9) Calc/filter in any other calc/filter that is used in the Workflow.
    sSQL = "SELECT tmpComponents.componentID" & _
      " FROM tmpComponents " & _
      " WHERE (tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
      "   OR tmpComponents.type = " & Trim(Str(giCOMPONENT_WORKFLOWFIELD)) & ")" & _
      " AND tmpComponents.fieldSelectionFilter = " & CStr(plngExprID)

    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsTemp
      Do While Not .EOF
        Set objComp = New CExprComponent
        objComp.ComponentID = rsTemp!ComponentID
        lngExprID = objComp.RootExpressionID
        Set objComp = Nothing
        
        If lngExprID > 0 Then
          fUsed = ExpressionUsedInWorkflow(lngExprID)
        End If
      
        If fUsed Then
          Exit Do
        End If
          
        .MoveNext
      Loop
    End With
    Set rsTemp = Nothing
  End If
    
  ExpressionUsedInWorkflow = fUsed
  
End Function

Public Function GetWebFormItemTypeName(piInputType As Integer) As String
  
  On Error GoTo ErrorTrap
  
  GetWebFormItemTypeName = vbNullString
  
  Select Case piInputType
    Case giWFFORMITEM_FORM
      GetWebFormItemTypeName = "Web Form"
    Case giWFFORMITEM_UNKNOWN
      GetWebFormItemTypeName = "Unknown"
    Case giWFFORMITEM_BUTTON
      GetWebFormItemTypeName = "Button"
    Case giWFFORMITEM_DBVALUE
      GetWebFormItemTypeName = "Database Value"
    Case giWFFORMITEM_LABEL
      GetWebFormItemTypeName = "Label"
    Case giWFFORMITEM_INPUTVALUE_CHAR
      GetWebFormItemTypeName = "Character"
    Case giWFFORMITEM_WFVALUE
      GetWebFormItemTypeName = "Workflow Value"
    Case giWFFORMITEM_INPUTVALUE_NUMERIC
      GetWebFormItemTypeName = "Numeric"
    Case giWFFORMITEM_INPUTVALUE_LOGIC
      GetWebFormItemTypeName = "Logic"
    Case giWFFORMITEM_INPUTVALUE_DATE
      GetWebFormItemTypeName = "Date"
    Case giWFFORMITEM_FRAME
      GetWebFormItemTypeName = "Frame"
    Case giWFFORMITEM_LINE
      GetWebFormItemTypeName = "Line"
    Case giWFFORMITEM_IMAGE
      GetWebFormItemTypeName = "Image"
    Case giWFFORMITEM_INPUTVALUE_GRID
      GetWebFormItemTypeName = "Record Selector"
    Case giWFFORMITEM_FORMATCODE
      GetWebFormItemTypeName = "Format Code"
    Case giWFFORMITEM_INPUTVALUE_DROPDOWN
      GetWebFormItemTypeName = "Dropdown"
    Case giWFFORMITEM_INPUTVALUE_LOOKUP
      GetWebFormItemTypeName = "Lookup"
    Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
      GetWebFormItemTypeName = "Option Group"
    Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
      GetWebFormItemTypeName = "File Upload"
    Case giWFFORMITEM_DBFILE
      GetWebFormItemTypeName = "Database Value"
    Case giWFFORMITEM_WFFILE
      GetWebFormItemTypeName = "Workflow Value"
End Select

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function




Private Function BaseTableColumnsUsedInDeleteTriggeredWorkflow(plngWorkflowID As Long) As Variant
  ' Return an array of the IDs of the base table columns used in the given workflow
  Dim alngColumnsUsed() As Long
  Dim lngBaseTableID As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim alngExprColumnsUsed() As Long
  
  ReDim alngColumnsUsed(0)

  lngBaseTableID = 0
  With recWorkflowEdit
    .Index = "idxWorkflowID"
    .Seek "=", plngWorkflowID

    If Not .NoMatch Then
      lngBaseTableID = !BaseTable
    End If
  End With
  
  If lngBaseTableID > 0 Then
    ' JPD - Email addresses handled on their own now.
    '  ----------------------------------------------------------------------------
    '  -- Determine which fields from the Deleted record are used in Email elements
    '  -- 1) Email items
    '  ----------------------------------------------------------------------------
    sSQL = "SELECT tmpWorkflowElementItems.dbColumnID" & _
      " FROM tmpWorkflowElementItems," & _
      "   tmpWorkflowElements," & _
      "   tmpColumns" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID) & _
      "   AND tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
      "   AND tmpWorkflowElementItems.dbColumnID = tmpColumns.columnID" & _
      "   AND tmpWorkflowElements.type = 3" & _
      "   AND tmpWorkflowElementItems.itemType = 1" & _
      "   AND tmpColumns.tableID = " & CStr(lngBaseTableID) & _
      "   AND tmpWorkflowElementItems.dbRecord = 4"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsTemp
      ' For each element item definition ...
      Do While (Not .EOF)
        fFound = False
        
        For lngLoop = 1 To UBound(alngColumnsUsed)
          If alngColumnsUsed(lngLoop) = !DBColumnID Then
            fFound = True
            Exit For
          End If
        Next lngLoop
      
        If Not fFound Then
          ReDim Preserve alngColumnsUsed(UBound(alngColumnsUsed) + 1)
          alngColumnsUsed(UBound(alngColumnsUsed)) = !DBColumnID
        End If
        
        .MoveNext
      Loop
    End With
    Set rsTemp = Nothing

    '  ----------------------------------------------------------------------------
    '  -- Determine which fields from the Deleted record are used in WebForm elements
    '  -- 1) WebForm DBValues
    '  ----------------------------------------------------------------------------
    sSQL = "SELECT tmpWorkflowElementItems.dbColumnID" & _
      " FROM tmpWorkflowElementItems," & _
      "   tmpWorkflowElements," & _
      "   tmpColumns" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID) & _
      "   AND tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
      "   AND tmpWorkflowElementItems.dbColumnID = tmpColumns.columnID" & _
      "   AND tmpWorkflowElements.type = 2" & _
      "   AND tmpWorkflowElementItems.itemType = 1" & _
      "   AND tmpColumns.tableID = " & CStr(lngBaseTableID) & _
      "   AND tmpWorkflowElementItems.dbRecord = 4"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsTemp
      ' For each element item definition ...
      Do While (Not .EOF)
        fFound = False
        
        For lngLoop = 1 To UBound(alngColumnsUsed)
          If alngColumnsUsed(lngLoop) = !DBColumnID Then
            fFound = True
            Exit For
          End If
        Next lngLoop
      
        If Not fFound Then
          ReDim Preserve alngColumnsUsed(UBound(alngColumnsUsed) + 1)
          alngColumnsUsed(UBound(alngColumnsUsed)) = !DBColumnID
        End If
        
        .MoveNext
      Loop
    End With
    Set rsTemp = Nothing

    '  ----------------------------------------------------------------------------
    '  -- Determine which fields from the Deleted record are used in StoredData elements
    '  -- 1) StoredData DBValues
    '  ----------------------------------------------------------------------------
    sSQL = "SELECT tmpWorkflowElementColumns.dbColumnID" & _
      " FROM tmpWorkflowElementColumns," & _
      "   tmpWorkflowElements," & _
      "   tmpColumns" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID) & _
      "   AND tmpWorkflowElementColumns.elementID = tmpWorkflowElements.ID" & _
      "   AND tmpWorkflowElementColumns.dbColumnID = tmpColumns.columnID" & _
      "   AND tmpWorkflowElements.type = 5" & _
      "   AND tmpWorkflowElementColumns.valueType = 2" & _
      "   AND tmpColumns.tableID = " & CStr(lngBaseTableID) & _
      "   AND tmpWorkflowElementColumns.dbRecord = 4"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsTemp
      ' For each element item definition ...
      Do While (Not .EOF)
        fFound = False
        
        For lngLoop = 1 To UBound(alngColumnsUsed)
          If alngColumnsUsed(lngLoop) = !DBColumnID Then
            fFound = True
            Exit For
          End If
        Next lngLoop
      
        If Not fFound Then
          ReDim Preserve alngColumnsUsed(UBound(alngColumnsUsed) + 1)
          alngColumnsUsed(UBound(alngColumnsUsed)) = !DBColumnID
        End If
        
        .MoveNext
      Loop
    End With
    Set rsTemp = Nothing
  
  End If
  
  '  ----------------------------------------------------------------------------
  '  -- Return a recordset of the columns in the deleted record's table that are used
  '  -- elsewhere in the Workflow
  '  ----------------------------------------------------------------------------
  BaseTableColumnsUsedInDeleteTriggeredWorkflow = alngColumnsUsed
  
End Function


Private Function BaseTableEmailAddressesUsedInDeleteTriggeredWorkflow(plngWorkflowID As Long) As Variant
  ' Return an array of the IDs of the base table columns used in the given workflow
  Dim alngEmailsUsed() As Long
  Dim lngBaseTableID As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  
  ReDim alngEmailsUsed(2, 0)
  ' Column 0 = emailID
  ' Column 1 = type (1=Column, 2=Calculated)
  ' Column 2 = column/expr ID
  lngBaseTableID = 0
  With recWorkflowEdit
    .Index = "idxWorkflowID"
    .Seek "=", plngWorkflowID

    If Not .NoMatch Then
      lngBaseTableID = !BaseTable
    End If
  End With
  
  If lngBaseTableID > 0 Then
    '  ----------------------------------------------------------------------------
    '  -- Determine which fields from the Deleted record are used in Email elements
    '  -- 1) Email To
    '  -- 2) Email Copy
    '  ----------------------------------------------------------------------------
    sSQL = "SELECT tmpEmailAddresses.emailID," & _
      "   tmpEmailAddresses.type," & _
      "   tmpEmailAddresses.columnID," & _
      "   tmpEmailAddresses.exprID" & _
      " FROM tmpWorkflowElements," & _
      "   tmpEmailAddresses" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID) & _
      "   AND (tmpWorkflowElements.emailID = tmpEmailAddresses.emailID" & _
      "     OR tmpWorkflowElements.emailCCID = tmpEmailAddresses.emailID)" & _
      "   AND tmpWorkflowElements.type = 3" & _
      "   AND tmpEmailAddresses.tableID = " & CStr(lngBaseTableID) & _
      "   AND tmpWorkflowElements.emailRecord = 4" & _
      "   AND ((tmpEmailAddresses.type = 1) OR (tmpEmailAddresses.type = 2))"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    With rsTemp
      ' For each element item definition ...
      Do While (Not .EOF)
        fFound = False
  
        For lngLoop = 1 To UBound(alngEmailsUsed, 2)
          If alngEmailsUsed(0, lngLoop) = !EmailID Then
            fFound = True
            Exit For
          End If
        Next lngLoop
  
        If Not fFound Then
          ReDim Preserve alngEmailsUsed(2, UBound(alngEmailsUsed, 2) + 1)
          alngEmailsUsed(0, UBound(alngEmailsUsed, 2)) = !EmailID
          alngEmailsUsed(1, UBound(alngEmailsUsed, 2)) = !Type
          If !Type = 1 Then
            ' Column
            alngEmailsUsed(2, UBound(alngEmailsUsed, 2)) = !ColumnID
          Else
            ' Calculated
            alngEmailsUsed(2, UBound(alngEmailsUsed, 2)) = !ExprID
          End If
        End If

        .MoveNext
      Loop
    End With
    Set rsTemp = Nothing

  End If
  
  '  ----------------------------------------------------------------------------
  '  -- Return a recordset of the email addresses based on the deleted record's
  '  -- table that are used in the Workflow
  '  ----------------------------------------------------------------------------
  BaseTableEmailAddressesUsedInDeleteTriggeredWorkflow = alngEmailsUsed
  
End Function



Private Function ColumnsUsedInExpression(plngExprID As Long) As Variant
  ' Return an array of the IDs of the columns used in the given expression
  Dim alngColumnsUsed() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim rsTemp2 As DAO.Recordset
  Dim alngExprColumnsUsed() As Long

  ReDim alngColumnsUsed(0)

  ' Record the columns used by field components.
  sSQL = "SELECT tmpComponents.fieldColumnID" & _
    " FROM tmpComponents" & _
    " WHERE tmpComponents.exprID = " & CStr(plngExprID) & _
    "   AND tmpComponents.type = " & CStr(giCOMPONENT_FIELD)
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  With rsTemp
    ' For each element item definition ...
    Do While (Not .EOF)
      fFound = False

      For lngLoop = 1 To UBound(alngColumnsUsed)
        If alngColumnsUsed(lngLoop) = !fieldColumnID Then
          fFound = True
          Exit For
        End If
      Next lngLoop

      If Not fFound Then
        ReDim Preserve alngColumnsUsed(UBound(alngColumnsUsed) + 1)
        alngColumnsUsed(UBound(alngColumnsUsed)) = !fieldColumnID
      End If

      .MoveNext
    Loop
  End With
  Set rsTemp = Nothing

  ' Check sub-expressions.
  sSQL = "SELECT tmpComponents.type," & _
    "   fieldSelectionFilter," & _
    "   componentID," & _
    "   calculationID," & _
    "   filterID" & _
    " FROM tmpComponents" & _
    " WHERE tmpComponents.exprID = " & CStr(plngExprID) & _
    "   AND ((tmpComponents.type = " & CStr(giCOMPONENT_FIELD) & " AND tmpComponents.fieldSelectionFilter > 0)" & _
    "     OR (tmpComponents.type = " & CStr(giCOMPONENT_FUNCTION) & ")" & _
    "     OR (tmpComponents.type = " & CStr(giCOMPONENT_CALCULATION) & ")" & _
    "     OR (tmpComponents.type = " & CStr(giCOMPONENT_FILTER) & "))"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  With rsTemp
    ' For each element item definition ...
    Do While (Not .EOF)

      If !Type = giCOMPONENT_FUNCTION Then
        ' Get the columns used in as follows:
        ' 1) Function component sub-expressions
        sSQL = "SELECT tmpExpressions.exprID" & _
          " FROM tmpExpressions" & _
          " WHERE tmpExpressions.parentComponentID = " & CStr(!ComponentID)
        Set rsTemp2 = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

        With rsTemp2
          ' For each element item definition ...
          Do While (Not .EOF)
            alngExprColumnsUsed = ColumnsUsedInExpression(!ExprID)
            
            For lngLoop2 = 1 To UBound(alngExprColumnsUsed)
              fFound = False
          
              For lngLoop = 1 To UBound(alngColumnsUsed)
                If alngColumnsUsed(lngLoop) = alngExprColumnsUsed(lngLoop2) Then
                  fFound = True
                  Exit For
                End If
              Next lngLoop
          
              If Not fFound Then
                ReDim Preserve alngColumnsUsed(UBound(alngColumnsUsed) + 1)
                alngColumnsUsed(UBound(alngColumnsUsed)) = alngExprColumnsUsed(lngLoop2)
              End If
            Next lngLoop2
            
            .MoveNext
          Loop
        End With
        Set rsTemp2 = Nothing
      Else
        ' Get the columns used in as follows:
        ' 1) Field component filters
        ' 2) Calculation components
        ' 3) Filter components
        Select Case !Type
          Case giCOMPONENT_FIELD
            alngExprColumnsUsed = ColumnsUsedInExpression(!FieldSelectionFilter)
          Case giCOMPONENT_CALCULATION
            alngExprColumnsUsed = ColumnsUsedInExpression(!CalculationID)
          Case giCOMPONENT_FILTER
            alngExprColumnsUsed = ColumnsUsedInExpression(!FilterID)
          Case Else
            ReDim alngExprColumnsUsed(0)
        End Select
        
        For lngLoop2 = 1 To UBound(alngExprColumnsUsed)
          fFound = False
      
          For lngLoop = 1 To UBound(alngColumnsUsed)
            If alngColumnsUsed(lngLoop) = alngExprColumnsUsed(lngLoop2) Then
              fFound = True
              Exit For
            End If
          Next lngLoop
      
          If Not fFound Then
            ReDim Preserve alngColumnsUsed(UBound(alngColumnsUsed) + 1)
            alngColumnsUsed(UBound(alngColumnsUsed)) = alngExprColumnsUsed(lngLoop2)
          End If
        Next lngLoop2
      End If

      .MoveNext
    Loop
  End With
  Set rsTemp = Nothing
  
  '  ----------------------------------------------------------------------------
  '  -- Return a recordset of the columns used in the given expression
  '  ----------------------------------------------------------------------------
  ColumnsUsedInExpression = alngColumnsUsed

End Function

Public Sub TableAscendants(plngTableID As Long, palngAscendants As Variant)
  ' Populate the array with the ascendants of the given table.
  ' NB. The given table is itself entered into the array at location (1)
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  
  ' Check if the table has already been added, and its ascendants already calculated.
  fFound = False
  
  For lngLoop = 1 To UBound(palngAscendants)
    If palngAscendants(lngLoop) = plngTableID Then
      fFound = True
      Exit For
    End If
  Next lngLoop
  
  If Not fFound Then
    ReDim Preserve palngAscendants(UBound(palngAscendants) + 1)
    palngAscendants(UBound(palngAscendants)) = plngTableID
  
    ' Get the given table's parents.
    sSQL = "SELECT parentID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.childID = " & CStr(plngTableID)
      
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    Do While Not rsTables.EOF
      TableAscendants rsTables!parentID, palngAscendants
      rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
  End If

End Sub


Public Sub TableDescendants(plngTableID As Long, palngDescendants As Variant)
  ' Populate the array with the descendants of the given table.
  ' NB. The given table is itself entered into the array at location (1)
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  
  ' Check if the table has already been added, and its descendants already calculated.
  fFound = False
  
  For lngLoop = 1 To UBound(palngDescendants)
    If palngDescendants(lngLoop) = plngTableID Then
      fFound = True
      Exit For
    End If
  Next lngLoop
  
  If Not fFound Then
    ReDim Preserve palngDescendants(UBound(palngDescendants) + 1)
    palngDescendants(UBound(palngDescendants)) = plngTableID
  
    ' Get the given table's children.
    sSQL = "SELECT childID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.parentID = " & CStr(plngTableID)
      
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    Do While Not rsTables.EOF
      TableDescendants rsTables!childID, palngDescendants
      rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
  End If

End Sub



Public Function SaveWorkflowLinks(lngTableID As Long) As Boolean

  Dim rsWorkflowLinks As ADODB.Recordset
  Dim rsWorkflowLinkColumns As ADODB.Recordset

  Set rsWorkflowLinks = New ADODB.Recordset
  Set rsWorkflowLinkColumns = New ADODB.Recordset

  rsWorkflowLinks.Open "SELECT * FROM ASRSysWorkflowTriggeredLinks", gADOCon, adOpenDynamic, adLockOptimistic
  rsWorkflowLinkColumns.Open "SELECT * FROM ASRSysWorkflowTriggeredLinkColumns", gADOCon, adOpenDynamic, adLockOptimistic

  With recWorkflowTriggeredLinks
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do While Not .EOF
        If !TableID = lngTableID Then
          If Not !Deleted Then
            rsWorkflowLinks.AddNew

            rsWorkflowLinks!LinkID = !LinkID
            rsWorkflowLinks!TableID = !TableID
            rsWorkflowLinks!FilterID = !FilterID
            rsWorkflowLinks!EffectiveDate = !EffectiveDate
            rsWorkflowLinks!Type = !Type
            rsWorkflowLinks!RecordInsert = !RecordInsert
            rsWorkflowLinks!RecordUpdate = !RecordUpdate
            rsWorkflowLinks!RecordDelete = !RecordDelete
            rsWorkflowLinks!DateColumn = !DateColumn
            rsWorkflowLinks!DateOffset = !DateOffset
            rsWorkflowLinks!DateOffsetPeriod = !DateOffsetPeriod
            rsWorkflowLinks!WorkflowID = !WorkflowID

            rsWorkflowLinks.Update
            rsWorkflowLinks.MoveLast

            With recWorkflowTriggeredLinkColumns
              If Not (.BOF And .EOF) Then
                .MoveFirst

                Do While Not .EOF
                  If !LinkID = recWorkflowTriggeredLinks!LinkID Then
                    rsWorkflowLinkColumns.AddNew

                    rsWorkflowLinkColumns!LinkID = !LinkID
                    rsWorkflowLinkColumns!ColumnID = !ColumnID
                    
                    rsWorkflowLinkColumns.Update
                  End If

                  .MoveNext
                Loop
              End If
            End With
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  rsWorkflowLinks.Close
  rsWorkflowLinkColumns.Close
  SaveWorkflowLinks = True

TidyUpAndExit:
  Set rsWorkflowLinks = Nothing
  Set rsWorkflowLinkColumns = Nothing

  Exit Function

LocalErr:
  SaveWorkflowLinks = False
  Resume TidyUpAndExit

End Function



Public Function GetRecordSelectionFromDescription(psRecordSelectionDescription As String) As WorkflowRecordSelectorTypes
  On Error GoTo ErrorTrap
  
  Dim iRecordSelection As WorkflowRecordSelectorTypes
  
  Select Case psRecordSelectionDescription
    Case GetRecordSelectionDescription(giWFRECSEL_INITIATOR):
      iRecordSelection = giWFRECSEL_INITIATOR
    Case GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD):
      iRecordSelection = giWFRECSEL_IDENTIFIEDRECORD
    Case GetRecordSelectionDescription(giWFRECSEL_ALL):
      iRecordSelection = giWFRECSEL_ALL
    Case GetRecordSelectionDescription(giWFRECSEL_UNIDENTIFIED):
      iRecordSelection = giWFRECSEL_UNIDENTIFIED
    Case GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
      iRecordSelection = giWFRECSEL_TRIGGEREDRECORD
    Case Else:
      iRecordSelection = giWFRECSEL_UNKNOWN
  End Select

TidyUpAndExit:
  GetRecordSelectionFromDescription = iRecordSelection
  Exit Function
  
ErrorTrap:
  iRecordSelection = giWFRECSEL_UNKNOWN
  Resume TidyUpAndExit
  
End Function



Public Function GetRecordSelectionDescription(piRecordSelection As WorkflowRecordSelectorTypes) As String
  On Error GoTo ErrorTrap
  
  Dim sDescription As String
  
  Const sDBRECORD_UNKNOWN_TEXT = "<unknown>"
  Const sDBRECORD_INITIATOR_TEXT = "Initiator's Record"
  Const sDBRECORD_PREV_RECORD_SELECTION = "Identified Record"
  Const sDBRECORD_ALL_TEXT = "All Records"
  Const sDBRECORD_UNIDENTIFIED_TEXT = "Unidentified"
  Const sDBRECORD_TRIGGEREDRECORD_TEXT = "Triggered Record"
  
  Select Case piRecordSelection
    Case giWFRECSEL_INITIATOR:
      sDescription = sDBRECORD_INITIATOR_TEXT
    Case giWFRECSEL_IDENTIFIEDRECORD:
      sDescription = sDBRECORD_PREV_RECORD_SELECTION
    Case giWFRECSEL_ALL:
      sDescription = sDBRECORD_ALL_TEXT
    Case giWFRECSEL_UNIDENTIFIED:
      sDescription = sDBRECORD_UNIDENTIFIED_TEXT
    Case giWFRECSEL_TRIGGEREDRECORD
      sDescription = sDBRECORD_TRIGGEREDRECORD_TEXT
    Case Else:
      sDescription = sDBRECORD_UNKNOWN_TEXT
  End Select

TidyUpAndExit:
  GetRecordSelectionDescription = sDescription
  Exit Function
  
ErrorTrap:
  sDescription = sDBRECORD_UNKNOWN_TEXT
  Resume TidyUpAndExit
  
End Function




Public Function GetDecisionCaptionDescription(piCaptionType As DecisionCaptionType, _
  pfTrueFlow As Boolean) As String
  
  On Error GoTo ErrorTrap
  
  Dim sDescription As String

  Const sDECISIONCAPTION_UNKNOWN = "<unknown>"
  
  Const sDECISIONCAPTION_YES = "Yes"
  Const sDECISIONCAPTION_NO = "No"
  
  Const sDECISIONCAPTION_1 = "1"
  Const sDECISIONCAPTION_0 = "0"
  
  Const sDECISIONCAPTION_TICK = "Tick"
  Const sDECISIONCAPTION_CROSS = "Cross"
  
  Const sDECISIONCAPTION_TRUE = "True"
  Const sDECISIONCAPTION_FALSE = "False"

  Select Case piCaptionType
    Case decisionCaption_Y_N
      sDescription = IIf(pfTrueFlow, sDECISIONCAPTION_YES, sDECISIONCAPTION_NO)
    Case decisionCaption_1_0
      sDescription = IIf(pfTrueFlow, sDECISIONCAPTION_1, sDECISIONCAPTION_0)
    Case decisionCaption_tick_cross
      sDescription = IIf(pfTrueFlow, sDECISIONCAPTION_TICK, sDECISIONCAPTION_CROSS)
    Case Else
      sDescription = IIf(pfTrueFlow, sDECISIONCAPTION_TRUE, sDECISIONCAPTION_FALSE)
  End Select
    
TidyUpAndExit:
  GetDecisionCaptionDescription = sDescription
  Exit Function
  
ErrorTrap:
  sDescription = sDECISIONCAPTION_UNKNOWN
  Resume TidyUpAndExit
  
End Function





Public Function GetWorkflowName(plngWorkflowID As Long) As String
  'Return the name of the given workflow.
  On Error GoTo ErrorTrap
  
  Dim sName As String
  
  sName = "<unknown>"
  
  With recWorkflowEdit
    .Index = "idxWorkflowID"
    .Seek "=", plngWorkflowID

    If Not .NoMatch Then
      sName = !Name
    End If
  End With
  
TidyUpAndExit:
  GetWorkflowName = sName
  Exit Function
  
ErrorTrap:
  sName = ""
  Resume TidyUpAndExit
  
End Function

Public Function GetWorkflowEnabled(plngWorkflowID As Long) As Boolean
  'Return the 'enabled' value of the given workflow.
  On Error GoTo ErrorTrap
  
  Dim fEnabled As Boolean
  
  fEnabled = False
  
  With recWorkflowEdit
    .Index = "idxWorkflowID"
    .Seek "=", plngWorkflowID

    If Not .NoMatch Then
      fEnabled = !Enabled
    End If
  End With
  
TidyUpAndExit:
  GetWorkflowEnabled = fEnabled
  Exit Function
  
ErrorTrap:
  fEnabled = False
  Resume TidyUpAndExit
  
End Function


Public Function CloneWorkflow(plngWorkflowID As Long, _
  pavCloneRegister As Variant) As Boolean
  ' Clone the current expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fGoodName As Boolean
  Dim iIndex As Integer
  Dim iCounter As Integer
  Dim lngWorkflowID As Long
  Dim lngNewID As Long
  Dim lngTempNewID As Long
  Dim sSQL As String
  Dim sSQL2 As String
  Dim sWorkflowName As String
  Dim rsWorkflow As DAO.Recordset
  Dim rsElements As DAO.Recordset
  Dim rsLinks As DAO.Recordset
  Dim rsElementItems As DAO.Recordset
  Dim rsElementItemValues As DAO.Recordset
  Dim rsElementColumns As DAO.Recordset
  Dim rsElementValidations As DAO.Recordset
  Dim alngElementIDs() As Long
  Dim iLoop As Integer
  Dim rsExpressions As DAO.Recordset
  Dim rsFilters As DAO.Recordset
  Dim objSourceExpr As CExpression
  Dim objNewExpr As CExpression
  
  ReDim alngElementIDs(1, 0)
  
  sSQL = "SELECT *" & _
    " FROM tmpWorkflows" & _
    " WHERE tmpWorkflows.ID = " & Trim(Str(plngWorkflowID))
  Set rsWorkflow = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  fOK = Not (rsWorkflow.BOF And rsWorkflow.EOF)
  
  If fOK Then
    ' Create a new workflow name.
    sWorkflowName = "Copy_of_" & rsWorkflow!Name
    ' Check that the workflow name is not already used.
    iCounter = 1
    fGoodName = False
    Do While Not fGoodName
      With recWorkflowEdit
        .Index = "idxName"
        .Seek "=", sWorkflowName, False
        If Not .NoMatch Then
          iCounter = iCounter + 1
          sWorkflowName = "Copy_" & Trim(Str(iCounter)) & "_of_" & rsWorkflow!Name
        Else
          fGoodName = True
        End If
      End With
    Loop
    ' Get a unique ID for the new record.
    lngWorkflowID = UniqueColumnValue("tmpWorkflows", "ID")

    ' Add a new record in the database for the copied screen definition.
    recWorkflowEdit.AddNew
  
    recWorkflowEdit!ID = lngWorkflowID
    recWorkflowEdit!Changed = False
    recWorkflowEdit!New = True
    recWorkflowEdit!Deleted = False
    recWorkflowEdit!Name = sWorkflowName
    recWorkflowEdit!Description = rsWorkflow.Fields("Description")
    recWorkflowEdit!Enabled = rsWorkflow.Fields("enabled")
    recWorkflowEdit!InitiationType = rsWorkflow.Fields("initiationType")
    
    lngNewID = IIf(IsNull(rsWorkflow.Fields("baseTable")), 0, rsWorkflow.Fields("baseTable"))
    If lngNewID > 0 Then
      For iIndex = 1 To UBound(pavCloneRegister, 2)
        If pavCloneRegister(1, iIndex) = "TABLE" And _
          pavCloneRegister(2, iIndex) = rsWorkflow.Fields("baseTable") Then
          lngNewID = pavCloneRegister(3, iIndex)
          Exit For
        End If
      Next iIndex
    End If
    recWorkflowEdit!BaseTable = lngNewID

    recWorkflowEdit.Update

    ' Remember the IDs of the original and copied workflows.
    iIndex = UBound(pavCloneRegister, 2) + 1
    ReDim Preserve pavCloneRegister(3, iIndex)
    pavCloneRegister(1, iIndex) = "WORKFLOW"
    pavCloneRegister(2, iIndex) = plngWorkflowID
    pavCloneRegister(3, iIndex) = lngWorkflowID
        
    sSQL = "SELECT tmpExpressions.exprID" & _
      " FROM tmpExpressions" & _
      " WHERE tmpExpressions.deleted = FALSE" & _
      " AND tmpExpressions.utilityID = " & Trim(Str(plngWorkflowID)) & _
      " AND tmpExpressions.parentComponentID = 0 " & _
      " AND (tmpExpressions.Type = " & CStr(giEXPR_WORKFLOWCALCULATION) & _
      "   OR tmpExpressions.Type = " & CStr(giEXPR_WORKFLOWSTATICFILTER) & _
      "   OR tmpExpressions.Type = " & CStr(giEXPR_WORKFLOWRUNTIMEFILTER) & ")"
    Set rsExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsExpressions
      ' For each expression definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original expression object.
        Set objSourceExpr = New CExpression
        objSourceExpr.ExpressionID = .Fields("exprID")

        Set objNewExpr = objSourceExpr.CloneExpression(pavCloneRegister)
        fOK = Not objNewExpr Is Nothing

        If fOK Then
          ' Copy properties from the original expression to the copy.
          objNewExpr.UtilityID = lngWorkflowID
          ' Write the copied expession definition to the database.
          fOK = objNewExpr.WriteExpression
        End If

        ' Remember the IDs of the original and copied orders.
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "EXPRESSION"
        pavCloneRegister(2, iIndex) = objSourceExpr.ExpressionID

        If fOK Then
          pavCloneRegister(3, iIndex) = objNewExpr.ExpressionID
        Else
          pavCloneRegister(3, iIndex) = 0

          fOK = True
        End If

        ' Disassociate object variables.
        Set objSourceExpr = Nothing
        Set objNewExpr = Nothing

        .MoveNext
      Loop

      .Close
    End With
    ' Disassociate object variables.
    Set rsExpressions = Nothing

    If fOK Then
      ' Update the copied expression field components with the new IDs of their filter expressions.
      sSQL = "SELECT tmpComponents.componentID, tmpComponents.fieldSelectionFilter" & _
        " FROM tmpComponents, tmpExpressions " & _
        " WHERE tmpExpressions.utilityID = " & Trim(Str(lngWorkflowID)) & _
        " AND tmpComponents.exprID = tmpExpressions.exprID" & _
        " AND (tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
        "   OR tmpComponents.type = " & Trim(Str(giCOMPONENT_WORKFLOWFIELD)) & ")" & _
        " AND tmpComponents.fieldSelectionFilter > 0"
      Set rsFilters = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      With rsFilters
        Do While (Not .EOF)
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = .Fields("fieldSelectionFilter") Then

              recCompEdit.Index = "idxCompID"
              recCompEdit.Seek "=", .Fields("componentID")
              If Not recCompEdit.NoMatch Then
                recCompEdit.Edit
                recCompEdit.Fields("fieldSelectionFilter") = pavCloneRegister(3, iIndex)
                recCompEdit.Update
              End If
              Exit For
            End If
          Next iIndex

          .MoveNext
        Loop

        .Close
      End With

      ' Disassociate object variables.
      Set rsFilters = Nothing
    End If
        
    ' Copy the workflow element definitions.
    sSQL = "SELECT *" & _
      " FROM tmpWorkflowElements" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID)
    Set rsElements = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElements
      ' For each workflow element definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied screen control definition.
        recWorkflowElementEdit.AddNew

        lngNewID = UniqueColumnValue("tmpWorkflowElements", "ID")
        recWorkflowElementEdit!ID = lngNewID

        recWorkflowElementEdit!WorkflowID = lngWorkflowID
        recWorkflowElementEdit!Type = .Fields("Type")
        recWorkflowElementEdit!Caption = .Fields("Caption")
        recWorkflowElementEdit!ConnectionPairID = .Fields("connectionPairID")
        recWorkflowElementEdit!LeftCoord = .Fields("leftCoord")
        recWorkflowElementEdit!TopCoord = .Fields("topCoord")
        recWorkflowElementEdit!Identifier = .Fields("Identifier")
        
        recWorkflowElementEdit!DecisionCaptionType = .Fields("DecisionCaptionType")
        recWorkflowElementEdit!TrueFlowType = .Fields("TrueFlowType")
        recWorkflowElementEdit!TrueFlowIdentifier = .Fields("TrueFlowIdentifier")
        

        lngTempNewID = IIf(IsNull(.Fields("TrueFlowExprID")), 0, .Fields("TrueFlowExprID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then

              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!TrueFlowExprID = lngTempNewID
        
        lngTempNewID = IIf(IsNull(.Fields("DescriptionExprID")), 0, .Fields("DescriptionExprID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then
  
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!DescriptionExprID = lngTempNewID
        
        recWorkflowElementEdit!DescHasWorkflowName = .Fields("DescHasWorkflowName")
        recWorkflowElementEdit!DescHasElementCaption = .Fields("DescHasElementCaption")
        
        recWorkflowElementEdit!DataAction = .Fields("DataAction")

        lngTempNewID = IIf(IsNull(.Fields("DataTableID")), 0, .Fields("DataTableID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "TABLE" And _
              pavCloneRegister(2, iIndex) = .Fields("DataTableID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!DataTableID = lngTempNewID
        
        recWorkflowElementEdit!DataRecord = .Fields("DataRecord")

        lngTempNewID = IIf(IsNull(.Fields("EmailID")), 0, .Fields("EmailID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EMAIL" And _
              pavCloneRegister(2, iIndex) = .Fields("EmailID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!EmailID = lngTempNewID
        
        lngTempNewID = IIf(IsNull(.Fields("EmailCCID")), 0, .Fields("EmailCCID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EMAIL" And _
              pavCloneRegister(2, iIndex) = .Fields("EmailCCID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!EmailCCID = lngTempNewID
        
        recWorkflowElementEdit!EmailRecord = .Fields("EmailRecord")
        recWorkflowElementEdit!WebFormFGColor = .Fields("WebFormFGColor")
        recWorkflowElementEdit!WebFormBGColor = .Fields("WebFormBGColor")
        recWorkflowElementEdit!WebFormBGImageID = .Fields("WebFormBGImageID")
        recWorkflowElementEdit!WebFormBGImageLocation = .Fields("WebFormBGImageLocation")
        recWorkflowElementEdit!WebFormDefaultFontName = .Fields("webFormDefaultFontName")
        recWorkflowElementEdit!WebFormDefaultFontSize = .Fields("webFormDefaultFontSize")
        recWorkflowElementEdit!WebFormDefaultFontBold = .Fields("webFormDefaultFontBold")
        recWorkflowElementEdit!WebFormDefaultFontItalic = .Fields("webFormDefaultFontItalic")
        recWorkflowElementEdit!WebFormDefaultFontStrikeThru = .Fields("webFormDefaultFontStrikeThru")
        recWorkflowElementEdit!WebFormDefaultFontUnderline = .Fields("webFormDefaultFontUnderline")
        recWorkflowElementEdit!WebFormWidth = .Fields("WebFormWidth")
        recWorkflowElementEdit!WebFormHeight = .Fields("WebFormHeight")
        recWorkflowElementEdit!RecSelWebFormIdentifier = .Fields("recSelWebFormIdentifier")
        recWorkflowElementEdit!RecSelIdentifier = .Fields("recSelIdentifier")
        
        recWorkflowElementEdit!SecondaryDataRecord = .Fields("SecondaryDataRecord")
        recWorkflowElementEdit!SecondaryRecSelWebFormIdentifier = .Fields("secondaryRecSelWebFormIdentifier")
        recWorkflowElementEdit!SecondaryRecSelIdentifier = .Fields("secondaryRecSelIdentifier")
        
        lngTempNewID = IIf(IsNull(.Fields("DataRecordTable")), 0, .Fields("DataRecordTable"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "TABLE" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!DataRecordTable = lngTempNewID

        lngTempNewID = IIf(IsNull(.Fields("secondaryDataRecordTable")), 0, .Fields("secondaryDataRecordTable"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "TABLE" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementEdit!SecondaryDataRecordTable = lngTempNewID
        recWorkflowElementEdit!UseAsTargetIdentifier = .Fields("UseAsTargetIdentifier")
        
        recWorkflowElementEdit!EMailSubject = .Fields("EmailSubject")
        recWorkflowElementEdit!TimeoutFrequency = .Fields("TimeoutFrequency")
        recWorkflowElementEdit!TimeoutPeriod = .Fields("TimeoutPeriod")
        recWorkflowElementEdit!TimeoutExcludeWeekend = .Fields("TimeoutExcludeWeekend")

        recWorkflowElementEdit!CompletionMessageType = .Fields("CompletionMessageType")
        recWorkflowElementEdit!CompletionMessage = .Fields("CompletionMessage")
        recWorkflowElementEdit!SavedForLaterMessageType = .Fields("SavedForLaterMessageType")
        recWorkflowElementEdit!SavedForLaterMessage = .Fields("SavedForLaterMessage")
        recWorkflowElementEdit!FollowOnFormsMessageType = .Fields("FollowOnFormsMessageType")
        recWorkflowElementEdit!FollowOnFormsMessage = .Fields("FollowOnFormsMessage")

        recWorkflowElementEdit.Update

        ReDim Preserve alngElementIDs(1, UBound(alngElementIDs, 2) + 1)
        alngElementIDs(0, UBound(alngElementIDs, 2)) = .Fields("ID")
        alngElementIDs(1, UBound(alngElementIDs, 2)) = lngNewID

        .MoveNext
      Loop
    End With
    Set rsElements = Nothing

    ' Ensure the connector elements have the new IDs.
    For iLoop = 1 To UBound(alngElementIDs, 2)
      sSQL = "UPDATE tmpWorkflowElements" & _
        " SET tmpWorkflowElements.connectionPairID = " & Trim(Str(alngElementIDs(1, iLoop))) & _
        " WHERE tmpWorkflowElements.workflowID = " & Trim(Str(lngWorkflowID)) & _
        "   AND tmpWorkflowElements.connectionPairID = " & Trim(Str(alngElementIDs(0, iLoop)))

      daoDb.Execute sSQL, dbFailOnError
    Next iLoop

    ' Copy the workflow element item definitions.
    sSQL = "SELECT tmpWorkflowElementItems.*" & _
      " FROM tmpWorkflowElementItems" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID)
    Set rsElementItems = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElementItems
      ' For each element item definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied element item definition.
        recWorkflowElementItemEdit.AddNew

        lngNewID = UniqueColumnValue("tmpWorkflowElementItems", "ID")
        recWorkflowElementItemEdit!ID = lngNewID

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("ElementID") Then
            recWorkflowElementItemEdit!elementid = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        recWorkflowElementItemEdit!Caption = .Fields("Caption")
        recWorkflowElementItemEdit!UseAsTargetIdentifier = .Fields("UseAsTargetIdentifier")

        lngTempNewID = IIf(IsNull(.Fields("DBColumnID")), 0, .Fields("DBColumnID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "COLUMN" And _
              pavCloneRegister(2, iIndex) = .Fields("DBColumnID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementItemEdit!DBColumnID = lngTempNewID
        
        recWorkflowElementItemEdit!DBRecord = .Fields("DBRecord")
        recWorkflowElementItemEdit!Identifier = .Fields("Identifier")
        recWorkflowElementItemEdit!InputType = .Fields("InputType")
        recWorkflowElementItemEdit!InputSize = .Fields("InputSize")
        recWorkflowElementItemEdit!InputDecimals = .Fields("InputDecimals")
        recWorkflowElementItemEdit!InputDefault = .Fields("InputDefault")
        recWorkflowElementItemEdit!WFFormIdentifier = .Fields("WFFormIdentifier")
        recWorkflowElementItemEdit!WFValueIdentifier = .Fields("WFValueIdentifier")
        recWorkflowElementItemEdit!ItemType = .Fields("ItemType")
        recWorkflowElementItemEdit!LeftCoord = .Fields("LeftCoord")
        recWorkflowElementItemEdit!TopCoord = .Fields("TopCoord")
        recWorkflowElementItemEdit!Width = .Fields("Width")
        recWorkflowElementItemEdit!Height = .Fields("Height")
        recWorkflowElementItemEdit!BackColor = .Fields("BackColor")
        recWorkflowElementItemEdit!ForeColor = .Fields("ForeColor")
        recWorkflowElementItemEdit!FontName = .Fields("FontName")
        recWorkflowElementItemEdit!FontSize = .Fields("FontSize")
        recWorkflowElementItemEdit!FontBold = .Fields("FontBold")
        recWorkflowElementItemEdit!FontItalic = .Fields("FontItalic")
        recWorkflowElementItemEdit!FontStrikethru = .Fields("FontStrikeThru")
        recWorkflowElementItemEdit!FontUnderline = .Fields("FontUnderline")
        recWorkflowElementItemEdit!PictureID = .Fields("PictureID")
        recWorkflowElementItemEdit!PictureBorder = .Fields("PictureBorder")
        recWorkflowElementItemEdit!Alignment = .Fields("Alignment")
        recWorkflowElementItemEdit!ZOrder = .Fields("ZOrder")
        recWorkflowElementItemEdit!TabIndex = .Fields("TabIndex")
        recWorkflowElementItemEdit!BackStyle = .Fields("BackStyle")
        recWorkflowElementItemEdit!BackColorEven = .Fields("BackColorEven")
        recWorkflowElementItemEdit!BackColorOdd = .Fields("BackColorOdd")
        recWorkflowElementItemEdit!ColumnHeaders = .Fields("ColumnHeaders")
        recWorkflowElementItemEdit!ForeColorEven = .Fields("ForeColorEven")
        recWorkflowElementItemEdit!ForeColorOdd = .Fields("ForeColorOdd")
        recWorkflowElementItemEdit!HeaderBackColor = .Fields("HeaderBackColor")
        recWorkflowElementItemEdit!HeadFontName = .Fields("HeadFontName")
        recWorkflowElementItemEdit!HeadFontSize = .Fields("HeadFontSize")
        recWorkflowElementItemEdit!HeadFontBold = .Fields("HeadFontBold")
        recWorkflowElementItemEdit!HeadFontItalic = .Fields("HeadFontItalic")
        recWorkflowElementItemEdit!HeadFontStrikeThru = .Fields("HeadFontStrikeThru")
        recWorkflowElementItemEdit!HeadFontUnderline = .Fields("HeadFontUnderline")
        recWorkflowElementItemEdit!HeadLines = .Fields("Headlines")

        lngTempNewID = IIf(IsNull(.Fields("TableID")), 0, .Fields("TableID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "TABLE" And _
              pavCloneRegister(2, iIndex) = .Fields("TableID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementItemEdit!TableID = lngTempNewID
        
        recWorkflowElementItemEdit!RecSelWebFormIdentifier = .Fields("recSelWebFormIdentifier")
        recWorkflowElementItemEdit!RecSelIdentifier = .Fields("recSelIdentifier")
        recWorkflowElementItemEdit!ForeColorHighlight = .Fields("ForeColorHighlight")
        recWorkflowElementItemEdit!BackColorHighlight = .Fields("BackColorHighlight")
        recWorkflowElementItemEdit!LookupTableID = .Fields("LookupTableID")
        recWorkflowElementItemEdit!LookupColumnID = .Fields("LookupColumnID")

        lngTempNewID = IIf(IsNull(.Fields("RecordTableID")), 0, .Fields("RecordTableID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "TABLE" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementItemEdit!RecordTableID = lngTempNewID

        recWorkflowElementItemEdit!Orientation = .Fields("Orientation")
        
        lngTempNewID = IIf(IsNull(.Fields("RecordOrderID")), 0, .Fields("RecordOrderID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "ORDER" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then

              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementItemEdit!RecordOrderID = lngTempNewID

        lngTempNewID = IIf(IsNull(.Fields("RecordFilterID")), 0, .Fields("RecordFilterID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then

              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementItemEdit!RecordFilterID = lngTempNewID
        
        recWorkflowElementItemEdit!Behaviour = .Fields("behaviour")
        recWorkflowElementItemEdit!Mandatory = .Fields("mandatory")
        
        lngTempNewID = IIf(IsNull(.Fields("CalcID")), 0, .Fields("CalcID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then

              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementItemEdit!CalcID = lngTempNewID
        
        recWorkflowElementItemEdit!CaptionType = .Fields("captionType")
        recWorkflowElementItemEdit!DefaultValueType = .Fields("defaultValueType")
        
        recWorkflowElementItemEdit!VerticalOffset = .Fields("VerticalOffset")
        recWorkflowElementItemEdit!VerticalOffsetBehaviour = .Fields("VerticalOffsetBehaviour")
        recWorkflowElementItemEdit!HorizontalOffset = .Fields("HorizontalOffset")
        recWorkflowElementItemEdit!HorizontalOffsetBehaviour = .Fields("HorizontalOffsetBehaviour")
        recWorkflowElementItemEdit!HeightBehaviour = .Fields("HeightBehaviour")
        recWorkflowElementItemEdit!WidthBehaviour = .Fields("WidthBehaviour")
        
        recWorkflowElementItemEdit.Update

        ' Copy the workflow element item definitions.
        sSQL2 = "SELECT tmpWorkflowElementItemValues.*" & _
          " FROM tmpWorkflowElementItemValues" & _
          " WHERE tmpWorkflowElementItemValues.itemID = " & Trim(Str(.Fields("ID")))
        Set rsElementItemValues = daoDb.OpenRecordset(sSQL2, dbOpenForwardOnly, dbReadOnly)

        With rsElementItemValues
          ' For each element item value definition ...
          Do While (Not .EOF)
            ' Add a new record in the database for the copied element item definition.
            recWorkflowElementItemValuesEdit.AddNew

            recWorkflowElementItemValuesEdit!itemID = lngNewID
            recWorkflowElementItemValuesEdit!value = .Fields("Value")
            recWorkflowElementItemValuesEdit!Sequence = .Fields("Sequence")

            recWorkflowElementItemValuesEdit.Update

            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsElementItemValues = Nothing

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsElementItems = Nothing

    ' Copy the workflow element column definitions.
    sSQL = "SELECT tmpWorkflowElementColumns.*" & _
      " FROM tmpWorkflowElementColumns" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementColumns.elementID = tmpWorkflowElements.ID" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID)
    Set rsElementColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElementColumns
      ' For each element column definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied element column definition.
        recWorkflowElementColumnEdit.AddNew

        lngNewID = UniqueColumnValue("tmpWorkflowElementColumns", "ID")
        recWorkflowElementColumnEdit!ID = lngNewID

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("ElementID") Then
            recWorkflowElementColumnEdit!elementid = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        lngTempNewID = IIf(IsNull(.Fields("columnID")), 0, .Fields("columnID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "COLUMN" And _
              pavCloneRegister(2, iIndex) = .Fields("columnID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementColumnEdit!ColumnID = lngTempNewID
        
        recWorkflowElementColumnEdit!ValueType = .Fields("ValueType")
        recWorkflowElementColumnEdit!value = .Fields("Value")
        recWorkflowElementColumnEdit!WFFormIdentifier = .Fields("WFFormIdentifier")
        recWorkflowElementColumnEdit!WFValueIdentifier = .Fields("WFValueIdentifier")

        lngTempNewID = IIf(IsNull(.Fields("DBColumnID")), 0, .Fields("DBColumnID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "COLUMN" And _
              pavCloneRegister(2, iIndex) = .Fields("DBColumnID") Then
              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementColumnEdit!DBColumnID = lngTempNewID

        recWorkflowElementColumnEdit!DBRecord = .Fields("DBRecord")

        lngTempNewID = IIf(IsNull(.Fields("CalcID")), 0, .Fields("CalcID"))
        
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then

              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementColumnEdit!CalcID = lngTempNewID

        recWorkflowElementColumnEdit.Update

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsElementColumns = Nothing

    ' Copy the workflow element validation definitions.
    sSQL = "SELECT tmpWorkflowElementValidations.*" & _
      " FROM tmpWorkflowElementValidations" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementValidations.elementID = tmpWorkflowElements.ID" & _
      " WHERE tmpWorkflowElements.workflowID = " & CStr(plngWorkflowID)
    Set rsElementValidations = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElementValidations
      ' For each element Validation definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied element Validation definition.
        recWorkflowElementValidationEdit.AddNew

        lngNewID = UniqueColumnValue("tmpWorkflowElementValidations", "ID")
        recWorkflowElementValidationEdit!ID = lngNewID

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("ElementID") Then
            recWorkflowElementValidationEdit!elementid = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        lngTempNewID = IIf(IsNull(.Fields("exprID")), 0, .Fields("exprID"))
        If lngTempNewID > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = lngTempNewID Then

              lngTempNewID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        recWorkflowElementValidationEdit!ExprID = lngTempNewID

        recWorkflowElementValidationEdit!Type = .Fields("Type")
        recWorkflowElementValidationEdit!Message = .Fields("message")

        recWorkflowElementValidationEdit.Update

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsElementValidations = Nothing

    ' Copy the workflow link definitions.
    sSQL = "SELECT *" & _
      " FROM tmpWorkflowLinks" & _
      " WHERE tmpWorkflowLinks.workflowID = " & CStr(plngWorkflowID)
    Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsLinks
      ' For each link definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied link definition.
        recWorkflowLinkEdit.AddNew

        lngNewID = UniqueColumnValue("tmpWorkflowLinks", "ID")
        recWorkflowLinkEdit!ID = lngNewID

        recWorkflowLinkEdit!WorkflowID = lngWorkflowID
        recWorkflowLinkEdit!StartOutboundFlowCode = .Fields("StartOutboundFlowCode")

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("StartElementID") Then
            recWorkflowLinkEdit!StartElementID = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("EndElementID") Then
            recWorkflowLinkEdit!EndElementID = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        recWorkflowLinkEdit.Update

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsLinks = Nothing
  End If
    
  rsWorkflow.Close

TidyUpAndExit:
  Set rsWorkflow = Nothing
  Set rsElements = Nothing
  Set rsLinks = Nothing
  Set rsElementItems = Nothing
  Set rsElementItemValues = Nothing
  Set rsElementColumns = Nothing
  Set rsElementValidations = Nothing
  
  CloneWorkflow = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


'Public Function TwipsToPixels(pValue As Single) As Long
'  TwipsToPixels = Round((pValue / 15), 0)
'End Function
'
'Public Function PixelsToTwips(pValue As Double) As Long
'  PixelsToTwips = (pValue * 15)
'End Function

Public Function FormatDescription(psFormatCode As Variant) As String
  On Error GoTo ErrorTrap
  
  Dim sDescription As String
  
  sDescription = "<unknown>"
  
  Select Case UCase(Trim(psFormatCode))
    Case "T"
      ' Tab
      sDescription = "Tab"
    
    Case "N"
      ' New line
      sDescription = "New line"
      
    Case "L"
      ' Draw a line
      sDescription = "Line"
      
  End Select

TidyUpAndExit:
  FormatDescription = sDescription
  Exit Function
  
ErrorTrap:
  sDescription = "<unknown>"
  Resume TidyUpAndExit
  
End Function


Public Function WebFormItemHasProperty(piItemType As WorkflowWebFormItemTypes, _
  piProperty As Integer) As Boolean
  ' Return TRUE if the given WebForm item type has the given property.
  
  Dim fHasProperty As Boolean
  
  fHasProperty = False
  
  Select Case piProperty
    Case WFITEMPROP_UNKNOWN ' 0 - used for the general 'properties' button
      fHasProperty = True

    Case WFITEMPROP_ALIGNMENT ' 1
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC)

    Case WFITEMPROP_BACKCOLOR ' 2
      fHasProperty = (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_WFVALUE) _
        Or (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_LINE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE)
  
    Case WFITEMPROP_BORDERSTYLE ' 3
      fHasProperty = (piItemType = giWFFORMITEM_IMAGE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_WFVALUE)
               
    Case WFITEMPROP_CAPTION ' 4
      fHasProperty = (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE)
    
    Case WFITEMPROP_FONT ' 5
      fHasProperty = (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_WFVALUE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE)
    
    Case WFITEMPROP_FORECOLOR ' 6
      fHasProperty = (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_WFVALUE) _
        Or (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE)
    
    Case WFITEMPROP_HEIGHT ' 7
      fHasProperty = (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_IMAGE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_LINE) _
        Or (piItemType = giWFFORMITEM_WFVALUE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE) _
        Or (piItemType = giWFFORMITEM_PAGETAB)
      
    Case WFITEMPROP_LEFT ' 8
      fHasProperty = (piItemType <> giWFFORMITEM_FORM)
    
    Case WFITEMPROP_PICTURE ' 9
      fHasProperty = (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_IMAGE)

    Case WFITEMPROP_TOP ' 10
      fHasProperty = (piItemType <> giWFFORMITEM_FORM)
    
    Case WFITEMPROP_WIDTH ' 11
      fHasProperty = (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_IMAGE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_LINE) _
        Or (piItemType = giWFFORMITEM_WFVALUE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE) _
        Or (piItemType = giWFFORMITEM_PAGETAB)
    
    Case WFITEMPROP_WFIDENTIFIER ' 12  ' Identifier of the Input/RecSel control in this WebForm
      fHasProperty = (piItemType = giWFFORMITEM_BUTTON) _
        Or (piItemType = giWFFORMITEM_FORM) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)

    Case WFITEMPROP_PICTURELOCATION ' 13
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
    
    Case WFITEMPROP_DEFAULTVALUE_CHAR ' 14
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_CHAR)

    Case WFITEMPROP_DBRECORD ' 15 ' RecordSelection type of a DBValue control
      ' NB. WFITEMPROP_RECSELTYPE (33) is the RecordSelection type of a RecSel control
      fHasProperty = (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_DBFILE)

    Case WFITEMPROP_SIZE ' 16
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)

    Case WFITEMPROP_DECIMALS ' 17
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC)

    Case WFITEMPROP_DEFAULTVALUE_DATE ' 18
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_DATE)

    Case WFITEMPROP_DEFAULTVALUE_LOGIC ' 19
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC)

    Case WFITEMPROP_DEFAULTVALUE_NUMERIC ' 20
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC)

    Case WFITEMPROP_BACKSTYLE ' 21
      fHasProperty = (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_FRAME) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_LABEL) _
        Or (piItemType = giWFFORMITEM_WFVALUE) _
        Or (piItemType = giWFFORMITEM_DBFILE) _
        Or (piItemType = giWFFORMITEM_WFFILE)
      
    Case WFITEMPROP_BACKCOLOREVEN ' 22
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_BACKCOLORODD ' 23
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_COLUMNHEADERS ' 24
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)
    
    Case WFITEMPROP_FORECOLOREVEN ' 25
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_FORECOLORODD ' 26
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_HEADERBACKCOLOR ' 27
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_HEADFONT ' 28
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_HEADLINES ' 29
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)
      
    Case WFITEMPROP_TABLEID ' 30
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_ELEMENTIDENTIFIER ' 31 ' Identifier of a preceding record identifying element (StoredData or WebForm with RecSel)
      fHasProperty = (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_DBFILE)

    Case WFITEMPROP_RECORDSELECTOR ' 32  ' Identifier of a RecSel control in the WebForm identified by WFITEMPROP_ELEMENTIDENTIFIER
      fHasProperty = (piItemType = giWFFORMITEM_DBVALUE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_DBFILE)
  
    Case WFITEMPROP_RECSELTYPE ' 33 ' RecordSelection type of a RecSel control
      ' NB. WFITEMPROP_DBRECORD (15) is the RecordSelection type of a DBValue control
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_BACKCOLORHIGHLIGHT ' 34
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_FORECOLORHIGHLIGHT ' 35
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)
    
    Case WFITEMPROP_TIMEOUT ' 36
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
    
    Case WFITEMPROP_CONTROLVALUELIST ' 37
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) Or _
        (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP)
    
    Case WFITEMPROP_DEFAULTVALUE_LIST ' 38
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP)

    Case WFITEMPROP_LOOKUPTABLEID ' 39
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)

    Case WFITEMPROP_LOOKUPCOLUMNID ' 40
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)

    Case WFITEMPROP_DEFAULTVALUE_LOOKUP ' 41
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)
    
    Case WFITEMPROP_RECORDTABLEID ' 42 ' Table identified by WFITEMPROP_ELEMENTIDENTIFIER/WFITEMPROP_RECORDSELECTOR (can be ascendant table of the one in the element/recsel)
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)
    
    Case WFITEMPROP_DESCRIPTION ' 43
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
      
    Case WFITEMPROP_ORIENTATION  ' 44
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_LINE)
  
    Case WFITEMPROP_RECORDORDER   ' 45
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)

    Case WFITEMPROP_RECORDFILTER   ' 46
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)
  
    Case WFITEMPROP_VALIDATION   ' 47
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
    
    Case WFITEMPROP_MANDATORY   ' 48
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_GRID) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)
  
    Case WFITEMPROP_DEFAULTVALUE_EXPRID   ' 49
      ' NB. Mutually exclusive to WFITEMPROP_CALCULATION as saved in the same column.
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC)
          
    Case WFITEMPROP_DEFAULTVALUE_WORKPATTERN ' 50
      ' Future dev
  
    Case WFITEMPROP_DESCRIPTION_WORKFLOWNAME ' 51
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
      
    Case WFITEMPROP_DESCRIPTION_ELEMENTCAPTION ' 52
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
      
    Case WFITEMPROP_SUBMITTYPE ' 53
      fHasProperty = (piItemType = giWFFORMITEM_BUTTON)
    
    Case WFITEMPROP_CALCULATION ' 54
      ' NB. Mutually exclusive to WFITEMPROP_DEFAULTVALUE_EXPRID as saved in the same column.
      fHasProperty = (piItemType = giWFFORMITEM_LABEL)
  
    Case WFITEMPROP_CAPTIONTYPE ' 55
      fHasProperty = (piItemType = giWFFORMITEM_LABEL)
  
    Case WFITEMPROP_DEFAULTVALUETYPE   ' 56
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_CHAR) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DATE) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOGIC) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
        Or (piItemType = giWFFORMITEM_INPUTVALUE_NUMERIC)
  
    Case WFITEMPROP_VERTICALOFFSET, WFITEMPROP_HORIZONTALOFFSET, _
      WFITEMPROP_VERTICALOFFSETBEHAVIOUR, WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR
        fHasProperty = (piItemType = giWFFORMITEM_IMAGE) _
          Or (piItemType = giWFFORMITEM_BUTTON) '57-60
    
    Case WFITEMPROP_HEIGHTBEHAVIOUR, WFITEMPROP_WIDTHBEHAVIOUR  '61-62
      fHasProperty = (piItemType = giWFFORMITEM_IMAGE)
      
    Case WFITEMPROP_PASSWORDTYPE '63
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_CHAR)
    
    Case WFITEMPROP_COMPLETIONMESSAGETYPE, _
      WFITEMPROP_COMPLETIONMESSAGE, _
      WFITEMPROP_SAVEDFORLATERMESSAGETYPE, _
      WFITEMPROP_SAVEDFORLATERMESSAGE, _
      WFITEMPROP_FOLLOWONFORMSMESSAGETYPE, _
      WFITEMPROP_FOLLOWONFORMSMESSAGE
      ' 64, 65, 66, 67, 68, 69
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
    
    Case WFITEMPROP_FILEEXTENSIONS ' 70
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)
  
    Case WFITEMPROP_LOOKUPFILTER ' 71
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)

    Case WFITEMPROP_LOOKUPFILTERCOLUMN ' 72
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)

    Case WFITEMPROP_LOOKUPFILTEROPERATOR ' 73
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)

    Case WFITEMPROP_LOOKUPFILTERVALUE ' 74
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)

    Case WFITEMPROP_TABCAPTION ' 76
      fHasProperty = (piItemType = giWFFORMITEM_PAGETAB)
    
    Case WFITEMPROP_LOOKUPORDER ' 77
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_LOOKUP)
      
    Case WFITEMPROP_HOTSPOT ' 78
      fHasProperty = (piItemType = giWFFORMITEM_FRAME)
      
    Case WFITEMPROP_USEASTARGETIDENTIFIER ' 79
      fHasProperty = (piItemType = giWFFORMITEM_INPUTVALUE_GRID)
      
    Case WFITEMPROP_REQUIRESAUTHENTICATION ' 80
      fHasProperty = (piItemType = giWFFORMITEM_FORM)
      
  End Select
  
  WebFormItemHasProperty = fHasProperty
  
End Function

Public Function WorkflowWebFormValidationTypeDescription(piType As Integer) As String
  On Error GoTo ErrorTrap
  
  Dim sDescription As String
  
  sDescription = "<unknown>"
  
  Select Case piType
    Case WORKFLOWWFVALIDATIONTYPE_ERROR
      sDescription = "Error"
    Case WORKFLOWWFVALIDATIONTYPE_WARNING
      sDescription = "Warning"
  End Select

TidyUpAndExit:
  WorkflowWebFormValidationTypeDescription = sDescription
  Exit Function
  
ErrorTrap:
  sDescription = "<unknown>"
  Resume TidyUpAndExit
  
End Function



Public Function WorkflowInitiationTypeDescription(piInitiationType As WorkflowInitiationTypes) As String
  On Error GoTo ErrorTrap
  
  Dim sDescription As String
  
  sDescription = "<unknown>"
  
  Select Case piInitiationType
    Case WORKFLOWINITIATIONTYPE_MANUAL
      sDescription = "Manual"
    Case WORKFLOWINITIATIONTYPE_TRIGGERED
      sDescription = "Triggered"
    Case WORKFLOWINITIATIONTYPE_EXTERNAL
      sDescription = "External"
  End Select

TidyUpAndExit:
  WorkflowInitiationTypeDescription = sDescription
  Exit Function
  
ErrorTrap:
  sDescription = "<unknown>"
  Resume TidyUpAndExit
  
End Function




Public Sub DefaultWorkflowSetup()
  ' Default the module setup parameters if required.
  On Error GoTo ErrorTrap
  
  Dim lngPersModulePersonnelTableID As Long
  Dim lngLoginColumnID As Long
  Dim lngSecondLoginColumnID As Long
  
  With recModuleSetup
    .Index = "idxModuleParameter"
      
    ' ------------------------------------------
    ' Default the Personnel Identification parameters as they moved from
    ' the Personnel module into the Workflow module proper
    ' ------------------------------------------
    ' Default the Workflow module Personnel table
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE
    If .NoMatch Then
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
      If Not .NoMatch Then
        lngPersModulePersonnelTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
      
        .AddNew
        !moduleKey = gsMODULEKEY_WORKFLOW
        !parameterkey = gsPARAMETERKEY_PERSONNELTABLE
        !ParameterType = gsPARAMETERTYPE_TABLEID
        !parametervalue = lngPersModulePersonnelTableID
        .Update
      End If
    End If

    ' Default the Login Name field.
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_LOGINNAME
    If .NoMatch Then
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME
      If Not .NoMatch Then
        lngLoginColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
      
        .AddNew
        !moduleKey = gsMODULEKEY_WORKFLOW
        !parameterkey = gsPARAMETERKEY_LOGINNAME
        !ParameterType = gsPARAMETERTYPE_COLUMNID
        !parametervalue = lngLoginColumnID
        .Update
      End If
    End If

    ' Default the SecondLogin Name field.
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_SECONDLOGINNAME
    If .NoMatch Then
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECONDLOGINNAME
      If Not .NoMatch Then
        lngSecondLoginColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
      
        .AddNew
        !moduleKey = gsMODULEKEY_WORKFLOW
        !parameterkey = gsPARAMETERKEY_SECONDLOGINNAME
        !ParameterType = gsPARAMETERTYPE_COLUMNID
        !parametervalue = lngSecondLoginColumnID
        .Update
      End If
    End If
  End With
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub




Public Function ConfigureWorkflowSpecifics() As Boolean
  ' Configure module specific objects (eg. stored procedures)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sErrorMessage As String

  mvar_fGeneralOK = True
  mvar_sGeneralMsg = ""
  
  fOK = (glngSQLVersion >= 8)
  If Not fOK Then
    mvar_fGeneralOK = False
    sErrorMessage = "Workflow is only available for SQL 2000 and later." & vbNewLine & mvar_sGeneralMsg
    
    fOK = (OutputMessage(sErrorMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
  End If
  
  If fOK Then
    ' Read the Workflow parameters.
    fOK = ReadWorkflowParameters
    If Not fOK Then
      mvar_fGeneralOK = False
      sErrorMessage = "Workflow specifics not correctly configured." & vbNewLine & _
        "Some functionality will be disabled if you do not change your configuration." & vbNewLine & mvar_sGeneralMsg
      
      fOK = (OutputMessage(sErrorMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
    End If
  End If
  
  'Make sure that we drop the workflow SPs
  DropWorkflowObjects
  
  ' Create the functions.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateUDF_AscendantRecordID
    If Not fOK Then
      DropFunction msAscendantRecordID_FUNCTIONNAME
    End If
  End If
  
  If fOK And mvar_fGeneralOK Then
    fOK = CreateUDF_GetLoginName
    If Not fOK Then
      DropFunction msGetLoginName_FUNCTIONNAME
    End If
  End If

  If fOK And mvar_fGeneralOK Then
    fOK = CreateUDF_ValidTableRecordID
    If Not fOK Then
      DropFunction msValidTableRecordID_FUNCTIONNAME
    End If
  End If
  
  ' Create the GetEmailAddresses stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_GetEmailAddresses
    If Not fOK Then
      DropProcedure msGetEmailAddresses_PROCEDURENAME
    End If
  End If
  
  ' Create the GetDelegatedRecords function.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateUDF_GetDelegatedRecords
    If Not fOK Then
      DropFunction msGetDelegatedRecords_FUNCTIONNAME
    End If
  End If
  
  ' Create the GetLoginName stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_GetLoginName
    If Not fOK Then
      DropProcedure msGetLoginName_PROCEDURENAME
    End If
  End If
  
  ' Create the Pending Step Check stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_CheckPendingSteps
    If Not fOK Then
      DropProcedure msCheckPendingSteps_PROCEDURENAME
    End If
  End If
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_IntranetCheckPendingSteps
    If Not fOK Then
      DropProcedure msIntCheckPendingSteps_PROCEDURENAME
    End If
  End If
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_WorkspaceCheckPendingSteps
    If Not fOK Then
      DropProcedure msWorkspaceCheckPendingSteps_PROCEDURENAME
    End If
  End If
  
  ' Create the OutOfOffice Check/Set stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_OutOfOfficeCheck
    If Not fOK Then
      DropProcedure msOutOfOfficeCheck_PROCEDURENAME
    End If
  End If
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_OutOfOfficeSet
    If Not fOK Then
      DropProcedure msOutOfOfficeSet_PROCEDURENAME
    End If
  End If
  
  If Not Application.PersonnelModule Then
    ' Might still need this, and be able to create it if done through Workflow setup.
    fOK = modPersonnelSpecifics.CreateSP_GetCurrentUserRecordID
  End If
  
  
TidyUpAndExit:
  ConfigureWorkflowSpecifics = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error configuring Workflow specifics"
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function CreateSP_IntranetCheckPendingSteps() As Boolean
  ' Create the Intranet Check Pending Steps stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msIntCheckPendingSteps_PROCEDURENAME & "]" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    DECLARE" & vbNewLine & _
    "        @sURL varchar(MAX)," & vbNewLine & _
    "        @sDescription varchar(MAX)," & vbNewLine & _
    "        @sCalcDescription varchar(MAX)," & vbNewLine & _
    "        @iInstanceID integer," & vbNewLine & _
    "        @iInstanceStepID integer," & vbNewLine & _
    "        @iElementID integer," & vbNewLine & _
    "        @hResult integer," & vbNewLine & _
    "        @objectToken integer," & vbNewLine & _
    "        @sQueryString varchar(MAX)," & vbNewLine & _
    "        @sParam1  varchar(MAX)," & vbNewLine & _
    "        @sServerName sysname," & vbNewLine & _
    "        @sDBName  sysname," & vbNewLine & _
    "        @sWorkflowName varchar(MAX);" & vbNewLine

  sProcSQL = sProcSQL & vbNewLine & _
    "    DECLARE @pass1 TABLE(" & vbNewLine & _
    "        [instanceID]  integer," & vbNewLine & _
    "        [elementID]   integer," & vbNewLine & _
    "        [stepID]      integer," & vbNewLine & _
    "        [name]        varchar(MAX)," & vbNewLine & _
    "        [description] varchar(MAX)," & vbNewLine & _
    "        [url]         nvarchar(MAX));" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "    DECLARE @steps TABLE (" & vbNewLine & _
    "        [name]           varchar(MAX)," & vbNewLine & _
    "        [description]    varchar(MAX)," & vbNewLine & _
    "        [URL]            varchar(MAX)," & vbNewLine & _
    "        [instanceID]     integer," & vbNewLine & _
    "        [elementID]      integer," & vbNewLine & _
    "        [instanceStepID] integer);" & vbNewLine & vbNewLine
      
  sProcSQL = sProcSQL & _
    "    SELECT @sURL = parameterValue" & vbNewLine & _
    "    FROM ASRSysModuleSetup" & vbNewLine & _
    "    WHERE moduleKey = 'MODULE_WORKFLOW' AND parameterKey = 'Param_URL';" & vbNewLine & vbNewLine & _
    "        IF upper(right(@sURL, 5)) <> '.ASPX'" & vbNewLine & _
    "            AND right(@sURL, 1) <> '/'" & vbNewLine & _
    "            AND len(@sURL) > 0" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            SET @sURL = @sURL + '/'" & vbNewLine & _
    "        END"
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "    SELECT @sParam1 = parameterValue" & vbNewLine & _
    "    FROM ASRSysModuleSetup" & vbNewLine & _
    "    WHERE moduleKey = 'MODULE_WORKFLOW' AND parameterKey = 'Param_Web1';" & vbNewLine & vbNewLine & _
    "    SET @sServerName = CONVERT(sysname,SERVERPROPERTY('servername'));" & vbNewLine & _
    "    SET @sDBName = DB_NAME();"
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "    IF (len(@sURL) > 0)" & vbNewLine & _
    "    BEGIN"
    
  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
        "        DECLARE @sEmailAddress_" & CStr(iCount) & " varchar(MAX)" & vbNewLine & _
        "        SELECT @sEmailAddress_" & CStr(iCount) & " = replace(upper(ltrim(rtrim(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & "))), ' ', '')" & vbNewLine & _
        "        FROM " & mvar_sLoginTable & vbNewLine & _
        "        WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = SUSER_SNAME()" & _
        IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = SUSER_SNAME()", "") & ")" & vbNewLine & _
        "            AND LEN(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & ") > 0"
    Next iCount
  End If
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "        INSERT @pass1" & vbNewLine & _
    "            SELECT ASRSysWorkflowInstanceSteps.instanceID," & vbNewLine & _
    "                ASRSysWorkflowInstanceSteps.elementID," & vbNewLine & _
    "                ASRSysWorkflowInstanceSteps.ID," & vbNewLine & _
    "                ASRSysWorkflows.name + ' - ' + ASRSysWorkflowElements.caption AS [description], " & vbNewLine & _
    "                ASRSysWorkflows.name, " & vbNewLine & _
    "                dbo.[udfASRNetGetWorkflowQueryString]( ASRSysWorkflowInstanceSteps.instanceID,  ASRSysWorkflowInstanceSteps.elementID, @sParam1, @sServerName, @sDBName)" & vbNewLine & _
    "            FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
    "            INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
    "            INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
    "            WHERE (ASRSysWorkflowInstanceSteps.Status = 2" & vbNewLine & _
    "                    OR ASRSysWorkflowInstanceSteps.Status = 7)" & vbNewLine & _
    "                AND (ASRSysWorkflowInstanceSteps.userName = SUSER_SNAME()"
    
  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & vbNewLine & _
        "                    OR (';' + replace(upper(ASRSysWorkflowInstanceSteps.userEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_" & CStr(iCount) & " + ';%'" & vbNewLine & _
        "                        AND len(@sEmailAddress_" & CStr(iCount) & ") > 0)" & vbNewLine & _
        "                    OR ((len(@sEmailAddress_" & CStr(iCount) & ") > 0)" & vbNewLine & _
        "                        AND ((SELECT COUNT(*)" & vbNewLine & _
        "                            FROM ASRSysWorkflowStepDelegation" & vbNewLine & _
        "                            WHERE stepID = ASRSysWorkflowInstanceSteps.ID" & vbNewLine & _
        "                                AND ';' + replace(upper(ASRSysWorkflowStepDelegation.delegateEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_" & CStr(iCount) & " + ';%') > 0))"
    Next iCount
  End If
    
  sProcSQL = sProcSQL & _
    ")"
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "        DECLARE steps_cursor CURSOR LOCAL FAST_FORWARD FOR SELECT * FROM @pass1;" & vbNewLine & _
    "        OPEN steps_cursor;" & vbNewLine & _
    "        FETCH NEXT FROM steps_cursor INTO @iInstanceID, @iElementID, @iInstanceStepID, @sDescription, @sWorkflowName, @sQueryString;" & vbNewLine & _
    "        WHILE (@@fetch_status = 0)" & vbNewLine & _
    "        BEGIN" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "            IF LEN(@sQueryString) > 0" & vbNewLine & _
    "            BEGIN" & vbNewLine & _
    "                EXEC [dbo].[spASRWorkflowStepDescription]" & vbNewLine & _
    "                    @iInstanceStepID," & vbNewLine & _
    "                    @sCalcDescription OUTPUT;" & vbNewLine & _
    "                IF LEN(@sCalcDescription) > 0 " & vbNewLine & _
    "                    SET @sDescription = @sCalcDescription;" & vbNewLine & vbNewLine & _
    "                INSERT INTO @steps ([description], [url], [instanceID], [elementID], [instanceStepID], [name])" & vbNewLine & _
    "                    VALUES (@sDescription, @sURL + '/?' + @sQueryString, @iInstanceID, @iElementID, @iInstanceStepID, @sWorkflowName);" & vbNewLine & _
    "            END" & vbNewLine & _
    "            FETCH NEXT FROM steps_cursor INTO @iInstanceID, @iElementID, @iInstanceStepID, @sDescription, @sWorkflowName, @sQueryString;" & vbNewLine & _
    "        END" & vbNewLine & _
    "        CLOSE steps_cursor;" & vbNewLine & _
    "        DEALLOCATE steps_cursor;" & vbNewLine & vbNewLine & _
    "    END" & vbNewLine & vbNewLine & _
    "    SELECT *" & vbNewLine & _
    "    FROM @steps" & vbNewLine & _
    "    ORDER BY [description];"
  
  sProcSQL = sProcSQL & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_IntranetCheckPendingSteps = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Intranet Check Pending Steps stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function

Private Function CreateUDF_GetDelegatedRecords() As Boolean
  ' Create the Get Delegated Email Addresses stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer

  fCreatedOK = True

  ' Construct the function creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module function.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE FUNCTION [dbo].[" & msGetDelegatedRecords_FUNCTIONNAME & "]" & vbNewLine & _
    "(" & vbNewLine & _
    "    @psOriginalRecipient varchar(MAX)" & vbNewLine & _
    ")" & vbNewLine & _
    "RETURNS @results TABLE (" & vbNewLine & _
    "    id integer," & vbNewLine & _
    "    emailAddress varchar(MAX)," & vbNewLine & _
    "    delegated bit," & vbNewLine & _
    "    delegatedTo varchar(MAX)" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & _
    "    DECLARE" & vbNewLine & _
    "        @sSingleRecipient varchar(MAX)," & vbNewLine & _
    "        @iCount integer," & vbNewLine & _
    "        @iIndex integer" & vbNewLine & vbNewLine & _
    "    DECLARE @recipients TABLE (" & vbNewLine & _
    "        recordID  integer," & vbNewLine & _
    "        emailAddress  varchar(MAX) COLLATE database_default," & vbNewLine & _
    "        delegated   bit," & vbNewLine & _
    "        delegatedTo   varchar(MAX) COLLATE database_default" & vbNewLine & _
    "    )" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "    SET @psOriginalRecipient = ltrim(rtrim(@psOriginalRecipient))" & vbNewLine & _
    "    WHILE len(@psOriginalRecipient) > 0" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        SET @iIndex = CHARINDEX(';', @psOriginalRecipient)" & vbNewLine & vbNewLine & _
    "        IF @iIndex = 0" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            SET @sSingleRecipient = @psOriginalRecipient" & vbNewLine & _
    "            SET @psOriginalRecipient = ''" & vbNewLine & _
    "        END" & vbNewLine & _
    "        ELSE" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            SET @sSingleRecipient = ltrim(rtrim(LEFT(@psOriginalRecipient, @iIndex - 1)))" & vbNewLine & _
    "            SET @psOriginalRecipient = ltrim(rtrim(SUBSTRING(@psOriginalRecipient, @iIndex + 1, len(@psOriginalRecipient) - (@iIndex - 1))))" & vbNewLine & _
    "        END" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "        SELECT @iCount = COUNT(*)" & vbNewLine & _
    "        FROM @recipients" & vbNewLine & _
    "        WHERE emailAddress = @sSingleRecipient" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "        IF (@iCount = 0) AND (len(@sSingleRecipient) > 0)" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            INSERT INTO @recipients (" & vbNewLine & _
    "                recordID," & vbNewLine & _
    "                emailAddress," & vbNewLine & _
    "                delegated," & vbNewLine & _
    "                delegatedTo)" & vbNewLine & _
    "            VALUES (" & vbNewLine & _
    "                0," & vbNewLine & _
    "                @sSingleRecipient," & vbNewLine & _
    "                0," & vbNewLine & _
    "                '')" & vbNewLine

  If (mvar_lngActivateDelegationColumn > 0) _
    And (mvar_lngDelegationEmail > 0) _
    And UBound(malngEmailColumns) > 0 Then
  
    sProcSQL = sProcSQL & vbNewLine & _
      "            INSERT INTO @recipients (" & vbNewLine & _
      "                recordID," & vbNewLine & _
      "                emailAddress," & vbNewLine & _
      "                delegated," & vbNewLine & _
      "                delegatedTo)" & vbNewLine & _
      "            SELECT " & mvar_sLoginTable & ".ID," & vbNewLine & _
      "                @sSingleRecipient," & vbNewLine & _
      "                " & mvar_sLoginTable & "." & mvar_sActivateDelegationColumn & "," & vbNewLine & _
      "                ''" & vbNewLine & _
      "            FROM " & mvar_sLoginTable & vbNewLine
      
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & _
        "            " & IIf(iCount = 1, "WHERE", "    OR") & " ltrim(rtrim(" & GetColumnName(malngEmailColumns(iCount), True) & ")) = @sSingleRecipient" & vbNewLine
    Next iCount
  End If
    
  sProcSQL = sProcSQL & _
    "        END" & vbNewLine & _
    "    END" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "    INSERT @results" & vbNewLine & _
    "        SELECT DISTINCT recordID," & vbNewLine & _
    "            emailAddress," & vbNewLine & _
    "            delegated," & vbNewLine & _
    "            delegatedTo" & vbNewLine & _
    "        FROM @recipients" & vbNewLine & vbNewLine & _
    "    RETURN" & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateUDF_GetDelegatedRecords = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Get Delegated Email Addresses function (Workflow)"
  Resume TidyUpAndExit

End Function


Private Function CreateSP_OutOfOfficeCheck() As Boolean
  ' Create the Get Out Of Office Check stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String

  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msOutOfOfficeCheck_PROCEDURENAME & "]" & vbNewLine & _
    "(" & vbNewLine & _
    "    @pfOutOfOffice bit output," & vbNewLine & _
    "    @piRecordCount integer output" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    DECLARE @iCount integer" & vbNewLine & _
    "    DECLARE @bIsOvernight bit" & vbNewLine & vbNewLine & _
    "    SET @pfOutOfOffice = 0" & vbNewLine & _
    "    SET @piRecordCount = 0" & vbNewLine & vbNewLine & _
    "    SELECT @bIsOvernight = [SettingValue] FROM ASRSYSSystemSettings" & vbNewLine & _
    "      WHERE [Section] = 'database' AND [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine
   
  If mvar_lngActivateDelegationColumn > 0 Then
    sProcSQL = sProcSQL & vbNewLine & _
      "    IF @bIsOvernight <> 1" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SELECT @piRecordCount = COUNT(*)" & vbNewLine & _
      "        FROM " & mvar_sLoginTable & vbNewLine & _
      "        WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = SUSER_SNAME()" & _
      IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = SUSER_SNAME()", "") & ")" & vbNewLine & vbNewLine & _
      "        SELECT @iCount = COUNT(*)" & vbNewLine & _
      "        FROM " & mvar_sLoginTable & vbNewLine & _
      "        WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = SUSER_SNAME()" & _
      IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = SUSER_SNAME()", "") & ")" & vbNewLine & vbNewLine & _
      "            AND " & mvar_sLoginTable & "." & mvar_sActivateDelegationColumn & " = 1" & vbNewLine & _
      "    END" & vbNewLine & vbNewLine & _
      "    IF @iCount > 0 SET @pfOutOfOffice = 1" & vbNewLine
  End If
  
  sProcSQL = sProcSQL & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_OutOfOfficeCheck = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Out Of Office Check stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function








Private Function CreateSP_OutOfOfficeSet() As Boolean
  ' Create the Get Out Of Office Set stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String

  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msOutOfOfficeSet_PROCEDURENAME & "]" & vbNewLine & _
    "(" & vbNewLine & _
    "    @pfOutOfOffice bit" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    DECLARE" & vbNewLine & _
    "        @sMailUserName varchar(MAX)" & vbNewLine

  If mvar_lngActivateDelegationColumn > 0 Then
    sProcSQL = sProcSQL & vbNewLine & _
      "    SET @sMailUserName = rtrim(system_user)" & vbNewLine & vbNewLine & _
      "    UPDATE " & mvar_sLoginTable & vbNewLine & _
      "    SET " & mvar_sLoginTable & "." & mvar_sActivateDelegationColumn & " = @pfOutOfOffice" & vbNewLine & _
      "    WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = SUSER_SNAME()" & _
      IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = SUSER_SNAME()", "") & ")" & vbNewLine & vbNewLine & _
      "    EXEC [dbo].[spASREmailImmediate] @sMailUserName" & vbNewLine
  End If

  sProcSQL = sProcSQL & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_OutOfOfficeSet = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Out Of Office Set stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function









Private Function CreateSP_GetEmailAddresses() As Boolean
  ' Create the Get Email Addresses stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer

  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msGetEmailAddresses_PROCEDURENAME & "]" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    DECLARE @iDummy integer" & vbNewLine & vbNewLine
    
  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & _
        IIf(iCount > 1, "    UNION" & vbNewLine, "") & _
        "    SELECT DISTINCT " & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & " AS [address]" & vbNewLine & _
        "    FROM " & mvar_sLoginTable & vbNewLine & _
        "    WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = SUSER_SNAME()" & _
        IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = SUSER_SNAME()", "") & ")" & vbNewLine & _
        "        AND len(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & ") > 0" & vbNewLine
    Next iCount
  End If

  sProcSQL = sProcSQL & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_GetEmailAddresses = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Get Email Addresses stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function





Private Function CreateUDF_GetLoginName() As Boolean
  ' Create the GetLoginName function.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer

  fCreatedOK = True

  ' Construct the function creation string.
  sProcSQL = _
    "-----------------------------------------------------" & vbNewLine & _
    "-- Workflow module function." & vbNewLine & _
    "-- Automatically generated by the System manager." & vbNewLine & _
    "----------------------------------------------------- " & vbNewLine & _
    "CREATE FUNCTION [dbo].[" & msGetLoginName_FUNCTIONNAME & "]" & vbNewLine & _
    "(" & vbNewLine & _
    String(1, vbTab) & "@piRecordID  integer" & vbNewLine & _
    ")" & vbNewLine & _
    "RETURNS varchar(MAX)" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    String(1, vbTab) & "DECLARE" & vbNewLine & _
    String(2, vbTab) & "@sLoginName   varchar(MAX)" & vbNewLine & vbNewLine
    
  If Len(mvar_sSecondLoginColumn) > 0 Then
    sProcSQL = sProcSQL & _
      String(1, vbTab) & "SELECT @sLoginName = CASE" & vbNewLine & _
      String(3, vbTab) & "WHEN LEN(ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '')) = 0 THEN ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '')" & vbNewLine & _
      String(3, vbTab) & "ELSE ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '')" & vbNewLine & _
      String(2, vbTab) & "END" & vbNewLine
  Else
    sProcSQL = sProcSQL & _
      String(1, vbTab) & "SELECT @sLoginName = ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '')" & vbNewLine
  End If
  
  sProcSQL = sProcSQL & _
    String(1, vbTab) & "FROM " & mvar_sLoginTable & vbNewLine & _
    String(1, vbTab) & "WHERE " & mvar_sLoginTable & ".ID = @piRecordID" & vbNewLine & vbNewLine & _
    String(1, vbTab) & "RETURN @sLoginName" & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateUDF_GetLoginName = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Get Login Name function (Workflow)"
  Resume TidyUpAndExit

End Function






Private Function CreateUDF_AscendantRecordID() As Boolean
  ' Create the Ascendant Record ID function.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  Dim sSQL As String
  Dim rsTables As DAO.Recordset

  fCreatedOK = True

  ' Construct the function creation string.
  sProcSQL = _
    "-----------------------------------------------------" & vbNewLine & _
    "-- Workflow module function." & vbNewLine & _
    "-- Automatically generated by the System manager." & vbNewLine & _
    "----------------------------------------------------- " & vbNewLine & _
    "CREATE FUNCTION [dbo].[" & msAscendantRecordID_FUNCTIONNAME & "]" & vbNewLine & _
    "(" & vbNewLine & _
    String(1, vbTab) & "@piBaseTableID  integer," & vbNewLine & _
    String(1, vbTab) & "@piBaseRecordID integer," & vbNewLine & _
    String(1, vbTab) & "@piParent1TableID integer," & vbNewLine & _
    String(1, vbTab) & "@piParent1RecordID  integer," & vbNewLine & _
    String(1, vbTab) & "@piParent2TableID integer," & vbNewLine & _
    String(1, vbTab) & "@piParent2RecordID  integer," & vbNewLine & _
    String(1, vbTab) & "@piRequiredTableID  integer" & vbNewLine & _
    ")" & vbNewLine & _
    "RETURNS integer" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & _
    String(1, vbTab) & "DECLARE" & vbNewLine & _
    String(2, vbTab) & "@iParentTableID   integer," & vbNewLine & _
    String(2, vbTab) & "@sSQL       nvarchar(MAX)," & vbNewLine & _
    String(2, vbTab) & "@sParamDefinition nvarchar(500)," & vbNewLine & _
    String(2, vbTab) & "@iParentRecordID  integer," & vbNewLine & _
    String(2, vbTab) & "@iRequiredRecordID  integer" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & _
    String(1, vbTab) & "SET @iRequiredRecordID = 0" & vbNewLine & _
    String(1, vbTab) & "SET @piParent1TableID = isnull(@piParent1TableID, 0)" & vbNewLine & _
    String(1, vbTab) & "SET @piParent1RecordID = isnull(@piParent1RecordID, 0)" & vbNewLine & _
    String(1, vbTab) & "SET @piParent2TableID = isnull(@piParent2TableID, 0)" & vbNewLine & _
    String(1, vbTab) & "SET @piParent2RecordID = isnull(@piParent2RecordID, 0)" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "IF @piBaseTableID = @piRequiredTableID" & vbNewLine & _
    String(1, vbTab) & "BEGIN" & vbNewLine & _
    String(2, vbTab) & "SET @iRequiredRecordID = @piBaseRecordID" & vbNewLine & _
    String(2, vbTab) & "RETURN @iRequiredRecordID" & vbNewLine & _
    String(1, vbTab) & "END" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "-- The base table is not the same as the required table." & vbNewLine & _
    String(1, vbTab) & "-- Check ascendant tables." & vbNewLine & _
    String(1, vbTab) & "DECLARE ascendantsCursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
    String(2, vbTab) & "SELECT ASRSysRelations.parentID" & vbNewLine & _
    String(2, vbTab) & "FROM ASRSysRelations" & vbNewLine & _
    String(2, vbTab) & "WHERE ASRSysRelations.childID = @piBaseTableID" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & _
    String(1, vbTab) & "OPEN ascendantsCursor" & vbNewLine & _
    String(1, vbTab) & "FETCH NEXT FROM ascendantsCursor INTO @iParentTableID" & vbNewLine & _
    String(1, vbTab) & "WHILE (@@fetch_status = 0) AND (@iRequiredRecordID = 0)" & vbNewLine & _
    String(1, vbTab) & "BEGIN" & vbNewLine & _
    String(2, vbTab) & "-- Get the related record in the parent table (if one exists)" & vbNewLine
    
    sSQL = _
      "SELECT tmpRelations.parentID," & _
      "   tmpRelations.childID," & _
      "   tmpTables.tableName" & _
      " FROM tmpRelations, tmpTables" & _
      " WHERE tmpRelations.childID = tmpTables.tableID" & _
      " AND tmpTables.deleted = FALSE" & _
      " ORDER BY tmpRelations.parentID, tmpRelations.childID"
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    While Not rsTables.EOF
      sProcSQL = sProcSQL & _
        String(2, vbTab) & "IF (@iParentTableID = " & rsTables!parentID & ") AND (@piBaseTableID = " & rsTables!childID & ")" & vbNewLine & _
        String(2, vbTab) & "BEGIN" & vbNewLine & _
        String(3, vbTab) & "SELECT @iParentRecordID = isnull(ID_" & CStr(rsTables!parentID) & ", 0)" & vbNewLine & _
        String(3, vbTab) & "FROM " & rsTables!TableName & vbNewLine & _
        String(3, vbTab) & "WHERE ID = @piBaseRecordID" & vbNewLine & _
        String(2, vbTab) & "END" & vbNewLine & vbNewLine
      
      rsTables.MoveNext
    Wend
    rsTables.Close
    Set rsTables = Nothing
    
  sProcSQL = sProcSQL & _
    String(2, vbTab) & "IF @iParentRecordID > 0" & vbNewLine & _
    String(2, vbTab) & "BEGIN" & vbNewLine & _
    String(3, vbTab) & "SELECT @iRequiredRecordID = [dbo].[udf_ASRWorkflowAscendantRecordID](" & vbNewLine & _
    String(4, vbTab) & "@iParentTableID," & vbNewLine & _
    String(4, vbTab) & "@iParentRecordID," & vbNewLine & _
    String(4, vbTab) & "0," & vbNewLine & _
    String(4, vbTab) & "0," & vbNewLine & _
    String(4, vbTab) & "0," & vbNewLine & _
    String(4, vbTab) & "0," & vbNewLine & _
    String(4, vbTab) & "@piRequiredTableID)" & vbNewLine & _
    String(2, vbTab) & "END" & vbNewLine & vbNewLine & _
    String(2, vbTab) & "FETCH NEXT FROM ascendantsCursor INTO @iParentTableID" & vbNewLine & _
    String(1, vbTab) & "END" & vbNewLine & _
    String(1, vbTab) & "CLOSE ascendantsCursor" & vbNewLine & _
    String(1, vbTab) & "DEALLOCATE ascendantsCursor" & vbNewLine & vbNewLine
      
  sProcSQL = sProcSQL & _
    String(1, vbTab) & "IF (@iRequiredRecordID = 0)" & vbNewLine & _
    String(2, vbTab) & "AND (@piParent1TableID > 0)" & vbNewLine & _
    String(2, vbTab) & "AND (@piParent1RecordID > 0)" & vbNewLine & _
    String(1, vbTab) & "BEGIN" & vbNewLine & _
    String(2, vbTab) & "SELECT @iRequiredRecordID = [dbo].[udf_ASRWorkflowAscendantRecordID](" & vbNewLine & _
    String(3, vbTab) & "@piParent1TableID," & vbNewLine & _
    String(3, vbTab) & "@piParent1RecordID," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "@piRequiredTableID)" & vbNewLine & _
    String(1, vbTab) & "END" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "IF (@iRequiredRecordID = 0)" & vbNewLine & _
    String(2, vbTab) & "AND (@piParent2TableID > 0)" & vbNewLine & _
    String(2, vbTab) & "AND (@piParent2RecordID > 0)" & vbNewLine & _
    String(1, vbTab) & "BEGIN" & vbNewLine & _
    String(2, vbTab) & "SELECT @iRequiredRecordID = [dbo].[udf_ASRWorkflowAscendantRecordID](" & vbNewLine & _
    String(3, vbTab) & "@piParent2TableID," & vbNewLine & _
    String(3, vbTab) & "@piParent2RecordID," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "0," & vbNewLine & _
    String(3, vbTab) & "@piRequiredTableID)" & vbNewLine & _
    String(1, vbTab) & "END" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "-- Return the result of the function" & vbNewLine & _
    String(1, vbTab) & "RETURN @iRequiredRecordID" & vbNewLine & _
    "END" & vbNewLine

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateUDF_AscendantRecordID = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Ascendant Record ID function (Workflow)"
  Resume TidyUpAndExit

End Function







Private Function CreateUDF_ValidTableRecordID() As Boolean
  ' Create the Valida Table Record ID function.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  Dim sSQL As String
  Dim rsTables As DAO.Recordset

  fCreatedOK = True

  ' Construct the function creation string.
  sProcSQL = _
    "-----------------------------------------------------" & vbNewLine & _
    "-- Workflow module function." & vbNewLine & _
    "-- Automatically generated by the System manager." & vbNewLine & _
    "----------------------------------------------------- " & vbNewLine & _
    "CREATE FUNCTION [dbo].[" & msValidTableRecordID_FUNCTIONNAME & "]" & vbNewLine & _
    "(" & vbNewLine & _
    String(1, vbTab) & "@piTableID  integer," & vbNewLine & _
    String(1, vbTab) & "@piRecordID integer" & vbNewLine & _
    ")" & vbNewLine & _
    "RETURNS bit" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "DECLARE" & vbNewLine & _
    String(2, vbTab) & "@fValid bit," & vbNewLine & _
    String(2, vbTab) & "@iRecCount integer" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "SET @iRecCount = 0" & vbNewLine & _
    String(1, vbTab) & "SET @fValid = 0" & vbNewLine & vbNewLine

  sSQL = _
    "SELECT tmpTables.tableID, tmpTables.tableName" & _
    " FROM tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    " ORDER BY tmpTables.tableID"
  Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  While Not rsTables.EOF
    sProcSQL = sProcSQL & _
      String(2, vbTab) & "IF (@piTableID = " & rsTables!TableID & ")" & vbNewLine & _
      String(2, vbTab) & "BEGIN" & vbNewLine & _
      String(3, vbTab) & "SELECT @iRecCount = COUNT(*)" & vbNewLine & _
      String(3, vbTab) & "FROM " & rsTables!TableName & vbNewLine & _
      String(3, vbTab) & "WHERE ID = @piRecordID" & vbNewLine & _
      String(2, vbTab) & "END" & vbNewLine & vbNewLine
  
    rsTables.MoveNext
  Wend
  rsTables.Close
  Set rsTables = Nothing

  sProcSQL = sProcSQL & _
    String(1, vbTab) & "IF @iRecCount > 0 SET @fValid = 1" & vbNewLine & vbNewLine & _
    String(1, vbTab) & "-- Return the result of the function" & vbNewLine & _
    String(1, vbTab) & "RETURN @fValid" & vbNewLine & _
    "END" & vbNewLine

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateUDF_ValidTableRecordID = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Valid Table Record ID function (Workflow)"
  Resume TidyUpAndExit

End Function








Private Function CreateSP_GetLoginName() As Boolean
  ' Create the GetLoginName stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String

  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msGetLoginName_PROCEDURENAME & "] (" & vbNewLine & _
    "    @piRecordID integer," & vbNewLine & _
    "    @psLoginName varchar(MAX) OUTPUT" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  If Len(mvar_sSecondLoginColumn) > 0 Then
    sProcSQL = sProcSQL & _
      "    SELECT @psLoginName = CASE" & vbNewLine & _
      "            WHEN LEN(ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '')) = 0 THEN ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '')" & vbNewLine & _
      "            ELSE ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '')" & vbNewLine & _
      "        END" & vbNewLine
  Else
    sProcSQL = sProcSQL & _
      "    SELECT @psLoginName = ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '')" & vbNewLine
  End If
  
  sProcSQL = sProcSQL & _
    "    FROM " & mvar_sLoginTable & vbNewLine & _
    "    WHERE " & mvar_sLoginTable & ".ID = @piRecordID" & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_GetLoginName = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Get Login Name stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function





Private Function CreateSP_CheckPendingSteps() As Boolean
  ' Create the Check Pending Steps stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msCheckPendingSteps_PROCEDURENAME & "]" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    DECLARE @bIsOvernight bit;" & vbNewLine & vbNewLine & _
    "    SELECT @bIsOvernight = [SettingValue] FROM ASRSYSSystemSettings" & vbNewLine & _
    "      WHERE [Section] = 'database' AND [SettingKey] = 'updatingdatedependantcolumns';" & vbNewLine & vbNewLine & _
    "    IF @bIsOvernight <> 1" & vbNewLine & _
    "    BEGIN" & vbNewLine

  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
        "        DECLARE @sEmailAddress_" & CStr(iCount) & " varchar(MAX)" & vbNewLine & _
        "        SELECT @sEmailAddress_" & CStr(iCount) & " = replace(upper(ltrim(rtrim(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & "))), ' ', '')" & vbNewLine & _
        "        FROM " & mvar_sLoginTable & vbNewLine & _
        "        WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = SUSER_SNAME()" & _
        IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = SUSER_SNAME()", "") & ")" & vbNewLine & _
        "            AND len(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & ") > 0"
    Next iCount
  End If
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "        SELECT ASRSysWorkflowInstanceSteps.instanceID," & vbNewLine & _
    "            ASRSysWorkflowInstanceSteps.elementID," & vbNewLine & _
    "            ASRSysWorkflowInstanceSteps.ID" & vbNewLine & _
    "        FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
    "        INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
    "        INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
    "        WHERE (ASRSysWorkflowInstanceSteps.Status = 2" & vbNewLine & _
    "                OR ASRSysWorkflowInstanceSteps.Status = 7)" & vbNewLine & _
    "            AND (ASRSysWorkflowInstanceSteps.userName = SUSER_SNAME()"
    
  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & vbNewLine & _
        "                OR (';' + replace(upper(ASRSysWorkflowInstanceSteps.userEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_" & CStr(iCount) & " + ';%'" & vbNewLine & _
        "                    AND len(@sEmailAddress_" & CStr(iCount) & ") > 0)" & vbNewLine & _
        "                OR ((len(@sEmailAddress_" & CStr(iCount) & ") > 0)" & vbNewLine & _
        "                    AND ((SELECT COUNT(*)" & vbNewLine & _
        "                        FROM ASRSysWorkflowStepDelegation" & vbNewLine & _
        "                        WHERE stepID = ASRSysWorkflowInstanceSteps.ID" & vbNewLine & _
        "                            AND ';' + replace(upper(ASRSysWorkflowStepDelegation.delegateEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_" & CStr(iCount) & " + ';%') > 0))"
  Next iCount
  End If
    
  sProcSQL = sProcSQL & _
    ")" & vbNewLine & "    END" & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_CheckPendingSteps = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Check Pending Steps stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function






Public Sub DropWorkflowObjects()
  ' Drop Workflow stored procedures
  DropProcedure msCheckPendingSteps_PROCEDURENAME
  DropProcedure msIntCheckPendingSteps_PROCEDURENAME
  DropProcedure msWorkspaceCheckPendingSteps_PROCEDURENAME
  DropProcedure msGetEmailAddresses_PROCEDURENAME
  DropProcedure msGetDelegatedRecords_PROCEDURENAME
  DropProcedure msOutOfOfficeCheck_PROCEDURENAME
  DropProcedure msOutOfOfficeSet_PROCEDURENAME
  DropProcedure msGetLoginName_PROCEDURENAME

  ' Drop Workflow functions
  DropFunction msAscendantRecordID_FUNCTIONNAME
  DropFunction msGetLoginName_FUNCTIONNAME
  DropFunction msValidTableRecordID_FUNCTIONNAME
  DropFunction msGetDelegatedRecords_FUNCTIONNAME
End Sub

Private Function ReadWorkflowParameters() As Boolean
  ' Read the configured Workflow parameters into member variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngLoginColumn As Long
  Dim lngSecondLoginColumn As Long
  Dim lngLoginTable As Long
  Dim lngColumnID As Long
  Dim lngTableID As Long
  Dim sUser As String
  Dim sPassword As String
  
  ReDim malngEmailColumns(0)
  
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the URL. Essential.
    mvar_sURL = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_URL, "")
    
    fOK = (Len(mvar_sURL) > 0)
    If Not fOK Then
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  URL not defined."
    End If
    
    'JPD 20070615 Fault 12313
    If fOK Then
      ' Get the Workflow web site user and password.
      ReadWebLogon sUser, sPassword
      
      fOK = (Len(sUser) > 0)
      If Not fOK Then
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  Web site login not defined."
      End If
    End If
    
    If fOK Then
      ' Get the login column(s). Essential to have at least one.
      lngLoginColumn = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_LOGINNAME, 0)
      lngSecondLoginColumn = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_SECONDLOGINNAME, 0)
   
      fOK = ((lngLoginColumn > 0) Or (lngSecondLoginColumn > 0))
      If Not fOK Then
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  Login column not defined."
      End If
      
      If ((lngLoginColumn = 0) And (lngSecondLoginColumn > 0)) Then
        lngLoginColumn = lngSecondLoginColumn
        lngSecondLoginColumn = 0
      End If
    End If
    
    If fOK Then
      lngLoginTable = GetTableIDFromColumnID(lngLoginColumn)
      mvar_sLoginColumn = GetColumnName(lngLoginColumn, True)
      mvar_sSecondLoginColumn = IIf(lngSecondLoginColumn > 0, GetColumnName(lngSecondLoginColumn, True), "")
      mvar_sLoginTable = GetTableName(lngLoginTable)

      ' Get the email columns. Not essential, so don't set fOK to false if they're not defined.
      .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_EMAILCOLUMN

      If Not .NoMatch Then
        Do While Not .EOF
          If (!moduleKey <> gsMODULEKEY_WORKFLOW) Or _
            (!parameterkey <> gsPARAMETERKEY_EMAILCOLUMN) Then

            Exit Do
          End If

          lngColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, CLng(!parametervalue))

          If lngColumnID > 0 Then
            lngTableID = GetTableIDFromColumnID(lngColumnID)
            
            If lngTableID = lngLoginTable Then
              ReDim Preserve malngEmailColumns(UBound(malngEmailColumns) + 1)
              malngEmailColumns(UBound(malngEmailColumns)) = lngColumnID
            End If
          End If
          
          .MoveNext
        Loop
      End If
    
      ' Get the Delegation Activation column. Not essential, so don't set fOK to false if it's not defined.
      .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_DELEGATIONACTIVATEDCOLUMN
      If .NoMatch Then
        mvar_lngActivateDelegationColumn = 0
      Else
        mvar_lngActivateDelegationColumn = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sActivateDelegationColumn = GetColumnName(mvar_lngActivateDelegationColumn, True)
      End If
    
      ' Get the Delegation email . Not essential, so don't set fOK to false if it's not defined.
      .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_DELEGATEEMAIL
      If .NoMatch Then
        mvar_lngDelegationEmail = 0
      Else
        mvar_lngDelegationEmail = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      End If
      
    End If
    
  End With

TidyUpAndExit:
  ReadWorkflowParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading workflow parameters"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function ReadWebLogon(strUserName As String, strPassword As String) As Boolean

  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long

  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Workflow web site user and password.
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_WEBPARAM1
    If .NoMatch Then
      strInput = ""
    Else
      strInput = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, "", !parametervalue)
    End If
  End With

  If strInput = vbNullString Then
    Exit Function
  End If

  lngStart = Len(strInput) - 12
  strEKey = Mid(strInput, lngStart + 1, 10)
  strLens = Right(strInput, 2)
  strInput = XOREncript(Left(strInput, lngStart), strEKey)

  lngStart = 1
  lngFinish = Asc(Mid(strLens, 1, 1)) - 127
  strUserName = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 2, 1)) - 127
  strPassword = Mid(strInput, lngStart, lngFinish)

End Function

'Private Function XOREncript(strInput, strKey) As String
'
'  Dim lngCount As Long
'  Dim strOutput As String
'  Dim strChar As String
'
'  For lngCount = 1 To Len(strInput)
'    strChar = Mid(strKey, lngCount Mod Len(strKey) + 1, 1)
'    strOutput = strOutput & Chr(Asc(strChar) Xor Asc(Mid(strInput, lngCount, 1)))
'  Next
'
'  XOREncript = strOutput
'
'End Function


Public Function SaveWebLogon(strUserName As String, _
  strPassword As String) As Boolean

  Dim strInput As String
  Dim strOutput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngCount As Long
  Dim iChar As Integer

  strOutput = strUserName & strPassword
  strLens = Chr(Len(strUserName) + 127) & Chr(Len(strPassword) + 127)
  
  ' AE20080229 Fault #12939, #12959
  Do While strInput = vbNullString _
    Or (CBool(InStr(strInput, Chr(0))) Or CBool(InStr(strInput, Chr(144))))
    
    strInput = vbNullString
    strEKey = vbNullString
    For lngCount = 1 To 10
      'strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
      iChar = 0
      iChar = Int(Rnd * 255) + 1
      strEKey = strEKey & Chr(iChar)
    Next

    strInput = XOREncript(strOutput, strEKey) & strEKey & strLens
  Loop
  strOutput = strInput
  
  ' Save the Web Site Login details.
  With recModuleSetup
    .Index = "idxModuleParameter"
  
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_WEBPARAM1
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_WORKFLOW
      !parameterkey = gsPARAMETERKEY_WEBPARAM1
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_ENCYPTED
    !parametervalue = strOutput
    .Update
  End With
  
End Function

Public Function EncryptQueryString(plngInstanceID As Long, _
  plngStepID As Long, _
  psUser As String, _
  psPassword As String) As String
  
  On Error GoTo ErrorTrap
  
  Dim sKey As String
  Dim sEncryptedString As String
  Dim sSourceString As String
  Dim sServerName As String
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset

  Const ENCRYPTIONKEY = "jmltn"

  ' Get the server name - gsServerName may be '.'
  ' which screws up the Workflow queryString if the web site is not
  ' on the same server as the SQL database.
  'sSQL = "SELECT @@SERVERNAME AS [serverName]"
  sSQL = "SELECT SERVERPROPERTY('servername') AS [serverName]"
  Set rsTemp = New ADODB.Recordset
  rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsTemp
    If Not (.EOF And .BOF) Then
      If IsNull(rsTemp!ServerName) Then
        sServerName = gsServerName
      Else
        sServerName = rsTemp!ServerName
      End If
    Else
      sServerName = gsServerName
    End If
    .Close
  End With
  Set rsTemp = Nothing

  sKey = ENCRYPTIONKEY
  sSourceString = CStr(plngInstanceID) & _
    vbTab & CStr(plngStepID) & _
    vbTab & psUser & _
    vbTab & psPassword & _
    vbTab & sServerName & _
    vbTab & gsDatabaseName

  sEncryptedString = EncryptString(sSourceString, sKey, True)
  sEncryptedString = CompactString(sEncryptedString)

TidyUpAndExit:
  EncryptQueryString = sEncryptedString
  Exit Function

ErrorTrap:
  sEncryptedString = ""
  Resume TidyUpAndExit
  
End Function
Public Function GetWorkflowQueryString(plngInstanceID As Long, _
  plngStepID As Long, _
  Optional psUserName As Variant, _
  Optional psPassword As Variant) As String
  ' Get the QueryString details required to externally initiate the Workflow.
  ' For externally initiated workflows:
  '      plngInstance = -1 * workflowID
  '      plngStepID = -1
  Dim sUser As String
  Dim sPassword As String
  Dim sEncryptedString As String
  
  On Error GoTo ErrorTrap
  
  If IsMissing(psUserName) And IsMissing(psPassword) Then
    ' Get the Web Site login UID and password
    ReadWebLogon sUser, sPassword
  Else
    sUser = psUserName
    sPassword = psPassword
  End If
  
  If Len(Trim(sUser)) > 0 Then
    sEncryptedString = EncryptQueryString(plngInstanceID, plngStepID, sUser, sPassword)
  End If
  
TidyUpAndExit:
  GetWorkflowQueryString = sEncryptedString
  Exit Function
  
ErrorTrap:
  sEncryptedString = ""
  Resume TidyUpAndExit
  
End Function
Public Function GetWorkflowURL() As String
  ' Get the URL details required to externally initiate the Workflow.
  On Error GoTo ErrorTrap
  
  Dim sURL As String
  
  sURL = ""
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' ------------------------------------------
    ' Read the Web Site parameters
    ' ------------------------------------------
    ' Get the Web Site URL
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_URL
    If Not .NoMatch Then
      sURL = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, "", Trim(!parametervalue))
    
      If UCase(Right(sURL, 5)) <> ".ASPX" _
        And Right(sURL, 1) <> "/" _
        And Len(sURL) > 0 Then
        
        sURL = sURL + "/"
      End If
    End If
  End With
  
TidyUpAndExit:
  GetWorkflowURL = sURL
  Exit Function
  
ErrorTrap:
  sURL = ""
  Resume TidyUpAndExit
    
End Function

Private Function CompactString(psSourceString As String) As String
  ' Compact the encrypted string.
  ' psSourceString is a string of the hexadecimal values of the Ascii codes for each character in the encrypted string.
  ' In this string each character in the encrypted string is represented as 2 hex digits.
  ' As it's a string of hex characters all characters are in the range 0-9, A-F
  ' Valid hypertext link characters are 0-9, A-Z, a-z and some others (we'll be using $ and @).
  ' Take advantage of this by implementing our own base64 encoding as follows:
  Dim sCompactedString As String
  Dim sSubString As String
  Dim sModifiedSourceString As String
  Dim iValue As Integer
  Dim iTemp As Integer
  Dim sNewString As String
  
  sCompactedString = ""
  sModifiedSourceString = psSourceString
  Do While Len(sModifiedSourceString) > 0
    ' Read the hex characters in chunks of 3 (ie. possible values 0 - 4095)
    ' This chunk of 3 Hex characters can then be translated into 2 base64 characters (ie. still have possible values 0 - 4095)
    ' Woohoo! We've reduced the length of the encrypted string by about one third!
    sNewString = ""
    sSubString = Left(sModifiedSourceString & "000", 3)
    sModifiedSourceString = Mid(sModifiedSourceString, 4)
    iValue = val("&H" & sSubString)
    
    ' Use our own base64 digit set.
    ' Base64 digit values 0-9 are represented as 0-9
    ' Base64 digit values 10-35 are represented as A-Z
    ' Base64 digit values 36-61 are represented as a-z
    ' Base64 digit value 62 is represented as $
    ' Base64 digit value 63 is represented as @
    
    iTemp = iValue Mod 64
    If iTemp = 63 Then
      sNewString = "@"
    ElseIf iTemp = 62 Then
      sNewString = "$"
    ElseIf iTemp >= 36 Then
      sNewString = Chr(iTemp + 61)
    ElseIf iTemp >= 10 Then
      sNewString = Chr(iTemp + 55)
    Else
      sNewString = Chr(iTemp + 48)
    End If
    
    iTemp = (iValue - iTemp) / 64
    
    If iTemp = 63 Then
      sNewString = "@" & sNewString
    ElseIf iTemp = 62 Then
      sNewString = "$" & sNewString
    ElseIf iTemp >= 36 Then
      sNewString = Chr(iTemp + 61) & sNewString
    ElseIf iTemp >= 10 Then
      sNewString = Chr(iTemp + 55) & sNewString
    Else
      sNewString = Chr(iTemp + 48) & sNewString
    End If
    
    sCompactedString = sCompactedString & sNewString
  Loop
 
  ' Append the number of characters to ignore, to the compacted string
  CompactString = sCompactedString & CStr((3 - (Len(psSourceString) Mod 3)) Mod 3)
  
End Function


Public Function EncryptString(psText As String, _
  Optional psKey As String, _
  Optional pbOutputInHex As Boolean) As String
  
  Dim abytArray() As Byte
  Dim abytKey() As Byte
  Dim abytOut() As Byte

  psText = psText & " "
  abytArray() = StrConv(psText, vbFromUnicode)
  abytKey() = StrConv(psKey, vbFromUnicode)
  abytOut() = EncryptByte(abytArray(), abytKey())
  EncryptString = StrConv(abytOut(), vbUnicode)

  If pbOutputInHex = True Then EncryptString = EnHex(EncryptString)

End Function
Public Function EnHex(psData As String) As String
  Dim dblCount As Double
  Dim sTemp As String
  
  Reset
  
  For dblCount = 1 To Len(psData)
    sTemp = Hex$(Asc(Mid$(psData, dblCount, 1)))
    If Len(sTemp) < 2 Then sTemp = "0" & sTemp
    Append sTemp
  Next
  
  EnHex = GData
  
  Reset
  
End Function


Private Sub Append(ByRef psStringData As String, Optional plngLength As Long)
  Dim lngDataLength As Long
  
  If plngLength > 0 Then
    lngDataLength = plngLength
  Else
    lngDataLength = Len(psStringData)
  End If
  
  If lngDataLength + mlngHiByte > mlngHiBound Then
    mlngHiBound = mlngHiBound + 1024
    ReDim Preserve mabytArray(mlngHiBound)
  End If
  
  CopyMemory ByVal VarPtr(mabytArray(mlngHiByte)), ByVal psStringData, lngDataLength
  mlngHiByte = mlngHiByte + lngDataLength
    
End Sub


Private Property Get GData() As String
  Dim sStringData As String
  
  sStringData = Space(mlngHiByte)
  CopyMemory ByVal sStringData, ByVal VarPtr(mabytArray(0)), mlngHiByte
  GData = sStringData
  
End Property


Private Sub Reset()
  mlngHiByte = 0
  mlngHiBound = 1024
  ReDim mabytArray(mlngHiBound)
  
End Sub


Public Function EncryptByte(pabytText() As Byte, pabytKey() As Byte)
  Dim abytTemp() As Byte
  Dim iTemp As Integer
  Dim iLoop As Long
  Dim iBound As Integer
  
  Call InitTbl

  ReDim abytTemp((UBound(pabytText)) + 4)
  Randomize
  abytTemp(0) = Int((Rnd * 254) + 1)
  abytTemp(1) = Int((Rnd * 254) + 1)
  abytTemp(2) = Int((Rnd * 254) + 1)
  abytTemp(3) = Int((Rnd * 254) + 1)
  abytTemp(4) = Int((Rnd * 254) + 1)

  Call CopyMemory(abytTemp(5), pabytText(0), UBound(pabytText))

  ReDim pabytText(UBound(abytTemp)) As Byte
  pabytText() = abytTemp()
  ReDim abytTemp(0)
  iBound = (UBound(pabytKey) - 1)
  iTemp = 0

  For iLoop = 0 To UBound(pabytText) - 1
    If iTemp = iBound Then iTemp = 0
    pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp)))
    pabytText(iLoop + 1) = mabytXTable(pabytText(iLoop), pabytText(iLoop + 1))
    pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp + 1)))
    iTemp = iTemp + 1
  Next iLoop

  EncryptByte = pabytText()

End Function




Private Sub InitTbl()
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  
  If mfInitTrue = True Then Exit Sub
  
  For i = 0 To 255
    For j = 0 To 255
      mabytXTable(i, j) = CByte(i Xor j)
      mabytAddTable(i, j) = CByte((i + j) Mod 255)
    Next j
  Next i
  
  mfInitTrue = True
  
End Sub



Public Function WorkflowsWithStatus(plngWorkflowID As Long, piStatus As WorkflowInstanceStatus) As Boolean
  ' Return TRUE if the given workflow has instances with the given status.
  On Error GoTo ErrorTrap
  
  Dim fWithStatus As Boolean
  Dim rsInfo As New ADODB.Recordset
  Dim sSQL As String
  
  ' Check if the workflow is in use.
  sSQL = "SELECT COUNT(*) AS recCount" & _
    " FROM ASRSysWorkflowInstances" & _
    " WHERE workflowID = " & CStr(plngWorkflowID) & _
    " AND status = " & CStr(piStatus)
  rsInfo.Open sSQL, gADOCon, adOpenKeyset, adLockReadOnly
  
  fWithStatus = (rsInfo!reccount > 0)
  
  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  WorkflowsWithStatus = fWithStatus
  Exit Function
  
ErrorTrap:
  fWithStatus = False
  Resume TidyUpAndExit

End Function

Public Function MergeControlValues(ByVal psList_Crlf As String) As String
  MergeControlValues = Replace(psList_Crlf, vbCrLf, vbTab)
End Function

Public Function SplitControlValues(ByVal psList_TAB As String) As String
  SplitControlValues = Replace(psList_TAB, vbTab, vbCrLf)
End Function

Public Function WorkflowTableTriggerCode(plngTableID As Long, _
  pfAction As WFTriggerRelatedRecord) As String
  
  Dim sCode As String
  Dim sSubCode As String
  Dim sDeleteTriggerSelectCode As String
  Dim sDeleteTriggerInsertCode As String
  Dim fNeeded As Boolean
  Dim iCount As Integer
  Dim sEffectiveDate As String
  Dim sFilter As String
  Dim iIndent As Integer
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim rsColumnUsed As New ADODB.Recordset
  Dim alngColumnsUsed() As Long
  Dim alngEmailsUsed() As Long
  Dim lngLoop As Long
  Dim lngEmailID As Long
  Dim lngColumnID As Long
  Dim strColumnName As String
  Dim iDataType As Integer
  Dim iSize As Long
  Dim iDecimals As Integer
  Dim strVariableName As String
  Dim objExpr As CExpression
  
  iCount = 0
  sCode = ""
  iIndent = 2
  
  With recWorkflowTriggeredLinks
    .Index = "idxTableID"
    .Seek "=", plngTableID

    If Not .NoMatch Then
      Do While !TableID = plngTableID
        If (Not !Deleted) And (!Type = WORKFLOWTRIGGERLINKTYPE_RECORD) Then
        
          fNeeded = False
          
          recWorkflowEdit.Index = "idxWorkflowID"
          recWorkflowEdit.Seek "=", !WorkflowID
          
          If Not recWorkflowEdit.NoMatch Then
            If recWorkflowEdit!Enabled And (Not recWorkflowEdit!Deleted) Then
            
              Select Case pfAction
                Case WFRELATEDRECORD_INSERT
                  fNeeded = .Fields("recordInsert").value
                Case WFRELATEDRECORD_UPDATE
                  ' Disabled 'update' type link - not required and causes too much hassle when the overnight job runs.
                  'fNeeded = .Fields("recordUpdate").Value
                  fNeeded = False
                Case WFRELATEDRECORD_DELETE
                  fNeeded = .Fields("recordDelete").value
              End Select
            End If
          End If
          
          If fNeeded Then
            sSubCode = ""
            sDeleteTriggerSelectCode = ""
            sDeleteTriggerInsertCode = ""
            iCount = iCount + 1
            
            If Not IsNull(!EffectiveDate) Then
              sEffectiveDate = Replace(Format(!EffectiveDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
              iIndent = iIndent + 1
            Else
              sEffectiveDate = vbNullString
            End If
            
            If !FilterID > 0 Then
              sFilter = GetSQLFilter(!FilterID, GetTableName(plngTableID))
              iIndent = iIndent + 1
            Else
              sFilter = vbNullString
            End If
            
            ' Create the basic trigger code
            If pfAction <> WFRELATEDRECORD_DELETE Then
              sSubCode = sSubCode & _
                String(iIndent, vbTab) & "DELETE FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine & _
                String(iIndent, vbTab) & "WHERE dateInitiated IS null AND recordID=@recordID AND linkID=" & .Fields("linkID").value & vbNewLine & vbNewLine
            End If
            
            sSubCode = sSubCode & _
              String(iIndent, vbTab) & "INSERT dbo.[ASRSysWorkflowQueue](LinkID,RecordID,DateDue,UserName,[Immediate],RecalculateRecordDesc,RecordDesc, parent1TableID, parent1RecordID, parent2TableID, parent2RecordID, instanceID)" & vbNewLine & _
              String(iIndent, vbTab) & "VALUES (" & .Fields("linkID").value & ",@recordID, getDate()," & _
                "CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE ltrim(rtrim(SYSTEM_USER)) END," & _
                "1, " & IIf(pfAction = WFRELATEDRECORD_DELETE, "0", "1") & ", @recordDesc, @parent1TableID, @parent1RecordID, @parent2TableID, @parent2RecordID, 0)" & vbNewLine
            
            If pfAction = WFRELATEDRECORD_DELETE Then
              alngColumnsUsed = BaseTableColumnsUsedInDeleteTriggeredWorkflow(!WorkflowID)
              For lngLoop = 1 To UBound(alngColumnsUsed)
                lngColumnID = alngColumnsUsed(lngLoop)
                strColumnName = GetColumnName(alngColumnsUsed(lngLoop), True)
                iDataType = GetColumnDataType(alngColumnsUsed(lngLoop))
                iSize = GetColumnSize(alngColumnsUsed(lngLoop), False)
                iDecimals = GetColumnSize(alngColumnsUsed(lngLoop), True)
            
                strVariableName = "@sWFTemp_" & CStr(!LinkID) & "_" & CStr(lngColumnID)
            
                sSubCode = sSubCode & vbNewLine & _
                  String(iIndent, vbTab) & "DECLARE " & strVariableName & " varchar(MAX)" & vbNewLine
            
                Select Case iDataType
                  Case dtVARCHAR
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
            
                  Case dtLONGVARCHAR
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                  
                  Case dtINTEGER
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                  
                  Case dtNUMERIC
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                  
                  Case dtTIMESTAMP
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & ", 101), '')"
                  
                  Case dtBIT
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                  
                  Case dtVARBINARY, dtLONGVARBINARY
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                  
                  Case Else
                    sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                      String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                End Select
              
                sDeleteTriggerInsertCode = sDeleteTriggerInsertCode & vbNewLine & _
                  String(iIndent, vbTab) & "INSERT INTO dbo.[ASRSysWorkflowQueueColumns]" & vbNewLine & _
                  String(iIndent + 1, vbTab) & "(queueID, columnID, columnValue, emailID)" & vbNewLine & _
                  String(iIndent + 1, vbTab) & "SELECT max(queueID), " & CStr(lngColumnID) & ", " & strVariableName & ", 0 FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine
              Next lngLoop
            
              alngEmailsUsed = BaseTableEmailAddressesUsedInDeleteTriggeredWorkflow(!WorkflowID)
              For lngLoop = 1 To UBound(alngEmailsUsed, 2)
                lngEmailID = alngEmailsUsed(0, lngLoop)

                strVariableName = "@sWFTempEmail_" & CStr(!LinkID) & "_" & CStr(lngEmailID)

                sSubCode = sSubCode & vbNewLine & _
                  String(iIndent, vbTab) & "DECLARE " & strVariableName & " varchar(MAX)" & vbNewLine

                If alngEmailsUsed(1, lngLoop) = 1 Then
                  'Column
                  strColumnName = GetColumnName(alngEmailsUsed(2, lngLoop), True)
                  
                  sDeleteTriggerSelectCode = sDeleteTriggerSelectCode & IIf(Len(sDeleteTriggerSelectCode) = 0, "", ", ") & vbNewLine & _
                    String(iIndent + 1, vbTab) & strVariableName & " = ISNULL(CONVERT(varchar(MAX), deleted." & strColumnName & "), '')"
                    
                  sDeleteTriggerInsertCode = sDeleteTriggerInsertCode & vbNewLine & _
                    String(iIndent, vbTab) & "INSERT INTO [dbo].[ASRSysWorkflowQueueColumns]" & vbNewLine & _
                    String(iIndent + 1, vbTab) & "(queueID, emailID, columnValue, columnID)" & vbNewLine & _
                    String(iIndent + 1, vbTab) & "SELECT max(queueID), " & CStr(lngEmailID) & ", " & strVariableName & ", 0 FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine
                Else
                  'Calculated
                  Set objExpr = New CExpression
                  With objExpr
                    .ExpressionID = alngEmailsUsed(2, lngLoop)
                    If .ConstructExpression Then
                      sSubCode = sSubCode & vbNewLine & _
                        String(iIndent, vbTab) & "SET @id = @recordID" & vbNewLine & _
                        .StoredProcedureCode(strVariableName, "deleted") & vbNewLine
                    End If
                  End With
    
                  sSubCode = sSubCode & vbNewLine & _
                    String(iIndent, vbTab) & "INSERT INTO dbo.[ASRSysWorkflowQueueColumns]" & vbNewLine & _
                    String(iIndent + 1, vbTab) & "(queueID, emailID, columnValue, columnID)" & vbNewLine & _
                    String(iIndent + 1, vbTab) & "SELECT max(queueID), " & CStr(lngEmailID) & ", " & strVariableName & ", 0 FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine
                End If
              Next lngLoop
              
              If Len(sDeleteTriggerSelectCode) > 0 Then
                sSubCode = sSubCode & vbNewLine & _
                  String(iIndent, vbTab) & "SELECT" & sDeleteTriggerSelectCode & vbNewLine & _
                  String(iIndent, vbTab) & "FROM deleted" & vbNewLine & _
                  String(iIndent, vbTab) & "WHERE deleted.id = @recordid" & vbNewLine & _
                  sDeleteTriggerInsertCode
              End If
            End If

            
            ' Add the filter code (if required)
            If sFilter <> vbNullString Then
              iIndent = iIndent - 1
              
              sSubCode = _
                String(iIndent, vbTab) & "IF " & sFilter & vbNewLine & _
                String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
                sSubCode & _
                String(iIndent, vbTab) & "END" & vbNewLine
            End If
            
            If sEffectiveDate <> vbNullString Then
              iIndent = iIndent - 1
              
              sSubCode = _
                String(iIndent, vbTab) & "IF DateDiff(day, '" & sEffectiveDate & "', getDate()) >= 0" & vbNewLine & _
                String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
                sSubCode & _
                String(iIndent, vbTab) & "END" & vbNewLine
            End If
            
            sCode = sCode & vbNewLine & _
              sSubCode
          End If
        End If

        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop
    End If
  End With

  If iCount > 0 Then
    sCode = _
      String(iIndent, vbTab) & "/* Table Level Workflow Trigger" & IIf(iCount > 1, "s", "") & " */" & _
      sCode
  End If
  
  If Len(sCode) > 0 Then
    sSubCode = String(iIndent, vbTab) & "SET @parent1TableID = 0" & vbNewLine & _
      String(iIndent, vbTab) & "SET @parent1RecordID = 0" & vbNewLine & _
      String(iIndent, vbTab) & "SET @parent2TableID = 0" & vbNewLine & _
      String(iIndent, vbTab) & "SET @parent2RecordID = 0" & vbNewLine & vbNewLine

    sSQL = "SELECT TOP 2 parentID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.childID = " & CStr(plngTableID)
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    iCount = 1
    Do While Not rsTables.EOF
      sSubCode = sSubCode & _
        String(iIndent, vbTab) & "SELECT @parent" & CStr(iCount) & "TableID = " & CStr(rsTables!parentID) & "," & vbNewLine & _
        String(iIndent, vbTab) & vbTab & "@parent" & CStr(iCount) & "RecordID = isnull(ID_" & CStr(rsTables!parentID) & ", 0)" & vbNewLine & _
        String(iIndent, vbTab) & vbTab & "FROM " & IIf(pfAction = WFRELATEDRECORD_DELETE, "deleted", GetTableName(plngTableID)) & vbNewLine & _
        String(iIndent, vbTab) & vbTab & "WHERE ID = @recordID" & vbNewLine & vbNewLine
     
      iCount = iCount + 1
      rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
  
    sCode = _
      sSubCode & _
      sCode
  End If

  WorkflowTableTriggerCode = sCode
  
End Function



Private Sub CreateWorkflowProcsForLink(lngTableID As Long, _
  sCurrentTable As String, _
  lngLinkID As Long, _
  lngRecordDescExprID As Long, _
  ByRef alngAuditColumns As Variant, _
  ByRef sDeclareInsCols As SystemMgr.cStringBuilder, _
  ByRef sDeclareDelCols As SystemMgr.cStringBuilder, _
  ByRef sSelectInsCols As SystemMgr.cStringBuilder, _
  ByRef sSelectDelCols As SystemMgr.cStringBuilder, _
  ByRef sFetchInsCols As SystemMgr.cStringBuilder, _
  ByRef sFetchDelCols As SystemMgr.cStringBuilder)

  Dim strColumnName As String
  Dim fColumnOK As Boolean
  Dim sConvertInsCols As String
  Dim fColFound As Boolean
  Dim iLoop As Integer
  Dim iColumnLoop As Integer
  Dim iDataType As Integer
  Dim lngSize As Long
  Dim iDecimals As Integer
  Dim lngColumnID As Long
  Dim iLinkType As WorkflowTriggerLinkType
  Dim avarColumns() As Variant
  Dim strCheckCode As String
  Dim strTriggerCode As String
  Dim iIndent As Integer
  Dim sEffectiveDate As String
  Dim sFilter As String
  Dim sImmediate As String
  Dim strTemp As String
  Dim strRebuildTemp As String
  Dim strRebuildDeclare As String
  Dim strVariableName As String
  Dim strColumnValuesInsert As String
  Dim lngRecDescID As Long
  
  On Error GoTo LocalErr

  lngRecDescID = IIf(IsNull(recTabEdit!RecordDescExprID), 0, recTabEdit!RecordDescExprID)

  iLinkType = recWorkflowTriggeredLinks!Type
  iIndent = 2
  
  ReDim avarColumns(4, 0)
  ' Column 0 = ColumnID
  ' Column 1 = ColumnName
  ' Column 2 = DataType
  ' Column 3 = Size
  ' Column 4 = Decimals
  
  Select Case iLinkType
    Case WORKFLOWTRIGGERLINKTYPE_COLUMN
      With recWorkflowTriggeredLinkColumns
        .Index = "idxLinkID"
        .Seek "=", recWorkflowTriggeredLinks!LinkID

        If Not .NoMatch Then
          Do While Not .EOF
            If !LinkID <> recWorkflowTriggeredLinks!LinkID Then
              Exit Do
            End If
            
            With recColEdit
              .Index = "idxColumnID"
              .Seek "=", IIf(IsNull(recWorkflowTriggeredLinkColumns!ColumnID), 0, recWorkflowTriggeredLinkColumns!ColumnID)
          
              If Not .NoMatch Then
                ReDim Preserve avarColumns(4, UBound(avarColumns, 2) + 1)
                avarColumns(0, UBound(avarColumns, 2)) = !ColumnID
                avarColumns(1, UBound(avarColumns, 2)) = !ColumnName
                avarColumns(2, UBound(avarColumns, 2)) = !DataType
                avarColumns(3, UBound(avarColumns, 2)) = !Size
                avarColumns(4, UBound(avarColumns, 2)) = !Decimals
              End If
            End With
            
            .MoveNext
          Loop
        End If
      End With
      
    Case WORKFLOWTRIGGERLINKTYPE_DATE
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", IIf(IsNull(recWorkflowTriggeredLinks!DateColumn), 0, recWorkflowTriggeredLinks!DateColumn)
    
        If Not .NoMatch Then
          ReDim Preserve avarColumns(4, UBound(avarColumns, 2) + 1)
          avarColumns(0, UBound(avarColumns, 2)) = !ColumnID
          avarColumns(1, UBound(avarColumns, 2)) = !ColumnName
          avarColumns(2, UBound(avarColumns, 2)) = !DataType
          avarColumns(3, UBound(avarColumns, 2)) = !Size
          avarColumns(4, UBound(avarColumns, 2)) = !Decimals
        End If
      End With
  End Select
  
  If UBound(avarColumns, 2) > 0 Then
    For iColumnLoop = 1 To UBound(avarColumns, 2)
      lngColumnID = CLng(avarColumns(0, iColumnLoop))
      strColumnName = CStr(avarColumns(1, iColumnLoop))
      iDataType = CInt(avarColumns(2, iColumnLoop))
      lngSize = avarColumns(3, iColumnLoop)
      iDecimals = CInt(avarColumns(4, iColumnLoop))
  
      sConvertInsCols = ""

      fColFound = False

      ' Check if the column has already been declared and added to the select and fetch strings
      For iLoop = 1 To UBound(alngAuditColumns)
        If alngAuditColumns(iLoop) = lngColumnID Then
          fColFound = True
          Exit For
        End If
      Next iLoop

      If Not fColFound Then
        ReDim Preserve alngAuditColumns(UBound(alngAuditColumns) + 1)
        alngAuditColumns(UBound(alngAuditColumns)) = lngColumnID

        sSelectInsCols.Append ", inserted." & strColumnName
        sSelectDelCols.Append ", deleted." & strColumnName
        sFetchInsCols.Append ", @insCol_" & Trim$(Str$(lngColumnID))
        sFetchDelCols.Append ", @delCol_" & Trim$(Str$(lngColumnID))
  
        sDeclareInsCols.Append "," & vbNewLine & "        @insCol_" & Trim$(Str$(lngColumnID))
        sDeclareDelCols.Append "," & vbNewLine & "        @delCol_" & Trim$(Str$(lngColumnID))
      End If

      Select Case iDataType
        Case dtVARCHAR
          If Not fColFound Then
            sDeclareInsCols.Append " varchar(MAX)"
            sDeclareDelCols.Append " varchar(MAX)"
          End If
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"

        Case dtLONGVARCHAR
          If Not fColFound Then
            sDeclareInsCols.Append " varchar(14)"
            sDeclareDelCols.Append " varchar(14)"
          End If
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"

        Case dtINTEGER
          If Not fColFound Then
            sDeclareInsCols.Append " integer"
            sDeclareDelCols.Append " integer"
          End If
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"

        Case dtNUMERIC
          If Not fColFound Then
            sDeclareInsCols.Append " numeric(" & Trim$(Str$(lngSize)) & ", " & Trim$(Str$(iDecimals)) & ")"
            sDeclareDelCols.Append " numeric(" & Trim$(Str$(lngSize)) & ", " & Trim$(Str$(iDecimals)) & ")"
          End If
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"

        Case dtTIMESTAMP
          If Not fColFound Then
            sDeclareInsCols.Append " datetime"
            sDeclareDelCols.Append " datetime"
          End If
          'sConvertInsCols = "ISNULL(CONVERT(varchar(3000), LEFT(DATENAME(month, @insCol_" & Trim$(Str$(lngColumnID)) & "),3) + ' ' + CONVERT(varchar(3000),DATEPART(day, @insCol_" & Trim$(Str$(lngColumnID)) & ")) + ' ' + CONVERT(varchar(3000),DATEPART(year, @insCol_" & Trim$(Str$(lngColumnID)) & "))), '')"
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & ", 101), '')"

        Case dtBIT
          If Not fColFound Then
            sDeclareInsCols.Append " bit"
            sDeclareDelCols.Append " bit"
          End If
          'sConvertInsCols = "ISNULL(CONVERT(varchar(3000), CASE @insCol_" & Trim$(Str$(lngColumnID)) & " WHEN 1 THEN 'True' WHEN 0 THEN 'False' END), '')"
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"

        Case dtVARBINARY, dtLONGVARBINARY
          If Not fColFound Then
            sDeclareInsCols.Append " varchar(3000)"
            sDeclareDelCols.Append " varchar(3000)"
          End If
          sConvertInsCols = "ISNULL(CONVERT(varchar(3000), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"

        Case Else
          If Not fColFound Then
            sDeclareInsCols.Append " varchar(MAX)"
            sDeclareDelCols.Append " varchar(MAX)"
          End If
          sConvertInsCols = "ISNULL(CONVERT(varchar(MAX), @insCol_" & Trim$(Str$(lngColumnID)) & "), '')"
      End Select
  
      strVariableName = "@sWFTemp_" & CStr(lngLinkID) & "_" & CStr(lngColumnID)
      strRebuildDeclare = strRebuildDeclare & vbNewLine & _
        vbTab & vbTab & "DECLARE " & strVariableName & " varchar(MAX)" & vbNewLine




      'strCheckCode = strCheckCode & vbNewLine & _
        vbTab & vbTab & "DECLARE " & strVariableName & " varchar(MAX)" & vbNewLine & _
        vbTab & vbTab & "SET " & strVariableName & " = " & sConvertInsCols & vbNewLine & _
        vbTab & vbTab & "IF (@insCol_" & Trim$(Str$(lngColumnID)) & " <> @delCol_" & Trim$(Str$(lngColumnID)) & ") OR " & vbNewLine & _
        vbTab & vbTab & "  ((@insCol_" & Trim$(Str$(lngColumnID)) & " IS null) AND (NOT @delCol_" & Trim$(Str$(lngColumnID)) & " IS null)) OR " & vbNewLine & _
        vbTab & vbTab & "  ((NOT @insCol_" & Trim$(Str$(lngColumnID)) & " IS null) AND (@delCol_" & Trim$(Str$(lngColumnID)) & " IS null))" & vbNewLine & _
        vbTab & vbTab & "BEGIN" & vbNewLine & _
        vbTab & vbTab & vbTab & "SET @fWFTrigger = 1" & vbNewLine & _
        vbTab & vbTab & "END" & vbNewLine
      strCheckCode = strCheckCode & vbNewLine & _
        vbTab & vbTab & "DECLARE " & strVariableName & " varchar(MAX)" & vbNewLine & _
        vbTab & vbTab & "SET " & strVariableName & " = " & sConvertInsCols & vbNewLine

      Select Case iDataType
      Case dtNUMERIC, dtINTEGER, dtBIT
        strCheckCode = strCheckCode & vbNewLine & _
          vbTab & vbTab & "IF (isnull(@insCol_" & CStr(lngColumnID) & ",0) <> isnull(@delCol_" & CStr(lngColumnID) & ",0))" & vbNewLine
      Case Else
        strCheckCode = strCheckCode & vbNewLine & _
          vbTab & vbTab & "IF (isnull(@insCol_" & CStr(lngColumnID) & ",'') <> isnull(@delCol_" & CStr(lngColumnID) & ",''))" & vbNewLine
      End Select
        
      strCheckCode = strCheckCode & vbNewLine & _
        vbTab & vbTab & "BEGIN" & vbNewLine & _
        vbTab & vbTab & vbTab & "SET @fWFTrigger = 1" & vbNewLine & _
        vbTab & vbTab & "END" & vbNewLine




      strColumnValuesInsert = strColumnValuesInsert & vbNewLine & _
        vbTab & vbTab & "INSERT INTO dbo.[ASRSysWorkflowQueueColumns]" & vbNewLine & _
        vbTab & vbTab & vbTab & "(queueID, columnID, columnValue, emailID)" & vbNewLine & _
        vbTab & vbTab & vbTab & "SELECT max(queueID), " & CStr(lngColumnID) & ", " & strVariableName & ", 0 FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine
    Next iColumnLoop
  
    If Not IsNull(recWorkflowTriggeredLinks!EffectiveDate) Then
      sEffectiveDate = Replace(Format(recWorkflowTriggeredLinks!EffectiveDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
      iIndent = iIndent + 1
    Else
      sEffectiveDate = vbNullString
    End If

    If recWorkflowTriggeredLinks!FilterID > 0 Then
      sFilter = GetSQLFilter(recWorkflowTriggeredLinks!FilterID, GetTableName(lngTableID))
      iIndent = iIndent + 1
    Else
      sFilter = vbNullString
    End If
    
    sImmediate = _
      String(iIndent + 1, vbTab) & "INSERT dbo.[ASRSysWorkflowQueue](LinkID,RecordID,DateDue,UserName,[Immediate],RecalculateRecordDesc, recordDesc, parent1TableID, parent1RecordID, parent2TableID, parent2RecordID, instanceID)" & vbNewLine & _
      String(iIndent + 1, vbTab) & "VALUES (" & CStr(lngLinkID) & ",@recordID, getDate()," & _
        "CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE ltrim(rtrim(SYSTEM_USER)) END," & _
        "1, 1, @recordDesc, @parent1TableID, @parent1RecordID, @parent2TableID, @parent2RecordID, 0)" & vbNewLine & _
      strColumnValuesInsert

    strRebuildTemp = _
      String(iIndent, vbTab) & "INSERT dbo.[ASRSysWorkflowQueue](LinkID,RecordID,DateDue,UserName,[Immediate],RecalculateRecordDesc, recordDesc, parent1TableID, parent1RecordID, parent2TableID, parent2RecordID, instanceID)" & vbNewLine & _
      String(iIndent, vbTab) & "VALUES (" & CStr(lngLinkID) & ",@recordID, @dtWFLinkDate," & _
        "CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE ltrim(rtrim(SYSTEM_USER)) END," & _
        "0, 1, @recordDesc, @parent1TableID, @parent1RecordID, @parent2TableID, @parent2RecordID, 0)" & vbNewLine & _
      strColumnValuesInsert

    Select Case iLinkType
      Case WORKFLOWTRIGGERLINKTYPE_COLUMN
        strTriggerCode = sImmediate

        ' Add the filter code (if required)
        If sFilter <> vbNullString Then
          iIndent = iIndent - 1

          strTriggerCode = _
            String(iIndent, vbTab) & "IF " & sFilter & vbNewLine & _
            String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
            strTriggerCode & _
            String(iIndent, vbTab) & "END" & vbNewLine
        End If

        If sEffectiveDate <> vbNullString Then
          iIndent = iIndent - 1

          strTriggerCode = _
            String(iIndent + 1, vbTab) & "IF DateDiff(day, '" & sEffectiveDate & "', getDate()) >= 0" & vbNewLine & _
            String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine & _
            strTriggerCode & _
            String(iIndent + 1, vbTab) & "END" & vbNewLine
        End If

      Case WORKFLOWTRIGGERLINKTYPE_DATE
'        strTriggerCode = _
'          String(iIndent, vbTab) & "IF (DateDiff(day, @dtWFLinkDate, getdate()) >= 0) OR" & vbNewLine & _
'          String(iIndent, vbTab) & vbTab & "(@sWFLastSent IS NOT NULL)" & vbNewLine & _
'          String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
'          sImmediate & vbNewLine & _
'          String(iIndent, vbTab) & "END" & vbNewLine & _
'          String(iIndent, vbTab) & "ELSE" & vbNewLine & _
'          String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
'          strRebuildTemp & vbNewLine & _
'          String(iIndent, vbTab) & "END" & vbNewLine
        strTriggerCode = _
          String(iIndent, vbTab) & "IF (DateDiff(day, @dtWFLinkDate, getdate()) >= 0)" & vbNewLine & _
          String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
          sImmediate & vbNewLine & _
          String(iIndent, vbTab) & "END" & vbNewLine & _
          String(iIndent, vbTab) & "ELSE" & vbNewLine & _
          String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
          strRebuildTemp & vbNewLine & _
          String(iIndent, vbTab) & "END" & vbNewLine

        strTemp = _
          String(iIndent + 1, vbTab) & "SELECT TOP 1 @sWFLastSent = ASRSysWorkflowQueueColumns.ColumnValue" & vbNewLine & _
          String(iIndent + 1, vbTab) & "FROM dbo.[ASRSysWorkflowQueueColumns]" & vbNewLine & _
          String(iIndent + 1, vbTab) & "INNER JOIN ASRSysWorkflowQueue ON ASRSysWorkflowQueueColumns.queueID = ASRSysWorkflowQueue.queueID" & vbNewLine & _
          String(iIndent + 1, vbTab) & "WHERE ASRSysWorkflowQueue.recordID = @recordid" & vbNewLine & _
          String(iIndent + 2, vbTab) & "AND ASRSysWorkflowQueue.linkID = " & CStr(lngLinkID) & vbNewLine & _
          String(iIndent + 2, vbTab) & "AND ASRSysWorkflowQueueColumns.columnID = " & CStr(lngColumnID) & vbNewLine & vbNewLine & _
          String(iIndent + 1, vbTab) & "ORDER BY ASRSysWorkflowQueue.dateInitiated DESC" & vbNewLine & vbNewLine & _
          String(iIndent + 1, vbTab) & "IF ((DateDiff(day, @dtWFPurgeDate, @dtWFLinkDate) >= 0 OR @dtWFPurgeDate IS NULL)" & vbNewLine
        
'        strTriggerCode = strTemp & _
'          String(iIndent + 2, vbTab) & "OR (@sWFLastSent IS NOT NULL)) " & vbNewLine & _
'          IIf(sEffectiveDate <> vbNullString, _
'            String(iIndent + 2, vbTab) & "AND (DateDiff(day, '" & sEffectiveDate & "', @dtWFLinkDate) >= 0)", "") & vbNewLine & _
'          String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine & _
'          strTriggerCode & vbNewLine & _
'          String(iIndent + 1, vbTab) & "END" & vbNewLine
        strTriggerCode = strTemp & _
          String(iIndent + 2, vbTab) & ") " & vbNewLine & _
          IIf(sEffectiveDate <> vbNullString, _
            String(iIndent + 2, vbTab) & "AND (DateDiff(day, '" & sEffectiveDate & "', @dtWFLinkDate) >= 0)", "") & vbNewLine & _
          String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine & _
          strTriggerCode & vbNewLine & _
          String(iIndent + 1, vbTab) & "END" & vbNewLine

        strRebuildTemp = strTemp & _
          String(iIndent + 2, vbTab) & ") AND (IsNull(@sWFLastSent,'') <> IsNull(" & strVariableName & ",''))" & vbNewLine & _
          IIf(sEffectiveDate <> vbNullString, _
            String(iIndent + 2, vbTab) & "AND (DateDiff(day, '" & sEffectiveDate & "', @dtWFLinkDate) >= 0)", "") & vbNewLine & _
          String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine & _
          strRebuildTemp & vbNewLine & _
          String(iIndent + 1, vbTab) & "END" & vbNewLine
    
        If Abs(recWorkflowTriggeredLinks!DateOffset) Then
          strTemp = _
            String(iIndent, vbTab) & "SET @dtWFLinkDate =" & vbNewLine
          
          If recWorkflowTriggeredLinks!DateOffset < 0 Then
            strTemp = strTemp + _
              String(iIndent + 1, vbTab) & "CASE" & vbNewLine & _
              String(iIndent + 2, vbTab) & "WHEN dateadd(" & _
                Choose(recWorkflowTriggeredLinks!DateOffsetPeriod + 1, "dd", "ww", "mm", "yy") & "," & _
                -recWorkflowTriggeredLinks!DateOffset & ", '01/01/1753')" & _
                " > @dtWFLinkDate THEN '01/01/1753'" & vbNewLine & _
              String(iIndent + 2, vbTab) & "ELSE dateadd(" & _
                Choose(recWorkflowTriggeredLinks!DateOffsetPeriod + 1, "dd", "ww", "mm", "yy") & "," & _
                recWorkflowTriggeredLinks!DateOffset & ", @dtWFLinkDate)" & vbNewLine & _
              String(iIndent + 1, vbTab) & "END" & vbNewLine & vbNewLine
          ElseIf recWorkflowTriggeredLinks!DateOffset > 0 Then
            strTemp = strTemp + _
              String(iIndent + 1, vbTab) & "CASE" & vbNewLine & _
              String(iIndent + 2, vbTab) & "WHEN dateadd(" & _
                Choose(recWorkflowTriggeredLinks!DateOffsetPeriod + 1, "dd", "ww", "mm", "yy") & "," & _
                -recWorkflowTriggeredLinks!DateOffset & ", '12/31/9999')" & _
                " < @dtWFLinkDate THEN '12/31/9999'" & vbNewLine & _
              String(iIndent + 2, vbTab) & "ELSE dateadd(" & _
                Choose(recWorkflowTriggeredLinks!DateOffsetPeriod + 1, "dd", "ww", "mm", "yy") & "," & _
                recWorkflowTriggeredLinks!DateOffset & ", @dtWFLinkDate)" & vbNewLine & _
              String(iIndent + 1, vbTab) & "END" & vbNewLine & vbNewLine
          Else
            strTemp = strTemp + _
              String(iIndent + 1, vbTab) & "dateadd(" & _
                Choose(recWorkflowTriggeredLinks!DateOffsetPeriod + 1, "dd", "ww", "mm", "yy") & "," & _
                recWorkflowTriggeredLinks!DateOffset & ", @dtWFLinkDate)" & vbNewLine & vbNewLine
          End If
          
          strTriggerCode = strTemp & strTriggerCode
          strRebuildTemp = strTemp & strRebuildTemp
        End If

      strTriggerCode = _
        String(iIndent, vbTab) & "IF NOT @insCol_" & CStr(lngColumnID) & " IS null" & vbNewLine & _
        String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
        String(iIndent + 1, vbTab) & "SET @dtWFLinkDate = IsNull(convert(datetime,@insCol_" & CStr(lngColumnID) & "),getdate())" & vbNewLine & vbNewLine & _
        strTriggerCode & _
        String(iIndent, vbTab) & "END" & vbNewLine

      strRebuildTemp = _
        String(iIndent, vbTab) & "SELECT @dtWFLinkDate = " & strColumnName & "" & vbNewLine & _
        String(iIndent, vbTab) & "FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewLine & vbNewLine & _
        String(iIndent, vbTab) & "IF NOT @dtWFLinkDate IS null" & vbNewLine & _
        String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
        String(iIndent + 1, vbTab) & "SET " & strVariableName & " = CONVERT(varchar(3000), @dtWFLinkDate, 101)" & vbNewLine & _
        String(iIndent + 1, vbTab) & "SELECT @dtWFLinkDate = IsNull(convert(datetime," & strColumnName & "),getdate()) FROM " & sCurrentTable & " WHERE id = @recordID" & vbNewLine & vbNewLine & _
        strRebuildTemp & _
        String(iIndent, vbTab) & "END" & vbNewLine

        ' Add the filter code (if required)
        If sFilter <> vbNullString Then
          iIndent = iIndent - 1

          strTriggerCode = _
            String(iIndent, vbTab) & "IF " & sFilter & vbNewLine & _
            String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
            strTriggerCode & _
            String(iIndent, vbTab) & "END" & vbNewLine

          strRebuildTemp = _
            String(iIndent, vbTab) & "IF " & sFilter & vbNewLine & _
            String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
            strRebuildTemp & _
            String(iIndent, vbTab) & "END" & vbNewLine
        End If

      msRebuildLinkCode = strRebuildDeclare & _
        msRebuildLinkCode & vbNewLine & _
        String(iIndent, vbTab) & "-- " & GetWorkflowName(recWorkflowTriggeredLinks!WorkflowID) & vbNewLine & _
        strRebuildTemp & vbNewLine
    End Select


    'MH20070730
    'msInsertLinkTemp = vbNewLine & _
      vbTab & vbTab & "-- " & GetWorkflowName(recWorkflowTriggeredLinks!WorkflowID) & vbNewLine & _
      vbTab & vbTab & "SET @fWFTrigger = 0" & vbNewLine & _
      strCheckCode & vbNewLine & _
      vbTab & vbTab & "IF @fWFTrigger = 1" & vbNewLine & _
      vbTab & vbTab & "BEGIN" & vbNewLine & _
      vbTab & vbTab & vbTab & "IF (@fUpdatingDateDependentColumns = 1) AND (len(@recordDesc) = 0)" & vbNewLine & _
      vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & "IF EXISTS(SELECT Name FROM sysobjects WHERE type = 'P' AND name = 'sp_ASRExpr_" & Trim$(Str$(lngRecDescID)) & "')" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & vbTab & "EXEC @hResult = dbo.sp_ASRExpr_" & Trim$(Str$(lngRecDescID)) & " @recordDesc OUTPUT, @recordID" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & vbTab & "IF @hResult <> 0 SET @recordDesc = ''" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & vbTab & "SET @recordDesc = CONVERT(varchar(255), @recordDesc)" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & "END" & vbNewLine & _
      vbTab & vbTab & vbTab & vbTab & "ELSE SET @recordDesc = ''" & vbNewLine & _
      vbTab & vbTab & vbTab & "END" & vbNewLine & vbNewLine & _
      vbTab & vbTab & vbTab & "DELETE FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine & _
      vbTab & vbTab & vbTab & "WHERE dateInitiated IS Null AND recordID=@recordID AND linkID = " & CStr(lngLinkID) & vbNewLine & vbNewLine & _
      strTriggerCode & vbNewLine & _
      vbTab & vbTab & "END" & vbNewLine
    msInsertLinkTemp = vbNewLine & _
      vbTab & vbTab & "-- " & GetWorkflowName(recWorkflowTriggeredLinks!WorkflowID) & vbNewLine & _
      vbTab & vbTab & "SET @fWFTrigger = 0" & vbNewLine & _
      strCheckCode & vbNewLine & _
      vbTab & vbTab & "IF @fWFTrigger = 1" & vbNewLine & _
      vbTab & vbTab & "BEGIN" & vbNewLine & _
      vbTab & vbTab & vbTab & "DELETE FROM dbo.[ASRSysWorkflowQueue]" & vbNewLine & _
      vbTab & vbTab & vbTab & "WHERE dateInitiated IS Null AND recordID=@recordID AND linkID = " & CStr(lngLinkID) & vbNewLine & vbNewLine & _
      strTriggerCode & vbNewLine & _
      vbTab & vbTab & "END" & vbNewLine
  
  
  
    msUpdateLinkTemp = msInsertLinkTemp
  End If
  
Exit Sub

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
  End If

End Sub



Public Sub CreateWorkflowProcsForTable(pLngCurrentTableID As Long, _
  sCurrentTable As String, _
  lngRecordDescExprID As Long, _
  ByRef alngAuditColumns As Variant, _
  ByRef sDeclareInsCols As SystemMgr.cStringBuilder, _
  ByRef sDeclareDelCols As SystemMgr.cStringBuilder, _
  ByRef sSelectInsCols As SystemMgr.cStringBuilder, _
  ByRef sSelectDelCols As SystemMgr.cStringBuilder, _
  ByRef sFetchInsCols As SystemMgr.cStringBuilder, _
  ByRef sFetchDelCols As SystemMgr.cStringBuilder)
  
  Dim sTemp As String
  Dim sTemp_Trigger As String
  Dim sTemp_Rebuild As String
  Dim lngRecDescID As Long
  Dim sSQL As String
  Dim sSubCode As String
  Dim rsTables As DAO.Recordset
  Dim iCount As Integer
  
  On Error GoTo LocalErr

  recTabEdit.Index = "idxTableID"
  recTabEdit.Seek "=", pLngCurrentTableID
  lngRecDescID = IIf(IsNull(recTabEdit!RecordDescExprID), 0, recTabEdit!RecordDescExprID)

  msInsertLinkCode = vbNullString
  msUpdateLinkCode = vbNullString
  msRebuildLinkCode = vbNullString

  With recWorkflowTriggeredLinks
    .Index = "idxTableID"
    .Seek "=", pLngCurrentTableID

    If Not .NoMatch Then
      Do While !TableID = pLngCurrentTableID
        If (Not !Deleted) _
          And ((!Type = WORKFLOWTRIGGERLINKTYPE_COLUMN) _
            Or (!Type = WORKFLOWTRIGGERLINKTYPE_DATE)) Then

          recWorkflowEdit.Index = "idxWorkflowID"
          recWorkflowEdit.Seek "=", !WorkflowID
          
          If Not recWorkflowEdit.NoMatch Then
            If recWorkflowEdit!Enabled And (Not recWorkflowEdit!Deleted) Then

              msInsertLinkTemp = vbNullString
              msUpdateLinkTemp = vbNullString
              
              CreateWorkflowProcsForLink pLngCurrentTableID, sCurrentTable, !LinkID, lngRecordDescExprID, alngAuditColumns, _
                sDeclareInsCols, sDeclareDelCols, _
                sSelectInsCols, sSelectDelCols, _
                sFetchInsCols, sFetchDelCols
    
              msInsertLinkCode = msInsertLinkCode & msInsertLinkTemp
              msUpdateLinkCode = msUpdateLinkCode & msUpdateLinkTemp
            End If
          End If
        End If
        
        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop
    End If
  End With
  
  sTemp = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('[dbo].[spASRWorkflowRebuild_" & CStr(pLngCurrentTableID) & "]')" & _
    "     AND sysstat & 0xf = 4)" & _
    " DROP PROCEDURE [dbo].[spASRWorkflowRebuild_" & CStr(pLngCurrentTableID) & "]"
  gADOCon.Execute sTemp, , adExecuteNoRecords

  If msUpdateLinkCode <> vbNullString Then
    sTemp = vbNewLine & _
      vbTab & vbTab & "DECLARE @dtWFLinkDate datetime," & vbNewLine & _
      vbTab & vbTab & vbTab & "@dtWFPurgeDate datetime," & vbNewLine & _
      vbTab & vbTab & vbTab & "@sWFLastSent varchar(MAX)," & vbNewLine & _
      vbTab & vbTab & vbTab & "@fWFTrigger bit," & vbNewLine & _
      vbTab & vbTab & vbTab & "@sWFUserName varchar(255)" & vbNewLine & vbNewLine & _
      vbTab & vbTab & "SELECT @sWFUserName = rtrim(system_user)" & vbNewLine & vbNewLine & _
      vbTab & vbTab & "EXEC [dbo].[sp_ASRPurgeDate] @dtWFPurgeDate OUTPUT, 'WORKFLOW'" & vbNewLine
    
    sTemp_Trigger = ""
    sTemp_Rebuild = ""
    
    sSQL = "SELECT TOP 2 parentID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.childID = " & CStr(pLngCurrentTableID)
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    iCount = 1
    Do While Not rsTables.EOF
      sTemp_Trigger = sTemp_Trigger & _
        vbTab & vbTab & "SET @parent" & CStr(iCount) & "TableID = " & CStr(rsTables!parentID) & vbNewLine & _
        vbTab & vbTab & "SET @parent" & CStr(iCount) & "RecordID = @insParentID_" & CStr(rsTables!parentID) & vbNewLine & vbNewLine
      
      sTemp_Rebuild = sTemp_Rebuild & _
        IIf(Len(sTemp_Rebuild) > 0, ",", "") & _
        vbTab & vbTab & vbTab & "@parent" & CStr(iCount) & "TableID = " & CStr(rsTables!parentID) & "," & vbNewLine & _
        vbTab & vbTab & vbTab & "@parent" & CStr(iCount) & "RecordID = isnull(ID_" & CStr(rsTables!parentID) & ", 0)"

      'sTemp = sTemp & _
      '  vbTab & vbTab & "SELECT @parent" & CStr(iCount) & "TableID = " & CStr(rsTables!parentID) & "," & vbNewLine & _
      '  vbTab & vbTab & vbTab & "@parent" & CStr(iCount) & "RecordID = isnull(ID_" & CStr(rsTables!parentID) & ", 0)" & vbNewLine & _
      '  vbTab & vbTab & vbTab & "FROM " & GetTableName(pLngCurrentTableID) & vbNewLine & _
      '  vbTab & vbTab & vbTab & "WHERE ID = @recordid" & vbNewLine & vbNewLine

      iCount = iCount + 1
      rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
    
    If Len(sTemp_Rebuild) > 0 Then
      sTemp_Rebuild = _
        vbTab & vbTab & "SELECT" & vbNewLine & _
        sTemp_Rebuild & vbNewLine & _
        vbTab & vbTab & vbTab & "FROM " & GetTableName(pLngCurrentTableID) & vbNewLine & _
        vbTab & vbTab & vbTab & "WHERE ID = @recordid" & vbNewLine & vbNewLine
    End If
    
    msInsertLinkCode = sTemp & sTemp_Trigger & msInsertLinkCode
    msUpdateLinkCode = sTemp & sTemp_Trigger & msUpdateLinkCode

    If msRebuildLinkCode <> vbNullString Then
    
      sSubCode = _
        "/* -----------------------------------------------*/" & vbNewLine & _
        "/* Workflow Rebuild stored procedure.      */" & vbNewLine & _
        "/* Automatically generated by the System Manager. */" & vbNewLine & _
        "/* ---------------------------------------------- */" & vbNewLine & _
        "CREATE PROCEDURE [dbo].[spASRWorkflowRebuild_" & CStr(pLngCurrentTableID) & "]" & vbNewLine & _
        vbTab & "(@recordid int)" & vbNewLine & _
        "AS" & vbNewLine & _
        "BEGIN" & vbNewLine & _
        vbTab & "DECLARE @recordDesc varchar(MAX)," & vbNewLine & _
        vbTab & vbTab & "@hResult int," & vbNewLine & _
        vbTab & vbTab & "@parent1TableID integer," & vbNewLine & _
        vbTab & vbTab & "@parent1RecordID integer," & vbNewLine & _
        vbTab & vbTab & "@parent2TableID integer," & vbNewLine & _
        vbTab & vbTab & "@parent2RecordID integer" & vbNewLine & vbNewLine & _
        vbTab & "SET @parent1TableID = 0" & vbNewLine & _
        vbTab & "SET @parent1RecordID = 0" & vbNewLine & _
        vbTab & "SET @parent2TableID = 0" & vbNewLine & _
        vbTab & "SET @parent2RecordID = 0" & vbNewLine & vbNewLine

      msRebuildLinkCode = sSubCode & _
        sTemp & _
        sTemp_Rebuild & _
        vbTab & "IF EXISTS(SELECT Name FROM sysobjects WHERE sysstat & 0xf = 4 AND id = object_id('sp_ASRExpr_" & Trim$(Str$(lngRecDescID)) & "'))" & vbNewLine & _
        vbTab & "BEGIN" & vbNewLine & _
        vbTab & vbTab & "EXEC @hResult = [dbo].[sp_ASRExpr_" & Trim$(Str$(lngRecDescID)) & "] @recordDesc OUTPUT, @recordID" & vbNewLine & _
        vbTab & vbTab & "IF @hResult <> 0 SET @recordDesc = ''" & vbNewLine & _
        vbTab & vbTab & vbTab & "SET @recordDesc = CONVERT(varchar(255), @recordDesc)" & vbNewLine & _
        vbTab & vbTab & "END" & vbNewLine & _
        vbTab & vbTab & "ELSE SET @recordDesc = ''" & vbNewLine & _
        msRebuildLinkCode & vbNewLine & _
        "END"
    
      gADOCon.Execute "IF EXISTS (SELECT Name FROM sysobjects" & _
        "   WHERE id = object_id('spASRWorkflowRebuild_" & CStr(pLngCurrentTableID) & "')" & _
        "     AND sysstat & 0xf = 4)" & _
        " DROP PROCEDURE spASRWorkflowRebuild_" & CStr(pLngCurrentTableID), , adExecuteNoRecords
  
      gADOCon.Execute msRebuildLinkCode, , adExecuteNoRecords
    End If
  End If
  
  Exit Sub

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
  End If

End Sub


Public Function CreateSP_WorkflowParentRecord() As Boolean

  Const strSPName As String = "spASRSysWorkflowParentRecord"

  Dim strSQL As String
  Dim iIndent As Integer
  Dim sChildTableName As String
  Dim fOK As Boolean

  On Error GoTo ErrorTrap

  fOK = True
  iIndent = 1

  DropProcedure strSPName

  strSQL = vbNullString
  With recRelEdit
    If Not (.EOF And .BOF) Then
      .MoveFirst
      
      Do While Not .EOF
        sChildTableName = GetTableName(!childID)
        
        strSQL = strSQL & _
          String(iIndent, vbTab) & IIf(Len(strSQL) > 0, "", "") & "IF (@piChildTableID = " & CStr(!childID) & ") AND (@piParentTableID = " & CStr(!parentID) & ")" & vbNewLine & _
          String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
          String(iIndent + 1, vbTab) & "SELECT @piParentRecordID = isnull(ID_" & CStr(!parentID) & ", 0)" & vbNewLine & _
          String(iIndent + 1, vbTab) & "FROM " & sChildTableName & vbNewLine & _
          String(iIndent + 1, vbTab) & "WHERE ID = @piChildRecordID" & vbNewLine
  
        strSQL = strSQL & _
          String(iIndent, vbTab) & "END" & vbNewLine

        .MoveNext
      Loop
    End If
  End With

  iIndent = 1
  strSQL = _
    "------------------------------------------------------" & vbNewLine & _
    "-- Workflow parent record stored procedure." & vbNewLine & _
    "-- Automatically generated by the System Manager." & vbNewLine & _
    "------------------------------------------------------" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & strSPName & "]" & vbNewLine & _
    "(" & vbNewLine & _
    String(iIndent, vbTab) & "@piChildTableID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@piChildRecordID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@piParentTableID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@piParentRecordID integer OUTPUT" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    String(iIndent, vbTab) & "SET @piParentRecordID = 0" & vbNewLine & _
    strSQL & _
    "END"
  
  gADOCon.Execute strSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_WorkflowParentRecord = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Workflow parent record stored procedure"
  fOK = False
  Resume TidyUpAndExit

End Function



Public Function CreateSP_WorkflowCalculation() As Boolean

  Const strSPName As String = "spASRSysWorkflowCalculation"
  
  Dim strSQL As String
  Dim objStoredProc As SystemMgr.cStringBuilder
  Dim iIndent As Integer
  Dim strCalcSP As String
  Dim fOK As Boolean

  On Error GoTo ErrorTrap

  Set objStoredProc = New SystemMgr.cStringBuilder
  fOK = True
  iIndent = 1
  
  DropProcedure strSPName

  strSQL = vbNullString
  With recExprEdit
    .Index = "idxExprID"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    objStoredProc.TheString = vbNullString
    
    Do While Not .EOF
      
      If (Not !Deleted) _
        And (!Type = giEXPR_WORKFLOWCALCULATION) _
        And (!ParentComponentID = 0) Then
        
        strCalcSP = "[dbo].[sp_ASRExpr_" & CStr(!ExprID) & "]"
        
        objStoredProc.Append _
          String(iIndent, vbTab) & "IF @piExprID = " & CStr(!ExprID) & vbNewLine & _
          String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
          String(iIndent + 1, vbTab) & "-- " & Trim(.Fields("Name").value) & vbNewLine & _
          String(iIndent + 1, vbTab) & "SET @piResultType = " & CStr(!ReturnType) & vbNewLine & vbNewLine & _
          String(iIndent + 1, vbTab) & "IF EXISTS (SELECT Name FROM sysobjects WHERE sysstat & 0xf = 4 AND id = object_id('" & strCalcSP & "'))" & vbNewLine & _
          String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine
            
        Select Case !ReturnType
          Case giEXPRVALUE_NUMERIC  ' 2
            objStoredProc.Append _
              String(iIndent + 2, vbTab) & "EXEC " & strCalcSP & " @pfltResult OUTPUT, @piInstanceID, @piTempElement;" & vbNewLine & _
              String(iIndent + 2, vbTab) & "IF @pfltResult IS NULL SET @pfltResult = 0;" & vbNewLine
          
          Case giEXPRVALUE_LOGIC ' 3
            objStoredProc.Append _
              String(iIndent + 2, vbTab) & "EXEC " & strCalcSP & " @pfResult OUTPUT, @piInstanceID, @piTempElement" & vbNewLine & _
              String(iIndent + 2, vbTab) & "IF @pfResult IS NULL SET @pfResult = 0;" & vbNewLine
          
          Case giEXPRVALUE_DATE ' 4
            objStoredProc.Append _
              String(iIndent + 2, vbTab) & "EXEC " & strCalcSP & " @pdtResult OUTPUT, @piInstanceID, @piTempElement;" & vbNewLine

          Case Else ' giEXPRVALUE_CHARACTER 1
            objStoredProc.Append _
              String(iIndent + 2, vbTab) & "EXEC " & strCalcSP & " @psResult OUTPUT, @piInstanceID, @piTempElement;" & vbNewLine & _
              String(iIndent + 2, vbTab) & "IF @psResult IS NULL SET @psResult = '';" & vbNewLine
        End Select
        
        objStoredProc.Append _
          String(iIndent + 1, vbTab) & "END" & vbNewLine & _
          String(iIndent + 1, vbTab) & "RETURN" & vbNewLine & _
          String(iIndent, vbTab) & "END" & vbNewLine
      
      End If
           
      .MoveNext
    Loop
  End With
  
  iIndent = 1
  strSQL = _
    "------------------------------------------------------" & vbNewLine & _
    "-- Workflow calculation stored procedure." & vbNewLine & _
    "-- Automatically generated by the System Manager." & vbNewLine & _
    "------------------------------------------------------" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & strSPName & "]" & vbNewLine & _
    "(" & vbNewLine & _
    String(iIndent, vbTab) & "@piInstanceID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@piExprID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@piResultType integer OUTPUT," & vbNewLine & _
    String(iIndent, vbTab) & "@psResult varchar(MAX) OUTPUT," & vbNewLine & _
    String(iIndent, vbTab) & "@pfResult bit OUTPUT," & vbNewLine & _
    String(iIndent, vbTab) & "@pdtResult datetime OUTPUT," & vbNewLine & _
    String(iIndent, vbTab) & "@pfltResult float OUTPUT," & vbNewLine & _
    String(iIndent, vbTab) & "@piTempElement int" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    String(iIndent, vbTab) & "SET @psResult = '';" & vbNewLine & _
    String(iIndent, vbTab) & "SET @pfResult = 0;" & vbNewLine & _
    String(iIndent, vbTab) & "SET @pfltResult = 0;" & vbNewLine & vbNewLine & _
    objStoredProc.ToString & _
    "END"
  gADOCon.Execute strSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_WorkflowCalculation = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Workflow calculations stored procedure"
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function CreateSP_WorkflowWebFormValidation() As Boolean

  Const strSPName As String = "spASRSysWorkflowWebFormValidation"

  Dim strSPSQL As String
  Dim strValidationSQL As String
  Dim strWebFormSQL As String
  Dim iIndent As Integer
  Dim fOK As Boolean
  Dim iValueType As ExpressionValueTypes
  Dim iSQLDataType As SQLDataType
  Dim strIdentifier As String
  
  On Error GoTo ErrorTrap

  fOK = True
  iIndent = 1

  DropProcedure strSPName

  strValidationSQL = vbNullString
  strSPSQL = vbNullString
  
  With recWorkflowEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      ' JPD 2010/03/18 Jira HRPRO-821
      If !Enabled Or WorkflowsWithStatus(recWorkflowEdit!ID, giWFSTATUS_INPROGRESS) Then

        With recWorkflowElementEdit
          .Index = "idxWorkflowID"
          .Seek ">=", recWorkflowEdit!ID
      
          If Not .NoMatch Then
            Do While Not .EOF
              'If no more elements for this workflow exit loop
              If !WorkflowID <> recWorkflowEdit!ID Then
                Exit Do
              End If
      
              If recWorkflowElementEdit!Type = elem_WebForm Then
                strWebFormSQL = vbNullString
                
                ' Form validations
                recWorkflowElementValidationEdit.Index = "idxElementID"
                recWorkflowElementValidationEdit.Seek ">=", recWorkflowElementEdit!ID

                If Not recWorkflowElementValidationEdit.NoMatch Then
                  Do While Not recWorkflowElementValidationEdit.EOF
                    'If no more Validations for this element exit loop
                    If recWorkflowElementValidationEdit!elementid <> recWorkflowElementEdit!ID Then
                      Exit Do
                    End If

                    strWebFormSQL = strWebFormSQL & _
                      String(3, vbTab) & "SET @fResult= 0" & vbNewLine & vbNewLine & _
                      String(3, vbTab) & "EXEC [dbo].[spASRSysWorkflowCalculation]" & vbNewLine & _
                      String(4, vbTab) & "@piInstanceID," & vbNewLine & _
                      String(4, vbTab) & CStr(recWorkflowElementValidationEdit!ExprID) & "," & vbNewLine & _
                      String(4, vbTab) & "@iResultType OUTPUT," & vbNewLine & _
                      String(4, vbTab) & "@sResult OUTPUT," & vbNewLine & _
                      String(4, vbTab) & "@fResult OUTPUT," & vbNewLine & _
                      String(4, vbTab) & "@dtResult OUTPUT," & vbNewLine & _
                      String(4, vbTab) & "@fltResult OUTPUT," & vbNewLine & _
                      String(4, vbTab) & CStr(recWorkflowElementEdit!ID) & vbNewLine & vbNewLine
                    
                    strWebFormSQL = strWebFormSQL & _
                      String(3, vbTab) & "IF @fResult = 0" & vbNewLine & _
                      String(3, vbTab) & "BEGIN" & vbNewLine & _
                      String(4, vbTab) & "INSERT INTO @messages" & vbNewLine & _
                      String(5, vbTab) & "([message], [failureType])" & vbNewLine & _
                      String(4, vbTab) & "VALUES ('" & Replace(recWorkflowElementValidationEdit!Message, "'", "''") & "', " & CStr(recWorkflowElementValidationEdit!Type) & ")" & vbNewLine & _
                      String(3, vbTab) & "END" & vbNewLine & vbNewLine

                    'Get next element Validation definition
                    recWorkflowElementValidationEdit.MoveNext
                  Loop
                End If
                
                ' Mandatory checks
                recWorkflowElementItemEdit.Index = "idxElementID"
                recWorkflowElementItemEdit.Seek ">=", recWorkflowElementEdit!ID
                
                If Not recWorkflowElementItemEdit.NoMatch Then
                  Do While Not recWorkflowElementItemEdit.EOF
                    'If no more items for this element exit loop
                    If recWorkflowElementItemEdit!elementid <> recWorkflowElementEdit!ID Then
                      Exit Do
                    End If
          
                    If recWorkflowElementItemEdit!Mandatory Then
                      strWebFormSQL = strWebFormSQL & _
                        String(3, vbTab) & "SELECT @sValue = ltrim(rtrim(isnull(tempValue, '')))" & vbNewLine & _
                        String(3, vbTab) & "FROM ASRSysWorkflowInstanceValues" & vbNewLine & _
                        String(3, vbTab) & "WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID" & vbNewLine & _
                        String(4, vbTab) & "AND ASRSysWorkflowInstanceValues.elementID = @piElementID" & vbNewLine & _
                        String(4, vbTab) & "AND ASRSysWorkflowInstanceValues.identifier = '" & Replace(recWorkflowElementItemEdit!Identifier, "'", "''") & "'" & vbNewLine & vbNewLine

                      iValueType = giEXPRVALUE_UNDEFINED

                      Select Case recWorkflowElementItemEdit!ItemType
                        Case giWFFORMITEM_INPUTVALUE_CHAR, _
                          giWFFORMITEM_INPUTVALUE_DROPDOWN

                          iValueType = giEXPRVALUE_CHARACTER
                        
                        Case giWFFORMITEM_INPUTVALUE_NUMERIC, _
                          giWFFORMITEM_INPUTVALUE_GRID

                          iValueType = giEXPRVALUE_NUMERIC
                        
                        Case giWFFORMITEM_INPUTVALUE_LOOKUP
                          iSQLDataType = GetColumnDataType(recWorkflowElementItemEdit!LookupColumnID)
                          
                          Select Case iSQLDataType
                            Case dtVARCHAR, dtLONGVARCHAR
                              iValueType = giEXPRVALUE_CHARACTER
                            Case dtTIMESTAMP
                              iValueType = giEXPRVALUE_DATE
                            Case dtINTEGER, dtNUMERIC
                              iValueType = giEXPRVALUE_NUMERIC
                          End Select

                        Case giWFFORMITEM_INPUTVALUE_DATE
                          iValueType = giEXPRVALUE_DATE

                        Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
                          iValueType = giEXPRVALUE_NUMERIC
                      End Select
                      
                      strIdentifier = Replace(recWorkflowElementItemEdit!Identifier, "'", "''")
                      
                      Select Case iValueType
                        Case giEXPRVALUE_CHARACTER
                          strWebFormSQL = strWebFormSQL & _
                            String(3, vbTab) & "IF len(@sValue) = 0" & vbNewLine & _
                            String(3, vbTab) & "BEGIN" & vbNewLine & _
                            String(4, vbTab) & "INSERT INTO @messages" & vbNewLine & _
                            String(5, vbTab) & "([message], [failureType])" & vbNewLine & _
                            String(4, vbTab) & "VALUES ('The ''" & strIdentifier & "'' item is mandatory.', 0)" & vbNewLine & _
                            String(3, vbTab) & "END" & vbNewLine & vbNewLine
                        
                        Case giEXPRVALUE_NUMERIC
                          strWebFormSQL = strWebFormSQL & _
                            String(3, vbTab) & "IF isnumeric(@sValue) = 1" & vbNewLine & _
                            String(3, vbTab) & "BEGIN" & vbNewLine & _
                            String(4, vbTab) & "IF convert(float, @sValue) = 0" & vbNewLine & _
                            String(4, vbTab) & "BEGIN" & vbNewLine & _
                            String(5, vbTab) & "INSERT INTO @messages" & vbNewLine & _
                            String(6, vbTab) & "([message], [failureType])" & vbNewLine & _
                            String(5, vbTab) & "VALUES ('The ''" & strIdentifier & "'' item is mandatory.', 0)" & vbNewLine & _
                            String(4, vbTab) & "END" & vbNewLine & _
                            String(3, vbTab) & "END" & vbNewLine & _
                            String(3, vbTab) & "ELSE" & vbNewLine & _
                            String(3, vbTab) & "BEGIN" & vbNewLine & _
                            String(4, vbTab) & "INSERT INTO @messages" & vbNewLine & _
                            String(5, vbTab) & "([message], [failureType])" & vbNewLine & _
                            String(4, vbTab) & "VALUES ('The ''" & strIdentifier & "'' item is mandatory.', 0)" & vbNewLine & _
                            String(3, vbTab) & "END" & vbNewLine & vbNewLine

                        Case giEXPRVALUE_DATE
                          strWebFormSQL = strWebFormSQL & _
                            String(3, vbTab) & "IF UPPER(@sValue) = 'NULL'" & vbNewLine & _
                            String(4, vbTab) & "OR len(@sValue) = 0" & vbNewLine & _
                            String(3, vbTab) & "BEGIN" & vbNewLine & _
                            String(4, vbTab) & "INSERT INTO @messages" & vbNewLine & _
                            String(5, vbTab) & "([message], [failureType])" & vbNewLine & _
                            String(4, vbTab) & "VALUES ('The ''" & strIdentifier & "'' item is mandatory.', 0)" & vbNewLine & _
                            String(3, vbTab) & "END" & vbNewLine & vbNewLine
                      End Select
                    End If
                 
                    'Get next element item definition
                    recWorkflowElementItemEdit.MoveNext
                  Loop
                End If
              
                If Len(strWebFormSQL) > 0 Then
                  strValidationSQL = strValidationSQL & _
                    String(2, vbTab) & "------------------------------------------------------" & vbNewLine & _
                    String(2, vbTab) & "-- Workflow '" & recWorkflowEdit!Name & "', Web Form '" & recWorkflowElementEdit!Identifier & "'" & vbNewLine & _
                    String(2, vbTab) & "------------------------------------------------------" & vbNewLine & _
                    String(2, vbTab) & "IF @piElementID = " & CStr(recWorkflowElementEdit!ID) & vbNewLine & _
                    String(2, vbTab) & "BEGIN" & vbNewLine & _
                    strWebFormSQL & vbNewLine & _
                    String(2, vbTab) & "END" & vbNewLine
                End If
              End If
        
              'Get next element definition
              .MoveNext
            Loop
          End If
        End With
      End If
      
      .MoveNext
    Loop
  End With

  iIndent = 1
  strSPSQL = _
    "------------------------------------------------------" & vbNewLine & _
    "-- Workflow Web Form validation stored procedure." & vbNewLine & _
    "-- Automatically generated by the System Manager." & vbNewLine & _
    "------------------------------------------------------" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & strSPName & "]" & vbNewLine & _
    "(" & vbNewLine & _
    String(iIndent, vbTab) & "@piInstanceID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@piElementID integer," & vbNewLine & _
    String(iIndent, vbTab) & "@psFormInput1 varchar(MAX)" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  strSPSQL = strSPSQL & _
    String(iIndent, vbTab) & "DECLARE @sValue   varchar(MAX)," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iIndex1    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iIndex2    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iTemp    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iTableID    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iParent1TableID    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iParent1RecordID    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iParent2TableID    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iParent2RecordID    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iItemType    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iBehaviour   integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@sIdentifier    varchar(MAX)," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@fSubmitted    bit," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@iResultType    integer," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@sResult    varchar(MAX)," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@fResult    bit," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@dtResult    datetime," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@fltResult    float," & vbNewLine & _
    String(iIndent + 1, vbTab) & "@sID      varchar(MAX)" & vbNewLine & vbNewLine & _
    String(iIndent, vbTab) & "SET @fSubmitted = 0" & vbNewLine & vbNewLine & _
    String(iIndent, vbTab) & "DECLARE @messages TABLE (" & vbNewLine & _
    String(iIndent + 1, vbTab) & "[message] varchar(MAX) COLLATE database_default," & vbNewLine & _
    String(iIndent + 1, vbTab) & "[failureType] integer" & vbNewLine & _
    String(iIndent, vbTab) & ")" & vbNewLine & vbNewLine

  strSPSQL = strSPSQL & _
    String(iIndent, vbTab) & "-- Put the submitted form values into the ASRSysWorkflowInstanceValues table." & vbNewLine & _
    String(iIndent, vbTab) & "WHILE (charindex(CHAR(9), @psFormInput1) > 0)" & vbNewLine & _
    String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @iIndex1 = charindex(CHAR(9), @psFormInput1);" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1);" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''', '''''');" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1);" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2);" & vbNewLine & vbNewLine & _
    String(iIndent + 1, vbTab) & "-- Get the WebForm item type, etc." & vbNewLine & _
    String(iIndent + 1, vbTab) & "SELECT @sIdentifier = EI.identifier," & vbNewLine & _
    String(iIndent + 2, vbTab) & "@iItemType = EI.itemType," & vbNewLine & _
    String(iIndent + 2, vbTab) & "@iTableID = EI.tableID," & vbNewLine & _
    String(iIndent + 2, vbTab) & "@iBehaviour = EI.behaviour" & vbNewLine & _
    String(iIndent + 1, vbTab) & "FROM ASRSysWorkflowElementItems EI" & vbNewLine & _
    String(iIndent + 1, vbTab) & "WHERE EI.ID = convert(integer, @sID)" & vbNewLine & vbNewLine
   
  strSPSQL = strSPSQL & _
    String(iIndent + 1, vbTab) & "SET @iParent1TableID = 0" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @iParent1RecordID = 0" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @iParent2TableID = 0" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET @iParent2RecordID = 0" & vbNewLine & vbNewLine & _
    String(iIndent + 1, vbTab) & "IF @iItemType = 11 -- Record Selector" & vbNewLine & _
    String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine & _
    String(iIndent + 2, vbTab) & "SET @iTemp = convert(integer, isnull(@sValue, '0'))" & vbNewLine & vbNewLine & _
    String(iIndent + 2, vbTab) & "-- Record the selected record's parent details." & vbNewLine & _
    String(iIndent + 2, vbTab) & "exec [dbo].[spASRGetParentDetails]" & vbNewLine & _
    String(iIndent + 3, vbTab) & "@iTableID," & vbNewLine & _
    String(iIndent + 3, vbTab) & "@iTemp," & vbNewLine & _
    String(iIndent + 3, vbTab) & "@iParent1TableID  OUTPUT," & vbNewLine & _
    String(iIndent + 3, vbTab) & "@iParent1RecordID OUTPUT," & vbNewLine & _
    String(iIndent + 3, vbTab) & "@iParent2TableID  OUTPUT," & vbNewLine & _
    String(iIndent + 3, vbTab) & "@iParent2RecordID OUTPUT" & vbNewLine & _
    String(iIndent + 1, vbTab) & "END" & vbNewLine

  strSPSQL = strSPSQL & _
    String(iIndent + 1, vbTab) & "ELSE" & vbNewLine & _
    String(iIndent + 1, vbTab) & "IF (@iItemType = 0) and (@iBehaviour = 0) AND (@sValue = '1')-- Submit Button pressed" & vbNewLine & _
    String(iIndent + 1, vbTab) & "BEGIN" & vbNewLine & _
    String(iIndent + 2, vbTab) & "SET @fSubmitted = 1" & vbNewLine & _
    String(iIndent + 1, vbTab) & "END" & vbNewLine & vbNewLine

  strSPSQL = strSPSQL & _
    String(iIndent + 1, vbTab) & "UPDATE ASRSysWorkflowInstanceValues" & vbNewLine & _
    String(iIndent + 1, vbTab) & "SET ASRSysWorkflowInstanceValues.tempValue = @sValue," & vbNewLine & _
    String(iIndent + 2, vbTab) & "ASRSysWorkflowInstanceValues.tempParent1TableID = @iParent1TableID," & vbNewLine & _
    String(iIndent + 2, vbTab) & "ASRSysWorkflowInstanceValues.tempParent1RecordID = @iParent1RecordID," & vbNewLine & _
    String(iIndent + 2, vbTab) & "ASRSysWorkflowInstanceValues.tempParent2TableID = @iParent2TableID," & vbNewLine & _
    String(iIndent + 2, vbTab) & "ASRSysWorkflowInstanceValues.tempParent2RecordID = @iParent2RecordID" & vbNewLine & _
    String(iIndent + 1, vbTab) & "WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID" & vbNewLine & _
    String(iIndent + 2, vbTab) & "AND ASRSysWorkflowInstanceValues.elementID = @piElementID" & vbNewLine & _
    String(iIndent + 2, vbTab) & "AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier" & vbNewLine & _
    String(iIndent, vbTab) & "END" & vbNewLine & vbNewLine

  If Len(strValidationSQL) > 0 Then
    strSPSQL = strSPSQL & _
      String(iIndent, vbTab) & "IF @fSubmitted = 1" & vbNewLine & _
      String(iIndent, vbTab) & "BEGIN" & vbNewLine & _
      strValidationSQL & _
      String(iIndent, vbTab) & "END" & vbNewLine & vbNewLine
  End If
  
  strSPSQL = strSPSQL & _
    String(iIndent, vbTab) & "SELECT DISTINCT [message]," & vbNewLine & _
    String(iIndent + 1, vbTab) & "[failureType]" & vbNewLine & _
    String(iIndent, vbTab) & "FROM @messages" & vbNewLine & _
    String(iIndent, vbTab) & "ORDER BY [failureType], [message]" & vbNewLine & _
    "END"
  gADOCon.Execute strSPSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_WorkflowWebFormValidation = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Workflow Web Form validation stored procedure"
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function CreateSP_WorkflowGetValidLoginsForStep() As Boolean

  Dim strSPName As String
  Dim strSPSQL As String
  Dim fOK As Boolean
  Dim lngModuleRefId As Long
  Dim bRequireAuthorization As Boolean
  
  Dim sPersonnelTableName As String
  Dim sWorkEmailColumnName As String
  Dim sSelfServiceLoginColumnName As String
  Dim bSetupOK As Boolean
 
  On Error GoTo ErrorTrap
  
  strSPName = "spASRWorkflowGetValidLoginsForStep"

  fOK = DropProcedure(strSPName)

  If fOK Then
    lngModuleRefId = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
    sPersonnelTableName = GetTableName(lngModuleRefId)
       
    lngModuleRefId = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME, 0)
    sSelfServiceLoginColumnName = GetColumnName(lngModuleRefId, True)
              
    lngModuleRefId = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKEMAIL, 0)
    sWorkEmailColumnName = GetColumnName(lngModuleRefId, True)
  
    bRequireAuthorization = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_REQUIRESAUTHORIZATION, 0)
  
    bSetupOK = Len(sPersonnelTableName) > 0 And Len(sSelfServiceLoginColumnName) > 0 And Len(sWorkEmailColumnName)
   
    strSPSQL = _
      "------------------------------------------------------" & vbNewLine & _
      "-- Workflow Web Form authentication stored procedure." & vbNewLine & _
      "-- Automatically generated by the System Manager." & vbNewLine & _
      "------------------------------------------------------" & vbNewLine & _
      "CREATE PROCEDURE [dbo].[" & strSPName & "]" & vbNewLine & _
      "(" & vbNewLine & _
      "    @instanceId integer," & vbNewLine & _
      "    @elementId integer," & vbNewLine & _
      "    @requiresAuthorization bit = 0 OUTPUT)" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & _
      "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
      "    DECLARE @Email nvarchar(MAX)," & vbNewLine & _
      "            @initiatorID int;" & vbNewLine & vbNewLine
      
    If bSetupOK Then
      strSPSQL = strSPSQL & _
        "    SELECT @initiatorID = ISNULL(i.InitiatorID, 0), @Email = s.UserEmail + ';' + ISNULL(s.EmailCC, ''), @requiresAuthorization = ISNULL(e.RequiresAuthentication, 0)" & vbNewLine & _
        "        FROM ASRSysWorkflowInstanceSteps s" & vbNewLine & _
        "        INNER JOIN ASRSysWorkflowElements e ON e.ID = s.ElementID" & vbNewLine & _
        "        INNER JOIN ASRSysWorkflowInstances i ON s.InstanceID = i.ID" & vbNewLine & _
        "        WHERE s.elementid = @elementId AND s.instanceId = @instanceID;" & vbNewLine & vbNewLine
        
      If bRequireAuthorization Then
        strSPSQL = strSPSQL & _
          "    IF LEN(@Email) > 1 SET @requiresAuthorization = 1;" & vbNewLine & vbNewLine
      End If

      strSPSQL = strSPSQL & _
        "    IF @Email = ';'" & vbNewLine & _
        "        SELECT [" & sWorkEmailColumnName & "] AS [Email], ISNULL([" & sSelfServiceLoginColumnName & "], '') AS [Login]" & vbNewLine & _
        "            FROM [" & sPersonnelTableName & "]" & vbNewLine & _
        "            WHERE [id] = @initiatorID;" & vbNewLine & vbNewLine & _
        "    ELSE" & vbNewLine & _
        "       SELECT l.SplitColumn AS [Email], p.[" & sSelfServiceLoginColumnName & "] AS [Login]" & vbNewLine & _
        "           FROM dbo.[udfsysStringToTable](@Email, ';') l" & vbNewLine & _
        "           INNER JOIN [" & sPersonnelTableName & "] p ON p.[" & sWorkEmailColumnName & "] = l.SplitColumn" & vbNewLine & _
        "           WHERE l.SplitColumn <> '' AND p.[" & sSelfServiceLoginColumnName & "] <> '';" & vbNewLine & vbNewLine
    End If
    
    strSPSQL = strSPSQL & _
      " END"
    
     gADOCon.Execute strSPSQL, , adExecuteNoRecords
    
  End If

TidyUpAndExit:
  CreateSP_WorkflowGetValidLoginsForStep = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Workflow Web Form valid logins for step stored procedure"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function GetSQLFilter(lngFilterID As Long, sCurrentTable As String) As String
  Dim fOK As Boolean
  Dim objExpr As CExpression
  Dim strFilterRunTimeCode As String

  GetSQLFilter = vbNullString

  'Filter
  Set objExpr = New CExpression
  With objExpr
    .ExpressionID = lngFilterID
    .ConstructExpression
    fOK = .RuntimeFilterCode(strFilterRunTimeCode, False)

    strFilterRunTimeCode = Replace(strFilterRunTimeCode, vbNewLine, " ")
      
    GetSQLFilter = "@recordID IN " & _
          "(" & strFilterRunTimeCode & ")"
  End With
  Set objExpr = Nothing

End Function

Public Function WorkflowLinkTriggerCode_Insert() As String
  WorkflowLinkTriggerCode_Insert = msInsertLinkCode
End Function

Public Function WorkflowLinkTriggerCode_Update() As String
  WorkflowLinkTriggerCode_Update = msUpdateLinkCode
End Function

Public Function IsWorkflowControl(ctrl As VB.Control)

  IsWorkflowControl = _
       (TypeOf ctrl Is COAWF_Link) _
    Or (TypeOf ctrl Is COAWF_Decision) _
    Or (TypeOf ctrl Is COAWF_Webform) _
    Or (TypeOf ctrl Is COAWF_Email) _
    Or (TypeOf ctrl Is COAWF_StoredData) _
    Or (TypeOf ctrl Is COAWF_BeginEnd) _
    Or (TypeOf ctrl Is COAWF_Junction)

End Function

Public Function IsWorkflowElement(ctrl As VB.Control)

  IsWorkflowElement = _
       (TypeOf ctrl Is COAWF_Decision) _
    Or (TypeOf ctrl Is COAWF_Webform) _
    Or (TypeOf ctrl Is COAWF_Email) _
    Or (TypeOf ctrl Is COAWF_StoredData) _
    Or (TypeOf ctrl Is COAWF_BeginEnd) _
    Or (TypeOf ctrl Is COAWF_Junction)

End Function

Public Function CreateSP_WorkflowSelfServiceRecord() As Boolean

  Const strSPName As String = "spASRWorkflowGetSelfServiceRecordID"

  On Error GoTo ErrorTrap

  Dim sPersonnelTableName As String
  Dim sSelfServiceLoginColumnName As String
  Dim lngModuleRefId As Long
  Dim bOK As Boolean
  Dim strSPSQL As String

  bOK = DropProcedure(strSPName)

  If bOK Then
    lngModuleRefId = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
    sPersonnelTableName = GetTableName(lngModuleRefId)
       
    lngModuleRefId = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME, 0)
    sSelfServiceLoginColumnName = GetColumnName(lngModuleRefId, True)
      
    If Len(sPersonnelTableName) > 0 And Len(sSelfServiceLoginColumnName) > 0 Then
      strSPSQL = _
       "------------------------------------------------------" & vbNewLine & _
       "-- Workflow Identification stored procedure." & vbNewLine & _
       "-- Automatically generated by the System Manager." & vbNewLine & _
       "------------------------------------------------------" & vbNewLine & _
       "CREATE PROCEDURE [dbo].[" & strSPName & "]" & vbNewLine & _
       "(" & vbNewLine & _
       "    @psUserName nvarchar(255)," & vbNewLine & _
       "    @piRecordID integer OUTPUT," & vbNewLine & _
       "    @piRecordCount integer OUTPUT)" & vbNewLine & _
       "AS" & vbNewLine & _
       "BEGIN" & vbNewLine & _
       "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
       "    SET @piRecordID = 0;" & vbNewLine & _
       "    SELECT @piRecordCount = COUNT(ID) FROM [" & sPersonnelTableName & "] WHERE [" & sSelfServiceLoginColumnName & "] = @psUserName;" & vbNewLine & _
       "    IF @piRecordCount = 1" & vbNewLine & _
       "       SELECT @piRecordID = [ID] FROM [" & sPersonnelTableName & "] WHERE [" & sSelfServiceLoginColumnName & "] = @psUserName;" & vbNewLine & _
       "END"
 
      gADOCon.Execute strSPSQL, , adExecuteNoRecords
    End If
  End If

TidyUpAndExit:
  CreateSP_WorkflowSelfServiceRecord = bOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Workflow Self Service Record stored procedure"
  bOK = False
  Resume TidyUpAndExit
 
End Function

Private Function CreateSP_WorkspaceCheckPendingSteps() As Boolean
  ' Create the Intranet Check Pending Steps stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Workflow module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msWorkspaceCheckPendingSteps_PROCEDURENAME & "]" & vbNewLine & _
    "    (@username nvarchar(255))" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    DECLARE" & vbNewLine & _
    "        @sURL varchar(MAX)," & vbNewLine & _
    "        @sDescription varchar(MAX)," & vbNewLine & _
    "        @sCalcDescription varchar(MAX)," & vbNewLine & _
    "        @iInstanceID integer," & vbNewLine & _
    "        @iInstanceStepID integer," & vbNewLine & _
    "        @iElementID integer," & vbNewLine & _
    "        @hResult integer," & vbNewLine & _
    "        @objectToken integer," & vbNewLine & _
    "        @sQueryString varchar(MAX)," & vbNewLine & _
    "        @sParam1  varchar(MAX)," & vbNewLine & _
    "        @sServerName sysname," & vbNewLine & _
    "        @sDBName  sysname," & vbNewLine & _
    "        @sWorkflowName varchar(MAX);" & vbNewLine

  sProcSQL = sProcSQL & vbNewLine & _
    "    DECLARE @pass1 TABLE(" & vbNewLine & _
    "        [instanceID]  integer," & vbNewLine & _
    "        [elementID]   integer," & vbNewLine & _
    "        [stepID]      integer," & vbNewLine & _
    "        [name]        varchar(MAX)," & vbNewLine & _
    "        [description] varchar(MAX)," & vbNewLine & _
    "        [url]         nvarchar(MAX));" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "    DECLARE @steps TABLE (" & vbNewLine & _
    "        [name]           varchar(MAX)," & vbNewLine & _
    "        [description]    varchar(MAX)," & vbNewLine & _
    "        [URL]            varchar(MAX)," & vbNewLine & _
    "        [instanceID]     integer," & vbNewLine & _
    "        [elementID]      integer," & vbNewLine & _
    "        [instanceStepID] integer);" & vbNewLine & vbNewLine
      
  sProcSQL = sProcSQL & _
    "    SELECT @sURL = parameterValue" & vbNewLine & _
    "    FROM ASRSysModuleSetup" & vbNewLine & _
    "    WHERE moduleKey = 'MODULE_WORKFLOW' AND parameterKey = 'Param_URL';" & vbNewLine & vbNewLine & _
    "        IF upper(right(@sURL, 5)) <> '.ASPX'" & vbNewLine & _
    "            AND right(@sURL, 1) <> '/'" & vbNewLine & _
    "            AND len(@sURL) > 0" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            SET @sURL = @sURL + '/'" & vbNewLine & _
    "        END"
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "    SELECT @sParam1 = parameterValue" & vbNewLine & _
    "    FROM ASRSysModuleSetup" & vbNewLine & _
    "    WHERE moduleKey = 'MODULE_WORKFLOW' AND parameterKey = 'Param_Web1';" & vbNewLine & vbNewLine & _
    "    SET @sServerName = CONVERT(sysname,SERVERPROPERTY('servername'));" & vbNewLine & _
    "    SET @sDBName = DB_NAME();"
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "    IF (len(@sURL) > 0)" & vbNewLine & _
    "    BEGIN"
    
  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
        "        DECLARE @sEmailAddress_" & CStr(iCount) & " varchar(MAX)" & vbNewLine & _
        "        SELECT @sEmailAddress_" & CStr(iCount) & " = replace(upper(ltrim(rtrim(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & "))), ' ', '')" & vbNewLine & _
        "        FROM " & mvar_sLoginTable & vbNewLine & _
        "        WHERE (ISNULL(" & mvar_sLoginTable & "." & mvar_sLoginColumn & ", '') = @username" & _
        IIf(Len(mvar_sSecondLoginColumn) > 0, vbNewLine & "            OR ISNULL(" & mvar_sLoginTable & "." & mvar_sSecondLoginColumn & ", '') = @username", "") & ")" & vbNewLine & _
        "            AND LEN(" & mvar_sLoginTable & "." & GetColumnName(malngEmailColumns(iCount), True) & ") > 0"
    Next iCount
  End If
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "        INSERT @pass1" & vbNewLine & _
    "            SELECT ASRSysWorkflowInstanceSteps.instanceID," & vbNewLine & _
    "                ASRSysWorkflowInstanceSteps.elementID," & vbNewLine & _
    "                ASRSysWorkflowInstanceSteps.ID," & vbNewLine & _
    "                ASRSysWorkflows.name + ' - ' + ASRSysWorkflowElements.caption AS [description], " & vbNewLine & _
    "                ASRSysWorkflows.name, " & vbNewLine & _
    "                dbo.[udfASRNetGetWorkflowQueryString]( ASRSysWorkflowInstanceSteps.instanceID,  ASRSysWorkflowInstanceSteps.elementID, @sParam1, @sServerName, @sDBName)" & vbNewLine & _
    "            FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
    "            INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
    "            INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
    "            WHERE (ASRSysWorkflowInstanceSteps.Status = 2" & vbNewLine & _
    "                    OR ASRSysWorkflowInstanceSteps.Status = 7)" & vbNewLine & _
    "                AND (ASRSysWorkflowInstanceSteps.userName = @username"
    
  If UBound(malngEmailColumns) > 0 Then
    For iCount = 1 To UBound(malngEmailColumns)
      sProcSQL = sProcSQL & vbNewLine & _
        "                    OR (';' + replace(upper(ASRSysWorkflowInstanceSteps.userEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_" & CStr(iCount) & " + ';%'" & vbNewLine & _
        "                        AND len(@sEmailAddress_" & CStr(iCount) & ") > 0)" & vbNewLine & _
        "                    OR ((len(@sEmailAddress_" & CStr(iCount) & ") > 0)" & vbNewLine & _
        "                        AND ((SELECT COUNT(*)" & vbNewLine & _
        "                            FROM ASRSysWorkflowStepDelegation" & vbNewLine & _
        "                            WHERE stepID = ASRSysWorkflowInstanceSteps.ID" & vbNewLine & _
        "                                AND ';' + replace(upper(ASRSysWorkflowStepDelegation.delegateEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_" & CStr(iCount) & " + ';%') > 0))"
    Next iCount
  End If
    
  sProcSQL = sProcSQL & _
    ")"
    
  sProcSQL = sProcSQL & vbNewLine & vbNewLine & _
    "        DECLARE steps_cursor CURSOR LOCAL FAST_FORWARD FOR SELECT * FROM @pass1;" & vbNewLine & _
    "        OPEN steps_cursor;" & vbNewLine & _
    "        FETCH NEXT FROM steps_cursor INTO @iInstanceID, @iElementID, @iInstanceStepID, @sDescription, @sWorkflowName, @sQueryString;" & vbNewLine & _
    "        WHILE (@@fetch_status = 0)" & vbNewLine & _
    "        BEGIN" & vbNewLine & vbNewLine

  sProcSQL = sProcSQL & _
    "            IF LEN(@sQueryString) > 0" & vbNewLine & _
    "            BEGIN" & vbNewLine & _
    "                EXEC [dbo].[spASRWorkflowStepDescription]" & vbNewLine & _
    "                    @iInstanceStepID," & vbNewLine & _
    "                    @sCalcDescription OUTPUT;" & vbNewLine & _
    "                IF LEN(@sCalcDescription) > 0 " & vbNewLine & _
    "                    SET @sDescription = @sCalcDescription;" & vbNewLine & vbNewLine & _
    "                INSERT INTO @steps ([description], [url], [instanceID], [elementID], [instanceStepID], [name])" & vbNewLine & _
    "                    VALUES (@sDescription, @sURL + '/?' + @sQueryString, @iInstanceID, @iElementID, @iInstanceStepID, @sWorkflowName);" & vbNewLine & _
    "            END" & vbNewLine & _
    "            FETCH NEXT FROM steps_cursor INTO @iInstanceID, @iElementID, @iInstanceStepID, @sDescription, @sWorkflowName, @sQueryString;" & vbNewLine & _
    "        END" & vbNewLine & _
    "        CLOSE steps_cursor;" & vbNewLine & _
    "        DEALLOCATE steps_cursor;" & vbNewLine & vbNewLine & _
    "    END" & vbNewLine & vbNewLine & _
    "    SELECT *" & vbNewLine & _
    "    FROM @steps" & vbNewLine & _
    "    ORDER BY [description];"
  
  sProcSQL = sProcSQL & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_WorkspaceCheckPendingSteps = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Workspace Check Pending Steps stored procedure (Workflow)"
  Resume TidyUpAndExit

End Function

