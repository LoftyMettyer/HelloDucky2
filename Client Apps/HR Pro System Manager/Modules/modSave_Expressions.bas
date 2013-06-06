Attribute VB_Name = "modSave_Expressions"
Option Explicit

Public Function SaveExpressions(pfRefreshDatabase As Boolean) As Boolean
  ' Save the new and modified Expressions to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngExprID As Long
  Dim lngRecordCount As Long
  Dim fSave As Boolean
  
  fOK = True
  
  With recExprEdit
    .Index = "idxExprID"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      lngRecordCount = .RecordCount
    End If
    
    OutputCurrentProcess2 vbNullString, lngRecordCount
    
    Do While fOK And Not .EOF
      lngExprID = .Fields("exprID").value
      
      If !Deleted Then
        OutputCurrentProcess2 .Fields("Name").value
        fOK = ExpressionDelete
      
      ElseIf !New Then
        OutputCurrentProcess2 .Fields("Name").value
        fOK = ExpressionNew

      Else
        fSave = !Changed _
          Or pfRefreshDatabase _
          Or Application.ChangedTableName _
          Or Application.ChangedColumnName
          
        If (Not fSave) _
          And (!Type = giEXPR_WORKFLOWCALCULATION _
            Or !Type = giEXPR_WORKFLOWSTATICFILTER _
            Or !Type = giEXPR_WORKFLOWRUNTIMEFILTER) Then
          
          ' Check if the workflow's changed.
          recWorkflowEdit.Index = "idxWorkflowID"
          recWorkflowEdit.Seek "=", !UtilityID
             
          If Not recWorkflowEdit.NoMatch Then
            fSave = recWorkflowEdit!Changed
          End If

        End If
        
        If fSave Then
          If .Fields("ParentComponentID").value = 0 Then
            OutputCurrentProcess2 .Fields("Name").value
          End If
          fOK = ExpressionSave
        Else
          OutputCurrentProcess2 vbNullString
        End If
      End If
      
      ' Ensure that we are positioned on the correct record
      ' as the recExprEdit recordset may have been repositioned.
      .Index = "idxExprID"
      .Seek ">", lngExprID
      .MoveNext
      fOK = fOK And Not gobjProgress.Cancelled
    
      gobjProgress.UpdateProgress2
          
    Loop
  End With

TidyUpAndExit:
  SaveExpressions = fOK
  Exit Function

ErrorTrap:
  OutputError "Error saving expressions"
  fOK = False
  Resume TidyUpAndExit

End Function


Function ExpressionSave() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = ExpressionDelete
  
  If fOK Then
    fOK = ExpressionNew
  End If

TidyUpAndExit:
  ExpressionSave = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error updating expressions"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ExpressionDelete()
  On Error GoTo ErrorTrap
  
  Dim fDeleteOK As Boolean
  Dim lngExprID As Long
  Dim sProcedureName As String
  lngExprID = recExprEdit!ExprID
  
  gADOCon.Execute "DELETE FROM ASRSysExpressions WHERE exprID = " & Trim$(Str$(lngExprID)), , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysExprComponents WHERE exprID =" & Trim$(Str$(lngExprID)), , adCmdText + adExecuteNoRecords

  ' Drop any existing stored procedure with this name.
  sProcedureName = "sp_ASRExpr_" & lngExprID
  
  DropProcedure sProcedureName

  fDeleteOK = True
  
ExitExprDelete:
  'Set rsExistingProcedures = Nothing
  ExpressionDelete = fDeleteOK
  
  Exit Function

ErrorTrap:
  OutputError "Error deleting expression"
  Err = False
  
  fDeleteOK = False
  
  Resume ExitExprDelete
  
End Function


Private Function ExpressionNew() As Boolean
  On Error GoTo ErrorTrap
  
  Dim iColumn As Integer
  Dim iValidityCode As ExprValidationCodes
  Dim sName As String
  Dim fNotNeeded As Boolean
  Dim fNewExpr As Boolean
  Dim rsExpressions As ADODB.Recordset
  Dim rsComponents As ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim rsTemp As ADODB.Recordset
  Dim objExpression As CExpression
  Dim lngExprID As Long
  Dim lngTableID As Long
  Dim lngUtilityID As Long
  Dim sSQL As String
  Dim iType As Integer
  Dim fOK As Boolean
  Dim fWorkflowEnabled As Boolean
  
  Set rsExpressions = New ADODB.Recordset
  Set rsComponents = New ADODB.Recordset
  Set rsTables = New ADODB.Recordset
  Set rsTemp = New ADODB.Recordset
  
  lngExprID = recExprEdit.Fields("exprID").value
  fNewExpr = True
    
  'Open the expressions table
  rsExpressions.Open "ASRSysExpressions", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
  
  'JDM - 13/06/03 - Fault 5975 - Independant table calcs "disappearing"
  'JPD 20020122 Fault 3375
  If (recExprEdit.Fields("type").value = giEXPR_UTILRUNTIMEFILTER _
    Or (recExprEdit.Fields("type").value = giEXPR_RECORDINDEPENDANTCALC)) Then
    
    fNotNeeded = False
  ElseIf (recExprEdit.Fields("type").value = giEXPR_WORKFLOWCALCULATION) Then
    ' Check that the expression's Workflow still exists.
    lngUtilityID = IIf(IsNull(recExprEdit.Fields("utilityID").value), 0, recExprEdit.Fields("utilityID").value)
    sSQL = "SELECT ID, enabled" & _
      " FROM ASRSysWorkflows" & _
      " WHERE ID = " & Trim$(Str$(lngUtilityID))
    rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    fNotNeeded = (rsTemp.EOF And rsTemp.BOF)
    If Not fNotNeeded Then
      ' JPD 2010/03/18 Jira HRPRO-821
      fWorkflowEnabled = rsTemp!Enabled Or WorkflowsWithStatus(rsTemp!id, giWFSTATUS_INPROGRESS)
    End If
    rsTemp.Close
  Else
    If (recExprEdit.Fields("type").value = giEXPR_WORKFLOWSTATICFILTER) _
      Or (recExprEdit.Fields("type").value = giEXPR_WORKFLOWRUNTIMEFILTER) Then
      ' Check that the expression's Workflow still exists.
      lngUtilityID = IIf(IsNull(recExprEdit.Fields("utilityID").value), 0, recExprEdit.Fields("utilityID").value)
      sSQL = "SELECT ID, enabled" & _
        " FROM ASRSysWorkflows" & _
        " WHERE ID = " & Trim$(Str$(lngUtilityID))
      rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

      fNotNeeded = (rsTemp.EOF And rsTemp.BOF)
      If Not fNotNeeded Then
        ' JPD 2010/03/18 Jira HRPRO-821
        fWorkflowEnabled = rsTemp!Enabled Or WorkflowsWithStatus(rsTemp!id, giWFSTATUS_INPROGRESS)
      End If
      rsTemp.Close
    End If

    If Not fNotNeeded Then
      ' Check that the expression's base table still exists.
      lngTableID = recExprEdit.Fields("TableID").value
      sSQL = "SELECT tableID" & _
        " FROM ASRSysTables" & _
        " WHERE tableID = " & Trim$(Str$(lngTableID))
      rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      
      fNotNeeded = (rsTables.EOF And rsTables.BOF)
      rsTables.Close
    End If
  End If
  
  If Not fNotNeeded Then
    'Add new expression definition
    With rsExpressions
      .AddNew
      
      For iColumn = 0 To .Fields.Count - 1
        sName = UCase$(Trim(.Fields(iColumn).Name))
        If (Not IsNull(recExprEdit.Fields(sName))) And _
          (sName <> "TIMESTAMP") Then
          .Fields(iColumn) = recExprEdit.Fields(sName)
        End If
      Next iColumn
      
      .Update
          
      ' Create the SQL stored procedure for the expression only for
      ' root caclulation type expressions. ie. not the expression that form the
      ' paramters of function components.
      If (recExprEdit.Fields("parentComponentID").value = 0) Then

        iType = recExprEdit.Fields("type").value

        Select Case iType
          Case giEXPR_COLUMNCALCULATION, _
            giEXPR_STATICFILTER, _
            giEXPR_RECORDVALIDATION, _
            giEXPR_RECORDDESCRIPTION, _
            giEXPR_OUTLOOKFOLDER, _
            giEXPR_OUTLOOKSUBJECT, _
            giEXPR_EMAIL

            Set objExpression = New CExpression
            objExpression.ExpressionID = lngExprID
            
            fOK = objExpression.CreateStoredProcedure
            
            If Not fOK Then
              fNewExpr = False
              
              OutputError "Expression : " & objExpression.Name & vbNewLine & vbNewLine & _
                "Error creating stored procedure."
            End If
          
          Case giEXPR_WORKFLOWRUNTIMEFILTER
            Set objExpression = New CExpression
            objExpression.ExpressionID = lngExprID
            
            fOK = True
            If fWorkflowEnabled And ExpressionUsedInWorkflow(lngExprID) Then
              fOK = objExpression.CreateWorkflowUDF
            End If
            
            If Not fOK Then
              fNewExpr = False
              
              OutputError "Expression : " & objExpression.Name & vbNewLine & vbNewLine & _
                "Error creating Workflow UDF."
            End If
          
          Case giEXPR_WORKFLOWCALCULATION, _
            giEXPR_WORKFLOWSTATICFILTER

            Set objExpression = New CExpression
            objExpression.ExpressionID = lngExprID
            
            fOK = True
            If fWorkflowEnabled And ExpressionUsedInWorkflow(lngExprID) Then
              fOK = objExpression.CreateStoredProcedure
            End If
        
            If Not fOK Then
              fNewExpr = False
              
              OutputError "Expression : " & objExpression.Name & vbNewLine & vbNewLine & _
                "Error creating Workflow stored procedure."
            End If
          
          Case giEXPR_VIEWFILTER
            Set objExpression = New CExpression
            objExpression.ExpressionID = lngExprID
            objExpression.ConstructExpression
            
            iValidityCode = objExpression.ValidateExpression(True)
          
            If iValidityCode <> giEXPRVALIDATION_NOERRORS Then
              fNewExpr = False
              
              OutputError "Expression : " & objExpression.Name & vbNewLine & vbNewLine & _
                objExpression.ValidityMessage(iValidityCode)
            End If
        
          Case giEXPR_DEFAULTVALUE
            Set objExpression = New CExpression
            objExpression.ExpressionID = lngExprID
            objExpression.ConstructExpression
            
            If Not objExpression.CreateDefaultValueStoredProcedure Then
              fNewExpr = False
              OutputError "Expression : " & objExpression.Name & vbNewLine & vbNewLine & _
                "Error creating default value stored procedure."
            End If
        End Select
        
        Set objExpression = Nothing
      End If
    End With
    rsExpressions.Close
      
    If fNewExpr Then
      ' Open the components definition table
      rsComponents.Open "ASRSysExprComponents", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
      
      ' Add definitions for the components of this expression.
      recCompEdit.Index = "idxExprID"
      recCompEdit.Seek ">=", lngExprID
      
      If Not recCompEdit.NoMatch Then
        Do While Not recCompEdit.EOF
          'If no more components for this expression exit loop
          If recCompEdit!ExprID <> lngExprID Then
            Exit Do
          End If
          
          'Add component definition
          With rsComponents
            .AddNew
            
            For iColumn = 0 To .Fields.Count - 1
              sName = .Fields(iColumn).Name
              
              If Not IsNull(recCompEdit.Fields(sName).value) Then
                .Fields(iColumn).value = recCompEdit.Fields(sName).value
              End If
            Next iColumn
            
            .Update
          End With
          
          'Get next item definition
          recCompEdit.MoveNext
        Loop
      End If
      
      rsComponents.Close
    End If
  End If
  
ExitNewExpr:
  Set rsTables = Nothing
  Set objExpression = Nothing
  Set rsExpressions = Nothing
  Set rsComponents = Nothing
  
  ExpressionNew = fNewExpr
  
  Exit Function

ErrorTrap:
  'gobjProgress.Visible = False
  OutputError "Error creating expression"
  Err = False
  
  fNewExpr = False
  
  Resume ExitNewExpr
  
End Function
