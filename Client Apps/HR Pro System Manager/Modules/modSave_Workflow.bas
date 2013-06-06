Attribute VB_Name = "modSave_Workflow"
Option Explicit

Public Function SaveWorkflows() As Boolean
  ' Save the new or modified workflows definitions.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  fOK = True

  With recWorkflowEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
        fOK = WorkflowDelete
      ElseIf !New Then
        fOK = WorkflowNew
      ElseIf !Changed Then
        fOK = WorkflowSave
      End If

      .MoveNext
    Loop
  End With

  If fOK Then
    fOK = CreateSP_WorkflowCalculation
  End If
  
  If fOK Then
    fOK = CreateSP_WorkflowParentRecord
  End If
  
  If fOK Then
    fOK = CreateSP_WorkflowWebFormValidation
  End If
  
TidyUpAndExit:
  SaveWorkflows = fOK
  Exit Function

ErrorTrap:
  OutputError "Error saving workflow definitions"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function WorkflowSave() As Boolean
  ' Save the current Workflow record to the server database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  fOK = WorkflowDelete
  If fOK Then
    fOK = WorkflowNew
  End If

TidyUpAndExit:
  WorkflowSave = fOK
  Exit Function

ErrorTrap:
  OutputError "Error updating workflow"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function WorkflowDelete() As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim lngWorkflowID As Long

  lngWorkflowID = recWorkflowEdit!ID

  gADOCon.Execute "DELETE FROM ASRSysWorkflows WHERE ID=" & lngWorkflowID, , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysWorkflowElementItems WHERE elementID IN (SELECT ID FROM ASRSysWorkflowElements WHERE workflowID=" & lngWorkflowID & ")", , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysWorkflowElementItemValues WHERE itemID NOT IN (SELECT ID FROM ASRSysWorkflowElementItems)", , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysWorkflowElementColumns WHERE elementID IN (SELECT ID FROM ASRSysWorkflowElements WHERE workflowID=" & lngWorkflowID & ")", , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysWorkflowElementValidations WHERE elementID IN (SELECT ID FROM ASRSysWorkflowElements WHERE workflowID=" & lngWorkflowID & ")", , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysWorkflowElements WHERE workflowID=" & lngWorkflowID, , adCmdText + adExecuteNoRecords
  gADOCon.Execute "DELETE FROM ASRSysWorkflowLinks WHERE workflowID=" & lngWorkflowID, , adCmdText + adExecuteNoRecords
  
  If recWorkflowEdit!Deleted Then
    ' NB. Deleting the ASRSysWorkflowInstances record will trigger the deletion of related
    ' records in ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues
    gADOCon.Execute "DELETE FROM ASRSysWorkflowInstances" & _
      " WHERE workflowID=" & lngWorkflowID, , _
      adCmdText + adExecuteNoRecords
  
    ' NB. Deleting the ASRSysWorkflowTriggeredLinks record will NO LONGER trigger the deletion of related
    ' records in ASRSysWorkflowQueue and ASRSysWorkflowQueueColumns, so we need to do it manually here.
    gADOCon.Execute "DELETE FROM ASRSysWorkflowQueue" & _
      " WHERE linkID IN (SELECT linkID FROM ASRSysWorkflowTriggeredLinks WHERE workflowID = " & lngWorkflowID & ")", , _
      adCmdText + adExecuteNoRecords
    
    gADOCon.Execute "DELETE FROM ASRSysWorkflowTriggeredLinks" & _
      " WHERE workflowID=" & lngWorkflowID, , _
      adCmdText + adExecuteNoRecords
  Else
    ' NB. Deleting the ASRSysWorkflowInstances record will trigger the deletion of related
    ' records in ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues
    gADOCon.Execute "DELETE FROM ASRSysWorkflowInstances" & _
      " WHERE workflowID=" & lngWorkflowID & _
      "   AND status <> " & CStr(giWFSTATUS_INPROGRESS), , _
      adCmdText + adExecuteNoRecords
  End If
  
  fOK = True

TidyUpAndExit:
  WorkflowDelete = fOK
  Exit Function

ErrorTrap:
  OutputError "Error deleting workflow"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function WorkflowNew() As Boolean
  ' Save the current workflow definition to the server database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim sName As String
  Dim rsWorkflows As New ADODB.Recordset
  Dim rsElements As New ADODB.Recordset
  Dim rsElementItems As New ADODB.Recordset
  Dim rsElementItemValues As New ADODB.Recordset
  Dim rsElementColumns As New ADODB.Recordset
  Dim rsElementValidations As New ADODB.Recordset
  Dim rsLinks As New ADODB.Recordset

  fOK = True

  ' Open the Workflows table on the server.
  rsWorkflows.Open "ASRSysWorkflows", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
    
  With rsWorkflows
    .AddNew
    For iLoop = 0 To .Fields.Count - 1
      sName = .Fields(iLoop).Name
      If Not IsNull(recWorkflowEdit.Fields(sName).Value) Then
        .Fields(iLoop).Value = recWorkflowEdit.Fields(sName).Value
      End If
    Next iLoop
    .Update
    .Close
  End With


  ' Open the Workflow Elements table on the server.
  rsElements.Open "ASRSysWorkflowElements", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

  With recWorkflowElementEdit
    .Index = "idxWorkflowID"
    .Seek ">=", recWorkflowEdit!ID

    If Not .NoMatch Then
      Do While Not .EOF
        'If no more elements for this workflow exit loop
        If !WorkflowID <> recWorkflowEdit!ID Then
          Exit Do
        End If

        'Add element details to element table
        rsElements.AddNew
        For iLoop = 0 To rsElements.Fields.Count - 1
          sName = rsElements.Fields(iLoop).Name
          If Not IsNull(.Fields(sName).Value) Then
            rsElements.Fields(iLoop).Value = .Fields(sName).Value
          End If
        Next iLoop
        rsElements.Update

        ' Add the Element Items (if required)
        rsElementItems.Open "ASRSysWorkflowElementItems", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

        recWorkflowElementItemEdit.Index = "idxElementID"
        recWorkflowElementItemEdit.Seek ">=", recWorkflowElementEdit!ID

        If Not recWorkflowElementItemEdit.NoMatch Then
          Do While Not recWorkflowElementItemEdit.EOF
            'If no more items for this element exit loop
            If recWorkflowElementItemEdit!elementID <> recWorkflowElementEdit!ID Then
              Exit Do
            End If

            'Add element item details to element item table
            rsElementItems.AddNew
            For iLoop = 0 To rsElementItems.Fields.Count - 1
              sName = rsElementItems.Fields(iLoop).Name
              If Not IsNull(recWorkflowElementItemEdit.Fields(sName).Value) Then
                rsElementItems.Fields(iLoop).Value = recWorkflowElementItemEdit.Fields(sName).Value
              End If
             Next iLoop
            rsElementItems.Update

            
             ' Add the Element Item Control Values (if required)
            rsElementItemValues.Open "ASRSysWorkflowElementItemValues", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
            
            recWorkflowElementItemValuesEdit.Index = "idxItemID"
            recWorkflowElementItemValuesEdit.Seek ">=", recWorkflowElementItemEdit!ID
            
            If Not recWorkflowElementItemValuesEdit.NoMatch Then
              Do While Not recWorkflowElementItemValuesEdit.EOF
                'If no more item values for this element item exit loop
                If recWorkflowElementItemValuesEdit!itemID <> recWorkflowElementItemEdit!ID Then
                  Exit Do
                End If
            
                'Add item value details to element item value table
                rsElementItemValues.AddNew
                For iLoop = 0 To rsElementItemValues.Fields.Count - 1
                  sName = rsElementItemValues.Fields(iLoop).Name
                  If Not IsNull(recWorkflowElementItemValuesEdit.Fields(sName).Value) Then
                    rsElementItemValues.Fields(iLoop).Value = recWorkflowElementItemValuesEdit.Fields(sName).Value
                  End If
                 Next iLoop
                rsElementItemValues.Update
                
                'Get next item control value
                recWorkflowElementItemValuesEdit.MoveNext
              Loop
            End If
            rsElementItemValues.Close
            
            'Get next element item definition
            recWorkflowElementItemEdit.MoveNext
           Loop
        End If
        rsElementItems.Close

        ' Add the Element Columns (if required)
        rsElementColumns.Open "ASRSysWorkflowElementColumns", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

        recWorkflowElementColumnEdit.Index = "idxElementID"
        recWorkflowElementColumnEdit.Seek ">=", recWorkflowElementEdit!ID

        If Not recWorkflowElementColumnEdit.NoMatch Then
          Do While Not recWorkflowElementColumnEdit.EOF
            'If no more columns for this element exit loop
            If recWorkflowElementColumnEdit!elementID <> recWorkflowElementEdit!ID Then
              Exit Do
            End If

            'Add element column details to element column table
            rsElementColumns.AddNew
            For iLoop = 0 To rsElementColumns.Fields.Count - 1
              sName = rsElementColumns.Fields(iLoop).Name
              If Not IsNull(recWorkflowElementColumnEdit.Fields(sName)) Then
                rsElementColumns.Fields(iLoop) = recWorkflowElementColumnEdit.Fields(sName)
              End If
             Next iLoop
            rsElementColumns.Update

            'Get next element column definition
            recWorkflowElementColumnEdit.MoveNext
           Loop
        End If
        rsElementColumns.Close

        ' Add the Element Validations (if required)
        rsElementValidations.Open "ASRSysWorkflowElementValidations", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

        recWorkflowElementValidationEdit.Index = "idxElementID"
        recWorkflowElementValidationEdit.Seek ">=", recWorkflowElementEdit!ID

        If Not recWorkflowElementValidationEdit.NoMatch Then
          Do While Not recWorkflowElementValidationEdit.EOF
            'If no more Validations for this element exit loop
            If recWorkflowElementValidationEdit!elementID <> recWorkflowElementEdit!ID Then
              Exit Do
            End If

            'Add element Validation details to element Validation table
            rsElementValidations.AddNew
            For iLoop = 0 To rsElementValidations.Fields.Count - 1
              sName = rsElementValidations.Fields(iLoop).Name
              If Not IsNull(recWorkflowElementValidationEdit.Fields(sName)) Then
                rsElementValidations.Fields(iLoop) = recWorkflowElementValidationEdit.Fields(sName)
              End If
             Next iLoop
            rsElementValidations.Update

            'Get next element Validation definition
            recWorkflowElementValidationEdit.MoveNext
           Loop
        End If
        rsElementValidations.Close

        'Get next element definition
        .MoveNext
      Loop
    End If
  End With
  rsElements.Close

  ' Open the Workflow Links table on the server.
  rsLinks.Open "ASRSysWorkflowLinks", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
    
  With recWorkflowLinkEdit
    .Index = "idxWorkflowID"
    .Seek ">=", recWorkflowEdit!ID

    If Not .NoMatch Then
      Do While Not .EOF
        'If no more links for this workflow exit loop
        If !WorkflowID <> recWorkflowEdit!ID Then
          Exit Do
        End If

        'Add link details to links table
        rsLinks.AddNew
        For iLoop = 0 To rsLinks.Fields.Count - 1
          sName = rsLinks.Fields(iLoop).Name
          If Not IsNull(.Fields(sName)) Then
            rsLinks.Fields(iLoop) = .Fields(sName)
          End If
        Next iLoop
        rsLinks.Update

        'Get next link definition
        .MoveNext
      Loop
    End If
  End With
  rsLinks.Close

TidyUpAndExit:
  Set rsLinks = Nothing
  Set rsWorkflows = Nothing
  Set rsElementItems = Nothing
  Set rsElementItemValues = Nothing
  Set rsElementColumns = Nothing
  Set rsElementValidations = Nothing
  Set rsElements = Nothing
  WorkflowNew = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating workflow"
  fOK = False
  Resume TidyUpAndExit

End Function

