VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpression"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Expression definition variables.
Private mlngExpressionID As Long
Private msExpressionName As String
Private mlngBaseTableID As Long
Private miReturnType As ExpressionValueTypes
Private miExpressionType As ExpressionTypes
Private mlngParentComponentID As Long
Private msOwner As String
Private msAccess As String
Private msDescription As String

Private mfNew As Boolean
Private mfChanged As Boolean
Private mfDeleted As Boolean

' Class handling variables.
Private mcolComponents As Collection
Private mobjBadComponent As CExprComponent
Private mobjBaseComponent As CExprComponent
Private mlngCalculatedColumnID As Long

Private miEvaluatedReturnType As ExpressionValueTypes

Private msErrorMessage As String
Private mlngBottom As Long

Public glngPageNum As Long

Private mfrmUse As frmUsage

Private mlngUtilityID As Long
Private mlngUtilityBaseTable As Long
Private miWorkflowInitiationType As WorkflowInitiationTypes
Private mactlPrecedingWorkflowElements() As VB.Control
Private mactlAllWorkflowElements() As VB.Control

Private mdtLastSave As Date

Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  Dim iLoop1 As Integer

  ContainsExpression = False

  For iLoop1 = 1 To mcolComponents.Count
    If ContainsExpression Then
      Exit For
    End If
    
    With mcolComponents.Item(iLoop1)
      ContainsExpression = .ContainsExpression(plngExprID)
    End With
  Next iLoop1
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function



Public Function AddOperatorFunctionComponent(piComponentType As ExpressionComponentTypes, _
  piOpFuncID As Integer) As CExprComponent
  ' JPD20021108 Fault 3287
  ' Add a new component to the expression.
  ' Returns the new component object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As CExprComponent

  ' Instantiate a component object.
  Set objComponent = New CExprComponent
  
  ' Initialse the new component's properties.
  Set objComponent.ParentExpression = Me
  
  ' Instruct the new component to handle its own definition.
  fOK = objComponent.NewOperatorFunctionComponent(piComponentType, piOpFuncID)
  
  If fOK Then
    ' If the component definition was confirmed then
    ' add the new component to the expression's component
    ' collection.
    mcolComponents.Add objComponent
  End If
  
TidyUpAndExit:
  If fOK Then
    Set AddOperatorFunctionComponent = objComponent
  Else
    Set AddOperatorFunctionComponent = Nothing
  End If
  
  Set objComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Property Get ErrorMessage() As String
  ErrorMessage = msErrorMessage
End Property

Public Function CopyComponent() As CExpression
  ' Copies the selected component and all of it's children.

  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim fOK As Boolean
  Dim objCopyComponent As New CExpression

    fOK = True
    'objCopyComponent.ResetConstructedFlag (True)
    objCopyComponent.Name = msExpressionName
    objCopyComponent.BaseTableID = BaseTableID
    objCopyComponent.ExpressionType = ExpressionType
    objCopyComponent.ReturnType = ReturnType
    objCopyComponent.ExpressionID = ExpressionID
    
    'JPD 20070402 Fault 12089
    If (ExpressionType = giEXPR_WORKFLOWCALCULATION) _
      Or (ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
      Or (ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
    
      objCopyComponent.UtilityID = UtilityID
      objCopyComponent.UtilityBaseTable = UtilityBaseTable
      objCopyComponent.WorkflowInitiationType = WorkflowInitiationType
      objCopyComponent.PrecedingWorkflowElements = PrecedingWorkflowElements
      objCopyComponent.AllWorkflowElements = AllWorkflowElements
    End If
    
    ' Copy the children
    For iCount = 1 To mcolComponents.Count
        objCopyComponent.PasteComponent mcolComponents(iCount), mcolComponents(iCount), True
    Next iCount
  
TidyUpAndExit:
  If fOK Then
    Set CopyComponent = objCopyComponent
  Else
    Set CopyComponent = Nothing
  End If
  Set objCopyComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Sub ExpressionsUsedInThisExpression(palngExpressions As Variant)
  ' Return an array of the columns used this expression.
  Dim objComponent As CExprComponent
  
  For Each objComponent In mcolComponents
    ' Get the expressions components to check if they contain the given column.
    objComponent.ExpressionsUsedInThisComponent palngExpressions
  Next objComponent
  Set objComponent = Nothing
  
End Sub

Public Property Get LastSave() As Date
  LastSave = mdtLastSave
  
End Property

Public Function PasteComponent(pobjClipboard As CExprComponent, pobjPreviousComponent As CExprComponent, pbPasteBelow As Boolean) As CExprComponent
    ' Insert a component into the current expression
    ' just before the given component. Return the new component.
    On Error GoTo ErrorTrap
  
    Dim fOK As Boolean
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim objComponent As CExprComponent
  
    fOK = True
    iIndex = -1
  
    ' Calculate where the pasted component should be placed in the collection
    If pobjPreviousComponent.ComponentType = giCOMPONENT_EXPRESSION Then
        iIndex = 1
    Else
        ' Find the given component in the component collection.
        For iLoop = 1 To mcolComponents.Count
            If pobjPreviousComponent Is mcolComponents.Item(iLoop) Then
                iIndex = iLoop
                Exit For
            End If
        Next iLoop

        ' Paste component after the currently selected component
        If iIndex > -1 And pbPasteBelow Then
            iIndex = iIndex + 1
        End If

    End If

    ' Instantiate a component object.
    Set objComponent = pobjClipboard.CopyComponent
    Set objComponent.ParentExpression = Me

    If fOK Then
        ' Insert the new component into the expression's
        ' component collection before the given control.
        If (iIndex < 0) Or (iIndex > mcolComponents.Count) Then
            mcolComponents.Add objComponent
        Else
            mcolComponents.Add objComponent, , iIndex
        End If
    End If
  
TidyUpAndExit:
    If fOK Then
        Set PasteComponent = objComponent
    Else
        Set PasteComponent = Nothing
    End If
    Set objComponent = Nothing
    Exit Function

ErrorTrap:
    fOK = False
    Resume TidyUpAndExit

End Function

Public Property Get EvaluatedReturnType() As ExpressionValueTypes
  EvaluatedReturnType = miEvaluatedReturnType
End Property

'MH20010330 Allow "LET" when cloning the expression
Public Property Let EvaluatedReturnType(ByVal vNewValue As ExpressionValueTypes)
  miEvaluatedReturnType = vNewValue
End Property



Private Function ValidateOperatorParameters(plngOperatorID As Long, piResultType As ExpressionValueTypes, _
  piParam1Type As Integer, piParam2Type As Integer) As Boolean
  ' Validate the given operator with the given parameters.
  ' Return the result type in the piResultType parameter.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Validate the parameter types for the given operator.
  Select Case plngOperatorID
    Case 1 ' PLUS
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 2 ' MINUS
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 3 ' TIMES BY
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 4 ' DIVIDED BY
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC
    
    Case 5 ' AND
      fOK = (piParam1Type = giEXPRVALUE_LOGIC) And _
        (piParam2Type = giEXPRVALUE_LOGIC)
      piResultType = giEXPRVALUE_LOGIC

    Case 6 ' OR
      fOK = (piParam1Type = giEXPRVALUE_LOGIC) And _
        (piParam2Type = giEXPRVALUE_LOGIC)
      piResultType = giEXPRVALUE_LOGIC

    Case 7 ' IS EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_LOGIC) And (piParam2Type = giEXPRVALUE_LOGIC)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 8 ' IS NOT EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_LOGIC) And (piParam2Type = giEXPRVALUE_LOGIC)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 9 ' IS LESS THAN
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 10 ' IS GREATER THAN
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC
    
    Case 11 ' IS LESS THAN OR EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC
    
    Case 12 ' IS GREATER THAN OR EQUAL TO
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE)) Or _
        ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER))
      piResultType = giEXPRVALUE_LOGIC

    Case 13 ' NOT
      fOK = (piParam1Type = giEXPRVALUE_LOGIC)
      piResultType = giEXPRVALUE_LOGIC
    
    Case 14 ' IS CONTAINED WITHIN
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_LOGIC

    Case 15 ' TO THE POWER OF
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 16 ' MODULAS
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 17 ' CONCATENATED WITH
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER
    
    Case Else ' Unknown operator
      fOK = False
  End Select
  
TidyUpAndExit:
  If Not fOK Then
    piResultType = giEXPR_UNKNOWNTYPE
  End If
  
  ValidateOperatorParameters = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get CalculatedColumnID() As Long
  ' Return the Calculated Column ID.
  CalculatedColumnID = mlngCalculatedColumnID
  
End Property

Public Property Let CalculatedColumnID(ByVal plngNewValue As Long)
  ' Set the Calculated Column ID.
  mlngCalculatedColumnID = plngNewValue
  
End Property


Public Function ViewFilterCode() As String
  ' Create the code required for implementing a view filter.
  ' TRUE is returned if the code was constructed okay.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sFilterCode As String
  Dim sPrefix As String
  Dim alngSourceTables() As Long
  
  fOK = ConstructExpression
  
  If fOK Then
    ReDim alngSourceTables(2, 0)
    fOK = RuntimeCode(sFilterCode, alngSourceTables, False, sPrefix)
  End If
  
TidyUpAndExit:
  If fOK Then
    ViewFilterCode = sFilterCode & " = 1"
  Else
    ViewFilterCode = vbNullString
  End If
  
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the component type.
  ComponentType = giCOMPONENT_EXPRESSION
  
End Property


Public Function CloneExpression(pavCloneRegister As Variant) As CExpression
  ' Clone the current expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewExprID As Long
  Dim lngNewTableID As Long
  Dim lngNewUtilityID As Long
  Dim objClonedExpr As CExpression
  Dim objSourceComponent As CExprComponent
  Dim objCopyComponent As CExprComponent
  Dim colComponents As Collection
  
  ' Read the order definition.
  fOK = ConstructExpression

  If fOK Then
    ' Instantiate the clone expression.
    Set objClonedExpr = New CExpression
    ' Clone the current order's properties.
    With objClonedExpr
      .IsNew = True
      .IsDeleted = False
      .IsChanged = False
      .ExpressionID = 0
      .Name = Name
      .Owner = msOwner
      .Access = msAccess
      .Description = msDescription
      
      ' Get the new table ID from the clone register.
      lngNewTableID = mlngBaseTableID
      For iIndex = 1 To UBound(pavCloneRegister, 2)
        If pavCloneRegister(1, iIndex) = "TABLE" And _
          pavCloneRegister(2, iIndex) = mlngBaseTableID Then
          lngNewTableID = pavCloneRegister(3, iIndex)
          Exit For
        End If
      Next iIndex
      .BaseTableID = lngNewTableID
      
      ' Get the new table ID from the clone register.
      lngNewUtilityID = mlngUtilityID
      For iIndex = 1 To UBound(pavCloneRegister, 2)
        If pavCloneRegister(1, iIndex) = "WORKFLOW" And _
          pavCloneRegister(2, iIndex) = mlngUtilityID Then
          lngNewUtilityID = pavCloneRegister(3, iIndex)
          Exit For
        End If
      Next iIndex
      .UtilityID = lngNewUtilityID
      
      .ClearComponents
'If ASRDEVELOPMENT And ReturnType = 0 Then
'  Stop
'End If
      .ReturnType = ReturnType
      .EvaluatedReturnType = ReturnType   'MH20010330
      .ExpressionType = ExpressionType
      
      If mlngParentComponentID > 0 Then
        lngNewExprID = 0
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = mlngParentComponentID Then
            lngNewExprID = pavCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
        .ParentComponentID = lngNewExprID
      Else
        .ParentComponentID = 0
      End If
      
      ' Instantiate a new collection of components.
      Set colComponents = New Collection
      
      ' Clone the components.
      For Each objSourceComponent In mcolComponents
        ' Instantiate the cloned order item.

        Set objCopyComponent = objSourceComponent.CloneComponent(pavCloneRegister)
        fOK = Not objCopyComponent Is Nothing
        
        If fOK Then
          Set objCopyComponent.ParentExpression = objClonedExpr
          colComponents.Add objCopyComponent
        End If
        
        ' Disassociate object variables.
        Set objCopyComponent = Nothing
        
        If Not fOK Then
          Exit For
        End If
      Next objSourceComponent
      
      Set .Components = colComponents
      
      ' Disassociate object variables.
      Set colComponents = Nothing
    End With
  End If
  
TidyUpAndExit:
  
  'If ASRDEVELOPMENT Then
  '  If objClonedExpr.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS Then
  '    MsgBox objClonedExpr.ExpressionTypeName & ": '" & objClonedExpr.Name & "'" & vbnewline & objClonedExpr.ValidityMessage(objClonedExpr.ValidateExpression(True)), vbCritical, "ASRDEVELOPMENT"
  '    Stop
  '  End If
  'End If
  
  
  If fOK Then
    Set CloneExpression = objClonedExpr
  Else
    Set CloneExpression = Nothing
  End If
  ' Disassociate object variables.
  Set objClonedExpr = Nothing
  Set objSourceComponent = Nothing
  Set objCopyComponent = Nothing
  Set colComponents = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Sub ClearComponents()
  ' Clear the expression's component collection.
  On Error GoTo ErrorTrap
  
  ' Remove all components from the collection.
  Do While mcolComponents.Count > 0
    'mcolComponents(1).Terminate    'MH20060727
    mcolComponents.Remove 1
  Loop
  Set mcolComponents = Nothing
  
  ' Re-instantiate the collection.
  Set mcolComponents = New Collection
  
  Exit Sub

ErrorTrap:
  Err = False
  
End Sub

Public Function CalculatedColumnsThatUseThisExpression(ByRef pvColumns As Variant) As Variant
  ' Return an array of the calculated columns that use this expression.
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim rsCheck As DAO.Recordset
  Dim objExpr As CExpression
  
  If mlngExpressionID > 0 Then
    ' Get the columns that directly use this expression as the calculation.
    sSQL = "SELECT columnID" & _
      " FROM tmpColumns" & _
      " WHERE calcExprID = " & Trim$(Str$(mlngExpressionID)) & _
      " AND deleted = false"
    Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsCheck
      Do While Not .EOF
        iNextIndex = UBound(pvColumns) + 1
        ReDim Preserve pvColumns(iNextIndex)
        pvColumns(iNextIndex) = !ColumnID
        
        .MoveNext
      Loop
    
      .Close
    End With
    Set rsCheck = Nothing
  End If
  
  If (mlngExpressionID > 0) Then
    ' Get the columns that indirectly use this expression in other expressions thatn are used as the calculation.
    sSQL = "SELECT exprID" & _
      " FROM tmpComponents" & _
      " WHERE calculationID = " & Trim$(Str$(mlngExpressionID)) & _
      " OR filterID = " & Trim$(Str$(mlngExpressionID)) & _
      " OR fieldSelectionFilter = " & Trim$(Str$(mlngExpressionID))
    Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsCheck
      Do While Not .EOF
        Set objExpr = New CExpression
        objExpr.ExpressionID = !ExprID
        objExpr.CalculatedColumnsThatUseThisExpression pvColumns
        .MoveNext
      Loop
  
      .Close
    End With
    Set rsCheck = Nothing
  End If
  
End Function

Public Sub ColumnsUsedInThisExpression(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  Dim objComponent As CExprComponent
  
  For Each objComponent In mcolComponents
    ' Get the expressions components to check if they contain the given column.
    objComponent.ColumnsUsedInThisComponent pvColumns
  Next objComponent
  Set objComponent = Nothing
  
End Sub



Public Property Get IsChanged() As Boolean
  ' Return the 'changed expression' property.
  IsChanged = mfChanged
  
End Property
Public Property Let IsChanged(pfSetting As Boolean)
  ' Set the 'changed expression' property.
  mfChanged = pfSetting
  
End Property

Public Property Get IsNew() As Boolean
  ' Return the 'new expression' property.
  IsNew = mfNew

End Property
Public Property Let IsNew(pfNewValue As Boolean)
  ' Set the 'new expression' property.
  mfNew = pfNewValue

End Property

Public Property Get IsDeleted() As Boolean
  ' Return the 'deleted expression' property.
  IsDeleted = mfDeleted
  
End Property
Public Property Let IsDeleted(pfNewValue As Boolean)
  ' Set the 'deleted expression' property.
  mfDeleted = pfNewValue
  
End Property







Public Property Get BaseTableID() As Long
  ' Return the expressions base table ID.
  BaseTableID = mlngBaseTableID
  
End Property




Public Property Let BaseTableID(ByVal plngNewValue As Long)
  ' Set the expression base table property.
  mlngBaseTableID = plngNewValue

End Property


Private Sub Class_Initialize()
  ' Create a new collection to hold the expression's components.
  Set mcolComponents = New Collection

End Sub

Public Property Get Name() As String
  ' Return the expression name.
  Name = msExpressionName
  
End Property
Public Property Get Description() As String
  ' Return the expression's description.
  Description = msDescription
  
End Property

Public Property Let Description(psNewValue As String)
  ' Set the expression's descriprion property.
  msDescription = psNewValue
  
End Property

Public Property Let Name(ByVal psNewValue As String)
  ' Set the expression name.
  msExpressionName = psNewValue
  
End Property

Public Function AddComponent() As CExprComponent
  ' Add a new component to the expression.
  ' Returns the new component object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As CExprComponent

  ' Instantiate a component object.
  Set objComponent = New CExprComponent
  
  ' Initialse the new component's properties.
  Set objComponent.ParentExpression = Me
  
  ' Instruct the new component to handle its own definition.
  fOK = objComponent.NewComponent
  
  If fOK Then
    ' If the component definition was confirmed then
    ' add the new component to the expression's component
    ' collection.
    mcolComponents.Add objComponent
  End If
  
TidyUpAndExit:
  If fOK Then
    Set AddComponent = objComponent
  Else
    Set AddComponent = Nothing
  End If
  
  Set objComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function NewExpression() As Boolean
  ' Handle the definition of a new expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmExpr
  
  fOK = True
  
  ' Initialize the properties for a new expression.
  InitializeExpression
  
  ' Display the expression definition form.
  Set frmEdit = New frmExpr
  With frmEdit
    Set .Expression = Me
    .Show vbModal
  
    fOK = Not .Cancelled
  End With
  
  If fOK Then
    ' Write the new expression to the database.
    fOK = WriteExpression_Transaction
    
    ' If the write failed then re-initialize the
    ' properties for a new expression.
    If Not fOK Then
      InitializeExpression
    End If
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set frmEdit = Nothing
  NewExpression = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Function SelectExpression(Optional pvarReadOnly As Variant) As Boolean
  ' Display the expression selection form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fExit As Boolean
  Dim lngOldExpressionID As Long
  Dim frmSelectExpr As frmSelection
  Dim fReadOnly As Boolean
  
  If IsMissing(pvarReadOnly) Then
    fReadOnly = False
  Else
    fReadOnly = CBool(pvarReadOnly)
  End If
    
  fOK = (mlngBaseTableID > 0) Or (miExpressionType = giEXPR_WORKFLOWCALCULATION)
  
  If fOK Then
    Set frmSelectExpr = New frmSelection
    fExit = False
    lngOldExpressionID = mlngExpressionID
    
    ' Loop until an expression is selected, deselected, or the selection
    ' form is cancelled.
    Do While Not fExit
      frmSelectExpr.ForcedReadOnly = fReadOnly
    
      ' Set the selection form type.
      Set frmSelectExpr.Expression = Me

      ' Display the selection form.
      frmSelectExpr.Show vbModal
      
      ' Execute the required operation.
      Select Case frmSelectExpr.Action
        ' Add a new expression.
        Case edtAdd
          NewExpression
          
        ' Edit the selected expression.
        Case edtEdit
          mlngExpressionID = frmSelectExpr.SelectedID
          EditExpression fReadOnly
          
        ' Copy the selected expression.
        Case edtCopy
          mlngExpressionID = frmSelectExpr.SelectedID
          CopyExpression
          
        ' Delete the selected expression.
        Case edtDelete
          mlngExpressionID = frmSelectExpr.SelectedID
          ' Confirm the deletion.
          If MsgBox("Are you sure you want to delete this " & _
            ExpressionTypeName & " ?", _
            vbQuestion + vbYesNo, App.ProductName) = vbYes Then

            DeleteExpression_Transaction
          End If
          
        ' Print the selected expression.
        Case edtPrint
          mlngExpressionID = frmSelectExpr.SelectedID
          PrintExpression
          
        ' Select the selected expression and return to the calling form.
        Case edtSelect
          mlngExpressionID = frmSelectExpr.SelectedID
          ' Check that the selection is valid.
          If ValidateSelection Then
            fOK = True
            fExit = True
          End If
          
        ' Deselect all expressions and return to the calling form.
        Case edtDeselect
          mlngExpressionID = -1
          fOK = True
          fExit = True
          
        ' Exit the selection form and return to the calling form.
        Case edtCancel
          ' Check if the original expression still exists.
          With recExprEdit
            .Index = "idxExprID"
            .Seek "=", lngOldExpressionID, False
    
            If .NoMatch Then
              mlngExpressionID = 0
            Else
              mlngExpressionID = lngOldExpressionID
            End If
          End With
          
          fOK = False
          fExit = True
          
      End Select
    Loop
    
    Set frmSelectExpr = Nothing
  End If

TidyUpAndExit:
  SelectExpression = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function RuntimeCalculationCode(palngSourceTables As Variant, _
  psCalcCode As String) As Boolean
  ' Return TRUE if the Calculation code was created okay.
  ' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sRuntimeSQL As String
  Dim sPrefix As String
  
  ' Construct the expression from the database definition.
  fOK = ConstructExpression

  If fOK Then
    ' Get the Calculation code.
    fOK = RuntimeCode(sRuntimeSQL, palngSourceTables, False, sPrefix)
  End If

  If fOK Then
    fOK = (ValidateExpression(True) = giEXPRVALIDATION_NOERRORS)
  End If

  If fOK And (miReturnType = giEXPRVALUE_LOGIC) Then
    sRuntimeSQL = "convert(bit, " & sRuntimeSQL & ")"
  End If

TidyUpAndExit:
  If fOK Then
    psCalcCode = sRuntimeSQL
  Else
    psCalcCode = vbNullString
  End If
  RuntimeCalculationCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function





Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
  ' Return the SQL code that defines the expression.
  ' Used when creating the 'where clause' for view definitions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iParameter1Index As Integer
  Dim iParameter2Index As Integer
  Dim iMinOperatorPrecedence As Integer
  Dim iMaxOperatorPrecedence As Integer
  Dim sCode As String
  Dim sComponentCode As String
  Dim vParameter1 As Variant
  Dim vParameter2 As Variant
  Dim avValues() As Variant
  Dim mbCastAsFloat As Boolean

  mbCastAsFloat = False
  fOK = True
  sCode = vbNullString
  
  iMinOperatorPrecedence = -1
  iMaxOperatorPrecedence = -1
  
  ' Create an array of the components in the expression.
  ' Column 1 = operator id.
  ' Column 2 = component where clause code.
  ReDim avValues(2, mcolComponents.Count)
  For iLoop1 = 1 To mcolComponents.Count
    With mcolComponents.Item(iLoop1)
      ' If the current component is an operator then read the operator id into the array.
      If .ComponentType = giCOMPONENT_OPERATOR Then
        avValues(1, iLoop1) = .Component.OperatorID
        iMinOperatorPrecedence = IIf(iMinOperatorPrecedence > .Component.Precedence Or iMinOperatorPrecedence = -1, _
          .Component.Precedence, iMinOperatorPrecedence)
        iMaxOperatorPrecedence = IIf(iMaxOperatorPrecedence < .Component.Precedence Or iMaxOperatorPrecedence = -1, _
          .Component.Precedence, iMaxOperatorPrecedence)
      End If
                             
      fOK = .RuntimeCode(sComponentCode, palngSourceTables, pfValidating, psPrefix)
      
      If fOK Then
        avValues(2, iLoop1) = sComponentCode
      End If
    End With
    
    If Not fOK Then
      Exit For
    End If
  Next iLoop1
  
  If fOK Then
    ' Loop throught the expression's components checking that they are valid.
    ' Evaluate operators in the correct order.
    For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
      For iLoop2 = 1 To mcolComponents.Count
        With mcolComponents.Item(iLoop2)
          If .ComponentType = giCOMPONENT_OPERATOR Then
            If .Component.Precedence = iLoop1 Then
              ' Check that the operator has the correct parameter types.
              ' Read the value that follows the current operator.
              iParameter1Index = 0
              iParameter2Index = 0
              
              ' Read the index of the first parameter.
              For iLoop3 = iLoop2 + 1 To UBound(avValues, 2)
                If avValues(2, iLoop3) <> vbNullString Then
                  iParameter1Index = iLoop3
                  Exit For
                End If
              Next iLoop3
                    
              ' If a parameter has been found then read its value.
              ' Otherwise the expression is invalid.
              If iParameter1Index > 0 Then
                vParameter1 = avValues(2, iParameter1Index)
              End If
                
              ' Read a second parameter if required.
              If (.Component.OperandCount = 2) Then
                vParameter2 = vParameter1
                iParameter2Index = iParameter1Index
                iParameter1Index = 0
                
                ' Read the index of the parameter's value if there is one.
                For iLoop3 = iLoop2 - 1 To 1 Step -1
                  If avValues(2, iLoop3) <> vbNullString Then
                    iParameter1Index = iLoop3
                    Exit For
                  End If
                Next iLoop3
                  
                ' If a parameter has been found then read its value.
                ' Otherwise the expression is invalid.
                If iParameter1Index > 0 Then
                  vParameter1 = avValues(2, iParameter1Index)
                  
                  'JDM - 26/11/01 - Fault 3191 - Need tro cast values as float for division operators
                  If .Component.CastAsFloat Then
                    vParameter1 = "Cast(" & vParameter1 & " As Float)"
                  End If
                
                End If
              
                ' Update the array to reflect the constructed SQL code.
                avValues(1, iLoop2) = vbNullString
                If .Component.SQLType = "O" Then
                  If .Component.CheckDivideByZero Then
                    ' JPD20020415 Fault 3192
                    Select Case .Component.OperatorID
                      Case 16 'Modulus
                        avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbNewLine & _
                          vParameter1 & " - (CAST((" & vParameter1 & " / " & vParameter2 & ") AS INT) * " & vParameter2 & ")" & vbNewLine & _
                          ") end)"
                      Case Else
                        avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbNewLine & _
                          vParameter1 & " " & vbNewLine & _
                          avValues(2, iLoop2) & " " & vbNewLine & _
                          vParameter2 & vbNewLine & _
                          ") end)"
                    End Select
                  Else
                    Select Case .Component.OperatorID
                      Case 5 'And
                        avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) AND (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
                      Case 6 'Or
                        avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) OR (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
                      Case Else
                        avValues(2, iLoop2) = "(" & vbNewLine & _
                          vParameter1 & " " & vbNewLine & _
                          avValues(2, iLoop2) & " " & vbNewLine & _
                          vParameter2 & vbNewLine & _
                          ")"
                    End Select
                  End If
                Else
                  avValues(2, iLoop2) = _
                    IIf(LenB(.Component.SQLFixedParam1) <> 0, "(", vbNullString) & _
                    avValues(2, iLoop2) & vbNewLine & _
                    "(" & vbNewLine & _
                    vParameter1 & vbNewLine & _
                    ", " & vbNewLine & _
                    vParameter2 & vbNewLine & _
                    ")" & _
                    IIf(LenB(.Component.SQLFixedParam1) <> 0, " " & .Component.SQLFixedParam1 & ")", vbNullString)
                End If
                avValues(2, iParameter1Index) = vbNullString
                avValues(2, iParameter2Index) = vbNullString
              Else
                ' Update the array to reflect the constructed SQL code.
                avValues(1, iLoop2) = vbNullString
                If .Component.SQLType = "O" Then
                  Select Case .Component.OperatorID
                    Case 13 'Not
                      avValues(2, iLoop2) = "(CASE WHEN " & vParameter1 & " = 1 THEN 0 ELSE 1 END)"
                    Case Else
                      avValues(2, iLoop2) = "(" & vbNewLine & _
                        avValues(2, iLoop2) & " " & vbNewLine & _
                        vParameter1 & vbNewLine & _
                        ")"
                  End Select
                Else
                  avValues(2, iLoop2) = _
                    IIf(LenB(.Component.SQLFixedParam1) <> 0, "(", vbNullString) & _
                    avValues(2, iLoop2) & vbNewLine & _
                    "(" & vbNewLine & _
                    vParameter1 & vbNewLine & _
                    ")" & _
                    IIf(LenB(.Component.SQLFixedParam1) <> 0, " " & .Component.SQLFixedParam1 & ")", vbNullString)
                End If
                avValues(2, iParameter1Index) = vbNullString
              End If

              If (miExpressionType = giEXPR_VIEWFILTER) Or _
                (miExpressionType = giEXPR_RUNTIMEFILTER) Or _
                (miExpressionType = giEXPR_LINKFILTER) Or _
                (miExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
                (miExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
                
                If (.Component.ReturnType = giEXPRVALUE_LOGIC) And _
                  ((.Component.OperatorID <> 5) And _
                  (.Component.OperatorID <> 6) And _
                  (.Component.OperatorID <> 13)) Then
                  avValues(2, iLoop2) = "(CASE WHEN (" & avValues(2, iLoop2) & ") THEN 1 ELSE 0 END)"
                End If
              End If
            End If
          End If
        End With
      Next iLoop2
    Next iLoop1
     
    For iLoop1 = 1 To UBound(avValues, 2)
      If avValues(2, iLoop1) <> vbNullString Then
        sCode = avValues(2, iLoop1)
        Exit For
      End If
    Next iLoop1
  End If
  
TidyUpAndExit:
  If fOK Then
    psRuntimeCode = sCode
  Else
    psRuntimeCode = vbNullString
  End If
  RuntimeCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function RuntimeFilterCode(psFilterCode As String, _
  Optional pfValidating As Boolean, _
  Optional psPrefix As String) As Boolean
  ' Return TRUE if the filter code was created okay.
  ' Return the runtime filter SQL code in the parameter 'psFilterCode'.
  ' The filter code is to be used to validate the expression if the 'pfValidating' parameter is TRUE.
  ' This is used to suppress prompting the user for promted values, when we are only validating the expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iNextIndex As Integer
  Dim sWhereCode As String
  Dim sBaseTableSource As String
  Dim sJoinTableSource As String
  Dim sRuntimeFilterSQL As String
  Dim alngSourceTables() As Long
  Dim avRelatedTables() As Variant
  
  fOK = True
  
  ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If
  
  ' Construct the expression from the database definition.
'  fOK = ConstructExpression

  If fOK Then

    'MH20100324
    If glngExpressionTableIDForDeleteTrigger = mlngBaseTableID Then
      sBaseTableSource = "deleted"
    Else
      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mlngBaseTableID
      
        sBaseTableSource = !TableName
      End With
    End If

    sRuntimeFilterSQL = "SELECT DISTINCT " & sBaseTableSource & ".id" & vbNewLine & _
      "FROM " & sBaseTableSource & " " & vbNewLine
      
    ' Create an array of the IDs of the tables/view referred to in the expression.
    ' This is used for joining all of the tables/views used.
    ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
    ' Column 2 = table/view ID.
    ReDim alngSourceTables(2, 0)
    
    ' Get the filter code.
    fOK = RuntimeCode(sWhereCode, alngSourceTables, pfValidating, psPrefix)
  End If

  If fOK Then
    ' Create an array of the tables related to the expression base table.
    ' Used when Joining any other tables/view used.
    ' Column 1 = 'parent' if the expression's base table is the parent of the other table
    '            'child' if the expression's base table is the child of the other table
    ' Column 2 = ID of the other table
    ReDim avRelatedTables(2, 0)
    With recRelEdit
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do While Not .EOF
          If !parentID = mlngBaseTableID Then
            iNextIndex = UBound(avRelatedTables, 2) + 1
            ReDim Preserve avRelatedTables(2, iNextIndex)
            avRelatedTables(1, iNextIndex) = "parent"
            avRelatedTables(2, iNextIndex) = !childID
          ElseIf !childID = mlngBaseTableID Then
            iNextIndex = UBound(avRelatedTables, 2) + 1
            ReDim Preserve avRelatedTables(2, iNextIndex)
            avRelatedTables(1, iNextIndex) = "child"
            avRelatedTables(2, iNextIndex) = !parentID
          End If
        
          .MoveNext
        Loop
      End If
    End With
    
    ' Join any other tables used.
    For iLoop1 = 1 To UBound(alngSourceTables, 2)
      ' Join a table on a parent/child related to the base table.
      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", alngSourceTables(2, iLoop1)
      
        sJoinTableSource = !TableName
      End With
      
      For iLoop2 = 1 To UBound(avRelatedTables, 2)
        If avRelatedTables(2, iLoop2) = alngSourceTables(2, iLoop1) Then
            
          If avRelatedTables(1, iLoop2) = "parent" Then
            sRuntimeFilterSQL = sRuntimeFilterSQL & _
              "LEFT OUTER JOIN " & sJoinTableSource & " ON " & sBaseTableSource & ".id = " & sJoinTableSource & ".id_" & Trim$(Str$(mlngBaseTableID)) & " " & vbNewLine
          Else
            sRuntimeFilterSQL = sRuntimeFilterSQL & _
              "LEFT OUTER JOIN " & sJoinTableSource & " ON " & sBaseTableSource & ".id_" & Trim$(Str$(avRelatedTables(2, iLoop2))) & " = " & sJoinTableSource & ".id " & vbNewLine
          End If
          
          Exit For
        End If
      Next iLoop2
    Next iLoop1
        
    ' Add the filter 'where' clause code.
    If LenB(sWhereCode) <> 0 Then
      sWhereCode = sWhereCode & " = 1"
      
      sRuntimeFilterSQL = sRuntimeFilterSQL & _
        "WHERE " & vbNewLine & _
        sWhereCode & vbNewLine
    End If
  End If

TidyUpAndExit:
  If fOK Then
    psFilterCode = sRuntimeFilterSQL
  Else
    psFilterCode = vbNullString
  End If
  RuntimeFilterCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function












Public Function ValidateSelection() As Boolean
  ' Validate the expression section.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objExpr As CExpression
  Dim rsCheck As DAO.Recordset
  
  fOK = True
  
  ' Construct the expression to print.
  fOK = ConstructExpression
  
'  If fOK Then
'    ' If the expression is being selected to be a column calculation, check that
'    ' the column that is using this expression is not a component of the expression itself.
'    If (miExpressionType = giEXPR_COLUMNCALCULATION) And _
'      (mlngCalculatedColumnID > 0) Then
'
'      ' Get any components in this expression that use the given column.
'      sSQL = "SELECT *" & _
'        " FROM tmpComponents" & _
'        " WHERE exprID = " & trim$(str$(mlngExpressionID)) & _
'        " AND type = " & trim$(str$(giCOMPONENT_FIELD)) & _
'        " AND fieldColumnID = " & trim$(str$(mlngCalculatedColumnID))
'      Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'      fOK = (rsCheck.EOF And rsCheck.BOF)
'      rsCheck.Close
'      Set rsCheck = Nothing
'
'      If Not fOK Then
'        ' Tell the user why the expression can't be selected.
'        MsgBox "This " & LCase(ExpressionTypeName) & " cannot be selected." & vbCr & _
'          "A calculated column cannot use an expression which uses that column.", _
'          vbExclamation + vbOKOnly, App.ProductName
'      Else
'        ' Check any function sub-expressions.
'        sSQL = "SELECT tmpExpressions.exprID" & _
'          " FROM tmpExpressions, tmpComponents" & _
'          " WHERE tmpExpressions.parentComponentID = tmpComponents.componentID" & _
'          " AND tmpComponents.exprID = " & trim$(str$(mlngExpressionID)) & _
'          " AND tmpExpressions.deleted = FALSE"
'        Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'        With rsCheck
'          Do While (Not .EOF) And fOK
'            ' Instantiate each function parameter expression.
'            ' Instruct the function parameter expression to delete its components.
'            Set objExpr = New CExpression
'            objExpr.ExpressionID = !exprID
'            objExpr.CalculatedColumnID = mlngCalculatedColumnID
'            fOK = objExpr.ValidateSelection
'            Set objExpr = Nothing
'
'            .MoveNext
'          Loop
'
'          .Close
'        End With
'        Set rsCheck = Nothing
'      End If
'
'      If fOK Then
'        ' Check any calculation sub-expressions.
'        sSQL = "SELECT calculationID" & _
'          " FROM tmpComponents" & _
'          " WHERE exprID = " & trim$(str$(mlngExpressionID)) & _
'          " AND calculationID > 0"
'        Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'        With rsCheck
'          Do While (Not .EOF) And fOK
'            ' Instantiate each function parameter expression.
'            ' Instruct the function parameter expression to delete its components.
'            Set objExpr = New CExpression
'            objExpr.ExpressionID = !CalculationID
'            objExpr.CalculatedColumnID = mlngCalculatedColumnID
'            fOK = objExpr.ValidateSelection
'            Set objExpr = Nothing
'
'            .MoveNext
'          Loop
'
'          .Close
'        End With
'        Set rsCheck = Nothing
'      End If
'    End If
'  End If
  
  ValidateSelection = fOK

End Function


Public Function StoredProcedureCode(psVarName As String, Optional strBaseTable As String) As String
  ' Create the expression stored procedure.
  On Error GoTo ErrorTrap
  
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iParameter1Index As Integer
  Dim iParameter2Index As Integer
  Dim iMinOperatorPrecedence As Integer
  Dim iMaxOperatorPrecedence As Integer
  Dim sProcCode As SystemMgr.cStringBuilder
  Dim sVarName As String
  Dim sIfNullCode As String
  Dim vParameter1 As Variant
  Dim vParameter2 As Variant
  Dim avValues() As Variant

  Set sProcCode = New SystemMgr.cStringBuilder


  'If ASRDEVELOPMENT Then
  '  If mlngExpressionID = 247 Then
  '    strBaseTable = "deleted"
  '    Stop
  '  End If
  'End If


  iMinOperatorPrecedence = -1
  iMaxOperatorPrecedence = -1
  
  sProcCode.TheString = "    /* Evaluate expression " & Trim$(Str$(mlngExpressionID)) & ". */" & vbNewLine
  
  ' Create an array of the component types, operator ids and values.
  ' Column 1 = operator ID
  ' Column 2 = component stored procedure code
  ReDim avValues(3, mcolComponents.Count)
  For iLoop1 = 1 To mcolComponents.Count
    With mcolComponents.Item(iLoop1)
      ' If the current component is an operator then read the operator id into the array.
      If .ComponentType = giCOMPONENT_OPERATOR Then
        avValues(1, iLoop1) = .Component.OperatorID
        iMinOperatorPrecedence = IIf(iMinOperatorPrecedence > .Component.Precedence Or iMinOperatorPrecedence = -1, _
          .Component.Precedence, iMinOperatorPrecedence)
        iMaxOperatorPrecedence = IIf(iMaxOperatorPrecedence < .Component.Precedence Or iMaxOperatorPrecedence = -1, _
          .Component.Precedence, iMaxOperatorPrecedence)
      Else
        
        sProcCode.Append .StoredProcedureCode(strBaseTable) & vbNewLine & vbNewLine
        avValues(2, iLoop1) = "@component" & Trim$(Str$(.ComponentID))
      End If
    
      ' Read the component return type into the array.
      If .ComponentType = giCOMPONENT_WORKFLOWVALUE Then
        avValues(3, iLoop1) = .Component.ReturnType_FromDatabase
      Else
        avValues(3, iLoop1) = .ReturnType
      End If
    End With
  Next iLoop1
  
  ' Loop throught the expression's components.
  ' Evaluate operators in the correct order.
  For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
    For iLoop2 = 1 To mcolComponents.Count
      With mcolComponents.Item(iLoop2)
        If .ComponentType = giCOMPONENT_OPERATOR Then
          If .Component.Precedence = iLoop1 Then
            ' Read the current operator's parameters.
            iParameter1Index = 0
            iParameter2Index = 0
            
            ' Read the index of the first parameter.
            For iLoop3 = iLoop2 + 1 To UBound(avValues, 2)
              If avValues(2, iLoop3) <> vbNullString Then
                iParameter1Index = iLoop3
                Exit For
              End If
            Next iLoop3
                  
            ' If a parameter has been found then read its value.
            If iParameter1Index > 0 Then
              vParameter1 = avValues(2, iParameter1Index)
            End If
              
            ' Read a second parameter if required.
            If (.Component.OperandCount = 2) Then
              vParameter2 = vParameter1
              iParameter2Index = iParameter1Index
              iParameter1Index = 0
              
              ' Read the index of the parameter's value if there is one.
              For iLoop3 = iLoop2 - 1 To 1 Step -1
                If avValues(2, iLoop3) <> vbNullString Then
                  iParameter1Index = iLoop3
                  Exit For
                End If
              Next iLoop3
                
              ' If a parameter has been found then read its value.
              If iParameter1Index > 0 Then
                vParameter1 = avValues(2, iParameter1Index)
              End If
            End If
              
            sProcCode.Append "    /* Evaluate operator component " & Trim$(Str$(.ComponentID)) & ". */" & vbNewLine
            sVarName = "@component" & Trim$(Str$(.ComponentID))
            Select Case .ReturnType
              Case giEXPRVALUE_CHARACTER
                sProcCode.Append "    DECLARE " & sVarName & " nvarchar(MAX)" & vbNewLine
                sIfNullCode = "SET " & sVarName & " = ''" & vbNewLine
              Case giEXPRVALUE_DATE
                sProcCode.Append "    DECLARE " & sVarName & " datetime" & vbNewLine
                sIfNullCode = "SET " & sVarName & " = null" & vbNewLine
              Case giEXPRVALUE_NUMERIC
                sProcCode.Append "    DECLARE " & sVarName & " float" & vbNewLine
                sIfNullCode = "SET " & sVarName & " = 0" & vbNewLine
              Case giEXPRVALUE_LOGIC
                sProcCode.Append "    DECLARE " & sVarName & " bit" & vbNewLine
                sIfNullCode = "SET " & sVarName & " = 0" & vbNewLine
            End Select
    
            If (.Component.OperandCount = 2) Then
              If .Component.UnknownParameterTypes Then
                sProcCode.Append _
                  "    EXEC dbo." & .Component.SPName & "_" & Trim$(Str$(avValues(3, iParameter1Index))) & "_" & Trim$(Str$(avValues(3, iParameter2Index))) & " " & sVarName & " OUTPUT, " & vParameter1 & ", " & vParameter2
              Else
                sProcCode.Append _
                  "    EXEC dbo." & .Component.SPName & " " & sVarName & " OUTPUT, " & vParameter1 & ", " & vParameter2
              End If
            Else
              If .Component.UnknownParameterTypes Then
                sProcCode.Append _
                  "    EXEC dbo." & .Component.SPName & "_" & Trim$(Str$(avValues(3, iParameter1Index))) & " " & sVarName & " OUTPUT, " & vParameter1
              Else
                sProcCode.Append _
                  "    EXEC dbo." & .Component.SPName & " " & sVarName & " OUTPUT, " & vParameter1
              End If
            End If
            sProcCode.Append vbNewLine & "    IF " & sVarName & " IS NULL " & sIfNullCode
            
            ' Update the array to reflect the evaluated operation.
            avValues(1, iLoop2) = vbNullString
            avValues(2, iLoop2) = sVarName
            avValues(2, iParameter1Index) = vbNullString
            avValues(2, iParameter2Index) = vbNullString
          End If
        End If
      End With
    Next iLoop2
  Next iLoop1
  
  For iLoop1 = 1 To UBound(avValues, 2)
    If avValues(2, iLoop1) <> vbNullString Then
      sProcCode.Append "    SET " & psVarName & " = " & avValues(2, iLoop1)
      Exit For
    End If
  Next iLoop1
    
TidyUpAndExit:
  StoredProcedureCode = sProcCode.ToString
  Exit Function
  
ErrorTrap:
  sProcCode.TheString = vbNullString
  Resume TidyUpAndExit

End Function


Public Function InsertComponent(pobjComponent As CExprComponent) As CExprComponent
  ' Insert a component into the current expression
  ' just before the given component. Return the new component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim objComp As CExprComponent
  
  fOK = True
  iIndex = -1
  
  ' Find the given component in the component collection.
  For iLoop = 1 To mcolComponents.Count
    If pobjComponent Is mcolComponents.Item(iLoop) Then
      iIndex = iLoop
      Exit For
    End If
  Next iLoop
    
  ' Instantiate a component object.
  Set objComp = New CExprComponent
  Set objComp.ParentExpression = Me
  
  ' Instruct the new componnent to handle its definition.
  fOK = objComp.NewComponent
  
  If fOK Then
    ' Insert the new component into the expression's
    ' component collection before the given control.
    If iIndex < 0 Then
      mcolComponents.Add objComp
    Else
      mcolComponents.Add objComp, , iIndex
    End If
  End If

TidyUpAndExit:
  If fOK Then
    Set InsertComponent = objComp
  Else
    Set InsertComponent = Nothing
  End If
  Set objComp = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Function ModifyComponent(pobjComponent As CExprComponent) As CExprComponent
  ' Edit the given component. Return the new component if the changes
  ' were confirmed.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim objCopyComponent As New CExprComponent
  
  iIndex = -1
  
  ' Find the given component in the component collection.
  For iLoop = 1 To mcolComponents.Count
    If pobjComponent Is mcolComponents.Item(iLoop) Then
      iIndex = iLoop
      Exit For
    End If
  Next iLoop
  
  ' Create a copy of the given component in case the
  ' changes are cancelled.
  Set objCopyComponent = pobjComponent.CopyComponent
  fOK = Not objCopyComponent Is Nothing

  If fOK Then
    ' Instruct the copied component to handle the changes.
    fOK = objCopyComponent.ModifyComponent
  End If
  
  If fOK Then
    ' If the changes were confirmed then insert the copied
    ' component into the expressions component collection ...
    If iIndex < 0 Then
      mcolComponents.Add objCopyComponent
    Else
      mcolComponents.Add objCopyComponent, , iIndex
    End If

    ' ... and delete the original component from the collection.
    fOK = DeleteComponent(pobjComponent)
  End If
  
TidyUpAndExit:
  If fOK Then
    Set ModifyComponent = objCopyComponent
  Else
    Set ModifyComponent = Nothing
  End If
  
  ' Disassociate object variables.
  Set objCopyComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function DeleteComponent(pobjComponent As CExprComponent) As Boolean
  ' Remove the given component from the expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  fOK = True
  iIndex = 0
  
  ' Find the given component in the component collection.
  For iLoop = 1 To mcolComponents.Count
    If pobjComponent Is mcolComponents.Item(iLoop) Then
      iIndex = iLoop
      Exit For
    End If
  Next iLoop
    
  ' Delete the current component if it has been found.
  If iIndex > 0 Then
    mcolComponents.Remove iIndex
  End If

TidyUpAndExit:
  DeleteComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Function DeleteExpression(Optional pfUsageChecks As Variant) As Boolean
  ' Delete the expression.
  On Error GoTo ErrorTrap
  
  Dim fDeleteOK As Boolean
  Dim intObjectType As UsageCheckObject
  Dim fUsageChecks As Boolean
  Dim fShowMsg As Boolean
  
  fDeleteOK = False
  fUsageChecks = True
  
  If Not IsMissing(pfUsageChecks) Then
    fUsageChecks = CBool(pfUsageChecks)
  End If
  
  Select Case Me.ExpressionType
    Case ExpressionTypes.giEXPR_COLUMNCALCULATION
      intObjectType = UsageCheckObject.Calculation
    Case ExpressionTypes.giEXPR_VIEWFILTER
      intObjectType = UsageCheckObject.Filter
    Case ExpressionTypes.giEXPR_RUNTIMEFILTER
      intObjectType = UsageCheckObject.Filter
    Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
      intObjectType = UsageCheckObject.Calculation
  End Select
  
  'TM19012004
  Set mfrmUse = New frmUsage
  mfrmUse.ResetList
  ' Check that the expression can be deleted.
  ' ie. is not used anywhere.
  ' AE20080325 Fault #13039
  If ReadExpressionDetails Then
  
    fShowMsg = False
    If fUsageChecks Then
      fShowMsg = ExpressionIsUsed(mlngExpressionID)
    End If
    
    If fShowMsg Then
      Screen.MousePointer = vbDefault
      mfrmUse.ShowMessage Me.Name & " " & Me.ExpressionTypeName, "The " & Me.ExpressionTypeName & " cannot be deleted as the " & Me.ExpressionTypeName & " is used by the following:", intObjectType
      fDeleteOK = False
    Else
      ' Create the expression hieararchy to delete.
      ' AE20080325 Fault #13039
      'If ReadExpressionDetails Then
        ' Check if the expression is already marked as deleted.
        If mfDeleted Then
          fDeleteOK = True
        Else
          ' Mark the expression as deleted.
          mfDeleted = True
            
          ClearComponents
            
          ' Write the updated expression to the database.
          fDeleteOK = WriteExpression
        End If
      'End If
    End If
  End If
  
TidyUpAndExit:
  UnLoad mfrmUse
  Set mfrmUse = Nothing
  
  If fDeleteOK Then
    mfDeleted = True
    Application.Changed = True
  End If

  DeleteExpression = fDeleteOK
  Exit Function

ErrorTrap:
  fDeleteOK = False
  Err = False
  Resume TidyUpAndExit
  
End Function
Private Function DeleteExpression_Transaction() As Boolean
  ' Transaction wrapper for the 'DeleteExpression' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = DeleteExpression
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteExpression_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub PrintExpression()

  Dim fOK As Boolean

  ' Construct the expression to print.
  fOK = ConstructExpression

  ' Print the expression
  If fOK Then
    PrintExpressionWithoutConstructing
  End If
  
End Sub

Public Sub PrintExpressionWithoutConstructing()

  ' Print the expression.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim fOK As Boolean
  Dim sTableName As String
  Dim objComponent As CExprComponent
  Dim rsInfo As DAO.Recordset
  Dim objPrinter As clsPrintDef
  
  fOK = True
  glngPageNum = 0
  
  sSQL = "SELECT tableName" & _
    " FROM tmpTables" & _
    " WHERE tmpTables.tableID = " & Trim$(Str$(mlngBaseTableID))
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsInfo
    If Not (.EOF And .BOF) Then
      sTableName = !TableName
    Else
      sTableName = "<unknown>"
    End If
  
    .Close
  End With
  Set rsInfo = Nothing
   
  ' Load the printer object
  Set objPrinter = New clsPrintDef
  With objPrinter
    If .IsOK Then
      If .PrintStart(True) Then
        .PrintHeader ExpressionTypeName & " Definition : " & Trim(msExpressionName)
        .PrintNormal "Description : " & msDescription
        .PrintNormal
        .PrintNormal "Owner : " & msOwner
        .PrintNormal "Access : " & AccessDescription(msAccess)
        'Select Case msAccess
        '  Case "RW": .PrintNormal "Access : Read / Write"
        '  Case "RO": .PrintNormal "Access : Read only"
        '  Case "HD": .PrintNormal "Access : Hidden"
        'End Select
        .PrintNormal
        .PrintNormal "Base Table : " & Trim(sTableName)
        .PrintNormal
        
        .PrintTitle "Components"
        
        ' Print the components.
        Printer.FontBold = False
        For Each objComponent In mcolComponents
          fOK = objComponent.PrintComponent(1)
          If Not fOK Then
            Printer.KillDoc
            Exit For
          Else
            mlngBottom = CalculateBottomOfPage
            CheckEndOfPage2 mlngBottom, False
          End If
        Next
        
        ' Print the footer
        '.PrintEnd
        'NHRD23062003 Fault 3408 Page Footer
        mlngBottom = CalculateBottomOfPage
        Printer.CurrentY = mlngBottom + 1
        CheckEndOfPage2 mlngBottom, False
        Printer.EndDoc
        
        Dim objDefPrinter As cSetDfltPrinter
        
        Set objDefPrinter = New cSetDfltPrinter
        Do
          objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
        Loop While Printer.DeviceName <> gstrDefaultPrinterName
        Set objDefPrinter = Nothing
        
        Set objComponent = Nothing
      End If
    End If
  End With
  Set objPrinter = Nothing
  
TidyUpAndExit:
  Set objComponent = Nothing
  If Not fOK Then
    MsgBox "Unable to print the " & ExpressionTypeName & " '" & Name & "'." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As CExprComponent

  fOK = True
  
  glngBottom = CalculateBottomOfPage

  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  Printer.Print "Parameter : " & ComponentDescription
  
  CheckEndOfPage2 glngBottom
  
  ' Print the components.
  For Each objComponent In mcolComponents
    fOK = objComponent.PrintComponent(piLevel + 1)
    If Not fOK Then
      Printer.KillDoc
      Exit For
    End If
  Next

TidyUpAndExit:
  Set objComponent = Nothing
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub CopyExpression()
  ' Create a copy of the expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sName As String
  Dim frmEdit As frmExpr
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = False
  
  ' Construct the expression to copy.
  fOK = ConstructExpression
  
  If fOK Then
    ' Flag that the expression is new.
    mfNew = True
    mfChanged = False
    mfDeleted = False
    mlngExpressionID = 0
    
    ' Initialise the copied expression's name.
    sName = msExpressionName
    msExpressionName = "Copy_of_" & Trim(sName)

    'MH20010704 Fault 2059
    msOwner = gsUserName

    ' Display the form for editting the copied expression.
    Set frmEdit = New frmExpr
    Set frmEdit.Expression = Me
    frmEdit.Show vbModal
    
    ' If the copy (and changes) was confirmed then write the
    ' new expression to the database.
    fOK = Not frmEdit.Cancelled
    ' Disassociate object variables.
    Set frmEdit = Nothing
    
    If fOK Then
      fOK = WriteExpression
    End If
    
    ' Reset the expression selection variables.
    If Not fOK Then
      mfNew = True
      mfChanged = False
      mfDeleted = False
      mlngExpressionID = 0
      msExpressionName = sName
    End If
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  Set frmEdit = Nothing
  Exit Sub

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Sub


Public Property Get ComponentCount() As Integer
' JPD - NOT REQUIRED
'  ' Return the number of components in the expression.
'  ComponentCount = mcolComponents.Count
'
End Property






Public Property Get ComponentDescription() As String
  ' Return the expression name.
  ComponentDescription = msExpressionName
  
End Property

Public Property Get ExpressionType() As ExpressionTypes
  ' Return the expression's parent type property.
  ExpressionType = miExpressionType
  
End Property

Public Property Let ExpressionType(ByVal piNewValue As ExpressionTypes)
  ' Set the expression's parent type property.
  miExpressionType = piNewValue
  
End Property

Public Function ExpressionTypeName() As String
  ExpressionTypeName = ExprTypeName(miExpressionType)

End Function

Public Property Get ExpressionID() As Long
  ' Return the expression ID.
  ExpressionID = mlngExpressionID
  
End Property

Public Function ExpressionContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the expression makes any reference to the given column.
  ' Used to check for recursion.
  Dim fContains As Boolean
  Dim objComponent As CExprComponent
  
  fContains = False
  For Each objComponent In mcolComponents
    ' Get the expressions components to check if they contain the given column.
    fContains = objComponent.ComponentContainsColumn(plngColumnID)

    If fContains Then
      Exit For
    End If
  Next objComponent
  Set objComponent = Nothing
  
  ExpressionContainsColumn = fContains
  
End Function

Public Property Let ExpressionID(ByVal plngNewValue As Long)
  ' Set the expression ID.
  mlngExpressionID = plngNewValue
  
End Property
Public Function ReadExpressionDetails() As Boolean
  ' Locate the current expression in the database and read its
  ' parameters.
  ' NB. Use ConstructExpression to build the full expression.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean

  ' Locate the expression's database record.
  If FindExpression Then
    With recExprEdit
      ' Read the expression's properties.
      msExpressionName = IIf(IsNull(!Name), vbNullString, !Name)
      mlngBaseTableID = IIf(IsNull(!TableID), 0, !TableID)
      miReturnType = IIf(IsNull(!ReturnType), giEXPRVALUE_UNDEFINED, !ReturnType)
      miExpressionType = IIf(IsNull(!Type), giEXPR_COLUMNCALCULATION, !Type)
      
      If miExpressionType = giEXPR_RUNTIMECALCULATION Then
        miReturnType = giEXPRVALUE_UNDEFINED
      End If
      
      mlngParentComponentID = IIf(IsNull(!ParentComponentID), 0, !ParentComponentID)
      msOwner = IIf(IsNull(!userName), vbNullString, !userName)
      msAccess = IIf(IsNull(!Access), ACCESS_READWRITE, !Access)
      msDescription = IIf(IsNull(!Description), vbNullString, !Description)
      
      mlngUtilityID = IIf(IsNull(!UtilityID), 0, !UtilityID)
      
      mfNew = !New
      mfChanged = !Changed
      mfDeleted = !Deleted
      
      mdtLastSave = !LastSave
      
      fReadOk = True
    End With
  
  Else
    fReadOk = False
  End If
    
TidyUpAndExit:
  ReadExpressionDetails = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit
    
End Function

Public Function ConstructExpression() As Boolean
  ' Read the expression definition from the database.
  ' Construct the hierarchy of component class objects
  ' from this definition.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngComponentID As Long
  Dim objComponent As CExprComponent

  ' Read the expression's details .
  fOK = ReadExpressionDetails
  
  If fOK Then
    ' Clear the expressions collection of components.
    ClearComponents
    
    ' Read the expressions component collection from the database.
    With recCompEdit
      ' Locate the relevent records in the record set.
      .Index = "idxExprID"
      .Seek ">=", mlngExpressionID

      If Not .NoMatch Then
        Do While (Not .EOF) And fOK
          If !ExprID <> mlngExpressionID Then
            Exit Do
          End If
    
          lngComponentID = recCompEdit!ComponentID
          ' Instantiate a new component object.
          Set objComponent = New CExprComponent
           
          With objComponent
            Set .ParentExpression = Me
            .ComponentID = lngComponentID
            
            ' Instruct the new component to read it's own
            ' definition from the database.
            fOK = objComponent.ConstructComponent
          End With
          
          If fOK Then
            ' If the component definition was read correctly then
            ' add the new component to the expression's
            ' component collection.
            mcolComponents.Add objComponent
          End If
      
          ' Reposition the record pointer.
          .Index = "idxExprID"
          .Seek "=", mlngExpressionID, lngComponentID
            
          ' Disassociate object variables.
          Set objComponent = Nothing
                  
          .MoveNext
        Loop
      End If
    End With
  End If

TidyUpAndExit:
  Set objComponent = Nothing
  ConstructExpression = fOK
  Exit Function

ErrorTrap:
  fOK = False
  MsgBox "Unable to construct the expression.", _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  Resume TidyUpAndExit
    
End Function

Public Function WriteExpression() As Boolean
  ' Write the expression definition to the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objItem As CExprComponent
    
  With recExprEdit
    If (mfNew And Not mfChanged) Then
      
      'MH20010712
      'Replace the select max from access table with new method....
      'check the SQL database directly (incase anybody in Data Manager
      'has created a new expression.
      '''mlngExpressionID = UniqueColumnValue("tmpExpressions", "exprID")
      mlngExpressionID = GetUniqueID("Expressions", "tmpExpressions", "exprID")

      ' Add a record for new expression.
      fOK = (mlngExpressionID > 0)
      
      If fOK Then
        .AddNew
        !ExprID = mlngExpressionID
      End If
    Else
      ' Find the database record of the expression.
      fOK = FindExpression
      
      If fOK Then
        .Edit
      End If
    End If
    
    ' Write the expression definition to the database.
    If fOK Then
    
      mfChanged = True
      
      !Name = msExpressionName
      !TableID = mlngBaseTableID
      !ReturnType = miEvaluatedReturnType ' miReturnType
      !ReturnSize = 0 ' Not used.
      !ReturnDecimals = 0 ' Not used.
      !Type = miExpressionType
      !ParentComponentID = mlngParentComponentID
      !userName = msOwner
      !Access = msAccess
      !Description = msDescription

      !UtilityID = mlngUtilityID

      !New = mfNew
      !Changed = mfChanged
      !Deleted = mfDeleted
      
      !LastSave = Now
      
      .Update
      
    End If
    
  End With
  
  If fOK Then
    ' Delete the expression's existing components from the database.
    DeleteExistingComponents
    
    ' Add any components for this expression.
    For Each objItem In mcolComponents
      objItem.WriteComponent
    Next
  End If

TidyUpAndExit:
  Set objItem = Nothing
  WriteExpression = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Property Let ParentComponentID(ByVal plngNewValue As Long)
  ' Set the Parent component ID.
  mlngParentComponentID = plngNewValue
  
End Property



Private Sub EditExpression(pfReadOnly As Boolean)
  ' Edit the current expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmExpr
  
  fOK = False
  
  ' Construct the expression from the database definition.
  fOK = ConstructExpression
  
  If fOK Then
    ' Display the expression edit form.
    Set frmEdit = New frmExpr
    frmEdit.ForcedReadOnly = pfReadOnly
    Set frmEdit.Expression = Me
    
    frmEdit.Show vbModal
    
    ' If the changes were confirmed then write the changes to the database.
    fOK = Not frmEdit.Cancelled
    
    If fOK Then
      fOK = WriteExpression_Transaction
    End If
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  Set frmEdit = Nothing
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Public Function WriteExpression_Transaction() As Boolean
  ' Transaction wrapper for the 'WriteExpression' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = WriteExpression
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  WriteExpression_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function FindExpression() As Boolean
  ' Position the database pointer on the current expression's record.
  On Error GoTo ErrorTrap
  
  With recExprEdit
    .Index = "idxExprID"
    .Seek "=", mlngExpressionID, False
    FindExpression = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindExpression = False

End Function




Private Sub InitializeExpression()
  ' Initialize the properties for a new expression,
  ' and clear the expression's component collection.
  mlngExpressionID = 0
  
  msExpressionName = vbNullString
  mlngParentComponentID = 0
  msOwner = gsUserName
  msAccess = ACCESS_READWRITE
  msDescription = vbNullString
  
  mfNew = True
  mfChanged = False
  mfDeleted = False
  
  ' Clear any existing components from
  ' the expression's component collection.
  ClearComponents

End Sub

Public Function Initialise(plngBaseTableID As Long, _
  plngExpressionID As Long, _
  piType As Integer, _
  piReturnType As Integer) As Boolean
  ' Initialise the expression object.
  ' Return TRUE if everything was initialised okay.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  mlngBaseTableID = plngBaseTableID
  mlngExpressionID = plngExpressionID
  msExpressionName = vbNullString
  miExpressionType = piType
  miReturnType = piReturnType

  ReDim mactlPrecedingWorkflowElements(0)
  ReDim mactlAllWorkflowElements(0)

TidyUpAndExit:
  Initialise = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Let AllWorkflowElements(pactlElements As Variant)
  mactlAllWorkflowElements = pactlElements
End Property

Public Property Let PrecedingWorkflowElements(pactlElements As Variant)
  mactlPrecedingWorkflowElements = pactlElements
End Property

Public Property Get AllWorkflowElements() As Variant
  AllWorkflowElements = mactlAllWorkflowElements
End Property

Public Property Get PrecedingWorkflowElements() As Variant
  PrecedingWorkflowElements = mactlPrecedingWorkflowElements
End Property

Public Property Get BadComponent() As CExprComponent
  ' Return the component last caused the expression to fail its validity check.
  Set BadComponent = mobjBadComponent

End Property
Public Property Get BaseComponent() As CExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(pobjNewValue As CExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property

Public Function Component(piIndex As Integer) As CExprComponent
' JPD - NOT REQUIRED ?
'  ' Return the component at the given index in the expression's
'  ' component collection.
'  If piIndex > 0 And piIndex <= mcolComponents.Count Then
'    Set Component = mcolComponents.Item(piIndex)
'  Else
'    Set Component = Nothing
'  End If

End Function

Public Sub DeleteExistingComponents()
  ' Delete the expression's components and sub-expression's
  ' (ie. function parameter expressions) from the database.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsSubExpressions As DAO.Recordset
  Dim objExpr As CExpression
    
  ' Get the expression's function components from the database.
  sSQL = "SELECT tmpExpressions.exprID" & _
    " FROM tmpExpressions, tmpComponents" & _
    " WHERE tmpExpressions.parentComponentID=tmpComponents.componentID" & _
    " AND tmpComponents.exprID=" & Trim$(Str$(mlngExpressionID))
  Set rsSubExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  ' For each of the expression's function's parameter expressions ...
  With rsSubExpressions
    Do While Not .EOF
      ' Instantiate each function parameter expression.
      Set objExpr = New CExpression
      objExpr.ExpressionID = !ExprID
          
      ' Locate the expression record in the database.
      If objExpr.FindExpression Then
        ' Instruct the function parameter expression to delete
        ' its components
        objExpr.DeleteExistingComponents

        ' Mark the function parameter expression as deleted.
        With recExprEdit
          .Index = "idxExprID"
          .Seek "=", objExpr.ExpressionID, False

          If Not .NoMatch Then
            .Edit
            !Deleted = True
            .Update
          End If
        End With
      End If
      Set objExpr = Nothing
      
      .MoveNext
    Loop
    
    .Close
  End With
  
  ' Delete all existing components for this expression from the database.
  daoDb.Execute "DELETE FROM tmpComponents WHERE exprID=" & mlngExpressionID, dbFailOnError
  
TidyUpAndExit:
  Set objExpr = Nothing
  Set rsSubExpressions = Nothing
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub

Public Function ExpressionIsUsed(plngExprID As Long) As Boolean

  ' Return true if the expression is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Expressions may be used in the following contexts :
  '
  '   Table Record Descriptions
  '   Column Calculations & Validations
  '   Column Default Values
  '   View filters.
  '   Filters and Calculations in other expressions.
  '   Cross-tab filters
  '   Custom Report filters and calculations
  '   Data transfer filters
  '   Export filters
  '   Global function filters and calculations
  '   Mail Merge filters and calculations
  '   Workflow - Decision element flow calculations
  '   Workflow - Web Form descriptions
  '   Workflow - Web Form record selector filters
  '   Workflow - Web Form validations
  '   Workflow - Web Form label calulations
  '   Workflow - Web Form input value default values
  On Error GoTo ErrorTrap
  
  Dim sColumnName As String
  Dim sTableName As String
  Dim sExprName As String
  Dim lngExprBaseTableID As Long
  Dim sExprParentTable As String
  Dim sExprType As String
  Dim objExpr As CExpression
  Dim objComp As CExprComponent
  Dim lngCompID As Long
  Dim lngExprID As Long
  Dim lngTableID As Long
  Dim fUsed As Boolean
  Dim sSQL As String
  Dim rsADOCheck1 As ADODB.Recordset
  Dim rsADOCheck2 As ADODB.Recordset
  Dim rsADOCheck3 As ADODB.Recordset
  Dim rsADOCheck4 As ADODB.Recordset
  Dim rsADOCheck5 As ADODB.Recordset
  Dim rsADOCheck6 As ADODB.Recordset
  Dim rsADOCheck7 As ADODB.Recordset
  Dim rsADOCheck8 As ADODB.Recordset
  Dim rsADOCheck9 As ADODB.Recordset
  Dim rsADOCheck10 As ADODB.Recordset
  Dim rsADOCheck11 As ADODB.Recordset
  Dim rsADOCheck12 As ADODB.Recordset
  Dim rsADOCheck13 As ADODB.Recordset
  Dim rsADOCheck14 As ADODB.Recordset
  Dim rsADOCheck15 As ADODB.Recordset
  
  Dim rsDAOCheck As DAO.Recordset
  Dim wfTempElement As VB.Control
  Dim lngLoop As Long
  Dim iLoop2 As Integer
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim asValidations() As String
  
  Dim sGlobalFunctionType As String
  Dim sMessage As String

  sMessage = vbNullString

  fUsed = False
  
    ' Check that the expression is not used as a record description.
    With recTabEdit
      .Index = "idxTableID"
      .MoveFirst
      ' Find any tables that use this expression for its record
      ' description.
      Do While Not .EOF
        If (!RecordDescExprID = plngExprID) And _
          (Not !Deleted) Then
          fUsed = True
          mfrmUse.AddToList "Record Description : " & !TableName
        End If
        .MoveNext
      Loop
    End With
  
  
    ' Find any columns that use this expression as a Calculation,
    ' a Got Focus clause, or a Lost Focus clause.
    sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpColumns.tableID" & _
      " FROM tmpColumns" & _
      " WHERE deleted = FALSE" & _
      " AND (tmpColumns.calcExprID = " & Trim$(Str$(plngExprID)) & _
      " OR tmpColumns.lostFocusExprID = " & Trim$(Str$(plngExprID)) & ")"
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOCheck.BOF And rsDAOCheck.EOF) Then
      fUsed = True
      Do Until rsDAOCheck.EOF
        ' Get the column and table names.
        sColumnName = rsDAOCheck!ColumnName
          
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", rsDAOCheck!TableID
          
        If Not recTabEdit.NoMatch Then
          sTableName = recTabEdit!TableName
        Else
          sTableName = "<unknown>"
        End If
        mfrmUse.AddToList "Column : " & sColumnName & " <" & sTableName & ">"
        rsDAOCheck.MoveNext
      Loop
    End If
    rsDAOCheck.Close
    
    
    ' Find any columns that use this expression as a default value.
    sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpColumns.tableID" & _
      " FROM tmpColumns" & _
      " WHERE deleted = FALSE" & _
      " AND tmpColumns.dfltValueExprID = " & Trim$(Str$(plngExprID))
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOCheck.BOF And rsDAOCheck.EOF) Then
      fUsed = True
      Do Until rsDAOCheck.EOF
        ' Get the column and table names.
        sColumnName = rsDAOCheck!ColumnName
          
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", rsDAOCheck!TableID
          
        If Not recTabEdit.NoMatch Then
          sTableName = recTabEdit!TableName
        Else
          sTableName = "<unknown>"
        End If
        mfrmUse.AddToList ("Default Value : " & sColumnName & " <" & sTableName & ">")
        rsDAOCheck.MoveNext
      Loop
    End If
    rsDAOCheck.Close
    
    
    ' Find any views that use this expression as the view filter.
    sSQL = "SELECT DISTINCT tmpViews.viewName, tmpViews.viewTableID" & _
      " FROM tmpViews" & _
      " WHERE deleted = FALSE" & _
      " AND expressionID = " & Trim$(Str$(plngExprID))
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOCheck.BOF And rsDAOCheck.EOF) Then
      fUsed = True
      Do Until rsDAOCheck.EOF
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", rsDAOCheck!ViewTableID
          
        If Not recTabEdit.NoMatch Then
          sTableName = recTabEdit!TableName
        Else
          sTableName = "<unknown>"
        End If
        mfrmUse.AddToList ("View Filter : " & rsDAOCheck!viewName & " <" & sTableName & ">")
        rsDAOCheck.MoveNext
      Loop
    End If
    rsDAOCheck.Close
    
    
    ' Check that it is not used as a calculation in another expression,
    ' or as the filter in another expression.
    sSQL = "SELECT DISTINCT componentID" & _
      " FROM tmpComponents" & _
      " WHERE (calculationID = " & Trim$(Str$(plngExprID)) & ")" & _
      " OR (filterID = " & Trim$(Str$(plngExprID)) & ")" & _
      " OR (fieldSelectionFilter = " & Trim$(Str$(plngExprID)) & _
      " AND (type = " & Trim$(Str$(giCOMPONENT_FIELD)) & _
      " OR type = " & Trim$(Str$(giCOMPONENT_WORKFLOWFIELD)) & "))"
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOCheck.BOF And rsDAOCheck.EOF) Then
      fUsed = True
      Do Until rsDAOCheck.EOF
      
        Set objComp = New CExprComponent
        objComp.ComponentID = rsDAOCheck!ComponentID
        lngExprID = objComp.RootExpressionID
        Set objComp = Nothing
        
        ' Get the expression name and type description.
        Set objExpr = New CExpression
        objExpr.ExpressionID = lngExprID
    
        If objExpr.ReadExpressionDetails Then
          sExprName = objExpr.Name
          sExprType = objExpr.ExpressionTypeName
          lngExprBaseTableID = objExpr.BaseTableID
                
          If lngExprBaseTableID > 0 Then
            ' Get the expression's parent table name.
            recTabEdit.Index = "idxTableID"
            recTabEdit.Seek "=", lngExprBaseTableID
               
            If Not recTabEdit.NoMatch Then
              sExprParentTable = recTabEdit!TableName
            End If
            
            mfrmUse.AddToList (sExprType & " : " & sExprName & " <" & sExprParentTable & ">")
          Else
            mfrmUse.AddToList (sExprType & " : " & sExprName)
          End If
        End If
        ' Disassociate object variables.
        Set objExpr = Nothing
        rsDAOCheck.MoveNext
      Loop
    End If
    rsDAOCheck.Close
  
  
    ' Check that the expression is not used as a filter for a Cross-Tab.
    sSQL = "SELECT DISTINCT name" & _
        " FROM ASRSysCrossTab" & _
        " WHERE filterID = " & Trim$(Str$(plngExprID)) & _
        " AND CrossTabType <> " & ctt9GridBox
    Set rsADOCheck1 = New ADODB.Recordset
    rsADOCheck1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck1
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Cross Tab : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck1 = Nothing
  
      ' Check that the expression is not used as a filter for a 9-Box Grid Report
    sSQL = "SELECT DISTINCT name" & _
        " FROM ASRSysCrossTab" & _
        " WHERE filterID = " & Trim$(Str$(plngExprID)) & _
        " AND CrossTabType = " & ctt9GridBox
    Set rsADOCheck15 = New ADODB.Recordset
    rsADOCheck15.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck15
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "9-Box Grid Report : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck15 = Nothing
  
  
    ' Check that the expression is not used as a filter for a Custom Report.
    sSQL = "SELECT DISTINCT ASRSysCustomReportsName.name" & _
      " FROM ASRSysCustomReportsName" & _
      " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
      " WHERE ASRSysCustomReportsName.filter = " & Trim$(Str$(plngExprID)) & _
      " OR ASRSysCustomReportsName.parent1Filter = " & Trim$(Str$(plngExprID)) & _
      " OR ASRSysCustomReportsName.parent2Filter = " & Trim$(Str$(plngExprID)) & _
      " OR ASRSYSCustomReportsChildDetails.childFilter = " & Trim$(Str$(plngExprID))
    Set rsADOCheck2 = New ADODB.Recordset
    rsADOCheck2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck2
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Custom Report : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck2 = Nothing
  
    
    ' Check that the expression is not used as a calculation in a Custom Report.
    sSQL = "SELECT DISTINCT ASRSysCustomReportsName.name" & _
      " FROM ASRSysCustomReportsDetails" & _
      " INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsDetails.customReportID = ASRSysCustomReportsName.ID" & _
      " WHERE colExprID = " & Trim$(Str$(plngExprID)) & _
      " AND UPPER(type) = 'E'"
    Set rsADOCheck3 = New ADODB.Recordset
    rsADOCheck3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck3
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Custom Report : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck3 = Nothing
  
    
    ' Check that the expression is not used as a filter for a Data Transfer.
    sSQL = "SELECT DISTINCT name" & _
      " FROM ASRSysDataTransferName" & _
      " WHERE filterID = " & Trim$(Str$(plngExprID))
    Set rsADOCheck4 = New ADODB.Recordset
    rsADOCheck4.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck4
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Data Transfer : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck4 = Nothing


    ' Check that the expression is not used as a filter for an Export.
    sSQL = "SELECT DISTINCT name" & _
      " FROM ASRSysExportName" & _
      " WHERE filter = " & Trim$(Str$(plngExprID)) & _
      " OR parent1Filter = " & Trim$(Str$(plngExprID)) & _
      " OR parent2Filter = " & Trim$(Str$(plngExprID)) & _
      " OR childFilter = " & Trim$(Str$(plngExprID))
    Set rsADOCheck5 = New ADODB.Recordset
    rsADOCheck5.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck5
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Export : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck5 = Nothing
  
    
    ' Check that the expression is not used as a calculation in an Export.
    sSQL = "SELECT DISTINCT ASRSysExportName.name" & _
      " FROM ASRSysExportDetails" & _
      " INNER JOIN ASRSysExportName ON ASRSysExportDetails.exportID = ASRSysExportName.ID" & _
      " WHERE colExprID = " & Trim$(Str$(plngExprID)) & _
      " AND UPPER(type) = 'E'"
    Set rsADOCheck6 = New ADODB.Recordset
    rsADOCheck6.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck6
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Export : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck6 = Nothing
  
    
    ' Check that the expression is not used as a filter for a Global function.
    sSQL = "SELECT DISTINCT name, UPPER(type)" & _
      " FROM ASRSysGlobalFunctions" & _
      " WHERE filterID = " & Trim$(Str$(plngExprID))
    Set rsADOCheck7 = New ADODB.Recordset
    rsADOCheck7.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck7
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          ' Tell the user why the expression cannot be deleted.
          Select Case !Type
            Case "U"
              sGlobalFunctionType = "Update"
            Case "A"
              sGlobalFunctionType = "Add"
            Case "D"
              sGlobalFunctionType = "Delete"
            Case Else
              sGlobalFunctionType = "Function"
          End Select
          mfrmUse.AddToList "Global" & sGlobalFunctionType & " : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck7 = Nothing
  
  
    ' Check that the expression is not used as a calculation for a Global function.
    sSQL = "SELECT DISTINCT ASRSysGlobalFunctions.name, UPPER(ASRSysGlobalFunctions.type)" & _
      " FROM ASRSysGlobalItems" & _
      " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID" & _
      " WHERE ASRSysGlobalItems.exprID = " & Trim$(Str$(plngExprID)) & _
      " AND ASRSysGlobalItems.valueType = 4"
    Set rsADOCheck8 = New ADODB.Recordset
    rsADOCheck8.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck8
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          ' Tell the user why the expression cannot be deleted.
          Select Case !Type
            Case "U"
              sGlobalFunctionType = "Update"
            Case "A"
              sGlobalFunctionType = "Add"
            Case "D"
              sGlobalFunctionType = "Delete"
            Case Else
              sGlobalFunctionType = "Function"
          End Select
          mfrmUse.AddToList "Global" & sGlobalFunctionType & " : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck8 = Nothing
    
    
    ' Check that the expression is not used as a filter for a Mail Merge.
    sSQL = "SELECT DISTINCT name, isLabel" & _
      " FROM ASRSysMailMergeName" & _
      " WHERE filterID = " & Trim$(Str$(plngExprID))
    Set rsADOCheck9 = New ADODB.Recordset
    rsADOCheck9.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck9
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          If !IsLabel Then
            mfrmUse.AddToList "Envelope & Label : " & !Name
          Else
            mfrmUse.AddToList "Mail Merge : " & !Name
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck9 = Nothing
  
  
    ' Check that the expression is not used as a calculation for a Mail Merge.
    sSQL = "SELECT DISTINCT ASRSysMailMergeName.name" & _
      " FROM ASRSysMailMergeColumns" & _
      " INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeColumns.mailMergeID = ASRSysMailMergeName.mailMergeID" & _
      " WHERE ASRSysMailMergeColumns.columnID = " & Trim$(Str$(plngExprID)) & _
      " AND UPPER(ASRSysMailMergeColumns.type) = 'E'"
    Set rsADOCheck10 = New ADODB.Recordset
    rsADOCheck10.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck10
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          If !IsLabel Then
            mfrmUse.AddToList "Envelope & Label : " & !Name
          Else
            mfrmUse.AddToList "Mail Merge : " & !Name
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck10 = Nothing
    
    ' JDM - 10/11/2004 - Fault 9415 - Payroll setup
    'AE20071204 Fault #12658
    If IsModuleEnabled(modAccord) Then
      sSQL = "SELECT rtrim(t.TransferType) + ' - ' + f.Description AS Description" & _
             " FROM tmpAccordTransferFieldDefinitions f " & _
             " INNER JOIN tmpAccordTransferTypes t ON t.TransferTypeID = f.TransferTypeID" & _
             " WHERE ASRExprID = " & Me.ExpressionID & " AND ASRMapType = 1" & _
             " UNION SELECT rtrim(TransferType) + ' filter' AS Description" & _
             " FROM tmpAccordTransferTypes WHERE FilterID = " & Me.ExpressionID
      Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
      With rsDAOCheck
        If Not (.EOF And .BOF) Then
          fUsed = True
          Do Until .EOF
            mfrmUse.AddToList "Payroll Configuration : " & Trim(!Description)
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsDAOCheck = Nothing
    End If

    'TM20011003 Fault 2191 - 2
    'Make check for the expression calc used in email definitions.
    sSQL = "SELECT DISTINCT tmpEmailAddresses.name " & _
           "FROM tmpEmailAddresses " & _
           "WHERE tmpEmailAddresses.ExprID = " & Trim$(Str$(plngExprID)) & _
           " AND tmpEmailAddresses.Deleted = FALSE "
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
    With rsDAOCheck
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Email definition : " & !Name
          .MoveNext
        Loop
      End If
    
      .Close
    End With
    Set rsDAOCheck = Nothing


    'TM20011003 Fault 2191 - 3
    'Make check for the expression filter used in email link definitions.
    If ExpressionType = giEXPR_LINKFILTER Then
      'sSQL = "SELECT DISTINCT tmpEmailLinks.LinkID, tmpEmailLinks.Title , tmpColumns.ColumnName, tmpTables.TableName  " & _
             "FROM tmpEmailLinks, tmpColumns, tmpTables " & _
             "WHERE tmpEmailLinks.FilterID = " & CStr(plngExprID) & _
             "  AND tmpEmailLinks.ColumnID = tmpColumns.ColumnID " & _
             "  AND tmpColumns.TableID = tmpTables.TableID "
      sSQL = "SELECT DISTINCT tmpEmailLinks.LinkID, tmpEmailLinks.Title, tmpTables.TableName " & _
             "FROM tmpEmailLinks, tmpTables " & _
             "WHERE tmpEmailLinks.FilterID = " & CStr(plngExprID) & _
             "  AND tmpEmailLinks.TableID = tmpTables.TableID "
      Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
      With rsDAOCheck
        If Not (.EOF And .BOF) Then
          fUsed = True
          Do Until .EOF
            mfrmUse.AddToList "Email Link definition : " & !Title & " <" & !TableName & ">"
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsDAOCheck = Nothing


    'TM20011003 Fault 2191 - 4
    'Make check for the expression used in diary link definitions.
      sSQL = "SELECT DISTINCT tmpDiary.diaryID, tmpDiary.comment , tmpColumns.ColumnName, tmpTables.TableName " & _
             "FROM tmpDiary, tmpColumns, tmpTables " & _
             "WHERE tmpDiary.FilterID = " & CStr(plngExprID) & _
             "  AND tmpDiary.ColumnID = tmpColumns.ColumnID " & _
             "  AND tmpColumns.TableID = tmpTables.TableID "
      Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
      With rsDAOCheck
        If Not (.EOF And .BOF) Then
          fUsed = True
          Do Until .EOF
            mfrmUse.AddToList "Diary Link Definition : " & !Comment & " <" & !TableName & "." & !ColumnName & ">"
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsDAOCheck = Nothing
    
    
      'MH20040324 Check Outlook Calendar Link filters
      sSQL = "SELECT DISTINCT tmpOutlookLinks.LinkID, tmpOutlookLinks.Title, tmpTables.TableName " & _
             "FROM tmpOutlookLinks, tmpTables " & _
             "WHERE tmpOutlookLinks.FilterID = " & CStr(plngExprID) & _
             "  AND tmpOutlookLinks.TableID = tmpTables.TableID" & _
             "  AND tmpOutlookLinks.Deleted = False"
      Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
      With rsDAOCheck
        If Not (.EOF And .BOF) Then
          fUsed = True
          Do Until .EOF
            mfrmUse.AddToList "Outlook Calendar Link Definition : " & !Title & " <" & !TableName & ">"
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsDAOCheck = Nothing
    
      ' Check Workflow Triggered Link filters
      sSQL = "SELECT DISTINCT tmpWorkflows.name, tmpTables.TableName " & _
             "FROM tmpWorkflowTriggeredLinks, tmpTables, tmpWorkflows " & _
             "WHERE tmpWorkflowTriggeredLinks.FilterID = " & CStr(plngExprID) & _
             "  AND tmpWorkflowTriggeredLinks.TableID = tmpTables.TableID" & _
             "  AND tmpWorkflowTriggeredLinks.Deleted = False" & _
             "  AND tmpWorkflowTriggeredLinks.WorkflowID = tmpWorkflows.ID"
      Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
      With rsDAOCheck
        If Not (.EOF And .BOF) Then
          fUsed = True
          Do Until .EOF
            mfrmUse.AddToList "Workflow Triggered Link Definition : " & !Name & " <" & !TableName & ">"
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsDAOCheck = Nothing
    
    End If



    'MH20095027 Email Link content
    sSQL = "SELECT DISTINCT tmpTables.TableName, tmpEmailLinks.Title " & _
           "FROM tmpLinkContent, tmpEmailLinks, tmpTables " & _
           "WHERE (tmpLinkContent.FieldID = " & CStr(plngExprID) & " AND tmpLinkContent.FieldCode = 'E' " & _
           "  AND (tmpLinkContent.ContentID = tmpEmailLinks.SubjectContentID OR tmpLinkContent.ContentID = tmpEmailLinks.BodyContentID)) " & _
           "  AND tmpEmailLinks.TableID = tmpTables.TableID " & _
           "  AND tmpEmailLinks.Deleted = False "
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
    With rsDAOCheck
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Email Link : " & !Title & " <" & !TableName & ">"
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsDAOCheck = Nothing
    
    
    
    'MH20040324 Check Outlook Calendar Link subjects
    sSQL = "SELECT DISTINCT tmpOutlookLinks.LinkID, tmpOutlookLinks.Title, tmpTables.TableName " & _
           "FROM tmpOutlookLinks, tmpTables " & _
           "WHERE tmpOutlookLinks.Subject = " & CStr(plngExprID) & _
           "  AND tmpOutlookLinks.TableID = tmpTables.TableID" & _
           "  AND tmpOutlookLinks.Deleted = False"
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
    With rsDAOCheck
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Outlook Calendar Link Definition : " & !Title & " <" & !TableName & ">"
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsDAOCheck = Nothing
    
    
    
    'MH20040324 Check Outlook Folders calcs
    sSQL = "SELECT DISTINCT tmpOutlookFolders.FolderID, tmpOutlookFolders.Name, tmpTables.TableName " & _
           "FROM tmpOutlookFolders, tmpTables " & _
           "WHERE tmpOutlookFolders.ExprID = " & CStr(plngExprID) & _
           "  AND tmpOutlookFolders.TableID = tmpTables.TableID" & _
           "  AND tmpOutlookFolders.Deleted = False"
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
    With rsDAOCheck
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Outlook Folder : " & !Name & " <" & !TableName & ">"
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsDAOCheck = Nothing
    
    
    
    ' Check that the expression is not used as a filter for a Record Profile.
    sSQL = "SELECT DISTINCT ASRSysRecordProfileName.name" & _
      " FROM ASRSysRecordProfileName" & _
      " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
      " WHERE ASRSysRecordProfileName.filterID = " & Trim$(Str$(plngExprID)) & _
      " OR ASRSYSRecordProfileTables.filterID = " & Trim$(Str$(plngExprID))
    Set rsADOCheck11 = New ADODB.Recordset
    rsADOCheck11.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck11
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Record Profile : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck11 = Nothing
 
  
    ' Check that the expression is not used as a filter for a Calendar Report.
    sSQL = "SELECT DISTINCT ASRSysCalendarReports.name" & _
      " FROM ASRSysCalendarReports" & _
      " LEFT OUTER JOIN ASRSYSCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSYSCalendarReportEvents.calendarReportID" & _
      " WHERE ASRSysCalendarReports.filter = " & Trim$(Str$(plngExprID)) & _
      " OR ASRSYSCalendarReportEvents.filterID = " & Trim$(Str$(plngExprID))
    Set rsADOCheck12 = New ADODB.Recordset
    rsADOCheck12.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck12
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Calendar Report : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck12 = Nothing
    

    ' Check that the expression is not used as a calculation for a Calendar Report.
    sSQL = "SELECT DISTINCT ASRSysCalendarReports.name" & _
      " FROM ASRSysCalendarReports" & _
      " WHERE ASRSysCalendarReports.DescriptionExpr = " & Trim$(Str$(plngExprID))
    Set rsADOCheck13 = New ADODB.Recordset
    rsADOCheck13.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck13
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Calendar Report : " & !Name
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck13 = Nothing
    
    
    ' Check that the expression is not used as a filter for a Match Report.
    sSQL = "SELECT DISTINCT ASRSysMatchReportName.Name, ASRSysMatchReportName.MatchReportType" & _
        " FROM ASRSysMatchReportName" & _
        " WHERE ASRSysMatchReportName.Table1Filter = " & Trim$(Str$(plngExprID)) & _
        " OR ASRSysMatchReportName.Table2Filter = " & Trim$(Str$(plngExprID))
    Set rsADOCheck14 = New ADODB.Recordset
    rsADOCheck14.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsADOCheck14
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          Select Case !matchReportType
            Case 1: mfrmUse.AddToList "Succession Planning : " & !Name
            Case 2: mfrmUse.AddToList "Career Progression : " & !Name
            Case Else: mfrmUse.AddToList "Match Report : " & !Name
          End Select
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsADOCheck14 = Nothing
    
    If ExpressionType = giEXPR_WORKFLOWCALCULATION _
      Or ExpressionType = giEXPR_WORKFLOWSTATICFILTER _
      Or ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER Then
      
      If IsArray(mactlAllWorkflowElements) Then
        For lngLoop = 1 To UBound(mactlAllWorkflowElements)
          Set wfTempElement = mactlAllWorkflowElements(lngLoop)
          
          Select Case wfTempElement.ElementType
            ' Need to add in more check as expressions move into more areas of Workflow.
            Case elem_Decision
              If wfTempElement.DecisionFlowType = decisionFlowType_Expression Then
                If wfTempElement.DecisionFlowExpressionID = plngExprID Then
                  fUsed = True
                  mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & wfTempElement.Caption & ") : " & GetDecisionCaptionDescription(wfTempElement.DecisionCaptionType, True) & " flow calculation"
                End If
              End If
              
            Case elem_WebForm
              If wfTempElement.DescriptionExprID = plngExprID Then
                fUsed = True
                mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & wfTempElement.Caption & ") : Description calculation"
              End If
              
              asItems = wfTempElement.Items
              For iLoop2 = 1 To UBound(asItems, 2)
                If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) _
                  And (CLng(asItems(53, iLoop2)) = plngExprID) Then
                
                  fUsed = True
                  mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & wfTempElement.Caption & " - " & asItems(9, iLoop2) & ") : Record filter"
                ElseIf (asItems(2, iLoop2) = giWFFORMITEM_LABEL) _
                  And (CLng(asItems(56, iLoop2)) = plngExprID) Then
                
                  fUsed = True
                  mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & wfTempElement.Caption & ") : Label calculation"
                End If
                
                If ((asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_CHAR) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_DATE) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                    Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_NUMERIC)) _
                  And (CLng(asItems(56, iLoop2)) = plngExprID) Then
  
                  fUsed = True
                  mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & wfTempElement.Caption & " - " & asItems(9, iLoop2) & ") : Default value calculation"
                End If
              Next iLoop2
          
              asValidations = wfTempElement.Validations
              For iLoop2 = 1 To UBound(asValidations, 2)
                If CLng(asValidations(1, iLoop2)) = plngExprID Then
                  fUsed = True
                  mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & wfTempElement.Caption & ") : Validation calculation"
                End If
              Next iLoop2
            
            ' AE20080310 Fault #12990
            Case elem_Email
              ' ------------------------------------
              ' Email element items
              ' ------------------------------------
              asItems = wfTempElement.Items
              
              For iLoop2 = 1 To UBound(asItems, 2)
                ' Field calculations
                If asItems(2, iLoop2) = giWFEMAILITEM_CALCULATION Then
                  If CLng(asItems(56, iLoop2)) > 0 Then
                    If CLng(asItems(56, iLoop2)) = plngExprID Then
                      fUsed = True
                      mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & GetExpressionName(CLng(asItems(56, iLoop2))) & ") : Calculation"
                    End If
                  End If
                End If
              Next iLoop2
              
            ' AE20080310 Fault #12990
            Case elem_StoredData
              ' ------------------------------------
              ' Stored data element columns
              ' ------------------------------------
              avColumns = wfTempElement.DataColumns
              
              For iLoop2 = 1 To UBound(avColumns, 2)
                ' Field calculations
                If avColumns(4, iLoop2) = giWFDATAVALUE_CALC Then
                  If CLng(avColumns(10, iLoop2)) > 0 Then
                    If CLng(avColumns(10, iLoop2)) = plngExprID Then
                      fUsed = True
                      mfrmUse.AddToList wfTempElement.ElementTypeDescription & " (" & GetColumnName(CLng(avColumns(3, iLoop2))) & ") : Column calculation"
                    End If
                  End If
                End If
              Next iLoop2
              
          End Select
        
          Set wfTempElement = Nothing
        Next lngLoop
      End If
    End If
    
TidyUpAndExit:
  ' Disassociate object variables.
  Set objComp = Nothing
  Set objExpr = Nothing
  Set rsDAOCheck = Nothing
  Set rsADOCheck1 = Nothing
  Set rsADOCheck2 = Nothing
  Set rsADOCheck3 = Nothing
  Set rsADOCheck4 = Nothing
  Set rsADOCheck5 = Nothing
  Set rsADOCheck6 = Nothing
  Set rsADOCheck7 = Nothing
  Set rsADOCheck8 = Nothing
  Set rsADOCheck9 = Nothing
  Set rsADOCheck10 = Nothing
  Set rsADOCheck11 = Nothing
  Set rsADOCheck12 = Nothing
  Set rsADOCheck13 = Nothing
  Set rsADOCheck14 = Nothing
  Set rsADOCheck15 = Nothing
  
  ExpressionIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit
  
End Function




Public Function ValidateExpression(pfTopLevel As Boolean) As ExprValidationCodes
  ' Validate the expression. Return a code defining the validity of the expression.
  ' NB. This function is also good for evaluating the return type of an expression
  ' which has definite return type (eg. function sub-expressions, runtime calcs, etc).
  On Error GoTo ErrorTrap
  
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iParam1Type As Integer
  Dim iParam2Type As Integer
  Dim iParameter1Index As Integer
  Dim iParameter2Index As Integer
  Dim iParam1ReturnType As Integer
  Dim iParam2ReturnType As Integer
  Dim iOperatorReturnType As ExpressionValueTypes
  Dim iBadLogicColumnIndex As Integer
  Dim iMinOperatorPrecedence As Integer
  Dim iMaxOperatorPrecedence As Integer
  Dim iValidationCode As ExprValidationCodes
  Dim iEvaluatedReturnType As ExpressionValueTypes
  Dim sSPCode As String
  Dim sFilterCode As String
  Dim aiDummyValues() As Integer
  Dim iTempReturnType As Integer
  
  ' Initialise variables.
  iBadLogicColumnIndex = 0
  iMinOperatorPrecedence = -1
  iMaxOperatorPrecedence = -1
  iValidationCode = giEXPRVALIDATION_NOERRORS
  Set mobjBadComponent = Nothing
  
  ' If there are no expression components then report the error.
  If mcolComponents.Count = 0 Then
    iValidationCode = giEXPRVALIDATION_NOCOMPONENTS
  End If

  ' Create an array of the component return types and operator ids.
  ' Index 1 = operator id.
  ' Index 2 = return type.
  '
  ' Eg. the expression
  ' 'abc'
  ' Concatenated with
  ' Function 'uppercase'
  '   <parameter>
  '      Field 'personnel.surname'
  '
  ' will be represented in the array as
  '   -1,  giEXPRVALUE_CHARACTER
  '   17,  giEXPRVALUE_CHARACTER
  '   -1,  giEXPRVALUE_CHARACTER
  '
  ' The operators are then evaluated to leave the array as :
  '   -1,  -1
  '   -1,  giEXPRVALUE_CHARACTER
  '   -1,  -1
  '
  ' The one remaining value in the second column, after all operators have been evaluated
  ' is the return type of the expression.
  ReDim aiDummyValues(2, mcolComponents.Count)
  
  For iLoop1 = 1 To mcolComponents.Count
    ' Stop validating the expression if we already know its invalid.
    If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
      Exit For
    End If
    
    With mcolComponents.Item(iLoop1)
      ' If the current component is an operator then read the operator id into the array.
      If .ComponentType = giCOMPONENT_OPERATOR Then
        aiDummyValues(1, iLoop1) = .Component.OperatorID
        
        ' Remember the min and max operator precedence levels for later.
        iMinOperatorPrecedence = IIf((iMinOperatorPrecedence > .Component.Precedence) Or (iMinOperatorPrecedence = -1), _
          .Component.Precedence, iMinOperatorPrecedence)
        iMaxOperatorPrecedence = IIf((iMaxOperatorPrecedence < .Component.Precedence) Or (iMaxOperatorPrecedence = -1), _
          .Component.Precedence, iMaxOperatorPrecedence)
      Else
        aiDummyValues(1, iLoop1) = -1
      End If
      
      If .ComponentType = giCOMPONENT_FUNCTION Then
        ' Validate the function.
        ' NB. This also determines the function's return type if not already known.
        iValidationCode = .Component.ValidateFunction
        If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
          If .Component.BadComponent Is Nothing Then
            Set mobjBadComponent = mcolComponents.Item(iLoop1)
          Else
            Set mobjBadComponent = .Component.BadComponent
          End If
          Exit For
        End If
      End If
      
      ' Read the component return type into the array.
      aiDummyValues(2, iLoop1) = .ReturnType
    End With
  Next iLoop1
  
  ' Loop throught the expression's components checking that they are valid.
  ' Evaluate operators in the correct order of precedence.
  For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
    ' Stop validating the expression if we already know it's invalid.
    If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
      Exit For
    End If
    
    For iLoop2 = 1 To mcolComponents.Count
      ' Stop validating the expression if we already know it's invalid.
      If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
        Exit For
      End If
      
      With mcolComponents.Item(iLoop2)
        If .ComponentType = giCOMPONENT_OPERATOR Then
          If .Component.Precedence = iLoop1 Then
            ' Check that the operator has the correct parameter types.
            ' Read the dummy value that follows the current operator.
            iParameter1Index = 0
            iParameter2Index = 0
            For iLoop3 = iLoop2 + 1 To UBound(aiDummyValues, 2)
              ' If an operator follows the operator then the expression is invalid.
              If aiDummyValues(1, iLoop3) > 0 Then
                Set mobjBadComponent = mcolComponents.Item(iLoop2)
                iValidationCode = giEXPRVALIDATION_MISSINGOPERAND
                Exit For
              End If
                
              ' Read the index of the parameter.
              If aiDummyValues(2, iLoop3) > -1 Then
                iParameter1Index = iLoop3
                Exit For
              End If
            Next iLoop3
                  
            ' If a parameter has been found then read its dummy value.
            ' Otherwise the expression is invalid.
            If iParameter1Index = 0 Then
              Set mobjBadComponent = mcolComponents.Item(iLoop2)
              iValidationCode = giEXPRVALIDATION_MISSINGOPERAND
            End If
              
            ' Read a second parameter if required.
            If (.Component.OperandCount = 2) And _
              (iValidationCode = giEXPRVALIDATION_NOERRORS) Then
              
              iParameter2Index = iParameter1Index
                
              ' Read the dummy value that precedes the current operator.
              iParameter1Index = 0
              For iLoop3 = iLoop2 - 1 To 1 Step -1
                ' If an operator follows the operator then the expression is invalid.
                If aiDummyValues(1, iLoop3) > 0 Then
                  Set mobjBadComponent = mcolComponents.Item(iLoop2)
                  iValidationCode = giEXPRVALIDATION_MISSINGOPERAND
                  Exit For
                End If
                  
                ' Read the index of the parameter.
                If aiDummyValues(2, iLoop3) > -1 Then
                  iParameter1Index = iLoop3
                  Exit For
                End If
              Next iLoop3
                
              ' If a parameter has been found then read its dummy value.
              ' Otherwise the expression is invalid.
              If iParameter1Index = 0 Then
                Set mobjBadComponent = mcolComponents.Item(iLoop2)
                iValidationCode = giEXPRVALIDATION_MISSINGOPERAND
              End If
            End If
             
            ' Validate the operator by evaluating it with the dummy parmameters.
            ' NB. This also determines the operator's return type if not already known.
            ' Only try to evaluate the dummy operation if we still think
            ' it is valid.
            If iValidationCode = giEXPRVALIDATION_NOERRORS Then
              If Not ValidateOperatorParameters(.Component.OperatorID, iOperatorReturnType, aiDummyValues(2, iParameter1Index), _
                IIf(.Component.OperandCount = 2, aiDummyValues(2, iParameter2Index), giEXPRVALUE_UNDEFINED)) Then
                
                iValidationCode = giEXPRVALIDATION_OPERANDTYPEMISMATCH
                Set mobjBadComponent = mcolComponents.Item(iLoop2)
              Else
                ' Check that operators with logic parameters are valid.
                If (iBadLogicColumnIndex = 0) And (iParameter2Index > 0) Then
                  iParam1Type = mcolComponents.Item(iParameter1Index).ComponentType
                  iParam1ReturnType = mcolComponents.Item(iParameter1Index).ReturnType
                  iParam2Type = mcolComponents.Item(iParameter2Index).ComponentType
                  iParam2ReturnType = mcolComponents.Item(iParameter2Index).ReturnType
  
                  If ((iParam1Type = giCOMPONENT_FIELD) And (iParam1ReturnType = giEXPRVALUE_LOGIC)) And _
                    (((iParam2Type <> giCOMPONENT_FIELD) And (iParam2Type <> giCOMPONENT_VALUE)) Or _
                    (iParam2ReturnType <> giEXPRVALUE_LOGIC)) Then
  
                    iBadLogicColumnIndex = iParameter1Index
                  End If
  
                  If ((iParam2Type = giCOMPONENT_FIELD) And (iParam2ReturnType = giEXPRVALUE_LOGIC)) And _
                    (((iParam1Type <> giCOMPONENT_FIELD) And (iParam1Type <> giCOMPONENT_VALUE)) Or _
                    (iParam1ReturnType <> giEXPRVALUE_LOGIC)) Then
  
                    iBadLogicColumnIndex = iParameter2Index
                  End If
                End If
              
                ' Update the array to reflect the evaluated operation.
                aiDummyValues(1, iLoop2) = -1
                aiDummyValues(2, iParameter1Index) = -1
                aiDummyValues(2, iParameter2Index) = -1
              End If
            End If
          End If
        End If
      End With
    Next iLoop2
  Next iLoop1
  
  ' Check the expression has valid syntax (ie. if the components have evaluated to a single value).
  ' Get the evaluated return type while we're at it.
  If iValidationCode = giEXPRVALIDATION_NOERRORS Then
    iEvaluatedReturnType = giEXPRVALUE_UNDEFINED
    
    For iLoop1 = 1 To UBound(aiDummyValues, 2)
      If aiDummyValues(2, iLoop1) > -1 Then
        ' If the expression has more than one component after evaluating
        ' all of the operators then the expression is invalid.
        If iEvaluatedReturnType <> giEXPRVALUE_UNDEFINED Then
          iValidationCode = giEXPRVALIDATION_SYNTAXERROR
          Exit For
        End If
        
        iEvaluatedReturnType = aiDummyValues(2, iLoop1)
      End If
    Next iLoop1
  End If
  
  ' Set the expression's return type if it is not already set.
  If iValidationCode = giEXPRVALIDATION_NOERRORS Then
    If (miReturnType = giEXPRVALUE_UNDEFINED) Or (miReturnType = giEXPRVALUE_BYREF_UNDEFINED) Then
      miReturnType = iEvaluatedReturnType
    End If
  End If
  
  ' RH 23/10 - Set module level variable, so can access as property
  miEvaluatedReturnType = iEvaluatedReturnType
  
  ' Check the evaluated return type matches the pre-set return type.
  If (iValidationCode = giEXPRVALIDATION_NOERRORS) And (iEvaluatedReturnType <> miReturnType) Then
    iValidationCode = giEXPRVALIDATION_EXPRTYPEMISMATCH
  End If
  
  ' JPD20020419 Fault 3687
  ' Run the filter's SQL code to see if it is valid.
  If (iValidationCode = giEXPRVALIDATION_NOERRORS) And _
    pfTopLevel Then
    iTempReturnType = miReturnType
    iValidationCode = ValidateSQLCode
    miReturnType = iTempReturnType
  End If
  
TidyUpAndExit:
  ValidateExpression = iValidationCode
  Exit Function
  
ErrorTrap:
  iValidationCode = giEXPRVALIDATION_UNKNOWNERROR
  Resume TidyUpAndExit

End Function
Public Function ValidateSQLCode() As ExprValidationCodes
  ' Validate the expression's SQL code. This picks up on errors such as too many nested levels of the CASE statement.
  On Error GoTo ErrorTrap
  
  Dim lngCalcViews() As Long
  Dim objCalcExpr As CExpression
  Dim intCount As Integer
  Dim sSource As String
  Dim sSPCode As String
  Dim strJoinCode As String
  Dim iValidationCode As ExprValidationCodes
  Dim sSQLCode As String
  Dim lngOriginalExprID As Long
  Dim sOriginalSQLCode As String
  Dim alngSourceTables() As Long
  Dim sProcName As String
  Dim intStart As Integer
  Dim intFound As Integer
  
  iValidationCode = giEXPRVALIDATION_NOERRORS
    
  If ((miExpressionType = giEXPR_VIEWFILTER) Or _
    (miExpressionType = giEXPR_RUNTIMEFILTER) Or _
    (miExpressionType = giEXPR_LINKFILTER)) Then
      
    If RuntimeFilterCode(sSQLCode, True) Then

      On Error GoTo SQLCodeErrorTrap
        
      ' Drop the test stored procedure.
      DropProcedure "tmpsp_ASRExprTest"
        
      ' Create the test stored procedure to see if the filter expression is valid.
      sSPCode = " CREATE PROCEDURE tmpsp_ASRExprTest AS " & sSQLCode
      gADOCon.Execute sSPCode, , adExecuteNoRecords
      
      ' Drop the test stored procedure.
      DropProcedure "tmpsp_ASRExprTest"
    
      On Error GoTo ErrorTrap
    Else
      iValidationCode = giEXPRVALIDATION_FILTEREVALUATION
    End If
  End If
  
TidyUpAndExit:
  ValidateSQLCode = iValidationCode
  Exit Function
        
SQLCodeErrorTrap:
  iValidationCode = giEXPRVALIDATION_SQLERROR
  msErrorMessage = Err.Description
  
  intFound = 0
  Do
    intStart = intFound
    intFound = InStr(intStart + 1, msErrorMessage, "]")
  Loop While intFound > 0
  
  If intStart > 0 And intStart < LenB(Trim(msErrorMessage)) Then
    msErrorMessage = Trim(Mid(msErrorMessage, intStart + 1))
  End If

  Set mobjBadComponent = Nothing
  Resume TidyUpAndExit

ErrorTrap:
  iValidationCode = giEXPRVALIDATION_UNKNOWNERROR
  Resume TidyUpAndExit
        
End Function

Public Function UsesCalculation(plngExprID As Long) As Variant
  ' Return TRUE if the expression (or any calculations within it) uses the given expressionn.
  Dim sSQL As String
  Dim rsCheck As DAO.Recordset
  Dim objExpr As CExpression
  
  ' Check if this expression uses the given expression as a calculation.
  sSQL = "SELECT *" & _
    " FROM tmpComponents" & _
    " WHERE calculationID = " & Trim$(Str$(plngExprID)) & _
    " AND exprID = " & Trim$(Str$(mlngExpressionID))
  Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsCheck
    UsesCalculation = Not (.EOF And .BOF)
    .Close
  End With
  Set rsCheck = Nothing
  
  If Not UsesCalculation Then
    ' Check if this expression uses the any calculations that use the given expression as a calculation.
    sSQL = "SELECT calculationID" & _
      " FROM tmpComponents" & _
      " WHERE exprID = " & Trim$(Str$(mlngExpressionID)) & _
      " AND calculationID > 0 "
  Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsCheck
      Do While (Not UsesCalculation) And (Not .EOF)
        Set objExpr = New CExpression
        objExpr.ExpressionID = !CalculationID
        UsesCalculation = objExpr.UsesCalculation(plngExprID)
        .MoveNext
      Loop
  
      .Close
    End With
    Set rsCheck = Nothing
  End If
  
End Function



Public Function EvaluateExpression(plngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Evaluate the expression. Return the evaluate result
'  On Error GoTo BasicErrorTrap
'
'  Dim iLoop1 As Integer
'  Dim iLoop2 As Integer
'  Dim iLoop3 As Integer
'  Dim vResult As Variant
'  Dim avValues() As Variant
'  Dim vParameter1 As Variant
'  Dim vParameter2 As Variant
'  Dim iParameter1Index As Integer
'  Dim iParameter2Index As Integer
'  Dim objEvaluator As New ASRHRProEvaluator.Evaluator
'  Dim iMinOperatorPrecedence As Integer
'  Dim iMaxOperatorPrecedence As Integer
'
'  iMinOperatorPrecedence = -1
'  iMaxOperatorPrecedence = -1
'
'  ' Create an array of the component types, operator ids and values.
'  For iLoop1 = 1 To mcolComponents.Count
'
'    ReDim Preserve avValues(4, iLoop1)
'
'    With mcolComponents.Item(iLoop1)
'      ' If the current component is an operator then read the operator id into
'      ' the array.
'      If .ComponentType = giCOMPONENT_OPERATOR Then
'        avValues(1, iLoop1) = .Component.OperatorID
'        iMinOperatorPrecedence = IIf(iMinOperatorPrecedence > .Component.Precedence Or iMinOperatorPrecedence = -1, _
'          .Component.Precedence, iMinOperatorPrecedence)
'        iMaxOperatorPrecedence = IIf(iMaxOperatorPrecedence < .Component.Precedence Or iMaxOperatorPrecedence = -1, _
'          .Component.Precedence, iMaxOperatorPrecedence)
'          avValues(4, iLoop1) = False
'      Else
'        ' Write the component's evaluated value into the array.
'        avValues(2, iLoop1) = .EvaluateComponent(pLngRecordID, pfDoPrompts)
'        avValues(4, iLoop1) = True
'      End If
'
'      ' Read the component return type into the array.
'      avValues(3, iLoop1) = .ReturnType
'
'    End With
'  Next iLoop1
'
'  ' Loop throught the expression's components checking that they
'  ' are valid.
'  ' Evaluate operators in the correct order.
'  For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
'
'    For iLoop2 = 1 To mcolComponents.Count
'
'      With mcolComponents.Item(iLoop2)
'
'        If .ComponentType = giCOMPONENT_OPERATOR Then
'
'          If .Component.Precedence = iLoop1 Then
'
'            ' Check that the operator has the correct parameter types.
'            ' Read the value that follows the current operator.
'            iParameter1Index = 0
'            iParameter2Index = 0
'            For iLoop3 = iLoop2 + 1 To UBound(avValues, 2)
'
'              ' Read the index of the parameter's value if there is one.
'              If avValues(4, iLoop3) Then
'                iParameter1Index = iLoop3
'                Exit For
'              End If
'
'            Next iLoop3
'
'            ' If a parameter has been found then read its value.
'            ' Otherwise the expression is invalid.
'            If iParameter1Index > 0 Then
'              vParameter1 = avValues(2, iParameter1Index)
'            End If
'
'            ' Read a second parameter if required.
'            If .Component.OperandCount = 2 Then
'              vParameter2 = vParameter1
'              iParameter2Index = iParameter1Index
'
'              ' Read the value that precedes the current operator.
'              iParameter1Index = 0
'              For iLoop3 = iLoop2 - 1 To 1 Step -1
'                ' Read the index of the parameter's value if there is one.
'                If avValues(4, iLoop3) Then
'                  iParameter1Index = iLoop3
'                  Exit For
'                End If
'              Next iLoop3
'
'              ' If a parameter has been found then read its value.
'              ' Otherwise the expression is invalid.
'              If iParameter1Index > 0 Then
'                vParameter1 = avValues(2, iParameter1Index)
'              End If
'            End If
'
'            ' If the evaluation fails then trap the error.
'            On Error GoTo EvaluatorErrorTrap
'
'            vResult = objEvaluator.Evaluate(2, CLng(.Component.OperatorID), _
'              vParameter1, vParameter2)
'
'            On Error GoTo BasicErrorTrap
'
'            ' Update the array to reflect the evaluated operation.
'            avValues(1, iLoop2) = vbNullString
'            avValues(2, iLoop2) = vResult
'            avValues(4, iLoop2) = True
'            avValues(4, iParameter1Index) = False
'            avValues(4, iParameter2Index) = False
'
'          End If
'        End If
'      End With
'    Next iLoop2
'  Next iLoop1
'
'  For iLoop1 = 1 To UBound(avValues, 2)
'    If avValues(4, iLoop1) Then
'      vResult = avValues(2, iLoop1)
'    End If
'  Next iLoop1
'
'TidyUpAndExit:
'  ' Disassociate object variables.
'  Set objEvaluator = Nothing
'
'  ' Trim the evaluated result to the specified size.
'  EvaluateExpression = vResult
'  Exit Function
'
'EvaluatorErrorTrap:
'  ' The operator evaluation has failed.
'  Resume TidyUpAndExit
'
'BasicErrorTrap:
'  Resume TidyUpAndExit

End Function

Public Function CreateStoredProcedure() As Boolean
  ' Create the expression stored procedure.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sProcedureName As String
  Dim sSPCode As SystemMgr.cStringBuilder

  Set sSPCode = New SystemMgr.cStringBuilder

  ' Construct the expression from the database definition.
  fOK = ConstructExpression
   
  If fOK Then
  
    sProcedureName = "sp_ASRExpr_" & mlngExpressionID
    
    ' Drop any existing stored procedure with this name.
    DropProcedure sProcedureName
    
    ' Construct the stored procedure creation string.
    ' NB. The stored procedure needs to be owned by 'dbo' so that all
    ' uses can run it, and it does not get affected by permission changes.
    
    sSPCode.TheString = "/* ------------------------------------------------ */" & vbNewLine & _
     "/* Expression procedure (" & msExpressionName & "). */" & vbNewLine & _
     "/* Automatically generated by the System manager.   */" & vbNewLine & _
     "/* ------------------------------------------------ */" & vbNewLine & vbNewLine & _
     "CREATE PROCEDURE dbo." & sProcedureName & " (" & vbNewLine
    
    ' Get the expression's return type.
    Select Case miReturnType
    Case giEXPRVALUE_CHARACTER
      sSPCode.Append "    @result nvarchar(MAX) OUTPUT," & vbNewLine
    Case giEXPRVALUE_NUMERIC
      sSPCode.Append "    @result float OUTPUT," & vbNewLine
    Case giEXPRVALUE_DATE
      sSPCode.Append "    @result datetime OUTPUT," & vbNewLine
    Case giEXPRVALUE_LOGIC
      sSPCode.Append "    @result bit OUTPUT," & vbNewLine
    Case Else
      If ASRDEVELOPMENT Then
        gobjProgress.Visible = False
        MsgBox "Type Not Found Building Expression", vbCritical, "ASRDEVELOPMENT"
        Stop
      End If
    End Select
  
    ' Add the standard header code to the stored procedure.
    Select Case miExpressionType
      Case giEXPR_WORKFLOWCALCULATION
        sSPCode.Append _
          "    @piInstanceID int," & vbNewLine & _
          "    @piTempElement int)" & vbNewLine
      Case giEXPR_WORKFLOWSTATICFILTER
        sSPCode.Append _
          "    @id int," & vbNewLine & _
          "    @piInstanceID int," & vbNewLine & _
          "    @piTempElement int)" & vbNewLine
      Case Else
        sSPCode.Append _
          "    @id int)" & vbNewLine
    End Select
    
    sSPCode.Append "AS" & vbNewLine & vbNewLine & _
      "    /* Check if the nesting of the stored procedures is getting near its maximum. */" & vbNewLine & _
      "    /* NB. The maximum nesting level is 32. We stop running ASR Expression stored procedures at */" & vbNewLine & _
      "    /* 30 as they might call some other non-expression stored procedures. We leave 2 nesting levels */" & vbNewLine & _
      "    /* for these other stored procedures to run in. */" & vbNewLine & _
      "    SET NOCOUNT ON" & vbNewLine & _
      "    IF @@nestlevel >= 30 RETURN" & vbNewLine & vbNewLine & _
      "    /* Declare local variables. */" & vbNewLine & _
      "    DECLARE @hResult integer," & vbNewLine & _
      "        @source varChar(30)," & vbNewLine & _
      "        @desc varChar(200)," & vbNewLine & _
      "        @loopCounter float," & vbNewLine & _
      "        @filterRecordID integer," & vbNewLine & _
      "        @filterResult integer," & vbNewLine & _
      "        @HierarchyID integer," & vbNewLine & _
      "        @PersonnelID integer" & vbNewLine & vbNewLine

    If miExpressionType = giEXPR_WORKFLOWCALCULATION Then
      sSPCode.Append "    DECLARE @ID integer," & vbNewLine & _
        "        @iWorkflowID integer," & vbNewLine & _
        "        @iBaseTableID integer," & vbNewLine & _
        "        @iBaseRecordID integer," & vbNewLine & _
        "        @iParent1TableID integer," & vbNewLine & _
        "        @iParent1RecordID integer," & vbNewLine & _
        "        @iParent2TableID integer," & vbNewLine & _
        "        @iParent2RecordID integer," & vbNewLine & _
        "        @iRequiredTableID integer," & vbNewLine & _
        "        @fDeletedValue bit," & vbNewLine & _
        "        @sValue nvarchar(MAX)," & vbNewLine & _
        "        @fValidRecordID bit," & vbNewLine & _
        "        @iCount integer" & vbNewLine & vbNewLine
    ElseIf miExpressionType = giEXPR_WORKFLOWSTATICFILTER Then
      sSPCode.Append "    DECLARE @iWorkflowID integer," & vbNewLine & _
        "        @iBaseTableID integer," & vbNewLine & _
        "        @iBaseRecordID integer," & vbNewLine & _
        "        @iParent1TableID integer," & vbNewLine & _
        "        @iParent1RecordID integer," & vbNewLine & _
        "        @iParent2TableID integer," & vbNewLine & _
        "        @iParent2RecordID integer," & vbNewLine & _
        "        @iRequiredTableID integer," & vbNewLine & _
        "        @fDeletedValue bit," & vbNewLine & _
        "        @sValue nvarchar(MAX)," & vbNewLine & _
        "        @fValidRecordID bit," & vbNewLine & _
        "        @iCount integer" & vbNewLine & vbNewLine
    End If

    sSPCode.Append "    SET @hResult = 0" & vbNewLine & vbNewLine
      
    ' Add the expression's stored procedure code (ie. the code for the individual components).
    sSPCode.Append StoredProcedureCode("@result") & vbNewLine & vbNewLine
      
    ' Add the standard footer code to the stored procedure.
    sSPCode.Append _
      "    /* Jump over the error handling code. */" & vbNewLine & _
      "    GOTO Done" & vbNewLine & vbNewLine & _
      "Err:" & vbNewLine & _
      "    EXEC sp_OAGetErrorInfo null, @source OUT, @desc OUT" & vbNewLine & _
      "    SELECT hResult=CONVERT(binary(4),@hResult), source=@source, description=@desc" & vbNewLine & vbNewLine & _
      "Done:" & vbNewLine & _
      "    RETURN (@hResult)" & vbNewLine

    ' Run the code to create the stored procedure.
    On Error GoTo ExecutionError
    gADOCon.Execute sSPCode.ToString, , adExecuteNoRecords
  End If
  
TidyUpAndExit:
  CreateStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

ExecutionError:
  ' Flag the failure.
  fOK = False
  'MsgBox "Unable to create the stored procedure '" & msExpressionName & "'.", vbExclamation + vbOKOnly
  OutputError "Unable to create the stored procedure '" & msExpressionName & "'."

  ' Dump the failed stored procedure code into a file for debugging.
  DumpStoredProcedure (sSPCode.ToString)
  
  Resume TidyUpAndExit

End Function

Public Function CreateWorkflowUDF() As Boolean
  ' Create the filter UDF.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sFilterCode As String
  Dim sUDFName As String
  Dim UDFCode As SystemMgr.cStringBuilder
  Dim alngSourceTables() As Long
  Dim sPrefix As String

  Set UDFCode = New SystemMgr.cStringBuilder

  ' Construct the expression from the database definition.
  fOK = ConstructExpression
  
  If fOK Then
  
    sUDFName = "udf_ASRWFExpr_" & mlngExpressionID
    
    ' Drop any existing stored procedure with this name.
    DropFunction sUDFName
    
    ' Construct the runtime stored procedure creation string.
    ' NB. The stored procedure needs to be owned by 'dbo' so that all
    ' uses can run it, and it does not get affected by permission changes.

    UDFCode.TheString = "------------------------------------------------" & vbNewLine & _
      "-- Workflow expression filter (" & msExpressionName & ")." & vbNewLine & _
      "-- Automatically generated by the System Manager." & vbNewLine & _
      "------------------------------------------------" & vbNewLine & vbNewLine & _
      "CREATE FUNCTION [dbo].[" & sUDFName & "] (" & vbNewLine & _
      "    @piInstanceID integer" & vbNewLine & _
      ")" & vbNewLine & _
      "RETURNS @results TABLE (id integer)" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & _
      "    DECLARE" & vbNewLine & _
      "        @iWorkflowID integer," & vbNewLine & _
      "        @iBaseTableID integer," & vbNewLine & _
      "        @iBaseRecordID integer," & vbNewLine & _
      "        @iParent1TableID integer," & vbNewLine & _
      "        @iParent1RecordID integer," & vbNewLine & _
      "        @iParent2TableID integer," & vbNewLine & _
      "        @iParent2RecordID integer," & vbNewLine & _
      "        @dtDate datetime," & vbNewLine & _
      "        @sValue nvarchar(MAX)," & vbNewLine & _
      "        @fValidRecordID bit," & vbNewLine & _
      "        @iCount integer" & vbNewLine & vbNewLine & _
      "    SELECT @dtDate = [dbo].[udfASRGetDate]()" & vbNewLine & vbNewLine

    fOK = RuntimeFilterCode(sFilterCode, False, sPrefix)

    If fOK Then
      UDFCode.Append _
        IIf(Len(sPrefix) > 0, sPrefix & vbNewLine & vbNewLine, "") & _
        "    INSERT @results" & vbNewLine & _
        "    " & sFilterCode & vbNewLine & vbNewLine
    
      UDFCode.Append _
        "    RETURN" & vbNewLine & _
        "END"
  
      ' Run the code to create the stored procedure.
      On Error GoTo ExecutionError
      gADOCon.Execute UDFCode.ToString, , adExecuteNoRecords
    End If
  End If
  
TidyUpAndExit:
  CreateWorkflowUDF = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

ExecutionError:
  ' Flag the failure.
  fOK = False
  OutputError "Unable to create the Workflow UDF '" & msExpressionName & "'."

  ' Dump the failed stored procedure code into a file for debugging.
  DumpStoredProcedure (UDFCode.ToString)
  
  Resume TidyUpAndExit

End Function














Public Function CreateDefaultValueStoredProcedure() As Boolean
  ' Create the expression stored procedure.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As New SystemMgr.cStringBuilder
  Dim sProcedureName As String

  ' Construct the expression from the database definition.
  fOK = ConstructExpression
  
  If fOK Then
    sProcedureName = "sp_ASRDfltExpr_" & mlngExpressionID
    
    ' Drop any existing stored procedure with this name.
    DropProcedure sProcedureName
    
    ' Construct the stored procedure creation string.
    ' NB. The stored procedure needs to be owned by 'dbo' so that all
    ' uses can run it, and it does not get affected by permission changes.
    sSQL.TheString = "/* -------------------------------------------------------------------------------- */" & vbNewLine & _
      "/* Default value expression stored procedure. */" & vbNewLine & _
      "/* -------------------------------------------------------------------------------- */" & vbNewLine & _
      "CREATE PROCEDURE dbo." & Replace(sProcedureName, "'", "''") & " (" & vbNewLine
    
    ' Get the expression's return type.
    Select Case miReturnType
      Case giEXPRVALUE_CHARACTER
        sSQL.Append "    @result nvarchar(MAX) OUTPUT"
      Case giEXPRVALUE_NUMERIC
        sSQL.Append "    @result float OUTPUT"
      Case giEXPRVALUE_DATE
        sSQL.Append "    @result datetime OUTPUT"
      Case giEXPRVALUE_LOGIC
        sSQL.Append "    @result bit OUTPUT"
    End Select
  
    ' Add the parent table id parameters.
    With recRelEdit
      .Index = "idxChildParentID"
      .Seek ">=", mlngBaseTableID
    
      If Not .NoMatch Then
        Do While Not .EOF
    
          If !childID <> mlngBaseTableID Then
            Exit Do
          End If
        
          sSQL.Append "," & vbNewLine & _
            "    @id_" & Trim$(Str$(!parentID)) & " int"
                  
          .MoveNext
        Loop
      End If
    End With
    
    ' Add the standard header code to the stored procedure.
    sSQL.Append ")" & vbNewLine & _
      "AS" & vbNewLine & vbNewLine & _
      "    /* Check if the nesting of the stored procedures is getting near its maximum. */" & vbNewLine & _
      "    /* NB. The maximum nesting level is 32. We stop running ASR Expression stored procedures at */" & vbNewLine & _
      "    /* 30 as they might call some other non-expression stored procedures. We leave 2 nesting levels */" & vbNewLine & _
      "    /* for these other stored procedures to run in. */" & vbNewLine & _
      "    IF @@nestlevel >= 30 RETURN" & vbNewLine & vbNewLine & _
      "    /* Declare local variables. */" & vbNewLine & _
      "    DECLARE @hResult integer = 0," & vbNewLine & _
      "        @piInstanceID integer = 0," & vbNewLine & _
      "        @source varChar(30)," & vbNewLine & _
      "        @desc varChar(200)," & vbNewLine & _
      "        @loopCounter float," & vbNewLine & _
      "        @filterRecordID integer," & vbNewLine & _
      "        @filterResult integer" & vbNewLine & vbNewLine
      
    ' Add the expression's stored procedure code (ie. the code for the individual components).
    sSQL.Append _
      StoredProcedureCode("@result") & vbNewLine & vbNewLine
      
    ' Add the standard footer code to the stored procedure.
    sSQL.Append _
      "    /* Jump over the error handling code. */" & vbNewLine & _
      "    GOTO Done" & vbNewLine & vbNewLine & _
      "Err:" & vbNewLine & _
      "    EXEC sp_OAGetErrorInfo null, @source OUT, @desc OUT" & vbNewLine & _
      "    SELECT hResult=CONVERT(binary(4),@hResult), source=@source, description=@desc" & vbNewLine & vbNewLine & _
      "Done:" & vbNewLine & _
      "    RETURN (@hResult)" & vbNewLine

    ' Run the code to create the stored procedure.
    On Error GoTo ExecutionError
    gADOCon.Execute sSQL.ToString, , adExecuteNoRecords
  End If
  
TidyUpAndExit:
  CreateDefaultValueStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

ExecutionError:
  ' Flag the failure.
  fOK = False
  MsgBox "Unable to create the default value stored procedure '" & msExpressionName & "'.", vbExclamation + vbOKOnly
  
  ' Dump the failed stored procedure code into a file for debugging.
  DumpStoredProcedure (sSQL)
  
  Resume TidyUpAndExit

End Function














Public Function ValidityMessage(piInvalidityCode As ExprValidationCodes)
  ' Return the text nmessage that describes the given
  ' expression invalidity code.
  
  Select Case piInvalidityCode
    
    Case giEXPRVALIDATION_NOERRORS
      ValidityMessage = "No errors."
    
    Case giEXPRVALIDATION_MISSINGOPERAND
      ValidityMessage = "Missing operand."
    
    Case giEXPRVALIDATION_SYNTAXERROR
      ValidityMessage = "Syntax error."
    
    Case giEXPRVALIDATION_EXPRTYPEMISMATCH
      ValidityMessage = "Return type mismatch."
    
    Case giEXPRVALIDATION_UNKNOWNERROR
      ValidityMessage = "Unknown error."
    
    Case giEXPRVALIDATION_OPERANDTYPEMISMATCH
      ValidityMessage = "Operand type mismatch."
    
    Case giEXPRVALIDATION_PARAMETERTYPEMISMATCH
      ValidityMessage = "Parameter type mismatch."
    
    Case giEXPRVALIDATION_NOCOMPONENTS
      ValidityMessage = "The " & LCase(ExpressionTypeName) & " must have at least one component."
    
    Case giEXPRVALIDATION_PARAMETERSYNTAXERROR
      ValidityMessage = "Function parameter syntax error."
    
    Case giEXPRVALIDATION_PARAMETERNOCOMPONENTS
      ValidityMessage = "The function parameter expression must have at least one component."
    
    Case giEXPRVALIDATION_FILTEREVALUATION
      ' JPD20020419 Fault 3687
      'ValidityMessage = "Logic components must be compared with other logic components."
      ValidityMessage = "Error creating SQL runtime code."
    
    ' JPD20020419 Fault 3687
    Case giEXPRVALIDATION_SQLERROR
      ValidityMessage = "The complexity of the " & LCase(ExpressionTypeName) & " has caused the following SQL error : " & vbNewLine & vbNewLine & _
        "'" & msErrorMessage & "'" & vbNewLine & vbNewLine & _
        "Try simplifying the " & LCase(ExpressionTypeName) & "."
    
    ' JPD20020419 Fault 3687
    Case giEXPRVALIDATION_ASSOCSQLERROR
      ValidityMessage = "The complexity of this " & LCase(ExpressionTypeName) & " would cause an expression that uses this " & LCase(ExpressionTypeName) & " to suffer from the following SQL error : " & vbNewLine & vbNewLine & _
        "'" & msErrorMessage & "'" & vbNewLine & vbNewLine & _
        "Try simplifying this " & LCase(ExpressionTypeName) & "."

    Case giEXPRVALIDATION_INVALIDRECORDIDENTIFICATION
      ValidityMessage = "Invalid function record identification."
    
    Case Else
      ValidityMessage = "The function parameter expression must have at least one component."
    
  End Select

End Function



Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the expression's return type.
  ReturnType = miReturnType
  
End Property

Public Property Let ReturnType(piNewValue As ExpressionValueTypes)
  ' Set the expression's return type.
  miReturnType = piNewValue
  
End Property

Private Sub DumpStoredProcedure(psDumpString As String)
  ' Dump the given string to a file for debugging.
  
  ' Open the log file.
  Open "OpenHR Stored Procedure.log" For Output Lock Write As #1 ' Close before reopening in another mode.
  
  ' Write the given string to the log file.
  If LenB(psDumpString) <> 0 Then
    Print #1, psDumpString
  Else
    Print #1, "No string given"
  End If
  
  ' Close the dump log file.
  Close #1
  
End Sub

Public Property Get Components() As Collection
  ' Return the component collection.
  Set Components = mcolComponents
  
End Property

Public Property Set Components(ByVal pCollection As Collection)
  ' Set the component collection.
  Set mcolComponents = pCollection

End Property

Public Function ParentExpressionName() As String
  ' Return the name of the parent expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sName As String
  Dim objParentComponent As CExprComponent
  
  fOK = True
  
  If mlngParentComponentID > 0 Then
    Set objParentComponent = New CExprComponent
    With objParentComponent
      .ComponentID = mlngParentComponentID
      If .ReadComponentDetails Then
        sName = .ParentExpressionName
      Else
        fOK = False
      End If
    End With
    Set objParentComponent = Nothing
  Else
    sName = msExpressionName
  End If
  
TidyUpAndExit:
  Set objParentComponent = Nothing
  If fOK Then
    ParentExpressionName = sName
  Else
    ParentExpressionName = "<unknown>"
  End If
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get Owner() As String
  ' Return the expression owner.
  Owner = msOwner

End Property

Public Property Let Owner(ByVal psNewValue As String)
  ' Set the expression owner.
  msOwner = psNewValue

End Property

'Public Property Get Access() As AccessCodes
Public Property Get Access() As String
  ' Return the access code.
  Access = msAccess
  
End Property

Public Property Let Access(ByVal psNewValue As String)
  ' Set the access code.
  msAccess = psNewValue

End Property

Public Function CopyExpressionToClipboard() As Boolean

  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim fOK As Boolean
  Dim sTableName As String
  Dim objComponent As CExprComponent
  Dim rsInfo As DAO.Recordset
  Dim msClipboardText As String
  
  fOK = True
  
  sSQL = "SELECT tableName" & _
    " FROM tmpTables" & _
    " WHERE tmpTables.tableID = " & Trim$(Str$(mlngBaseTableID))
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsInfo
    If Not (.EOF And .BOF) Then
      sTableName = !TableName
    Else
      sTableName = "<unknown>"
    End If
  
    .Close
  End With
  Set rsInfo = Nothing
   
  ' Send details to the clipboard
  msClipboardText = ExpressionTypeName & " Definition : " & Trim(msExpressionName) & vbNewLine
  msClipboardText = msClipboardText & "Description : " & msDescription & vbNewLine & vbNewLine
  msClipboardText = msClipboardText & "Owner : " & msOwner & vbNewLine
  msClipboardText = msClipboardText & "Access : " & AccessDescription(msAccess)
  'Select Case msAccess
  '  Case "RW": msClipboardText = msClipboardText & "Access : Read / Write"
  '  Case "RO": msClipboardText = msClipboardText & "Access : Read only"
  '  Case "HD": msClipboardText = msClipboardText & "Access : Hidden"
  'End Select
  msClipboardText = msClipboardText & vbNewLine & vbNewLine & "Base Table : " & Trim(sTableName) & vbNewLine & vbNewLine
  
  msClipboardText = msClipboardText & "Components" & vbNewLine & vbNewLine
  
  ' Set the clipboard
  Clipboard.SetText Clipboard.GetText & msClipboardText
  
  ' Copy the components to the clipboard
  For Each objComponent In mcolComponents
    fOK = objComponent.CopyToClipboard(1)
  Next
  Set objComponent = Nothing
  
TidyUpAndExit:
  Set objComponent = Nothing
  If Not fOK Then
    MsgBox "Unable to send the " & ExpressionTypeName & " '" & Name & "' to the clipboard." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  CopyExpressionToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As CExprComponent
  Dim sClipboardMessage As String

  fOK = True
  sClipboardMessage = String(piLevel, vbTab) & "Parameter : " & ComponentDescription & vbNewLine
  
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
  ' Copy the function's parameter expressions to clipboard
  For Each objComponent In mcolComponents
    fOK = objComponent.Component.CopyToClipboard(piLevel + 1)
  Next objComponent

TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function RootExpressionName() As String
  ' Return the name of the root expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objComp As CExpression
  Dim rsExpressions As DAO.Recordset
  Dim sName As String

If mlngParentComponentID = 0 Then
  sName = msExpressionName
  fOK = True
Else

  sSQL = "SELECT tmpExpressions.name, tmpExpressions.parentComponentID, tmpExpressions.exprID" & _
    " FROM tmpExpressions, tmpComponents" & _
    " WHERE tmpComponents.componentID = " & Trim$(Str$(mlngParentComponentID)) & _
    " AND tmpExpressions.exprID = tmpComponents.exprID"

  Set rsExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsExpressions
    If Not (.EOF And .BOF) Then
      If !ParentComponentID = 0 Then
        sName = !Name
      Else
    
        ' If the parent expression is not a top-level expression then
        ' find the parent expression's parent expression. Confused yet ?
        Set objComp = New CExpression
        objComp.ExpressionID = !ExprID
        objComp.ReadExpressionDetails
        sName = objComp.RootExpressionName
        Set objComp = Nothing
    
      End If
      fOK = True
    Else
      fOK = False
    End If
  
    .Close
  End With
  
End If
  

TidyUpAndExit:
  Set objComp = Nothing
  RootExpressionName = IIf(fOK, sName, "<unknown>")
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

'Public Function CalculateBottomOfPage() As Long
'With Printer
'    CalculateBottomOfPage = .ScaleHeight - (giPRINT_YINDENT)
'End With
'End Function
'Public Function CheckEndOfPage2(Optional mlngBottom As Long, Optional fReset As Boolean)
'  If Printer.CurrentY > mlngBottom Then
'    Call FooterText2
'    Printer.NewPage
'
'    If fReset Then glngPageNum = 0
'
'    Printer.CurrentY = giPRINT_YINDENT
'    Printer.CurrentX = giPRINT_XINDENT
'  End If
'End Function
'Public Function FooterText2()
'
'  Dim strPageNum As String
'
'  glngPageNum = glngPageNum + 1
'  strPageNum = "Page " & CStr(glngPageNum)
'
'  Printer.FontSize = 8
'  Printer.Print " "
'  Printer.FontBold = False
'
'  Printer.CurrentX = giPRINT_XINDENT
'  Printer.Print "Printed on " & Format(Now, DateFormat) & _
'                " at " & Format(Now, "hh:nn") & " by " & gsUserName;
'
'  Printer.CurrentX = (Printer.ScaleWidth - giPRINT_XINDENT) - Printer.TextWidth(strPageNum)
'  Printer.Print strPageNum
'
'  Printer.FontSize = 10
'
'End Function


Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mcolComponents = Nothing
  Set mobjBadComponent = Nothing
  Set mobjBaseComponent = Nothing

End Sub


'MH20060727
'Private Sub Class_Terminate()
'  Terminate
'End Sub
'
'Public Sub Terminate()
'
'  If Not mcolComponents Is Nothing Then
'    Do While mcolComponents.Count > 0
'      mcolComponents(1).Terminate
'      mcolComponents.Remove 1
'    Loop
'    Set mcolComponents = Nothing
'  End If
'
'  Set mobjBadComponent = Nothing
'  Set mobjBaseComponent = Nothing
'  Set mfrmUse = Nothing
'
'End Sub

Public Property Get UtilityID() As Long
  UtilityID = mlngUtilityID
  
End Property

Public Property Let UtilityID(ByVal plngNewValue As Long)
  mlngUtilityID = plngNewValue
  
End Property

Public Property Get UtilityBaseTable() As Long
  UtilityBaseTable = mlngUtilityBaseTable

End Property

Public Property Let UtilityBaseTable(ByVal plngNewValue As Long)
  mlngUtilityBaseTable = plngNewValue

End Property

Public Property Get WorkflowInitiationType() As WorkflowInitiationTypes
  WorkflowInitiationType = miWorkflowInitiationType

End Property

Public Property Let WorkflowInitiationType(ByVal piNewValue As WorkflowInitiationTypes)
  miWorkflowInitiationType = piNewValue
  
End Property


Public Function ShowExpressionUsage() As Boolean
    'NPG20080229 Fault 12944
    'Returns true if expression is used (false if not!)
    'Also, shows the usage form if expression is used

    Dim blnIsUsed As Boolean

    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    
    blnIsUsed = ExpressionIsUsed(mlngExpressionID)
    
    If blnIsUsed Then
      Screen.MousePointer = vbDefault
      mfrmUse.ShowMessage Me.Name & " " & Me.ExpressionTypeName, "The return type of this " & LCase(Me.ExpressionTypeName) & " cannot be modified as it is used by the following:", Calculation
    End If
    
    UnLoad mfrmUse
    Set mfrmUse = Nothing

    ShowExpressionUsage = blnIsUsed

End Function

' Code stub for .NET migration
Public Function GetSQLElementCode() As String
  GetSQLElementCode = vbNullString
End Function
