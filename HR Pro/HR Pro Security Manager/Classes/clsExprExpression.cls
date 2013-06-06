VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Expression definition variables.
Private mlngExpressionID As Long
Private msExpressionName As String
Private mlngBaseTableID As Long
Private miReturnType As Integer
Private miExpressionType As Integer
Private mlngParentComponentID As Long
Private msOwner As String
Private msAccess As String
Private msDescription As String
Private mlngTimeStamp As Long
Private msBaseTableName As String
Private mbViewInColour As Boolean
Private mbExpandedNode As Boolean
Private mlngActionType As Long

Public mfDontUpdateTimeStamp As Boolean

' Class handling variables.
Private mfConstructed As Boolean
Private mcolComponents As Collection
Private mobjBadComponent As clsExprComponent
Private mobjBaseComponent As clsExprComponent

Private malngColumns() As Long

Private msErrorMessage As String

Public Function HiddenElements() As Boolean
  ' Checks the expression for hidden components.
  On Error GoTo ErrorTrap

  Dim blnHidden As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim rsCheck As New ADODB.Recordset
  Dim objComponent As clsExprComponent

  blnHidden = False
  
  ' Loop through each component in the expression
  For Each objComponent In mcolComponents
    ' If its a calculation or filter
    Select Case objComponent.ComponentType
      Case giCOMPONENT_CALCULATION
        sSQL = "SELECT ExprID, Username, Access FROM AsrSysExpressions WHERE ExprID = " & objComponent.Component.CalculationID
        rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        blnHidden = (rsCheck!Access = ACCESS_HIDDEN)
        rsCheck.Close
        
      Case giCOMPONENT_FILTER
        sSQL = "SELECT ExprID, Username, Access FROM AsrSysExpressions WHERE ExprID = " & objComponent.Component.FilterID
        rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        blnHidden = (rsCheck!Access = ACCESS_HIDDEN)
        rsCheck.Close
      
      Case giCOMPONENT_FIELD
        If objComponent.Component.SelectionFilterID > 0 Then
          sSQL = "SELECT ExprID, Username, Access FROM AsrSysExpressions WHERE ExprID = " & objComponent.Component.SelectionFilterID
          rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
          blnHidden = (rsCheck!Access = ACCESS_HIDDEN)
          rsCheck.Close
        End If
    
      Case giCOMPONENT_FUNCTION
        blnHidden = objComponent.Component.HiddenElements
    End Select
    
    If blnHidden Then Exit For
  Next objComponent

  HiddenElements = blnHidden

TidyUpAndExit:
  Set rsCheck = Nothing
  Exit Function

ErrorTrap:
  MsgBox "Error validating expression (checking for hidden components)." & vbCrLf & Err.Description, _
    vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Function



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
  piOpFuncID As Integer) As clsExprComponent
  ' JPD20021108 Fault 3287
  ' Add a new component to the expression.
  ' Returns the new component object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As clsExprComponent

  ' Instantiate a component object.
  Set objComponent = New clsExprComponent
  
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
Public Property Get ComponentDescription() As String
  ' Return the expression name.
  ComponentDescription = msExpressionName
  
End Property

Public Sub ResetConstructedFlag(fValue)

  mfConstructed = fValue
  
End Sub


Public Property Get ExpressionID() As Long
  ' Return the expression ID.
  ExpressionID = mlngExpressionID
  
End Property

Public Property Let ExpressionID(ByVal plngNewValue As Long)
  ' Set the expression ID.
  If mlngExpressionID <> plngNewValue Then
    mlngExpressionID = plngNewValue
    mfConstructed = False
  End If
  
End Property

Public Property Get BaseTableID() As Long
  ' Return the expressions base table ID.
  BaseTableID = mlngBaseTableID
  
End Property

Public Property Let BaseTableID(ByVal plngNewValue As Long)
  ' Set the expression base table property.
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
  
  If mlngBaseTableID <> plngNewValue Then
    mlngBaseTableID = plngNewValue

    ' Read the parent table name.
    sSQL = "SELECT tableName" & _
      " FROM ASRSysTables" & _
      " WHERE tableID = " & Trim(Str(mlngBaseTableID))
      
    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    If Not (rsInfo.EOF And rsInfo.BOF) Then
      msBaseTableName = rsInfo!TableName
    End If
            
    rsInfo.Close
    Set rsInfo = Nothing
  End If

End Property

Public Property Get ReturnType() As Integer
  ' Return the expression's return type.
  ReturnType = miReturnType
  
End Property

Public Property Let ReturnType(ByVal piNewValue As Integer)
  ' Set the expression's return type.
  miReturnType = piNewValue

End Property

Public Property Get ExpressionType() As Integer
  ' Return the expression's parent type property.
  ExpressionType = miExpressionType
  
End Property

Public Property Let ExpressionType(ByVal piNewValue As Integer)
  ' Set the expression's type property.
  miExpressionType = piNewValue
  
End Property

Public Property Let ColumnList(palngColumns As Variant)
  malngColumns = palngColumns
  
End Property



Public Property Get ColumnList() As Variant
  ColumnList = malngColumns
  
End Property




Public Function EvaluateExpression(palngColumnValues As Variant, pvResult As Variant) As Boolean
  ' Evaluate the expression with values from the column value array. Return the evaluate result
  ' NB. Assumes the expression is already constructed.
  On Error GoTo BasicErrorTrap

  Dim fOK As Boolean
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim avValues() As Variant
  Dim vParameter1 As Variant
  Dim vParameter2 As Variant
  Dim iParameter1Index As Integer
  Dim iParameter2Index As Integer
  Dim iMinOperatorPrecedence As Integer
  Dim iMaxOperatorPrecedence As Integer
  Dim vEvaluatedValue As Variant
  
  iMinOperatorPrecedence = -1
  iMaxOperatorPrecedence = -1

  fOK = True
  
  ' Create an array of the components in the expression.
  ' Column 1 = operator id.
  ' Column 2 = evaluated value.
  ' Column 3 = evaluated flag.
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
        avValues(3, iLoop1) = False
      Else
        ' Write the component's evaluated value into the array.
        fOK = .EvaluateComponent(palngColumnValues, vEvaluatedValue)
        
        If fOK Then
          avValues(2, iLoop1) = vEvaluatedValue
          avValues(3, iLoop1) = True
        Else
          Exit For
        End If
      End If
    End With
  Next iLoop1

  If fOK Then
    ' Loop throught the expression's components.
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
  
              For iLoop3 = iLoop2 + 1 To UBound(avValues, 2)
                ' Read the index of the parameter's value if there is one.
                If avValues(3, iLoop3) Then
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
  
                ' Read the value that precedes the current operator.
                iParameter1Index = 0
                For iLoop3 = iLoop2 - 1 To 1 Step -1
                  ' Read the index of the parameter's value if there is one.
                  If avValues(3, iLoop3) Then
                    iParameter1Index = iLoop3
                    Exit For
                  End If
                Next iLoop3
  
                If iParameter1Index > 0 Then
                  vParameter1 = avValues(2, iParameter1Index)
                End If
              End If
  
              ' If the evaluation fails then trap the error.
              fOK = .Component.EvaluateComponent(vParameter1, vParameter2, vEvaluatedValue)
          
              If fOK Then
                ' Update the array to reflect the evaluated operation.
                avValues(1, iLoop2) = vbNullString
                avValues(2, iLoop2) = vEvaluatedValue
                avValues(3, iLoop2) = True
                avValues(3, iParameter1Index) = False
                avValues(3, iParameter2Index) = False
              Else
                Exit For
              End If
            End If
          End If
        End With
      Next iLoop2
      
      If Not fOK Then
        Exit For
      End If
    Next iLoop1
  
    If fOK Then
      For iLoop1 = 1 To UBound(avValues, 2)
        If avValues(3, iLoop1) Then
          vEvaluatedValue = avValues(2, iLoop1)
        End If
      Next iLoop1
    End If
  End If
  
TidyUpAndExit:
  ' Trim the evaluated result to the specified size.
  If fOK Then
    pvResult = vEvaluatedValue
  End If
  EvaluateExpression = fOK
  Exit Function

BasicErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Private Sub Class_Initialize()
  ' Create a new collection to hold the expression's components.
  Set mcolComponents = New Collection
  mfConstructed = False
  mbExpandedNode = False
  
End Sub


Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mcolComponents = Nothing
  Set mobjBadComponent = Nothing
  Set mobjBaseComponent = Nothing

End Sub



Public Property Get Name() As String
  ' Return the expression name.
  If Not mfConstructed Then
    ConstructExpression
  End If
  
  Name = msExpressionName
  
End Property

Public Property Let Name(ByVal psNewValue As String)
  ' Set the expression name.
  msExpressionName = psNewValue

End Property

Public Function DeleteComponent(pobjComponent As clsExprComponent) As Boolean
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

Public Function AddComponent() As clsExprComponent
  ' Add a new component to the expression.
  ' Returns the new component object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As clsExprComponent

  ' Instantiate a component object.
  Set objComponent = New clsExprComponent
  
  ' Initialse the new component's properties.
  Set objComponent.ParentExpression = Me
  
  ' Get the new component to handle its own definition.
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
  ' Disassociate object variables.
  Set objComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property


Public Function SelectExpression(pfLockTable As Boolean, Optional plngOptions As Long) As Boolean
  ' Display the expression selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fExit As Boolean
  Dim lngOldExpressionID As Long
  Dim sSQL As String
  Dim rsExpressions As New ADODB.Recordset
  Dim frmSelectExpr As frmDefSel
  Dim strSQL As String

  fOK = (mlngBaseTableID > 0) Or (Not pfLockTable)

  If fOK Then
    Set frmSelectExpr = New frmDefSel

    fExit = False
    lngOldExpressionID = mlngExpressionID

    With frmSelectExpr
       
        ' Set the selection form properties.
        '.LockTable = pfLockTable

        'Set .Expression = Me

        ' Set the caption if necessary - only really needed because by default the
        ' caption is SELECT XXXXX, but when the user chooses CALCULATIONS from the
        ' utility menu, they arent really selecting one, so using this, we can change
        ' the form caption to Maintain Calculations or something like that !
        'If psCaption <> "" Then .Caption = psCaption


        strSQL = " type = " & CStr(ExpressionType) & _
                 " AND (returnType = " & CStr(ReturnType) & _
                      " OR type = " & CStr(giEXPR_RUNTIMECALCULATION) & ")" & _
                 " AND parentComponentID = 0"

        .SelectedID = mlngExpressionID
      
      ' Loop until an expression is selected, deselected, or the selection
      ' form is cancelled.
      Do While Not fExit

        .TableID = BaseTableID
        .TableComboEnabled = (Not pfLockTable)
        .TableComboVisible = True
        
        If plngOptions = 0 Then
          .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtSelect + edtDeselect + edtProperties
        Else
          .Options = plngOptions
        End If
        
        .EnableRun = False
        
        Select Case ExpressionType
          Case giEXPR_RUNTIMEFILTER
            .ShowList "FILTERS", strSQL
  
          Case giEXPR_RUNTIMECALCULATION
            .ShowList "CALCULATIONS", strSQL
        End Select

        ' Display the selection form.
        .Show vbModal
        BaseTableID = .TableID

        ' Execute the required operation.
        Select Case .Action
          ' Add a new expression.
          Case edtAdd
            NewExpression
            .SelectedID = ExpressionID

          ' Edit the selected expression.
          Case edtEdit
            ExpressionID = .SelectedID
            
            'JPD 20030818 Fault 6555
            mfConstructed = False
                        
            If .FromCopy Then
              CopyExpression
              If ExpressionID > 0 Then
                .SelectedID = ExpressionID
              End If
            Else
              EditExpression
            End If
          
          Case edtDelete
            ExpressionID = .SelectedID
            
            DeleteExpression
          
          ' Copy the selected expression.
          Case edtCopy
            ExpressionID = .SelectedID
            CopyExpression

          ' Print the selected expression.
          Case edtPrint
            ExpressionID = .SelectedID
            PrintExpression

          ' Select the selected expression and return to the calling form.
          Case edtSelect
            ExpressionID = .SelectedID

            ' Check that the selection is valid.
            If ValidateSelection Then
              fOK = True
              fExit = True
            End If

          ' Deselect all expressions and return to the calling form.
          Case edtDeselect
            ExpressionID = 0
            msExpressionName = ""
            fOK = True
            fExit = True

          ' Exit the selection form and return to the calling form.
          Case edtCancel
            ' Check if the original expression still exists.
            sSQL = "SELECT name" & _
              " FROM ASRSysExpressions" & _
              " WHERE exprID = " & Trim(Str(lngOldExpressionID))
            rsExpressions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
            
            If rsExpressions.EOF And rsExpressions.BOF Then
              ExpressionID = 0
              msExpressionName = ""
            Else
              ExpressionID = lngOldExpressionID
              msExpressionName = rsExpressions!Name
            End If

            rsExpressions.Close
            fOK = False
            fExit = True
        End Select
      Loop
    End With

    Set frmSelectExpr = Nothing
  End If

TidyUpAndExit:
  SelectExpression = fOK
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
  Dim frmEdit As frmExpression
  
  'TM20010801 Fault 2617
  'Set up ActionType, to be used in the frmExpression constructor.
  Me.ActionType = edtCopy
  
  fOK = False
  
  ' Construct the expression to copy.
  Screen.MousePointer = vbHourglass
  fOK = ConstructExpression
  Screen.MousePointer = vbDefault
  
  If fOK Then
    ' Initialise the copied expression's name.
    sName = msExpressionName
    msExpressionName = "Copy_of_" & Trim(sName)

    ' Display the expression edit form.
    Set frmEdit = New frmExpression
    Set frmEdit.Expression = Me
    
    If Not frmEdit.Cancelled Then
      mlngExpressionID = 0
      
      'MH15062000
      msOwner = gsUserName
  
      frmEdit.Show vbModal
    End If
    
    ' If the changes were confirmed then write the changes to the database.
    fOK = Not frmEdit.Cancelled
    
    If fOK Then
      fOK = WriteExpression_Transaction
    
      ' JDM - 26/11/01 - Fault 3202 - Stores details about this expression
      Select Case miExpressionType
        Case giEXPR_RUNTIMECALCULATION
          Call UtilCreated(utlCalculation, Me.ExpressionID)
        Case giEXPR_RUNTIMEFILTER
          Call UtilCreated(utlFilter, Me.ExpressionID)
      End Select
      
    Else
      mfConstructed = False
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

Public Sub QuickCopyExpression(Optional psName As String)
  ' Create a copy of the expression without displaying the edit form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sName As String
  
  'TM20010801 Fault 2617
  'Set up ActionType, to be used in the frmExpression constructor.
  Me.ActionType = edtCopy
  
  fOK = False
  
  ' Construct the expression to copy.
  fOK = ConstructExpression
  
  If fOK Then
    ' Initialise the copied expression's name.
    If IsMissing(psName) Then
      sName = msExpressionName
      msExpressionName = "Copy_of_" & Trim(sName)
    Else
      msExpressionName = psName
    End If
    
    mlngExpressionID = 0
      
    'MH15062000
    msOwner = gsUserName
  End If
    
  If fOK Then
    fOK = WriteExpression_Transaction
  
    ' JDM - 26/11/01 - Fault 3202 - Stores details about this expression
    Select Case miExpressionType
      Case giEXPR_RUNTIMECALCULATION
        Call UtilCreated(utlCalculation, Me.ExpressionID)
      Case giEXPR_RUNTIMEFILTER
        Call UtilCreated(utlFilter, Me.ExpressionID)
    End Select
  
  Else
    mfConstructed = False
  End If
    
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Public Function CopyComponent() As clsExprExpression
  ' Copies the selected component and all of it's children.

  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim fOK As Boolean
  Dim objCopyComponent As New clsExprExpression

    fOK = True
    objCopyComponent.ResetConstructedFlag (True)
    objCopyComponent.Name = msExpressionName
    objCopyComponent.BaseTableID = BaseTableID
    objCopyComponent.ExpressionType = ExpressionType
    objCopyComponent.ReturnType = ReturnType
    objCopyComponent.ColumnList = ColumnList

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

Public Function DeleteExpression(Optional pfIntransaction As Boolean) As Boolean
  ' Delete the expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fUsed As Boolean
  Dim fInTransaction As Boolean
  Dim sSQL As String
  Dim frmExpr As frmExpression
  
  If IsMissing(pfIntransaction) Then
    pfIntransaction = True
  End If
  
  fUsed = False
  fInTransaction = False
  
  Me.ActionType = edtDelete
  
  ' Construct the expression from the database definition.
  Screen.MousePointer = vbHourglass
  fOK = ConstructExpression
  Screen.MousePointer = vbDefault

  'TM20010801 Fault 2617
  'Create expression object so that the frmExpression constructor can check the Access
  'and the ownership of the selected expression and handle accordingly.
  If fOK Then
    ' Display the expression edit form.
    Set frmExpr = New frmExpression
    Set frmExpr.Expression = Me
    fOK = Not frmExpr.Cancelled
    Set frmExpr = Nothing
  End If
  
  If Not fOK Then
    Exit Function
  End If
  
  If fOK Then
    If pfIntransaction Then
      fInTransaction = True
      ' Begin the transaction of data.
      gADOCon.BeginTrans
    End If
    
    ' Check that the expression can be deleted.
    ' ie. is not used anywhere.
    fUsed = ExpressionIsUsed
    
    If Not fUsed Then
      ' Delete the expression's components.
      fOK = DeleteExistingComponents
      
      ' Delete the expression itself.
      If fOK Then
        sSQL = "DELETE FROM ASRSysExpressions" & _
          " WHERE exprID = " & Trim(Str(mlngExpressionID))
        gADOCon.Execute sSQL, , adCmdText
      End If
    End If
  End If

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK And (Not fUsed) Then
    If fInTransaction Then
      gADOCon.CommitTrans
    End If
  Else
    If (Not fUsed) Then
      MsgBox "Error deleting the expression.", _
        vbOKOnly, App.ProductName
    End If
    If fInTransaction Then
      gADOCon.RollbackTrans
    End If
  End If
  DeleteExpression = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ExpressionIsUsed() As Boolean
  ' Return true if the expression is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Expressions may be used in the following contexts :
  '
  ' Table definitions - each table can have a record description defined.
  ' Column definitions - calculated columns require a calculation to be defined. All columns can have a validation check defined.
  ' Views - A filter can be defined for which records can be seen through a view.
  ' Expression definitions - expressions can refer to columns in child tables of the expression's base table. A filter for which child records are read can be defined. Expressions can also refer directly to other expressions.
  ' Cross-Tabs - A filter can be defined for which records are included in the cross-tab.
  ' Custom Reports - Filters can be defined for which records are included in the report, from the report's base table, and any parent or child tables that are also referred to. Custom reports can include calculations.
  ' Data Transfer - A filter can be defined for which records are included in the data transfer.
  ' Export - Filters can be defined for which records are included in the export, from the export's base table, and any parent or child tables that are also referred to. Exports can include calculations.
  ' Global Functions - A filter can be defined for which records are included in the global function. Global functions can include calculations.
  ' Mail Merge - A filter can be defined for which records are included in the mail merge. Mail merge definitions can include calculations.
  ' Record Profile - A filter can be defined for which records are included in the record profile.
  ' Calendar Report - A filter can be defined for which records are included in the Calendar Report.
  
  On Error GoTo ErrorTrap
  
  Dim sExprName As String
  Dim sExprParentTable As String
  Dim sExprType As String
  Dim sGlobalFunctionType As String
  Dim objComp As clsExprComponent
  Dim lngRootExprID As Long
  Dim fUsed As Boolean
  Dim rsCheck As New ADODB.Recordset
  Dim sSQL As String
  Dim sType As String
  
  fUsed = False
  
  ' Check that the expression is not used as a record description for a table.
  If Not fUsed Then
    sSQL = "SELECT ASRSysTables.tableName" & _
      " FROM ASRSysTables" & _
      " WHERE ASRSysTables.recordDescExprID = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    With rsCheck
  
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as the record description for the '" & .Fields("tableName") & "' table.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Find any columns that use this expression as a Calculation,
    ' a Got Focus clause, or a Lost Focus clause.
    sSQL = "SELECT ASRSysColumns.columnName, ASRSysTables.tableName" & _
      " FROM ASRSysColumns" & _
      " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
      " WHERE (ASRSysColumns.calcExprID = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSysColumns.lostFocusExprID = " & Trim(Str(mlngExpressionID)) & ")"
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used by the '" & !ColumnName & "' column in the '" & !TableName & "' table.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Find any views that use this expression as the view filter.
    sSQL = "SELECT ASRSysViews.viewName, ASRSysTables.tableName" & _
      " FROM ASRSysViews" & _
      " INNER JOIN ASRSysTables ON ASRSysViews.viewTableID = ASRSysTables.tableID" & _
      " WHERE ASRSysViews.expressionID = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as the filter for the '" & !ViewName & "' view, based on the '" & !TableName & "' table.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that it is not used as a calculation in another expression,
    ' or as the filter in another expression.
    sSQL = "SELECT componentID" & _
      " FROM ASRSysExprComponents" & _
      " WHERE (calculationID = " & Trim(Str(mlngExpressionID)) & ")" & _
      " OR (fieldSelectionFilter = " & Trim(Str(mlngExpressionID)) & _
      "   AND type = " & Trim(Str(giCOMPONENT_FIELD)) & ")"
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
      
      If fUsed Then
        Set objComp = New clsExprComponent
        objComp.ComponentID = !ComponentID
        lngRootExprID = objComp.RootExpressionID
      End If
      
      .Close
    End With
        
    If fUsed Then
      sExprName = "<unknown>"
      sExprType = "<unknown>"
      sExprParentTable = "<unknown>"
      
      ' Get the expression definition.
      sSQL = "SELECT ASRSysExpressions.name, ASRSysTables.tableName, ASRSysExpressions.type" & _
        " FROM ASRSysExpressions" & _
        " INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & _
        " WHERE ASRSysExpressions.exprID = " & Trim(Str(lngRootExprID))
      rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      With rsCheck
        If Not (.EOF And .BOF) Then
          sExprName = !Name
          sExprParentTable = !TableName
          sExprType = LCase(ExpressionTypeName(!Type))
        End If
      
        .Close
      End With
      
      ' Tell the user why the expression cannot be deleted.
      MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
        "It is used by the " & sExprType & " '" & sExprName & "'," & vbCr & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.ProductName
    End If
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Cross-Tab.
    sSQL = "SELECT name" & _
      " FROM ASRSysCrossTab" & _
      " WHERE filterID = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as the filter for the '" & !Name & "' cross-tab.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Custom Report.
    ' JPD20021112 Fault 4735
    'sSQL = "SELECT name" & _
      " FROM ASRSysCustomReportsName" & _
      " WHERE filter = " & Trim(Str(mlngExpressionID)) & _
      " OR parent1Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR parent2Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR childFilter = " & Trim(Str(mlngExpressionID))
    sSQL = "SELECT ASRSysCustomReportsName.name" & _
      " FROM ASRSysCustomReportsName" & _
      " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
      " WHERE ASRSysCustomReportsName.filter = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSysCustomReportsName.parent1Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSysCustomReportsName.parent2Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSYSCustomReportsChildDetails.childFilter = " & Trim(Str(mlngExpressionID))

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a filter in the '" & !Name & "' custom report.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a calculation in a Custom Report.
    sSQL = "SELECT ASRSysCustomReportsName.name" & _
      " FROM ASRSysCustomReportsDetails" & _
      " INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsDetails.customReportID = ASRSysCustomReportsName.ID" & _
      " WHERE colExprID = " & Trim(Str(mlngExpressionID)) & _
      " AND UPPER(type) = 'E'"

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a calculation in the '" & !Name & "' custom report.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Data Transfer.
    sSQL = "SELECT name" & _
      " FROM ASRSysDataTransferName" & _
      " WHERE filterID = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as the filter for the '" & !Name & "' data transfer.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a filter for an Export.
    sSQL = "SELECT name" & _
      " FROM ASRSysExportName" & _
      " WHERE filter = " & Trim(Str(mlngExpressionID)) & _
      " OR parent1Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR parent2Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR childFilter = " & Trim(Str(mlngExpressionID))

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a filter in the '" & !Name & "' export.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a calculation in an Export.
    sSQL = "SELECT ASRSysExportName.name" & _
      " FROM ASRSysExportDetails" & _
      " INNER JOIN ASRSysExportName ON ASRSysExportDetails.exportID = ASRSysExportName.ID" & _
      " WHERE colExprID = " & Trim(Str(mlngExpressionID)) & _
      " AND UPPER(type) = 'X'"

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a calculation in the '" & !Name & "' export.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Global function.
    sSQL = "SELECT name, UPPER(type)" & _
      " FROM ASRSysGlobalFunctions" & _
      " WHERE filterID = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        Select Case !Type
          Case "U"
            sGlobalFunctionType = "update"
          Case "A"
            sGlobalFunctionType = "add"
          Case "D"
            sGlobalFunctionType = "delete"
          Case Else
            sGlobalFunctionType = "function"
        End Select
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as the filter for the '" & !Name & "' global " & sGlobalFunctionType & ".", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  ' Check that the expression is not used as a calculation for a Global function.
  If Not fUsed Then
    sSQL = "SELECT ASRSysGlobalFunctions.name, UPPER(ASRSysGlobalFunctions.type)" & _
      " FROM ASRSysGlobalItems" & _
      " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID" & _
      " WHERE ASRSysGlobalItems.exprID = " & Trim(Str(mlngExpressionID)) & _
      " AND ASRSysGlobalItems.valueType = 4"
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        Select Case !Type
          Case "U"
            sGlobalFunctionType = "update"
          Case "A"
            sGlobalFunctionType = "add"
          Case "D"
            sGlobalFunctionType = "delete"
          Case Else
            sGlobalFunctionType = "function"
        End Select
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a calculation in the '" & !Name & "' global " & sGlobalFunctionType & ".", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Mail Merge.
    sSQL = "SELECT name" & _
      " FROM ASRSysMailMergeName" & _
      " WHERE filterID = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as the filter for the '" & !Name & "' mail merge.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  ' Check that the expression is not used as a calculation for a Mail Merge.
  If Not fUsed Then
    sSQL = "SELECT ASRSysMailMergeName.name" & _
      " FROM ASRSysMailMergeColumns" & _
      " INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeColumns.mailMergeID = ASRSysMailMergeName.mailMergeID" & _
      " WHERE ASRSysMailMergeColumns.columnID = " & Trim(Str(mlngExpressionID)) & _
      " AND UPPER(ASRSysMailMergeColumns.type) = 'E'"
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a calculation in the '" & !Name & "' mail merge.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If
  
  'record profile
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Record Profile.
    sSQL = "SELECT ASRSysRecordProfileName.name" & _
      " FROM ASRSysRecordProfileName" & _
      " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
      " WHERE ASRSysRecordProfileName.filterID = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSYSRecordProfileTables.filterID = " & Trim(Str(mlngExpressionID))

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a filter in the '" & !Name & "' record profile.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If

  'Match Report/Career Progression/Succesion Planning
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Match Report/Career Progression/Succesion Planning.
    sSQL = "SELECT ASRSysMatchReportName.name, ASRSysMatchReportName.matchReportType" & _
      " FROM ASRSysMatchReportName" & _
      " WHERE ASRSysMatchReportName.table1Filter = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSysMatchReportName.table2Filter = " & Trim(Str(mlngExpressionID))

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
      
      If fUsed Then
        Select Case !matchReportType
          Case 1
            sType = "Succession Planning"
          Case 2
            sType = "Career Progression"
          Case Else
            sType = "Match Report"
        End Select
      
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a filter in the '" & !Name & "' " & sType & ".", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If

  'calendar report
  If Not fUsed Then
    ' Check that the expression is not used as a filter for a Calendar Report.
    sSQL = "SELECT ASRSysCalendarReports.name" & _
      " FROM ASRSysCalendarReports " & _
      " LEFT OUTER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.CalendarReportID" & _
      " WHERE ASRSysCalendarReports.filter = " & Trim(Str(mlngExpressionID)) & _
      " OR ASRSysCalendarReportEvents.filterID = " & Trim(Str(mlngExpressionID))

    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used as a filter in the '" & !Name & "' calendar report.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If

  If Not fUsed Then
    ' Check that the expression is not used as a calculation for a Calendar Report.
    sSQL = "SELECT AsrSysCalendarReports.Name" & _
      " FROM AsrSysCalendarReports" & _
      " WHERE AsrSysCalendarReports.DescriptionExpr = " & Trim(Str(mlngExpressionID)) & _
      "   OR AsrSysCalendarReports.StartDateExpr = " & Trim(Str(mlngExpressionID)) & _
      "   OR AsrSysCalendarReports.EndDateExpr = " & Trim(Str(mlngExpressionID))
    rsCheck.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the expression cannot be deleted.
        MsgBox "This " & LCase(ExpressionTypeName(miExpressionType)) & " cannot be deleted." & vbCr & _
          "It is used in the '" & !Name & "' calendar report.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
  End If


TidyUpAndExit:
  ' Disassociate object variables.
  Set rsCheck = Nothing
  Set objComp = Nothing
  
  ExpressionIsUsed = fUsed
  Exit Function

ErrorTrap:
  MsgBox "Error checking if the expression is used.", _
    vbOKOnly + vbExclamation, App.ProductName
  fUsed = True
  Resume TidyUpAndExit
  
End Function





Public Function ValidityMessage(piInvalidityCode As Integer)
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
      ValidityMessage = "The " & LCase(ExpressionTypeName(miExpressionType)) & " must have at least one component."
    
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
      ValidityMessage = "The complexity of the " & LCase(ExpressionTypeName(miExpressionType)) & " has caused the following SQL error : " & vbCrLf & vbCrLf & _
        "'" & msErrorMessage & "'" & vbCrLf & vbCrLf & _
        "Try simplifying the " & LCase(ExpressionTypeName(miExpressionType)) & "."
    
    ' JPD20020419 Fault 3687
    Case giEXPRVALIDATION_ASSOCSQLERROR
      ValidityMessage = "The complexity of this " & LCase(ExpressionTypeName(miExpressionType)) & " would cause an expression that uses this " & LCase(ExpressionTypeName(miExpressionType)) & " to suffer from the following SQL error : " & vbCrLf & vbCrLf & _
        "'" & msErrorMessage & "'" & vbCrLf & vbCrLf & _
        "Try simplifying this " & LCase(ExpressionTypeName(miExpressionType)) & "."
    
    Case Else
      ValidityMessage = "The function parameter expression must have at least one component."
    
  End Select

End Function

Public Sub NewExpression(Optional ByRef pfCancelled As Boolean)
  ' Handle the definition of a new expression.
  ' Optionally return a flag showing if the expression creation was cancelled.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmExpression
  
  fOK = True
  
  ' Initialize the properties for a new expression.
  InitialiseExpression
  
  ' Display the expression definition form.
  Set frmEdit = New frmExpression
  With frmEdit
    Set .Expression = Me
    .Show vbModal
  
    fOK = Not .Cancelled
    
    If Not IsMissing(pfCancelled) Then
      pfCancelled = .Cancelled
    End If
  End With
  
  If fOK Then
    ' Write the new expression to the database.
    fOK = WriteExpression_Transaction

    ' JDM - 26/11/01 - Fault 3202 - Stores details about this expression
    Select Case miExpressionType
      Case giEXPR_RUNTIMECALCULATION
        Call UtilCreated(utlCalculation, Me.ExpressionID)
      Case giEXPR_RUNTIMEFILTER
        Call UtilCreated(utlFilter, Me.ExpressionID)
    End Select

    ' If the write failed then re-initialize the
    ' properties for a new expression.
    If Not fOK Then
      InitialiseExpression
    End If
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
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
  
  ' Begin the transaction of data.
  gADOCon.BeginTrans
  
  fOK = WriteExpression
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If
  WriteExpression_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function WriteExpression() As Boolean
  ' Write the expression definition to the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objComponent As clsExprComponent
    
  fOK = True
  
  If mlngExpressionID = 0 Then
    
    'MH20010712 Need keep manual record of allocated IDs incase users
    'in SYS MGR have created expressions but not yet saved changes
    'ExpressionID = UniqueColumnValue("ASRSysExpressions", "exprID")
    'ExpressionID = GetUniqueID("Expressions", "ASRSysExpressions", "exprID")
    
    'JPD20010911 Setting ExpressionID clears the mfConstructed flag, which we don't want.
    ' So just set the mlngExpressionID variable. NB. The mfConstructed flag is only reset
    ' when the code is stepped through, not when run without breakpoints. So no real
    ' runtime error, but it just didn't make logical sense.
    mlngExpressionID = GetUniqueID("Expressions", "ASRSysExpressions", "exprID")
    
    ' Add a record for the new expression.
    fOK = (mlngExpressionID > 0)
    
    If fOK Then
      sSQL = "INSERT INTO ASRSysExpressions" & _
        " (exprID, name, TableID, returnType, returnSize, returnDecimals, " & _
        " type, parentComponentID, Username, access, description, ViewInColour, ExpandedNode)" & _
        " VALUES(" & Trim(Str(mlngExpressionID)) & ", " & _
        "'" & Replace(Trim(msExpressionName), "'", "''") & "', " & _
        Trim(Str(mlngBaseTableID)) & ", " & _
        IIf(miExpressionType = giEXPR_RUNTIMECALCULATION, Trim(Str(giEXPRVALUE_UNDEFINED)), Trim(Str(miReturnType))) & ", " & _
        "0,0, " & _
        Trim(Str(miExpressionType)) & ", " & _
        Trim(Str(mlngParentComponentID)) & ", " & _
        "'" & Replace(Trim(msOwner), "'", "''") & "', " & _
        "'" & Trim(msAccess) & "', " & _
        "'" & Replace(Trim(msDescription), "'", "'") & "', " & _
        IIf(mbViewInColour, "1, ", "0, ") & _
        IIf(mbExpandedNode, "1", "0") & ")"
      gADOCon.Execute sSQL, , adCmdText

    End If
  Else
    'JPD 20050812 Fault 10166
    sSQL = "UPDATE ASRSysExpressions" & _
      " SET name = '" & Replace(Trim(msExpressionName), "'", "''") & "'," & _
      " TableID = " & Trim(Str(mlngBaseTableID)) & "," & _
      " returnType = " & IIf(miExpressionType = giEXPR_RUNTIMECALCULATION, Trim(Str(giEXPRVALUE_UNDEFINED)), Trim(Str(miReturnType))) & "," & _
      " returnSize = 0," & _
      " returnDecimals = 0," & _
      " type = " & Trim(Str(miExpressionType)) & "," & _
      " parentComponentID = " & Trim(Str(mlngParentComponentID)) & "," & _
      " Username = '" & Replace(Trim(msOwner), "'", "''") & "'," & _
      " access = '" & Trim(msAccess) & "'," & _
      " description = '" & Replace(Trim(msDescription), "'", "''") & "', " & _
      " ViewInColour = " & IIf(mbViewInColour, "1", "0") & _
      " WHERE exprID = " & Trim(Str(mlngExpressionID))
    
      '" owner = '" & Trim(msOwner) & "'," & _

    gADOCon.Execute sSQL, , adCmdText
  End If

  If fOK Then
    ' Delete the expression's existing components from the database.
    fOK = DeleteExistingComponents
    
    If fOK Then
      ' Add any components for this expression.
      For Each objComponent In mcolComponents
        Set objComponent.ParentExpression = Me
        fOK = objComponent.WriteComponent
        
        If Not fOK Then
          Exit For
        End If
      Next
    End If
  End If

TidyUpAndExit:
  Set objComponent = Nothing
  WriteExpression = fOK
  Exit Function

ErrorTrap:
  MsgBox "Error saving the expression.", _
    vbOKOnly + vbExclamation, App.ProductName
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
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
  sCode = ""
  
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
                             
      ' JPD20020419 Fault 3687
      fOK = .RuntimeCode(sComponentCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
      
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
                        avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbCrLf & _
                          vParameter1 & " - (CAST((" & vParameter1 & " / " & vParameter2 & ") AS INT) * " & vParameter2 & ")" & vbCrLf & _
                          ") end)"
                      Case Else
                        avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbCrLf & _
                          vParameter1 & " " & vbCrLf & _
                          avValues(2, iLoop2) & " " & vbCrLf & _
                          vParameter2 & vbCrLf & _
                          ") end)"
                    End Select
                  Else
                    Select Case .Component.OperatorID
                      Case 5 'And
                        avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) AND (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
                      Case 6 'Or
                        avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) OR (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
                      Case Else
                        avValues(2, iLoop2) = "(" & vbCrLf & _
                          vParameter1 & " " & vbCrLf & _
                          avValues(2, iLoop2) & " " & vbCrLf & _
                          vParameter2 & vbCrLf & _
                          ")"
                    End Select
                  End If
                Else
                  avValues(2, iLoop2) = _
                    IIf(Len(.Component.SQLFixedParam1) > 0, "(", "") & _
                    avValues(2, iLoop2) & vbCrLf & _
                    "(" & vbCrLf & _
                    vParameter1 & vbCrLf & _
                    ", " & vbCrLf & _
                    vParameter2 & vbCrLf & _
                    ")" & _
                    IIf(Len(.Component.SQLFixedParam1) > 0, " " & .Component.SQLFixedParam1 & ")", "")
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
                      avValues(2, iLoop2) = "(" & vbCrLf & _
                        avValues(2, iLoop2) & " " & vbCrLf & _
                        vParameter1 & vbCrLf & _
                        ")"
                  End Select
                Else
                  avValues(2, iLoop2) = _
                    IIf(Len(.Component.SQLFixedParam1) > 0, "(", "") & _
                    avValues(2, iLoop2) & vbCrLf & _
                    "(" & vbCrLf & _
                    vParameter1 & vbCrLf & _
                    ")" & _
                    IIf(Len(.Component.SQLFixedParam1) > 0, " " & .Component.SQLFixedParam1 & ")", "")
                End If
                avValues(2, iParameter1Index) = vbNullString
              End If
              
              If (miExpressionType = giEXPR_RUNTIMECALCULATION) Or _
                (miExpressionType = giEXPR_RUNTIMEFILTER) Then
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
    psRuntimeCode = ""
  End If
  RuntimeCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Function RuntimeFilterCode(psFilterCode As String, _
  pfApplyPermissions As Boolean, _
  Optional pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  ' Return TRUE if the filter code was created okay.
  ' Return the runtime filter SQL code in the parameter 'psFilterCode'.
  ' Apply permissions to the filter code only if the 'pfApplyPermissions' parameter is TRUE.
  ' The filter code is to be used to validate the expression if the 'pfValidating' parameter is TRUE.
  ' This is used to suppress prompting the user for promted values, when we are only validating the expression.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sWhereCode As String
  Dim sBaseTableSource As String
  Dim sRuntimeFilterSQL As String
  Dim alngSourceTables() As Long
  Dim avRelatedTables() As Variant
  Dim rsInfo As New ADODB.Recordset
  
  ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If

  ' Construct the expression from the database definition.
  fOK = ConstructExpression
  
  If fOK Then
    sBaseTableSource = msBaseTableName

    sRuntimeFilterSQL = "SELECT DISTINCT " & sBaseTableSource & ".id" & vbCrLf & _
      "FROM " & sBaseTableSource & " " & vbCrLf

    ' Create an array of the IDs of the tables/view referred to in the expression.
    ' This is used for joining all of the tables/views used.
    ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
    ' Column 2 = table/view ID.
    ReDim alngSourceTables(2, 0)

    ' Get the filter code.
    ' JPD20020419 Fault 3687
    fOK = RuntimeCode(sWhereCode, alngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If

  If fOK Then
    ' Create an array of the tables related to the expression base table.
    ' Used when Joining any other tables/view used.
    ' Column 1 = 'parent' if the expression's base table is the parent of the other table
    '            'child' if the expression's base table is the child of the other table
    ' Column 2 = ID of the other table
    ReDim avRelatedTables(2, 0)
    sSQL = "SELECT 'parent' AS relationship, childID AS tableID" & _
      " FROM ASRSysRelations" & _
      " WHERE parentID = " & Trim(Str(mlngBaseTableID)) & _
      " UNION" & _
      " SELECT 'child' AS relationship, parentID AS tableID" & _
      " FROM ASRSysRelations" & _
      " WHERE childID = " & Trim(Str(mlngBaseTableID))
    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    With rsInfo
      Do While Not .EOF
        iNextIndex = UBound(avRelatedTables, 2) + 1
        ReDim Preserve avRelatedTables(2, iNextIndex)
        avRelatedTables(1, iNextIndex) = !relationship
        avRelatedTables(2, iNextIndex) = !TableID
        .MoveNext
      Loop
      .Close
    End With
    Set rsInfo = Nothing
    
    ' Join any other tables used.
    For iLoop1 = 1 To UBound(alngSourceTables, 2)
      ' Join a table on a parent/child related to the base table.
      'With recTabEdit
'        .Index = "idxTableID"
'        .Seek "=", alngSourceTables(2, iLoop1)
'
'        sJoinTableSource = !TableName
      'End With
      
'      For iLoop2 = 1 To UBound(avRelatedTables, 2)
'        If avRelatedTables(2, iLoop2) = alngSourceTables(2, iLoop1) Then
'
'          If avRelatedTables(1, iLoop2) = "parent" Then
'            sRuntimeFilterSQL = sRuntimeFilterSQL & _
'              "LEFT OUTER JOIN " & sJoinTableSource & " ON " & sBaseTableSource & ".id = " & sJoinTableSource & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbCrLf
'          Else
'            sRuntimeFilterSQL = sRuntimeFilterSQL & _
'              "LEFT OUTER JOIN " & sJoinTableSource & " ON " & sBaseTableSource & ".id_" & Trim(Str(avRelatedTables(2, iLoop2))) & " = " & sJoinTableSource & ".id " & vbCrLf
'          End If
'
'          Exit For
'        End If
'      Next iLoop2
    Next iLoop1
    
'
'    ' Join any other tables/views used.
'    For iLoop1 = 1 To UBound(alngSourceTables, 2)
'      If alngSourceTables(1, iLoop1) = 0 Then
'        Set objTableView = gcoTablePrivileges.FindTableID(alngSourceTables(2, iLoop1))
'      Else
'        Set objTableView = gcoTablePrivileges.FindViewID(alngSourceTables(2, iLoop1))
'      End If
'
'      If objTableView.TableID = mlngBaseTableID Then
'        ' Join a view on the base table.
'        If Not pfApplyPermissions Then
'          sRuntimeFilterSQL = sRuntimeFilterSQL & _
'            "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id = " & objTableView.TableName & ".id" & vbCrLf
'        Else
'          sRuntimeFilterSQL = sRuntimeFilterSQL & _
'            "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id = " & objTableView.RealSource & ".id" & vbCrLf
'        End If
'      Else
'        ' Join a table/view on a parent/child related to the base table.
'        For iLoop2 = 1 To UBound(avRelatedTables, 2)
'          If avRelatedTables(2, iLoop2) = objTableView.TableID Then
'
'            If Not pfApplyPermissions Then
'              If avRelatedTables(1, iLoop2) = "parent" Then
'                sRuntimeFilterSQL = sRuntimeFilterSQL & _
'                  "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id = " & objTableView.TableName & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbCrLf
'              Else
'                sRuntimeFilterSQL = sRuntimeFilterSQL & _
'                  "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id_" & Trim(Str(avRelatedTables(2, iLoop2))) & " = " & objTableView.TableName & ".id " & vbCrLf
'              End If
'            Else
'              If avRelatedTables(1, iLoop2) = "parent" Then
'                sRuntimeFilterSQL = sRuntimeFilterSQL & _
'                  "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id = " & objTableView.RealSource & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbCrLf
'              Else
'                sRuntimeFilterSQL = sRuntimeFilterSQL & _
'                  "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id_" & Trim(Str(avRelatedTables(2, iLoop2))) & " = " & objTableView.RealSource & ".id " & vbCrLf
'              End If
'            End If
'
'            Exit For
'          End If
'        Next iLoop2
'      End If
'
'      Set objTableView = Nothing
'    Next iLoop1
'
    ' Add the filter 'where' clause code.
    If Len(sWhereCode) > 0 Then
      sWhereCode = sWhereCode & " = 1"

      sRuntimeFilterSQL = sRuntimeFilterSQL & _
        "WHERE " & vbCrLf & _
        sWhereCode & vbCrLf
    End If
  End If

TidyUpAndExit:
  If fOK Then
    psFilterCode = sRuntimeFilterSQL
  Else
    psFilterCode = ""
  End If
  RuntimeFilterCode = fOK

  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Public Function RuntimeCalculationCode(palngSourceTables As Variant, _
  psCalcCode As String, _
  pfApplyPermissions As Boolean, _
  Optional pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  ' Return TRUE if the Calculation code was created okay.
  ' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
  ' Apply permissions to the Calculation code only if the 'pfApplyPermissions' parameter is TRUE.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sRuntimeSQL As String

  ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If

  ' Construct the expression from the database definition.
  fOK = ConstructExpression

  If fOK Then
    ' Get the Calculation code.
    ' JPD20020419 Fault 3687
    fOK = RuntimeCode(sRuntimeSQL, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If

  If fOK Then
    If pfApplyPermissions Then
      fOK = (ValidateExpression(True) = giEXPRVALIDATION_NOERRORS)
    End If
  End If

  If fOK And (miReturnType = giEXPRVALUE_LOGIC) Then
    sRuntimeSQL = "convert(bit, " & sRuntimeSQL & ")"
  End If

TidyUpAndExit:
  If fOK Then
    psCalcCode = sRuntimeSQL
  Else
    psCalcCode = ""
  End If
  RuntimeCalculationCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

















Public Property Let ParentComponentID(ByVal plngNewValue As Long)
  ' Set the Parent component ID.
  mlngParentComponentID = plngNewValue
  
End Property

Public Function DeleteExistingComponents() As Boolean
  ' Delete the expression's components and sub-expression's
  ' (ie. function parameter expressions) from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sDeletedExpressionIDs As String
  Dim rsSubExpressions As ADODB.Recordset
  Dim objExpr As clsExprExpression
  
  fOK = True
  sDeletedExpressionIDs = ""
  
  ' Get the expression's function components from the database.
  sSQL = "SELECT ASRSysExpressions.exprID" & _
    " FROM ASRSysExpressions" & _
    " INNER JOIN ASRSysExprComponents" & _
    "   ON ASRSysExpressions.parentComponentID = ASRSysExprComponents.componentID" & _
    " AND ASRSysExprComponents.exprID = " & Trim(Str(mlngExpressionID))
  Set rsSubExpressions = modExpression.OpenRecordset(sSQL, adOpenDynamic, adLockReadOnly)
  
  With rsSubExpressions
    Do While (Not .EOF) And fOK
      ' Instantiate each function parameter expression.
      ' Instruct the function parameter expression to delete its components.
      Set objExpr = New clsExprExpression
      objExpr.ExpressionID = !ExprID
      fOK = objExpr.DeleteExistingComponents
      Set objExpr = Nothing

      ' Add the ID of the sub-expression to the string of sub-expressions to be deleted.
      sDeletedExpressionIDs = sDeletedExpressionIDs & _
        IIf(Len(sDeletedExpressionIDs) > 0, ", ", "") & _
        Trim(Str(!ExprID))
              
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsSubExpressions = Nothing
  
  If Len(sDeletedExpressionIDs) > 0 Then
    ' Delete all existing sub-expressions for this expression from the database.
    sSQL = "DELETE FROM ASRSysExpressions" & _
      " WHERE exprID IN (" & sDeletedExpressionIDs & ")"
    gADOCon.Execute sSQL, , adCmdText
  End If
  
  ' Delete all existing components for this expression from the database.
  sSQL = "DELETE FROM ASRSysExprComponents" & _
    " WHERE exprID = " & Trim(Str(mlngExpressionID))
  gADOCon.Execute sSQL, , adCmdText
  
TidyUpAndExit:
  Set objExpr = Nothing
  Set rsSubExpressions = Nothing
  DeleteExistingComponents = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function




Public Function ModifyComponent(pobjComponent As clsExprComponent) As clsExprComponent
  ' Edit the given component. Return the new component if the changes
  ' were confirmed.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim objCopyComponent As New clsExprComponent
  
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
  Set objCopyComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Property Get ComponentType() As Integer
  ' Return the component type.
  ComponentType = giCOMPONENT_EXPRESSION
  
End Property

Public Function PasteComponent(pobjClipboard As clsExprComponent, pobjPreviousComponent As clsExprComponent, pbPasteBelow As Boolean) As clsExprComponent
    ' Insert a component into the current expression
    ' just before the given component. Return the new component.
    On Error GoTo ErrorTrap
  
    Dim fOK As Boolean
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim objComponent As clsExprComponent
  
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

Public Function InsertComponent(pobjComponent As clsExprComponent) As clsExprComponent
  ' Insert a component into the current expression
  ' just before the given component. Return the new component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim objComponent As clsExprComponent
  
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
  Set objComponent = New clsExprComponent
  Set objComponent.ParentExpression = Me

  ' Instruct the new component to handle its definition.
  fOK = objComponent.NewComponent
  
  If fOK Then
    ' Insert the new component into the expression's
    ' component collection before the given control.
    If iIndex < 0 Then
      mcolComponents.Add objComponent
    Else
      mcolComponents.Add objComponent, , iIndex
    End If
  End If
  
TidyUpAndExit:
  If fOK Then
    Set InsertComponent = objComponent
  Else
    Set InsertComponent = Nothing
  End If
  Set objComponent = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Get Components() As Collection
  ' Return the component collection.
  Set Components = mcolComponents
  
End Property

Public Property Set Components(pcolCollection As Collection)
  ' Set the component collection.
  Set mcolComponents = pcolCollection

End Property

Public Sub EditExpression(Optional ByRef pfCancelled As Boolean)
  ' Edit the selected expression.
  ' Optionally return a flag showing if the editing was cancelled.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmExpr As frmExpression
  
  Me.ActionType = edtEdit
 
  ' Construct the expression from the database definition.
  Screen.MousePointer = vbHourglass
  fOK = ConstructExpression
  Screen.MousePointer = vbDefault
  
  If fOK Then
    ' Display the expression edit form.
    Set frmExpr = New frmExpression
    Set frmExpr.Expression = Me
    
    If Not frmExpr.Cancelled Then
      frmExpr.Show vbModal
    End If
    
    ' If the changes were confirmed then write the changes to the database.
    fOK = Not frmExpr.Cancelled
    
    If Not IsMissing(pfCancelled) Then
      pfCancelled = frmExpr.Cancelled
    End If
    
    If fOK Then
      fOK = WriteExpression_Transaction
    
      ' JDM - 26/11/01 - Fault 3202 - Stores details about this expression
      Select Case miExpressionType
        Case giEXPR_RUNTIMECALCULATION
          Call UtilUpdateLastSaved(utlCalculation, Me.ExpressionID)
        Case giEXPR_RUNTIMEFILTER
          Call UtilUpdateLastSaved(utlFilter, Me.ExpressionID)
      End Select
    
    Else
      mfConstructed = False
    End If
  End If
  
TidyUpAndExit:
  Set frmExpr = Nothing
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Public Property Get Owner() As String
  ' Return the expression owner.
  Owner = msOwner

  If Not mobjBaseComponent Is Nothing Then
    If Not mobjBaseComponent.ParentExpression Is Nothing Then
      Owner = mobjBaseComponent.ParentExpression.Owner
    End If
  End If

End Property

Public Property Let Owner(ByVal psNewValue As String)
  ' Set the expression owner.
  msOwner = psNewValue

End Property

Public Function ValidateExpression(pfTopLevel As Boolean) As Integer
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
  ' null,  giEXPRVALUE_CHARACTER
  '   17,  giEXPRVALUE_CHARACTER
  ' null,  giEXPRVALUE_CHARACTER
  '
  ' The operators are then evaluated to leave the array as :
  ' null,  null
  ' null,  giEXPRVALUE_CHARACTER
  ' null,  null
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

Public Function ValidateSQLCode(Optional plngFixedExprID As Long, Optional psFixedSQLCode As String) As ExprValidationCodes
  ' Validate the expression's SQL code. This picks up on errors such as too many nested levels of the CASE statement.
  On Error GoTo ErrorTrap
  
  Dim lngCalcViews() As Long
  Dim objCalcExpr As clsExprExpression
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
  
  iValidationCode = giEXPRVALIDATION_NOERRORS
    
  If ((miExpressionType = giEXPR_VIEWFILTER) Or _
    (miExpressionType = giEXPR_RUNTIMEFILTER) Or _
    (miExpressionType = giEXPR_RUNTIMECALCULATION)) Then
      
    mfConstructed = True
      
    If ((miExpressionType = giEXPR_VIEWFILTER) Or (miExpressionType = giEXPR_RUNTIMEFILTER)) Then
      If RuntimeFilterCode(sSQLCode, False, True, plngFixedExprID, psFixedSQLCode) Then
  
        On Error GoTo SQLCodeErrorTrap
          
        sProcName = UniqueSQLObjectName("tmpsp_ASRExprTest", 4)
          
        ' Create the test stored procedure to see if the filter expression is valid.
        sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
        gADOCon.Execute sSPCode, , adCmdText
        
        'TM20020605 Fault 3756
        ' Drop the test stored procedure.
'        sSPCode = "IF EXISTS (SELECT *" & _
'          " FROM sysobjects" & _
'          " WHERE type = 'P'" & _
'          " AND name = '" & sProcName & "')" & _
'          " DROP PROCEDURE " & sProcName
'        gADOCon.Execute sSPCode, , adCmdText
        DropUniqueSQLObject sProcName, 4
        
        On Error GoTo ErrorTrap
      Else
        iValidationCode = giEXPRVALIDATION_FILTEREVALUATION
      End If
    End If
    
    If (miExpressionType = giEXPR_RUNTIMECALCULATION) Then
      ReDim lngCalcViews(2, 0)

      If RuntimeCalculationCode(lngCalcViews, sSQLCode, False, True, plngFixedExprID, psFixedSQLCode) Then
        ' Add the required views to the JOIN code.
        strJoinCode = vbNullString
'''        For intCount = 1 To UBound(lngCalcViews, 2)
'''          If lngCalcViews(1, intCount) = 1 Then
'''            sSource = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount)).RealSource
'''
'''            strJoinCode = strJoinCode & vbCrLf & _
'''              " LEFT OUTER JOIN " & sSource & _
'''              " ON " & msBaseTableName & ".ID = " & sSource & ".ID"
'''          End If
'''        Next
  
'        If gbEnableUDFFunctions Then
'          objCalcExpr.UDFCalculationCode lngCalcViews, mvarUDFsRequired(), True
'        End If
  
        sSQLCode = "SELECT " & sSQLCode & _
          " FROM " & msBaseTableName & _
          strJoinCode
  
        On Error GoTo SQLCodeErrorTrap
        
        sProcName = UniqueSQLObjectName("tmpsp_ASRExprTest", 4)

        ' Create the test stored procedure to see if the filter expression is valid.
        sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
        gADOCon.Execute sSPCode, , adExecuteNoRecords
      
        'TM20020605 Fault 3756
        ' Drop the test stored procedure.
'        sSPCode = "IF EXISTS (SELECT *" & _
'          " FROM sysobjects" & _
'          " WHERE type = 'P'" & _
'          " AND name = '" & sProcName & "')" & _
'          " DROP PROCEDURE " & sProcName
'        gADOCon.Execute sSPCode, , adCmdText
        DropUniqueSQLObject sProcName, 4
        
        On Error GoTo ErrorTrap
      Else
        iValidationCode = giEXPRVALIDATION_FILTEREVALUATION
      End If
    End If
  
    If iValidationCode = giEXPRVALIDATION_NOERRORS Then
      ' Need to check if all calcs/filters that use this filter are still okay.
      If (IsMissing(plngFixedExprID) And IsMissing(psFixedSQLCode)) Or _
        ((plngFixedExprID = 0) And (psFixedSQLCode = "")) Then
        lngOriginalExprID = mlngExpressionID
        
        ' Create an array of the IDs of the tables/view referred to in the expression.
        ' This is used for joining all of the tables/views used.
        ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
        ' Column 2 = table/view ID.
        ReDim alngSourceTables(2, 0)
            
        RuntimeCode sSQLCode, alngSourceTables, False, True, plngFixedExprID, psFixedSQLCode
        sOriginalSQLCode = sSQLCode
      Else
        lngOriginalExprID = plngFixedExprID
        sOriginalSQLCode = psFixedSQLCode
      End If
      
      iValidationCode = ValidateAssociatedExpressionsSQLCode(lngOriginalExprID, sOriginalSQLCode)
    End If
  End If
  
TidyUpAndExit:
  ValidateSQLCode = iValidationCode
  Exit Function
        
SQLCodeErrorTrap:
  If (IsMissing(plngFixedExprID) And IsMissing(psFixedSQLCode)) Or _
    ((plngFixedExprID = 0) And (psFixedSQLCode = "")) Then
    iValidationCode = giEXPRVALIDATION_SQLERROR
  Else
    iValidationCode = giEXPRVALIDATION_ASSOCSQLERROR
  End If
  msErrorMessage = Err.Description
  Set mobjBadComponent = Nothing
  Resume TidyUpAndExit

ErrorTrap:
  iValidationCode = giEXPRVALIDATION_UNKNOWNERROR
  Resume TidyUpAndExit
        
End Function



Private Function ValidateAssociatedExpressionsSQLCode(plngFixedExpressionID As Long, psFixedSQLCode As String) As ExprValidationCodes
  ' Validate the SQL code for any expressions that use this expression.
  ' This picks up on errors such as too many nested levels of the CASE statement.
  Dim iValidationCode As ExprValidationCodes
  Dim sSQL As String
  Dim rsTemp As New ADODB.Recordset
  Dim objComp As clsExprComponent
  Dim objExpr As clsExprExpression

  iValidationCode = giEXPRVALIDATION_NOERRORS
  
  ' Do nothing if this is a new expression
  If mlngExpressionID = 0 Then
    ValidateAssociatedExpressionsSQLCode = iValidationCode
    Exit Function
  End If
    
  sSQL = "SELECT componentID" & _
    " FROM ASRSysExprComponents" & _
    " WHERE calculationID = " & mlngExpressionID & _
    " OR filterID = " & mlngExpressionID & _
    " OR (fieldSelectionFilter = " & mlngExpressionID & " AND type = " & CStr(giCOMPONENT_FIELD) & ")"
  rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsTemp
    Do While (Not .EOF) And (iValidationCode = giEXPRVALIDATION_NOERRORS)
      Set objComp = New clsExprComponent
      objComp.ComponentID = !ComponentID

      Set objExpr = New clsExprExpression
      objExpr.ExpressionID = objComp.RootExpressionID
      objExpr.ConstructExpression
      iValidationCode = objExpr.ValidateSQLCode(plngFixedExpressionID, psFixedSQLCode)
      If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
        msErrorMessage = objExpr.ErrorMessage
      End If
      Set objExpr = Nothing

      Set objComp = Nothing
      
      .MoveNext
    Loop
    .Close
  End With
  Set rsTemp = Nothing
        
  ValidateAssociatedExpressionsSQLCode = iValidationCode
  
End Function

Public Property Get BadComponent() As clsExprComponent
  ' Return the component last caused the expression to fail its validity check.
  Set BadComponent = mobjBadComponent

End Property





Public Function ConstructExpression() As Boolean
  ' Read the expression definition from the database and
  ' construct the hierarchy of component class objects.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objComponent As clsExprComponent
  Dim rsExpression As ADODB.Recordset
  Dim rsComponents As ADODB.Recordset
  
  fOK = True
  
  ' Do nothing if the expression is already constructed.
  If mfConstructed Then
    If miExpressionType = giEXPR_RUNTIMECALCULATION Then
      miReturnType = giEXPRVALUE_UNDEFINED
    End If
  
    If mlngExpressionID > 0 Then
      ' Get the expression timestamp.
      sSQL = sSQL & "SELECT CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp" & _
        " FROM ASRSysExpressions" & _
        " WHERE exprID = " & Trim(Str(mlngExpressionID))
      Set rsExpression = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      
      With rsExpression
        fOK = Not (.EOF And .BOF)
        If fOK Then
          If Not mfDontUpdateTimeStamp Then mlngTimeStamp = IIf(IsNull(!intTimeStamp), 0, !intTimeStamp)
        End If
        .Close
      End With

    End If
  Else
    ' Get the expression definition.
    If miExpressionType = giEXPR_UTILRUNTIMEFILTER Then
      ' Utility runtime filters are not tied to a base table.
      sSQL = sSQL & "SELECT ASRSysExpressions.name," & _
        " 0 AS tableID," & _
        " ASRSysExpressions.returnType," & _
        " ASRSysExpressions.type," & _
        " ASRSysExpressions.parentComponentID," & _
        " ASRSysExpressions.Username," & _
        " ASRSysExpressions.access," & _
        " ASRSysExpressions.description," & _
        " ASRSysExpressions.ViewInColour," & _
        " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & _
        " '' AS tableName" & _
        " FROM ASRSysExpressions" & _
        " WHERE exprID = " & Trim(Str(mlngExpressionID))
    Else
      sSQL = sSQL & "SELECT ASRSysExpressions.name," & _
        " ASRSysExpressions.TableID," & _
        " ASRSysExpressions.returnType," & _
        " ASRSysExpressions.type," & _
        " ASRSysExpressions.parentComponentID," & _
        " ASRSysExpressions.Username," & _
        " ASRSysExpressions.access," & _
        " ASRSysExpressions.description," & _
        " ASRSysExpressions.ViewInColour," & _
        " CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp," & _
        " ASRSysTables.tableName" & _
        " FROM ASRSysExpressions" & _
        " INNER JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & _
        " WHERE exprID = " & Trim(Str(mlngExpressionID))
    End If

    Set rsExpression = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsExpression
      fOK = Not (.EOF And .BOF)
      If fOK Then
        ' Read the expression's properties.
        msExpressionName = IIf(IsNull(!Name), "", !Name)
        mlngBaseTableID = IIf(IsNull(!TableID), 0, !TableID)
        miReturnType = IIf(IsNull(!ReturnType), giEXPRVALUE_UNDEFINED, !ReturnType)
        miExpressionType = IIf(IsNull(!Type), giEXPR_UNKNOWNTYPE, !Type)
        
        If miExpressionType = giEXPR_RUNTIMECALCULATION Then
          miReturnType = giEXPRVALUE_UNDEFINED
        End If

        mlngParentComponentID = IIf(IsNull(!ParentComponentID), 0, !ParentComponentID)
        'msOwner = IIf(IsNull(!Owner), gsUserName, !Owner)
        msOwner = IIf(IsNull(!UserName), gsUserName, !UserName)
        msAccess = IIf(IsNull(!Access), ACCESS_READWRITE, !Access)
        msDescription = IIf(IsNull(!Description), "", !Description)
        mlngTimeStamp = IIf(IsNull(!intTimeStamp), 0, !intTimeStamp)
        msBaseTableName = IIf(IsNull(!TableName), "", !TableName)
        mbViewInColour = IIf(IsNull(!ViewInColour), False, !ViewInColour)

      Else
        ' Initialise the expression.
        InitialiseExpression
      End If
    
      .Close
    End With
  
    If fOK Then
      ' Clear the expressions collection of components.
      ClearComponents
  
      ' Get the expression definition.
      sSQL = "SELECT *" & _
        " FROM ASRSysExprComponents" & _
        " WHERE exprID = " & Trim(Str(mlngExpressionID)) & _
        " ORDER BY componentID"
      Set rsComponents = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
      Do While (Not rsComponents.EOF) And fOK
        ' Instantiate a new component object.
        Set objComponent = New clsExprComponent
  
        With objComponent
          ' Initialise the new component's properties.
          Set .ParentExpression = Me
          .ComponentID = rsComponents!ComponentID
  
          ' Instruct the new component to read it's own definition from the database.
          fOK = .ConstructComponent(rsComponents)
        End With
  
        If fOK Then
          ' If the component definition was read correctly then
          ' add the new component to the expression's component collection.
          mcolComponents.Add objComponent
        End If
  
        ' Disassociate object variables.
        Set objComponent = Nothing
  
        rsComponents.MoveNext
      Loop
  
      rsComponents.Close
    End If
  End If
  
TidyUpAndExit:
  mfConstructed = fOK
  Set rsExpression = Nothing
  Set rsComponents = Nothing
  Set objComponent = Nothing
  ConstructExpression = fOK
  Exit Function

ErrorTrap:
  fOK = False
  MsgBox "Error constructing the expression.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function

Public Property Get Access() As String
  ' Return the access code.
  Access = msAccess
  
End Property

Public Property Let Access(ByVal psNewValue As String)
  ' Set the access code.
  msAccess = psNewValue

  If Not mobjBaseComponent Is Nothing Then
    If Not mobjBaseComponent.ParentExpression Is Nothing Then
      mobjBaseComponent.ParentExpression.Access = psNewValue
    End If
  End If

End Property

Public Property Get ActionType() As Long
  ' Return the Action Type
  ActionType = mlngActionType
  
End Property

Public Property Let ActionType(ByVal plngNewValue As Long)
  ' Set the Action Type
  mlngActionType = plngNewValue

End Property

Private Sub InitialiseExpression()
  ' Initialize the properties for a new expression,
  ' and clear the expression's component collection.
  ExpressionID = 0
  
  msExpressionName = ""
  mlngParentComponentID = 0
  msOwner = gsUserName
  
  'JDM - 10/12/01 - Fault 3225 - In Sec manager default all new calcs to be hidden
  'msAccess = "RW"
  msAccess = ACCESS_HIDDEN
  
  msDescription = ""
  mlngTimeStamp = 0
  
  mfConstructed = True

  ' Clear any existing components from
  ' the expression's component collection.
  ClearComponents

End Sub




Public Sub ClearComponents()
  ' Clear the expression's component collection.
  On Error GoTo ErrorTrap
  
  ' Remove all components from the collection.
  Do While mcolComponents.Count > 0
    mcolComponents.Remove 1
  Loop
  Set mcolComponents = Nothing
  
  ' Re-instantiate the collection.
  Set mcolComponents = New Collection
  
  Exit Sub

ErrorTrap:
  Err = False
  
End Sub







Public Property Get Description() As String
  ' Return the expression's description.
  Description = msDescription
  
End Property

Public Property Let Description(psNewValue As String)
  ' Set the expression's descriprion property.
  msDescription = psNewValue
  
End Property

Public Property Get Timestamp() As Long
  ' Return the expression's timestamp value.
  Timestamp = mlngTimeStamp
  
End Property

Public Property Let Timestamp(plngNewValue As Long)
  ' Set the expression's timestamp property.
  mlngTimeStamp = plngNewValue
  
End Property

Public Function Initialise(plngBaseTableID As Long, _
  plngExpressionID As Long, _
  piType As Integer, _
  piReturnType As Integer) As Boolean
  ' Initialise the expression object.
  ' Return TRUE if everything was initialised okay.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  BaseTableID = plngBaseTableID
  ExpressionID = plngExpressionID
  miExpressionType = piType
  miReturnType = piReturnType
  
ReDim malngColumns(0)

TidyUpAndExit:
  Initialise = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ValidateSelection() As Boolean
  ' Validate the expression section.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objExpr As clsExprExpression
  Dim rsCheck As ADODB.Recordset
  Dim iValidationCode As Integer
  
  fOK = True
  
  ' Construct the expression to print.
  Screen.MousePointer = vbHourglass
  fOK = ConstructExpression
  
  If fOK Then
    iValidationCode = ValidateExpr(Me, True)
    
    Select Case iValidationCode
    Case 1: SetExprAccess Me.ExpressionID, ACCESS_HIDDEN
    Case 2: SetExprAccess Me.ExpressionID, ACCESS_HIDDEN
    End Select
    
    fOK = (iValidationCode < 2)
  End If
  
  Screen.MousePointer = vbDefault
    
TidyUpAndExit:
  ValidateSelection = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit
  
End Function

Public Property Get BaseTableName() As String
  ' Return the name of the expression's base table.
  BaseTableName = msBaseTableName
  
End Property

Public Property Let BaseTableName(ByVal psNewValue As String)
  ' Set the name of the expression's base table.
  msBaseTableName = psNewValue

End Property

Public Property Get ViewInColour() As Boolean

    ViewInColour = mbViewInColour

End Property

Public Property Let ViewInColour(ByVal bNewValue As Boolean)

    mbViewInColour = bNewValue

End Property

Public Property Get ExpandedNode() As Boolean

ExpandedNode = mbExpandedNode

End Property

Public Property Let ExpandedNode(ByVal bNewValue As Boolean)

    mbExpandedNode = bNewValue

End Property


Public Function CopyExpressionToClipboard() As Boolean

  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim fOK As Boolean
  Dim sTableName As String
  Dim objComponent As clsExprComponent
  Dim msClipboardText As String
  
  fOK = ConstructExpression
  
  ' Send details to the clipboard
  msClipboardText = ExpressionTypeName(miExpressionType) & " Definition : " & Trim(msExpressionName) & vbCrLf
  msClipboardText = msClipboardText & "Description : " & msDescription & vbCrLf & vbCrLf
  msClipboardText = msClipboardText & "Owner : " & msOwner & vbCrLf
  msClipboardText = msClipboardText & "Access : " & AccessDescription(msAccess)
  'Select Case msAccess
  '  Case "RW": msClipboardText = msClipboardText & "Access : Read / Write"
  '  Case "RO": msClipboardText = msClipboardText & "Access : Read only"
  '  Case "HD": msClipboardText = msClipboardText & "Access : Hidden"
  'End Select
  msClipboardText = msClipboardText & vbCrLf & vbCrLf & "Base Table : " & Trim(msBaseTableName) & vbCrLf & vbCrLf
  
  msClipboardText = msClipboardText & "Components" & vbCrLf & vbCrLf
  
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
    MsgBox "Unable to send the " & ExpressionTypeName(miExpressionType) & " '" & Name & "' to the clipboard." & vbCr & vbCr & _
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
  Dim objComponent As clsExprComponent
  Dim sClipboardMessage As String

  fOK = True
  sClipboardMessage = String(piLevel, vbTab) & "Parameter : " & ComponentDescription & vbCrLf
  
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

Public Sub PrintExpression()
  ' Print the expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objComponent As clsExprComponent
  Dim objPrintDef As clsPrintDef

  Dim frmExpr As frmExpression

  ' Construct the expression to print.
  Screen.MousePointer = vbHourglass
  fOK = ConstructExpression
  Screen.MousePointer = vbDefault

  'TM20010801 Fault 2617
  'Create expression object so that the frmExpression constructor can check the Access
  'and the ownership of the selected expression and handle accordingly.
  If fOK Then
    ' Display the expression edit form.
    Set frmExpr = New frmExpression
    Set frmExpr.Expression = Me
    fOK = Not frmExpr.Cancelled
    Set frmExpr = Nothing
  End If
  
  If Not fOK Then
    Exit Sub
  End If

  If fOK Then
      
    Set objPrintDef = New clsPrintDef

    If objPrintDef.IsOK Then
    
      With objPrintDef
        If .PrintStart(False) Then
          .PrintHeader ExpressionTypeName(miExpressionType) & " Definition : " & Trim(msExpressionName)
    
          .PrintNormal "Description : " & Trim(msDescription)
          .PrintNormal
    
          .PrintNormal "Owner : " & Trim(msOwner)
          
          .PrintNormal "Access : " & AccessDescription(msAccess)
          'Select Case msAccess
          '  Case "RW": .PrintNormal "Access : Read / Write"
          '  Case "RO": .PrintNormal "Access : Read only"
          '  Case "HD": .PrintNormal "Access : Hidden"
          'End Select
          
          .PrintNormal
          .PrintNormal "Base Table : " & Trim(msBaseTableName)
          .PrintNormal
    
          .PrintTitle "Components"
          Printer.FontBold = False
          
          ' Print the components.
          For Each objComponent In mcolComponents
            fOK = objComponent.PrintComponent(1)
            If Not fOK Then
              Printer.KillDoc
              Exit For
            End If
          Next
          Set objComponent = Nothing
          
          'End of document
          .PrintEnd
          .PrintConfirm ExpressionTypeName(miExpressionType) & " : " & Trim(msExpressionName), ExpressionTypeName(miExpressionType) & " Definition"
          
        End If
      End With
    End If
  
  End If
  
TidyUpAndExit:
  Set objComponent = Nothing
  If Not fOK Then
    MsgBox "Unable to print the " & ExpressionTypeName(miExpressionType) & " '" & Name & "'." & vbCr & vbCr & _
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
  Dim objComponent As clsExprComponent

  fOK = True

  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Parameter : " & ComponentDescription
  End With
  
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


' Creates a UDF for this expression if its required
Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean

  Dim iLoop1 As Integer
  Dim fOK As Boolean

  fOK = True

  For iLoop1 = 1 To mcolComponents.Count
    With mcolComponents(iLoop1)
      
      ' Add the created UDFs to the total list
      fOK = .UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
    
    End With
    
    If Not fOK Then
      Exit For
    End If

  Next iLoop1

  UDFCode = fOK

End Function


Public Function UDFCalculationCode(palngSourceTables As Variant, _
  psCalcCode() As String, _
  pfApplyPermissions As Boolean, _
  Optional pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  ' Return TRUE if the Calculation code was created okay.
  ' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
  ' Apply permissions to the Calculation code only if the 'pfApplyPermissions' parameter is TRUE.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sRuntimeSQL As String

  ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If
  
  ' Construct the expression from the database definition.
  fOK = ConstructExpression

  If fOK Then
    ' Get the Calculation code.
    ' JPD20020419 Fault 3687
    fOK = UDFCode(psCalcCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If

  If fOK Then
    If pfApplyPermissions Then
      fOK = (ValidateExpression(True) = giEXPRVALIDATION_NOERRORS)
    End If
  End If

TidyUpAndExit:
  If Not fOK Then
    psCalcCode(UBound(psCalcCode)) = ""
  End If
  UDFCalculationCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function UDFFilterCode(pastrFilterCode() As String, _
  pfApplyPermissions As Boolean, _
  Optional pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  ' Return TRUE if the filter code was created okay.
  ' Return the runtime filter SQL code in the parameter 'pastrFilterCode'.
  ' Apply permissions to the filter code only if the 'pfApplyPermissions' parameter is TRUE.
  ' The filter code is to be used to validate the expression if the 'pfValidating' parameter is TRUE.
  ' This is used to suppress prompting the user for promted values, when we are only validating the expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sWhereCode As String
  Dim sBaseTableSource As String
  Dim sRuntimeFilterSQL As String
  Dim alngSourceTables() As Long
  Dim avRelatedTables() As Variant
  Dim rsInfo As Recordset
  
  ' Check if the 'validating' parameter is set.
  ' If not, set it to FALSE.
  If IsMissing(pfValidating) Then
    pfValidating = False
  End If
  
  ' Construct the expression from the database definition.
  fOK = ConstructExpression

  If fOK Then
    sBaseTableSource = msBaseTableName
       
    ' Create an array of the IDs of the tables/view referred to in the expression.
    ' This is used for joining all of the tables/views used.
    ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
    ' Column 2 = table/view ID.
    ReDim alngSourceTables(2, 0)
    
    ' Get the filter code.
    fOK = UDFCode(pastrFilterCode(), alngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End If

TidyUpAndExit:
  If Not fOK Then
    pastrFilterCode(UBound(pastrFilterCode)) = ""
  End If
  UDFFilterCode = fOK
  
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


