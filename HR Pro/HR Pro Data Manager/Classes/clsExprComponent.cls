VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngComponentID As Long
Private miComponentType As ExpressionComponentTypes

' Class handling variables.
Private mobjParentExpression As clsExprExpression
Private mvComponent As Variant
Private mastrUDFFunctions() As String

' Definition for expanded/unexpanded status of the component
Private mbExpanded As Boolean
Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ContainsExpression = mvComponent.ContainsExpression(plngExprID)
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function



Public Function NewOperatorFunctionComponent(piComponentType As ExpressionComponentTypes, _
  piOpFuncID As Integer) As Boolean
  ' JPD20021108 Fault 3287
  ' Define a new operator/function component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Initialize the properties for a new expression.
  InitializeComponent
  
  ComponentType = piComponentType
  
  If piComponentType = giCOMPONENT_OPERATOR Then
    mvComponent.OperatorID = piOpFuncID
  Else
    mvComponent.FunctionID = piOpFuncID
  End If

TidyUpAndExit:
  NewOperatorFunctionComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function EvaluateComponent(palngColumnValues As Variant, pvResult As Variant) As Boolean
  ' Return the evaluated value of the component.
  EvaluateComponent = mvComponent.EvaluateComponent(palngColumnValues, pvResult)
  
End Function

Public Property Let ExpandedNode(ByVal bNewValue As Boolean)
    'Set whether this component node is expanded or not
    mbExpanded = bNewValue

    Select Case Me.ComponentType
        Case giCOMPONENT_FUNCTION
            Me.Component.ExpandedNode = bNewValue

        Case giCOMPONENT_EXPRESSION
            Me.Component.ExpandedNode = bNewValue

    End Select

End Property
Public Property Get ExpandedNode() As Boolean
    'Return whether this node is expanded or not
    ExpandedNode = mbExpanded

End Property



Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  ' Return the runtime filter SQL code for the component.
  RuntimeCode = mvComponent.RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
   
End Function




Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = Component.PrintComponent(piLevel)
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Function WriteComponent() As Boolean
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngNewID As Long
  
  
  ' Update the real component expression id property, and give it
  ' a unique component id.
  
  'MH20010712 Need keep manual record of allocated IDs incase users
  'in SYS MGR have created expressions but not yet saved changes
  'lngNewID = UniqueColumnValue("ASRSysExprComponents", "componentID")
  lngNewID = GetUniqueID("ExprComponents", "ASRSysExprComponents", "componentID")



  fOK = (lngNewID > 0)
  
  If fOK Then
    mlngComponentID = lngNewID
    Set mvComponent.BaseComponent = Me
      
    ' Instruct the real component to write its definition to the
    ' component recordset.
    fOK = mvComponent.WriteComponent
  End If
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Get Component() As Variant
  ' Return the real component object.
  Set Component = mvComponent
  
End Property

Public Function NewComponent() As Boolean
  ' Define a new component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmExprComponent
  
  ' Initialize the properties for a new expression.
  InitializeComponent
  
  ' Display the component definition form.
  Set frmEdit = New frmExprComponent
  With frmEdit
    Set .Component = Me
    .Show vbModal
    fOK = Not .Cancelled
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set frmEdit = Nothing
  NewComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function CopyComponent() As clsExprComponent
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objCopyComponent As New clsExprComponent
  
  ' Copy the component's basic properties.
  With objCopyComponent
    .ComponentType = miComponentType
    Set .ParentExpression = mobjParentExpression
  
    ' Instruct the original component to copy itself.
    Set .Component = mvComponent.CopyComponent
    Set .Component.BaseComponent = objCopyComponent
  
    fOK = Not .Component Is Nothing
  End With

    'Copy whether this object is in expanded mode.
    objCopyComponent.ExpandedNode = mbExpanded

  
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


Public Function ModifyComponent() As Boolean
  ' Edit the component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmExprComponent
  
  ' Display the component definition form.
  Set frmEdit = New frmExprComponent
  With frmEdit
    Set .Component = Me
    .Show vbModal
    fOK = Not .Cancelled
  End With
  
TidyUpAndExit:
  Set frmEdit = Nothing
  ModifyComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Sub InitializeComponent()
  ' Initialize the properties for a new component.
  mlngComponentID = 0
  
'  If (mobjParentExpression.ExpressionType = ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC) Then
'    ComponentType = giCOMPONENT_OPERATOR
'  Else
    ComponentType = giCOMPONENT_FIELD
'  End If
  
  Set mvComponent.BaseComponent = Me

End Sub

Public Property Get ComponentDescription() As String
  ' Return a text description of the component.
  ComponentDescription = mvComponent.ComponentDescription

End Property

Public Property Set Component(pvNewComponent As Variant)
  ' Set the real component object.
  If Not pvNewComponent Is Nothing Then
    Set mvComponent = pvNewComponent
    miComponentType = mvComponent.ComponentType
  End If
  
End Property

Public Function ConstructComponent(prsComponents As ADODB.Recordset) As Boolean
  ' Read the component definition from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean

  fOK = True
  
  ' Initialise the component with the definition from the database.
  ComponentType = prsComponents!Type
  
  With mvComponent
    Select Case miComponentType
      Case giCOMPONENT_FIELD
        .TableID = prsComponents!fieldTableID
        .ColumnID = prsComponents!fieldColumnID
        .FieldPassType = prsComponents!fieldPassBy
        .SelectionType = prsComponents!fieldSelectionRecord
        .SelectionLine = prsComponents!fieldSelectionLine
        .SelectionOrderID = prsComponents!fieldSelectionOrderID
        .SelectionFilterID = prsComponents!FieldSelectionFilter
  
      Case giCOMPONENT_FUNCTION
        .FunctionID = prsComponents!FunctionID
       
      Case giCOMPONENT_CALCULATION
        .CalculationID = prsComponents!CalculationID
    
      Case giCOMPONENT_VALUE
        .ReturnType = prsComponents!ValueType
        Select Case prsComponents!ValueType
          Case giEXPRVALUE_CHARACTER
            .Value = IIf(IsNull(prsComponents!valueCharacter), "", prsComponents!valueCharacter)
          Case giEXPRVALUE_NUMERIC
            .Value = IIf(IsNull(prsComponents!valueNumeric), 0, prsComponents!valueNumeric)
          Case giEXPRVALUE_LOGIC
            .Value = IIf(IsNull(prsComponents!valueLogic), True, prsComponents!valueLogic)
          Case giEXPRVALUE_DATE
            
            'MH20010201 Fault 1576
            '.Value = IIf(IsNull(prsComponents!valueDate), Date, prsComponents!valueDate)
            .Value = prsComponents!valueDate
        End Select
        
      Case giCOMPONENT_OPERATOR
        .OperatorID = prsComponents!OperatorID
      
      Case giCOMPONENT_TABLEVALUE
        ' Do nothing as Table Value components are treated as Value components.
        .TableID = prsComponents!LookupTableID
        .ColumnID = prsComponents!LookupColumnID
        .ReturnType = prsComponents!ValueType

        Select Case prsComponents!ValueType
          Case giEXPRVALUE_CHARACTER
            .Value = IIf(IsNull(prsComponents!valueCharacter), "", prsComponents!valueCharacter)
          Case giEXPRVALUE_NUMERIC
            .Value = IIf(IsNull(prsComponents!valueNumeric), 0, prsComponents!valueNumeric)
          Case giEXPRVALUE_LOGIC
            .Value = IIf(IsNull(prsComponents!valueLogic), True, prsComponents!valueLogic)
          Case giEXPRVALUE_DATE
            .Value = prsComponents!valueDate
        End Select
      
      Case giCOMPONENT_PROMPTEDVALUE
        .Prompt = IIf(IsNull(prsComponents!promptDescription), "", prsComponents!promptDescription)
        .ValueType = IIf(IsNull(prsComponents!ValueType), giEXPRVALUE_CHARACTER, prsComponents!ValueType)
        .ReturnSize = IIf(IsNull(prsComponents!promptSize), 1, prsComponents!promptSize)
        .ReturnDecimals = IIf(IsNull(prsComponents!promptDecimals), 0, prsComponents!promptDecimals)
        .ValueFormat = IIf(IsNull(prsComponents!promptMask), "", prsComponents!promptMask)
        .DefaultDateType = IIf(IsNull(prsComponents!PromptDateType), 0, prsComponents!PromptDateType)
        
        Select Case prsComponents!ValueType
          Case giEXPRVALUE_CHARACTER
            .DefaultValue = IIf(IsNull(prsComponents!valueCharacter), "", prsComponents!valueCharacter)
          Case giEXPRVALUE_NUMERIC
            .DefaultValue = IIf(IsNull(prsComponents!valueNumeric), 0, prsComponents!valueNumeric)
          Case giEXPRVALUE_LOGIC
            .DefaultValue = IIf(IsNull(prsComponents!valueLogic), False, prsComponents!valueLogic)
          Case giEXPRVALUE_DATE
            'JPD 20031211 Fault 7806
            '.DefaultValue = IIf(IsNull(prsComponents!valueDate), Date, prsComponents!valueDate)
            .DefaultValue = IIf(IsNull(prsComponents!valueDate), 0, prsComponents!valueDate)
          Case giEXPRVALUE_TABLEVALUE
            .DefaultValue = IIf(IsNull(prsComponents!valueCharacter), "", prsComponents!valueCharacter)
        End Select
        .LookupColumn = IIf(IsNull(prsComponents!fieldColumnID), 0, prsComponents!fieldColumnID)
        
      Case giCOMPONENT_CUSTOMCALC
        ' Not required.
  
      Case giCOMPONENT_EXPRESSION
        ' Sub-expressions are handled via the Function component class.

      Case giCOMPONENT_FILTER
        ' Load information for filters
        .FilterID = prsComponents!FilterID

    End Select
      
  End With

TidyUpAndExit:
  ConstructComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Function RootExpressionID() As Long
  ' Return the id of the expression which contains this component.
  ' NB. We are not returning the id of the immediate parent expression;
  ' rather the top-level parent expression. Return 0 if we are unable to
  ' determine the root expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngRootExprID As Long
  Dim sSQL As String
  Dim objComp As clsExprComponent
  Dim rsExpressions As Recordset
  
  sSQL = "SELECT ASRSysExpressions.parentComponentID, ASRSysExpressions.exprID" & _
    " FROM ASRSysExpressions" & _
    " JOIN ASRSysExprComponents ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID" & _
    " WHERE ASRSysExprComponents.componentID = " & Trim(Str(mlngComponentID))
  Set rsExpressions = datGeneral.GetRecords(sSQL)
  With rsExpressions
    fOK = Not (.EOF And .BOF)
  
    If fOK Then
      ' See if the parent expression is a top level expression.
      If !ParentComponentID = 0 Then
        lngRootExprID = !ExprID
      Else
        ' If the parent expression is not a top-level expression then
        ' find the parent expression's parent expression. Confused yet ?
        Set objComp = New clsExprComponent
        objComp.ComponentID = !ParentComponentID
        lngRootExprID = objComp.RootExpressionID
        Set objComp = Nothing
      End If
    End If
  
    .Close
  End With
  
TidyUpAndExit:
  Set rsExpressions = Nothing
  Set objComp = Nothing
  If fOK Then
    RootExpressionID = lngRootExprID
  Else
    RootExpressionID = 0
  End If
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Property Let ComponentType(ByVal piNewValue As Integer)
  ' Set the component type property.
  If miComponentType <> piNewValue Then
    miComponentType = piNewValue
  
    Set mvComponent = Nothing
    
    ' Instantiate the correct type of component object for
    ' the given component type.
    Select Case miComponentType
    
      Case giCOMPONENT_FIELD
        Set mvComponent = New clsExprField
        
      Case giCOMPONENT_FUNCTION
        Set mvComponent = New clsExprFunction

      Case giCOMPONENT_CALCULATION
        Set mvComponent = New clsExprCalculation
      
      Case giCOMPONENT_VALUE
        Set mvComponent = New clsExprValue
      
      Case giCOMPONENT_OPERATOR
        Set mvComponent = New clsExprOperator
      
      Case giCOMPONENT_TABLEVALUE
        Set mvComponent = New clsExprTableLookup
      
      Case giCOMPONENT_PROMPTEDVALUE
        Set mvComponent = New clsExprPromptedValue
      
      Case giCOMPONENT_CUSTOMCALC
        ' Not required.
  
      Case giCOMPONENT_EXPRESSION
        Set mvComponent = New clsExprExpression

      Case giCOMPONENT_FILTER
        Set mvComponent = New clsExprFilter

    End Select
  
    If Not mvComponent Is Nothing Then
      Set mvComponent.BaseComponent = Me
    End If
  End If
  
End Property

Public Property Get ComponentType() As Integer
  ' Return the component type property.
  ComponentType = miComponentType

End Property


Public Property Get ReturnType() As Integer
  ' Return the component's return type.
  ReturnType = mvComponent.ReturnType
  
End Property


Public Property Let ComponentID(ByVal plngNewValue As Long)
  ' Set the component id property.
  mlngComponentID = plngNewValue

End Property

Public Property Get ComponentID() As Long
  ' Return the component id property.
  ComponentID = mlngComponentID

End Property


Public Property Get ParentExpression() As clsExprExpression
  ' Return the component's parent expression.
  Set ParentExpression = mobjParentExpression
  
End Property

Public Property Set ParentExpression(ByVal pobjNewValue As clsExprExpression)
  ' Set the component's parent expression property.
  Set mobjParentExpression = pobjNewValue
  
End Property


Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = Component.CopyToClipboard(piLevel)
  
TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

' UDF code for this component
Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  ' Return the runtime filter SQL code for the component.
  UDFCode = mvComponent.UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  
End Function

