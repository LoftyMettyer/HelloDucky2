VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprComponent"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngComponentID As Long
Private miComponentType As ExpressionComponentTypes

' Class handling variables.
Private mobjParentExpression As CExpression
Private mvComponent As Variant
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
Public Property Let ExpandedNode(ByVal bNewValue As Boolean)
    'Set whether this component node is expanded or not
    mbExpanded = bNewValue

End Property
Public Property Get ExpandedNode() As Boolean
    'Return whether this node is expanded or not
    ExpandedNode = mbExpanded

End Property

Public Function ParentExpressionName() As String
  ' Return the name of the parent expression.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sName As String
  
  fOK = True
  
  With mobjParentExpression
    If .ReadExpressionDetails Then
      sName = .ParentExpressionName
    Else
      fOK = False
    End If
  End With
  
TidyUpAndExit:
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







Public Property Get ParentExpression() As CExpression
  ' Return the component's parent expression.
  Set ParentExpression = mobjParentExpression
  
End Property

Public Property Set ParentExpression(ByVal pobjNewValue As CExpression)
  ' Set the component's parent expression property.
  Set mobjParentExpression = pobjNewValue
  
End Property

Private Sub InitializeComponent()
  ' Initialize the properties for a new component.
  mlngComponentID = 0
  
  If (mobjParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
    And (mobjParentExpression.ReturnType < 100) Then
    
    ComponentType = giCOMPONENT_WORKFLOWFIELD
  Else
    ComponentType = giCOMPONENT_FIELD
  End If
  
  Set mvComponent.BaseComponent = Me

End Sub
Public Function NewComponent() As Boolean
  ' Define a new component.
  Dim frmEdit As frmComponent
  
  ' Initialize the properties for a new expression.
  InitializeComponent
  
  ' Display the component definition form.
  Set frmEdit = New frmComponent
  Set frmEdit.Component = Me
  frmEdit.Show vbModal
  
  NewComponent = Not frmEdit.Cancelled
  
  ' Disassociate object variables.
  Set frmEdit = Nothing

End Function





Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the component type property.
  ComponentType = miComponentType

End Property

Public Property Let ComponentType(ByVal piNewValue As ExpressionComponentTypes)
  ' Set the component type property.
  If miComponentType <> piNewValue Then
    miComponentType = piNewValue
  
    Set mvComponent = Nothing
    
    ' Instantiate the correct type of component object for
    ' the given component type.
    Select Case miComponentType
    
      Case giCOMPONENT_FIELD
        Set mvComponent = New CExprField

      Case giCOMPONENT_FUNCTION
        Set mvComponent = New CExprFunction
        
      Case giCOMPONENT_CALCULATION
        Set mvComponent = New CExprCalculation
      
      Case giCOMPONENT_VALUE
        Set mvComponent = New CExprValue
      
      Case giCOMPONENT_OPERATOR
        Set mvComponent = New CExprOperator
      
      Case giCOMPONENT_TABLEVALUE
        Set mvComponent = New CExprTableValue
        miComponentType = giCOMPONENT_TABLEVALUE
      
      'MH20010320 Removed as not required ???
      'Case giCOMPONENT_PROMPTEDVALUE
      '  Set mvComponent = New CExprPromptedValue
      
      Case giCOMPONENT_CUSTOMCALC
        ' Not required.
  
      Case giCOMPONENT_EXPRESSION
        Set mvComponent = New CExpression

      Case giCOMPONENT_FILTER
        Set mvComponent = New CExprFilter

      Case giCOMPONENT_WORKFLOWVALUE
        Set mvComponent = New CExprWorkflowValue
      
      Case giCOMPONENT_WORKFLOWFIELD
        Set mvComponent = New CExprWorkflowField

    End Select
    
    If Not mvComponent Is Nothing Then
      Set mvComponent.BaseComponent = Me
    End If
  End If
  
End Property
Public Function ModifyComponent() As Boolean
  ' Edit the component.
  Dim frmEdit As frmComponent
  
  ' Display the component definition form.
  Set frmEdit = New frmComponent
  Set frmEdit.Component = Me
  frmEdit.Show vbModal
  
  ModifyComponent = Not frmEdit.Cancelled
  
  ' Disassociate object variables.
  Set frmEdit = Nothing
  
End Function


Public Property Get Component() As Variant
  ' Return the real component object.
  Set Component = mvComponent
  
End Property


Public Property Set Component(pvNewComponent As Variant)
  ' Set the real component object.
  If Not pvNewComponent Is Nothing Then
    Set mvComponent = pvNewComponent
    miComponentType = mvComponent.ComponentType
  End If
  
End Property
Public Function CopyComponent() As CExprComponent
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objCopyComponent As New CExprComponent
  
  ' Copy the component's basic properties.
  With objCopyComponent
    .ComponentType = miComponentType
    Set .ParentExpression = mobjParentExpression
  
    ' Instruct the original component to copy itself.
    Set .Component = mvComponent.CopyComponent
    Set .Component.BaseComponent = objCopyComponent
  End With

    'Copy whether this object is in expanded mode.
    objCopyComponent.ExpandedNode = mbExpanded
  
  Set CopyComponent = objCopyComponent
   
  ' Disassociate object variables.
  Set objCopyComponent = Nothing
   
End Function

Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NOT REQUIRED
'  ' Return the evaluated value of the component.
'  EvaluateComponent = mvComponent.EvaluateComponent(pLngRecordID, pfDoPrompts)
   
End Function


Public Function WriteComponent() As Boolean
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngNewID As Long
    
  ' Update the real component expression id property, and give it
  ' a unique component id.


  'MH20010712
  'lngNewID = Database.UniqueColumnValue("tmpComponents", "componentID")
  lngNewID = GetUniqueID("ExprComponents", "tmpComponents", "componentID")
  
  
  
  ComponentID = lngNewID
  Set mvComponent.BaseComponent = Me
  
  ' Instruct the real component to write its definition to the
  ' component recordset.
  fOK = mvComponent.WriteComponent
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function








Public Function ConstructComponent() As Boolean
  ' Read the component definition from the component recordset.
  ' NB. This function constructs the hiearchy of object that
  ' make up the real component definition. Use the ReadComponentDetails
  ' function if you just want to get the basic component details.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Read the basic component details.
  If ReadComponentDetails Then
  
    ' Instruct the real component object to construct the
    ' hierarchy of objects that define it.
    mvComponent.ConstructComponent
  End If
   
TidyUpAndExit:
  ConstructComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function
Public Function ReadComponentDetails() As Boolean
  ' Read the basic component details from the component recordset.
  ' NB. This function does not construct the hiearchy of object that
  ' make up the real component definition. Use the ConstructComponent
  ' function to do this.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Read the component record if it exists.
  If FindComponent Then
    With recCompEdit
      ComponentType = !Type
    End With
  End If
   
TidyUpAndExit:
  ReadComponentDetails = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function

Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
  
  ' Return the runtime filter SQL code for the component.
  RuntimeCode = mvComponent.RuntimeCode(psRuntimeCode, _
    palngSourceTables, _
    pfValidating, _
    psPrefix)
   
End Function


Public Function StoredProcedureCode(Optional strBaseTable As String) As String
  StoredProcedureCode = mvComponent.StoredProcedureCode(strBaseTable)
   
End Function

Public Function RootExpressionID() As Long
  ' Return the id of the expression which contains this component.
  ' NB. We are not returning the id of the immediate parent expression;
  ' rather the top-level parent expression. Return 0 if we are unable to
  ' determine the root expression.
  On Error GoTo ErrorTrap
  
  Dim lngRootExprID As Long
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objComp As CExprComponent
  Dim rsExpressions As dao.Recordset
  
  sSQL = "SELECT tmpExpressions.parentComponentID, tmpExpressions.exprID" & _
    " FROM tmpExpressions, tmpComponents" & _
    " WHERE tmpComponents.componentID = " & Trim(Str(mlngComponentID)) & _
    " AND tmpExpressions.exprID = tmpComponents.exprID"
  Set rsExpressions = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  With rsExpressions
    fOK = Not (.EOF And .BOF)
  
    If fOK Then
      ' See if the parent expression is a top level expression.
      If !ParentComponentID = 0 Then
        lngRootExprID = !ExprID
      Else
        ' If the parent expression is not a top-level expression then
        ' find the parent expression's parent expression. Confused yet ?
        Set objComp = New CExprComponent
        objComp.ComponentID = !ParentComponentID
        lngRootExprID = objComp.RootExpressionID
        'objComp.Terminate   'MH20060727
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


Private Function FindComponent() As Boolean
  ' Locate the current component in the component recordset.
  On Error GoTo ErrorTrap
  
  With recCompEdit
    .Index = "idxCompID"
    .Seek "=", mlngComponentID
    FindComponent = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindComponent = False

End Function



Public Property Let ComponentID(ByVal plngNewValue As Long)
  ' Set the component id property.
  mlngComponentID = plngNewValue

End Property

Public Property Get ComponentID() As Long
  ' Return the component id property.
  ComponentID = mlngComponentID

End Property


Public Sub ColumnsUsedInThisComponent(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  Component.ColumnsUsedInThisComponent pvColumns
  
End Sub


Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  Component.ExpressionsUsedInThisComponent palngExpressions
  
End Sub



Public Function ComponentContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the component makes any reference to the given column.
  ' Used to check for recursion.
  ComponentContainsColumn = Component.ComponentContainsColumn(plngColumnID)
  
End Function



Public Property Get ComponentDescription() As String
  ' Return a text description of the component.
  ComponentDescription = mvComponent.ComponentDescription

End Property

Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the component's return type.
  ReturnType = mvComponent.ReturnType
  
End Property


Public Function CloneComponent(pavCloneRegister As Variant) As CExprComponent
  ' Clone the current component and return the clone.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewTableID As Long
  Dim vComponent As Variant
  Dim objCloneComponent As New CExprComponent
  
  fOK = True
  
  ' Clone the component's basic properties.
  With objCloneComponent
    .ComponentType = miComponentType
  
    ' Instruct the original component to copy itself.
    If miComponentType = giCOMPONENT_EXPRESSION Then
      Set vComponent = mvComponent.CloneExpression(pavCloneRegister)
    Else
      Set vComponent = mvComponent.CloneComponent(pavCloneRegister)
    End If
    
    fOK = Not vComponent Is Nothing
    
    If fOK Then
      Set .Component = vComponent
    End If
    
  End With
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objCloneComponent
  Else
    Set CloneComponent = Nothing
  End If
  
  ' Disassociate object variables.
  Set objCloneComponent = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
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

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Print the component definition.
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

Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mvComponent = Nothing
  
End Sub


'MH20060727
'Private Sub Class_Terminate()
'  Terminate
'End Sub
'
'Public Sub Terminate()
'  Set mobjParentExpression = Nothing
'  If Not mvComponent Is Nothing Then
'    mvComponent.Terminate
'    Set mvComponent = Nothing
'  End If
'End Sub

' Code stub for .NET migration
Public Function GetSQLElementCode() As String
  GetSQLElementCode = vbNullString
End Function
