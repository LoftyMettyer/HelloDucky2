VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprOperator"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngOperatorID As Long
Private msOperatorName As String
Private miReturnType As ExpressionValueTypes
Private miOperandCount As Integer
Private miPrecedence As Integer
Private msSPName As String
Private msSQLCode As String
Private msSQLType As String
Private mfUnknownParameterTypes As Boolean
Private mfCheckDivideByZero As Boolean
Private msSQLFixedParam1 As String
Private mbCastAsFloat As Boolean

' Class handling variables.
Private mobjBaseComponent As CExprComponent



Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ContainsExpression = False
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function







Public Property Get BaseComponent() As CExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property





Public Property Set BaseComponent(ByVal pobjNewValue As CExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property


Public Function ComponentContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the component makes any reference to the given column.
  ' Used to check for recursion.
  ComponentContainsColumn = False
  
End Function

Public Sub ColumnsUsedInThisComponent(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  ' Do nothing for operators.
  
End Sub


Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  ' Do nothing for operators.
  
End Sub



Public Property Get ComponentDescription() As String
  ' Return the operator description.
  ComponentDescription = msOperatorName
  
End Property

Public Function StoredProcedureCode(Optional strBaseTable As String) As String
  ' Operators are implemented directly into the expression's stored procedure code.
  ' No code is required from the Operator component itself..
  StoredProcedureCode = vbNullString

End Function

Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean)
' JPD - NO LONGER REQUIRED.
'  ' Return nothing. Operators are evaluated by the
'  ' expression class.
'  EvaluateComponent = vbNullString
  
End Function

Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the operator's return type.
  ReturnType = miReturnType

End Property
Public Property Get Precedence() As Integer
  ' Return the operator precedence property.
  Precedence = miPrecedence
  
End Property

Public Property Get SQLType() As String
  ' Return the operator SQL Type property.
  SQLType = msSQLType
  
End Property


Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
  ' Return the SQL code for the component.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sCode As String
  
  fOK = True
  
  sCode = msSQLCode

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


Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the Operator component type.
  ComponentType = giCOMPONENT_OPERATOR
  
End Property

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objOperatorCopy As New CExprOperator

  ' Copy the component's basic properties.
  objOperatorCopy.OperatorID = mlngOperatorID
  
  Set CopyComponent = objOperatorCopy
   
  ' Disassociate object variables.
  Set objOperatorCopy = Nothing
   
End Function
Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objOperatorClone As New CExprOperator

  ' Copy the component's basic properties.
  objOperatorClone.OperatorID = mlngOperatorID
  
  fOK = True
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objOperatorClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objOperatorClone = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
    
  fOK = True
  
  With recCompEdit
    .AddNew
    !ComponentID = mobjBaseComponent.ComponentID
    !ExprID = mobjBaseComponent.ParentExpression.ExpressionID
    !Type = giCOMPONENT_OPERATOR
    !OperatorID = mlngOperatorID
  
    .Update
  End With

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Function ConstructComponent() As Boolean
  ' Read the Operator component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean
  
  fReadOk = False
  
  ' Read the operator component record.
  OperatorID = recCompEdit!OperatorID
  
  fReadOk = True
  
TidyUpAndExit:
  ConstructComponent = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
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
  Printer.Print "Operator : " & ComponentDescription
  
  CheckEndOfPage2 glngBottom
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function






Public Property Get OperatorID() As Long
  ' Return the operator ID property.
  OperatorID = mlngOperatorID
  
End Property

Public Property Let OperatorID(ByVal plngNewValue As Long)
  ' Set the operator ID property.
  mlngOperatorID = plngNewValue
  
  ' Read the operator definition from the database.
  ReadOperator

End Property

Private Sub ReadOperator()
  ' Read the operator definition from the operators table.
  On Error GoTo ErrorTrap
  
  Dim objOperatorDef As clsOperatorDef
  Dim objOperatorDefParameter As clsOperatorDefParameter
  
  ' Set default values.
  msOperatorName = "<unknown>"
  miReturnType = giEXPRVALUE_UNDEFINED
  miPrecedence = 1
  miOperandCount = 0
  msSPName = ""
  msSQLCode = ""
  msSQLType = ""
  mfCheckDivideByZero = False
  msSQLFixedParam1 = ""
  mbCastAsFloat = False
  
  ' Get the operator info. from the collection.
  gobjOperatorDefs.Initialise
  If gobjOperatorDefs.IsValidID(mlngOperatorID) Then
    Set objOperatorDef = gobjOperatorDefs.Item("O" & Trim$(Str(mlngOperatorID)))
    
    msOperatorName = objOperatorDef.Name
    miReturnType = objOperatorDef.ReturnType
    miPrecedence = objOperatorDef.Precedence
    miOperandCount = objOperatorDef.OperandCount
    msSPName = objOperatorDef.SPName
    msSQLCode = objOperatorDef.SQLCode
    msSQLType = objOperatorDef.SQLType
    mfCheckDivideByZero = objOperatorDef.CheckDivideByZero
    msSQLFixedParam1 = objOperatorDef.SQLFixedParam1
    mbCastAsFloat = objOperatorDef.CastAsFloat
  
    mfUnknownParameterTypes = False
    For Each objOperatorDefParameter In objOperatorDef.Parameters
      If objOperatorDefParameter.ParameterType = giEXPRVALUE_UNDEFINED Then
        mfUnknownParameterTypes = True
        Exit For
      End If
    Next objOperatorDefParameter
    Set objOperatorDefParameter = Nothing
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub
Public Property Get SQLFixedParam1() As String
  ' Return the first fixed SQL parameter.
  SQLFixedParam1 = msSQLFixedParam1
  
End Property

Public Property Let SQLFixedParam1(ByVal psNewValue As String)
  ' Set the first fixed SQL parameter.
  msSQLFixedParam1 = psNewValue

End Property

Public Property Get CheckDivideByZero() As Boolean
  ' Return the 'check for divide by zero' flag.
  CheckDivideByZero = mfCheckDivideByZero

End Property

Public Property Let CheckDivideByZero(ByVal pfNewValue As Boolean)
  ' Set the 'check for divide by zero' flag.
  mfCheckDivideByZero = pfNewValue

End Property

Public Property Get OperandCount() As Integer
  ' Return the operator's operand count.
  OperandCount = miOperandCount
  
End Property


Public Property Get SQLCode() As String
  ' Return the operator's SQL code.
  SQLCode = msSQLCode
  
End Property

Public Property Let SQLCode(ByVal psNewValue As String)
  ' Set the operator's SQL code.
  msSQLCode = psNewValue
  
End Property

Public Property Get SPName() As String
  ' Return the operator's stored procedure name.
  SPName = msSPName
  
End Property

Public Property Let SPName(ByVal psNewValue As String)
  ' Set the operator's stored procedure name.
  msSPName = psNewValue

End Property

Public Property Get UnknownParameterTypes() As Boolean
  UnknownParameterTypes = mfUnknownParameterTypes
  
End Property

Public Property Let UnknownParameterTypes(ByVal pfNewValue As Boolean)
  mfUnknownParameterTypes = pfNewValue
  
End Property

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String$(piLevel, vbTab) & ComponentDescription & vbNewLine
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get CastAsFloat() As Boolean

  ' Property used to force surrounding values into using numeric (i.e. 0.00 as opposed to 0)
  ' Is necessary in order to get correct values back from SQL when dealing with division signs
  CastAsFloat = mbCastAsFloat

End Property


'MH20060727
'Private Sub Class_Terminate()
'  Terminate
'End Sub
'
'Public Sub Terminate()
'  Set mobjBaseComponent = Nothing
'End Sub
