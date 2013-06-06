VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprOperator"
Attribute VB_GlobalNameSpace = False
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
Private mobjBaseComponent As clsExprComponent

Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ContainsExpression = False
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  COAMsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function






Public Function EvaluateComponent(pvParameter1 As Variant, pvParameter2 As Variant, pvResult As Variant) As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngPosition As Long
  
  fOK = True
  
  Select Case mlngOperatorID
    Case 1
      ' Plus
      pvResult = pvParameter1 + pvParameter2
    
    Case 2
      ' Minus
      pvResult = pvParameter1 - pvParameter2
    
    Case 3
      ' Times By
      pvResult = pvParameter1 * pvParameter2
    
    Case 4
      ' Divided By
      pvResult = pvParameter1 / pvParameter2
    
    Case 5
      ' And
      pvResult = pvParameter1 And pvParameter2

    Case 6
      ' Or
      pvResult = pvParameter1 Or pvParameter2

    Case 7
      'Is Equal To'
      pvResult = (pvParameter1 = pvParameter2)

    Case 8
      ' Is NOT Equal To
      pvResult = (pvParameter1 <> pvParameter2)

    Case 9
      ' Is Less Than
      pvResult = (pvParameter1 < pvParameter2)

    Case 10
      ' Is Greater Than
      pvResult = (pvParameter1 > pvParameter2)

    Case 11
      ' Is Less Than OR Equal To
      pvResult = (pvParameter1 <= pvParameter2)

    Case 12
      ' Is Greater Than OR Equal To
      pvResult = (pvParameter1 >= pvParameter2)

    Case 13
      ' Not
      pvResult = (Not pvParameter1)
    
    Case 14
      ' Is Contained Within
      lngPosition = InStr(1, pvParameter1, pvParameter2)
      
      If IsNull(lngPosition) Then
        pvResult = False
      Else
        pvResult = (lngPosition > 0)
      End If
    
    Case 15
      ' To The Power Of
      pvResult = (pvParameter1 ^ pvParameter2)

    Case 16
      ' Modulas
      pvResult = (pvParameter1 Mod pvParameter2)

    Case 17
      ' Concatenated With
      pvResult = (pvParameter1 & pvParameter2)
    
    Case Else
      ' Unrecognised operator.
      fOK = False
  End Select

TidyUpAndExit:
  EvaluateComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get SQLType() As String
  ' Return the operator SQL Type property.
  SQLType = msSQLType
  
End Property


Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Position the printing.
  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    'Printer.Print ComponentDescription
    Printer.Print "Operator : " & ComponentDescription
  End With
  
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
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type, operatorID, valueLogic)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_OPERATOR)) & "," & _
    " " & Trim(Str(mlngOperatorID)) & "," & _
    " 0)"
  gADOCon.Execute sSQL, , adCmdText
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Get ReturnType() As Integer
  ' Return the operator's return type.
  ReturnType = miReturnType

End Property

Public Property Get ComponentType() As Integer
  ' Return the Operator component type.
  ComponentType = giCOMPONENT_OPERATOR
  
End Property

Public Property Get Precedence() As Integer
  ' Return the operator precedence property.
  Precedence = miPrecedence
  
End Property


Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objOperatorCopy As New clsExprOperator

  ' Copy the component's basic properties.
  objOperatorCopy.OperatorID = mlngOperatorID
  
  Set CopyComponent = objOperatorCopy
   
  ' Disassociate object variables.
  Set objOperatorCopy = Nothing
   
End Function

Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property

Public Property Get OperatorID() As Long
  ' Return the operator ID property.
  OperatorID = mlngOperatorID
  
End Property

Public Property Let OperatorID(ByVal plngNewValue As Long)
  ' Setthe operator ID property.
  mlngOperatorID = plngNewValue
  
  ' Read the operator definition from the database.
  ReadOperator

End Property

Private Sub ReadOperator()
  ' Read the operator definition from the database.
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
  
  ' Get the order definition.
  gobjOperatorDefs.Initialise
  If gobjOperatorDefs.IsValidID(mlngOperatorID) Then
    Set objOperatorDef = gobjOperatorDefs.Item("O" & Trim(Str(mlngOperatorID)))
    
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
  Resume TidyUpAndExit

End Sub
Public Property Get UnknownParameterTypes() As Boolean
  UnknownParameterTypes = mfUnknownParameterTypes
  
End Property

Public Property Let UnknownParameterTypes(ByVal pfNewValue As Boolean)
  mfUnknownParameterTypes = pfNewValue
  
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


Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
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
    psRuntimeCode = ""
  End If
  
  RuntimeCode = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get OperandCount() As Integer
  ' Return the operator's operand count.
  OperandCount = miOperandCount
  
End Property

Public Property Get ComponentDescription() As String
  ' Return the operator description.
  ComponentDescription = msOperatorName
  
End Property

Public Property Get CheckDivideByZero() As Boolean
  ' Return the 'check for divide by zero' flag.
  CheckDivideByZero = mfCheckDivideByZero

End Property

Public Property Let CheckDivideByZero(ByVal pfNewValue As Boolean)
  ' Set the 'check for divide by zero' flag.
  mfCheckDivideByZero = pfNewValue

End Property

Public Property Get SQLFixedParam1() As String
  ' Return the first fixed SQL parameter.
  SQLFixedParam1 = msSQLFixedParam1
  
End Property

Public Property Let SQLFixedParam1(ByVal psNewValue As String)
  ' Set the first fixed SQL parameter.
  msSQLFixedParam1 = psNewValue

End Property

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String(piLevel, vbTab) & ComponentDescription & vbCrLf
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

Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  UDFCode = True

End Function


