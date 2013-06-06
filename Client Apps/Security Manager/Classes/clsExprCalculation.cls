VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngCalculationID As Long
Private msCalculationName As String
Private miReturnType As Integer

' Class handling variables.
Private mobjBaseComponent As clsExprComponent


Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ' Check if the calc component IS the one we're checking for.
  ContainsExpression = (plngExprID = mlngCalculationID)
  
  If Not ContainsExpression Then
    ' The calc component IS NOT the one we're checking for.
    ' Check if it contains the one we're looking for.
    ContainsExpression = HasExpressionComponent(mlngCalculationID, plngExprID)
  End If
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function





Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean

  Dim objCalc As clsExprExpression

  If mlngCalculationID = plngFixedExprID Then
    RuntimeCode = True
    psRuntimeCode = psFixedSQLCode
  Else
    ' Instantiate the calculation expression.
    Set objCalc = New clsExprExpression
    
    With objCalc
      ' Construct the calculation expression.
      .ExpressionID = mlngCalculationID
      .ConstructExpression
      RuntimeCode = .RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
    End With
      
    Set objCalc = Nothing
  End If
  
End Function





Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True

  ' Position the printing.
  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Calculation : " & ComponentDescription
  End With
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objCalcCopy As New clsExprCalculation

  ' Copy the component's basic properties.
  objCalcCopy.CalculationID = mlngCalculationID
  
  Set CopyComponent = objCalcCopy
   
  ' Disassociate object variables.
  Set objCalcCopy = Nothing
   
End Function

Public Property Get ComponentType() As Integer
  ' Return the component type.
  ComponentType = giCOMPONENT_CALCULATION
  
End Property


Public Property Get ReturnType() As Integer
  ' Return the calculation's return type.
Dim objCalc As clsExprExpression

' Instantiate the calculation expression.
Set objCalc = New clsExprExpression

With objCalc
  ' Construct the calculation expression.
  .ExpressionID = mlngCalculationID
  .ConstructExpression
  .ValidateExpression False
  miReturnType = .ReturnType
End With
  
Set objCalc = Nothing

  ReturnType = miReturnType

End Property


Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Get ComponentDescription() As String
  ' Return the calculation's name.
  ComponentDescription = msCalculationName
  
End Property



Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property





Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type, calculationID, valueLogic)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_CALCULATION)) & "," & _
    " " & Trim(Str(mlngCalculationID)) & ", " & _
    " 0)"
  gADOCon.Execute sSQL, , adCmdText
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Public Property Get CalculationID() As Long
  ' Return the calculation ID property.
  CalculationID = mlngCalculationID
  
End Property

Public Property Let CalculationID(ByVal plngNewValue As Long)
  ' Set the calculation ID property.
  mlngCalculationID = plngNewValue
  
  ReadCalculation
  
End Property

Private Sub ReadCalculation()
  ' Read the calculation definition from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsCalculation As Recordset
  
  ' Set default values.
  msCalculationName = "<unknown>"
  
  ' Get the calculation definition.
  sSQL = "SELECT name, returnType" & _
    " FROM ASRSysExpressions" & _
    " WHERE exprID = " & Trim(Str(mlngCalculationID))
  Set rsCalculation = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsCalculation
    fOK = Not (.EOF And .BOF)
  
    If fOK Then
      msCalculationName = !Name
      miReturnType = !ReturnType
    End If
  
    .Close
  End With

TidyUpAndExit:
  Set rsCalculation = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String(piLevel, vbTab) & "Calculation : " & ComponentDescription & vbCrLf
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  Dim objCalc As clsExprExpression

  If mlngCalculationID = plngFixedExprID Then
    UDFCode = True
  Else
    ' Instantiate the calculation expression.
    Set objCalc = New clsExprExpression
    
    With objCalc
      ' Construct the calculation expression.
      .ExpressionID = mlngCalculationID
      .ConstructExpression
      UDFCode = .UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
    End With
      
    Set objCalc = Nothing
  End If

End Function


