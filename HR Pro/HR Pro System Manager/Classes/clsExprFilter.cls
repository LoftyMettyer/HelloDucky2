VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngFilterID As Long
Private msFilterName As String
Private miReturnType As Integer

' Class handling variables.
Private mobjBaseComponent As CExprComponent

Public Sub ColumnsUsedInThisComponent(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  Dim objCalcExpr As CExpression
  
  ' Check if the column is used in the component expression.
  Set objCalcExpr = New CExpression
  With objCalcExpr
    .ExpressionID = mlngFilterID
    .ConstructExpression
    .ColumnsUsedInThisExpression pvColumns
  End With
  Set objCalcExpr = Nothing
  
End Sub

Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  Dim objCalcExpr As CExpression
  Dim lngLoop As Long
  Dim fFound As Boolean
  
  If mlngFilterID > 0 Then
    fFound = False
    For lngLoop = 1 To UBound(palngExpressions)
      If palngExpressions(lngLoop) = mlngFilterID Then
        fFound = True
        Exit For
      End If
    Next lngLoop
    
    If Not fFound Then
      ReDim Preserve palngExpressions(UBound(palngExpressions) + 1)
      palngExpressions(UBound(palngExpressions)) = mlngFilterID
      
      ' Get the expressions used in the component expression.
      Set objCalcExpr = New CExpression
      With objCalcExpr
        .ExpressionID = mlngFilterID
        .ConstructExpression
        .ExpressionsUsedInThisExpression palngExpressions
      End With
      Set objCalcExpr = Nothing
    End If
  End If
  
End Sub

Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  ' Check if the calc component IS the one we're checking for.
  ContainsExpression = (plngExprID = mlngFilterID)
  
  If Not ContainsExpression Then
    ' The calc component IS NOT the one we're checking for.
    ' Check if it contains the one we're looking for.
    ContainsExpression = HasExpressionComponent(mlngFilterID, plngExprID)
  End If
  
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

Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True

  glngBottom = CalculateBottomOfPage

  ' Position the printing.
  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  Printer.Print "Filter : " & ComponentDescription
  
  CheckEndOfPage2 glngBottom
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean)
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated value of the Filter component.
'  On Error GoTo ErrorTrap
'
'  Dim cCalc As CExpression
'  Dim vResult As Variant
'
'  vResult = vbNullString
'
'  ' Instantiate the Filter expression.
'  Set cCalc = New CExpression
'
'  With cCalc
'
'    ' Construct the Filter expression.
'    .ExpressionID = mlngFilterID
'    If .ConstructExpression Then
'      vResult = .EvaluateExpression(pLngRecordID, pfDoPrompts)
'    End If
'
'  End With
'
'TidyUpAndExit:
'  Set cCalc = Nothing
'  EvaluateComponent = vResult
'  Exit Function
'
'ErrorTrap:
'  vResult = vbNullString
'  Err = False
'  Resume TidyUpAndExit
  
End Function
Public Function StoredProcedureCode(Optional strBaseTable As String) As String
  ' Return a string describing the component in a stored procedure.
  On Error GoTo ErrorTrap
  
  Dim objCalc As CExpression
  Dim sProcCode As SystemMgr.cStringBuilder
  Dim sVarName As String
  Dim sIfNullCode  As String
  
  Set sProcCode = New SystemMgr.cStringBuilder
  sVarName = "@component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  sProcCode.TheString = "    /* Evaluate Filter component " & Trim$(Str$(mobjBaseComponent.ComponentID)) & ". */" & vbNewLine
  
  ' Instantiate the Filter expression.
  Set objCalc = New CExpression
  With objCalc
    ' Construct the Filter expression.
    .ExpressionID = mlngFilterID
    
    If .ConstructExpression Then
      ' Get the expression's return type.
      Select Case .ReturnType
        Case giEXPRVALUE_CHARACTER
          sProcCode.Append "    DECLARE " & sVarName & " varchar(MAX)" & vbNewLine
          sIfNullCode = "SET " & sVarName & " = ''"
        Case giEXPRVALUE_NUMERIC
          sProcCode.Append "    DECLARE " & sVarName & " float" & vbNewLine
          sIfNullCode = "SET " & sVarName & " = 0"
        Case giEXPRVALUE_LOGIC
          sProcCode.Append "    DECLARE " & sVarName & " bit" & vbNewLine
          sIfNullCode = "SET " & sVarName & " = 0"
        Case giEXPRVALUE_DATE
          sProcCode.Append "    DECLARE " & sVarName & " datetime" & vbNewLine
          sIfNullCode = "SET " & sVarName & " = null"
      End Select
  
      sProcCode.Append "    IF EXISTS (SELECT Name FROM sysobjects WHERE type = 'P'" & _
        " AND name = 'sp_ASRExpr_" & Trim$(Str$(mlngFilterID)) & "')" & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        EXEC @hResult = sp_ASRExpr_" & Trim$(Str$(mlngFilterID)) & " " & sVarName & " OUTPUT, @id" & vbNewLine & _
        "        IF " & sVarName & " IS NULL " & sIfNullCode & vbNewLine & _
        "    END" & vbNewLine & _
        "    ELSE " & sIfNullCode & vbNewLine
    End If
  End With
  
TidyUpAndExit:
  Set objCalc = Nothing
  StoredProcedureCode = sProcCode.ToString
  Set sProcCode = Nothing
  Exit Function

ErrorTrap:
  sProcCode.TheString = ""
  Err = False
  Resume TidyUpAndExit

End Function


Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the Filter's return type.
  ReturnType = miReturnType

End Property

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objCalcCopy As New CExprFilter

  ' Copy the component's basic properties.
  objCalcCopy.FilterID = mlngFilterID
  
  Set CopyComponent = objCalcCopy
   
  ' Disassociate object variables.
  Set objCalcCopy = Nothing
   
End Function
Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clone the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objCalcClone As New CExprFilter

  fOK = True
  
  ' Copy the component's basic properties.
  objCalcClone.FilterID = mlngFilterID
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objCalcClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objCalcClone = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ComponentContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the component makes any reference to the given column.
  ' Used to check for recursion.
  Dim objCalcExpr As CExpression
  
  ' Check if the column is used in the component expression.
  Set objCalcExpr = New CExpression
  With objCalcExpr
    .ExpressionID = mlngFilterID
    .ConstructExpression
    ComponentContainsColumn = .ExpressionContainsColumn(plngColumnID)
  End With
  Set objCalcExpr = Nothing
  
End Function


Public Property Get ComponentDescription() As String
  ' Return the Filter's name.
  ComponentDescription = msFilterName
  
End Property


Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the component type.
  ComponentType = giCOMPONENT_FILTER
  
End Property




Public Property Get FilterID() As Long
  ' Return the Filter ID property.
  FilterID = mlngFilterID
  
End Property

Public Property Let FilterID(ByVal plngNewValue As Long)
  ' Set the Filter ID property.
  mlngFilterID = plngNewValue
  
  ReadFilter
  
End Property

Public Function ConstructComponent() As Boolean
  ' Read the Filter component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean
  
  fReadOk = False
  
  ' Read the Filter component record.
'  mlngFilterID = recCompEdit!FilterID
  FilterID = recCompEdit!FilterID
  
  fReadOk = True
   
TidyUpAndExit:
  ConstructComponent = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function

Private Sub ReadFilter()
  ' Read the Filter definition from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Set default values.
  msFilterName = "<unknown>"
  
  ' Get the Filter definition.
  recExprEdit.Index = "idxExprID"
  recExprEdit.Seek "=", mlngFilterID, False

  fOK = Not recExprEdit.NoMatch
  If fOK Then
    msFilterName = recExprEdit!Name
    miReturnType = recExprEdit!ReturnType
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub


Public Function WriteComponent() As Boolean
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  With recCompEdit
    'Add a record for the new component.
    .AddNew
    !ComponentID = mobjBaseComponent.ComponentID
    !ExprID = mobjBaseComponent.ParentExpression.ExpressionID
    !Type = giCOMPONENT_FILTER
    !FilterID = mlngFilterID
  
    .Update
  End With

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function

Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String

  fOK = True
  sClipboardMessage = String(piLevel, vbTab) & "Filter : " & ComponentDescription & vbNewLine
  
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


'MH20060727
'Private Sub Class_Terminate()
'  Terminate
'End Sub
'
'Public Sub Terminate()
'  Set mobjBaseComponent = Nothing
'End Sub

' Code stub for .NET migration
Public Function GetSQLElementCode() As String
  GetSQLElementCode = vbNullString
End Function
