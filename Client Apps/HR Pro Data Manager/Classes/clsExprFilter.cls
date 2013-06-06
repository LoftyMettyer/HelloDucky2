VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprFilter"
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
Private mobjBaseComponent As clsExprComponent


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





Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean

  Dim objFilter As clsExprExpression
  Dim strRuntimeCode As String
  Dim bOK As Boolean
  
  ' Instantiate and generate the runtime for the filter expression.
  Set objFilter = New clsExprExpression
  With objFilter
    .ExpressionID = mlngFilterID
    .ConstructExpression
    bOK = .RuntimeCode(strRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
  End With
    
  ' Return different value depending on passed in parameters
  If mlngFilterID = plngFixedExprID Then
    psRuntimeCode = psFixedSQLCode
    RuntimeCode = True
  Else
    psRuntimeCode = strRuntimeCode
    RuntimeCode = bOK
  End If
    
  Set objFilter = Nothing
  
  
  
' JDM - Fault 8750 - Changed the code around because when passing in fixed string the table array was not being populated

'  If mlngFilterID = plngFixedExprID Then
'    psRuntimeCode = psFixedSQLCode
'    RuntimeCode = True
'  Else
'
'    ' Instantiate the filter expression.
'    Set objFilter = New clsExprExpression
'
'    With objFilter
'      ' Construct the filter expression.
'      .ExpressionID = mlngFilterID
'      .ConstructExpression
'      RuntimeCode = .RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
'    End With
'
'    Set objFilter = Nothing
'  End If

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
    Printer.Print "Filter : " & ComponentDescription
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
  Dim objFilterCopy As New clsExprFilter

  ' Copy the component's basic properties.
  objFilterCopy.FilterID = mlngFilterID
  
  Set CopyComponent = objFilterCopy
   
  ' Disassociate object variables.
  Set objFilterCopy = Nothing
   
End Function

Public Property Get ComponentType() As Integer
  ' Return the component type.
  ComponentType = giCOMPONENT_FILTER
  
End Property


Public Property Get ReturnType() As Integer
  ' Return the filter's return type.
Dim objFilter As clsExprExpression

' Instantiate the filter expression.
Set objFilter = New clsExprExpression

With objFilter
  ' Construct the filter expression.
  .ExpressionID = mlngFilterID
  .ConstructExpression
  .ValidateExpression False
  miReturnType = .ReturnType
End With
  
Set objFilter = Nothing

  ReturnType = miReturnType

End Property


Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Get ComponentDescription() As String
  ' Return the filter's name.
  ComponentDescription = msFilterName
  
End Property



Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property





Public Function WriteComponent() As Boolean
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type, FilterID, valueLogic)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_FILTER)) & "," & _
    " " & Trim(Str(mlngFilterID)) & ", " & _
    " 0)"
  gADOCon.Execute sSQL, , adCmdText
  
TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Public Property Get FilterID() As Long
  ' Return the filter ID property.
  FilterID = mlngFilterID
  
End Property

Public Property Let FilterID(ByVal plngNewValue As Long)
  ' Set the filter ID property.
  mlngFilterID = plngNewValue
  
  ReadFilter
  
End Property

Private Sub ReadFilter()
  ' Read the filter definition from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsFilter As Recordset
  
  ' Set default values.
  msFilterName = "<unknown>"
  
  ' Get the filter definition.
  sSQL = "SELECT name, returnType" & _
    " FROM ASRSysExpressions" & _
    " WHERE exprID = " & Trim(Str(mlngFilterID))
  Set rsFilter = datGeneral.GetRecords(sSQL)
  With rsFilter
    fOK = Not (.EOF And .BOF)
  
    If fOK Then
      msFilterName = !Name
      miReturnType = !ReturnType
    End If
  
    .Close
  End With

TidyUpAndExit:
  Set rsFilter = Nothing
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
  
  sClipboardMessage = String(piLevel, vbTab) & "Filter : " & ComponentDescription & vbCrLf
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

  Dim objFilter As clsExprExpression

  If mlngFilterID = plngFixedExprID Then
    UDFCode = True
  Else
    ' Instantiate the filter expression.
    Set objFilter = New clsExprExpression
    
    With objFilter
      ' Construct the filter expression.
      .ExpressionID = mlngFilterID
      .ConstructExpression
      UDFCode = .UDFCode(psRuntimeCode(), palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
    End With
      
    Set objFilter = Nothing
  End If

End Function


