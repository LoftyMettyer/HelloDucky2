VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private miType As ExpressionValueTypes
Private mdblNumericValue As Double
Private msCharacterValue As String
Private mfLogicValue As Boolean

'MH20010201 Fault 1576
'Private mdtDateValue As Date
Private mdtDateValue As Variant

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





Public Function EvaluateComponent(palngColumnValues As Variant, pvResult As Variant) As Boolean
  ' Return the Value of the component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  fOK = True
  
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      ' JPD20021024 Fault 4653
      pvResult = RTrim(UCase(msCharacterValue))
      'pvResult = msCharacterValue
    Case giEXPRVALUE_NUMERIC
      pvResult = mdblNumericValue
    Case giEXPRVALUE_LOGIC
      pvResult = mfLogicValue
    Case giEXPRVALUE_DATE
      pvResult = mdtDateValue
    Case Else
      fOK = False
  End Select
    
TidyUpAndExit:
  EvaluateComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

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
  sCode = ""
  
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      sCode = "'" & Replace(msCharacterValue, "'", "''") & "'"
    Case giEXPRVALUE_NUMERIC
      sCode = Trim(Str(mdblNumericValue))
    Case giEXPRVALUE_LOGIC
      sCode = IIf(mfLogicValue, "1", "0")
    Case giEXPRVALUE_DATE
      'JPD 20041118 Fault 8231
      'MH20010201 Fault 1576
      'sCode = "convert(datetime, '" & Format(mdtDateValue, "MM/dd/yyyy") & "')"
      'sCode = IIf(IsNull(mdtDateValue), "null", "convert(datetime, '" & Format(mdtDateValue, "MM/dd/yyyy") & "')")
      sCode = IIf(IsNull(mdtDateValue), "null", "convert(datetime, '" & Replace(Format(mdtDateValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')")
  End Select

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
    Printer.Print "Value : " & ComponentDescription
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
  
  'JPD 20041118 Fault 8231
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type," & _
    " valueType, valueCharacter, valueNumeric, valueLogic, valuedate)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_VALUE)) & "," & _
    " " & Trim(Str(miType)) & "," & _
    " '" & Replace(msCharacterValue, "'", "''") & "'," & _
    " " & Trim(Str(mdblNumericValue)) & "," & _
    " " & IIf(mfLogicValue, "1", "0") & "," & _
    " " & IIf(IsNull(mdtDateValue), "null", "'" & Replace(Format(mdtDateValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'") & ")"
  
    'MH20010201 Fault 1576
    '" '" & Format(mdtDateValue, "MM/dd/yyyy") & "')"
  
  gADOCon.Execute sSQL, , adCmdText

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    COAMsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
  End If
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objValueCopy As New clsExprValue
  
  ' Copy the component's basic properties.
  With objValueCopy
    .ReturnType = miType
    .Value = Value
  End With
  
  Set CopyComponent = objValueCopy
   
  ' Disassociate object variables.
  Set objValueCopy = Nothing
   
End Function

Public Property Get ComponentType() As Integer
  ' Return the component type.
  ComponentType = giCOMPONENT_VALUE
  
End Property

Public Property Get Value() As Variant
  ' Return the value property.
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      Value = msCharacterValue
    Case giEXPRVALUE_NUMERIC
      Value = mdblNumericValue
    Case giEXPRVALUE_LOGIC
      Value = mfLogicValue
    Case giEXPRVALUE_DATE
      Value = mdtDateValue
    Case Else
      Value = ""
  End Select
  
End Property

Public Property Get ComponentDescription() As String
  ' Return the component description.
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      ComponentDescription = Chr(34) & msCharacterValue & Chr(34)
    Case giEXPRVALUE_NUMERIC
      'MH20010130 Fault 1610
      'ComponentDescription = Trim(Str(mdblNumericValue))
      ComponentDescription = datGeneral.ConvertNumberForDisplay(mdblNumericValue)
    Case giEXPRVALUE_LOGIC
      ComponentDescription = IIf(mfLogicValue, "True", "False")
    Case giEXPRVALUE_DATE
      'MH20010201 Fault 1576
      'ComponentDescription = Format(mdtDateValue, "Long Date")
      ComponentDescription = IIf(IsNull(mdtDateValue), "Empty Date", Format(mdtDateValue, "Long Date"))
    Case Else
      ComponentDescription = ""
  End Select

End Property

Public Property Let Value(ByVal pvNewValue As Variant)
  ' Set the value property.
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      msCharacterValue = pvNewValue
    Case giEXPRVALUE_NUMERIC
      mdblNumericValue = pvNewValue
    Case giEXPRVALUE_LOGIC
      mfLogicValue = pvNewValue
    Case giEXPRVALUE_DATE
      mdtDateValue = pvNewValue
  End Select

End Property

Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property

Public Property Get ReturnType() As Integer
  ' Return the value type property.
  ReturnType = miType

End Property

Public Property Let ReturnType(ByVal piNewValue As Integer)
  ' Set the value type property.
  miType = piNewValue

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

Public Function UDFCode(psRuntimeCode() As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  
  UDFCode = True

End Function


