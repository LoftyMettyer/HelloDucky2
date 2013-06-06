VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprTableLookup"
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

Private miTableID As Integer
Private miColumnID As Integer

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
  MsgBox "Error checking for cyclic expressions.", _
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
      pvResult = msCharacterValue
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
      'JPD 20041117 Fault 9484
      'MH20010201 Fault 1576
      'sCode = "convert(datetime, '" & Format(mdtDateValue, "MM/dd/yyyy") & "')"
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
  
  Dim sSQL As String
  Dim rsInfo As Recordset
  
  Dim sSrchColumnName As String
  Dim sSrchTableName As String
  Dim sSrchColumnId As String
  
  fOK = True
  
  ' Position the printing.
  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    'Printer.Print ComponentDescription
    Printer.Print "Table Lookup : " & ComponentDescription
    
    ' Get the column and table name.
    sSQL = "SELECT ASRSysColumns.ColumnID, ASRSysColumns.ColumnName, ASRSysTables.TableName" & _
      " FROM ASRSysColumns" & _
      " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
      " WHERE ASRSysColumns.columnID = " & miColumnID
      
    'MH20020702 datGeneral does not exist in SECMGR and will not compiled.
    'Changed this just to do build!
    'Set rsInfo = datGeneral.GetRecords(sSQL)
    Set rsInfo = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    With rsInfo
        If Trim(Str(!ColumnID)) = miColumnID Then
            sSrchColumnId = !ColumnID
            sSrchColumnName = !ColumnName
            sSrchTableName = !TableName
        End If
        .Close
    End With
    
    Set rsInfo = Nothing
    
    .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Table Name : " & sSrchTableName
    
    .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Column Name : " & sSrchColumnName
    
    .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    'JPD 20030411 Fault 5232
    'Printer.Print "Default Value : " & msCharacterValue
    Printer.Print "Value : " & ComponentDescription
  End With
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  'JPD 20041117 Fault 9484
  sSQL = "INSERT INTO ASRSysExprComponents" & _
    " (componentID, exprID, type," & _
    " valueType, valueCharacter, valueNumeric, valueLogic, valuedate," & _
    " LookupTableID, LookupColumnID)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_TABLEVALUE)) & "," & _
    " " & Trim(Str(miType)) & "," & _
    " '" & Replace(msCharacterValue, "'", "''") & "'," & _
    " " & Trim(Str(mdblNumericValue)) & "," & _
    " " & IIf(mfLogicValue, "1", "0") & "," & _
    " " & IIf(IsNull(mdtDateValue), "null", "'" & Replace(Format(mdtDateValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'") & _
    ", " & Trim(Str(miTableID)) & ", " & Trim(Str(miColumnID)) & ")"
  
  gADOCon.Execute sSQL, , adCmdText

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
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
  Dim objValueCopy As New clsExprTableLookup
  
  ' Copy the component's basic properties.
  With objValueCopy
    .ReturnType = miType
    .Value = Value
    .ColumnID = ColumnID
    .TableID = TableID
  End With
  
  Set CopyComponent = objValueCopy
   
  ' Disassociate object variables.
  Set objValueCopy = Nothing
   
End Function

Public Property Get ComponentType() As Integer
  ' Return the component type.
  ComponentType = giCOMPONENT_TABLEVALUE

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
      ComponentDescription = ConvertNumberForDisplay(mdblNumericValue)
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




Public Property Get ColumnID() As Integer

    ColumnID = miColumnID

End Property

Public Property Let ColumnID(ByVal iNewValue As Integer)

    miColumnID = iNewValue

End Property

Public Property Get TableID() As Integer

    TableID = miTableID

End Property

Public Property Let TableID(ByVal iNewValue As Integer)

    miTableID = iNewValue

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



