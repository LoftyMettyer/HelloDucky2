VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprTableValue"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private miType As ExpressionValueTypes
Private mdblNumericValue As Double
Private msCharacterValue As String
Private mfLogicValue As Boolean
Private mdtDateValue As Date

' Class handling variables.
Private mobjBaseComponent As CExprComponent

Private miTableID As Integer
Private miColumnID As Integer


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
  ' Do nothing for table values.
  
End Sub



Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  ' Do nothing for table values.
  
End Sub




Public Property Get ComponentDescription() As String
  ' Return the component description.
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      ComponentDescription = Chr(34) & msCharacterValue & Chr(34)
    Case giEXPRVALUE_NUMERIC
      'MH20010130 Fault 1610
      'ComponentDescription = Trim(Str(mdblNumericValue))
      ComponentDescription = Trim(UI.ConvertNumberForDisplay(mdblNumericValue))
    Case giEXPRVALUE_LOGIC
      ComponentDescription = IIf(mfLogicValue, "True", "False")
    Case giEXPRVALUE_DATE
      ComponentDescription = Format(mdtDateValue, "Long Date")
    Case Else
      ComponentDescription = ""
  End Select

End Property

Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the component type.
  ComponentType = giCOMPONENT_TABLEVALUE
  
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


Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objValueCopy As New CExprTableValue
  
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

Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objValueClone As New CExprValue
  
  fOK = True
  
  ' Copy the component's basic properties.
  With objValueClone
    .ReturnType = miType
    .Value = Value
  End With
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objValueClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objValueClone = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function ConstructComponent() As Boolean
  ' Read the Value component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean

  fReadOk = False
  
  With recCompEdit
    ' Read the value component record
    miType = !ValueType
      
    Select Case miType
      Case giEXPRVALUE_CHARACTER
        msCharacterValue = IIf(IsNull(!valueCharacter), "", !valueCharacter)
          
      Case giEXPRVALUE_NUMERIC
        mdblNumericValue = IIf(IsNull(!valueNumeric), 0, !valueNumeric)
    
      Case giEXPRVALUE_LOGIC
        mfLogicValue = IIf(IsNull(!valueLogic), True, !valueLogic)
      
      Case giEXPRVALUE_DATE
        mdtDateValue = IIf(IsNull(!valuedate), Date, !valuedate)
    End Select
  
    ' JDM - 22/08/01 - Fault 2042 - Not remembering it's a lookup table
    miTableID = !LookupTableID
    miColumnID = !LookupColumnID
    
  End With
   
  fReadOk = True

TidyUpAndExit:
  ConstructComponent = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function


Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the value type property.
  ReturnType = miType

End Property

Public Property Let ReturnType(ByVal piNewValue As ExpressionValueTypes)
  ' Set the value type property.
  miType = piNewValue

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





Private Sub Class_Initialize()

  ' Initialize properties.
  msCharacterValue = ""
  mdblNumericValue = 0
  mfLogicValue = True
  mdtDateValue = Date
  
End Sub











Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
    
  fOK = True
  
  With recCompEdit
    .AddNew
    !ComponentID = mobjBaseComponent.ComponentID
    !ExprID = mobjBaseComponent.ParentExpression.ExpressionID
    
    ' JDM - 22/08/01 - Fault 2042 - Not remembering it's a lookup table
    !Type = giCOMPONENT_TABLEVALUE
    '!Type = giCOMPONENT_VALUE
    !ValueType = miType

    !LookupTableID = miTableID
    !LookupColumnID = miColumnID
      
    Select Case miType
      Case giEXPRVALUE_CHARACTER
        !valueCharacter = msCharacterValue
          
      Case giEXPRVALUE_NUMERIC
        !valueNumeric = mdblNumericValue
  
      Case giEXPRVALUE_LOGIC
        !valueLogic = mfLogicValue
    
      Case giEXPRVALUE_DATE
        !valuedate = mdtDateValue
    End Select
  
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




Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean)
' JPD - NO LONGER REQUIRED.
'  ' Return the Value of the component.
'  Select Case miType
'    Case giEXPRVALUE_CHARACTER
'      EvaluateComponent = msCharacterValue
'    Case giEXPRVALUE_NUMERIC
'      EvaluateComponent = mdblNumericValue
'    Case giEXPRVALUE_LOGIC
'      EvaluateComponent = mfLogicValue
'    Case giEXPRVALUE_DATE
'      EvaluateComponent = mdtDateValue
'    Case Else
'      EvaluateComponent = vbNullString
'  End Select
   
End Function


Public Function StoredProcedureCode(Optional strBaseTable As String) As String
  ' Return a string describing the component in a stored procedure.
  ' ie. Code that declares the component variable, and assigns the required value to it.
  Dim sVarName As String
  Dim sProcCode As String
  
  sVarName = "@component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      sProcCode = "    DECLARE " & sVarName & " nvarchar(MAX)" & vbNewLine & _
        "    SET " & sVarName & " = '" & Replace(msCharacterValue, "'", "''") & "'"
      
    Case giEXPRVALUE_NUMERIC
      sProcCode = "    DECLARE " & sVarName & " float" & vbNewLine & _
        "    SET " & sVarName & " = " & Trim(Str(mdblNumericValue))

    Case giEXPRVALUE_LOGIC
      sProcCode = "    DECLARE " & sVarName & " bit" & vbNewLine & _
        "    SET " & sVarName & " = " & IIf(mfLogicValue, "1", "0")
  
    Case giEXPRVALUE_DATE
      sProcCode = "    DECLARE " & sVarName & " datetime" & vbNewLine & _
        "    SET " & sVarName & " = convert(datetime, '" & Replace(Format(mdtDateValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')"
      
    Case Else
      sProcCode = ""
  End Select

  StoredProcedureCode = sProcCode

End Function

Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
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
      sCode = "convert(datetime, '" & Replace(Format(mdtDateValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')"
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
  
'  Dim fOK As Boolean
'
'  fOK = True
'
'  ' Position the printing.
'  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
'  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
'  Printer.Print ComponentDescription

  Dim fOK As Boolean
  
'  Dim sSQL As String
'  Dim rsInfo As Recordset
  
  Dim sSrchColumnName As String
  Dim sSrchTableName As String
  Dim sSrchColumnId As String
  
  fOK = True
  
  glngBottom = CalculateBottomOfPage
  
  ' Position the printing.
  With Printer
    .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    'Printer.Print ComponentDescription
    Printer.Print "Table Lookup : " & ComponentDescription
    
  CheckEndOfPage2 glngBottom
    
'    ' Get the column and table name.
'    sSQL = "SELECT ASRSysColumns.ColumnID, ASRSysColumns.ColumnName, ASRSysTables.TableName" & _
'      " FROM ASRSysColumns" & _
'      " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
'      " WHERE ASRSysColumns.columnID = " & miColumnID
      
    'MH20020702 datGeneral does not exist in SYSMGR and will not compiled.
    'Changed this just to do build!
    'Set rsInfo = datGeneral.GetRecords(sSQL)
'    Set rsInfo = rdoCon.OpenResultset(sSQL, rdOpenDynamic, rdConcurReadOnly, rdExecDirect)
'
'    With rsInfo
'        If Trim(Str(!ColumnID)) = miColumnID Then
'            sSrchColumnId = !ColumnID
'            sSrchColumnName = !ColumnName
'            sSrchTableName = !TableName
'        End If
'        .Close
'    End With
'
'    Set rsInfo = Nothing

    'JPD 20030411 Fault 5321
    With recTabEdit
      .Index = "idxTableID"
      .Seek "=", miTableID
      
      If Not .NoMatch Then
        sSrchTableName = !TableName
      Else
        sSrchTableName = "<unknown>"
      End If
    End With
        
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", miColumnID
      
      If Not .NoMatch Then
        sSrchColumnName = .Fields("columnName")
      Else
        sSrchColumnName = "<unknown>"
      End If
    End With
    
    .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Table Name : " & sSrchTableName
    
    CheckEndOfPage2 glngBottom
    
    .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    Printer.Print "Column Name : " & sSrchColumnName
    
    CheckEndOfPage2 glngBottom
    
    .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    'JPD 20030411 Fault 5323
    'Printer.Print "Default Value : " & msCharacterValue
    Printer.Print "Value : " & ComponentDescription
    
    CheckEndOfPage2 glngBottom
  End With
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sClipboardMessage As String
  
  fOK = True
  
  sClipboardMessage = String(piLevel, vbTab) & ComponentDescription & vbNewLine
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
