VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExprPromptedValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private msPrompt As String
Private miType As ExpressionValueTypes
Private miReturnSize As Integer
Private miReturnDecimals As Integer
Private msFormat As String
Private mlngLookupColumnID As Long

Private msDefaultCharacterValue As String
Private mdblDefaultNumericValue As Double
Private mfDefaultLogicValue As Boolean
Private mdtDefaultDateValue As Date
Private miDefaultDateType As Integer

' Class handling variables.
Private mobjBaseComponent As clsExprComponent
Private mvLastEvaluatedValue As Variant

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






Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfApplyPermissions As Boolean, _
  pfValidating As Boolean, _
  Optional plngFixedExprID As Long, _
  Optional psFixedSQLCode As String) As Boolean
  ' Return the SQL code for the component.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fInProgress As Boolean
  Dim sCode As String
  Dim frmPrompt As frmExprPromptedValue
  Dim vDefValue As Variant
  
  fOK = True
  sCode = ""
   
  ' Do not display the prompt form if we are just validating the expression.
  
  'MH20030815 Fault 5660
  'Don't pop up with prompted values if we are in batch mode...
  'If pfValidating Then
  'If pfValidating Or gblnBatchMode Then
  If pfValidating Or gblnBatchJobsOnly Then
    
    vDefValue = DefaultValue
  
    Select Case ReturnType
      Case giEXPRVALUE_CHARACTER
        
        'MH20061214 Fault 11806
        'sCode = "'" & IIf(IsEmpty(vDefValue) Or VarType(vDefValue) <> vbString, gsDUMMY_CHARACTER, vDefValue) & "'"
        sCode = IIf(IsEmpty(vDefValue) Or VarType(vDefValue) <> vbString, gsDUMMY_CHARACTER, vDefValue)
        sCode = "'" & Replace(sCode, "'", "''") & "'"
        
      Case giEXPRVALUE_NUMERIC
        sCode = Trim(Str(IIf(IsEmpty(vDefValue) Or VarType(vDefValue) <> vbDouble, gsDUMMY_NUMERIC, vDefValue)))
      Case giEXPRVALUE_LOGIC
        sCode = IIf(IIf(IsEmpty(vDefValue) Or VarType(vDefValue) <> vbBoolean, gsDUMMY_LOGIC, vDefValue), "1", "0")
      Case giEXPRVALUE_DATE
        'JPD 20041118 Fault 8231
        'sCode = "convert(datetime, '" & Format(IIf(IsEmpty(vDefValue) Or Not IsDate(vDefValue), CDate(gsDUMMY_DATE), vDefValue), "MM/dd/yyyy") & "')"
        sCode = "convert(datetime, '" & Replace(Format(IIf(IsEmpty(vDefValue) Or Not IsDate(vDefValue), CDate(gsDUMMY_DATE), vDefValue), "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')"
    End Select
  Else
'    ' Hide the progress bar if it is visible.
'    fInProgress = False
'    If Not gobjProgress Is Nothing Then
'      If gobjProgress.Visible Then
'        fInProgress = True
'        gobjProgress.Visible = False
'      End If
'    End If
'
'    Set frmPrompt = New frmExprPromptedValue
'    Set frmPrompt.Component = Me
'    frmPrompt.Show vbModal
'    mvLastEvaluatedValue = frmPrompt.Value
'    Set frmPrompt = Nothing
  
    Set frmPrompt = New frmExprPromptedValue
    Set frmPrompt.Component = Me
    
    ' Hide the progress bar if it is visible.
    fInProgress = False
    If Not gobjProgress Is Nothing Then
      If gobjProgress.Visible Then
        fInProgress = True
        gobjProgress.Visible = False
        
        '14/08/2001 MH Fault 2680
        'Don't show the prompted value if we cancel on progress bar
        If gobjProgress.Cancelled = False Then
          frmPrompt.Show vbModal
        End If
      
      Else
        frmPrompt.Show vbModal
      
      End If
    Else
      frmPrompt.Show vbModal
    
    End If
    
    mvLastEvaluatedValue = frmPrompt.Value
    Set frmPrompt = Nothing
    
    ' JDM - 16/06/03 - Fault 5974 - Prompt appearing twice - remember last setting
    gvLastPromptedValue = mvLastEvaluatedValue
    
    Select Case ReturnType
      Case giEXPRVALUE_CHARACTER
        sCode = "'" & Replace(CStr(mvLastEvaluatedValue), "'", "''") & "'"
      Case giEXPRVALUE_NUMERIC
        sCode = Trim(Str(mvLastEvaluatedValue))
      Case giEXPRVALUE_LOGIC
        sCode = IIf(mvLastEvaluatedValue, "1", "0")
      Case giEXPRVALUE_DATE
        'JPD 20041118 Fault 8231
        sCode = "convert(datetime, '" & Replace(Format(CDate(mvLastEvaluatedValue), "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')"
    End Select
    
    If fInProgress Then
      gobjProgress.Visible = True
    End If
  End If
  
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
    
    Dim strFormatUsedHere As String
    Dim strDefaultCharacterValue As String
    Dim strTypeUsedHere As String
    Dim sSrchColumnName As String
    Dim sSrchTableName As String
    Dim sSrchColumnId As String
    
    fOK = True
    
    'What data type used for Prompted Value?
    'The type will determine the details printed
    Select Case miType
        Case giEXPRVALUE_CHARACTER
            strTypeUsedHere = "Character"
            strDefaultCharacterValue = msDefaultCharacterValue
            strFormatUsedHere = IIf(msFormat = "", "<none>", msFormat)
        Case giEXPRVALUE_NUMERIC
            strTypeUsedHere = "Numeric"
            strDefaultCharacterValue = mdblDefaultNumericValue
            'strFormatUsedHere = IIf(msFormat = "", "<none>", msFormat)
        Case giEXPRVALUE_LOGIC
            strTypeUsedHere = "Logic"
            strDefaultCharacterValue = CStr(mfDefaultLogicValue)
        Case giEXPRVALUE_DATE
            strTypeUsedHere = "Date"
            'NHRD26022004 Fault 4133
            'strDefaultCharacterValue = mdtDefaultDateValue
            strDefaultCharacterValue = DefaultValue
        Case giEXPRVALUE_TABLEVALUE
            strTypeUsedHere = "Lookup-table-value"
        Case Else
            strTypeUsedHere = "NON-STANDARD DATATYPE"
    End Select
  
    ' Position the printing.
    With Printer
        .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
        .CurrentY = .CurrentY + giPRINT_YSPACE
        Printer.Print "Prompted Value: " & ComponentDescription
    
        .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
        .CurrentY = .CurrentY + giPRINT_YSPACE
        Printer.Print "Default Value : " & strDefaultCharacterValue
        
        If strTypeUsedHere = "Character" Or strTypeUsedHere = "Numeric" Then
          If strTypeUsedHere = "Character" Then ' You can't format a numeric prompted value
            .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
            .CurrentY = .CurrentY + giPRINT_YSPACE
            Printer.Print "Format/Masking : " & strFormatUsedHere
          End If
          
          .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
          .CurrentY = .CurrentY + giPRINT_YSPACE
          
          'NHRD - 05042003 - Fault 4135
          If miReturnDecimals > 0 Then
            Printer.Print "Return Size : " & miReturnSize & ", Decimals : " & miReturnDecimals
          Else
            Printer.Print "Return Size : " & miReturnSize
          End If
        End If
        
        If strTypeUsedHere = "Lookup-table-value" Then
            ' Get the column and table name.
            sSQL = "SELECT ASRSysColumns.ColumnID, ASRSysColumns.ColumnName, ASRSysTables.TableName" & _
              " FROM ASRSysColumns" & _
              " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
              " WHERE ASRSysColumns.columnID = " & mlngLookupColumnID
              
            Set rsInfo = datGeneral.GetRecords(sSQL)
            
            With rsInfo
                If Trim(Str(!ColumnID)) = mlngLookupColumnID Then
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
        End If
    
        .CurrentX = giPRINT_XINDENT + giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
        .CurrentY = .CurrentY + giPRINT_YSPACE
        Printer.Print "Type : " & strTypeUsedHere
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
    " (componentID, exprID, type, promptDescription," & _
    " valueType, promptSize, promptDecimals, promptMask," & _
    " valueCharacter, valueNumeric, valueLogic, valueDate, fieldColumnID,PromptDateType)" & _
    " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & _
    " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & _
    " " & Trim(Str(giCOMPONENT_PROMPTEDVALUE)) & "," & _
    " '" & Replace(Trim(msPrompt), "'", "''") & "'," & _
    " " & Trim(Str(miType)) & "," & _
    " " & Trim(Str(miReturnSize)) & "," & _
    " " & Trim(Str(miReturnDecimals)) & "," & _
    " '" & Replace(Trim(msFormat), "'", "''") & "'," & _
    " '" & Replace(Trim(msDefaultCharacterValue), "'", "''") & "'," & _
    " " & Trim(Str(mdblDefaultNumericValue)) & "," & _
    " " & IIf(mfDefaultLogicValue, "1", "0") & "," & _
    " '" & Replace(Format(mdtDefaultDateValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'," & _
    " " & Trim(Str(mlngLookupColumnID)) & "," & _
    " " & Trim(Str(miDefaultDateType)) & _
    ")"
    
    
    
  gADOCon.Execute sSQL, , adCmdText

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Public Function CopyComponent()
  ' Copies the selected component.
  ' When editing a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim dtDate As Date
  Dim objPromptedValueCopy As New clsExprPromptedValue

  ' Copy the component's basic properties.
  With objPromptedValueCopy
    .Prompt = msPrompt
    .ValueType = miType
    .ReturnSize = miReturnSize
    .ReturnDecimals = miReturnDecimals
    .ValueFormat = msFormat
    .DefaultValue = DefaultValue
    .LookupColumn = mlngLookupColumnID
    .DefaultDateType = miDefaultDateType
  End With
  
  Set CopyComponent = objPromptedValueCopy
   
  ' Disassociate object variables.
  Set objPromptedValueCopy = Nothing
   
End Function

Public Property Get LookupColumn() As Long
  ' Return the Lookup Column ID.
  LookupColumn = mlngLookupColumnID

End Property

Public Property Let LookupColumn(ByVal plngNewValue As Long)
  ' Set the Lookup Column ID.
  mlngLookupColumnID = plngNewValue

End Property

Public Property Get ValueFormat() As String
  ' Return the ValueFormat property.
  ValueFormat = msFormat

End Property

Public Property Let ValueFormat(ByVal psNewValue As String)
  ' Set the ValueFormat property.
  msFormat = psNewValue

End Property

Public Property Get DefaultValue() As Variant

  ' Return the default value property.
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      DefaultValue = msDefaultCharacterValue
    Case giEXPRVALUE_NUMERIC
      DefaultValue = mdblDefaultNumericValue
    Case giEXPRVALUE_LOGIC
      DefaultValue = mfDefaultLogicValue
    Case giEXPRVALUE_DATE
      
      ' What type of default date is it?
      Select Case miDefaultDateType
        Case 0
          DefaultValue = mdtDefaultDateValue
        Case 1
          DefaultValue = Format(Now(), "dd/mm/yyyy")
        Case 2
          DefaultValue = Now - Day(Now) + 1
          DefaultValue = Format(DefaultValue, "dd/mm/yyyy")
        Case 3
          DefaultValue = DateSerial(Year(Now), Month(Now), DateSerial(Year(Now), Month(Now) + 1, 1) _
              - DateSerial(Year(Now), Month(Now), 1))
        Case 4
          DefaultValue = DateSerial(Year(Now), 1, 1)
        Case 5
          DefaultValue = DateSerial(Year(Now), 12, 31)
      End Select
      
    Case giEXPRVALUE_TABLEVALUE
      DefaultValue = msDefaultCharacterValue
    Case Else
      DefaultValue = ""
  End Select
  
End Property
Public Property Get ComponentType() As Integer
  ' Return the Prompted Value component type.
  ComponentType = giCOMPONENT_PROMPTEDVALUE
  
End Property

Public Property Get ReturnType() As Integer
  ' Return the return type property.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iType As ExpressionValueTypes
  Dim sSQL As String
  Dim rsColumn As Recordset
  
  fOK = True
  
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      iType = giEXPRVALUE_CHARACTER
      
    Case giEXPRVALUE_NUMERIC
      iType = giEXPRVALUE_NUMERIC
      
    Case giEXPRVALUE_LOGIC
      iType = giEXPRVALUE_LOGIC
      
    Case giEXPRVALUE_DATE
      iType = giEXPRVALUE_DATE
      
    Case giEXPRVALUE_TABLEVALUE
      ' Get the lookup column's return type.
      sSQL = "SELECT dataType" & _
        " FROM ASRSysColumns" & _
        " WHERE columnID = " & Trim(Str(mlngLookupColumnID))
      Set rsColumn = datGeneral.GetRecords(sSQL)
      With rsColumn
    
        fOK = Not (.EOF And .BOF)
      
        If fOK Then
          Select Case !DataType
            Case sqlNumeric, sqlInteger
              iType = giEXPRVALUE_NUMERIC
            Case sqlDate
              iType = giEXPRVALUE_DATE
            Case sqlVarChar, sqlLongVarChar
              iType = giEXPRVALUE_CHARACTER
            Case sqlBoolean
              iType = giEXPRVALUE_LOGIC
            Case sqlOle
              iType = giEXPRVALUE_OLE
            Case sqlVarBinary
              iType = giEXPRVALUE_PHOTO
            Case Else
              fOK = False
          End Select
        End If
      
        .Close
      End With
      Set rsColumn = Nothing
      
    Case Else
      fOK = False
  End Select
  
TidyUpAndExit:
  If fOK Then
    ReturnType = iType
  Else
    ReturnType = giEXPRVALUE_UNDEFINED
  End If
  Exit Property
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Property


Public Property Let DefaultValue(ByVal pvNewValue As Variant)
  ' Set the value property.
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      msDefaultCharacterValue = pvNewValue
    Case giEXPRVALUE_NUMERIC
      mdblDefaultNumericValue = pvNewValue
    Case giEXPRVALUE_LOGIC
      mfDefaultLogicValue = pvNewValue
    Case giEXPRVALUE_DATE
      If IsNull(pvNewValue) Then
        mdtDefaultDateValue = 0
      Else
        mdtDefaultDateValue = pvNewValue
      End If
    Case giEXPRVALUE_TABLEVALUE
      msDefaultCharacterValue = pvNewValue
  End Select

End Property

Public Property Get ReturnDecimals() As Integer
  ' Return the return number of decimals.
  ReturnDecimals = miReturnDecimals
  
End Property

Public Property Let ReturnDecimals(ByVal piNewValue As Integer)
  ' Set the return number of decimals.
  miReturnDecimals = piNewValue

End Property

Public Property Get ValueType() As Integer
  ' Return the type property.
  ValueType = miType

End Property

Public Property Get ReturnSize() As Integer
  ' Return the return size.
  ReturnSize = miReturnSize

End Property

Public Property Let ReturnSize(ByVal piNewValue As Integer)
  ' Set the return size.
  miReturnSize = piNewValue

End Property

Public Property Let ReturnType(ByVal piNewValue As Integer)
  ' Set the return type.
  miType = piNewValue

End Property


Public Property Let ValueType(ByVal piNewValue As Integer)
  ' Set the type property.
  miType = piNewValue

End Property

Public Property Get ComponentDescription() As String
  ' Return the component description.
  Dim sDescription As String
  
  sDescription = msPrompt & " : "
  
  Select Case miType
    Case giEXPRVALUE_CHARACTER
      sDescription = sDescription & "<string>"
      
    Case giEXPRVALUE_NUMERIC
      sDescription = sDescription & "<numeric>"
      
    Case giEXPRVALUE_LOGIC
      sDescription = sDescription & "<logic>"
      
    Case giEXPRVALUE_DATE
      sDescription = sDescription & "<date>"
      
    Case giEXPRVALUE_TABLEVALUE
      sDescription = sDescription & "<lookup value>"
  End Select

  ComponentDescription = sDescription
  
End Property




Public Property Get BaseComponent() As clsExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(ByVal pobjNewValue As clsExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property

Public Property Get Prompt() As String
  ' Return the Prompt property.
  Prompt = msPrompt

End Property
Public Property Let Prompt(ByVal psNewValue As String)
  ' Set the Prompt property.
  msPrompt = psNewValue

End Property

' Read the default date type
Public Property Get DefaultDateType() As Integer
  DefaultDateType = miDefaultDateType
End Property

' Set the default date type
Public Property Let DefaultDateType(ByVal piNewValue As Integer)
  miDefaultDateType = piNewValue
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

