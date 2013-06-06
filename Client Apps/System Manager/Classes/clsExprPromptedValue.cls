VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprPromptedValue"
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

' Class handling variables.
Private mobjBaseComponent As CExprComponent
Private mvLastEvaluatedValue As Variant






Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Position the printing.
  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  Printer.Print ComponentDescription
  
TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function StoredProcedureCode() As String
  ' Prompted Values aren't used in calculation and so
  ' are not required in stored procedures.
  StoredProcedureCode = ""

End Function



Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the Prompted Value component type.
  ComponentType = giCOMPONENT_PROMPTEDVALUE
  
End Property


Public Function ConstructComponent() As Boolean
  ' Read the Prompted Value component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean
  
  fReadOk = False
  
  ' Read the Prompted Value component record.
  msPrompt = recCompEdit!promptDescription
  miType = recCompEdit!ValueType
  miReturnSize = recCompEdit!promptSize
  miReturnDecimals = recCompEdit!promptDecimals
  msFormat = recCompEdit!promptMask
      
  msDefaultCharacterValue = recCompEdit!valueCharacter
  mdblDefaultNumericValue = recCompEdit!valueNumeric
  mfDefaultLogicValue = recCompEdit!valueLogic
  mdtDefaultDateValue = recCompEdit!valuedate
  
  mlngLookupColumnID = recCompEdit!fieldColumnID

  fReadOk = True
  
TidyUpAndExit:
  ConstructComponent = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function

Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated result of the Prompted Value.
'  Dim vResult As Variant
'  Dim frmPrompt As frmPromptedValue
'  Dim dtDateValue As Date
'
'  ' Display the Prompted value form if required.
'  If pfDoPrompts Then
'    Set frmPrompt = New frmPromptedValue
'    Set frmPrompt.Component = Me
'    frmPrompt.Show vbModal
'    EvaluateComponent = frmPrompt.Value
'    Set frmPrompt = Nothing
'    mvLastEvaluatedValue = EvaluateComponent
'  End If
'
'  If IsNull(mvLastEvaluatedValue) Then
'    Select Case ReturnType
'      Case giEXPRVALUE_CHARACTER
'        mvLastEvaluatedValue = ""
'      Case giEXPRVALUE_NUMERIC
'        mvLastEvaluatedValue = 0
'      Case giEXPRVALUE_LOGIC
'        mvLastEvaluatedValue = False
'      Case giEXPRVALUE_DATE
'        mvLastEvaluatedValue = dtDateValue
'    End Select
'  End If
'
'  EvaluateComponent = mvLastEvaluatedValue
  
End Function

Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the return type property.
  Dim iType As ExpressionValueTypes

  ' Initialise the default return type.
  iType = giEXPRVALUE_UNDEFINED
  
  Select Case miType
    Case giEXPRVALUE_CHARACTER ' Character
      iType = giEXPRVALUE_CHARACTER
      
    Case giEXPRVALUE_NUMERIC ' Numeric
      iType = giEXPRVALUE_NUMERIC
      
    Case giEXPRVALUE_LOGIC ' Logic
      iType = giEXPRVALUE_LOGIC
      
    Case giEXPRVALUE_DATE ' Date
      iType = giEXPRVALUE_DATE
      
    Case giEXPRVALUE_TABLEVALUE ' Table Value
      ' Get the lookup column's return type.
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", mlngLookupColumnID
        
        If Not .NoMatch Then
          Select Case .Fields("dataType")
            Case rdTypeVARCHAR, rdTypeLONGVARCHAR
              iType = giEXPRVALUE_CHARACTER
            Case rdTypeTIMESTAMP
              iType = giEXPRVALUE_DATE
            Case rdTypeINTEGER
              iType = giEXPRVALUE_NUMERIC
            Case rdTypeBIT
              iType = giEXPRVALUE_LOGIC
            Case rdTypeNUMERIC
              iType = giEXPRVALUE_NUMERIC
          End Select
        End If
      End With
  End Select
  
  ReturnType = iType
  
End Property

Public Function RuntimeCode(psRuntimeCode As String, palngSourceTables As Variant, pfValidating As Boolean) As Boolean
  ' Return the SQL code for the component.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sCode As String
  Dim frmPrompt As frmPromptedValue
  
  fOK = True
  sCode = ""
  
  ' Do not display the prompt form if we are just validating the expression.
  If pfValidating Then
    Select Case ReturnType
      Case giEXPRVALUE_CHARACTER
        sCode = "'" & gsDUMMY_CHARACTER & "'"
      Case giEXPRVALUE_NUMERIC
        sCode = Trim(Str(gsDUMMY_NUMERIC))
      Case giEXPRVALUE_LOGIC
        sCode = IIf(gsDUMMY_LOGIC, "1", "0")
      Case giEXPRVALUE_DATE
        sCode = "convert(datetime, '" & Format(CDate(gsDUMMY_DATE), "MM/dd/yyyy") & "')"
    End Select
  Else
    Set frmPrompt = New frmPromptedValue
    Set frmPrompt.Component = Me
    frmPrompt.Show vbModal
    mvLastEvaluatedValue = frmPrompt.Value
    Set frmPrompt = Nothing
  
    Select Case ReturnType
      Case giEXPRVALUE_CHARACTER
        sCode = "'" & Replace(CStr(mvLastEvaluatedValue), "'", "''") & "'"
      Case giEXPRVALUE_NUMERIC
        sCode = Trim(Str(mvLastEvaluatedValue))
      Case giEXPRVALUE_LOGIC
        sCode = IIf(mvLastEvaluatedValue, "1", "0")
      Case giEXPRVALUE_DATE
        sCode = "convert(datetime, '" & Format(CDate(mvLastEvaluatedValue), "MM/dd/yyyy") & "')"
    End Select
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



Public Function WriteComponent()
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
    
  fOK = True
  
  With recCompEdit
    .AddNew
    !ComponentID = mobjBaseComponent.ComponentID
    !exprID = mobjBaseComponent.ParentExpression.ExpressionID
    !Type = giCOMPONENT_PROMPTEDVALUE
    !promptDescription = msPrompt
    !ValueType = miType
    !promptSize = miReturnSize
    !promptDecimals = miReturnDecimals
    !promptMask = msFormat

    !valueCharacter = msDefaultCharacterValue
    !valueNumeric = mdblDefaultNumericValue
    !valueLogic = mfDefaultLogicValue
    !valuedate = mdtDefaultDateValue

    !fieldColumnID = mlngLookupColumnID

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

Public Property Get ValueType() As ExpressionValueTypes
  ' Return the type property.
  ValueType = miType

End Property

Public Property Get ReturnSize() As Integer
  ' Return the return size.
  ReturnSize = miReturnSize

End Property
Public Property Get LookupColumn() As Long
  ' Return the Lookup Column ID.
  LookupColumn = mlngLookupColumnID

End Property
Public Property Get LookupTable() As Long
  ' Return the Lookup Table ID.
  On Error GoTo ErrorTrap
  
  Dim objColumn As Column
  Dim lngTableID As Long

  ' Initialize the default table ID.
  lngTableID = 0
  
  Set objColumn = New Column
  objColumn.ColumnID = mlngLookupColumnID
  If objColumn.ReadColumn Then
    lngTableID = objColumn.TableID
  End If

TidyUpAndExit:
  Set objColumn = Nothing
  LookupTable = lngTableID
  Exit Property
  
ErrorTrap:
  lngTableID = 0
  Resume TidyUpAndExit
  
End Property

Public Property Let ReturnSize(ByVal piNewValue As Integer)
  ' Set the return size.
  miReturnSize = piNewValue

End Property
Public Property Let LookupColumn(ByVal plngNewValue As Long)
  ' Set the Lookup Column ID.
  mlngLookupColumnID = plngNewValue

End Property

Public Property Get DefaultValue() As Variant

  ' Return the default value property.
  Select Case miType
    Case giEXPRVALUE_CHARACTER ' Character
      DefaultValue = msDefaultCharacterValue
    Case giEXPRVALUE_NUMERIC ' Numeric
      DefaultValue = mdblDefaultNumericValue
    Case giEXPRVALUE_LOGIC ' Logic
      DefaultValue = mfDefaultLogicValue
    Case giEXPRVALUE_DATE ' Date
      DefaultValue = mdtDefaultDateValue
    Case giEXPRVALUE_TABLEVALUE ' Table Value
      DefaultValue = msDefaultCharacterValue
    Case Else
      DefaultValue = ""
  End Select
  
End Property
Public Property Let DefaultValue(ByVal pvNewValue As Variant)
  ' Set the value property.
  Select Case miType
    Case giEXPRVALUE_CHARACTER ' Character
      msDefaultCharacterValue = pvNewValue
    Case giEXPRVALUE_NUMERIC ' Numeric
      mdblDefaultNumericValue = pvNewValue
    Case giEXPRVALUE_LOGIC ' Logic
      mfDefaultLogicValue = pvNewValue
    Case giEXPRVALUE_DATE ' Date
      If IsNull(pvNewValue) Then
        mdtDefaultDateValue = 0
      Else
        mdtDefaultDateValue = pvNewValue
      End If
    Case giEXPRVALUE_TABLEVALUE ' Table Value
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



Public Property Let ValueType(ByVal piNewValue As ExpressionValueTypes)
  ' Set the type property.
  miType = piNewValue

End Property


Public Property Get Prompt() As String
  ' Return the Prompt property.
  Prompt = msPrompt

End Property
Public Property Get ValueFormat() As String
  ' Return the Mask property.
  ValueFormat = msFormat

End Property

Public Property Let Prompt(ByVal psNewValue As String)
  ' Set the Prompt property.
  msPrompt = psNewValue

End Property

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editing a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objPromptedValueCopy As New CExprPromptedValue

  ' Copy the component's basic properties.
  With objPromptedValueCopy
    .Prompt = msPrompt
    .ValueType = miType
    .ReturnSize = miReturnSize
    .ReturnDecimals = miReturnDecimals
    .ValueFormat = msFormat
    .DefaultValue = DefaultValue
    .LookupColumn = mlngLookupColumnID
  End With
  
  Set CopyComponent = objPromptedValueCopy
   
  ' Disassociate object variables.
  Set objPromptedValueCopy = Nothing
   
End Function
Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objPromptedValueClone As New CExprPromptedValue

  fOK = True
  
  ' Clone the component's basic properties.
  With objPromptedValueClone
    .Prompt = msPrompt
    .ValueType = miType
    .ReturnSize = miReturnSize
    .ReturnDecimals = miReturnDecimals
    .ValueFormat = msFormat
    .DefaultValue = DefaultValue
    .LookupColumn = mlngLookupColumnID
  End With
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objPromptedValueClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objPromptedValueClone = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
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
  ComponentContainsColumn = (miType = giEXPRVALUE_TABLEVALUE) And _
    (mlngLookupColumnID = plngColumnID)
  
End Function

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
      sDescription = sDescription & "<table value>"
  End Select

  ComponentDescription = sDescription
  
End Property


Public Property Let ValueFormat(ByVal psNewValue As String)
  ' Set the Format property.
  msFormat = psNewValue

End Property

Private Sub Class_Initialize()
  ' Initialize properties.
  msDefaultCharacterValue = ""
  mdblDefaultNumericValue = 0
  mfDefaultLogicValue = True

End Sub


