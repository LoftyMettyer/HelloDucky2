VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExprFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mlngFunctionID As Long
Private msFunctionName As String
Private miReturnType As ExpressionValueTypes
Private msSPName As String
Private msUDFName As String

' Class handling variables.
Private mobjBaseComponent As CExprComponent
Private mcolParameters As Collection
Private mobjBadComponent As CExprComponent

' Definition for expanded/unexpanded status of the component
Private mbExpanded As Boolean

' Workflow parameters
Private miWorkflowRecordSelection As WorkflowRecordSelectorTypes
Private msWorkflowElementIdentifier As String
Private msWorkflowItemIdentifier As String
Private mlngWorkflowRecordTableID As Long

Public Function ContainsExpression(plngExprID As Long) As Boolean
  ' Retrun TRUE if the current expression (or any of its sub expressions)
  ' contains the given expression. This ensures no cyclic expressions get created.
  'JPD 20040504 Fault 8599
  On Error GoTo ErrorTrap
  
  Dim objParameter As CExprComponent
  Dim objSubExpression As CExpression
  
  ContainsExpression = False
  
  For Each objParameter In mcolParameters
    Set objSubExpression = objParameter.Component
    
    ContainsExpression = objSubExpression.ContainsExpression(plngExprID)
        
    'objSubExpression.Terminate   'MH20060727 New code not implemented
    Set objSubExpression = Nothing
    
    If ContainsExpression Then
      Exit For
    End If
  Next objParameter
  'objParameter.Terminate       'MH20060727  New code not implemented
  Set objParameter = Nothing
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error checking for cyclic expressions.", _
    vbOKOnly + vbExclamation, App.ProductName
  Err = False
  Resume TidyUpAndExit
    
End Function








'Private mlngBottom As Long
Public Property Get ExpandedNode() As Boolean
    'Return whether this node is expanded or not
    ExpandedNode = mbExpanded

End Property

Public Property Let ExpandedNode(ByVal bNewValue As Boolean)
    'Set whether this component node is expanded or not
    mbExpanded = bNewValue

End Property



Private Function ValidateFunctionParameters(plngFunctionID As Variant, piResultType As ExpressionValueTypes, _
  Optional piParam1Type As Integer, Optional piParam2Type As Integer, _
  Optional piParam3Type As Integer, Optional piParam4Type As Integer, _
  Optional piParam5Type As Integer, Optional piParam6Type As Integer) As Boolean
  ' Validate the given function with the given parameters.
  ' Return the result type in the piResultType parameter.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  
  fOK = True

  ' Get the parameter types.
  piParam1Type = IIf(IsMissing(piParam1Type), giEXPRVALUE_UNDEFINED, piParam1Type)
  piParam2Type = IIf(IsMissing(piParam2Type), giEXPRVALUE_UNDEFINED, piParam2Type)
  piParam3Type = IIf(IsMissing(piParam3Type), giEXPRVALUE_UNDEFINED, piParam3Type)
  piParam4Type = IIf(IsMissing(piParam4Type), giEXPRVALUE_UNDEFINED, piParam4Type)
  piParam5Type = IIf(IsMissing(piParam5Type), giEXPRVALUE_UNDEFINED, piParam5Type)
  piParam6Type = IIf(IsMissing(piParam6Type), giEXPRVALUE_UNDEFINED, piParam6Type)

  ' Validate the parameter types for the given function.
  Select Case plngFunctionID
    Case 1 ' SYSTEM DATE
      fOK = True
      piResultType = giEXPRVALUE_DATE

    Case 2 ' CONVERT TO UPPERCASE
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 3 ' CONVERT NUMERIC TO STRING
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 4 ' IF, THEN, ELSE
      fOK = (piParam1Type = giEXPRVALUE_LOGIC) And _
        (piParam2Type = piParam3Type)
      piResultType = piParam2Type

    Case 5 ' REMOVE LEADING AND TRAINING SPACES
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 6 ' EXTRACT CHARACTERS FROM THE LEFT
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 7 ' LENGTH OF STRING
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 8 ' CONVERT TO LOWERCASE
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 9 ' MAXIMUM
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 10 ' MINIMUM
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 11 ' SEARCH FOR CHARACTER STRING
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 12 ' CAPITALIZE INITIALS
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 13 'EXTRACT CHARACTERS FROM THE RIGHT
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 14 ' EXTRACT PART OF A STRING
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC) And _
        (piParam3Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_CHARACTER

    Case 15 ' SYSTEM TIME
      fOK = True
      piResultType = giEXPRVALUE_CHARACTER

    Case 16 ' IS FIELD EMPTY
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) Or _
        (piParam1Type = giEXPRVALUE_NUMERIC) Or _
        (piParam1Type = giEXPRVALUE_LOGIC) Or _
        (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 17 ' CURRENT USER
      fOK = True
      piResultType = giEXPRVALUE_CHARACTER

    Case 18 ' WHOLE YEARS UNTIL CURRENT DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 19 ' REMAINING MONTHS SINCE WHOLE YEARS
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 20 ' INITIALS FROM FORENAMES
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 21 ' FIRST NAME FROM FORENAMES
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 22 ' WEEKDAYS FROM START AND END DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 23 ' ADD MONTHS TO DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_DATE

    Case 24 ' ADD YEARS TO DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_DATE

    Case 25 ' CONVERT CHARACTER TO NUMERIC
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 26 ' WHOLE MONTHS BETWEEN TWO DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 27 ' PARENTHESESES
      fOK = True
      piResultType = piParam1Type

    Case 28 ' DAY OF THE WEEK
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 29 ' NUMBER OF WORKING DAYS PER WEEK
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 30 ' ABSENCE DURATION
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_CHARACTER) And _
        (piParam3Type = giEXPRVALUE_DATE) And _
        (piParam4Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

'      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
'        (piParam2Type = giEXPRVALUE_CHARACTER) And _
'        (piParam3Type = giEXPRVALUE_DATE) And _
'        (piParam4Type = giEXPRVALUE_CHARACTER) And _
'        (piParam5Type = giEXPRVALUE_CHARACTER) And _
'        (piParam6Type = giEXPRVALUE_BYREF_DATE)
'      piResultType = giEXPRVALUE_NUMERIC

    Case 31 ' ROUND DOWN TO NEAREST WHOLE NUMBER
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 32 ' YEAR OF DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 33 ' MONTH OF DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 34 ' DAY OF DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 35 ' NICE DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_CHARACTER

    Case 36 ' NICE TIME
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case 37 ' ROUND DATE TO START OF NEAREST MONTH
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    Case 38 ' IS BETWEEN
      fOK = ((piParam1Type = giEXPRVALUE_DATE) And (piParam2Type = giEXPRVALUE_DATE) And (piParam3Type = giEXPRVALUE_DATE)) Or _
       ((piParam1Type = giEXPRVALUE_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC) And (piParam3Type = giEXPRVALUE_NUMERIC))
      piResultType = giEXPRVALUE_LOGIC

    Case 39 ' SERVICE YEARS
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 40 ' SERVICE MONTHS
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC

    Case 41 ' STATUTORY REDUNDANCY PAY
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam3Type = giEXPRVALUE_DATE) And _
        (piParam4Type = giEXPRVALUE_NUMERIC) And _
        (piParam5Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 42 ' GET FIELD FROM DATABASE RECORD
      fOK = (((piParam1Type = giEXPRVALUE_BYREF_CHARACTER) And (piParam2Type = giEXPRVALUE_CHARACTER)) Or _
        ((piParam1Type = giEXPRVALUE_BYREF_NUMERIC) And (piParam2Type = giEXPRVALUE_NUMERIC)) Or _
        ((piParam1Type = giEXPRVALUE_BYREF_LOGIC) And (piParam2Type = giEXPRVALUE_LOGIC)) Or _
        ((piParam1Type = giEXPRVALUE_BYREF_DATE) And (piParam2Type = giEXPRVALUE_DATE))) And _
        ((piParam3Type = giEXPRVALUE_BYREF_CHARACTER) Or _
        (piParam3Type = giEXPRVALUE_BYREF_NUMERIC) Or _
        (piParam3Type = giEXPRVALUE_BYREF_LOGIC) Or _
        (piParam3Type = giEXPRVALUE_BYREF_DATE))
      piResultType = (piParam3Type - giEXPRVALUE_BYREF_OFFSET)

    Case 43 ' GET UNIQUE CODE
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 44 ' ADD DAYS TO DATE
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_DATE

    Case 45 ' DAYS BETWEEN TWO DATES
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC
      
    Case 46 ' WORKING DAYS BETWEEN TWO DATES (INC BHOLS)
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC
      
'      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
'        (piParam2Type = giEXPRVALUE_DATE) And _
'        (piParam3Type = giEXPRVALUE_CHARACTER) And _
'        (piParam4Type = giEXPRVALUE_CHARACTER)
'      piResultType = giEXPRVALUE_NUMERIC
      
    Case 47 ' ABSENCE BETWEEN TWO DATES (INC BHOLS)
      
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam3Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC
      
'      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
'        (piParam2Type = giEXPRVALUE_DATE) And _
'        (piParam3Type = giEXPRVALUE_CHARACTER) And _
'        (piParam4Type = giEXPRVALUE_CHARACTER) And _
'        (piParam5Type = giEXPRVALUE_CHARACTER)
'      piResultType = giEXPRVALUE_NUMERIC
    
    Case 48 ' ROUND UP TO NEAREST WHOLE NUMBER
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    Case 49 ' ROUND TO NEAREST NUMBER
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_NUMERIC)
      piResultType = giEXPRVALUE_NUMERIC

    'MH20010515
    Case 50 ' IS OVERNIGHT PROCESS
      fOK = True
      piResultType = giEXPRVALUE_LOGIC

    Case 51 ' CONVERT CURRENCY
      fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
        (piParam2Type = giEXPRVALUE_CHARACTER) And _
        (piParam3Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 52 ' Field Last Changed Date
      fOK = (piParam1Type = giEXPRVALUE_BYREF_CHARACTER Or _
        piParam1Type = giEXPRVALUE_BYREF_NUMERIC Or _
        piParam1Type = giEXPRVALUE_BYREF_LOGIC Or _
        piParam1Type = giEXPRVALUE_BYREF_DATE)
      piResultType = giEXPRVALUE_DATE
    
    Case 53 ' Field changed between two dates
      fOK = ((piParam1Type = giEXPRVALUE_BYREF_CHARACTER Or _
        piParam1Type = giEXPRVALUE_BYREF_NUMERIC Or _
        piParam1Type = giEXPRVALUE_BYREF_LOGIC Or _
        piParam1Type = giEXPRVALUE_BYREF_DATE)) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 54 'Whole years between two dates
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
            (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_NUMERIC
    
    ' JPD20021121 Fault 3177
    Case 55 ' First Day of Month
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021121 Fault 3177
    Case 56 ' Last Day of Month
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021121 Fault 3177
    Case 57 ' First Day of Year
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021121 Fault 3177
    Case 58 ' Last Day of Year
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_DATE

    ' JPD20021129 Fault 4337
    Case 59 ' NAME OF MONTH
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_CHARACTER
    
    ' JPD20021129 Fault 4337
    Case 60 ' NAME OF DAY
      fOK = (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_CHARACTER
    
    ' JPD20021129 Fault 3606
    Case 61 ' IS FIELD POPULATED
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) Or _
        (piParam1Type = giEXPRVALUE_NUMERIC) Or _
        (piParam1Type = giEXPRVALUE_LOGIC) Or _
        (piParam1Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 62 'PARENTAL LEAVE ENTITLEMENT
      fOK = True
      piResultType = giEXPRVALUE_NUMERIC

    Case 63 'PARENTAL LEAVE TAKEN
      fOK = True
      piResultType = giEXPRVALUE_NUMERIC

    Case 64 'MATERNITY EXPECTED RETURN DATE
      fOK = True
      piResultType = giEXPRVALUE_DATE

    Case 65 'IS POST SUBORDINATE OF
      'JPD 20040127 Hierarchy performance modifications
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 66 'IS POST SUBORDINATE OF USER
      'JPD 20040127 Hierarchy performance modifications
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 67 'IS PERSONNEL SUBORDINATE OF
      'JPD 20040127 Hierarchy performance modifications
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 68 'IS PERSONNEL SUBORDINATE OF USER
      'JPD 20040127 Hierarchy performance modifications
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 69 'HAS POST SUBORDINATE
      'JPD 20040127 Hierarchy performance modifications
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 70 'HAS POST SUBORDINATE USER
      'JPD 20040127 Hierarchy performance modifications
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 71 'HAS PERSONNEL SUBORDINATE
      'JPD 20040127 Hierarchy performance modifications
      fOK = False
      'Select Case IdentifyingColumnDataType
      '  Case sqlNumeric, sqlInteger
      '    fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      '  Case Else
      '    fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      'End Select
      piResultType = giEXPRVALUE_LOGIC

    Case 72 'HAS PERSONNEL SUBORDINATE USER
      'JPD 20040127 Hierarchy performance modifications
      fOK = True
      'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
      '      (piParam2Type = giEXPRVALUE_DATE)
      piResultType = giEXPRVALUE_LOGIC

    Case 73 'BRADFORD FACTOR
      fOK = (piParam1Type = giEXPRVALUE_DATE) And _
        (piParam2Type = giEXPRVALUE_DATE) And _
        (piParam3Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_NUMERIC

    Case 74 ' Does Record Exist (Workflow only)
      fOK = True
      piResultType = giEXPRVALUE_LOGIC

    Case 75 ' Is Valid NI Number
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_LOGIC

    Case 76 ' Is Valid Payroll Character Set
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
        (piParam2Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_LOGIC

    Case 77 ' Replace Characters in a String
      fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
      piResultType = giEXPRVALUE_CHARACTER

    Case Else ' Unknown function
      fOK = False
  End Select

TidyUpAndExit:
  If Not fOK Then
    piResultType = giEXPR_UNKNOWNTYPE
  End If

  ValidateFunctionParameters = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function PrintComponent(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objParameter As CExprComponent
  
  glngBottom = CalculateBottomOfPage
  
  fOK = True
  
  ' Position the printing.
  Printer.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
  Printer.CurrentY = Printer.CurrentY + giPRINT_YSPACE
  Printer.Print "Function : " & ComponentDescription
  
  CheckEndOfPage2 glngBottom
  
  ' Print the function's parameter expressions.
  For Each objParameter In mcolParameters
    objParameter.Component.PrintComponent (piLevel + 1)
  Next objParameter

TidyUpAndExit:
  PrintComponent = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function




Public Property Get ReturnType() As ExpressionValueTypes
  ' Return the function's return type.
  ReturnType = miReturnType

End Property


Public Property Get ComponentType() As ExpressionComponentTypes
  ' Return the 'function' component type.
  ComponentType = giCOMPONENT_FUNCTION
  
End Property


Public Property Get ComponentDescription() As String
  ' Return the Function description.
  ComponentDescription = msFunctionName
  
End Property

Public Function CopyComponent()
  ' Copies the selected component.
  ' When editting a component we actually copy the component first
  ' and edit the copy. If the changes are confirmed then the copy
  ' replaces the original. If the changes are cancelled then the
  ' copy is discarded.
  Dim objFunctionCopy As New CExprFunction
  
  ' Copy the component's basic properties.
  ' ie. the function id, not its parameters, etc.
  With objFunctionCopy
    .FunctionID = mlngFunctionID
  
    .WorkflowRecordSelectionType = miWorkflowRecordSelection
    .WorkflowElement = msWorkflowElementIdentifier
    .WorkflowItem = msWorkflowItemIdentifier
    .WorkflowRecordTableID = mlngWorkflowRecordTableID
  End With

   ' JDM - 06/02/01 - Now copies it's children so that cut'n paste works
    ' Copy all the child components
    Dim iCount As Integer
    Dim objParameter As CExprComponent
    For iCount = 1 To mcolParameters.Count
'        Set objParameter = New clsExprComponent
        objFunctionCopy.Parameters.Add (mcolParameters(iCount).CopyComponent)
    Next iCount
    Set objParameter = Nothing
  
  Set CopyComponent = objFunctionCopy
   
  ' Disassociate object variables.
  Set objFunctionCopy = Nothing
   
End Function

Public Function CloneComponent(pavCloneRegister As Variant)
  ' Clones the selected component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim lngIndex As Long
  Dim objFunctionClone As New CExprFunction
  Dim objParameterClone As CExprComponent
  Dim objParameter As CExprComponent
  Dim lngNewRecordTableID As Long
  
  fOK = True
  lngNewRecordTableID = 0
  
  'JPD 20070530 Fault 12292
  If (mobjBaseComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWCALCULATION _
      And mobjBaseComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWSTATICFILTER _
      And mobjBaseComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWRUNTIMEFILTER) _
    And (mlngFunctionID = 30 _
      Or mlngFunctionID = 46 _
      Or mlngFunctionID = 47) Then
  
    fOK = False
  End If
  
  ' Clone the component's basic properties.
  ' ie. the function id, not its parameters, etc.
  
  If fOK Then
    ' Get the new table ID from the clone register.
    For lngIndex = 1 To UBound(pavCloneRegister, 2)
      If (pavCloneRegister(1, lngIndex)) = "TABLE" And _
        (pavCloneRegister(2, lngIndex) = mlngWorkflowRecordTableID) And _
        (lngNewRecordTableID = 0) Then
        lngNewRecordTableID = pavCloneRegister(3, lngIndex)
      End If
    
      If ((mlngWorkflowRecordTableID = 0) Or (lngNewRecordTableID > 0)) Then
        Exit For
      End If
    Next lngIndex
    If ((mlngWorkflowRecordTableID > 0) And (lngNewRecordTableID = 0)) Then
      lngNewRecordTableID = mlngWorkflowRecordTableID
    End If
  
    objFunctionClone.FunctionID = mlngFunctionID
    
    objFunctionClone.WorkflowRecordSelectionType = miWorkflowRecordSelection
    objFunctionClone.WorkflowElement = msWorkflowElementIdentifier
    objFunctionClone.WorkflowItem = msWorkflowItemIdentifier
    objFunctionClone.WorkflowRecordTableID = lngNewRecordTableID
        
    ' Clone the parameter expressions.
    lngIndex = 1
    For Each objParameter In mcolParameters
      Set objParameterClone = objParameter.CloneComponent(pavCloneRegister)
    
      fOK = Not (objParameterClone Is Nothing)
      
      If fOK Then
  '      mcolParameters.Remove lngIndex
  '      If mcolParameters.Count >= lngIndex Then
  'objFunctionClone.Parameters.Remove lngIndex
        If objFunctionClone.Parameters.Count >= lngIndex Then
          objFunctionClone.Parameters.Add objParameterClone, , lngIndex
        Else
          objFunctionClone.Parameters.Add objParameterClone
        End If
        
        lngIndex = lngIndex + 1
      Else
        Exit For
      End If
    Next objParameter
    Set objParameter = Nothing
  End If
  
TidyUpAndExit:
  If fOK Then
    Set CloneComponent = objFunctionClone
  Else
    Set CloneComponent = Nothing
  End If
  ' Disassociate object variables.
  Set objFunctionClone = Nothing
  Set objParameter = Nothing
  Set objParameterClone = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get WorkflowRecordTableID() As Long
  WorkflowRecordTableID = mlngWorkflowRecordTableID
  
End Property

Public Property Let WorkflowRecordTableID(ByVal plngNewValue As Long)
  mlngWorkflowRecordTableID = plngNewValue
  
End Property
Public Property Get WorkflowRecordSelectionType() As WorkflowRecordSelectorTypes
  WorkflowRecordSelectionType = miWorkflowRecordSelection
  
End Property

Public Property Get WorkflowElement() As String
  ' Return the WorkflowElement property.
  WorkflowElement = msWorkflowElementIdentifier
  
End Property

Public Property Get WorkflowItem() As String
  ' Return the Workflow Item property.
  WorkflowItem = msWorkflowItemIdentifier
  
End Property


Public Property Let WorkflowElement(ByVal psNewValue As String)
  ' Set the Workflow Element property.
  msWorkflowElementIdentifier = psNewValue

End Property
Public Property Let WorkflowItem(ByVal psNewValue As String)
  ' Set the Workflow Item property.
  msWorkflowItemIdentifier = psNewValue

End Property
Public Property Let WorkflowRecordSelectionType(ByVal piNewValue As WorkflowRecordSelectorTypes)
  miWorkflowRecordSelection = piNewValue
  
End Property
Public Property Get BadComponent() As CExprComponent
  ' Return the component last caused the function to fail its validity check.
  Set BadComponent = mobjBadComponent

End Property

Public Property Get BaseComponent() As CExprComponent
  ' Return the component's base component object.
  Set BaseComponent = mobjBaseComponent
  
End Property

Public Property Set BaseComponent(ByVal pobjNewValue As CExprComponent)
  ' Set the component's base component object property.
  Set mobjBaseComponent = pobjNewValue
  
End Property


Public Sub ClearParameters()
  ' Clear the function's parameter collection.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Remove all components from the collection.
  Do While mcolParameters.Count > 0
    mcolParameters.Remove 1
  Loop
  Set mcolParameters = Nothing
  
  ' Re-instantiate the collection.
  Set mcolParameters = New Collection
  Exit Sub

ErrorTrap:
  fOK = False
  
End Sub


Public Function ComponentContainsColumn(plngColumnID As Long) As Boolean
  ' Return TRUE if the component makes any reference to the given column.
  ' Used to check for recursion.
  Dim objParameter As CExprComponent
      
  ComponentContainsColumn = False
  
  For Each objParameter In mcolParameters
    ComponentContainsColumn = objParameter.Component.ExpressionContainsColumn(plngColumnID)

    If ComponentContainsColumn Then
      Exit For
    End If
  Next objParameter
  Set objParameter = Nothing

End Function



Public Sub ColumnsUsedInThisComponent(ByRef pvColumns As Variant)
  ' Return an array of the columns used this expression.
  Dim objParameter As CExprComponent
      
  For Each objParameter In mcolParameters
    objParameter.Component.ColumnsUsedInThisExpression pvColumns
  Next objParameter
  Set objParameter = Nothing

End Sub





Public Sub ExpressionsUsedInThisComponent(ByRef palngExpressions As Variant)
  ' Return an array of the expressions used this expression.
  Dim objParameter As CExprComponent

  For Each objParameter In mcolParameters
    objParameter.Component.ExpressionsUsedInThisExpression palngExpressions
  Next objParameter
  Set objParameter = Nothing

End Sub






Public Function WriteComponent() As Boolean
  ' Write the component definition to the component recordset.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim objParameter As CExprComponent
  Dim objSubExpression As CExpression
  
  fOK = True
  
  With recCompEdit
    ' Add a record for new expression
    .AddNew
    !ComponentID = mobjBaseComponent.ComponentID
    !ExprID = mobjBaseComponent.ParentExpression.ExpressionID
    !Type = giCOMPONENT_FUNCTION
    !FunctionID = mlngFunctionID
  
    !WorkflowRecord = miWorkflowRecordSelection
    !WorkflowElement = msWorkflowElementIdentifier
    !WorkflowItem = msWorkflowItemIdentifier
    !WorkflowRecordTableID = mlngWorkflowRecordTableID
  
    .Update
  End With
  
  ' Write the function parameter expressions.
  For Each objParameter In mcolParameters
    Set objSubExpression = objParameter.Component
    objSubExpression.ParentComponentID = mobjBaseComponent.ComponentID
    objSubExpression.ExpressionID = 0
    objSubExpression.IsNew = True
    objSubExpression.IsChanged = False
    objSubExpression.UtilityID = mobjBaseComponent.ParentExpression.UtilityID
    
    fOK = objSubExpression.WriteExpression
        
    Set objSubExpression = Nothing
  
    If Not fOK Then
      Exit For
    End If
  Next objParameter
  Set objParameter = Nothing

TidyUpAndExit:
  WriteComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Function ConstructComponent() As Boolean
  ' Read the Function component record.
  ' NB. This function must be called after seeking the correct record in
  ' the recCompEdit recordset.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngComponentID As Long
  Dim lngExprID As Long
  Dim iIndex As Integer
  Dim objNewParameter As CExprComponent
  
  fOK = True
  
  ' Read the function component record
  mlngFunctionID = recCompEdit!FunctionID
  lngComponentID = recCompEdit!ComponentID
       
  miWorkflowRecordSelection = IIf(IsNull(recCompEdit!WorkflowRecord), giWFRECSEL_UNKNOWN, recCompEdit!WorkflowRecord)
  msWorkflowElementIdentifier = IIf(IsNull(recCompEdit!WorkflowElement), "", recCompEdit!WorkflowElement)
  msWorkflowItemIdentifier = IIf(IsNull(recCompEdit!WorkflowItem), "", recCompEdit!WorkflowItem)
  mlngWorkflowRecordTableID = IIf(IsNull(recCompEdit!WorkflowRecordTableID), 0, recCompEdit!WorkflowRecordTableID)
       
  ' Read the function definition.
  ReadFunction
            
TidyUpAndExit:
  ConstructComponent = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function


Public Property Get FunctionID() As Long
  ' Return the function ID property.
  FunctionID = mlngFunctionID
  
End Property


Public Property Let FunctionID(ByVal plngNewValue As Long)
  ' Set the function ID property.
  mlngFunctionID = plngNewValue
  
  ' Read the function definition from the database.
  ReadFunction
  
End Property







Public Property Get Parameters() As Collection
  ' Return the parameter collection.
  Set Parameters = mcolParameters
  
End Property

Public Property Set Parameters(ByVal pcolCollection As Collection)
  ' Set the parameter collection.
  Set mcolParameters = pcolCollection

End Property







Private Sub ReadFunction()
  ' Read the function definition from the database.
  On Error GoTo ErrorTrap
  
  ' Read the function name, return type, etc.
  ReadFunctionDetails
  
  ' Create the array of parameter components.
  ReadParameterDefinition

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Function ReadFunctionDetails() As Boolean
  ' Read the function details (not paramter info) from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objFunctionDef As clsFunctionDef
  
  fOK = True
  msFunctionName = "<unknown>"
  miReturnType = giEXPRVALUE_UNDEFINED
  msSPName = vbNullString
  msUDFName = vbNullString
  
  ' Clear the parameter collection.
  ClearParameters
  
  ' Get the function list.
  gobjFunctionDefs.Initialise
  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim$(Str$(mlngFunctionID)))
  
    msFunctionName = objFunctionDef.Name
    miReturnType = objFunctionDef.ReturnType
    msSPName = objFunctionDef.SPName
    msUDFName = objFunctionDef.UDFName
    
    If Not mobjBaseComponent Is Nothing Then
      If mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION Or _
        mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER Or _
        mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER Then
        
        Select Case mlngFunctionID
          Case 17
            msFunctionName = "Login Name"
            
          Case 66
            msFunctionName = "Is Post That Reports To Identified Person"
          
          Case 68
            msFunctionName = "Is Personnel That Reports To Identified Person"

          Case 70
            msFunctionName = "Is Post That Identified Person Reports To"

          Case 72
            msFunctionName = "Is Personnel That Identified Person Reports To"
        
        End Select
      End If
    End If
    
  Else
    fOK = False
  End If
  
TidyUpAndExit:
  'Set objFunctionDef = Nothing  'MH20060727  New code not implemented
  ReadFunctionDetails = fOK
  Exit Function
  
ErrorTrap:
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  fOK = False
  Resume TidyUpAndExit

End Function


Private Sub ReadParameterDefinition()
  ' Read the function's paramter definition from the database,
  ' and create an array of components to represent the parameters.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim sSQL As String
  Dim rsCustomisedParameters As DAO.Recordset
  Dim objNewParameter As CExprComponent
  Dim objFunctionDefParameter As clsFunctionDefParameter
  Dim objFunctionDef As clsFunctionDef
  Dim iMinIndex As Integer
  Dim iMaxIndex As Integer
  
  fOK = True
  
  ' Clear the parameter collection.
  ClearParameters

  ' Read the function parameter list from the collection.
  gobjFunctionDefs.Initialise
  
  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim$(Str$(mlngFunctionID)))
  
    iMinIndex = 0
    iMaxIndex = 0
    For Each objFunctionDefParameter In objFunctionDef.Parameters
      If objFunctionDefParameter.Index < iMinIndex Then iMinIndex = objFunctionDefParameter.Index
      If objFunctionDefParameter.Index > iMaxIndex Then iMaxIndex = objFunctionDefParameter.Index
    Next objFunctionDefParameter
    Set objFunctionDefParameter = Nothing
    If iMinIndex > iMaxIndex Then iMaxIndex = iMinIndex
    
    For iIndex = iMinIndex To iMaxIndex
      For Each objFunctionDefParameter In objFunctionDef.Parameters
        If objFunctionDefParameter.Index = iIndex Then
          mcolParameters.Add New CExprComponent
          With mcolParameters.Item(mcolParameters.Count)
            .ComponentType = giCOMPONENT_EXPRESSION
            .Component.Name = objFunctionDefParameter.Name
            .Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
            .Component.ReturnType = objFunctionDefParameter.ParameterType
            .Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
            
            If (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
              Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
              Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
            
              .Component.UtilityID = mobjBaseComponent.ParentExpression.UtilityID
              .Component.UtilityBaseTable = mobjBaseComponent.ParentExpression.UtilityBaseTable
              .Component.WorkflowInitiationType = mobjBaseComponent.ParentExpression.WorkflowInitiationType
              .Component.PrecedingWorkflowElements = mobjBaseComponent.ParentExpression.PrecedingWorkflowElements
              .Component.AllWorkflowElements = mobjBaseComponent.ParentExpression.AllWorkflowElements
            End If
          End With
        End If
      Next objFunctionDefParameter
      Set objFunctionDefParameter = Nothing
    Next iIndex
  End If
  
  ' Get the customised function parameter definitions if they exist.
  If (mobjBaseComponent.ComponentID > 0) Then

    iIndex = 1
    sSQL = "SELECT *" & _
      " FROM tmpExpressions" & _
      " WHERE parentComponentID = " & Trim$(Str$(mobjBaseComponent.ComponentID)) & _
      " AND deleted = False" & _
      " ORDER BY exprID"
    Set rsCustomisedParameters = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsCustomisedParameters
      Do While (Not .EOF) And fOK

        ' Instantiate a new component object.
        Set objNewParameter = New CExprComponent
  
        ' Construct the hierarchy of objects that define the parameter.
        objNewParameter.ComponentType = giCOMPONENT_EXPRESSION
        objNewParameter.Component.ExpressionID = !ExprID
        objNewParameter.Component.BaseTableID = mobjBaseComponent.ParentExpression.BaseTableID
  
            
        If (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
          Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
          Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
        
          objNewParameter.Component.UtilityID = mobjBaseComponent.ParentExpression.UtilityID
          objNewParameter.Component.UtilityBaseTable = mobjBaseComponent.ParentExpression.UtilityBaseTable
          objNewParameter.Component.WorkflowInitiationType = mobjBaseComponent.ParentExpression.WorkflowInitiationType
          objNewParameter.Component.PrecedingWorkflowElements = mobjBaseComponent.ParentExpression.PrecedingWorkflowElements
          objNewParameter.Component.AllWorkflowElements = mobjBaseComponent.ParentExpression.AllWorkflowElements
        End If
  
        fOK = objNewParameter.Component.ConstructExpression
  
        If fOK Then
          objNewParameter.Component.ExpressionType = mobjBaseComponent.ParentExpression.ExpressionType
  
          ' Reset the new expression's return type.
          objNewParameter.Component.ReturnType = mcolParameters.Item(iIndex).Component.ReturnType
  
          ' Insert the new expression into the function's parameter array.
          mcolParameters.Remove iIndex
          If mcolParameters.Count >= iIndex Then
            mcolParameters.Add objNewParameter, , iIndex
          Else
            mcolParameters.Add objNewParameter
          End If
        End If

        iIndex = iIndex + 1
        Set objNewParameter = Nothing

        .MoveNext
      Loop
      
      .Close
    End With
  End If
    
TidyUpAndExit:
  'MH20060727  New code not implemented
  'Set objNewParameter = Nothing
  'Set objFunctionDefParameter = Nothing
  'Set objFunctionDef = Nothing

  Set rsCustomisedParameters = Nothing
  Exit Sub
  
ErrorTrap:
  ' Clear the parameter collection.
  ClearParameters
  fOK = False
  Resume TidyUpAndExit
      
End Sub



Public Function ValidateFunction() As ExprValidationCodes
  ' Validate the function. Return a code describing the validity.
  On Error GoTo BasicErrorTrap
  
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim lngLoop As Long
  Dim lngLoop3 As Long
  Dim iValidationCode As ExprValidationCodes
  Dim iFunctionReturnType As ExpressionValueTypes
  Dim aiDummyValues(6) As Integer
  Dim objSubExpression As CExpression
  Dim objParameter As CExprComponent
  Dim objFunctionDefParameter As clsFunctionDefParameter
  Dim objFunctionDef As clsFunctionDef
  Dim aWFPrecedingElements() As VB.Control
  Dim aWFAllElements() As VB.Control
  Dim fValidIdentification As Boolean
  Dim wfTemp As VB.Control
  Dim alngValidTables() As Long
  Dim asItems() As String
  Dim fFound As Boolean
  Dim lngRequiredTableID As Long
  
  iLoop = 0
  gobjFunctionDefs.Initialise
  fValidIdentification = True
  
  ' Initialise the validation code.
  iValidationCode = giEXPRVALIDATION_NOERRORS
  Set mobjBadComponent = Nothing

  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim$(Str$(mlngFunctionID)))

    ' Validate the function parameter expressions.
    For Each objParameter In mcolParameters
      iLoop = iLoop + 1
    
      Set objSubExpression = objParameter.Component
      With objSubExpression
        ' Validate the parameter expression.
        ' NB. Reset the sub-expression's return type to that defined by the parameter definition
        ' as it may be changeable. The evaluated return type will be determined when the
        ' sub-expression is validated.
        For Each objFunctionDefParameter In objFunctionDef.Parameters
          If objFunctionDefParameter.Index = iLoop Then
            objSubExpression.ReturnType = objFunctionDefParameter.ParameterType
            Exit For
          End If
        Next objFunctionDefParameter
        Set objFunctionDefParameter = Nothing

        iValidationCode = .ValidateExpression(False)
        
        If iValidationCode <> giEXPRVALIDATION_NOERRORS Then
          ' Interpret the parameter sub-expression validation code to reflect
          ' the fact that a function parameter was invalid.
          Select Case iValidationCode
            Case giEXPRVALIDATION_NOCOMPONENTS
              iValidationCode = giEXPRVALIDATION_PARAMETERNOCOMPONENTS
            Case giEXPRVALIDATION_SYNTAXERROR
              iValidationCode = giEXPRVALIDATION_PARAMETERSYNTAXERROR
            Case giEXPRVALIDATION_EXPRTYPEMISMATCH
              iValidationCode = giEXPRVALIDATION_PARAMETERTYPEMISMATCH
          End Select
          
          If .BadComponent Is Nothing Then
            Set mobjBadComponent = objParameter
          Else
            Set mobjBadComponent = .BadComponent
          End If
          Exit For
        End If
      
        ' Write the given return type into the array.
        aiDummyValues(iLoop) = .ReturnType
      End With
    Next objParameter
    Set objParameter = Nothing
  End If
  
    
  If (iValidationCode = giEXPRVALIDATION_NOERRORS) _
    And ((mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
      Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
      Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)) Then

    lngRequiredTableID = FunctionBaseTableID(mlngFunctionID)

    If (lngRequiredTableID <> 0) _
      And (miWorkflowRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) Then
    
      ' Check that the identifiying element/item is still valid.
      aWFPrecedingElements = mobjBaseComponent.ParentExpression.PrecedingWorkflowElements
      aWFAllElements = mobjBaseComponent.ParentExpression.AllWorkflowElements
  
      If UBound(aWFAllElements) >= 0 Then
        fValidIdentification = False
  
        For lngLoop = 2 To UBound(aWFPrecedingElements)
          If UCase(Trim(aWFPrecedingElements(lngLoop).Identifier)) = UCase(Trim(msWorkflowElementIdentifier)) Then
            Set wfTemp = aWFPrecedingElements(lngLoop)
  
            ReDim alngValidTables(0)
  
            If wfTemp.ElementType = elem_StoredData Then
              TableAscendants wfTemp.DataTableID, alngValidTables
  
              If wfTemp.DataAction = DATAACTION_DELETE Then
                ' Cannot do anything with a Deleted record, but can use its ascendants.
                ' Remove the table itself from the array of valid tables.
                alngValidTables(1) = 0
              End If

            ElseIf wfTemp.ElementType = elem_WebForm Then
              asItems = wfTemp.Items
  
              For iLoop2 = 1 To UBound(asItems, 2)
                If UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(msWorkflowItemIdentifier)) _
                  And asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
  
                  TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                  Exit For
                End If
              Next iLoop2
            End If
  
            fFound = False
            For lngLoop3 = 1 To UBound(alngValidTables)
              If (alngValidTables(lngLoop3) = lngRequiredTableID) _
                Or (lngRequiredTableID < 0) Then
                
                fFound = True
                Exit For
              End If
            Next lngLoop3
  
            If fFound Then
              fValidIdentification = True
            End If
  
            Exit For
          End If
        Next lngLoop
      End If
    End If
    
    If Not fValidIdentification Then
      iValidationCode = giEXPRVALIDATION_INVALIDRECORDIDENTIFICATION
    End If
  End If

  If iValidationCode = giEXPRVALIDATION_NOERRORS Then
    If Not ValidateFunctionParameters(mlngFunctionID, iFunctionReturnType, _
      aiDummyValues(1), aiDummyValues(2), aiDummyValues(3), _
      aiDummyValues(4), aiDummyValues(5), aiDummyValues(6)) Then
      
      iValidationCode = giEXPRVALIDATION_PARAMETERTYPEMISMATCH
    Else
      miReturnType = iFunctionReturnType
    End If
  End If

TidyUpAndExit:
  ' Disassociate object variables.
  Set objSubExpression = Nothing
  ValidateFunction = iValidationCode
  Exit Function
  
BasicErrorTrap:
  iValidationCode = giEXPRVALIDATION_UNKNOWNERROR
  Resume TidyUpAndExit

End Function

Private Function FunctionBaseTableID(plngFunctionID As Long) As Long
  ' Workflow record selection is only required for the following functions:
  ' Return:
  '   The ID of the required table
  '   0 if no specific table AND no record identification is required
  '   -1 if no specific table is required BUT record identification IS required
  Dim lngRequiredTableID As Long
  Dim lngPersonnelTableID As Long
  Dim lngDependantsTableID As Long
  Dim lngMaternityTableID As Long
  Dim lngHierarchyTableID As Long
  
  lngRequiredTableID = 0
  
  lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  lngDependantsTableID = GetModuleSetting(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE, 0)
  lngMaternityTableID = GetModuleSetting(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYTABLE, 0)
  lngHierarchyTableID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE, 0)
  
  Select Case plngFunctionID
    Case 17
      ' 17 - Current User
      If (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
        (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
        (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
    
        lngRequiredTableID = lngPersonnelTableID
      End If
    Case 30
      ' 30 - Absence Duration
      lngRequiredTableID = lngPersonnelTableID
    Case 46
      ' 46 - Working Days Between Two Dates
      lngRequiredTableID = lngPersonnelTableID
    Case 47
      ' 47 - Absence Between Two Dates
      lngRequiredTableID = lngPersonnelTableID
    Case 62
      ' 62 - Parental Leave Entitlement
      lngRequiredTableID = lngDependantsTableID
    Case 63
      ' 63 - Parental Leave Taken
      lngRequiredTableID = lngDependantsTableID
    Case 64
      ' 64 - Maternity Return Date
      lngRequiredTableID = lngMaternityTableID
    Case 66
      ' 66 - Post that Reports to Current User
      lngRequiredTableID = lngHierarchyTableID
    Case 68
      ' 68 - Personnel that Reports to Current User
      lngRequiredTableID = lngPersonnelTableID
    Case 70
      ' 70 - Post that Current User Reports to
      lngRequiredTableID = lngHierarchyTableID
    Case 72
      ' 72 - Personnel that Current User Reports to
      lngRequiredTableID = lngPersonnelTableID
    Case 73
      ' 73 - Bradford Factor
      lngRequiredTableID = lngPersonnelTableID
    Case 74
      ' 74 - Does Record Exist
      lngRequiredTableID = -1
  End Select

  FunctionBaseTableID = lngRequiredTableID
  
End Function



Public Function EvaluateComponent(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated result of the function.
'  On Error GoTo BasicErrorTrap
'
'  Dim iLoop As Integer
'  Dim sSQL As String
'  Dim sResultSQL As String
'  Dim sParameterSQL As String
'  Dim vResult As Variant
'  Dim avValues(6) As Variant
'  Dim objSubExpression As CExpression
'  Dim objEvaluator As New ASRHRProEvaluator.Evaluator
'  Dim rsResult As rdoResultset
'
'  If Not mfStandardFunction Then
'    vResult = NonStandardFunction_Evaluate(pLngRecordID, pfDoPrompts)
'  Else
'
'    ' Evaluate the function parameter expressions.
'    ' Write each parameter's evaluated value into the array.
'    sParameterSQL = ""
'    For iLoop = 1 To giParameterCount
'      Set objSubExpression = gaParameterArray(iLoop).Component
'      avValues(iLoop) = objSubExpression.EvaluateExpression(pLngRecordID, pfDoPrompts)
'      Set objSubExpression = Nothing
'
'      Select Case VarType(avValues(iLoop))
'        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
'          sParameterSQL = sParameterSQL & IIf(Len(sParameterSQL) > 0, ", ", "") & avValues(iLoop)
'        Case vbDate
'          sParameterSQL = sParameterSQL & IIf(Len(sParameterSQL) > 0, ", ", "") & "'" & Format(avValues(iLoop), "MM/dd/yyyy") & "'"
'        Case vbString
'          sParameterSQL = sParameterSQL & IIf(Len(sParameterSQL) > 0, ", ", "") & "'" & avValues(iLoop) & "'"
'        Case vbBoolean
'          sParameterSQL = sParameterSQL & IIf(Len(sParameterSQL) > 0, ", ", "") & IIf(avValues(iLoop), 1, 0)
'      End Select
'    Next iLoop
'
'    If Len(msSPName) > 0 Then
'      Select Case miReturnType
'        Case giEXPRVALUE_NUMERIC
'          sResultSQL = "DECLARE @result float" & vbnewline
'        Case giEXPRVALUE_DATE
'          sResultSQL = "DECLARE @result datetime" & vbnewline
'        Case giEXPRVALUE_CHARACTER
'          sResultSQL = "DECLARE @result varchar(MAX)" & vbnewline
'        Case giEXPRVALUE_LOGIC
'          sResultSQL = "DECLARE @result bit" & vbnewline
'      End Select
'
'      sSQL = sResultSQL & "EXEC " & msSPName & " @result OUTPUT" & IIf(Len(sParameterSQL) > 0, ", ", "") & sParameterSQL
'
'      Set rsResult = rdoCon.OpenResultset(sSQL, _
'        rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'      With rsResult
'        If Not (.EOF And .BOF) Then
'          Select Case miReturnType
'            Case giEXPRVALUE_NUMERIC
'              vResult = IIf(IsNull(!result), 0, !result)
'            Case giEXPRVALUE_DATE
'              vResult = IIf(IsNull(!result), Null, !result)
'            Case giEXPRVALUE_CHARACTER
'              vResult = IIf(IsNull(!result), "", !result)
'            Case giEXPRVALUE_LOGIC
'              vResult = IIf(IsNull(!result), False, !result)
'          End Select
'        Else
'          Select Case miReturnType
'            Case giEXPRVALUE_NUMERIC
'              vResult = 0
'            Case giEXPRVALUE_DATE
'              vResult = Null
'            Case giEXPRVALUE_CHARACTER
'              vResult = ""
'            Case giEXPRVALUE_LOGIC
'              vResult = False
'          End Select
'        End If
'        .Close
'      End With
'      Set rsResult = Nothing
'    Else
'      On Error GoTo EvaluatorErrorTrap
'
'      ' Send the parameter values to the function evaluator.
'      vResult = objEvaluator.Evaluate(1, mlngFunctionID, _
'        avValues(1), avValues(2), avValues(3), _
'        avValues(4), avValues(5), avValues(6))
'
'      On Error GoTo BasicErrorTrap
'    End If
'
'    ' Determine the return type of the function with the given parameters.
'    Select Case VarType(vResult)
'      Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
'        miReturnType = giEXPRVALUE_NUMERIC
'      Case vbDate
'        miReturnType = giEXPRVALUE_DATE
'      Case vbString
'        miReturnType = giEXPRVALUE_CHARACTER
'      Case vbBoolean
'        miReturnType = giEXPRVALUE_LOGIC
'    End Select
'
'  End If
'
'TidyUpAndExit:
'  ' Disassociate object variables.
'  Set objEvaluator = Nothing
'  Set objSubExpression = Nothing
'  EvaluateComponent = vResult
'  Exit Function
'
'EvaluatorErrorTrap:
'  ' The operator evaluation has failed.
'  Resume TidyUpAndExit
'
'BasicErrorTrap:
'  Resume TidyUpAndExit
'
End Function

Private Function NonStandardFunction_Evaluate(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Evaluate the non-standard functions.
'  Select Case mlngFunctionID
'    Case 42
'      NonStandardFunction_Evaluate = NonStandardFunction_Evaluate_42(pLngRecordID, pfDoPrompts)
'  End Select
  
End Function


Public Function StoredProcedureCode(Optional strBaseTable As String) As String
  ' Create the expression stored procedure.
  On Error GoTo ErrorTrap
  
  Dim fUnknownParameterType As Boolean
  Dim iLoop As Integer
  Dim sProcCode As SystemMgr.cStringBuilder
  Dim sIfNullCode As String
  Dim sVarName As String
  Dim sUnknownParameterSuffix As String
  Dim objSubExpression As CExpression
  Dim objParameter As CExprComponent
  Dim sBaseTableName As String
  Dim sPersonnelTableName As String
  Dim iPersTableID As Integer
  Dim iWFPersTableID As Integer
  Dim objFunctionDefParameter As clsFunctionDefParameter
  Dim objFunctionDef As clsFunctionDef
  Dim fOK As Boolean
  Dim sSrchColumnName As String
  Dim sRtnColumnName As String
  Dim sSrchTableName As String
  Dim lngSrchTableID As Long
  Dim lngRtnTableID As Long
  Dim fWorkflowExpression As Boolean
  Dim rsTemp As DAO.Recordset
  Dim sSQL As String
  Dim lngElementID As Long
  Dim iElementType As ElementType
  Dim lngRequiredTableID As Long
  Dim sRequiredTableName As String
  
  fUnknownParameterType = False
  sUnknownParameterSuffix = ""
  gobjFunctionDefs.Initialise
  
  fWorkflowExpression = (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
    Or (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER)
  
  Set sProcCode = New SystemMgr.cStringBuilder
  sVarName = "@component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  
  fUnknownParameterType = False
  If gobjFunctionDefs.IsValidID(mlngFunctionID) Then
    Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim$(Str$(mlngFunctionID)))
  
    For Each objFunctionDefParameter In objFunctionDef.Parameters
      If objFunctionDefParameter.ParameterType = giEXPRVALUE_UNDEFINED Then
        fUnknownParameterType = True
        Exit For
      End If
    Next objFunctionDefParameter
    Set objFunctionDefParameter = Nothing
  End If
  
  ' Determine the function return type if it is not explicitly defined.
  If miReturnType = giEXPRVALUE_UNDEFINED Then
    ValidateFunction
  End If

  Select Case miReturnType
    Case giEXPRVALUE_CHARACTER
      sProcCode.TheString = "    DECLARE " & sVarName & " varchar(MAX);" & vbNewLine
      sIfNullCode = "SET " & sVarName & " = '';"
    Case giEXPRVALUE_NUMERIC
      sProcCode.TheString = "    DECLARE " & sVarName & " float;" & vbNewLine
      sIfNullCode = "SET " & sVarName & " = 0;"
    Case giEXPRVALUE_LOGIC
      sProcCode.TheString = "    DECLARE " & sVarName & " bit;" & vbNewLine
      sIfNullCode = "SET " & sVarName & " = 0;"
    Case giEXPRVALUE_DATE
      sProcCode.TheString = "    DECLARE " & sVarName & " datetime;" & vbNewLine
      sIfNullCode = "SET " & sVarName & " = null;"
  End Select
        
  ' Construct the stored procedure code for the function's parameters.
  If mlngFunctionID = 4 Then ' If...Then...Else.. function handled differently.
    Set objSubExpression = mcolParameters.Item(1).Component
    sProcCode.Append "    DECLARE " & sVarName & "_1" & " bit;" & vbNewLine & vbNewLine & _
      objSubExpression.StoredProcedureCode(sVarName & "_1", strBaseTable) & vbNewLine & vbNewLine & _
      "    IF " & sVarName & "_1 = 1" & vbNewLine & _
      "    BEGIN" & vbNewLine
    Set objSubExpression = Nothing
      
    Set objSubExpression = mcolParameters.Item(2).Component
    objSubExpression.ValidateExpression (False)
      
    Select Case objSubExpression.ReturnType
      Case giEXPRVALUE_CHARACTER
        sProcCode.Append "    DECLARE " & sVarName & "_2" & " varchar(MAX);" & vbNewLine & vbNewLine
      Case giEXPRVALUE_NUMERIC
        sProcCode.Append "    DECLARE " & sVarName & "_2" & " float;" & vbNewLine & vbNewLine
      Case giEXPRVALUE_LOGIC
        sProcCode.Append "    DECLARE " & sVarName & "_2" & " bit;" & vbNewLine & vbNewLine
      Case giEXPRVALUE_DATE
        sProcCode.Append "    DECLARE " & sVarName & "_2" & " datetime;" & vbNewLine & vbNewLine
      Case giEXPRVALUE_BYREF_CHARACTER, giEXPRVALUE_BYREF_NUMERIC, giEXPRVALUE_BYREF_LOGIC, giEXPRVALUE_BYREF_DATE
        sProcCode.Append "    DECLARE " & sVarName & "_2" & " varchar(MAX);" & vbNewLine & vbNewLine
    End Select
      
    sProcCode.Append objSubExpression.StoredProcedureCode(sVarName & "_2", strBaseTable) & vbNewLine & vbNewLine
    Set objSubExpression = Nothing
    
    sProcCode.Append _
      "        SET " & sVarName & " = " & sVarName & "_2;" & vbNewLine & _
      "    END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "    BEGIN" & vbNewLine
      
    Set objSubExpression = mcolParameters.Item(3).Component
    objSubExpression.ValidateExpression (False)
      
    Select Case objSubExpression.ReturnType
      Case giEXPRVALUE_CHARACTER
        sProcCode.Append "    DECLARE " & sVarName & "_3" & " varchar(MAX);" & vbNewLine & vbNewLine
      Case giEXPRVALUE_NUMERIC
        sProcCode.Append "    DECLARE " & sVarName & "_3" & " float;" & vbNewLine & vbNewLine
      Case giEXPRVALUE_LOGIC
        sProcCode.Append "    DECLARE " & sVarName & "_3" & " bit;" & vbNewLine & vbNewLine
      Case giEXPRVALUE_DATE
        sProcCode.Append "    DECLARE " & sVarName & "_3" & " datetime;" & vbNewLine & vbNewLine
      Case giEXPRVALUE_BYREF_CHARACTER, giEXPRVALUE_BYREF_NUMERIC, giEXPRVALUE_BYREF_LOGIC, giEXPRVALUE_BYREF_DATE
        sProcCode.Append "    DECLARE " & sVarName & "_3" & " varchar(MAX);" & vbNewLine & vbNewLine
    End Select
      
    sProcCode.Append objSubExpression.StoredProcedureCode(sVarName & "_3", strBaseTable) & vbNewLine & vbNewLine
    Set objSubExpression = Nothing
    
    sProcCode.Append _
      "        SET " & sVarName & " = " & sVarName & "_3;" & vbNewLine & _
      "    END"
  Else
    ' Evaluate the function parameter expressions.
    iLoop = 0
    For Each objParameter In mcolParameters
      iLoop = iLoop + 1
      
      Set objSubExpression = objParameter.Component
    
      If objSubExpression.ReturnType = giEXPRVALUE_UNDEFINED Then
        objSubExpression.ValidateExpression (False)
      End If
  
      sUnknownParameterSuffix = sUnknownParameterSuffix & "_" & Trim$(Str$(objSubExpression.ReturnType))
      
      Select Case objSubExpression.ReturnType
        Case giEXPRVALUE_CHARACTER
          sProcCode.Append "    DECLARE " & sVarName & "_" & Trim$(Str$(iLoop)) & " varchar(MAX);" & vbNewLine & vbNewLine
        Case giEXPRVALUE_NUMERIC
          sProcCode.Append "    DECLARE " & sVarName & "_" & Trim$(Str$(iLoop)) & " float;" & vbNewLine & vbNewLine
        Case giEXPRVALUE_LOGIC
          sProcCode.Append "    DECLARE " & sVarName & "_" & Trim$(Str$(iLoop)) & " bit;" & vbNewLine & vbNewLine
        Case giEXPRVALUE_DATE
          sProcCode.Append "    DECLARE " & sVarName & "_" & Trim$(Str$(iLoop)) & " datetime;" & vbNewLine & vbNewLine
        Case giEXPRVALUE_BYREF_CHARACTER, giEXPRVALUE_BYREF_NUMERIC, giEXPRVALUE_BYREF_LOGIC, giEXPRVALUE_BYREF_DATE
          sProcCode.Append "    DECLARE " & sVarName & "_" & Trim$(Str$(iLoop)) & " varchar(MAX);" & vbNewLine & vbNewLine
        Case giEXPRVALUE_BYREF_UNDEFINED
          sProcCode.Append "    DECLARE " & sVarName & "_" & Trim$(Str$(iLoop)) & " int;" & vbNewLine & vbNewLine
        
      End Select
      
      sProcCode.Append objSubExpression.StoredProcedureCode(sVarName & "_" & Trim$(Str$(iLoop)), strBaseTable) & vbNewLine & vbNewLine
    
      Set objSubExpression = Nothing
    Next objParameter
    Set objParameter = Nothing
   
    ' Get the personnel table name
    iPersTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
    iWFPersTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
    
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", iPersTableID
    If Not recTabEdit.NoMatch Then
      sPersonnelTableName = recTabEdit!TableName
    End If
    
    ' RH - 18/12/00
    ' if the stored procedure name is sp_ASRFn_AbsenceDuration or
    ' sp_ASRFn_AbsenceBetweenTwoDates or sp_asrfn_workingdaysbetweentwodates then add
    ' this little bit that passes the personnel.ID to the function
    
    'JPD 20040106 Use IDs not spName
    'If LCase(msSPName) = "sp_asrfn_absenceduration" Or _
      LCase(msSPName) = "sp_asrfn_absencebetweentwodates" Or _
      LCase(msSPName) = "sp_asrfn_workingdaysbetweentwodates" Then
    'JPD 20040127 Hierarchy performance modifications
    'If (mlngFunctionID = 30) Or _
      (mlngFunctionID = 47) Or _
      (mlngFunctionID = 46) Or _
      (mlngFunctionID = 67) Or _
      (mlngFunctionID = 68) Or _
      (mlngFunctionID = 71) Or _
      (mlngFunctionID = 72) Then
    If (mlngFunctionID = 30) Or _
      (mlngFunctionID = 47) Or _
      (mlngFunctionID = 46) Or _
      (mlngFunctionID = 73) Then
  
      If fWorkflowExpression Then
        Select Case miWorkflowRecordSelection
          Case giWFRECSEL_INITIATOR ' 0 - Initiator's personnel table record
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
              "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
              "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
              "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
              "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
              "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
              "       @iBaseTableID = " & CStr(iWFPersTableID) & vbNewLine & _
              "    FROM ASRSysWorkflowInstances" & vbNewLine & _
              "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID;" & vbNewLine & vbNewLine
          
          Case giWFRECSEL_IDENTIFIEDRECORD ' 1 = Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
          
            sSQL = "SELECT TOP 1 tmpWorkflowElements.ID," & _
              "   tmpWorkflowElements.type" & _
              " FROM tmpWorkflowElements" & _
              " WHERE tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & _
              "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'"
          
            Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            With rsTemp
              If Not (.BOF And .EOF) Then
                lngElementID = !ID
                iElementType = !Type
              End If
              
              .Close
            End With
            Set rsTemp = Nothing
      
            If iElementType = elem_WebForm Then
              ' WebForm
              sProcCode.Append _
                "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
                "    SELECT @sValue = CASE" & vbNewLine & _
                "            WHEN IV.elementID = @piTempElement THEN ISNULL(IV.tempValue, '0')" & _
                "            ELSE ISNULL(IV.value, '0')" & vbNewLine & _
                "        END," & vbNewLine & _
                "        @iWorkflowID = Es.workflowID," & vbNewLine & _
                "        @iParent1TableID = CASE" & vbNewLine & _
                "            WHEN IV.elementID = @piTempElement THEN IV.tempParent1TableID" & _
                "            ELSE IV.parent1TableID" & vbNewLine & _
                "        END," & vbNewLine & _
                "        @iParent1RecordID = CASE" & vbNewLine & _
                "            WHEN IV.elementID = @piTempElement THEN IV.tempParent1RecordID" & _
                "            ELSE IV.parent1RecordID" & vbNewLine & _
                "        END," & vbNewLine & _
                "        @iParent2TableID = CASE" & vbNewLine & _
                "            WHEN IV.elementID = @piTempElement THEN IV.tempParent2TableID" & _
                "            ELSE IV.parent2TableID" & vbNewLine & _
                "        END," & vbNewLine & _
                "        @iParent2RecordID = CASE" & vbNewLine & _
                "            WHEN IV.elementID = @piTempElement THEN IV.tempParent2RecordID" & _
                "            ELSE IV.parent2RecordID" & vbNewLine & _
                "        END," & vbNewLine & _
                "        @iBaseTableID = EI.tableID" & vbNewLine
                
              sProcCode.Append _
                "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
                "    INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier" & vbNewLine & _
                "    INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID" & vbNewLine & _
                "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
                "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = '" & Replace(UCase(Trim(msWorkflowItemIdentifier)), "'", "''") & "'" & vbNewLine & _
                "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'" & vbNewLine & _
                "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
                "        AND Es.ID = IV.elementID" & vbNewLine & vbNewLine
            Else
              ' StoredData
              sProcCode.Append _
                "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
                "    SELECT @sValue = ISNULL(IV.value, '0')," & vbNewLine & _
                "       @iWorkflowID = Es.workflowID," & vbNewLine & _
                "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
                "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
                "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
                "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
                "       @iBaseTableID = isnull(Es.dataTableID, 0)" & vbNewLine & _
                "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
                "    INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID" & vbNewLine & _
                "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = UPPER(LTRIM(RTRIM(Es.identifier)))" & vbNewLine & _
                "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
                "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'" & vbNewLine & _
                "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & vbNewLine
            End If
            
            sProcCode.Append _
              "    SET @iBaseRecordID = " & vbNewLine & _
              "        CASE" & vbNewLine & _
              "            WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)" & vbNewLine & _
              "            ELSE 0" & vbNewLine & _
              "        END" & vbNewLine & vbNewLine

          Case giWFRECSEL_TRIGGEREDRECORD ' 4 = Triggered Base table record
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
              "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
              "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
              "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
              "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
              "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
              "       @iBaseTableID = isnull(ASRSysWorkflows.baseTable, 0)" & vbNewLine & _
              "    FROM ASRSysWorkflowInstances" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
              "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
        
          Case giWFRECSEL_UNIDENTIFIED ' 3 = Unidentified table record
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SET @iBaseRecordID = @id" & vbNewLine & _
              "    SET @iWorkflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "    SET @iParent1TableID = 0" & vbNewLine & _
              "    SET @iParent1RecordID = 0" & vbNewLine & _
              "    SET @iParent2TableID = 0" & vbNewLine & _
              "    SET @iParent2RecordID = 0" & vbNewLine & _
              "    SET @iBaseTableID = " & CStr(mobjBaseComponent.ParentExpression.BaseTableID) & vbNewLine & vbNewLine
        End Select
      
        If (miWorkflowRecordSelection = giWFRECSEL_INITIATOR) _
          Or (miWorkflowRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
          Or (miWorkflowRecordSelection = giWFRECSEL_TRIGGEREDRECORD) _
          Or (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
      
          sProcCode.Append vbNewLine & vbNewLine & _
            "    EXEC spASRWorkflowAscendantRecordID" & vbNewLine & _
            "        @iBaseTableID," & vbNewLine & _
            "        @iBaseRecordID," & vbNewLine & _
            "        @iParent1TableID," & vbNewLine & _
            "        @iParent1RecordID," & vbNewLine & _
            "        @iParent2TableID," & vbNewLine & _
            "        @iParent2RecordID," & vbNewLine & _
            "        " & CStr(iWFPersTableID) & "," & vbNewLine & _
            "        @PersonnelID  OUTPUT" & vbNewLine & vbNewLine
        End If
      Else
      
        ' Get the base table name for the expression
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", Me.BaseComponent.ParentExpression.BaseTableID
        If Not recTabEdit.NoMatch Then
          sBaseTableName = recTabEdit!TableName
        End If
      
        If LCase$(sBaseTableName) = LCase$(sPersonnelTableName) Then
          sProcCode.Append vbNewLine & vbNewLine & _
            "    /* Extra Section to pass in hidden ID parameter */" & vbNewLine & _
            "    SELECT @PersonnelID = @id" & vbNewLine & vbNewLine
        Else
          sProcCode.Append vbNewLine & vbNewLine & _
            "    /* Extra Section to pass in hidden ID parameter */" & vbNewLine & _
            "    SELECT @PersonnelID = ID_" & iPersTableID & " FROM " & sBaseTableName & " WHERE ID = @id" & vbNewLine & vbNewLine
        End If
      End If
    End If

    If fWorkflowExpression _
      And ((mlngFunctionID = 17) _
        Or (mlngFunctionID = 62) _
        Or (mlngFunctionID = 63) _
        Or (mlngFunctionID = 64) _
        Or (mlngFunctionID = 74)) Then
      ' 17 = Current User/Login Name
      ' 62 = Parental Leave Entitlement
      ' 63 = Parental Leave Taken
      ' 64 = Maternity return date
      ' 74 = Does Record Exist
      
      Select Case mlngFunctionID
        Case 17
          ' 17 = Current User/Login Name
          lngRequiredTableID = iWFPersTableID
        Case 62, 63
          ' 62 = Parental Leave Entitlement
          ' 63 = Parental Leave Taken
          recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE
          If recModuleSetup.NoMatch Then
            lngRequiredTableID = 0
          Else
            lngRequiredTableID = IIf(IsNull(recModuleSetup!parametervalue) Or Len(recModuleSetup!parametervalue) = 0, 0, recModuleSetup!parametervalue)
          End If
        Case 64
          ' Maternity return date
          recModuleSetup.Seek "=", gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYTABLE
          If recModuleSetup.NoMatch Then
            lngRequiredTableID = 0
          Else
            lngRequiredTableID = IIf(IsNull(recModuleSetup!parametervalue) Or Len(recModuleSetup!parametervalue) = 0, 0, recModuleSetup!parametervalue)
          End If
        Case 74
          ' 74 = Does Record Exist
          lngRequiredTableID = mlngWorkflowRecordTableID
      End Select
      
      Select Case miWorkflowRecordSelection
        Case giWFRECSEL_INITIATOR ' 0 - Initiator's personnel table record
          sProcCode.Append _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
            "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
            "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
            "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
            "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
            "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
            "       @iBaseTableID = " & CStr(iWFPersTableID) & vbNewLine & _
            "    FROM ASRSysWorkflowInstances" & vbNewLine & _
            "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
    
        Case giWFRECSEL_IDENTIFIEDRECORD ' 1 = Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
          sSQL = "SELECT TOP 1 tmpWorkflowElements.ID," & _
            "   tmpWorkflowElements.type" & _
            " FROM tmpWorkflowElements" & _
            " WHERE tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & _
            "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'"
    
          Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          With rsTemp
            If Not (.BOF And .EOF) Then
              lngElementID = !ID
              iElementType = !Type
            End If
    
            .Close
          End With
          Set rsTemp = Nothing
    
          If iElementType = elem_WebForm Then
            ' WebForm
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @sValue = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN ISNULL(IV.tempValue, '0')" & _
              "            ELSE ISNULL(IV.value, '0')" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iWorkflowID = Es.workflowID," & vbNewLine & _
              "        @iParent1TableID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent1TableID" & _
              "            ELSE IV.parent1TableID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iParent1RecordID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent1RecordID" & _
              "            ELSE IV.parent1RecordID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iParent2TableID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent2TableID" & _
              "            ELSE IV.parent2TableID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iParent2RecordID = CASE" & vbNewLine & _
              "            WHEN IV.elementID = @piTempElement THEN IV.tempParent2RecordID" & _
              "            ELSE IV.parent2RecordID" & vbNewLine & _
              "        END," & vbNewLine & _
              "        @iBaseTableID = EI.tableID" & vbNewLine
              
            sProcCode.Append _
              "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID" & vbNewLine & _
              "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = '" & Replace(UCase(Trim(msWorkflowItemIdentifier)), "'", "''") & "'" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'" & vbNewLine & _
              "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "        AND Es.ID = IV.elementID" & vbNewLine & vbNewLine
          Else
            ' StoredData
            sProcCode.Append _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @sValue = ISNULL(IV.value, '0')," & vbNewLine & _
              "       @iWorkflowID = Es.workflowID," & vbNewLine & _
              "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
              "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
              "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
              "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
              "       @iBaseTableID = isnull(Es.dataTableID, 0)" & vbNewLine & _
              "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = UPPER(LTRIM(RTRIM(Es.identifier)))" & vbNewLine & _
              "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'" & vbNewLine & _
              "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & vbNewLine
          End If
          
          sProcCode.Append _
            "    SET @iBaseRecordID = " & vbNewLine & _
            "        CASE" & vbNewLine & _
            "            WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)" & vbNewLine & _
            "            ELSE 0" & vbNewLine & _
              "        END" & vbNewLine & vbNewLine
    
        Case giWFRECSEL_TRIGGEREDRECORD ' 4 = Triggered Base table record
          sProcCode.Append _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
            "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
            "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
            "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
            "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
            "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
            "       @iBaseTableID = isnull(ASRSysWorkflows.baseTable, 0)" & vbNewLine & _
            "    FROM ASRSysWorkflowInstances" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
      
      End Select
    
      If (miWorkflowRecordSelection = giWFRECSEL_INITIATOR) _
        Or (miWorkflowRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
        Or (miWorkflowRecordSelection = giWFRECSEL_TRIGGEREDRECORD) Then
    
        sProcCode.Append vbNewLine & vbNewLine & _
          "    EXEC spASRWorkflowAscendantRecordID" & vbNewLine & _
          "        @iBaseTableID," & vbNewLine & _
          "        @iBaseRecordID," & vbNewLine & _
          "        @iParent1TableID," & vbNewLine & _
          "        @iParent1RecordID," & vbNewLine & _
          "        @iParent2TableID," & vbNewLine & _
          "        @iParent2RecordID," & vbNewLine & _
          "        " & CStr(lngRequiredTableID) & "," & vbNewLine & _
          "        @ID  OUTPUT" & vbNewLine & vbNewLine
      End If
    End If

    'JPD 20040127 Hierarchy performance modifications
    'If (mlngFunctionID = 65) Or _
    '  (mlngFunctionID = 66) Or _
    '  (mlngFunctionID = 69) Or _
    '  (mlngFunctionID = 70) Then
    '
    '  ' Get the base table name for the expression
    '  recTabEdit.Index = "idxTableID"
    '  recTabEdit.Seek "=", Me.BaseComponent.ParentExpression.BaseTableID
    '  If Not recTabEdit.NoMatch Then
    '    sBaseTableName = recTabEdit!TableName
    '  End If
    '
    '  ' Get the Hierarchy table name
    '  recModuleSetup.Index = "idxModuleParameter"
    '  recModuleSetup.Seek "=", gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE
    '  If Not recModuleSetup.NoMatch Then
    '    iTempID = recModuleSetup!parametervalue
    '  End If
    '  recTabEdit.Index = "idxTableID"
    '  recTabEdit.Seek "=", iTempID
    '  If Not recTabEdit.NoMatch Then
    '    sHierarchyTableName = recTabEdit!TableName
    '  End If
    '
    '  If LCase(sBaseTableName) = LCase(sHierarchyTableName) Then
    '      sProcCode = sProcCode & vbnewline & vbnewline & _
    '                  "    /* Extra Section to pass in hidden ID parameter */" & vbnewline & _
    '                  "    SELECT @HierarchyID = @id" & vbnewline & vbnewline
    '  Else
    '    sProcCode = sProcCode & vbnewline & vbnewline & _
    '                "    /* Extra Section to pass in hidden ID parameter */" & vbnewline & _
    '                "    SELECT @HierarchyID = ID_" & iTempID & " FROM " & sBaseTableName & " WHERE ID = @id" & vbnewline & vbnewline
    '  End If
    'End If
    
    ' JPD20021128 Fault 4809
    If mlngFunctionID = 42 Then
      ' Get Field From Database Record handled differently
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", val(mcolParameters(1).Component.Components(1).Component.ColumnID)
        fOK = Not .NoMatch
    
        If fOK Then
          sSrchColumnName = !ColumnName
          lngSrchTableID = !TableID
          
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", lngSrchTableID
                   
          fOK = Not recTabEdit.NoMatch
          
          If fOK Then
            sSrchTableName = recTabEdit!TableName
    
            .Seek "=", val(mcolParameters(3).Component.Components(1).Component.ColumnID)
            fOK = Not .NoMatch
      
            If fOK Then
              sRtnColumnName = !ColumnName
              lngRtnTableID = !TableID
      
              fOK = ((LenB(sSrchColumnName) <> 0) And (LenB(sRtnColumnName) <> 0)) And _
                (lngSrchTableID = lngRtnTableID)
            End If
          End If
        End If
    
        If fOK Then
          sProcCode.Append _
            "    SELECT TOP 1 " & sVarName & " = " & sSrchTableName & "." & sRtnColumnName & vbNewLine & _
            "    FROM " & sSrchTableName & vbNewLine & _
            "    WHERE (" & sSrchTableName & "." & sSrchColumnName & " = " & sVarName & "_2)" & vbNewLine
        End If
      End With
    ElseIf mlngFunctionID = 74 Then
      ' Does Record Exist (Workflow only)
      sRequiredTableName = GetTableName(lngRequiredTableID)
      sProcCode.Append _
        "    SELECT " & sVarName & " = " & vbNewLine & _
        "        CASE" & vbNewLine & _
        "            WHEN (SELECT COUNT(*) FROM " & sRequiredTableName & " WHERE ID = @ID) > 0 THEN 1" & vbNewLine & _
        "            ELSE 0" & vbNewLine & _
        "        END" & vbNewLine
    Else
      If LenB(msSPName) <> 0 Or LenB(msUDFName) <> 0 Then
        
        If LenB(msSPName) <> 0 Then
        
          ' Construct the stored procedure code for executing the function.
          If (mlngFunctionID = 17) And _
            fWorkflowExpression Then
          
            sProcCode.Append _
              "    IF EXISTS (SELECT Name" & _
              "        FROM sysobjects" & _
              "        WHERE id = object_id('spASRSysGetLoginName')" & _
              "        AND sysstat & 0xf = 4)" & _
              "        EXEC [spASRSysGetLoginName] @ID, " & sVarName & " OUTPUT"
          
          Else
            sProcCode.Append "    EXEC " & msSPName & _
              IIf(fUnknownParameterType, sUnknownParameterSuffix, "") & _
              " " & sVarName & " OUTPUT"
          
            iLoop = 0
            For Each objParameter In mcolParameters
              iLoop = iLoop + 1
              sProcCode.Append ", " & vbNewLine & "        " & sVarName & "_" & Trim$(Str$(iLoop))
            Next objParameter
            Set objParameter = Nothing
          
            If (mlngFunctionID = 30) Or _
              (mlngFunctionID = 47) Or _
              (mlngFunctionID = 46) Or _
              (mlngFunctionID = 73) Then
      
              sProcCode.Append ", " & vbNewLine & "        @PersonnelID" & vbNewLine
            End If
          End If

        ElseIf LenB(msUDFName) <> 0 Then
          sProcCode.Append "    SET " & IIf(fUnknownParameterType, sUnknownParameterSuffix, "") & sVarName & " = dbo." & msUDFName & "("
        
          For iLoop = 1 To mcolParameters.Count
            sProcCode.Append " " & sVarName & "_" & Trim$(Str$(iLoop)) & IIf(iLoop = mcolParameters.Count, "", ",")
          Next iLoop
        
          If mlngFunctionID = 73 Then
            sProcCode.Append ", " & "@PersonnelID)" & vbNewLine
          Else
            sProcCode.Append " @id)" & vbNewLine
          End If
        
        End If


        'JPD 20040127 Hierarchy performance modifications
        'If (mlngFunctionID = 65) Or _
        '  (mlngFunctionID = 66) Or _
        '  (mlngFunctionID = 69) Or _
        '  (mlngFunctionID = 70) Then
        '
        '  sProcCode = sProcCode & ", " & vbnewline & "        @HierarchyID" & vbnewline
        'End If

        ''''JDM - 16/01/02 - Added extra parameters for Audit functions
        ''If LCase(msSPName) = "sp_asrfn_auditfieldlastchangedate" Or _
          LCase(msSPName) = "sp_asrfn_auditfieldchangedbetweendates" Then
        If LCase$(msSPName) = "sp_asrfn_auditfieldlastchangedate" Or _
          LCase$(msSPName) = "sp_asrfn_auditfieldchangedbetweendates" Or _
          LCase$(msSPName) = "spasrsysfnparentalleaveentitlement" Or _
          LCase$(msSPName) = "spasrsysfnparentalleavetaken" Or _
          LCase$(msSPName) = "spasrsysfnmaternityexpectedreturn" Then
          sProcCode.Append ", " & vbNewLine & "        @ID" & vbNewLine
        End If

      Else
        ' Parentheses.
        sProcCode.Append "    SET " & sVarName & " = " & sVarName & "_1"
      End If
    End If
  End If
  
  sProcCode.Append vbNewLine & _
    "    IF " & sVarName & " IS NULL " & sIfNullCode

TidyUpAndExit:
  ' Disassociate object variables.
  Set objSubExpression = Nothing
  StoredProcedureCode = sProcCode.ToString
  Set sProcCode = Nothing
  Exit Function
  
ErrorTrap:
  sProcCode.TheString = ""
  Resume TidyUpAndExit

End Function



Private Sub Class_Initialize()
  ' Create a new collection to hold the function's parameters.
  Set mcolParameters = New Collection

End Sub


Public Function RuntimeCode(psRuntimeCode As String, _
  palngSourceTables As Variant, _
  pfValidating As Boolean, _
  psPrefix As String) As Boolean
  ' Return the SQL code for the component.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngSrchTableID As Long
  Dim lngRtnTableID As Long
  Dim sCode As String
  Dim sParamCode1 As String
  Dim sParamCode2 As String
  Dim sParamCode3 As String
  Dim sParamCode4 As String
  Dim sSrchColumnName As String
  Dim sRtnColumnName As String
  Dim sSrchTableName As String
  Dim strRemainString As String

  'Currency Conversion Values
  Dim sCConvTable As String
  Dim sCConvExRateCol As String
  Dim sCConvCurrDescCol As String
  Dim sCConvDecCol As String
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  
  Dim sHierarchyTableName As String
  Dim sPersonnelTableName As String
  Dim iPersonnelTableID As Integer
  Dim fWorkflowExpression As Boolean
  Dim iPersTableID As Integer
  Dim lngElementID As Long
  Dim iElementType As ElementType
  Dim lngRequiredTableID As Long
  Dim sRequiredTableName As String
  Dim lngHierarchyTableID As Long
  Dim lngPersonnelTableID As Long
  Dim lngDependantsTableID As Long
  Dim lngMaternityTableID As Long
  Dim sVarName As String
  Dim rsTemp As DAO.Recordset
  Dim strTempTableName As String
  Dim strTempTableID As String
  
  sSQL = vbNullString
  
  fOK = True
  sCode = vbNullString
  sVarName = "@component" & Trim$(Str$(mobjBaseComponent.ComponentID))
  
  fWorkflowExpression = (mobjBaseComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)
  
  If fWorkflowExpression Then
    lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
    lngDependantsTableID = GetModuleSetting(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE, 0)
    lngMaternityTableID = GetModuleSetting(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYTABLE, 0)
    lngHierarchyTableID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE, 0)
    
    lngRequiredTableID = -1
    
    Select Case mlngFunctionID
      Case 17
        ' 17 - Current User
        lngRequiredTableID = lngPersonnelTableID
      Case 30
        ' 30 - Absence Duration
        lngRequiredTableID = lngPersonnelTableID
      Case 46
        ' 46 - Working Days Between Two Dates
        lngRequiredTableID = lngPersonnelTableID
      Case 47
        ' 47 - Absence Between Two Dates
        lngRequiredTableID = lngPersonnelTableID
      Case 62
        ' 62 - Parental Leave Entitlement
        lngRequiredTableID = lngDependantsTableID
      Case 63
        ' 63 - Parental Leave Taken
        lngRequiredTableID = lngDependantsTableID
      Case 64
        ' 64 - Maternity Return Date
        lngRequiredTableID = lngMaternityTableID
      Case 66
        ' 66 - Post that Reports to Current User
        lngRequiredTableID = lngPersonnelTableID
      Case 68
        ' 68 - Personnel that Reports to Current User
        lngRequiredTableID = lngPersonnelTableID
      Case 70
        ' 70 - Post that Current User Reports to
        lngRequiredTableID = lngPersonnelTableID
      Case 72
        ' 72 - Personnel that Current User Reports to
        lngRequiredTableID = lngPersonnelTableID
      Case 73
        ' 73 - Bradford Factor
        lngRequiredTableID = lngPersonnelTableID
      Case 74
        ' 74 - Does Record Exist (Workflow only)
        lngRequiredTableID = mlngWorkflowRecordTableID
    End Select
    
    If lngRequiredTableID > 0 Then
      sRequiredTableName = GetTableName(lngRequiredTableID)
      
      psPrefix = psPrefix & _
        "    DECLARE " & sVarName & " integer" & vbNewLine
      
      Select Case miWorkflowRecordSelection
        Case giWFRECSEL_INITIATOR ' 0 - Initiator's personnel table record
          psPrefix = psPrefix & _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
            "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
            "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
            "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
            "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
            "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
            "       @iBaseTableID = " & CStr(lngPersonnelTableID) & vbNewLine & _
            "    FROM ASRSysWorkflowInstances" & vbNewLine & _
            "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
        
        Case giWFRECSEL_IDENTIFIEDRECORD ' 1 = Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
        
          sSQL = "SELECT TOP 1 tmpWorkflowElements.ID," & _
            "   tmpWorkflowElements.type" & _
            " FROM tmpWorkflowElements" & _
            " WHERE tmpWorkflowElements.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & _
            "   AND UCASE(LTRIM(RTRIM(tmpWorkflowElements.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'"
        
          Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          With rsTemp
            If Not (.BOF And .EOF) Then
              lngElementID = !ID
              iElementType = !Type
            End If
            
            .Close
          End With
          Set rsTemp = Nothing
    
          If iElementType = elem_WebForm Then
            ' WebForm
            psPrefix = psPrefix & _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @sValue = ISNULL(IV.value, '0')," & vbNewLine & _
              "       @iWorkflowID = Es.workflowID," & vbNewLine & _
              "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
              "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
              "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
              "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
              "       @iBaseTableID = EI.tableID" & vbNewLine & _
              "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID" & vbNewLine & _
              "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = '" & Replace(UCase(Trim(msWorkflowItemIdentifier)), "'", "''") & "'" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'" & vbNewLine & _
              "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "        AND Es.ID = IV.elementID" & vbNewLine & vbNewLine
          Else
            ' StoredData
            psPrefix = psPrefix & _
              "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
              "    SELECT @sValue = ISNULL(IV.value, '0')," & vbNewLine & _
              "       @iWorkflowID = Es.workflowID," & vbNewLine & _
              "       @iParent1TableID = IV.parent1TableID," & vbNewLine & _
              "       @iParent1RecordID = IV.parent1RecordID," & vbNewLine & _
              "       @iParent2TableID = IV.parent2TableID," & vbNewLine & _
              "       @iParent2RecordID = IV.parent2RecordID," & vbNewLine & _
              "       @iBaseTableID = isnull(Es.dataTableID, 0)" & vbNewLine & _
              "    FROM ASRSysWorkflowInstanceValues IV" & vbNewLine & _
              "    INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID" & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(IV.identifier))) = UPPER(LTRIM(RTRIM(Es.identifier)))" & vbNewLine & _
              "        AND Es.workflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
              "        AND UPPER(LTRIM(RTRIM(Es.identifier))) = '" & Replace(UCase(Trim(msWorkflowElementIdentifier)), "'", "''") & "'" & vbNewLine & _
              "    WHERE IV.instanceID = @piInstanceID" & vbNewLine & vbNewLine
          End If
          
          psPrefix = psPrefix & _
            "    SET @iBaseRecordID = " & vbNewLine & _
            "        CASE" & vbNewLine & _
            "            WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)" & vbNewLine & _
            "            ELSE 0" & vbNewLine & _
              "        END" & vbNewLine & vbNewLine
        
        Case giWFRECSEL_TRIGGEREDRECORD ' 4 = Triggered Base table record
          psPrefix = psPrefix & _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SELECT @iBaseRecordID = ASRSysWorkflowInstances.initiatorID," & vbNewLine & _
            "       @iWorkflowID = ASRSysWorkflowInstances.workflowID," & vbNewLine & _
            "       @iParent1TableID = ASRSysWorkflowInstances.parent1TableID," & vbNewLine & _
            "       @iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID," & vbNewLine & _
            "       @iParent2TableID = ASRSysWorkflowInstances.parent2TableID," & vbNewLine & _
            "       @iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID," & vbNewLine & _
            "       @iBaseTableID = isnull(ASRSysWorkflows.baseTable, 0)" & vbNewLine & _
            "    FROM ASRSysWorkflowInstances" & vbNewLine & _
            "    INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
            "    WHERE ASRSysWorkflowInstances.ID = @piInstanceID" & vbNewLine & vbNewLine
      
        Case giWFRECSEL_UNIDENTIFIED ' 3 = Unidentified table record
          psPrefix = psPrefix & _
            "    -- Extra Section to read Workflow record identifiers" & vbNewLine & _
            "    SET @iWorkflowID = " & CStr(mobjBaseComponent.ParentExpression.UtilityID) & vbNewLine & _
            "    SET @iParent1TableID = 0" & vbNewLine & _
            "    SET @iParent1RecordID = 0" & vbNewLine & _
            "    SET @iParent2TableID = 0" & vbNewLine & _
            "    SET @iParent2RecordID = 0" & vbNewLine & _
            "    SET @iBaseTableID = " & CStr(mobjBaseComponent.ParentExpression.BaseTableID) & vbNewLine & vbNewLine
      End Select
    
      If (miWorkflowRecordSelection = giWFRECSEL_INITIATOR) _
        Or (miWorkflowRecordSelection = giWFRECSEL_IDENTIFIEDRECORD) _
        Or (miWorkflowRecordSelection = giWFRECSEL_TRIGGEREDRECORD) Then
    
        psPrefix = psPrefix & vbNewLine & vbNewLine & _
          "    SET " & sVarName & " = dbo.udf_ASRWorkflowAscendantRecordID(" & vbNewLine & _
          "        @iBaseTableID," & vbNewLine & _
          "        @iBaseRecordID," & vbNewLine & _
          "        @iParent1TableID," & vbNewLine & _
          "        @iParent1RecordID," & vbNewLine & _
          "        @iParent2TableID," & vbNewLine & _
          "        @iParent2RecordID," & vbNewLine & _
          "        " & CStr(lngRequiredTableID) & ")" & vbNewLine & vbNewLine
      End If
    End If
  End If
  
  ' Get the first parameter's runtime code if required.
  If mcolParameters.Count >= 1 Then
    fOK = mcolParameters.Item(1).Component.RuntimeCode(sParamCode1, palngSourceTables, pfValidating, psPrefix)
  End If
  
  ' Get the second parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 2) Then
    fOK = mcolParameters.Item(2).Component.RuntimeCode(sParamCode2, palngSourceTables, pfValidating, psPrefix)
  End If
  
  ' Get the third parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 3) Then
    fOK = mcolParameters.Item(3).Component.RuntimeCode(sParamCode3, palngSourceTables, pfValidating, psPrefix)
  End If
    
  ' Get the fourth parameter's runtime code if required.
  If fOK And (mcolParameters.Count >= 4) Then
    fOK = mcolParameters.Item(4).Component.RuntimeCode(sParamCode4, palngSourceTables, pfValidating, psPrefix)
  End If
    
  If fOK Then
    Select Case mlngFunctionID
      Case 1  ' System date
        If fWorkflowExpression Then
          sCode = "@dtDate"
        Else
          sCode = "getdate()"
        End If
          
      Case 2  ' Convert to uppercase
        sCode = "upper(" & sParamCode1 & ")"
        
      Case 3  ' Convert numeric to string
        sCode = "IsNull(ltrim(str(" & sParamCode1 & ", 255, " & sParamCode2 & ")),'')"
  
      Case 4  ' If... Then... Else...
        sCode = "CASE WHEN (" & sParamCode1 & " = 1) THEN " & sParamCode2 & " ELSE " & sParamCode3 & " END"
    
      Case 5  ' Remove leading and trailing spaces
        sCode = "ltrim(rtrim(" & sParamCode1 & "))"
        
      Case 6  ' Extract characters from the left
        sCode = "left(" & sParamCode1 & ", " & sParamCode2 & ")"
        
      Case 7  ' Length of character field
        sCode = "len(" & sParamCode1 & ")"
        
      Case 8  ' Convert to lowercase
        sCode = "lower(" & sParamCode1 & ")"
  
      Case 9  ' Maximum
        sCode = "CASE WHEN (" & sParamCode1 & " > " & sParamCode2 & ") THEN " & sParamCode1 & " ELSE " & sParamCode2 & " END"
      
      Case 10  ' Minimum
        sCode = "CASE WHEN (" & sParamCode1 & " < " & sParamCode2 & ") THEN " & sParamCode1 & " ELSE " & sParamCode2 & " END"
      
      Case 11  ' Search for character string.
        sCode = "charindex(" & sParamCode2 & ", " & sParamCode1 & ")"
        
      Case 12  ' Capitalize Initials
        If fWorkflowExpression Then
          sCode = "(dbo.udf_ASRFn_CapitalizeInitials(" & sParamCode1 & "))"
        Else
          fOK = False
        End If
        
      Case 13  ' Extract characters from the right
        sCode = "right(" & sParamCode1 & ", " & sParamCode2 & ")"
        
      Case 14  ' Extract part of a character string
        sCode = "substring(" & sParamCode1 & ", " & sParamCode2 & ", " & sParamCode3 & ")"
  
      Case 15  ' System Time
        If fWorkflowExpression Then
          sCode = "convert(varchar(50), @dtDate, 8)"
        Else
          sCode = "convert(varchar(50), getdate(), 8)"
        End If
        
      Case 16  ' Is field empty
        sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"

        ' Validate the sub-expression. This is done, not to  validate the expression,
        ' but rather to determine the return type of the expression.
        mcolParameters(1).Component.ValidateExpression False
      
        Select Case mcolParameters(1).ReturnType
          Case giEXPRVALUE_CHARACTER
            sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
          Case giEXPRVALUE_NUMERIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
          Case giEXPRVALUE_LOGIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
        End Select
        
        sCode = sCode & " THEN 1 ELSE 0 END)"
  
      Case 17  ' Current user
        If fWorkflowExpression Then
          If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
            sCode = "(dbo.udf_ASRGetLoginName(" & sRequiredTableName & ".ID))"
          Else
            sCode = "(dbo.udf_ASRGetLoginName(" & sVarName & "))"
          End If
        Else
          sCode = "dbo.udfsys_getsystemuser()"
        End If
        
      Case 18  ' Whole Years Until Current Date
        If fWorkflowExpression Then
          sCode = "datediff(year," & sParamCode1 & ", @dtDate)" & _
            " - case" & _
            "       when datepart(month," & sParamCode1 & ") > datepart(month, @dtDate) then 1" & _
            "       when (datepart(month," & sParamCode1 & ") = datepart(month, @dtDate)) " & _
            "           and (datepart(day," & sParamCode1 & ") > datepart(day, @dtDate)) then 1" & _
            "       else 0" & _
            "   end"
        Else
          sCode = "datediff(year," & sParamCode1 & ", getdate())" & _
            " - case" & _
            "       when datepart(month," & sParamCode1 & ") > datepart(month, getdate()) then 1" & _
            "       when (datepart(month," & sParamCode1 & ") = datepart(month, getdate())) " & _
            "           and (datepart(day," & sParamCode1 & ") > datepart(day, getdate())) then 1" & _
            "       else 0" & _
            "   end"
        End If
        
      Case 19  ' Remaining Months Since Whole Years
        If fWorkflowExpression Then
          sCode = "datepart(month, @dtDate)" & _
            " - datepart(month, " & sParamCode1 & ")" & _
            " - case" & _
            "       when datepart(day," & sParamCode1 & ") > datepart(day, @dtDate) then 1" & _
            "       else 0" & _
            "   end" & _
            " + case" & _
            "       when (datepart(month, @dtDate)" & _
            "           - datepart(month," & sParamCode1 & ")" & _
            "           - case" & _
            "               when datepart(day," & sParamCode1 & ") > datepart(day, @dtDate) then 1" & _
            "               else 0" & _
            "             end) < 0 then 12" & _
            "       else 0" & _
            "   end"
        Else
          sCode = "datepart(month, getdate())" & _
            " - datepart(month, " & sParamCode1 & ")" & _
            " - case" & _
            "       when datepart(day," & sParamCode1 & ") > datepart(day, getdate()) then 1" & _
            "       else 0" & _
            "   end" & _
            " + case" & _
            "       when (datepart(month, getdate())" & _
            "           - datepart(month," & sParamCode1 & ")" & _
            "           - case" & _
            "               when datepart(day," & sParamCode1 & ") > datepart(day, getdate()) then 1" & _
            "               else 0" & _
            "             end) < 0 then 12" & _
            "       else 0" & _
            "   end"
        End If
      
      'JPD 20081103 Fault 13411
      Case 20  ' Initials from Forenames
        If fWorkflowExpression Then
          sCode = "(dbo.udf_ASRFn_InitialsFromForenames(" & sParamCode1 & "))"
        Else
          fOK = False
        End If
      
      Case 21  ' First Name from Forenames
        sCode = _
          "case" & _
          "    when charindex(' ', ltrim(" & sParamCode1 & ")) > 0 then substring(ltrim(" & sParamCode1 & "), 1, charindex(' ', ltrim(" & sParamCode1 & "))-1)" & _
          "    else ltrim(" & sParamCode1 & ")" & _
          "end"
      
      Case 22  ' Weekdays From Start and End Dates
        sCode = _
          " case" & _
          "    when datediff(day, " & sParamCode1 & ", " & sParamCode2 & ") <= 0 then 0" & _
          "    else datediff(day, " & sParamCode1 & ", " & sParamCode2 & " + 1)" & _
          "        - (2 * (datediff(day, " & sParamCode1 & " - (datepart(dw, " & sParamCode1 & ") - 1), " & _
          "                              " & sParamCode2 & " - (datepart(dw, " & sParamCode2 & ") - 1)) / 7))" & _
          "        - case" & _
          "              when datepart(dw, " & sParamCode1 & ") = 1 then 1" & _
          "              else 0" & _
          "          end" & _
          "        - case" & _
          "              when datepart(dw, " & sParamCode2 & ") = 7 then 1" & _
          "              else 0" & _
          "          end" & _
          " end"
    
      Case 23  ' Add months to date
        sCode = "dateadd(month, " & sParamCode2 & ", " & sParamCode1 & ")"
        
      Case 24  ' Add years to date
        sCode = "dateadd(year, " & sParamCode2 & ", " & sParamCode1 & ")"
  
      Case 25  ' Convert character to numeric.
        'JPD 20041213 Fault 9568
        'sCode = _
          " case" & _
          "    when isnumeric(" & sParamCode1 & ") = 1 then convert(float, " & sParamCode1 & ")" & _
          "    else 0" & _
          " end"
        sCode = _
          " case" & _
          "    when isnumeric(" & sParamCode1 & ") = 1 then convert(float, convert(money, " & sParamCode1 & "))" & _
          "    else 0" & _
          " end"
      
      Case 26  ' Whole Months between 2 Dates.
        sCode = _
          " case" & _
          "    when " & sParamCode1 & " >= " & sParamCode2 & " then 0" & _
          "    else datediff(month, " & sParamCode1 & ", " & sParamCode2 & ")" & _
          "        - case" & _
          "              when datepart(day, " & sParamCode2 & ") < datepart(day, " & sParamCode1 & ") then 1" & _
          "              else 0" & _
          "          end" & _
          " end"
    
      Case 27  ' Parentheses
        sCode = sParamCode1
  
      Case 28  ' Day of the week
        sCode = "DATEPART(weekday, " & sParamCode1 & ")"
      
      Case 29  ' Working Days per week
        sCode = "(convert(float, len(replace(left(" & sParamCode1 & ", 14), ' ', ''))) / 2)"
    
      Case 30 ' Absence Duration
        If fWorkflowExpression Then
          If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
            strTempTableName = GetTableName(mobjBaseComponent.ParentExpression.BaseTableID)
            If mobjBaseComponent.ParentExpression.BaseTableID = lngPersonnelTableID Then
              strTempTableID = "ID"
            Else
              strTempTableID = "ID_" & Trim(Str(lngPersonnelTableID))
            End If
  
            sCode = "(dbo.udf_ASRFn_AbsenceDuration(" & sParamCode1 & "," & sParamCode2 _
                & "," & sParamCode3 & "," & sParamCode4 & "," & strTempTableName & "." & strTempTableID & "))"
          Else
            sCode = "(dbo.udf_ASRFn_AbsenceDuration(" & sParamCode1 & "," & sParamCode2 _
                & "," & sParamCode3 & "," & sParamCode4 & "," & sVarName & "))"
          End If
        Else
          fOK = False
        End If
      
      Case 31  ' Round down to nearest whole number.
        sCode = "round(" & sParamCode1 & ", 0, 1)"
        
      Case 32  ' Year of date.
        sCode = "datepart(year, " & sParamCode1 & ")"
        
      Case 33  ' Month of date.
        sCode = "datepart(month, " & sParamCode1 & ")"
        
      Case 34  ' Day of date.
        sCode = "datepart(day, " & sParamCode1 & ")"
        
      Case 35  ' Nice Date
        sCode = _
          "datename(day, " & sParamCode1 & ") + ' ' + " & _
          "datename(month, " & sParamCode1 & ") + ' ' + " & _
          "datename(year, " & sParamCode1 & ")"
      
      Case 36  ' Nice Time
        'sCode = _
          "convert(varchar(2), datepart(hour, convert(datetime, " & sParamCode1 & ")) % 12) + ':' " & _
          "    + right('00' + datename(minute, convert(datetime, " & sParamCode1 & ")),2)" & _
          "    + case" & _
          "          when datepart(hour, convert(datetime, " & sParamCode1 & ")) > 11 then ' pm'" & _
          "          else ' am'" & _
          "      end"
        ' JPD20020617 Fault 3998
        sCode = _
          "case when len(ltrim(rtrim(" & sParamCode1 & "))) = 0 then ''" & _
          " else case when isdate(" & sParamCode1 & ") = 0 then '***'" & _
          " else (convert(varchar(2),((datepart(hour,convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) + 11) % 12) + 1)" & _
          " + ':' + right('00' + datename(minute, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)),2)" & _
          " + case when datepart(hour, convert(datetime, case when isdate(" & sParamCode1 & ") = 1 then " & sParamCode1 & " else '1:1' end)) > 11 then ' pm'" & _
          " else ' am' end) end end"
      
      Case 37  ' Round Date to Start of nearest month
        sCode = _
          " case" & _
          "     when datediff(day, (" & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1), " & sParamCode1 & ")" & _
          "         <= datediff(day, " & sParamCode1 & ", (dateadd(month, 1, " & sParamCode1 & ") - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1))" & _
          "         then " & sParamCode1 & " - datepart(day, " & sParamCode1 & ") + 1" & _
          "     else dateadd(month, 1, " & sParamCode1 & ")" & _
          "         - datepart(day, dateadd(month, 1, " & sParamCode1 & ")) + 1" & _
          " end"
      
      Case 38  ' Is Between
        sCode = _
          " case" & _
          "     when (" & sParamCode1 & " >= " & sParamCode2 & ")" & _
          "         and (" & sParamCode1 & " <= " & sParamCode3 & ") then 1" & _
          "     else 0" & _
          " end"
      
      Case 39  ' Service Years
        If fWorkflowExpression Then
          sCode = _
            "   datepart(year, case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end)" & _
            " - datepart(year, case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end)" & _
            " - case" & _
            "       when datepart(month, case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end)" & _
            "          > datepart(month, case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end) then 1" & _
            "      when (datepart(month, case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end)" & _
            "          = datepart(month, case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end))" & _
            "       and (datepart(day, case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end)" & _
            "          > datepart(day, case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end)) then 1" & _
            "      else 0" & _
            "  end"
        Else
          sCode = _
            "   datepart(year, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & _
            " - datepart(year, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
            " - case" & _
            "       when datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
            "          > datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) then 1" & _
            "      when (datepart(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
            "          = datepart(month, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end))" & _
            "       and (datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end)" & _
            "          > datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)) then 1" & _
            "      else 0" & _
            "  end"
        End If
        
      Case 40  ' Service Months
        If fWorkflowExpression Then
          sCode = _
            " (case" & _
            "    when case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end" & _
            "          >= case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end then 0" & _
            "    else datediff(month, case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end, case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end)" & _
            "        - case" & _
            "              when datepart(day, case when " & sParamCode2 & " is null then @dtDate else " & sParamCode2 & " end) < datepart(day, case when " & sParamCode1 & " is null then @dtDate else " & sParamCode1 & " end) then 1" & _
            "              else 0" & _
            "          end" & _
            " end) % 12"
        Else
          sCode = _
            " (case" & _
            "    when case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end" & _
            "          >= case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end then 0" & _
            "    else datediff(month, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end)" & _
            "        - case" & _
            "              when datepart(day, case when " & sParamCode2 & " is null then getdate() else " & sParamCode2 & " end) < datepart(day, case when " & sParamCode1 & " is null then getdate() else " & sParamCode1 & " end) then 1" & _
            "              else 0" & _
            "          end" & _
            " end) % 12"
        End If
      
      Case 42  ' Get field from database record.
        ' Get the column parameter definitions.
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", val(sParamCode1)
          fOK = Not .NoMatch
          
          If fOK Then
            sSrchColumnName = !ColumnName
            lngSrchTableID = !TableID
            
            recTabEdit.Index = "idxTableID"
            recTabEdit.Seek "=", lngSrchTableID
                     
            fOK = Not recTabEdit.NoMatch
            
            If fOK Then
              sSrchTableName = recTabEdit!TableName
            
              .Seek "=", val(sParamCode3)
              fOK = Not .NoMatch
            
              If fOK Then
                sRtnColumnName = !ColumnName
                lngRtnTableID = !TableID
              
                fOK = ((LenB(sSrchColumnName) <> 0) And (LenB(sRtnColumnName) <> 0)) And _
                  (lngSrchTableID = lngRtnTableID)
              End If
            End If
          End If
          
          If fOK Then
            sCode = "(SELECT TOP 1 dbo." & sSrchTableName & "." & sRtnColumnName & _
              " FROM dbo." & sSrchTableName & _
              " WHERE (dbo." & sSrchTableName & "." & sSrchColumnName & " = " & sParamCode2 & "))"
          End If
        End With

      Case 44  ' Add days to date.
        sCode = "dateadd(day, " & sParamCode2 & ", " & sParamCode1 & ")"
      
      Case 45  ' Days Between 2 Dates

        'MH20010220 Fault 1850
        'This function needs to be inclusive of both start and end
        'sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")"
        sCode = "datediff(dd, " & sParamCode1 & ", " & sParamCode2 & ")+1"

      Case 46  'Working days between two dates
        If fWorkflowExpression Then
          If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
            strTempTableName = GetTableName(mobjBaseComponent.ParentExpression.BaseTableID)
            
            If mobjBaseComponent.ParentExpression.BaseTableID = lngPersonnelTableID Then
              strTempTableID = "ID"
            Else
              strTempTableID = "ID_" & Trim(Str(lngPersonnelTableID))
            End If
          
            sCode = "(dbo.udf_ASRFn_WorkingDaysBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 & "," & strTempTableName & "." & strTempTableID & "))"
          Else
            sCode = "(dbo.udf_ASRFn_WorkingDaysBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 & "," & sVarName & "))"
          End If
        Else
          fOK = False
        End If

      Case 47  ' Absence between two dates
        If fWorkflowExpression Then
          If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
            strTempTableName = GetTableName(mobjBaseComponent.ParentExpression.BaseTableID)
  
            If mobjBaseComponent.ParentExpression.BaseTableID = lngPersonnelTableID Then
              strTempTableID = "ID"
            Else
              strTempTableID = "ID_" & Trim(Str(lngPersonnelTableID))
            End If
    
            sCode = "(dbo.udf_ASRFn_AbsenceBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 _
                & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "," & "convert(datetime,'" & Replace(Format(Now(), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')" & "))"
          Else
            sCode = "(dbo.udf_ASRFn_AbsenceBetweenTwoDates(" & sParamCode1 & "," & sParamCode2 _
                & "," & sParamCode3 & "," & sVarName & "," & "convert(datetime,'" & Replace(Format(Now(), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')" & "))"
          End If
        Else
          fOK = False
        End If

      Case 48  ' Round Up to nearest whole number.
        ' JPD20030116 Fault 4910
        sCode = "CASE WHEN (" & sParamCode1 & ") < 0 THEN floor(" & sParamCode1 & ")" & _
          " ELSE ceiling(" & sParamCode1 & ") END"

      Case 49  ' Round to nearest number.
        strRemainString = "(" + sParamCode1 + ") - ((floor(" + sParamCode1 + "/" + sParamCode2 + "))*" + sParamCode2 + ")"
        ' JPD20030116 Fault 4910
        sCode = "CASE WHEN (((" + sParamCode1 + ")<0) AND ((" + strRemainString + ")<=((" + sParamCode2 + ")/2.0)))" & _
          " OR (((" + sParamCode1 + ")>=0) AND ((" + strRemainString + ")<((" + sParamCode2 + ")/2.0)))" & _
          " THEN (" + sParamCode1 + ")-(" + strRemainString + ")" & _
          " ELSE (" + sParamCode1 + ")+(" + sParamCode2 + ")-(" + strRemainString + ") END"

        'MH20100629
        sCode = "CASE WHEN (" & sParamCode2 & " > 0) THEN " & sCode & " ELSE 0 END"

      'MH20010515
      Case 50  ' Is overnight process
        'sCode = "(SELECT UpdatingDateDependentColumns FROM ASRSysConfig)"
        sCode = "(SELECT SettingValue FROM ASRSysSystemSettings " & _
                 "WHERE [Section] = 'database' AND [SettingKey] = 'updatingdatedependantcolumns')"

      'TM20011022 Currency Implementation
      Case 51
      '*********** runtime code to go here *************

        ' Get the column parameter definitions.
        sSQL = vbNullString
        sSQL = sSQL & "SELECT ASRSysModuleSetup.*, ASRSysColumns.ColumnName, ASRSysTables.TableName "
        sSQL = sSQL & "FROM ASRSysModuleSetup "
        sSQL = sSQL & "     INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID "
        sSQL = sSQL & "                INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID "
        sSQL = sSQL & "WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY' "
        
        Set rsInfo = New ADODB.Recordset
        rsInfo.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly, adCmdText
        With rsInfo
          If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While Not .EOF
              sCConvTable = !TableName
              Select Case !parameterkey
              Case "Param_CurrencyNameColumn": sCConvCurrDescCol = !ColumnName
              Case "Param_ConversionValueColumn": sCConvExRateCol = !ColumnName
              Case "Param_DecimalColumn": sCConvDecCol = !ColumnName
              End Select
              .MoveNext
            Loop
            
            If (LenB(sCConvTable) <> 0) And (LenB(sCConvCurrDescCol) <> 0) And _
                (LenB(sCConvExRateCol) <> 0) And (LenB(sCConvDecCol) <> 0) Then
'              sCode = vbNullString
'              sCode = sCode & " ( CASE"
'              sCode = sCode & "  WHEN ((SELECT Currency_Conversion_Table.Units_per_Euro"
'              sCode = sCode & "                From Currency_Conversion_Table"
'              sCode = sCode & "                WHERE Currency_Conversion_Table.Currency_Name = " & sParamCode2 & ") <> 0) "
'              sCode = sCode & "  THEN (SELECT ROUND((" & sParamCode1
'              sCode = sCode & "              / "
'              sCode = sCode & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode2 & ") "
'              sCode = sCode & "              * "
'              sCode = sCode & "             (SELECT " & sCConvTable & "." & sCConvExRateCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & ")) "
'              sCode = sCode & "        , "
'              sCode = sCode & "        (SELECT " & sCConvTable & "." & sCConvDecCol & " FROM " & sCConvTable & " WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & ")) ) "
'              sCode = sCode & "  ELSE "
'              sCode = sCode & "    0"
'              sCode = sCode & "  END )"
              
              'AE20071204 Fault #12669
              sCode = vbNullString
              sCode = sCode & "ROUND(ISNULL((" & sParamCode1 & " / NULLIF((SELECT " & sCConvTable & "." & sCConvExRateCol
              sCode = sCode & "                                     FROM " & sCConvTable
              sCode = sCode & "                                     WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode2 & "), 0))"
              sCode = sCode & " * "
              sCode = sCode & "                                   (SELECT " & sCConvTable & "." & sCConvExRateCol
              sCode = sCode & "                                    FROM " & sCConvTable
              sCode = sCode & "                                    WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & "), 0)"
              sCode = sCode & "        , "
              sCode = sCode & "         ISNULL((SELECT " & sCConvTable & "." & sCConvDecCol
              sCode = sCode & "                 FROM " & sCConvTable
              sCode = sCode & "                 WHERE " & sCConvTable & "." & sCConvCurrDescCol & " = " & sParamCode3 & "), 0))"
            Else
              sCode = "null"
            End If
          Else
            sCode = "null"
          End If
          .Close
        End With
        Set rsInfo = Nothing
   
      ' Last field change date
      Case 52
        sCode = "(SELECT Top 1 DateTimeStamp FROM ASRSysAuditTrail WHERE ColumnID = " & sParamCode1
        sCode = sCode & " And " & GetColumnTableName(CLng(sParamCode1)) & ".ID = ASRSysAuditTrail.RecordID ORDER BY DateTimeStamp DESC)"
      
      ' Field changed between two dates
      Case 53
        sCode = " case when " _
          & " Exists(Select DateTimeStamp From ASRSysAuditTrail Where ColumnID = " & sParamCode1 _
          & " And " & GetColumnTableName(CLng(sParamCode1)) & ".ID = ASRSysAuditTrail.RecordID" _
          & " And DateTimeStamp >= " & sParamCode2 & " And DateTimeStamp <= " & sParamCode3 & ")" _
          & " then 1 else 0 end"
  
      'Whole years between two dates
      Case 54
        sCode = _
          " case " & _
          " when " & sParamCode1 & " >= " & sParamCode2 & _
          " then 0 " & _
          " else " & _
          "   datediff(year, " & sParamCode1 & ", " & sParamCode2 & ") " & _
          "   - " & _
          "   case " & _
          "   when DatePart(Month, " & sParamCode2 & ") < DatePart(Month, " & sParamCode1 & ") " & _
          "   then 1 " & _
          "   else " & _
          "     case " & _
          "     when DatePart(Month, " & sParamCode2 & ") = DatePart(Month, " & sParamCode1 & ") " & _
          "     then " & _
          "       case " & _
          "       when DatePart(Day, " & sParamCode2 & ") < DatePart(Day, " & sParamCode1 & ") " & _
          "       then 1 " & _
          "       else 0 " & _
          "       end " & _
          "     else " & _
          "       0 " & _
          "     end " & _
          "   end " & _
          " end "
          
      ' JPD20021121 Fault 3177
      Case 55  ' First Day of Month
        sCode = "dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")"
        
      ' JPD20021121 Fault 3177
      Case 56  ' Last Day of Month
        sCode = "dateadd(dd, -1, dateadd(mm, 1, dateadd(dd, 1 - datepart(dd, " & sParamCode1 & "), " & sParamCode1 & ")))"
        
      ' JPD20021121 Fault 3177
      Case 57  ' First Day of Year
        sCode = "dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")"
        
      ' JPD20021121 Fault 3177
      Case 58  ' Last Day of Year
        sCode = "dateadd(dd, -1, dateadd(yy, 1, dateadd(dd, 1 - datepart(dy, " & sParamCode1 & "), " & sParamCode1 & ")))"
        
      ' JPD20021129 Fault 4337
      Case 59  ' Name of Month.
        sCode = "datename(month, " & sParamCode1 & ")"
      
      ' JPD20021129 Fault 4337
      Case 60  ' Name of Day.
        sCode = "datename(weekday, " & sParamCode1 & ")"
              
      ' JPD20021129 Fault 3606
      Case 61  ' Is field populated
        sCode = "(CASE WHEN ((" & sParamCode1 & ") IS NULL)"

        ' Validate the sub-expression. This is done, not to  validate the expression,
        ' but rather to determine the return type of the expression.
        mcolParameters(1).Component.ValidateExpression False
      
        Select Case mcolParameters(1).ReturnType
          Case giEXPRVALUE_CHARACTER
            sCode = sCode & " OR ((" & sParamCode1 & ") = '')"
          Case giEXPRVALUE_NUMERIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
          Case giEXPRVALUE_LOGIC
            sCode = sCode & " OR ((" & sParamCode1 & ") = 0)"
        End Select
        
        sCode = sCode & " THEN 0 ELSE 1 END)"
  
      'Case 62 'PARENTAL LEAVE ENTITLEMENT
        ''' Workflow - future dev
      'Case 63 'PARENTAL LEAVE TAKEN
        ''' Workflow - future dev
      'Case 64 'MATERNITY EXPECTED RETURN DATE
        ''' Workflow - future dev
      'Case 65  ' Is Post Subordinate Of (NOT USED)
      
      Case 66  ' Is Post Subordinate Of User
        sHierarchyTableName = ""
        recModuleSetup.Index = "idxModuleParameter"
        recModuleSetup.Seek "=", gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE
        If Not recModuleSetup.NoMatch Then
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", recModuleSetup!parametervalue
          If Not recTabEdit.NoMatch Then
            sHierarchyTableName = recTabEdit!TableName
          End If
        End If

        If (LenB(sHierarchyTableName) <> 0) Then
          If fWorkflowExpression Then
            If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
              sCode = "CASE WHEN " & sHierarchyTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_IsPostSubordinateOfUser(" & sRequiredTableName & ".ID)) THEN 1 ELSE 0 END"
            Else
              sCode = "CASE WHEN " & sHierarchyTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_IsPostSubordinateOfUser(" & sVarName & ")) THEN 1 ELSE 0 END"
            End If
          Else
            sCode = "CASE WHEN " & sHierarchyTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPostSubordinateOfUser()) THEN 1 ELSE 0 END"
          End If
        Else
          sCode = "0"
        End If
              
      'Case 67  ' Is Personnel Subordinate Of (NOT USED)
      
      Case 68  ' Is Personnel Subordinate Of User
        sPersonnelTableName = ""
        recModuleSetup.Index = "idxModuleParameter"
        recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
        If Not recModuleSetup.NoMatch Then
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", recModuleSetup!parametervalue
          If Not recTabEdit.NoMatch Then
            sPersonnelTableName = recTabEdit!TableName
          End If
        End If

        If (LenB(sPersonnelTableName) <> 0) Then
          If fWorkflowExpression Then
            If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
              sCode = "CASE WHEN dbo." & sPersonnelTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_IsPersonnelSubordinateOfUser(" & sRequiredTableName & ".ID)) THEN 1 ELSE 0 END"
            Else
              sCode = "CASE WHEN dbo." & sPersonnelTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_IsPersonnelSubordinateOfUser(" & sVarName & ")) THEN 1 ELSE 0 END"
            End If
          Else
            sCode = "CASE WHEN dbo." & sPersonnelTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_IsPersonnelSubordinateOfUser()) THEN 1 ELSE 0 END"
          End If
        Else
          sCode = "0"
        End If
  
      'Case 69 'Has Post Subordinate (NOT USED)
      
      Case 70 'Has Post Subordinate User
        sHierarchyTableName = ""
        recModuleSetup.Index = "idxModuleParameter"
        recModuleSetup.Seek "=", gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE
        If Not recModuleSetup.NoMatch Then
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", recModuleSetup!parametervalue
          If Not recTabEdit.NoMatch Then
            sHierarchyTableName = recTabEdit!TableName
          End If
        End If

        If (LenB(sHierarchyTableName) <> 0) Then
          If fWorkflowExpression Then
            If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
              sCode = "CASE WHEN " & sHierarchyTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_HasPostSubordinateUser(" & sRequiredTableName & ".ID)) THEN 1 ELSE 0 END"
            Else
              sCode = "CASE WHEN " & sHierarchyTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_HasPostSubordinateUser(" & sVarName & ")) THEN 1 ELSE 0 END"
            End If
          Else
            sCode = "CASE WHEN " & sHierarchyTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPostSubordinateUser()) THEN 1 ELSE 0 END"
          End If
        Else
          sCode = "0"
        End If
  
      'Case 71 'Has Personnel Subordinate (NOT USED)
      
      Case 72 'Has Personnel Subordinate User
        sPersonnelTableName = ""
        recModuleSetup.Index = "idxModuleParameter"
        recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
        If Not recModuleSetup.NoMatch Then
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", recModuleSetup!parametervalue
          If Not recTabEdit.NoMatch Then
            sPersonnelTableName = recTabEdit!TableName
          End If
        End If

        If (LenB(sPersonnelTableName) <> 0) Then
          If fWorkflowExpression Then
            If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
              sCode = "CASE WHEN " & sPersonnelTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_HasPersonnelSubordinateUser(" & sRequiredTableName & ".ID)) THEN 1 ELSE 0 END"
            Else
              sCode = "CASE WHEN " & sPersonnelTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_ByID_HasPersonnelSubordinateUser(" & sVarName & ")) THEN 1 ELSE 0 END"
            End If
          Else
            sCode = "CASE WHEN " & sPersonnelTableName & ".id IN (SELECT id FROM dbo.udf_ASRFn_HasPersonnelSubordinateUser()) THEN 1 ELSE 0 END"
          End If
        Else
          sCode = "0"
        End If
      
      Case 73 ' Bradford Factor
        iPersonnelTableID = 0
        sPersonnelTableName = ""
        recModuleSetup.Index = "idxModuleParameter"
        recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
        If Not recModuleSetup.NoMatch Then
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", recModuleSetup!parametervalue
          If Not recTabEdit.NoMatch Then
            iPersonnelTableID = recTabEdit!TableID
            sPersonnelTableName = recTabEdit!TableName
          End If
        End If
        
        If mobjBaseComponent.ParentExpression.BaseTableID = iPersonnelTableID Then
          sSrchColumnName = "ID"
        Else
          sSrchColumnName = "ID_" & Trim(Str(iPersonnelTableID))
        End If
        
        If fWorkflowExpression Then
          If (miWorkflowRecordSelection = giWFRECSEL_UNIDENTIFIED) Then
            strTempTableName = GetTableName(mobjBaseComponent.ParentExpression.BaseTableID)
            If mobjBaseComponent.ParentExpression.BaseTableID = lngPersonnelTableID Then
              strTempTableID = "ID"
            Else
              strTempTableID = "ID_" & Trim(Str(lngPersonnelTableID))
            End If

            sCode = "(dbo.udf_ASRFn_BradfordFactor(" & sParamCode1 & "," & sParamCode2 _
                & "," & sParamCode3 & "," & strTempTableName & "." & strTempTableID & "))"
          Else
            sCode = "(dbo.udf_ASRFn_BradfordFactor(" & sParamCode1 & "," & sParamCode2 _
                & "," & sParamCode3 & "," & sVarName & "))"
          End If
        Else
          sCode = "(dbo.udf_ASRFn_BradfordFactor(" & sParamCode1 & "," & sParamCode2 _
              & "," & sParamCode3 & "," & sPersonnelTableName & "." & sSrchColumnName & "))"
        End If
        
      Case 74 ' Does Record Exist (Workflow only)
        sCode = "CASE WHEN (SELECT COUNT(*) FROM " & sRequiredTableName & " WHERE ID = " & sVarName & ") > 0 THEN 1 ELSE 0 END"
      
      Case Else
        fOK = False
        
    End Select
  End If
  
  If fOK Then
    ' We need to convert date values to varchars in the format 'mm/dd/yyyy'.
    If miReturnType = giEXPRVALUE_DATE Then
      sCode = "convert(" & vbNewLine & _
        "datetime, " & vbNewLine & _
        "convert(" & vbNewLine & _
        "varchar(20), " & vbNewLine & _
        sCode & "," & vbNewLine & _
        "101)" & vbNewLine & _
        ")"
    End If
  End If
  
TidyUpAndExit:
  If fOK Then
    
    ' JDM - 26/11/01 - Fault 3204 - Needs some brackets around these functions
    'psRuntimeCode = sCode
    psRuntimeCode = "(" & sCode & ")"

  Else
    psRuntimeCode = ""
  End If
  RuntimeCode = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function












Private Function NonStandardFunction_Evaluate_42(pLngRecordID As Long, pfDoPrompts As Boolean) As Variant
' JPD - NO LONGER REQUIRED.
'  ' Return the evaluated result of the function.
'  On Error GoTo BasicErrorTrap
'
'  Dim sSQL As String
'  Dim sParameterSQL As String
'  Dim vResult As Variant
'  Dim vValue As Variant
'  Dim objSubExpression As CExpression
'  Dim rsResult As rdoResultset
'
'  ' Evaluate the function parameter expressions.
'  ' Write each parameter's evaluated value into the array.
'  sParameterSQL = ""
'  Set objSubExpression = gaParameterArray(2).Component
'  vValue = objSubExpression.EvaluateExpression(pLngRecordID, pfDoPrompts)
'  Set objSubExpression = Nothing
'
'  Select Case VarType(vValue)
'    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
'      sParameterSQL = "null, null, " & vValue & ", null, "
'    Case vbDate
'      sParameterSQL = "null, null, null, '" & Format(vValue, "MM/dd/yyyy") & "', "
'    Case vbString
'      sParameterSQL = "'" & vValue & "', null, null, null, "
'    Case vbBoolean
'      sParameterSQL = "null, " & IIf(vValue, 1, 0) & ", null, null, "
'  End Select
'
'  sSQL = "DECLARE @charRes varchar(MAX)," & vbnewline & _
'    "    @bitRes bit," & vbnewline & _
'    "    @numRes float," & vbnewline & _
'    "    @dateRes datetime" & vbnewline & _
'    "EXEC sp_ASRFn_GetFieldFromDatabase @charRes OUTPUT, @bitRes OUTPUT, @numRes OUTPUT, @dateRes OUTPUT, " & _
'      trim$(str$(gaParameterArray(1).Component.Components(1).Component.ColumnID)) & ", " & _
'      sParameterSQL & _
'      trim$(str$(gaParameterArray(3).Component.Components(1).Component.ColumnID))
'
'    Set rsResult = rdoCon.OpenResultset(sSQL, _
'      rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
'    With rsResult
'      If Not (.EOF And .BOF) Then
'        Select Case miReturnType
'          Case giEXPRVALUE_NUMERIC
'            vResult = IIf(IsNull(!result), 0, !result)
'          Case giEXPRVALUE_DATE
'            vResult = IIf(IsNull(!result), Null, !result)
'          Case giEXPRVALUE_CHARACTER
'            vResult = IIf(IsNull(!result), "", !result)
'          Case giEXPRVALUE_LOGIC
'            vResult = IIf(IsNull(!result), False, !result)
'        End Select
'      Else
'        Select Case miReturnType
'          Case giEXPRVALUE_NUMERIC
'            vResult = 0
'          Case giEXPRVALUE_DATE
'            vResult = Null
'          Case giEXPRVALUE_CHARACTER
'            vResult = ""
'          Case giEXPRVALUE_LOGIC
'            vResult = False
'        End Select
'      End If
'      .Close
'    End With
'    Set rsResult = Nothing
'
'  ' Determine the return type of the function with the given parameters.
'  Select Case VarType(vResult)
'    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
'      miReturnType = giEXPRVALUE_NUMERIC
'    Case vbDate
'      miReturnType = giEXPRVALUE_DATE
'    Case vbString
'      miReturnType = giEXPRVALUE_CHARACTER
'    Case vbBoolean
'      miReturnType = giEXPRVALUE_LOGIC
'  End Select
'
'TidyUpAndExit:
'  ' Disassociate object variables.
'  Set objSubExpression = Nothing
'  NonStandardFunction_Evaluate_42 = vResult
'  Exit Function
'
'BasicErrorTrap:
'  Resume TidyUpAndExit
'
End Function
Public Function CopyToClipboard(piLevel As Integer) As Boolean
  ' Copy the component definition to the clipboard object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objParameter As CExprComponent
  Dim sClipboardMessage As String

  fOK = True
  sClipboardMessage = String(piLevel, vbTab) & "Function : " & ComponentDescription & vbNewLine
  
  Clipboard.SetText Clipboard.GetText & sClipboardMessage
  
  ' Print the function's parameter expressions.
  For Each objParameter In mcolParameters
    fOK = objParameter.Component.CopyToClipboard(piLevel + 1)
  Next objParameter

TidyUpAndExit:
  CopyToClipboard = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function GetColumnTableName(lngColumnID As Long) As String
  On Error GoTo ErrorTrap
  
  Dim lngTableID As Long
  
  GetColumnTableName = vbNullString
  
  'JPD 20051130 Fault 10621
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID

    If Not .NoMatch Then
      lngTableID = !TableID
      
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTableID

      If Not recTabEdit.NoMatch Then
        GetColumnTableName = recTabEdit!TableName
      End If
    End If
  End With
    
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mcolParameters = Nothing
  
End Sub


'MH20060727  New code not implemented
'Private Sub Class_Terminate()
'  Terminate
'End Sub
'
'Public Sub Terminate()
'
'  If Not mcolParameters Is Nothing Then
'    Do While mcolParameters.Count > 0
'      mcolParameters(1).Terminate
'      mcolParameters.Remove 1
'    Loop
'    Set mcolParameters = Nothing
'  End If
'
'  Set mobjBaseComponent = Nothing
'  Set mobjBadComponent = Nothing
'
'End Sub

' Code stub for .NET migration
Public Function GetSQLElementCode() As String
  GetSQLElementCode = vbNullString
End Function

