VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperatorDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_fInitialised As Boolean
Private mvar_colOperators As Collection

Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mvar_colOperators = New Collection

End Sub


Private Sub Class_Terminate()
  ' Destroy collection when this class is terminated.
  If Not mvar_colOperators Is Nothing Then
    Set mvar_colOperators = Nothing
  End If

End Sub


Private Sub Add(plngID As Long, _
  psName As String, _
  piReturnType As Integer, _
  piPrecedence As Integer, _
  piOperandCount As Integer, _
  psCategory As String, _
  psSQLCode As String, _
  psSQLType As String, _
  psSPName As String, _
  pfCheckDivideByZero As Boolean, _
  psSQLFixedParam1 As String, _
  pfCastAsFloat As Boolean, _
  psShortcutKeys As String)
  ' Add the given operator definition to the collection of operator definitions.
  
  Dim objNewMember As clsOperatorDef
  Dim objParameters As clsOperatorDefParams
  Dim rsParameters As New ADODB.Recordset
  Dim sSQL As String
  Dim iIndex As Integer
  Dim iType As Integer
  Dim sName As String

  ' Instantiate a new operator definition.
  Set objNewMember = New clsOperatorDef
  Set objParameters = New clsOperatorDefParams

  With objNewMember
    ' Set the properties passed into the method.
    .ID = plngID
    .Name = psName
    .ReturnType = piReturnType
    .Precedence = piPrecedence
    .OperandCount = piOperandCount
    .Category = psCategory
    .SQLCode = psSQLCode
    .SQLType = psSQLType
    .SPName = psSPName
    .CheckDivideByZero = pfCheckDivideByZero
    .SQLFixedParam1 = psSQLFixedParam1
    .CastAsFloat = pfCastAsFloat
    .ShortcutKeys = psShortcutKeys
  
    ' Read the operator parameters.
    sSQL = "SELECT * FROM ASRSysOperatorParameters" & _
      " WHERE operatorID = " & Trim(Str(plngID))
    rsParameters.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    With rsParameters
      Do While Not .EOF
        iIndex = IIf(IsNull(!parameterIndex), 0, !parameterIndex)
        iType = IIf(IsNull(!ParameterType), 0, !ParameterType)

        objParameters.Add iIndex, iType
        
        .MoveNext
      Loop
      .Close
    End With
    Set rsParameters = Nothing
    
    Set .Parameters = objParameters
  End With

  mvar_colOperators.Add objNewMember, "O" & Trim(Str(plngID))

  Set objNewMember = Nothing

End Sub


Public Property Get Collection() As Collection
  Set Collection = mvar_colOperators

End Property
Public Property Get Count() As Long
  Count = mvar_colOperators.Count
  
End Property


Public Property Get Item(pvIndexKey As Variant) As clsOperatorDef
  Set Item = mvar_colOperators(pvIndexKey)
  
End Property


Public Sub Remove(pvIndexKey As Variant)
  mvar_colOperators.Remove pvIndexKey
  
End Sub

Public Property Get Initialised() As Boolean
  Initialised = mvar_fInitialised

End Property

Public Property Let Initialised(ByVal pfNewValue As Boolean)
  mvar_fInitialised = pfNewValue

End Property

Public Function Initialise()
  ' Initialise the collection if it has not already been initialised.
  Dim sSQL As String
  Dim rsOperators As New ADODB.Recordset
  Dim lngID As Long
  Dim sName As String
  Dim iReturnType As Integer
  Dim iPrecedence As Integer
  Dim iOperandCount As Integer
  Dim sCategory As String
  Dim sSQLCode As String
  Dim sSQLType As String
  Dim sSPName As String
  Dim fCheckDivideByZero As Boolean
  Dim sSQLFixedParam1 As String
  Dim fCastAsFloat As Boolean
  Dim sShortcutKeys As String
  
  If Initialised Then Exit Function
  
  sSQL = "SELECT * FROM ASRSysOperators"
  rsOperators.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsOperators
    Do While Not .EOF
      
      lngID = IIf(IsNull(!OperatorID), 0, !OperatorID)
      sName = IIf(IsNull(!Name), "", !Name)
      iReturnType = IIf(IsNull(!ReturnType), 0, !ReturnType)
      iPrecedence = IIf(IsNull(!Precedence), 0, !Precedence)
      iOperandCount = IIf(IsNull(!OperandCount), 0, !OperandCount)
      sCategory = IIf(IsNull(!Category), "", !Category)
      sSQLCode = IIf(IsNull(!SQLCode), "", !SQLCode)
      sSQLType = IIf(IsNull(!SQLType), "", !SQLType)
      sSPName = IIf(IsNull(!SPName), "", !SPName)
      fCheckDivideByZero = IIf(IsNull(!CheckDivideByZero), False, !CheckDivideByZero)
      sSQLFixedParam1 = IIf(IsNull(!SQLFixedParam1), "", !SQLFixedParam1)
      fCastAsFloat = IIf(IsNull(!CastAsFloat), False, !CastAsFloat)
      sShortcutKeys = IIf(IsNull(!ShortcutKeys), "", UCase(!ShortcutKeys))
  
      If lngID > 0 Then
        Add lngID, sName, iReturnType, iPrecedence, iOperandCount, sCategory, sSQLCode, _
          sSQLType, sSPName, fCheckDivideByZero, sSQLFixedParam1, fCastAsFloat, sShortcutKeys
      End If
  
      .MoveNext
    Loop
    .Close
  End With
  Set rsOperators = Nothing
    
  Initialised = True
  
End Function
Public Function IsValidID(plngID As Long) As Boolean
  ' Return TRUE if the given operator ID exists in the collection.
  Dim Item As clsOperatorDef
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colOperators("O" & Trim(Str(plngID)))
  IsValidID = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValidID = False
  
End Function
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colOperators.[_NewEnum]
  
End Property


