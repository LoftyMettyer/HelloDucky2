VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctionDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_fInitialised As Boolean
Private mvar_colFunctions As Collection

Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mvar_colFunctions = New Collection

End Sub


Private Sub Class_Terminate()
  ' Destroy collection when this class is terminated.
  If Not mvar_colFunctions Is Nothing Then
    Set mvar_colFunctions = Nothing
  End If

End Sub


Private Sub Add(plngID As Long, _
  psName As String, _
  piReturnType As Integer, _
  pfTimeDependent As Boolean, _
  psCategory As String, _
  psSPName As String, _
  pfNonStandard As Boolean, _
  pfRuntime As Boolean, _
  pfUDF As Boolean, _
  psShortcutKeys As String, _
  psExcludeTypes As String, _
  psIncludeTypes As String)
  
  ' Add the given function definition to the collection of function definitions.
  Dim objNewMember As clsFunctionDef
  Dim objParameters As clsFunctionDefParams
  Dim rsParameters As New ADODB.Recordset
  Dim sSQL As String
  Dim iIndex As Integer
  Dim iType As Integer
  Dim sName As String

  ' Instantiate a new function definition.
  Set objNewMember = New clsFunctionDef
  Set objParameters = New clsFunctionDefParams

  With objNewMember
    ' Set the properties passed into the method.
    .ID = plngID
    .Name = psName
    .ReturnType = piReturnType
    .TimeDependent = pfTimeDependent
    .Category = psCategory
    .SPName = psSPName
    .NonStandard = pfNonStandard
    .Runtime = pfRuntime
    .ShortcutKeys = psShortcutKeys
    .UDF = pfUDF
    .ExcludeTypes = psExcludeTypes
    .IncludeTypes = psIncludeTypes
    
    ' Read the function parameters.
    sSQL = "SELECT * FROM ASRSysFunctionParameters" & _
      " WHERE functionID = " & Trim(Str(plngID))
    
    rsParameters.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    With rsParameters
      Do While Not .EOF
        iIndex = IIf(IsNull(!parameterIndex), 0, !parameterIndex)
        iType = IIf(IsNull(!ParameterType), 0, !ParameterType)
        sName = IIf(IsNull(!parameterName), "", !parameterName)

        objParameters.Add iIndex, iType, sName
        
        .MoveNext
      Loop
      .Close
    End With
    Set rsParameters = Nothing
    
    Set .Parameters = objParameters
  End With

  mvar_colFunctions.Add objNewMember, "F" & Trim(Str(plngID))

  Set objNewMember = Nothing

End Sub

Public Function IsValidID(plngID As Long) As Boolean
  ' Return TRUE if the given function ID exists in the collection.
  Dim Item As clsFunctionDef
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colFunctions("F" & Trim(Str(plngID)))
  IsValidID = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValidID = False
  
End Function

Public Property Get Collection() As Collection
  Set Collection = mvar_colFunctions

End Property
Public Property Get Count() As Long
  Count = mvar_colFunctions.Count
  
End Property


Public Property Get Item(pvIndexKey As Variant) As clsFunctionDef
  Set Item = mvar_colFunctions(pvIndexKey)
  
End Property


Public Sub Remove(pvIndexKey As Variant)
  mvar_colFunctions.Remove pvIndexKey
  
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
  Dim rsFunctions As New ADODB.Recordset
  Dim lngID As Long
  Dim sName As String
  Dim iReturnType As Integer
  Dim fTimeDependent As Boolean
  Dim sCategory As String
  Dim sSPName As String
  Dim fNonStandard As Boolean
  Dim fRuntime As Boolean
  Dim sShortcutKeys As String
  Dim fUDF As Boolean
  Dim sExcludeTypes As String
  Dim sIncludeTypes As String
  
  If Initialised Then Exit Function
  
  sSQL = "SELECT * FROM ASRSysFunctions"
  rsFunctions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsFunctions
    Do While Not .EOF
      
      lngID = IIf(IsNull(!FunctionID), 0, !FunctionID)
      sName = IIf(IsNull(!functionName), "", !functionName)
      iReturnType = IIf(IsNull(!ReturnType), 0, !ReturnType)
      fTimeDependent = IIf(IsNull(!TimeDependent), False, !TimeDependent)
      sCategory = IIf(IsNull(!Category), "", !Category)
      sSPName = IIf(IsNull(!SPName), "", !SPName)
      fNonStandard = IIf(IsNull(!NonStandard), False, !NonStandard)
      fRuntime = IIf(IsNull(!Runtime), False, !Runtime)
      sShortcutKeys = IIf(IsNull(!ShortcutKeys), "", UCase(!ShortcutKeys))
      fUDF = IIf(IsNull(!UDF), False, !UDF)
      sExcludeTypes = IIf(IsNull(!ExcludeExprTypes), False, !ExcludeExprTypes)
      sIncludeTypes = Trim(IIf(IsNull(!includeExprTypes), "", !includeExprTypes))

      If lngID > 0 Then
        Add lngID, sName, iReturnType, fTimeDependent, sCategory, _
          sSPName, fNonStandard, fRuntime, fUDF, sShortcutKeys, sExcludeTypes, sIncludeTypes
      End If
  
      .MoveNext
    Loop
    .Close
  End With
  Set rsFunctions = Nothing
    
  Initialised = True
  
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colFunctions.[_NewEnum]
End Property



