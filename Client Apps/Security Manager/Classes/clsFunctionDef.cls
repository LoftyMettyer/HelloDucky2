VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctionDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_lngID As Long
Private mvar_sName As String
Private mvar_iReturnType As Integer
Private mvar_fTimeDependent As Boolean
Private mvar_sCategory As String
Private mvar_sSPName As String
Private mvar_fNonStandard As Boolean
Private mvar_fRuntime As Boolean
Private mvar_sShortcutKeys As String
Private mvar_fUDF As Boolean
Private mvar_sExcludeTypes As String
Private mvar_sIncludeTypes As String
  
Private mvar_objParameters As clsFunctionDefParams

Public Property Get Parameters() As clsFunctionDefParams
  ' Create the parameters collection if it doesn't exist.
  If mvar_objParameters Is Nothing Then
    Set mvar_objParameters = New clsFunctionDefParams
  End If

  Set Parameters = mvar_objParameters
    
End Property

Public Property Set Parameters(pobjData As clsFunctionDefParams)
  Set mvar_objParameters = pobjData
End Property

Public Property Get Runtime() As Boolean
  Runtime = mvar_fRuntime
End Property

Public Property Get TimeDependent() As Boolean
  TimeDependent = mvar_fTimeDependent
End Property

Public Property Let NonStandard(ByVal pfNewValue As Boolean)
  mvar_fNonStandard = pfNewValue
End Property

Public Property Get Category() As String
  Category = mvar_sCategory
End Property

Public Property Let TimeDependent(ByVal pfNewValue As Boolean)
  mvar_fTimeDependent = pfNewValue
End Property

Public Property Get Name() As String
  Name = mvar_sName
End Property

Public Property Let Runtime(ByVal pfNewValue As Boolean)
  mvar_fRuntime = pfNewValue
End Property

Public Property Get NonStandard() As Boolean
  NonStandard = mvar_fNonStandard
End Property

Public Property Let Name(ByVal psNewValue As String)
  mvar_sName = psNewValue
End Property

Public Property Get ID() As Long
  ID = mvar_lngID
End Property

Public Property Get ShortcutKeys() As String
  ShortcutKeys = mvar_sShortcutKeys
End Property

Public Property Let ShortcutKeys(ByVal psNewValue As String)
  mvar_sShortcutKeys = psNewValue
End Property

Public Property Let SPName(ByVal psNewValue As String)
  mvar_sSPName = psNewValue
End Property

Public Property Let ReturnType(ByVal piNewValue As Integer)
  mvar_iReturnType = piNewValue
End Property

Public Property Get ReturnType() As Integer
  ReturnType = mvar_iReturnType
End Property

Public Property Get SPName() As String
  SPName = mvar_sSPName
End Property

Public Property Let ID(ByVal plngNewValue As Long)
  mvar_lngID = plngNewValue
End Property

Public Property Let Category(ByVal psNewValue As String)
  mvar_sCategory = psNewValue
End Property

Public Property Let UDF(ByVal pfNewValue As Boolean)
  mvar_fUDF = pfNewValue
End Property

Public Property Get UDF() As Boolean
  UDF = mvar_fUDF And gbEnableUDFFunctions
End Property

Public Property Get ExcludeTypes() As String
  ExcludeTypes = mvar_sExcludeTypes
End Property

Public Property Get IncludeTypes() As String
  IncludeTypes = mvar_sIncludeTypes
  
End Property


Public Property Let ExcludeTypes(ByVal psNewValue As String)
  mvar_sExcludeTypes = psNewValue
End Property
Public Property Let IncludeTypes(ByVal psNewValue As String)
  mvar_sIncludeTypes = psNewValue
  
End Property

