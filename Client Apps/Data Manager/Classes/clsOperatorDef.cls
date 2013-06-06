VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperatorDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_lngID As Long
Private mvar_sName As String
Private mvar_iReturnType As Integer
Private mvar_iPrecedence As Integer
Private mvar_iOperandCount As Integer
Private mvar_sCategory As String
Private mvar_sSQLCode As String
Private mvar_sSQLType As String
Private mvar_sSPName As String
Private mvar_fCheckDivideByZero As Boolean
Private mvar_sSQLFixedParam1 As String
Private mvar_fCastAsFloat As Boolean
Private mvar_sShortcutKeys As String

Private mvar_objParameters As clsOperatorDefParameters

Public Property Get Parameters() As clsOperatorDefParameters
  ' Create the parameters collection if it doesn't exist.
  If mvar_objParameters Is Nothing Then
    Set mvar_objParameters = New clsOperatorDefParameters
  End If

  Set Parameters = mvar_objParameters
    
End Property
Public Property Set Parameters(pobjData As clsOperatorDefParameters)
  Set mvar_objParameters = pobjData
    
End Property
Public Property Get SQLFixedParam1() As String
  SQLFixedParam1 = mvar_sSQLFixedParam1

End Property

Public Property Let SQLType(ByVal psNewValue As String)
  mvar_sSQLType = psNewValue

End Property

Public Property Get SQLCode() As String
  SQLCode = mvar_sSQLCode

End Property

Public Property Let CheckDivideByZero(ByVal pfNewValue As Boolean)
  mvar_fCheckDivideByZero = pfNewValue

End Property

Public Property Get Category() As String
  Category = mvar_sCategory

End Property

Public Property Let CastAsFloat(ByVal pfNewValue As Boolean)
  mvar_fCastAsFloat = pfNewValue

End Property

Public Property Get Name() As String
  Name = mvar_sName

End Property

Public Property Let SQLFixedParam1(ByVal psNewValue As String)
  mvar_sSQLFixedParam1 = psNewValue

End Property

Public Property Get CheckDivideByZero() As Boolean
  CheckDivideByZero = mvar_fCheckDivideByZero

End Property

Public Property Let Name(ByVal psNewValue As String)
  mvar_sName = psNewValue

End Property

Public Property Get ID() As Long
  ID = mvar_lngID

End Property

Public Property Let OperandCount(ByVal piNewValue As Integer)
  mvar_iOperandCount = piNewValue

End Property

Public Property Get ShortcutKeys() As String
  ShortcutKeys = mvar_sShortcutKeys

End Property

Public Property Let ShortcutKeys(ByVal psNewValue As String)
  mvar_sShortcutKeys = psNewValue

End Property

Public Property Get CastAsFloat() As Boolean
  CastAsFloat = mvar_fCastAsFloat

End Property

Public Property Let SPName(ByVal psNewValue As String)
  mvar_sSPName = psNewValue

End Property

Public Property Get Precedence() As Integer
  Precedence = mvar_iPrecedence

End Property

Public Property Let ReturnType(ByVal piNewValue As Integer)
  mvar_iReturnType = piNewValue

End Property

Public Property Get ReturnType() As Integer
  ReturnType = mvar_iReturnType

End Property

Public Property Let Precedence(ByVal piNewValue As Integer)
  mvar_iPrecedence = piNewValue

End Property

Public Property Get OperandCount() As Integer
  OperandCount = mvar_iOperandCount

End Property

Public Property Let SQLCode(ByVal psNewValue As String)
  mvar_sSQLCode = psNewValue

End Property

Public Property Get SPName() As String
  SPName = mvar_sSPName

End Property

Public Property Let ID(ByVal plngNewValue As Long)
  mvar_lngID = plngNewValue

End Property

Public Property Get SQLType() As String
  SQLType = mvar_sSQLType

End Property

Public Property Let Category(ByVal psNewValue As String)
  mvar_sCategory = psNewValue

End Property


