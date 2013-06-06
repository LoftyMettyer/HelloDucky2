VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctionDefParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_iIndex As Integer
Private mvar_iType As ExpressionValueTypes
Private mvar_sName As String


Public Property Get Name() As String
  Name = mvar_sName

End Property

Public Property Let Name(ByVal psNewValue As String)
  mvar_sName = psNewValue

End Property

Public Property Get Index() As Integer
  Index = mvar_iIndex

End Property



Public Property Let ParameterType(ByVal piNewValue As ExpressionValueTypes)
  mvar_iType = piNewValue

End Property

Public Property Get ParameterType() As ExpressionValueTypes
  ParameterType = mvar_iType

End Property






Public Property Let Index(ByVal piNewValue As Integer)
  mvar_iIndex = piNewValue

End Property







