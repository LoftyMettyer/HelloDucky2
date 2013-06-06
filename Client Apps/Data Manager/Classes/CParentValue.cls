VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParentValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbSelect As Boolean
Private mlIDColumn As Long
Private msTable As String
Private msColumn As String
Private msDefaultValue As String

Public Property Get CanSelect() As Boolean

    CanSelect = mbSelect

End Property

Public Property Let CanSelect(ByVal bSelect As Boolean)

    mbSelect = bSelect

End Property

Public Property Get IDColumn() As Long

    IDColumn = mlIDColumn

End Property

Public Property Let IDColumn(ByVal lIDColumn As Long)

    mlIDColumn = lIDColumn

End Property

Public Property Get Tablename() As String

    Tablename = msTable

End Property
Public Property Get DefaultValue() As String
  DefaultValue = msDefaultValue

End Property

Public Property Let Tablename(ByVal sTable As String)

    msTable = sTable

End Property

Public Property Let DefaultValue(ByVal psNewValue As String)

    msDefaultValue = psNewValue

End Property


Public Property Get Columnname() As String

    Columnname = msColumn

End Property

Public Property Let Columnname(ByVal sColumn As String)

    msColumn = sColumn

End Property
