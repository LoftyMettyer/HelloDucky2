VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSummaryField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private miSequence As Integer
Private mfStartOfGroup As Boolean
Private msColumnName As String
Private mlngColumnID As Long
Private mlngTableID As Long
Private miDataType As Integer
Private mlngSize As Long
Private miDecimals As Integer
Private miControlType As Integer
Private miColumnType As Integer
Private mfMultiline As Boolean
Private mlngWidth As Long
Private miAlignment As Integer
Private mbStartOfColumn As Boolean
Private mbUse1000Separator As Boolean
Private mbBlankIfZero As Boolean

Public Property Get Sequence() As Integer
  Sequence = miSequence
  
End Property

Public Property Let Sequence(ByVal piNewValue As Integer)
  miSequence = piNewValue
  
End Property

Public Property Get StartOfGroup() As Boolean
  StartOfGroup = mfStartOfGroup
  
End Property

Public Property Let StartOfGroup(ByVal pfNewValue As Boolean)
  mfStartOfGroup = pfNewValue
  
End Property

Public Property Get ColumnName() As String
  ColumnName = msColumnName
  
End Property

Public Property Let ColumnName(ByVal psNewValue As String)
  msColumnName = psNewValue
  
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
  
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  mlngColumnID = plngNewValue
  
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
  
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  mlngTableID = plngNewValue
  
End Property

Public Property Get DataType() As Integer
  DataType = miDataType
  
End Property

Public Property Get Alignment() As Integer
  Alignment = miAlignment
  
End Property


Public Property Let DataType(ByVal piNewValue As Integer)
  miDataType = piNewValue
  
End Property

Public Property Let Alignment(ByVal piNewValue As Integer)
  miAlignment = piNewValue
  
End Property

Public Property Get Size() As Long
  Size = mlngSize
End Property

Public Property Let Size(ByVal plngNewValue As Long)
  mlngSize = plngNewValue
End Property

Public Property Get Decimals() As Integer
  Decimals = miDecimals
  
End Property

Public Property Let Decimals(ByVal piNewValue As Integer)
  miDecimals = piNewValue
  
End Property

Public Property Get ControlType() As Integer
  ControlType = miControlType
  
End Property

Public Property Let ControlType(ByVal piNewValue As Integer)
  miControlType = piNewValue
  
End Property

Public Property Get ColumnType() As Integer
  ColumnType = miColumnType
  
End Property

Public Property Let ColumnType(ByVal piNewValue As Integer)
  miColumnType = piNewValue
  
End Property

Public Property Get Multiline() As Boolean
  Multiline = mfMultiline
  
End Property

Public Property Let Multiline(ByVal pfNewValue As Boolean)
  mfMultiline = pfNewValue
  
End Property

Public Property Get Width() As Long
  Width = mlngWidth
  
End Property

Public Property Let Width(ByVal plngNewValue As Long)
  mlngWidth = plngNewValue
  
End Property

Public Property Get StartOfColumn() As Boolean
  StartOfColumn = mbStartOfColumn
End Property

Public Property Let StartOfColumn(ByVal pbNewValue As Boolean)
  mbStartOfColumn = pbNewValue
End Property

Public Property Get Use1000Separator() As Boolean
  Use1000Separator = mbUse1000Separator
End Property

Public Property Let Use1000Separator(ByVal pbNewValue As Boolean)
  mbUse1000Separator = pbNewValue
End Property

Public Property Get BlankIfZero() As Boolean
  BlankIfZero = mbBlankIfZero
End Property

Public Property Let BlankIfZero(ByVal pbNewValue As Boolean)
  mbBlankIfZero = pbNewValue
End Property
