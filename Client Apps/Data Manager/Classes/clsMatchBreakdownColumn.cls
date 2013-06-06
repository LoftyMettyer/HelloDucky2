VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatchBreakdownColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strColumnHeading As String
Private lngColumnID As Long
Private lngSequence As Long
Private lngSize As Long
Private lngDecimals As Long
Private blnNumeric As Boolean

Public Property Get ColumnID() As Long
  ColumnID = lngColumnID
End Property

Public Property Let ColumnID(ByVal lngNewValue As Long)
  lngColumnID = lngNewValue
End Property

Public Property Get Sequence() As Long
  Sequence = lngSequence
End Property

Public Property Let Sequence(ByVal lngNewValue As Long)
  lngSequence = lngNewValue
End Property

Public Property Get Size() As Long
  Size = lngSize
End Property

Public Property Let Size(ByVal lngNewValue As Long)
  lngSize = lngNewValue
End Property

Public Property Get Decimals() As Long
  Decimals = lngDecimals
End Property

Public Property Let Decimals(ByVal lngNewValue As Long)
  lngDecimals = lngNewValue
End Property

Public Property Get Numeric() As Boolean
  Numeric = blnNumeric
End Property

Public Property Let Numeric(ByVal blnNewValue As Boolean)
  blnNumeric = blnNewValue
End Property

Public Property Get ColumnHeading() As String
  ColumnHeading = strColumnHeading
End Property

Public Property Let ColumnHeading(strNewValue As String)
  strColumnHeading = strNewValue
End Property
