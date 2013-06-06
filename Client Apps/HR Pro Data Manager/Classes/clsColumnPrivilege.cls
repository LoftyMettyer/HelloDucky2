VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumnPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msColumnName As String
Private mfSelect As Boolean
Private mfUpdate As Boolean
Private miColumnType As Integer
Private miDataType As Integer
Private mlngColumnID As Long
Private mfUniqueCheck As Boolean
Private mlDisplaySize As Long
Private mlSize As Long
Private mlDecimals As Long
Private mbUseThousandSeparator As Boolean
Private miOLEType As DataMgr.OLEType

Public Property Get AllowSelect() As Boolean
  AllowSelect = mfSelect

End Property

Public Property Let AllowSelect(ByVal pfSelect As Boolean)
  mfSelect = pfSelect

End Property

Public Property Get AllowUpdate() As Boolean
   AllowUpdate = mfUpdate

End Property

Public Property Let AllowUpdate(ByVal pfUpdate As Boolean)
  mfUpdate = pfUpdate

End Property

Public Property Get ColumnName() As String
  ColumnName = msColumnName
  
End Property

Public Property Let ColumnName(ByVal psNewValue As String)
  msColumnName = psNewValue

End Property


Public Property Get ColumnType() As Integer
  ColumnType = miColumnType
  
End Property

Public Property Let ColumnType(ByVal piNewValue As Integer)
  miColumnType = piNewValue
  
End Property

Public Property Get DataType() As Integer
  DataType = miDataType
  
End Property

Public Property Let DataType(ByVal piNewValue As Integer)
  miDataType = piNewValue
  
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
  
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  mlngColumnID = plngNewValue
  
End Property

Public Property Get UniqueCheck() As Boolean
  UniqueCheck = mfUniqueCheck
  
End Property

Public Property Let UniqueCheck(ByVal pfNewValue As Boolean)
  mfUniqueCheck = pfNewValue
  
End Property

Public Property Get DisplaySize() As Long
  DisplaySize = mlDisplaySize
End Property

Public Property Let DisplaySize(ByVal lngNewValue As Long)
  mlDisplaySize = lngNewValue
End Property

Public Property Get Size() As Long
  Size = mlSize
End Property

Public Property Let Size(ByVal lngNewValue As Long)
  mlSize = lngNewValue
End Property

Public Property Get Decimals() As Long
  Decimals = mlDecimals
End Property

Public Property Let Decimals(ByVal lngNewValue As Long)
  mlDecimals = lngNewValue
End Property

Public Property Get UseThousandSeparator() As Boolean
  UseThousandSeparator = mbUseThousandSeparator
End Property

Public Property Let UseThousandSeparator(ByVal pbNewValue As Boolean)
  mbUseThousandSeparator = pbNewValue
End Property

Public Property Get OLEType() As DataMgr.OLEType
  OLEType = miOLEType
End Property

Public Property Let OLEType(ByVal piNewValue As DataMgr.OLEType)
  miOLEType = piNewValue
End Property
