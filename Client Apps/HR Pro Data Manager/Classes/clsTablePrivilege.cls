VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTablePrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngTableID As Long
Private msTableName As String
Private miTableType As TableTypes
Private mlngDfltOrderID As Long
Private mlngRecDescID As Long

Private mlngViewID As Long
Private msViewName As String

Private mfIsTable As Boolean

Private msRealSource As String

Private mfSelect As Boolean
Private mfUpdate As Boolean
Private mfInsert As Boolean
Private mfDelete As Boolean
Private mfHideFromMenu As Boolean

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

Public Property Get AllowInsert() As Boolean
  AllowInsert = mfInsert

End Property

Public Property Let AllowInsert(ByVal pfInsert As Boolean)
  mfInsert = pfInsert

End Property

Public Property Get AllowDelete() As Boolean
  AllowDelete = mfDelete

End Property

Public Property Get TableID() As Long
  TableID = mlngTableID

End Property
Public Property Get ViewID() As Long
  ViewID = mlngViewID

End Property



Public Property Get RealSource() As String
  RealSource = msRealSource

End Property


Public Property Get IsTable() As Boolean
  IsTable = mfIsTable
  
End Property

Public Property Let AllowDelete(ByVal pfDelete As Boolean)
  mfDelete = pfDelete

End Property
Public Property Let TableID(ByVal plngNewValue As Long)
  mlngTableID = plngNewValue

End Property
Public Property Let ViewID(ByVal plngNewValue As Long)
  mlngViewID = plngNewValue

End Property

Public Property Let RealSource(ByVal psRealSource As String)
  msRealSource = psRealSource

End Property

Public Property Let IsTable(ByVal pfIsTable As Boolean)
  mfIsTable = pfIsTable

End Property



Public Property Get TableName() As String
  TableName = msTableName
  
End Property
Public Property Get ViewName() As String
  ViewName = msViewName
  
End Property

Public Property Let TableName(ByVal psNewValue As String)
  msTableName = psNewValue
  
End Property
Public Property Let ViewName(ByVal psNewValue As String)
  msViewName = psNewValue
  
End Property


Public Property Get TableType() As TableTypes
  TableType = miTableType
  
End Property

Public Property Let TableType(ByVal piNewValue As TableTypes)
  miTableType = piNewValue
  
End Property

Public Property Get DefaultOrderID() As Long
  DefaultOrderID = mlngDfltOrderID
  
End Property

Public Property Let DefaultOrderID(ByVal plngNewValue As Long)
  mlngDfltOrderID = plngNewValue
  
End Property

Public Property Get RecordDescriptionID() As Long
  RecordDescriptionID = mlngRecDescID
  
End Property

Public Property Let RecordDescriptionID(ByVal plngNewValue As Long)
  mlngRecDescID = plngNewValue
  
End Property

Public Property Get HideFromMenu() As Boolean
   HideFromMenu = mfHideFromMenu
End Property

Public Property Let HideFromMenu(ByVal pfNewValue As Boolean)
   mfHideFromMenu = pfNewValue
End Property
