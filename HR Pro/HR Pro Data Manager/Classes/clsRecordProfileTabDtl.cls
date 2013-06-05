VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordProfileTabDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngTableID As Long
Private msTableName As String
Private miIDPosition As Integer
Private mlngRecDescID As Long
Private mlngFilterID As Long
Private mlngPicklistID As Long
Private mlngOrderID As Long

Private mlngRelatedTableID As Long
Private miRelatedTableIDPosition As Integer
Private msRelationship As String
Private miGeneration As Integer
Private mfHasChildren As Boolean
Private mlngMaxRecords As Long
Private miOrientation As OrientationTypes
Private mfPageBreak As Boolean
Private miSequence As Integer

Private mlngLastRecordID As Long
Private mfHasHeadings As Boolean
Private miPhotoCount As Integer
Private mlngGridIndex As Long

Private mcolColumns As clsRecordProfileColDtls

Public Property Get TableID() As Long
  TableID = mlngTableID
  
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  mlngTableID = plngNewValue
  
End Property

Public Property Get TableName() As String
  TableName = msTableName
  
End Property

Public Property Let TableName(ByVal psNewValue As String)
  msTableName = psNewValue
  
End Property

Public Property Get RecordDescriptionID() As Long
  RecordDescriptionID = mlngRecDescID
  
End Property

Public Property Let RecordDescriptionID(ByVal plngNewValue As Long)
  mlngRecDescID = plngNewValue
  
End Property

Public Property Get FilterID() As Long
  FilterID = mlngFilterID
  
End Property

Public Property Let FilterID(ByVal plngNewValue As Long)
  mlngFilterID = plngNewValue
  
End Property

Public Property Get OrderID() As Long
  OrderID = mlngOrderID
  
End Property

Public Property Let OrderID(ByVal plngNewValue As Long)
  mlngOrderID = plngNewValue
  
End Property

Public Property Get RelatedTableID() As Long
  RelatedTableID = mlngRelatedTableID
  
End Property

Public Property Let RelatedTableID(ByVal plngNewValue As Long)
  mlngRelatedTableID = plngNewValue
  
End Property

Public Property Get RelatedTableIDPosition() As Integer
  RelatedTableIDPosition = miRelatedTableIDPosition
  
End Property

Public Property Let RelatedTableIDPosition(ByVal piNewValue As Integer)
  miRelatedTableIDPosition = piNewValue
  
End Property

Public Property Get Relationship() As String
  Relationship = msRelationship

End Property

Public Property Let Relationship(ByVal psNewValue As String)
  msRelationship = psNewValue
  
End Property

Public Property Get Generation() As Integer
  Generation = miGeneration
  
End Property

Public Property Let Generation(ByVal piNewValue As Integer)
  miGeneration = piNewValue
  
End Property

Public Property Get HasChildren() As Boolean
  HasChildren = mfHasChildren
  
End Property

Public Property Let HasChildren(ByVal pfNewValue As Boolean)
  mfHasChildren = pfNewValue
  
End Property

Public Property Get MaxRecords() As Long
  MaxRecords = mlngMaxRecords
  
End Property

Public Property Let MaxRecords(ByVal plngNewValue As Long)
  mlngMaxRecords = plngNewValue
  
End Property

Public Property Get Orientation() As OrientationTypes
  Orientation = miOrientation
  
End Property

Public Property Let Orientation(ByVal piNewValue As OrientationTypes)
  miOrientation = piNewValue
  
End Property

Public Property Get PageBreak() As Boolean
  PageBreak = mfPageBreak
  
End Property

Public Property Let PageBreak(ByVal pfNewValue As Boolean)
  mfPageBreak = pfNewValue
  
End Property

Public Property Get Sequence() As Integer
  Sequence = miSequence
  
End Property

Public Property Let Sequence(ByVal piNewValue As Integer)
  miSequence = piNewValue
  
End Property

Public Property Get LastRecordID() As Long
  LastRecordID = mlngLastRecordID
  
End Property

Public Property Let LastRecordID(ByVal plngNewValue As Long)
  mlngLastRecordID = plngNewValue
  
End Property

Private Sub Class_Initialize()
  Set mcolColumns = New clsRecordProfileColDtls

End Sub


Private Sub Class_Terminate()
  Set mcolColumns = Nothing

End Sub



Public Property Get PicklistID() As Long
  PicklistID = mlngPicklistID
  
End Property

Public Property Let PicklistID(ByVal plngNewValue As Long)
  mlngPicklistID = plngNewValue
  
End Property

Public Property Get Columns() As clsRecordProfileColDtls
  Set Columns = mcolColumns

End Property

Public Property Set Columns(pcolNewValue As clsRecordProfileColDtls)
  Set mcolColumns = pcolNewValue

End Property


Public Property Get IDPosition() As Integer
  IDPosition = miIDPosition
  
End Property

Public Property Let IDPosition(ByVal piNewValue As Integer)
  miIDPosition = piNewValue
  
End Property

Public Property Get HasHeadings() As Boolean
  HasHeadings = mfHasHeadings
  
End Property

Public Property Let HasHeadings(ByVal pfNewValue As Boolean)
  mfHasHeadings = pfNewValue
  
End Property

Public Property Get PhotoCount() As Integer
  PhotoCount = miPhotoCount
  
End Property

Public Property Let PhotoCount(ByVal piNewValue As Integer)
  miPhotoCount = piNewValue
  
End Property



Public Property Get GridIndex() As Long
  GridIndex = mlngGridIndex
  
End Property

Public Property Let GridIndex(ByVal plngNewValue As Long)
  mlngGridIndex = plngNewValue
  
End Property
