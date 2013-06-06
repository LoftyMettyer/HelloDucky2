VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordProfileTabDtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Add(plngTableID As Long, _
  psTableName As String, _
  plngRecDescID As Long, _
  plngFilterID As Long, _
  plngPicklistID As Long, _
  plngOrderID As Long, _
  plngMaxRecords As Long, _
  piOrientation As OrientationTypes, _
  pfPageBreak As Boolean, _
  piSequence As Integer) As clsRecordProfileTabDtl

  ' Add a new object to the collection
  Dim objNewMember As New clsRecordProfileTabDtl

  With objNewMember
    .FilterID = plngFilterID
    .Generation = 0
    .HasChildren = False
    .LastRecordID = 0
    .MaxRecords = plngMaxRecords
    .OrderID = plngOrderID
    .Orientation = piOrientation
    .PageBreak = pfPageBreak
    .PicklistID = plngPicklistID
    .RecordDescriptionID = plngRecDescID
    .RelatedTableID = 0
    .RelatedTableIDPosition = 0
    .Relationship = ""
    .Sequence = piSequence
    .TableID = plngTableID
    .TableName = psTableName
    .IDPosition = 0
    .HasHeadings = False
    .PhotoCount = 0
  End With
  
  mCol.Add objNewMember, CStr(plngTableID)

  Set Add = objNewMember

  Set objNewMember = Nothing
                    
End Function


Public Property Get Count() As Long
  ' Provide number of objects in the collection
  Count = mCol.Count
    
End Property







Public Property Get ItemByPosition(piPosition As Integer) As clsRecordProfileTabDtl
  ' Provide a reference to a specific item in the collection
  On Error Resume Next

  Set ItemByPosition = mCol.Item(piPosition)

End Property
Public Property Get BaseTable() As clsRecordProfileTabDtl
  ' Provide a reference to a specific item in the collection
  On Error Resume Next

  Set BaseTable = mCol.Item(1)

End Property

Public Property Get Item(psKey As String) As clsRecordProfileTabDtl
  ' Provide a reference to a specific item in the collection
  On Error Resume Next

  Set Item = mCol(psKey)

End Property




Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
  Set NewEnum = mCol.[_NewEnum]

End Property



Public Sub Remove(psKey As String)

  ' Remove a specific object from the collection
  ' Called when a user removes a column/header/separator from a record profile definition

  mCol.Remove psKey

End Sub



Public Sub RemoveAll()
  ' Remove all objects from the collection
  ' Called when a user removes all columns from a record profile definition

  Do While mCol.Count > 1
    mCol.Remove mCol.Count
  Loop

End Sub



Private Sub Class_Initialize()
  Set mCol = New Collection

End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing

End Sub


