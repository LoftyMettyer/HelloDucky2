VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordProfileColDtls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Add(psKey As String, _
  psType As String, _
  plngID As Long, _
  psHeading As String, _
  plngSize As Long, _
  piDecPlaces As Integer, _
  piDataType As SQLDataType, _
  plngTableID As Long, _
  psColumnName As String, _
  Optional piSequence As Integer, _
  Optional psBeforeKey As String) As clsRecordProfileColDtl
                    
  ' Add a new object to the collection
  ' Called when a user adds a column/heading/separator to a record profile definition
  Dim objNewMember As New clsRecordProfileColDtl

  With objNewMember
    .ColType = psType
    .ID = plngID
    .Heading = psHeading
    .Size = plngSize
    .DecPlaces = piDecPlaces
    .TableID = plngTableID
    .ColumnName = psColumnName
    .Sequence = piSequence
    .DataType = piDataType
    .Key = psKey
  End With
  
  If Len(psBeforeKey) > 0 Then
    mCol.Add objNewMember, psKey, psBeforeKey
  Else
    mCol.Add objNewMember, psKey
  End If
  
  Set Add = objNewMember

  Set objNewMember = Nothing
                    
End Function


Public Property Get Count() As Long
  ' Provide number of objects in the collection
  Count = mCol.Count
    
End Property


Public Property Get Item(psKey As String) As clsRecordProfileColDtl
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


