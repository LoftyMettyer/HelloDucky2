VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoryScreens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Private msTag As String

Private Sub Class_Initialize()
  Set mCol = New Collection

End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing

End Sub


Public Function Add(plngHistoryScreenID As Long, psHistoryScreenName As String, plngPictureID As Long, _
  plngTableID As Long, plngViewID As Long, psViewName As String, psTableName As String) As clsHistoryScreen
                    
  Dim obNewMember As clsHistoryScreen
    
  Set obNewMember = New clsHistoryScreen
  
  With obNewMember
    .HistoryScreenID = plngHistoryScreenID
    .HistoryScreenName = psHistoryScreenName
    .PictureID = plngPictureID
    .TableID = plngTableID
    .ViewID = plngViewID
    .ViewName = psViewName
    .TableName = psTableName
  End With
    
'  If Len(sKey) = 0 Then
    mCol.Add obNewMember
'  Else
'    mCol.Add obNewMember, sKey
'  End If
    
  Set Add = obNewMember
  Set obNewMember = Nothing
                    
End Function



Public Property Get Tag() As String
  ' Return the object's tag.
  Tag = msTag
  
End Property

Public Property Let Tag(ByVal psNewValue As String)
  ' Set the object's tag property.
  msTag = psNewValue
  
End Property

Public Property Get Count() As Long
  Count = mCol.Count
    
End Property

Public Property Get Item(vntIndexKey As Variant) As clsHistoryScreen
  Set Item = mCol(vntIndexKey)
    
End Property

Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
    
End Sub

