VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSITableViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable to hold collection.
Private mvar_colTableViews As Collection

Public Function Add(plngTableID As Long, _
                    plngViewID As Long, _
                    psTableViewName As String) As clsSSITableView
                    
  Dim objNewMember As clsSSITableView
  Dim sKey As String
  
  Set objNewMember = New clsSSITableView

  sKey = plngTableID & "_" & plngViewID

  ' Set the properties of the new table object.
  With objNewMember
    .TableID = plngTableID
    .ViewID = plngViewID
    .TableViewName = psTableViewName
  End With
  
  mvar_colTableViews.Add objNewMember, sKey

  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function

Public Property Get Item(ByRef pvIndexKey As Variant) As clsSSITableView
  
  On Error GoTo err_item
  
  Set Item = mvar_colTableViews(pvIndexKey)
  
Exit Property

err_item:
  Set Item = Nothing
  
End Property

Public Property Get Collection() As Collection
  Set Collection = mvar_colTableViews
End Property

Public Property Get Count() As Long
  Count = mvar_colTableViews.Count
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colTableViews.Remove pvIndexKey
End Sub

Public Sub Clear()
  Do While mvar_colTableViews.Count > 0
    mvar_colTableViews.Remove 1
  Loop
End Sub

Public Property Get NewEnum() As IUnknown
  Set NewEnum = mvar_colTableViews.[_NewEnum]
End Property

Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colTableViews = New Collection
End Sub

Private Sub Class_Terminate()
  ' Destroy the collection when this class is terminated.
  If Not mvar_colTableViews Is Nothing Then
    Set mvar_colTableViews = Nothing
  End If
End Sub


