VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable to hold collection.
Private mvar_colGroups As Collection

Public Function Add(psName As String) As clsSecurityGroup
  ' Add the given user group to the collection of user groups.
  Dim objNewMember As clsSecurityGroup
  Dim objTables As clsSecurityTables
  Dim objViews As clsSecurityTables

  ' Instantiate new collections for the group definition.
  Set objTables = New clsSecurityTables
  Set objViews = New clsSecurityTables

  ' Instantiate a new security group.
  Set objNewMember = New clsSecurityGroup

  With objNewMember
    ' Set the properties passed into the method.
    Set .Tables = objTables
    Set .Views = objViews
    
    .Name = psName
    .SystemManager = False
    .SecurityManager = False
  End With

   mvar_colGroups.Add objNewMember, psName

  ' Return the new group object.
  Set Add = objNewMember
  Set objNewMember = Nothing
  Set objTables = Nothing
  Set objViews = Nothing

End Function

Public Property Get Item(ByRef pvIndexKey As Variant) As clsSecurityGroup
Attribute Item.VB_UserMemId = 0
  Set Item = mvar_colGroups(pvIndexKey)
  
End Property

Public Property Get Count() As Long
  Count = mvar_colGroups.Count
  
End Property

Public Property Get Collection() As Collection
  Set Collection = mvar_colGroups

End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colGroups.Remove pvIndexKey
  
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = 0
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colGroups.[_NewEnum]
  
End Property

Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mvar_colGroups = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroy collection when this class is terminated.
  If Not mvar_colGroups Is Nothing Then
    Set mvar_colGroups = Nothing
  End If
  
End Sub


