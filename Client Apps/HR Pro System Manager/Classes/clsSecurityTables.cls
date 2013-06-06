VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable to hold collection.
Private mvar_colTables As Collection

Public Function Add(pobjColumns As clsSecurityColumns, _
  psName As String, _
  piTableType As Integer) As clsSecurityTable
  ' Instantiate a new table object in the table collection.
  Dim objNewMember As clsSecurityTable
  
  Set objNewMember = New clsSecurityTable

  ' Set the properties of the new table object.
  With objNewMember
    Set .Columns = pobjColumns
    .Name = psName
    .TableType = piTableType
    .ParentJoinType = 0
  End With
  
  mvar_colTables.Add objNewMember, psName

  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function

Public Property Get Item(ByRef pvIndexKey As Variant) As clsSecurityTable
Attribute Item.VB_UserMemId = 0
  On Error GoTo err_item
  
  Set Item = mvar_colTables(pvIndexKey)
Exit Property

err_item:
  Set Item = Nothing
  
End Property

Public Property Get Collection() As Collection
  Set Collection = mvar_colTables

End Property

Public Property Get Count() As Long
  Count = mvar_colTables.Count
  
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colTables.Remove pvIndexKey
  
End Sub

Public Sub Clear()
  Do While mvar_colTables.Count > 0
    mvar_colTables.Remove 1
  Loop
  
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = 0
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colTables.[_NewEnum]
  
End Property

Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colTables = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroy the collection when this class is terminated.
  If Not mvar_colTables Is Nothing Then
    Set mvar_colTables = Nothing
  End If
  
End Sub

Public Function IsValid(ByRef pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given security table exists in the collection.
  Dim Item As clsSecurityTable
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colTables(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function





Public Sub GrantAll()
  ' Grant all permissions on all columns in all tables in the collection.
  Dim objTable As clsSecurityTable
  
  For Each objTable In mvar_colTables
    With objTable
      .GrantAll
      .DeletePrivilege = True
      .InsertPrivilege = True
      .SelectPrivilege = giPRIVILEGES_ALLGRANTED
      .UpdatePrivilege = giPRIVILEGES_ALLGRANTED
    End With
  Next objTable
  Set objTable = Nothing
  
End Sub

