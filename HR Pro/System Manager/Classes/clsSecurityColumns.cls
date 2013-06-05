VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable to hold collection.
Private mvar_colColumns As Collection

Public Function Add(psKey As String) As clsSecurityColumn
  Dim objNewMember As clsSecurityColumn
  
  ' Create a new object.
  Set objNewMember = New clsSecurityColumn

  mvar_colColumns.Add objNewMember, psKey
  
  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function

Public Property Get Item(ByRef pvIndexKey As Variant) As clsSecurityColumn
Attribute Item.VB_UserMemId = 0
  Set Item = mvar_colColumns(pvIndexKey)
  
End Property

Public Function IsValid(ByRef pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given security column exists in the collection.
  Dim Item As clsSecurityColumn
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colColumns(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function



Public Property Get Count() As Long
  Count = mvar_colColumns.Count
  
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colColumns.Remove pvIndexKey
  
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = 0
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colColumns.[_NewEnum]
  
End Property

Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colColumns = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroys the collection when this class is terminated.
  If Not mvar_colColumns Is Nothing Then
    Set mvar_colColumns = Nothing
  End If
  
End Sub



Public Property Get Collection() As Collection
  Set Collection = mvar_colColumns

End Property

