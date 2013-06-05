VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"PropertyItem"
Attribute VB_Ext_KEY = "Member0" ,"PropertyItem"
Option Explicit

'Local variable to hold collection
Private mCol As Collection

Public Function Add(ByRef sName As String, Optional ByVal vValue As Variant) As Property
  Dim strKey As String
  
  'Create a new object
  Dim objNewMember As Property
  
  If LenB(Trim$(sName)) <> 0 Then
    Set objNewMember = New Property

    'Set the properties passed into the method
    objNewMember.Key = Trim$(sName)
    objNewMember.Name = sName
    If Not IsEmpty(vValue) Then objNewMember.Value = vValue
    
    mCol.Add objNewMember, objNewMember.Key

    'Return the object created
    Set Add = objNewMember
  End If
  Set objNewMember = Nothing

End Function

Public Property Get Item(ByRef vntIndexKey As Variant) As Property
Attribute Item.VB_UserMemId = 0
  On Error GoTo ErrorTrap
  
  Set Item = mCol(vntIndexKey)
  
  Exit Property
  
ErrorTrap:
  Set Item = Nothing
  Err = False
  
End Property

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  'Create the collection
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  'Destroy the collection
  Set mCol = Nothing
End Sub

