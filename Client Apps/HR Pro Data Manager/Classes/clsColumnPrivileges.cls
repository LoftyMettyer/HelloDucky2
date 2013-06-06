VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumnPrivileges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ColumnPrivilege"
Attribute VB_Ext_KEY = "Member0" ,"ColumnPrivilege"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private msTag As String

Public Function IsValid(pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given column exists in the collection.
  Dim Item As CColumnPrivilege
  
  On Error GoTo err_IsValid
  
  Set Item = mCol(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function



Public Function Add(pfSelect As Boolean, pfUpdate As Boolean, psColumnName As String, _
  piColumnType As Integer, piDataType As Integer, plngColumnID As Long, _
  pfUniqueCheck As Boolean, plngDisplaySize As Long, plngSize As Long, plngDecimals As Long, _
  pbUseThousandSeparator As Boolean, piOLEType As DataMgr.OLEType) As CColumnPrivilege
  'create a new object
  Dim objNewMember As CColumnPrivilege
  Set objNewMember = New CColumnPrivilege

  With objNewMember
    .ColumnName = psColumnName
    .AllowSelect = pfSelect
    .AllowUpdate = pfUpdate
    .ColumnType = piColumnType
    .DataType = piDataType
    .ColumnID = plngColumnID
    .UniqueCheck = pfUniqueCheck
    .DisplaySize = plngDisplaySize
    .Size = plngSize
    .Decimals = plngDecimals
    .UseThousandSeparator = pbUseThousandSeparator
    .OLEType = piOLEType
  End With
  
  mCol.Add objNewMember, psColumnName

  'return the object created
  Set Add = objNewMember
  Set objNewMember = Nothing

End Function

Public Function FindColumnID(plngColumnID As Long) As CColumnPrivilege
  ' Return the column privilege object with the given column ID.
  Dim objColumn As CColumnPrivilege
  Dim objRequiredColumn As CColumnPrivilege
  
  For Each objColumn In mCol
    If objColumn.ColumnID = plngColumnID Then
      Set objRequiredColumn = objColumn
      Exit For
    End If
  Next objColumn
  Set objColumn = Nothing
  
  Set FindColumnID = objRequiredColumn
  
End Function

Public Property Get Item(vntIndexKey As Variant) As CColumnPrivilege
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
    
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
  Set mCol = New Collection
    
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
    
End Sub


Public Property Get Tag() As String
  ' Return the object's tag.
  Tag = msTag
  
End Property

Public Property Let Tag(ByVal psNewValue As String)
  ' Set the object's tag property.
  msTag = psNewValue
  
End Property
