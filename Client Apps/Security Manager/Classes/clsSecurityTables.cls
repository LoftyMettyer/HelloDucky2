VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SecurityTable"
Attribute VB_Ext_KEY = "Member0" ,"SecurityTable"
Option Explicit

' Local variable to hold collection.
Private mvar_colTables As Collection

Public Function Add(pobjColumns As SecurityColumns, _
  psName As String, _
  piTableType As Integer, _
  Optional psKey As String) As SecurityTable
  
  ' Instantiate a new table object in the table collection.
  Dim objNewMember As SecurityTable
  
  Set objNewMember = New SecurityTable

  ' Set the properties of the new table object.
  With objNewMember
    Set .Columns = pobjColumns
    .Name = psName
    .TableType = piTableType
    .ParentJoinType = 0
    .ParentCount = 0
  End With
  
  If Len(psKey) = 0 Then
    mvar_colTables.Add objNewMember
  Else
    mvar_colTables.Add objNewMember, psKey
  End If

  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function

Public Property Get Item(ByRef pvIndexKey As Variant) As SecurityTable
Attribute Item.VB_UserMemId = 0
  On Error GoTo err_item
  
  Set Item = mvar_colTables(pvIndexKey)
Exit Property

err_item:
  Set Item = Nothing
  
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
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colTables.[_NewEnum]
  
End Property

Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colTables = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroy the collection when this class is terminated.
  Set mvar_colTables = Nothing
  
End Sub

Public Function IsValid(ByRef pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given security table exists in the collection.
  Dim Item As SecurityTable
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colTables(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function


Public Sub GrantAll(psUserGroup As String)
  ' Grant all permissions on all columns in all tables in the collection.
  Dim objTable As SecurityTable
  
  For Each objTable In mvar_colTables
    With objTable
      .GrantAll
      
      If Not .Changed Then
        .Changed = (Not .DeletePrivilege) Or _
          (Not .InsertPrivilege) Or _
          (.SelectPrivilege <> giPRIVILEGES_ALLGRANTED) Or _
          (.UpdatePrivilege <> giPRIVILEGES_ALLGRANTED)
      End If
      
      ' If the table/view is a top-level table/view then
      ' check if the select permission is changing from None to All.
      ' If so then flag all children of this table as changed. This is done as the
      ' permitted child views on the children need to be recalculated.
      ' Do nothing if the given table.view is a child or lookup table.
      If ((.TableType <> tabChild) And (.TableType <> tabLookup)) Then
         FlagChildrenChanged objTable, psUserGroup
      End If
        
      .DeletePrivilege = True
      .InsertPrivilege = True
      .SelectPrivilege = giPRIVILEGES_ALLGRANTED
      .UpdatePrivilege = giPRIVILEGES_ALLGRANTED
    End With
  Next objTable
  Set objTable = Nothing
  
End Sub


Public Function Clone() As SecurityTables

  Dim objSrcTable As SecurityTable
  Dim objDestTable As SecurityTable
  Dim objTempColumns As SecurityColumns

  Set Clone = New SecurityTables

  For Each objSrcTable In mvar_colTables
    
    With objSrcTable
    
      Set objTempColumns = New SecurityColumns
      Set objTempColumns = .Columns.Clone
    
      Set objDestTable = Clone.Add(objTempColumns, .Name, .TableType, .Name)
      
      objDestTable.Changed = True
      objDestTable.DeletePrivilege = .DeletePrivilege
      objDestTable.InsertPrivilege = .InsertPrivilege
      objDestTable.SelectPrivilege = .SelectPrivilege
      objDestTable.UpdatePrivilege = .UpdatePrivilege
      
      'JPD 20050810 Fault 10169
      objDestTable.TableID = .TableID
      objDestTable.ViewTableID = .ViewTableID
      
      ' JPD20020620 Fault 4045
      objDestTable.ParentJoinType = .ParentJoinType
      objDestTable.ParentCount = .ParentCount
      objDestTable.Columns_Initialised = .Columns_Initialised
      
    End With
  Next objSrcTable
  
  Set objSrcTable = Nothing
  
End Function


