VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SecurityColumn"
Attribute VB_Ext_KEY = "Member0" ,"SecurityColumn"
Option Explicit

' Local variable to hold collection.
Private mvar_colColumns As Collection
Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
'
  Dim fOK As Boolean
  Dim objColumn As SecurityColumn
'
  fOK = True

  For Each objColumn In mvar_colColumns
    If Not objColumn.Changed Then objColumn.PrintSecurity piLevel
    'Exit For
  Next objColumn
  
TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function Add(psName As String) As SecurityColumn
  Dim fAdded As Boolean
  Dim objColumn As SecurityColumn
  Dim objNewMember As SecurityColumn
  
  ' Create a new object.
  Set objNewMember = New SecurityColumn

  ' Add the new column into the collection, positioning it
  ' so that the collection is ordered alphabetically.
  fAdded = False
  For Each objColumn In mvar_colColumns
    If UCase(psName) < UCase(objColumn.Name) Then
      mvar_colColumns.Add objNewMember, psName, objColumn.Name
      fAdded = True
      Exit For
    End If
  Next objColumn
  Set objColumn = Nothing
  
  If Not fAdded Then
    mvar_colColumns.Add objNewMember, psName
  End If
  
  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function



Public Function IsValid(ByRef pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given security column exists in the collection.
  Dim Item As SecurityColumn
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colColumns(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function


Public Property Get Item(ByRef pvIndexKey As Variant) As SecurityColumn
Attribute Item.VB_UserMemId = 0
  Set Item = mvar_colColumns(pvIndexKey)
  
End Property

Public Property Get Count() As Long
  Count = mvar_colColumns.Count
  
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colColumns.Remove pvIndexKey
  
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colColumns.[_NewEnum]
  
End Property

Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colColumns = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroys the collection when this class is terminated.
  Set mvar_colColumns = Nothing
  
End Sub


Public Function Clone() As SecurityColumns

  Dim objSrcColumn As SecurityColumn
  Dim objDestColumn As SecurityColumn

  Set Clone = New SecurityColumns

  For Each objSrcColumn In mvar_colColumns
    
    Set objDestColumn = Clone.Add(objSrcColumn.Name)
    With objDestColumn
      .Changed = True
      .Name = objSrcColumn.Name
      .SelectPrivilege = objSrcColumn.SelectPrivilege
      .UpdatePrivilege = objSrcColumn.UpdatePrivilege
    End With
  
  Next objSrcColumn
  
  Set objSrcColumn = Nothing
  
End Function
