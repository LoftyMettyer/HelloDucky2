VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTablePrivileges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Property Get Collection() As Collection
  Set Collection = mCol
  
End Property

Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
  
End Property

Public Function Add(psTableName As String, plngTableID As Long, piTableType As Integer, _
  plngDfltOrderID As Long, plngRecDescID As Long, pfIsTable As Boolean, _
  plngViewID As Long, psViewName As String, pbIsRemoteView As Boolean) As CTablePrivilege
  ' Add a new member to the collection of table privileges.
  Dim sKey As String
  Dim objNewMember As CTablePrivilege
  
  Set objNewMember = New CTablePrivilege
  
  ' Initialise the privileges.
  With objNewMember
    .TableID = plngTableID
    .TableName = psTableName
    .TableType = piTableType
    .DefaultOrderID = plngDfltOrderID
    .RecordDescriptionID = plngRecDescID
    
    .IsTable = pfIsTable
    
    .ViewID = plngViewID
    .ViewName = psViewName
    
    If (Not pfIsTable) Then
      sKey = psViewName
    Else
      sKey = psTableName
    End If
       
    .AllowSelect = False
    .AllowUpdate = False
    .AllowDelete = False
    .AllowInsert = False
    
    .IsRemoteView = pbIsRemoteView
  End With
  
  mCol.Add objNewMember, sKey
    
  Set Add = objNewMember
  Set objNewMember = Nothing
                    
End Function

Public Property Get Item(ByRef vntIndexKey As Variant) As CTablePrivilege
  Set Item = mCol(vntIndexKey)
    
End Property

Public Property Get Count() As Long
    
    Count = mCol.Count
    
End Property

Public Sub Remove(vntIndexKey As Variant)

    mCol.Remove vntIndexKey
    
End Sub

Private Sub Class_Initialize()
  Set mCol = New Collection
    
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
    
End Sub





Public Function FindRealSource(psRealSource As String) As CTablePrivilege
  ' Return the table/view privilege object with the given real source.
  Dim objTable As CTablePrivilege
  Dim objRequiredTable As CTablePrivilege
  
  For Each objTable In mCol
    If objTable.RealSource = psRealSource Then
      Set objRequiredTable = objTable
      Exit For
    End If
  Next objTable
  Set objTable = Nothing
  
  Set FindRealSource = objRequiredTable
  
End Function
Public Function FindTableID(plngTableID As Long) As CTablePrivilege
  ' Return the table/view privilege object with the given table ID.
  Dim objTable As CTablePrivilege
  Dim objRequiredTable As CTablePrivilege
  
  For Each objTable In mCol
    ' JPD 6/9/00 This function has been modified to ensure that the object returned is for the
    ' given table, and not just a view on the given table.
'    If objTable.TableID = plngTableID Then
    If (objTable.TableID = plngTableID) And _
      (objTable.IsTable) Then
      Set objRequiredTable = objTable
      Exit For
    End If
  Next objTable
  Set objTable = Nothing
  
  Set FindTableID = objRequiredTable
  
End Function

Public Function IsValid(ByRef pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given key exists in the collection.
  Dim Item As CTablePrivilege
  
  On Error GoTo err_IsValid
  
  Set Item = mCol(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function


Public Function FindViewID(plngViewID As Long) As CTablePrivilege
  ' Return the table/view privilege object with the given table ID.
  Dim objView As CTablePrivilege
  Dim objRequiredView As CTablePrivilege
  
  For Each objView In mCol
    If objView.ViewID = plngViewID Then
      Set objRequiredView = objView
      Exit For
    End If
  Next objView
  Set objView = Nothing
  
  Set FindViewID = objRequiredView
  
End Function

