VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLookupValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection

Public Function Add(sLookupColName As String, _
  lParentHwnd As Long, _
  lChildHwnd As Long, _
  plColumnID As Long) As CLookupValue
                    
  Dim obNewMember As New CLookupValue
  
  'NHRD27062002 Fault 3779
  'The object information was being retained after the Error 457
  'I'm not sure if this is the most elegant fix but it seems to do the job.
  'TM11062004 - only remove the specific key, if it already exists.
  If IsValid(CStr(lParentHwnd)) Then
    Me.Collection.Remove CStr(lParentHwnd)
  End If
  
  obNewMember.LookupColName = sLookupColName
  obNewMember.ParentHwnd = lParentHwnd
  obNewMember.ChildHwnd = lChildHwnd
  obNewMember.ColumnID = plColumnID

  mCol.Add obNewMember, CStr(lParentHwnd)
  
  Set Add = obNewMember
  Set obNewMember = Nothing
                    
End Function

Public Function IsValid(pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given key exists in the collection.
  Dim Item As CLookupValue
  
  On Error GoTo err_IsValid
  
  Set Item = mCol(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function

Public Property Get Item(vntIndexKey As Variant) As CLookupValue

    Set Item = mCol(vntIndexKey)
    
End Property

Public Property Get Collection() As Collection

    Set Collection = mCol
    
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


