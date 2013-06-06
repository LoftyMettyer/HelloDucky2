VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParentValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Add(sColumn As String, sTable As String, bCanSelect As Boolean, lIDColumn As Long, _
                sTabColName As String, psDefaultValue As String) As CParentValue
                    
    Dim obNewMember As CParentValue
    
    Set obNewMember = New CParentValue
    
    obNewMember.ColumnName = sColumn
    obNewMember.TableName = sTable
    obNewMember.CanSelect = bCanSelect
    obNewMember.IDColumn = lIDColumn
    obNewMember.DefaultValue = psDefaultValue
    
    mCol.Add obNewMember, sTabColName
    
    Set Add = obNewMember
    Set obNewMember = Nothing
                    
End Function

Public Property Get Item(sTabColName As String) As CParentValue
    
    On Error Resume Next
    
    Set Item = mCol(sTabColName)
    
End Property

Public Property Get Count() As Long
    
    Count = mCol.Count
    
End Property

Public Sub Remove(sTabColName As String)

    mCol.Remove sTabColName
    
End Sub

Private Sub Class_Initialize()
    
    Set mCol = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set mCol = Nothing
    
End Sub

Public Property Get NewEnum() As IUnknown
    
    Set NewEnum = mCol.[_NewEnum]
    
End Property





