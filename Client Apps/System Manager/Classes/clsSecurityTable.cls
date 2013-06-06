VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_sName As String
Private mvar_fInsertPrivilege As Boolean
Private mvar_iSelectPrivilege As ColumnPrivilegeStates
Private mvar_iUpdatePrivilege As ColumnPrivilegeStates
Private mvar_fDeletePrivilege As Boolean
Private mvar_objColumns As clsSecurityColumns
Private mvar_iTableType As Integer
Private mvar_iParentJoinType As Integer

Public Sub GrantAll()
  ' Grant all permissions on all columns in this table/view.
  Dim iLoop As Integer
  Dim objColumn As clsSecurityColumns
  
  For iLoop = 1 To Columns.Count
    ' Grant permission to the column.
    Set objColumn = Columns
    With objColumn.Item(iLoop)
      .SelectPrivilege = True
      .UpdatePrivilege = True
    End With
    Set objColumn = Nothing
  Next iLoop
  
End Sub

Public Property Let DeletePrivilege(ByVal pfData As Boolean)
  mvar_fDeletePrivilege = pfData
End Property

Public Property Get DeletePrivilege() As Boolean
  DeletePrivilege = mvar_fDeletePrivilege
End Property

Public Property Let UpdatePrivilege(ByVal piData As ColumnPrivilegeStates)
  mvar_iUpdatePrivilege = piData
End Property

Public Property Get UpdatePrivilege() As ColumnPrivilegeStates
  Dim objColumn As clsSecurityColumn
  Dim fAllGranted As Boolean
  Dim fNoneGranted As Boolean
  
  fAllGranted = True
  fNoneGranted = True
  
  ' Check the select privilege of the table/view's columns.
  For Each objColumn In Columns.Collection
    If objColumn.UpdatePrivilege Then
      fNoneGranted = False
    Else
      fAllGranted = False
    End If
  Next
  
  If fAllGranted Then
    mvar_iUpdatePrivilege = giPRIVILEGES_ALLGRANTED
  ElseIf fNoneGranted Then
    mvar_iUpdatePrivilege = giPRIVILEGES_NONEGRANTED
  Else
    mvar_iUpdatePrivilege = giPRIVILEGES_SOMEGRANTED
  End If
  
  UpdatePrivilege = mvar_iUpdatePrivilege
  
End Property

Public Property Let SelectPrivilege(ByVal piData As ColumnPrivilegeStates)
  mvar_iSelectPrivilege = piData
  
End Property

Public Property Get SelectPrivilege() As ColumnPrivilegeStates
  Dim objColumn As clsSecurityColumn
  Dim fAllGranted As Boolean
  Dim fNoneGranted As Boolean
  
  fAllGranted = True
  fNoneGranted = True
  
  ' Check the select privilege of the table/view's columns.
  For Each objColumn In Columns.Collection
    If objColumn.SelectPrivilege Then
      fNoneGranted = False
    Else
      fAllGranted = False
    End If
  Next
  
  If fAllGranted Then
    mvar_iSelectPrivilege = giPRIVILEGES_ALLGRANTED
  ElseIf fNoneGranted Then
    mvar_iSelectPrivilege = giPRIVILEGES_NONEGRANTED
  Else
    mvar_iSelectPrivilege = giPRIVILEGES_SOMEGRANTED
  End If

  SelectPrivilege = mvar_iSelectPrivilege
  
End Property

Public Property Let InsertPrivilege(ByVal pfData As Boolean)
  mvar_fInsertPrivilege = pfData
  
End Property

Public Property Get InsertPrivilege() As Boolean
  InsertPrivilege = mvar_fInsertPrivilege
  
End Property

Public Property Let Name(ByVal psData As String)
  mvar_sName = psData
  
End Property

Public Property Get Name() As String
  Name = mvar_sName
End Property

Public Property Get Columns() As clsSecurityColumns
  If mvar_objColumns Is Nothing Then
    Set mvar_objColumns = New clsSecurityColumns
  End If

  Set Columns = mvar_objColumns
  
End Property

Public Property Set Columns(pobjData As clsSecurityColumns)
  Set mvar_objColumns = pobjData
End Property

Private Sub Class_Initialize()
  ' Initialise the permissions to revoked.
  mvar_fInsertPrivilege = False
  mvar_iSelectPrivilege = giPRIVILEGES_NONEGRANTED
  mvar_iUpdatePrivilege = giPRIVILEGES_NONEGRANTED
  mvar_fDeletePrivilege = False
  
End Sub

Private Sub Class_Terminate()
  ' Disassociate object variables.
  If Not mvar_objColumns Is Nothing Then
    Set mvar_objColumns = Nothing
  End If
End Sub

Public Property Get TableType() As Integer
  TableType = mvar_iTableType
End Property

Public Property Get ParentJoinType() As Integer
  ParentJoinType = mvar_iParentJoinType
End Property

Public Property Let TableType(ByVal piNewValue As Integer)
  mvar_iTableType = piNewValue
End Property

Public Property Let ParentJoinType(ByVal piNewValue As Integer)
  mvar_iParentJoinType = piNewValue
End Property

