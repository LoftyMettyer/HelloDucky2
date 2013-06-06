VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_fUpdatePrivilege As Boolean
Private mvar_fSelectPrivilege As Boolean
Private mvar_sName As String
Private mvar_lngColumnID As Long

Public Property Let Name(ByVal psData As String)
  mvar_sName = psData
  
End Property

Public Property Let ColumnID(ByVal plngData As Long)
  mvar_lngColumnID = plngData
  
End Property


Public Property Get Name() As String
  Name = mvar_sName

End Property

Public Property Get ColumnID() As Long
  ColumnID = mvar_lngColumnID

End Property


Public Property Let SelectPrivilege(ByVal pfData As Boolean)
  mvar_fSelectPrivilege = pfData
    
End Property

Public Property Get SelectPrivilege() As Boolean
  SelectPrivilege = mvar_fSelectPrivilege
    
End Property

Public Property Let UpdatePrivilege(ByVal pfData As Boolean)
  mvar_fUpdatePrivilege = pfData

End Property

Public Property Get UpdatePrivilege() As Boolean
  UpdatePrivilege = mvar_fUpdatePrivilege

End Property

