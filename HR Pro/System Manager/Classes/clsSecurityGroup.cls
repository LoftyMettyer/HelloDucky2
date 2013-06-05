VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_sName As String
Private mvar_fSystemManager As Boolean
Private mvar_fSecurityManager As Boolean
Private mvar_objTables As clsSecurityTables
Private mvar_objViews As clsSecurityTables










Public Property Get Views() As clsSecurityTables
  ' Create the Views collection if it doesn't exist.
  If mvar_objViews Is Nothing Then
    Set mvar_objViews = New clsSecurityTables
  End If

  Set Views = mvar_objViews
  
End Property


Public Property Set Views(pobjData As clsSecurityTables)
  Set mvar_objViews = pobjData
  
End Property





































Public Property Get Tables() As clsSecurityTables
  ' Create the Tables collection if it doesn't exist.
  If mvar_objTables Is Nothing Then
    Set mvar_objTables = New clsSecurityTables
  End If

  Set Tables = mvar_objTables
    
End Property


Public Property Set Tables(pobjData As clsSecurityTables)
  Set mvar_objTables = pobjData
    
End Property

Public Property Let Name(ByVal psData As String)
  mvar_sName = psData

End Property
Public Property Let SystemManager(ByVal pfData As Boolean)
  mvar_fSystemManager = pfData

End Property

Public Property Let SecurityManager(ByVal pfData As Boolean)
  mvar_fSecurityManager = pfData

End Property


Public Property Get SystemManager() As Boolean
  SystemManager = mvar_fSystemManager
  
End Property


Public Property Get SecurityManager() As Boolean
  SecurityManager = mvar_fSecurityManager
  
End Property



Public Property Get Name() As String
  Name = mvar_sName
  
End Property



Private Sub Class_Terminate()
  ' Disassociate object variables.
  If Not mvar_objTables Is Nothing Then
    Set mvar_objTables = Nothing
  End If
  
  If Not mvar_objViews Is Nothing Then
    Set mvar_objViews = Nothing
  End If
  
End Sub




