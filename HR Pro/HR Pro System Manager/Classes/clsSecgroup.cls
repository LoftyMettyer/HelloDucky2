VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrGroupName As String
Private mblnAllow As Boolean


Public Property Get GroupName() As String
  GroupName = mstrName
End Property

Public Property Let GroupName(val As String)
  mstrGroupName = val
End Property

Public Property Get Allow() As Boolean
  Allow = mblnAllow
End Property

Public Property Let Allow(val As Boolean)
  mblnAllow = val
End Property

