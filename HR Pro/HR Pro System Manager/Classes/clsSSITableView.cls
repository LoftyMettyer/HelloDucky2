VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSITableView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngTableID As Long
Private mlngViewID As Long
Private msTableViewName As String

Public Property Get TableID() As Long
 TableID = mlngTableID
End Property
Public Property Let TableID(NEW_VALUE As Long)
  mlngTableID = NEW_VALUE
End Property

Public Property Get ViewID() As Long
 ViewID = mlngViewID
End Property
Public Property Let ViewID(NEW_VALUE As Long)
  mlngViewID = NEW_VALUE
End Property

Public Property Get TableViewName() As String
 TableViewName = msTableViewName
End Property
Public Property Let TableViewName(NEW_VALUE As String)
  msTableViewName = NEW_VALUE
End Property


