VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrossTabRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private datData As clsDataAccess

Private mlngCrossTabID As Long

Private Sub Class_Initialize()
  Set datData = New HrPro.clsDataAccess
End Sub

Public Property Get CrossTabID() As Long
  CrossTabID = mlngCrossTabID
End Property

Public Property Let CrossTabID(ByVal lngNewID As Long)
  mlngCrossTabID = lngNewID
End Property

