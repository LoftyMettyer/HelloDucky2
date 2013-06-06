VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrossTabDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rsTables As New ADODB.Recordset
Private datData As clsDataAccess



Private Sub Class_Initialize()
  Set datData = New HrPro.clsDataAccess
End Sub


Public Property Get Loading() As Boolean
  Loading = mblnLoading
End Property

Public Property Let Loading(ByVal blnNewValue As Boolean)
  mblnLoading = blnNewValue
End Property


Public Property Get CrossTabID() As Long
  CrossTabID = mlngCrossTabID
End Property

Public Property Let CrossTabID(ByVal lngNewID As Long)
  mlngCrossTabID = lngNewID
End Property
