VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private datData As asruserdata.clsDataAccess
Private datData As HrPro.clsDataAccess

Private Sub Class_Initialize()
'  Set datData = New clsDataAccess
  Set datData = New HrPro.clsDataAccess

End Sub

Private Sub Class_Terminate()
  Set datData = Nothing

End Sub








