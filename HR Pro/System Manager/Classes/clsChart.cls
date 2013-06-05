VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim mlngChartTableID As Long

Public Property Get ChartTableID() As Long
  ChartTableID = mlngChartTableID
End Property

Public Property Let ChartTableID(ByVal lngNewValue As Long)
  mlngChartTableID = lngNewValue
End Property


