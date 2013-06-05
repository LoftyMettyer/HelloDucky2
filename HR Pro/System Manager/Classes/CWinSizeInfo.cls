VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinSizeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlhWnd As Long
Private miMaxX As Long
Private miMaxY As Long
Private miMinX As Long
Private miMinY As Long

Public Property Get hWnd() As Long
  hWnd = mlhWnd
End Property

Public Property Let hWnd(ByVal lngNewValue As Long)
  mlhWnd = lngNewValue
End Property

Public Property Get MinX() As Long
  MinX = miMinX
End Property

Public Property Let MinX(ByVal lngNewValue As Long)
  miMinX = lngNewValue
End Property

Public Property Get MinY() As Long
  MinY = miMinY
End Property

Public Property Let MinY(ByVal lngNewValue As Long)
  miMinY = lngNewValue
End Property

Public Property Get MaxX() As Long
  MaxX = miMaxX
End Property

Public Property Let MaxX(ByVal lngNewValue As Long)
 miMaxX = lngNewValue
End Property

Public Property Get MaxY() As Long
  MaxY = miMaxY
End Property

Public Property Let MaxY(ByVal lngNewValue As Long)
 miMaxY = lngNewValue
End Property

