VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoryScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngScreenID As Long
Private msScreenName As String
Private mlngPictureID As Long
Private mlngTableID As Long
Private mlngViewID As Long
Private msViewName As String
Private msTableName As String

Public Property Get HistoryScreenName() As String
  HistoryScreenName = msScreenName

End Property

Public Property Get ViewName() As String
  ViewName = msViewName

End Property
Public Property Get TableName() As String
  TableName = msTableName

End Property


Public Property Let TableID(ByVal plngValue As Long)
  mlngTableID = plngValue

End Property

Public Property Get ViewID() As Long
  ViewID = mlngViewID

End Property

Public Property Let PictureID(ByVal plngValue As Long)
  mlngPictureID = plngValue

End Property

Public Property Get PictureID() As Long
  PictureID = mlngPictureID

End Property

Public Property Let ViewID(ByVal plngValue As Long)
  mlngViewID = plngValue

End Property

Public Property Get TableID() As Long
  TableID = mlngTableID

End Property

Public Property Let ViewName(ByVal psValue As String)
  msViewName = psValue

End Property

Public Property Let TableName(ByVal psValue As String)
  msTableName = psValue

End Property


Public Property Let HistoryScreenName(ByVal psValue As String)
  msScreenName = psValue

End Property

Public Property Get HistoryScreenID() As Long
  HistoryScreenID = mlngScreenID

End Property

Public Property Let HistoryScreenID(ByVal plngValue As Long)
  mlngScreenID = plngValue

End Property


