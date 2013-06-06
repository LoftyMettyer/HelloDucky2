VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatchRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngTable1ID As Long
Private strTable1Name As String
Private strTable1RealSource As String
Private lngTable2ID As Long
Private strTable2Name As String
Private strTable2RealSource As String
Private lngRequiredExprID As Long
Private lngPreferredExprID As Long
Private lngMatchScoreExprID As Long
Private colBreakdownColumns As Collection


Private Sub Class_Initialize()
  Set colBreakdownColumns = New Collection
End Sub

Private Sub Class_Terminate()
  Set colBreakdownColumns = Nothing
End Sub

Public Property Get Table1ID() As Long
  Table1ID = lngTable1ID
End Property

Public Property Let Table1ID(ByVal lngNewValue As Long)
  lngTable1ID = lngNewValue
End Property

Public Property Get Table1Name() As String
  Table1Name = strTable1Name
End Property

Public Property Let Table1Name(ByVal strNewValue As String)
  strTable1Name = strNewValue
End Property

Public Property Get Table1RealSource() As String
  Table1RealSource = strTable1RealSource
End Property

Public Property Let Table1RealSource(ByVal strNewValue As String)
  strTable1RealSource = strNewValue
End Property

Public Property Get Table2ID() As Long
  Table2ID = lngTable2ID
End Property

Public Property Let Table2ID(ByVal lngNewValue As Long)
  lngTable2ID = lngNewValue
End Property

Public Property Get Table2Name() As String
  Table2Name = strTable2Name
End Property

Public Property Let Table2Name(ByVal strNewValue As String)
  strTable2Name = strNewValue
End Property

Public Property Get Table2RealSource() As String
  Table2RealSource = strTable2RealSource
End Property

Public Property Let Table2RealSource(ByVal strNewValue As String)
  strTable2RealSource = strNewValue
End Property

Public Property Get RequiredExprID() As Long
  RequiredExprID = lngRequiredExprID
End Property

Public Property Let RequiredExprID(ByVal lngNewValue As Long)
  lngRequiredExprID = lngNewValue
End Property

Public Property Get PreferredExprID() As Long
  PreferredExprID = lngPreferredExprID
End Property

Public Property Let PreferredExprID(ByVal lngNewValue As Long)
  lngPreferredExprID = lngNewValue
End Property

Public Property Get MatchScoreID() As Long
  MatchScoreID = lngMatchScoreExprID
End Property

Public Property Let MatchScoreID(ByVal lngNewValue As Long)
  lngMatchScoreExprID = lngNewValue
End Property

Public Property Get BreakdownColumns() As Collection
  Set BreakdownColumns = colBreakdownColumns
End Property

Public Property Let BreakdownColumns(colTemp As Collection)
  Set colBreakdownColumns = colTemp
End Property
