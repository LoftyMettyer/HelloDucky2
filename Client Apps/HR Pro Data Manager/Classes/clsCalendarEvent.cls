VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendarEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrEventKey As String
Private mstrName As String                'also used for Calendar Breakdown
Private mlngTableID As Long
Private mstrTableName As String
Private mlngFilterID As Long
Private mlngStartDateID As Long
Private mstrStartDateName As String       'also used for Calendar Breakdown
Private mlngStartSessionID As Long
Private mstrStartSessionName As String    'also used for Calendar Breakdown
Private mlngEndDateID As Long
Private mstrEndDateName As String         'also used for Calendar Breakdown
Private mlngEndSessionID As Long
Private mstrEndSessionName As String      'also used for Calendar Breakdown
Private mlngDurationID As Long
Private mstrDurationName As String        'also used for Calendar Breakdown
Private mintLegendType As Integer
Private mstrLegendCharacter As String     'also used for Calendar Breakdown
Private mlngLegendTableID As Long
Private mstrLegendTableName As String
Private mlngLegendColumnID As Long
Private mstrLegendColumnName As String
Private mlngLegendCodeID As Long
Private mstrLegendCodeName As String
Private mlngLegendEventTypeID As Long
Private mstrLegendEventTypeName As String
Private mlngDesc1ID As Long
Private mstrDesc1Name As String           'also used for Calendar Breakdown
Private mlngDesc2ID As Long
Private mstrDesc2Name As String           'also used for Calendar Breakdown

'extra properties added for Calendar Breakdown
Private mstrBaseDescription As String
Private mstrRegion As String
Private mstrWorkingPattern As String
Private mstrDesc1Value As String
Private mstrDesc2Value As String
Public Property Let Key(pstrEventKey As String)
  mstrEventKey = pstrEventKey
End Property
Public Property Get Key() As String
  Key = mstrEventKey
End Property
Public Property Let Name(pstrName As String)
  mstrName = pstrName
End Property
Public Property Get Name() As String
  Name = mstrName
End Property
Public Property Let TableID(plngTableID As Long)
  mlngTableID = plngTableID
End Property
Public Property Get TableID() As Long
  TableID = mlngTableID
End Property
Public Property Let TableName(pstrTableName As String)
  mstrTableName = pstrTableName
End Property
Public Property Get TableName() As String
  TableName = mstrTableName
End Property
Public Property Let FilterID(plngFilterID As Long)
  mlngFilterID = plngFilterID
End Property
Public Property Get FilterID() As Long
  FilterID = mlngFilterID
End Property
Public Property Let StartDateID(plngStartDateID As Long)
  mlngStartDateID = plngStartDateID
End Property
Public Property Get StartDateID() As Long
  StartDateID = mlngStartDateID
End Property
Public Property Let StartDateName(pstrStartDateName As String)
  mstrStartDateName = pstrStartDateName
End Property
Public Property Get StartDateName() As String
  StartDateName = mstrStartDateName
End Property
Public Property Let StartSessionID(plngStartSessionID As Long)
  mlngStartSessionID = plngStartSessionID
End Property
Public Property Get StartSessionID() As Long
  StartSessionID = mlngStartSessionID
End Property
Public Property Let StartSessionName(pstrStartSessionName As String)
  mstrStartSessionName = pstrStartSessionName
End Property
Public Property Get StartSessionName() As String
  StartSessionName = mstrStartSessionName
End Property
Public Property Let EndDateID(plngEndDateID As Long)
  mlngEndDateID = plngEndDateID
End Property
Public Property Get EndDateID() As Long
  EndDateID = mlngEndDateID
End Property
Public Property Let EndDateName(pstrEndDateName As String)
  mstrEndDateName = pstrEndDateName
End Property
Public Property Get EndDateName() As String
  EndDateName = mstrEndDateName
End Property
Public Property Let EndSessionID(plngEndSessionID As Long)
  mlngEndSessionID = plngEndSessionID
End Property
Public Property Get EndSessionID() As Long
  EndSessionID = mlngEndSessionID
End Property
Public Property Let EndSessionName(pstrEndSessionName As String)
  mstrEndSessionName = pstrEndSessionName
End Property
Public Property Get EndSessionName() As String
  EndSessionName = mstrEndSessionName
End Property
Public Property Let DurationID(plngDurationID As Long)
  mlngDurationID = plngDurationID
End Property
Public Property Get DurationID() As Long
  DurationID = mlngDurationID
End Property
Public Property Let DurationName(pstrDurationName As String)
  mstrDurationName = pstrDurationName
End Property
Public Property Get DurationName() As String
  DurationName = mstrDurationName
End Property
Public Property Let LegendType(pintLegendType As Integer)
  mintLegendType = pintLegendType
End Property
Public Property Get LegendType() As Integer
  LegendType = mintLegendType
End Property
Public Property Let LegendCharacter(pstrLegendCharacter As String)
  mstrLegendCharacter = pstrLegendCharacter
End Property
Public Property Get LegendCharacter() As String
  LegendCharacter = mstrLegendCharacter
End Property
Public Property Let LegendTableID(plngLegendTableID As Long)
  mlngLegendTableID = plngLegendTableID
End Property
Public Property Get LegendTableID() As Long
  LegendTableID = mlngLegendTableID
End Property
Public Property Let LegendTableName(pstrLegendTableName As String)
  mstrLegendTableName = pstrLegendTableName
End Property
Public Property Get LegendTableName() As String
  LegendTableName = mstrLegendTableName
End Property
Public Property Let LegendColumnID(plngLegendColumnID As Long)
  mlngLegendColumnID = plngLegendColumnID
End Property
Public Property Get LegendColumnID() As Long
  LegendColumnID = mlngLegendColumnID
End Property
Public Property Let LegendColumnName(pstrLegendColumnName As String)
  mstrLegendColumnName = pstrLegendColumnName
End Property
Public Property Get LegendColumnName() As String
  LegendColumnName = mstrLegendColumnName
End Property
Public Property Let LegendCodeID(plngLegendCodeID As Long)
  mlngLegendCodeID = plngLegendCodeID
End Property
Public Property Get LegendCodeID() As Long
  LegendCodeID = mlngLegendCodeID
End Property
Public Property Let LegendCodeName(pstrLegendCodeName As String)
  mstrLegendCodeName = pstrLegendCodeName
End Property
Public Property Get LegendCodeName() As String
  LegendCodeName = mstrLegendCodeName
End Property
Public Property Let LegendEventTypeID(plngLegendEventTypeID As Long)
  mlngLegendEventTypeID = plngLegendEventTypeID
End Property
Public Property Get LegendEventTypeID() As Long
  LegendEventTypeID = mlngLegendEventTypeID
End Property
Public Property Let LegendEventTypeName(pstrLegendEventTypeName As String)
  mstrLegendEventTypeName = pstrLegendEventTypeName
End Property
Public Property Get LegendEventTypeName() As String
  LegendEventTypeName = mstrLegendEventTypeName
End Property
Public Property Let Description1ID(plngDesc1ID As Long)
  mlngDesc1ID = plngDesc1ID
End Property
Public Property Get Description1ID() As Long
  Description1ID = mlngDesc1ID
End Property
Public Property Let Description1Name(pstrDesc1Name As String)
  mstrDesc1Name = pstrDesc1Name
End Property
Public Property Get Description1Name() As String
  Description1Name = mstrDesc1Name
End Property
Public Property Let Description2ID(plngDesc2ID As Long)
  mlngDesc2ID = plngDesc2ID
End Property
Public Property Get Description2ID() As Long
  Description2ID = mlngDesc2ID
End Property
Public Property Let Description2Name(pstrDesc2Name As String)
  mstrDesc2Name = pstrDesc2Name
End Property
Public Property Get Description2Name() As String
  Description2Name = mstrDesc2Name
End Property
Public Property Let BaseDescription(pstrBaseDescription As String)
  mstrBaseDescription = pstrBaseDescription
End Property
Public Property Get BaseDescription() As String
  BaseDescription = mstrBaseDescription
End Property
Public Property Let Region(pstrRegion As String)
  mstrRegion = pstrRegion
End Property
Public Property Get Region() As String
  Region = mstrRegion
End Property
Public Property Let WorkingPattern(pstrWorkingPattern As String)
  mstrWorkingPattern = pstrWorkingPattern
End Property
Public Property Get WorkingPattern() As String
  WorkingPattern = mstrWorkingPattern
End Property
Public Property Let Desc1Value(pstrDesc1Value As String)
  mstrDesc1Value = pstrDesc1Value
End Property
Public Property Get Desc1Value() As String
  Desc1Value = mstrDesc1Value
End Property
Public Property Let Desc2Value(pstrDesc2Value As String)
  mstrDesc2Value = pstrDesc2Value
End Property
Public Property Get Desc2Value() As String
  Desc2Value = mstrDesc2Value
End Property
