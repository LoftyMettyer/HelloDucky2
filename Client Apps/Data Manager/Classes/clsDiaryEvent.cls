VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiaryEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrEventTitle As String
Private mdtEventDate As Date
Private mstrEventTime As String
Private mstrEventNotes As String
Private mstrUserName As String
Private mblnAlarm As Boolean
Private mstrAccess As String
Private mlngCopiedFromID As Long
  
Public Property Get EventTitle() As String
  EventTitle = mstrEventTitle
End Property
  
Public Property Let EventTitle(ByVal strNewValue As String)
  mstrEventTitle = strNewValue
End Property
  
Public Property Get EventDate() As Date
  EventDate = mdtEventDate
End Property
  
Public Property Let EventDate(ByVal dtNewValue As Date)
  mdtEventDate = dtNewValue
End Property
  
Public Property Get EventTime() As String
  EventTime = mstrEventTime
End Property
  
Public Property Let EventTime(ByVal strNewValue As String)
  mstrEventTime = strNewValue
End Property

Public Property Get EventNotes() As String
  EventNotes = mstrEventNotes
End Property
  
Public Property Let EventNotes(ByVal strNewValue As String)
  mstrEventNotes = strNewValue
End Property
  
Public Property Get UserName() As String
  UserName = mstrUserName
End Property
  
Public Property Let UserName(ByVal strNewValue As String)
  mstrUserName = strNewValue
End Property
  
Public Property Get Alarm() As Boolean
  Alarm = mblnAlarm
End Property
  
Public Property Let Alarm(ByVal blnNewValue As Boolean)
  mblnAlarm = blnNewValue
End Property
  
Public Property Get Access() As String
  Access = mstrAccess
End Property
  
Public Property Let Access(ByVal strNewValue As String)
  mstrAccess = strNewValue
End Property
  
Public Property Get CopiedFromID() As Long
  CopiedFromID = mlngCopiedFromID
End Property
  
Public Property Let CopiedFromID(ByVal lngNewValue As Long)
  mlngCopiedFromID = lngNewValue
End Property
