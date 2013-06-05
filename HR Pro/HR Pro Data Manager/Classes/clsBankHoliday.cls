VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBankHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrRegion As String
Private mstrDescription As String
Private mdtHolidayDate As Date
Private mlngBaseRecordID As String
Public Property Get Region() As String
  Region = mstrRegion
End Property

Public Property Let Region(pstrRegion As String)
  mstrRegion = pstrRegion
End Property

Public Property Get Description() As String
  Description = mstrDescription
End Property

Public Property Let Description(pstrDescription As String)
  mstrDescription = pstrDescription
End Property

Public Property Get HolidayDate() As Date
  HolidayDate = mdtHolidayDate
End Property

Public Property Let HolidayDate(pdtHolidayDate As Date)
  mdtHolidayDate = pdtHolidayDate
End Property

Public Property Get BaseRecordID() As Long
  BaseRecordID = mlngBaseRecordID
End Property

Public Property Let BaseRecordID(plngBaseRecordID As Long)
  mlngBaseRecordID = plngBaseRecordID
End Property
