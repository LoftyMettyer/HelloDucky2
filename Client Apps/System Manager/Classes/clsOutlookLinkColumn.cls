VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutlookLinkColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintColumnID As Integer
Private mstrHeading As String
Private mintSequence As Integer

Public Property Get ColumnID() As Integer
  ColumnID = mintColumnID
End Property

Public Property Let ColumnID(ByVal intNewValue As Integer)
  mintColumnID = intNewValue
End Property

Public Property Get Heading() As String
  Heading = mstrHeading
End Property

Public Property Let Heading(ByVal strNewValue As String)
  mstrHeading = strNewValue
End Property

Public Property Get Sequence() As Integer
  Sequence = mintSequence
End Property

Public Property Let Sequence(ByVal intNewValue As Integer)
  mintSequence = intNewValue
End Property

