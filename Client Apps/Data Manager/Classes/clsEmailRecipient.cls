VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmailRecipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrFixedEmail As String
Private mstrEmailName As String

Public Property Get FixedEmail() As String
  FixedEmail = mstrFixedEmail
End Property
Public Property Let FixedEmail(pstrNewValue As String)
  mstrFixedEmail = pstrNewValue
End Property

Public Property Get EmailName() As String
  EmailName = mstrEmailName
End Property
Public Property Let EmailName(pstrNewValue As String)
  mstrEmailName = pstrNewValue
End Property

