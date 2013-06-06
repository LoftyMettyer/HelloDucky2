VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQAProProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mstrName As String
Private mstrModule As String
Private mdatTimeEntered As Date
Private mobjNextProc As clsQAProProcedure

' Name is the name of the procedure--
' note that it is a write-once property
Property Get Name() As String
    Name = mstrName
End Property
Property Let Name(strName As String)
    If mstrName = "" Then
        mstrName = strName
    End If
End Property

' Module is the name of the module this
' procedure is located in
Property Get Module() As String
    Module = mstrModule
End Property
Property Let Module(strModule As String)
    If mstrModule = "" Then
        mstrModule = strModule
    End If
End Property

' NextProc is used a pointer to the next
' procedure in the stack
Property Get NextProc() As clsQAProProcedure
    Set NextProc = mobjNextProc
End Property
Property Set NextProc(objProc As clsQAProProcedure)
    Set mobjNextProc = objProc
End Property

' TimeEntered is the date/time that the class
' instance was created
Property Get TimeEntered() As Date
    TimeEntered = mdatTimeEntered
End Property

Private Sub Class_Initialize()
    ' Set date/time entered
    mdatTimeEntered = Now
End Sub

