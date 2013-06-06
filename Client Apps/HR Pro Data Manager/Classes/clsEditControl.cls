VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEditControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCalcExpressionID As Long 'local copy
Private mvarWhenExpressionID As Long 'local copy
Private mvarValidExpressionID As Long 'local copy
Private mvarCalcTrigger As Integer 'local copy
Private mvarUserChanged As Boolean 'local copy
Private mvarBlankIfZero As Boolean 'local copy
Private mvarErrorMessage As String 'local copy
Private mbReadOnly As Boolean
Private mbMandatory As Boolean
Private msDefaultValue As String
Private msStatusBarMessage As String
Private miDataType As DataMgr.SQLDataType
'#RH 07/09/99
Private mbAfdEnabled As Boolean

Public Property Get AFDEnabled() As Boolean

AFDEnabled = mbAfdEnabled

End Property

Public Property Get DataType() As DataMgr.SQLDataType
  DataType = miDataType

End Property


Public Property Let DataType(ByVal piNewValue As DataMgr.SQLDataType)
  miDataType = piNewValue

End Property
Public Property Let AFDEnabled(ByVal bAfdEnabled As Boolean)

  mbAfdEnabled = bAfdEnabled

End Property
'#RH 07/09/99

Public Property Let BlankIfZero(ByVal vData As Boolean)
    On Error GoTo BlankIfZeroLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlankIfZero = 5
    mvarBlankIfZero = vData
    Exit Property


BlankIfZeroLetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:BlankIfZero Property Let")
  COAMsgBox "Error CEditControl:BlankIfZero Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Get BlankIfZero() As Boolean
    On Error GoTo BlankIfZeroGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlankIfZero
    BlankIfZero = mvarBlankIfZero
    Exit Property


BlankIfZeroGetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:BlankIfZero Property Get")
  COAMsgBox "Error CEditControl:BlankIfZero Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Let UserChanged(ByVal vData As Boolean)
    On Error GoTo UserChangedLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserChanged = 5
    mvarUserChanged = vData
    Exit Property


UserChangedLetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:UserChanged Property Let")
  COAMsgBox "Error CEditControl:UserChanged Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Get UserChanged() As Boolean
    On Error GoTo UserChangedGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserChanged
    UserChanged = mvarUserChanged
    Exit Property


UserChangedGetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:UserChanged Property Get")
  COAMsgBox "Error CEditControl:UserChanged Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property



Public Property Let CalcTrigger(ByVal vData As Integer)
    On Error GoTo CalcTriggerLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CalcTrigger = 5
    mvarCalcTrigger = vData
    Exit Property


CalcTriggerLetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:CalcTrigger Property Let")
  COAMsgBox "Error CEditControl:CalcTrigger Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Get CalcTrigger() As Integer
    On Error GoTo CalcTriggerGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CalcTrigger
    CalcTrigger = mvarCalcTrigger
    Exit Property


CalcTriggerGetErr:
   'Call RaiseError(MyUnhandledError, "CEditControl:CalcTrigger Property Get")
   COAMsgBox "Error CEditControl:CalcTrigger Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property



Public Property Let CalcExpressionID(ByVal vData As Long)
    On Error GoTo CalcExpressionIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CalcExpressionID = 5
    mvarCalcExpressionID = vData
    Exit Property


CalcExpressionIDLetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:CalcExpressionID Property Let")
  COAMsgBox "Error CEditControl:CalcExpressionID Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Get CalcExpressionID() As Long
    On Error GoTo CalcExpressionIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CalcExpressionID
    CalcExpressionID = mvarCalcExpressionID
    Exit Property


CalcExpressionIDGetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:CalcExpressionID Property Get")
  COAMsgBox "Error CEditControl:CalcExpressionID Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property



Public Property Let ValidExpressionID(ByVal vData As Long)
    On Error GoTo ValidExpressionIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ValidExpressionID = 5
    mvarValidExpressionID = vData
    Exit Property


ValidExpressionIDLetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:ValidExpressionID Property Let")
  COAMsgBox "Error CEditControl:ValidExpressionID Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Get ValidExpressionID() As Long
    On Error GoTo ValidExpressionIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ValidExpressionID
    ValidExpressionID = mvarValidExpressionID
    Exit Property


ValidExpressionIDGetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:ValidExpressionID Property Get")
  COAMsgBox "Error CEditControl:ValidExpressionID Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Let WhenExpressionID(ByVal vData As Long)
    On Error GoTo WhenExpressionIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WhenExpressionID = 5
    mvarWhenExpressionID = vData
    Exit Property


WhenExpressionIDLetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:WhenExpressionID Property Let")
  COAMsgBox "Error CEditControl:WhenExpressionID Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Get WhenExpressionID() As Long
    On Error GoTo WhenExpressionIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WhenExpressionID
    WhenExpressionID = mvarWhenExpressionID
    Exit Property


WhenExpressionIDGetErr:
  'Call RaiseError(MyUnhandledError, "CEditControl:WhenExpressionID Property Get")
  COAMsgBox "Error CEditControl:WhenExpressionID Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property




Private Sub Class_Initialize()
'    #If DebugMode Then
'        'get the next available class ID, and print out
'        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
'        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
'    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
'    #If DebugMode Then
'        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
'    #End If
End Sub

Public Property Let ReadOnly(ByVal bReadOnly As Boolean)

    mbReadOnly = bReadOnly

End Property

Public Property Get ReadOnly() As Boolean

    ReadOnly = mbReadOnly

End Property

Public Property Let Mandatory(ByVal bMandatory As Boolean)

    mbMandatory = bMandatory

End Property

Public Property Get Mandatory() As Boolean

    Mandatory = mbMandatory
    
End Property
Public Property Let DefaultValue(ByVal psDefaultValue As String)

    msDefaultValue = psDefaultValue

End Property

Public Property Get ErrorMessage() As String

    ErrorMessage = mvarErrorMessage
    
End Property

Public Property Get StatusBarMessage() As String

    StatusBarMessage = msStatusBarMessage
    
End Property


Public Property Let ErrorMessage(ByVal psErrorMessage As String)

    mvarErrorMessage = psErrorMessage

End Property
Public Property Let StatusBarMessage(ByVal psStatusBarMessage As String)

    msStatusBarMessage = psStatusBarMessage

End Property

Public Property Get DefaultValue() As String

    DefaultValue = msDefaultValue
    
End Property


