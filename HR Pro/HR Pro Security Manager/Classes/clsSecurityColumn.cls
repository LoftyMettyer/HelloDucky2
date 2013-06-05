VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 0

#If DebugMode Then
  'local variable to hold the serialized class ID that was created in Class_Initialize
  Private mvar_lngClassDebugID As Long
#End If

' Local variables to hold property values.
Private mvar_fUpdatePrivilege As Boolean
Private mvar_fSelectPrivilege As Boolean
Private mvar_fUpdateOriginalPrivilege As Boolean
Private mvar_fSelectOriginalPrivilege As Boolean
Private mvar_fChanged As Boolean
Private mvar_sName As String

Private mlngBottom As Long

Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim intSameLine As Integer
  Dim intOldFontSize As Integer
  Dim intCurrentX As Integer

  fOK = True

  With Printer
    intOldFontSize = .FontSize
    .FontSize = 6
    'check where your bottom is
    mlngBottom = CalculateBottomOfPage()
    intSameLine = .CurrentY + (giPRINT_YSPACE * 1)
        
    .CurrentX = giPRINT_XINDENT + ((piLevel + 1) * giPRINT_XSPACE)
    .CurrentY = intSameLine
    Printer.Print mvar_sName
        
    'Read
    intCurrentX = giPRINT_XINDENT + ((piLevel + 5) * giPRINT_XINDENT) + 200
    Call PrintPictureBox(intCurrentX, intSameLine, mvar_fSelectPrivilege)
    'Edit
    intCurrentX = giPRINT_XINDENT + ((piLevel + 6) * giPRINT_XINDENT) + 250
    Call PrintPictureBox(intCurrentX, intSameLine, mvar_fUpdatePrivilege)
    
    CheckEndOfPage2 (mlngBottom)
  End With
  
TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Let Name(ByVal psData As String)
  On Error GoTo NameLetErr
  
  mvar_sName = psData
  Exit Property

NameLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:Name Property Let")
  MsgBox "Error SecurityColumn:Name Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get Name() As String
  On Error GoTo NameGetErr
  
  Name = mvar_sName
  Exit Property

NameGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:Name Property Get")
  MsgBox "Error SecurityColumn:Name Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let Changed(ByVal pfData As Boolean)
  On Error GoTo ChangedLetErr
  
  mvar_fChanged = pfData
  Exit Property

ChangedLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:Changed Property Let")
  MsgBox "Error SecurityColumn:Changed Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get Changed() As Boolean
  On Error GoTo ChangedGetErr
  
  Changed = mvar_fChanged
  Exit Property

ChangedGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:Changed Property Get")
  MsgBox "Error SecurityColumn:Changed Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let SelectPrivilege(ByVal pfData As Boolean)
  On Error GoTo SelectPrivilegeLetErr
  
  mvar_fSelectPrivilege = pfData
  Exit Property

SelectPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:SelectPrivilege Property Let")
  MsgBox "Error SecurityColumn:SelectPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property

Public Property Get SelectPrivilege() As Boolean
  On Error GoTo SelectPrivilegeGetErr
  
  SelectPrivilege = mvar_fSelectPrivilege
  Exit Property

SelectPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:SelectPrivilege Property Get")
  MsgBox "Error SecurityColumn:SelectPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property

Public Property Let UpdatePrivilege(ByVal pfData As Boolean)
  On Error GoTo UpdatePrivilegeLetErr
  
  mvar_fUpdatePrivilege = pfData
  Exit Property

UpdatePrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:UpdatePrivilege Property Let")
  MsgBox "Error SecurityColumn:UpdatePrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get UpdatePrivilege() As Boolean
  On Error GoTo UpdatePrivilegeGetErr
  
  UpdatePrivilege = mvar_fUpdatePrivilege
  Exit Property

UpdatePrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:UpdatePrivilege Property Get")
  MsgBox "Error SecurityColumn:UpdatePrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let SelectOriginalPrivilege(ByVal pfData As Boolean)
  On Error GoTo SelectOriginalPrivilegeLetErr
  
  mvar_fSelectOriginalPrivilege = pfData
  Exit Property

SelectOriginalPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:SelectOriginalPrivilege Property Let")
  MsgBox "Error SecurityColumn:SelectOriginalPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property

Public Property Get SelectOriginalPrivilege() As Boolean
  On Error GoTo SelectOriginalPrivilegeGetErr
  
  SelectOriginalPrivilege = mvar_fSelectOriginalPrivilege
  Exit Property

SelectOriginalPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:SelectOriginalPrivilege Property Get")
  MsgBox "Error SecurityColumn:SelectOriginalPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property

Public Property Let UpdateOriginalPrivilege(ByVal pfData As Boolean)
  On Error GoTo UpdateOriginalPrivilegeLetErr
  
  mvar_fUpdateOriginalPrivilege = pfData
  Exit Property

UpdateOriginalPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:UpdateOriginalPrivilege Property Let")
  MsgBox "Error SecurityColumn:UpdateOriginalPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get UpdateOriginalPrivilege() As Boolean
  On Error GoTo UpdateOriginalPrivilegeGetErr
  
  UpdateOriginalPrivilege = mvar_fUpdateOriginalPrivilege
  Exit Property

UpdateOriginalPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityColumn:UpdateOriginalPrivilege Property Get")
  MsgBox "Error SecurityColumn:UpdateOriginalPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Private Sub Class_Initialize()
  #If DebugMode Then
    ' Get the next available class ID, and print out
    ' that the class was created successfully.
    mvar_lngClassDebugID = GetNextClassDebugID()
'    Debug.Print "'" & TypeName(Me) & "' instance " & mvar_lngClassDebugID & " created"
  #End If
  
End Sub

Private Sub Class_Terminate()
  ' The class is being destroyed.
  #If DebugMode Then
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mvar_lngClassDebugID) & " is terminating"
  #End If
  
End Sub


#If DebugMode Then
  Public Property Get ClassDebugID()
    ' If we are in debug mode, surface this property that consumers can query.
    ClassDebugID = mvar_lngClassDebugID
  End Property
#End If
