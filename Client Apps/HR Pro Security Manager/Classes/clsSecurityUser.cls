VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvar_sLogin As String
Private mvar_sUserName As String
Private mvar_sPassword As String
Private mvar_sMovedUserFrom As String
Private mvar_sMovedUserTo As String
Private mvar_fNewUser As Boolean
Private mvar_fChanged As Boolean
Private mvar_fDeleteUser As Boolean
Private mlngPersonnelID As Long
Private mbForcePasswordChange As Boolean
Private miLoginType As HrProSecurityMgr.LoginType
Private mbIsLocked As Boolean
Private mbIsMustChange As Boolean
Private mdPasswordLastSetTime As Date
Private mbIsExpired As Boolean
Private mbIsPolicyChecked As Boolean
Private mbIsExpirationChecked As Boolean

Public Property Let LoginType(ByVal piData As HrProSecurityMgr.LoginType)
  miLoginType = piData
End Property

Public Property Get LoginType() As HrProSecurityMgr.LoginType
  LoginType = miLoginType
End Property

Public Property Let Login(ByVal psData As String)
  mvar_sLogin = psData
End Property

Public Property Get Login() As String
  Login = mvar_sLogin
End Property

Public Property Let Password(ByVal psData As String)
  mvar_sPassword = psData
End Property

Public Property Get Password() As String
  Password = mvar_sPassword
End Property

Public Function PrintSecurity(piLevel As Integer) As Boolean
  
  Dim objIcon As IPictureDisp
  Dim strUserType As String
  Dim iCurrentY As Integer
  
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
'
  Dim fOK As Boolean
  Dim intOldFontSize As Integer
'
  fOK = True

  With Printer
    intOldFontSize = .FontSize
    .FontSize = 8
        
    .CurrentX = giPRINT_XINDENT + ((piLevel + 1) * giPRINT_XSPACE)
    .CurrentY = .CurrentY + giPRINT_YSPACE
    
    Select Case miLoginType
      Case iUSERTYPE_SQLLOGIN
        Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("USER_SQL").Picture
        strUserType = "SQL Server Login"
      
      Case iUSERTYPE_TRUSTEDGROUP
        Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("GROUP").Picture
        strUserType = "Windows Group"
      
      Case iUSERTYPE_TRUSTEDUSER
        Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("USER").Picture
        strUserType = "Windows User"
     
      ' NPG20090204 Fault 11931
      Case iUSERTYPE_ORPHANUSER
        Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("USER_ORPHAN").Picture
        strUserType = "Windows User (Orphan)"
     
      Case iUSERTYPE_ORPHANGROUP
        Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("GROUP_ORPHAN").Picture
        strUserType = "Windows Group (Orphan)"
     
    End Select
 
    .PaintPicture objIcon, Printer.CurrentX, Printer.CurrentY
    
    iCurrentY = .CurrentY
    .CurrentX = .CurrentX + objIcon.Width + 15
    Printer.Print mvar_sUserName
    
    .CurrentY = iCurrentY
    .CurrentX = (.ScaleWidth / 2)
    Printer.Print strUserType
    
    'Restore font size
    .FontSize = intOldFontSize
  End With
  
TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Let MovedUserTo(ByVal psData As String)
  mvar_sMovedUserTo = psData
End Property

Public Property Get MovedUserTo() As String
  MovedUserTo = mvar_sMovedUserTo
End Property

Public Property Let MovedUserFrom(ByVal psData As String)
  mvar_sMovedUserFrom = psData
End Property

Public Property Get MovedUserFrom() As String
  MovedUserFrom = mvar_sMovedUserFrom
End Property

Public Property Let NewUser(ByVal pfData As Boolean)
  mvar_fNewUser = pfData
End Property

Public Property Get NewUser() As Boolean
  NewUser = mvar_fNewUser
End Property

Public Property Let DeleteUser(ByVal pfData As Boolean)
  mvar_fDeleteUser = pfData
End Property

Public Property Get DeleteUser() As Boolean
  DeleteUser = mvar_fDeleteUser
End Property

Public Property Let Changed(ByVal pfData As Boolean)
  On Error GoTo ChangedLetErr
    
  mvar_fChanged = pfData
  Exit Property

ChangedLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityUser:Changed Property Let")
  MsgBox "Error SecurityUser:Changed Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get Changed() As Boolean
  On Error GoTo ChangedGetErr

  Changed = mvar_fChanged
  Exit Property

ChangedGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityUser:Changed Property Get")
  MsgBox "Error SecurityUser:Changed Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
    
End Property


Public Property Let UserName(ByVal psData As String)
  On Error GoTo UserNameLetErr
    
  mvar_sUserName = psData
  Exit Property

UserNameLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityUser:UserName Property Let")
  MsgBox "Error SecurityUser:UserName Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property


Public Property Get UserName() As String
  On Error GoTo UserNameGetErr
    
  UserName = mvar_sUserName
  Exit Property

UserNameGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityUser:UserName Property Get")
  MsgBox "Error SecurityUser:UserName Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let Personnel_RecordID(ByVal plngRecordID As Long)
  mlngPersonnelID = plngRecordID
End Property

Public Property Get Personnel_RecordID() As Long
  Personnel_RecordID = mlngPersonnelID
End Property

Public Property Get ForcePasswordChange() As Boolean
  ForcePasswordChange = mbForcePasswordChange
End Property

Public Property Let ForcePasswordChange(ByVal pbNewValue As Boolean)
  mbForcePasswordChange = pbNewValue
End Property

Public Property Get CheckPolicy() As Boolean
  CheckPolicy = mbIsPolicyChecked
End Property

Public Property Let CheckPolicy(ByVal bNewValue As Boolean)
  mbIsPolicyChecked = bNewValue
End Property

Public Function Clone() As SecurityUser

  Dim objNewUser As SecurityUser

  Set objNewUser = New SecurityUser
  
  ' Copy the settings
  With objNewUser
    .Changed = mvar_fChanged
    .DeleteUser = mvar_fDeleteUser
    .ForcePasswordChange = mbForcePasswordChange
    .CheckPolicy = mbIsPolicyChecked
    .Login = mvar_sLogin
    .MovedUserFrom = mvar_sMovedUserFrom
    .MovedUserTo = mvar_sMovedUserTo
    .NewUser = mvar_fNewUser
    .Password = mvar_sPassword
    .Personnel_RecordID = mlngPersonnelID
    .UserName = mvar_sUserName
    .LoginType = miLoginType
  End With
  
  ' Return the cloned security group
  Set Clone = objNewUser

End Function
