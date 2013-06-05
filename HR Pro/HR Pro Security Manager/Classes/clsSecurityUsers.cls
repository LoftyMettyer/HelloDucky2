VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SecurityUser"
Attribute VB_Ext_KEY = "Member0" ,"SecurityUser"
Option Explicit

' Local variables.
Private mvar_colUsers As Collection

Public Function Add(psUserName As String, _
  pfChanged As Boolean, _
  pfDeleteUser As Boolean, _
  pfNewUser As Boolean, _
  psMovedUserFrom As String, _
  psMovedUserTo As String, _
  psLogin As String, _
  psPassword As String, _
  Optional psKey As String, _
  Optional plngPersonnelRecordID As Long, _
  Optional pbForcePasswordChange, _
  Optional piLoginType As SecurityMgr.LoginType, _
  Optional pbIsPolicyChecked As Boolean) As SecurityUser
    
  Dim objNewMember As SecurityUser
  
  ' Instantiate a new user object.
  Set objNewMember = New SecurityUser

  'set the properties passed into the method
  With objNewMember
    .Login = psLogin
    .UserName = psUserName
    .Changed = pfChanged
    .DeleteUser = pfDeleteUser
    .NewUser = pfNewUser
    .MovedUserFrom = psMovedUserFrom
    .MovedUserTo = psMovedUserTo
    .Password = psPassword
    .Personnel_RecordID = plngPersonnelRecordID
    .ForcePasswordChange = IIf(IsMissing(pbForcePasswordChange), False, pbForcePasswordChange)
    .CheckPolicy = IIf(IsMissing(pbIsPolicyChecked), False, pbIsPolicyChecked)
    .LoginType = IIf(IsMissing(piLoginType), iUSERTYPE_SQLLOGIN, piLoginType)
  End With
  
  ' Add the new SecurityUser object to the collection.
  If Len(psKey) = 0 Then
      mvar_colUsers.Add objNewMember, psUserName
  Else
      mvar_colUsers.Add objNewMember, psKey
  End If

  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function
Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objUser As SecurityUser

  fOK = True

  For Each objUser In mvar_colUsers
    If Not objUser.DeleteUser Then objUser.PrintSecurity piLevel + 1
    'Exit For
  Next objUser

TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Property Get Item(ByRef pvIndexKey As Variant) As SecurityUser
Attribute Item.VB_UserMemId = 0
  On Error GoTo err_item
  
  Set Item = mvar_colUsers(pvIndexKey)
Exit Property

err_item:
  Set Item = Nothing
  
End Property



Public Property Get Count() As Long
  Count = mvar_colUsers.Count
  
End Property


Public Sub Remove(pvIndexKey As Variant)
  mvar_colUsers.Remove pvIndexKey
  
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colUsers.[_NewEnum]
    
End Property


Private Sub Class_Initialize()
  ' Instantiate the collection of SecurityUser objects.
  Set mvar_colUsers = New Collection
    
End Sub


Private Sub Class_Terminate()
    ' Destroy the collection when this class is terminates.
    Set mvar_colUsers = Nothing
    
End Sub

Public Function Clone(ByVal pbPerfectCopy As Boolean) As SecurityUsers

  Dim objUser As SecurityUser
  Dim objDestUser As SecurityUser
  Dim bChanged As Boolean
  Dim bDeleteUser As Boolean
  Dim bNewUser As Boolean

  Set Clone = New SecurityUsers

  If pbPerfectCopy Then
    For Each objUser In mvar_colUsers
      
      bChanged = objUser.Changed
      bDeleteUser = objUser.DeleteUser
      bNewUser = objUser.NewUser
      
      'Set objDestUser = New SecurityUser
      Set objDestUser = Clone.Add(objUser.UserName, bChanged, bDeleteUser, bNewUser, objUser.MovedUserFrom _
        , objUser.MovedUserTo, objUser.Login, objUser.Password, objUser.UserName, objUser.Personnel_RecordID _
        , objUser.ForcePasswordChange, objUser.LoginType, objUser.CheckPolicy)
  
    Next
  End If

  Set objUser = Nothing
  
End Function

Public Function Paste(ByVal pobjNewUser As SecurityUser) As Boolean

  Dim bOK As Boolean

  On Error GoTo ErrorTrap

  bOK = True
  mvar_colUsers.Add pobjNewUser, pobjNewUser.UserName

TidyUpAndExit:
  Paste = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function
