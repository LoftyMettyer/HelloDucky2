VERSION 5.00
Begin VB.Form frmPasswordMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User "
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8018
   Icon            =   "frmPasswordMaintenance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Caption         =   "Options :"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   1755
      Width           =   4830
      Begin VB.CheckBox chkForceChange 
         Caption         =   "&User must change password at next logon"
         Height          =   270
         Left            =   135
         TabIndex        =   6
         Top             =   930
         Value           =   1  'Checked
         Width           =   3990
      End
      Begin VB.CheckBox chkEnforcePasswordPolicy 
         Caption         =   "&Enforce password policy"
         Height          =   300
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   3450
      End
      Begin VB.CheckBox chkEnforceExpiry 
         Caption         =   "E&nforce password expiration"
         Enabled         =   0   'False
         Height          =   225
         Left            =   135
         TabIndex        =   5
         Top             =   630
         Width           =   3450
      End
      Begin VB.CheckBox chkAccountIsLocked 
         Caption         =   "&Login is locked out"
         Enabled         =   0   'False
         Height          =   345
         Left            =   135
         TabIndex        =   7
         Top             =   1200
         Width           =   3450
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3750
      TabIndex        =   9
      Top             =   3585
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   2490
      TabIndex        =   8
      Top             =   3585
      Width           =   1200
   End
   Begin VB.Frame fraReset 
      Caption         =   "Password : "
      Height          =   1650
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   4845
      Begin VB.ComboBox cboUser 
         Height          =   315
         Left            =   1980
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   2745
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1110
         Width           =   2715
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   375
         Width           =   435
      End
      Begin VB.Label lblConfirmPassword 
         Caption         =   "Confirm Password : "
         Height          =   300
         Left            =   180
         TabIndex        =   11
         Top             =   1170
         Width           =   1770
      End
      Begin VB.Label lblNewPasswird 
         Caption         =   "New Password :"
         Height          =   270
         Left            =   195
         TabIndex        =   10
         Top             =   780
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmPasswordMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUserName As String
Private mstrSecurityGroupName As String
Private mbShowAllUsers As Boolean

Public Property Let ShowAllUsers(ByVal bNewValue As Boolean)
  mbShowAllUsers = bNewValue
End Property

Public Property Let UserName(ByVal strNewValue As String)
  mstrUserName = strNewValue
End Property

Public Property Let SecurityGroup(ByVal strNewValue As String)
  mstrSecurityGroupName = strNewValue
End Property

Public Function Initialise() As Boolean

  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  
  Screen.MousePointer = vbHourglass
  
  If mbShowAllUsers Then
    Me.Caption = "Password Maintenance"
  Else
    'NHRD13012005 Fault 10723
    ' Seems a bit crap but - Left original code in just in case minds are changed in future.
    'Me.Caption = "Reset Password"
    Me.Caption = "Password Maintenance"
  End If
  
  chkEnforcePasswordPolicy.Enabled = (glngSQLVersion >= 9)
  'NHRD13012005 Fault 10722
  'chkEnforceExpiry.Enabled = (glngSQLVersion >= 9)
  chkEnforceExpiry.Enabled = (glngSQLVersion >= 9)
   
  If mbShowAllUsers Then
    For Each objGroup In gObjGroups
    
      ' Load the users into collections...may take a while
      If Not gObjGroups(objGroup.Name).Users_Initialised Then
        InitialiseUsersCollection gObjGroups(objGroup.Name)
      End If
    
      For Each objUser In gObjGroups(objGroup.Name).Users
        'JPD 20040224 Fault 8131
        'If (Not objUser.DeleteUser) And _
          (objUser.MovedUserTo = "") Then
        If (Not objUser.DeleteUser) And _
          (objUser.MovedUserTo = "") And _
          (Not objUser.NewUser) And _
          (objUser.LoginType = iUSERTYPE_SQLLOGIN) Then
          cboUser.AddItem objUser.UserName
        End If
      Next objUser
    
    Next objGroup
  Else
    cboUser.AddItem mstrUserName
  End If

  If cboUser.ListCount > 0 Then
    If LenB(mstrUserName) <> 0 And gbUserCanManageLogins Then
      SetComboText cboUser, mstrUserName
      ControlsDisableAll cboUser
      Initialise = True
    ElseIf gbUserCanManageLogins Then
      ' we can change/reset anybodies passwords
      cboUser.ListIndex = 0
      Initialise = True
    ElseIf Not gbUseWindowsAuthentication Then
      ' we can only change/reset our own passwords
      SetComboText cboUser, gsUserName
      cboUser.Enabled = False
      cboUser.BackColor = vbButtonFace
      Initialise = True
    Else
      Initialise = False
    End If
  Else
    MsgBox "There are no OpenHR users mapped to SQL logins. Please select at least one user.", vbExclamation + vbOKOnly, App.Title
    Initialise = False
  End If

  Screen.MousePointer = vbNormal
  Exit Function
  
Init_ERROR:
  
  Screen.MousePointer = vbNormal
  Initialise = False
  MsgBox "Error reading user information." & vbNewLine & vbNewLine & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  
End Function

'Private Function UserLoggedIn(sUserID As String) As String
'
'  Dim sSQL As String
'  Dim rsUsers As New ADODB.Recordset
'
'  On Error GoTo Error_Trap
'
'  sSQL = "SELECT DISTINCT hostname, loginame, program_name, hostprocess " & _
'    "FROM master..sysprocesses " & _
'    "WHERE program_name like 'OpenHR%' " & _
'    "  AND program_name NOT LIKE 'OpenHR Workflow%' " & _
'    "  AND LOWER(loginame) = '" & LCase(sUserID) & "' "
''                    "AND dbid in (" & _
''                        "SELECT dbid " & _
''                        "FROM master..sysdatabases " & _
''                        "WHERE name = '" & gsDatabaseName & "') "
'
'  rsUsers.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'  With rsUsers
'    If Not (.EOF And .BOF) Then
'      UserLoggedIn = Trim(!Program_name)
'    Else
'      UserLoggedIn = vbNullString
'    End If
'  End With
'  rsUsers.Close
'
'TidyUpAndExit:
'  Set rsUsers = Nothing
'  Exit Function
'
'Error_Trap:
'  MsgBox "Error validating current users.", vbExclamation + vbOKOnly, App.Title
'  UserLoggedIn = True
'  GoTo TidyUpAndExit
'
'End Function

Private Sub cboUser_Click()
  LoadPropertiesForUser
End Sub

Private Sub chkAccountIsLocked_Click()
  cmdOK.Enabled = True
End Sub

Private Sub chkEnforceExpiry_Click()

  If chkEnforceExpiry.Value = vbUnchecked Then
    chkForceChange.Value = vbUnchecked
  End If

  chkForceChange.Enabled = (chkEnforceExpiry.Value = vbChecked) Or glngSQLVersion < 9
  cmdOK.Enabled = True

End Sub

Private Sub chkEnforcePasswordPolicy_Click()

  If chkEnforcePasswordPolicy.Value = vbUnchecked Then
    chkEnforceExpiry.Value = vbUnchecked
    chkForceChange.Value = vbUnchecked
  End If

  chkEnforceExpiry.Enabled = (chkEnforcePasswordPolicy.Value = vbChecked)
  chkForceChange.Enabled = (chkEnforceExpiry.Value = vbChecked) Or glngSQLVersion < 9
  cmdOK.Enabled = True

End Sub

Private Sub chkForceChange_Click()
  cmdOK.Enabled = True
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()

  If ValidateIt Then
    If SaveIt Then
      Unload Me
    End If
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub optForceChange_Click()

  With txtPassword
    .Enabled = False
    .BackColor = vbButtonFace
  End With
  
End Sub

Private Sub optReset_Click()

  With txtPassword
    .Enabled = True
    .BackColor = vbWindowBackground
    .SetFocus
  End With
  
End Sub

Private Function ValidateIt() As Boolean

  Dim sHRProApp As String
  
  ' AE20080425 not used?
  'Dim bBypassPolicy As Boolean
  'bBypassPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)   ' Default - Off
  
  sHRProApp = vbNullString
  
  On Error GoTo Validate_ERROR
  
  'MH20061017 Fault 11376
  '''TM20020103 Fault 3317 - Check if the user is logged on or not!
  'sHRProApp = UserLoggedIn(Me.cboUser.Text)
  sHRProApp = GetCurrentUsersAppName(cboUser.Text)
  If sHRProApp <> vbNullString Then
    MsgBox "Cannot change the password for '" & Me.cboUser.Text & _
            "' as this user is logged into the " & sHRProApp & ".", _
            vbExclamation + vbOKOnly, App.Title
    ValidateIt = False
    Exit Function
  End If
  
  ' Confirmation
  If (Trim(txtPassword.Text) <> Trim(txtConfirmPassword.Text)) Then
    MsgBox "The confirmation password is not correct.", vbExclamation + vbOKOnly, App.ProductName
    txtConfirmPassword.SetFocus
    Exit Function
  End If
  
  ' Long enough
  If glngSQLVersion < 9 Then
    If CheckOverMaxLength = False Then
      ValidateIt = False
      Exit Function
    End If
    
''    ' Complex enough
''    If CheckPasswordComplexity(cboUser.Text, txtPassword.Text) = False Then
''      ValidateIt = False
''      MsgBox "The new password does not meet the policy requirement because it is not complicated enough.", vbExclamation + vbOKOnly, App.Title
''      Exit Function
''    End If
  End If
    
  ValidateIt = True
  Exit Function
  
Validate_ERROR:
  
  MsgBox "Error validating selection." & vbNewLine & vbNewLine & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  ValidateIt = False
  
End Function

Private Function SaveIt() As Boolean

  On Error GoTo Save_ERROR

  Dim rsInfo As New ADODB.Recordset
  Dim strNewPassword As String
  Dim sMessage As String
  Dim bOK As Boolean
  Dim sSQL As String

  bOK = True

  ' If user is resetting their own password (why !??!), update the global string
  If LCase(cboUser.Text) = LCase(gsUserName) Then
    gADOCon.Execute "EXEC sp_password '" & gsPassword & "', '" & Replace(Me.txtPassword.Text, "'", "''") & "'", , adExecuteNoRecords
    gsPassword = Replace(Me.txtPassword.Text, "'", "''")
  Else
  
    ' AE20080415 Fault #13102
    If mstrSecurityGroupName = vbNullString Then
      IsUserNameInUse cboUser.Text, gObjGroups, mstrSecurityGroupName
    End If
    ' AE20080325 Fault #12827
    ' If this is a user we've just created then update the collection
    If gObjGroups(mstrSecurityGroupName).Users(cboUser.Text).NewUser Then
      With gObjGroups(mstrSecurityGroupName).Users(cboUser.Text)
        .Changed = True
        .Password = txtPassword.Text
        .ForcePasswordChange = CBool(chkForceChange.Value = vbChecked)
        .CheckPolicy = CBool(chkEnforceExpiry.Value = vbChecked)
      End With
    
    ElseIf glngSQLVersion < 9 Then
      strNewPassword = IIf(Len(txtPassword.Text) = 0, "null", "'" & Replace(txtPassword.Text, "'", "''") & "'")
      gADOCon.Execute "EXEC sp_password NULL," & strNewPassword & ", '" & Replace(Me.cboUser.Text, "'", "''") & "'", , adExecuteNoRecords
    
      ' If we need to force user to change at next login
      If chkForceChange.Value = vbChecked Then
        rsInfo.Open "SELECT * FROM ASRSysPasswords WHERE Username = '" & Replace(LCase(cboUser.Text), "'", "''") & "'", gADOCon, adOpenForwardOnly, adLockOptimistic
        If rsInfo.BOF And rsInfo.EOF Then
          rsInfo.AddNew
        End If
        
        rsInfo!UserName = LCase(cboUser.Text)
        rsInfo!LastChanged = Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
        rsInfo!ForceChange = IIf(chkForceChange.Value = vbChecked, 1, 0)
        
        rsInfo.Update
        rsInfo.Close
        
      End If
      
      gObjGroups(mstrSecurityGroupName).Users(cboUser.Text).ForcePasswordChange = IIf(chkForceChange.Value = vbChecked, True, False)
    
    Else
      
      sSQL = "ALTER LOGIN [" & Trim(cboUser.Text) & "] WITH PASSWORD = " _
        & IIf(Len(txtPassword.Text) = 0, "''", "'" & Replace(txtPassword.Text, "'", "''") & "'") _
        & IIf(chkAccountIsLocked.Enabled = True And chkAccountIsLocked.Value = vbUnchecked, " UNLOCK", "") _
        & IIf(chkForceChange.Value = vbChecked, " MUST_CHANGE", "") _
        & ",DEFAULT_DATABASE=master" _
        & ",CHECK_POLICY=" & IIf(chkEnforcePasswordPolicy.Value = vbChecked, "ON", "OFF") _
        & ",CHECK_EXPIRATION=" & IIf(chkEnforceExpiry.Value = vbChecked, "ON", "OFF")
      gADOCon.Execute sSQL, , adExecuteNoRecords
    
    End If
    
  End If
  
  'MsgBox "The password for user '" & cboUser.Text & "' has been changed.", vbInformation + vbOKOnly, App.Title
  
TidyUpAndExit:
  SaveIt = bOK
  Set rsInfo = Nothing
  Exit Function

Save_ERROR:
  bOK = False
  sMessage = vbNullString
  
  If gADOCon.Errors.Count > 0 Then
    Select Case gADOCon.Errors(0).NativeError
        
        ' Password too short
        Case 15116
          sMessage = "The new password must be at least " & glngDomainMinimumLength & " characters long !"
        
        ' Password does not meet domain policy
        Case 15115, 15117, 15118
          sMessage = gADOCon.Errors(0).Description
        
        ' Someone else has changed the password
        Case 15151
          sMessage = "Old password incorrect."
      
        ' You can't change policy settings if the MUST_CHANGE has already been set on this account
        Case 15128
          sMessage = "You cannot change policy settings for this account because it has already " & _
             "been marked as requiring a password change."
          
      
        Case Else
          sMessage = gADOCon.Errors(0).Description
      
    End Select
  
  Else
    sMessage = "Error saving selection." & vbNewLine & vbNewLine & "(" & Err.Number & " - " & Err.Description & ")"
  End If
  
  MsgBox sMessage, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function


Private Function CheckOverMaxLength() As Boolean

  Dim lMinimumLength As Long      ' The minimum length for passwords
  
  ' First store the config info in local variables
  'Set rsInfo = rdoCon.OpenResultset("Select * From ASRSysConfig")
  'lMinimumLength = rsInfo!MinimumPasswordLength
  'Set rsInfo = Nothing
  lMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)

  If lMinimumLength = 0 Then
    CheckOverMaxLength = True
  ElseIf Len(Me.txtPassword.Text) >= lMinimumLength Then
    CheckOverMaxLength = True
  Else
    MsgBox "The new password must be at least " & lMinimumLength & " characters long.", vbExclamation + vbOKOnly, App.Title
    CheckOverMaxLength = False
  End If
  
End Function

Private Sub SetComboText(cboCombo As ComboBox, sText As String)

  Dim lCount As Long

  With cboCombo
    For lCount = 1 To .ListCount
      If LCase$(.List(lCount - 1)) = LCase$(sText) Then
        .ListIndex = lCount - 1
        Exit For
      End If
    Next
  End With

End Sub

Private Sub LoadPropertiesForUser()

  Dim rsUserInfo As New ADODB.Recordset
  Dim sSQL As String
  
  If glngSQLVersion >= 9 Then
    sSQL = "SELECT sys.syslogins.loginname " & _
      ", LOGINPROPERTY(sys.syslogins.loginname, 'IsLocked') AS IsLocked" & _
      ", LOGINPROPERTY(sys.syslogins.loginname, 'IsMustChange') AS IsMustChange" & _
      ", LOGINPROPERTY(sys.syslogins.loginname, 'IsExpired') AS IsExpired" & _
      ", LOGINPROPERTY(sys.syslogins.loginname, 'PasswordLastSetTime') AS PasswordLastSetTime" & _
      ", inf.is_policy_checked, inf.is_expiration_checked" & _
      " FROM sys.syslogins" & _
      " LEFT JOIN sys.sql_logins inf ON inf.sid = sys.syslogins.sid" & _
      " WHERE inf.name = '" & Replace(cboUser.Text, "'", "''") & "'"
  
    rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsUserInfo.EOF And rsUserInfo.BOF) Then
      chkEnforcePasswordPolicy.Value = IIf(rsUserInfo.Fields("is_policy_checked").Value = 0, vbUnchecked, vbChecked)
      chkEnforcePasswordPolicy_Click
      chkEnforceExpiry.Value = IIf(rsUserInfo.Fields("is_expiration_checked").Value = 0, vbUnchecked, vbChecked)
      chkForceChange.Value = IIf(rsUserInfo.Fields("IsMustChange").Value = 0, vbUnchecked, vbChecked)
      chkForceChange.Enabled = (chkEnforceExpiry.Value = vbChecked)
      chkAccountIsLocked.Value = IIf(rsUserInfo.Fields("IsLocked").Value = 0, vbUnchecked, vbChecked)
      chkAccountIsLocked.Enabled = (chkAccountIsLocked.Value = vbChecked)
    Else
      ' New user
      Dim bIgnoreDomainPolicy As Boolean
      bIgnoreDomainPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)
      
      If bIgnoreDomainPolicy Then
        chkEnforcePasswordPolicy.Value = vbUnchecked
        chkEnforceExpiry.Enabled = False
        chkEnforceExpiry.Value = vbUnchecked
        chkForceChange.Enabled = False
        chkForceChange.Value = vbUnchecked
        chkAccountIsLocked.Enabled = False
      Else
        chkEnforcePasswordPolicy.Value = vbChecked
        chkEnforceExpiry.Value = vbChecked
        chkForceChange.Value = vbChecked
        chkAccountIsLocked.Enabled = False
      End If
    End If

    rsUserInfo.Close
  Else
  
    IsUserNameInUse cboUser.Text, gObjGroups, mstrSecurityGroupName
    If Len(mstrSecurityGroupName) <> 0 Then
      chkForceChange.Value = IIf(gObjGroups(mstrSecurityGroupName).Users(cboUser.Text).ForcePasswordChange, vbChecked, vbUnchecked)
    End If

  End If

  Set rsUserInfo = Nothing

End Sub

Private Sub txtPassword_Change()
  chkForceChange.Enabled = (chkEnforceExpiry.Value = vbChecked) Or glngSQLVersion < 9
  cmdOK.Enabled = True
End Sub
