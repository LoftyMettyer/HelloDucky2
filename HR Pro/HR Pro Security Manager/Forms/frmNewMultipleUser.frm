VERSION 5.00
Begin VB.Form frmNewMultipleUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatic User Add"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1029
   Icon            =   "frmNewMultipleUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmGroup 
      Caption         =   "Group :"
      Height          =   765
      Left            =   90
      TabIndex        =   24
      Top             =   90
      Width           =   5025
      Begin VB.ComboBox cboSecurityGroups 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   3720
      End
      Begin VB.Label lblSecurityGroup 
         Caption         =   "Name :"
         Height          =   315
         Left            =   255
         TabIndex        =   0
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.Frame frmUserInfo 
      Caption         =   "User Information :"
      Height          =   2910
      Left            =   90
      TabIndex        =   2
      Top             =   915
      Width           =   5025
      Begin VB.ComboBox cboDomain 
         Height          =   315
         Left            =   1740
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1995
         Width           =   3120
      End
      Begin VB.CommandButton cmdWindowsUsername 
         Height          =   315
         Left            =   4530
         Picture         =   "frmNewMultipleUser.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtWindowsUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         TabIndex        =   14
         Top             =   2400
         Width           =   2790
      End
      Begin VB.OptionButton optAuthenticationMethod 
         Caption         =   "SQL Server Authentication"
         Height          =   330
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Value           =   -1  'True
         Width           =   3120
      End
      Begin VB.OptionButton optAuthenticationMethod 
         Caption         =   "Windows Authentication"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   1680
         Width           =   2655
      End
      Begin VB.OptionButton optAuthenticationMethod 
         Caption         =   "Windows Domain (Super Very Clever Mode)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   23
         Top             =   2955
         Width           =   3705
      End
      Begin VB.CommandButton cmdPassword 
         Height          =   315
         Left            =   4530
         Picture         =   "frmNewMultipleUser.frx":0084
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1125
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmdSQLUserName 
         Height          =   315
         Left            =   4530
         Picture         =   "frmNewMultipleUser.frx":00FC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1755
         TabIndex        =   9
         Top             =   1110
         Width           =   2775
      End
      Begin VB.TextBox txtSQLUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1755
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblDomain 
         Caption         =   "Domain :"
         Height          =   255
         Left            =   615
         TabIndex        =   11
         Top             =   2055
         Width           =   930
      End
      Begin VB.Label lblWindowsUserName 
         Caption         =   "User :"
         Height          =   315
         Left            =   615
         TabIndex        =   13
         Top             =   2445
         Width           =   960
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password :"
         Height          =   285
         Left            =   615
         TabIndex        =   8
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label lblSQLUserName 
         Caption         =   "User name :"
         Height          =   315
         Left            =   615
         TabIndex        =   5
         Top             =   765
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   405
      Left            =   2625
      TabIndex        =   21
      Top             =   5265
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3885
      TabIndex        =   22
      Top             =   5265
      Width           =   1200
   End
   Begin VB.Frame frmOptions 
      Caption         =   "Options :"
      Height          =   1215
      Left            =   90
      TabIndex        =   16
      Top             =   3930
      Width           =   5025
      Begin VB.TextBox txtFilter 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1185
         TabIndex        =   18
         Top             =   285
         Width           =   3345
      End
      Begin VB.CheckBox chkChangePassword 
         Caption         =   "&User must change password at next login"
         Height          =   270
         Left            =   285
         TabIndex        =   20
         Top             =   765
         Value           =   1  'Checked
         Width           =   3870
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   315
         Left            =   4530
         Picture         =   "frmNewMultipleUser.frx":0174
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter :"
         Height          =   285
         Left            =   300
         TabIndex        =   17
         Top             =   345
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmNewMultipleUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFilterExprID As Long
Private mlngSQLUserNameExprID As Long
Private mlngWindowsUserNameExprID As Long
Private mlngPasswordExprID As Long
Private mfCancelled As Boolean
Private mstrSecurityGroupName As String
Private mbForcePasswordChange As Boolean
Private miAuthenticationMode As HrProSecurityMgr.CreateUserMode
Private mstrDomainName As String
Private mbDisableDomainBrowsing As Boolean

Public Property Get ForceChangePassword() As Boolean
  ForceChangePassword = mbForcePasswordChange
End Property

Public Property Get ID_WindowsUserNameExpr() As Long
  ID_WindowsUserNameExpr = mlngWindowsUserNameExprID
End Property

Public Property Get ID_SQLUserNameExpr() As Long
  ID_SQLUserNameExpr = mlngSQLUserNameExprID
End Property

Public Property Get ID_PasswordExpr() As Long
  ID_PasswordExpr = mlngPasswordExprID
End Property

Public Property Get ID_FilterExpr() As Long
  ID_FilterExpr = mlngFilterExprID
End Property

Public Property Get SecurityGroupName() As String
  SecurityGroupName = mstrSecurityGroupName
End Property

Public Property Let SecurityGroupName(ByVal pstrNewValue As String)
  mstrSecurityGroupName = pstrNewValue
  
  ' Load the saved expression IDs for this security group
  miAuthenticationMode = GetSystemSetting("AutomaticAdd_AuthenticationMode", mstrSecurityGroupName, iUSERCREATE_SQLLOGIN)
  mlngSQLUserNameExprID = GetSystemSetting("AutomaticAdd_UserNameExprID", mstrSecurityGroupName, 0)
  mlngPasswordExprID = GetSystemSetting("AutomaticAdd_PasswordExprID", mstrSecurityGroupName, 0)
  mlngFilterExprID = GetSystemSetting("AutomaticAdd_FilterExprID", mstrSecurityGroupName, 0)
  mlngWindowsUserNameExprID = GetSystemSetting("AutomaticAdd_UserWindowsName", mstrSecurityGroupName, 0)
  mbForcePasswordChange = GetSystemSetting("AutomaticAdd_ChangePassword", mstrSecurityGroupName, True)
  mstrDomainName = GetSystemSetting("AutomaticAdd_DomainName", mstrSecurityGroupName, gstrServerDefaultDomain)
  
End Property

Private Sub cboDomain_Click()
  mstrDomainName = cboDomain.Text
End Sub

Private Sub chkChangePassword_Click()
  mbForcePasswordChange = chkChangePassword.Value
End Sub

Private Sub cmdCancel_Click()

  ' the user has selected to cancel
  Cancelled = True
  Unload Me

End Sub

Private Sub cmdOK_Click()

  Dim bOK As Boolean
  
  'AE20080917 Fault #13372
  If glngSQLVersion > 8 Then
    mstrDomainName = GetDomainFromFQDN(mstrDomainName)
  End If
  
  If Not gbUserCanManageLogins Then
    MsgBox "New users can only be created by a System Administrator.", vbInformation + vbOKOnly, App.Title
    Cancelled = True
  Else
    ' Save the entered expression IDs for this security group
    SaveSystemSetting "AutomaticAdd_AuthenticationMode", mstrSecurityGroupName, miAuthenticationMode
    SaveSystemSetting "AutomaticAdd_UserNameExprID", mstrSecurityGroupName, mlngSQLUserNameExprID
    SaveSystemSetting "AutomaticAdd_PasswordExprID", mstrSecurityGroupName, mlngPasswordExprID
    SaveSystemSetting "AutomaticAdd_FilterExprID", mstrSecurityGroupName, mlngFilterExprID
    SaveSystemSetting "AutomaticAdd_UserWindowsName", mstrSecurityGroupName, mlngWindowsUserNameExprID
    SaveSystemSetting "AutomaticAdd_ChangePassword", mstrSecurityGroupName, mbForcePasswordChange
    SaveSystemSetting "AutomaticAdd_DomainName", mstrSecurityGroupName, mstrDomainName
    Cancelled = False
  End If
  
  Unload Me

End Sub

Private Sub cmdPassword_Click()

  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
  Set objExpr = New clsExprExpression

  With objExpr
    If .Initialise(glngPersonnelTableID, mlngPasswordExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
      'NHRD11032004 Fault 6633 temporarily hide the frmNewMultipleUser screen to avoid
      'ghosting image when editing the resulting definition.  Used on the other buttons too.
'      Me.Hide
      If .SelectExpression(False, lngOptions) Then
        mlngPasswordExprID = .ExpressionID
        txtPassword.Text = .Name
      Else
        If .ActionType = edtDeselect Then
          mlngPasswordExprID = 0
          txtPassword.Text = ""
        Else
          'JPD 20040225 Fault 8097
          mlngPasswordExprID = .ExpressionID
          txtPassword.Text = .Name
        End If
      End If
    End If
'    Me.Show vbModal
  End With

  Set objExpr = Nothing
  
  ' Refresh the buttons
  RefreshButtons

End Sub

Private Sub cmdFilter_Click()

  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
  Set objExpr = New clsExprExpression

  With objExpr
    If .Initialise(glngPersonnelTableID, mlngFilterExprID, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
'      Me.Hide
      If .SelectExpression(False, lngOptions) Then
        mlngFilterExprID = .ExpressionID
        txtFilter.Text = .Name
      Else
        If .ActionType = edtDeselect Then
          mlngFilterExprID = 0
          txtFilter.Text = ""
        Else
          'JPD 20040225 Fault 8097
          mlngFilterExprID = .ExpressionID
          txtFilter.Text = .Name
        End If
      End If
    End If
'    Me.Show vbModal
  End With

  Set objExpr = Nothing

  ' Refresh the buttons
  RefreshButtons
  
End Sub

Private Sub cmdSQLUserName_Click()

  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
  Set objExpr = New clsExprExpression

  With objExpr
    If .Initialise(glngPersonnelTableID, mlngSQLUserNameExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
'      Me.Hide
      If .SelectExpression(False, lngOptions) Then
        mlngSQLUserNameExprID = .ExpressionID
        txtSQLUserName.Text = .Name
      Else
        If .ActionType = edtDeselect Then
          mlngSQLUserNameExprID = 0
          txtSQLUserName.Text = ""
        Else
          'JPD 20040225 Fault 8097
          mlngSQLUserNameExprID = .ExpressionID
          
          txtSQLUserName.Text = .Name
        End If
      End If
    End If
'    Me.Show vbModal
  End With

  Set objExpr = Nothing

  ' Refresh the buttons
  RefreshButtons

End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Public Property Let Cancelled(ByVal pbNewValue As Boolean)
  mfCancelled = pbNewValue
End Property


Private Sub PopulateSecurityGroupsCombo()

  Dim objSecurityGroup As SecurityGroup
  
  cboSecurityGroups.Clear
  For Each objSecurityGroup In gObjGroups
    cboSecurityGroups.AddItem objSecurityGroup.Name
  Next objSecurityGroup
  
  ' Set the passed in security group
  If Len(mstrSecurityGroupName) > 0 Then
    cboSecurityGroups.Text = mstrSecurityGroupName
  End If

End Sub

Private Sub cmdWindowsUsername_Click()

  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
  Set objExpr = New clsExprExpression

  With objExpr
    If .Initialise(glngPersonnelTableID, mlngWindowsUserNameExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
'      Me.Hide
      If .SelectExpression(False, lngOptions) Then
        mlngWindowsUserNameExprID = .ExpressionID
        txtWindowsUserName.Text = .Name
      Else
        If .ActionType = edtDeselect Then
          mlngWindowsUserNameExprID = 0
          txtWindowsUserName.Text = ""
        Else
          'JPD 20040225 Fault 8097
          mlngWindowsUserNameExprID = .ExpressionID
          
          txtWindowsUserName.Text = .Name
        End If
      End If
    End If
'    Me.Show
  End With

  Set objExpr = Nothing

  ' Refresh the buttons
  RefreshButtons

End Sub

Private Sub Form_Load()
  Dim iLoop As Integer
  Dim sSQL As String
  Dim strDefaultDomainName As String
  Dim rsRecords As New ADODB.Recordset
  Dim astrDomainList() As String
  Dim iResizeArray As Integer
  
  ReDim astrDomainList(0)

  ' Default exit as being cancelled
  Me.Cancelled = True

  mbDisableDomainBrowsing = GetSystemSetting("Misc", "AutoBuildDomainList", 0)

  ' If Windows only allowed
  If giSQLServerAuthenticationType = iWINDOWSONLY Then
    optAuthenticationMethod(0).Enabled = False
    optAuthenticationMethod(1).Value = True
  End If
 
  ' Is Windows Authentication disabled
  If Not gbCanUseWindowsAuthentication Then
    optAuthenticationMethod(1).Enabled = False
    lblDomain.Enabled = False
    cboDomain.Enabled = False
    cboDomain.BackColor = vbButtonFace
    lblWindowsUserName.Enabled = False
    txtWindowsUserName.Text = ""
    cmdWindowsUsername.Enabled = False
  End If
 
  ' If user logged in using authentication select that as the default
  If Not miAuthenticationMode = iUSERCREATE_SQLLOGIN Then
    optAuthenticationMethod(1).Value = True
  End If
 
  ' Populate domain list
  If Not mbDisableDomainBrowsing Then
    If cboDomain.Text = "" Then
      FillCombo cboDomain, , GetWindowsDomains
      SetComboText cboDomain, gstrServerDefaultDomain
    End If
        
    ' AE20080507 Fault #13150
    If cboDomain.Text = "" And cboDomain.ListCount > 0 Then
      cboDomain.ListIndex = 0
    End If
    
    mstrDomainName = cboDomain.Text
  Else
    InitialiseWindowsDomains
    mstrDomainName = ""
    cboDomain.Enabled = False
    cboDomain.BackColor = IIf(cboDomain.Enabled, vbWhite, vbButtonFace)
  End If

  'Populate the display
  PopulateSecurityGroupsCombo
  PopulateExpressions
  
  ' Update the password change option
  chkChangePassword_Click
  
  ' Refresh the buttons
  RefreshButtons
    
End Sub

Private Sub PopulateExpressions()

  Dim objExpr As New clsExprExpression
  Set objExpr = New clsExprExpression

  ' Update the Username textbox
  If objExpr.Initialise(glngPersonnelTableID, mlngSQLUserNameExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
    txtSQLUserName.Text = objExpr.Name
  End If

  If objExpr.Initialise(glngPersonnelTableID, mlngPasswordExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
    txtPassword.Text = objExpr.Name
  End If
  
  If objExpr.Initialise(glngPersonnelTableID, mlngFilterExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
    txtFilter.Text = objExpr.Name
  End If

  If objExpr.Initialise(glngPersonnelTableID, mlngWindowsUserNameExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER) Then
    txtWindowsUserName.Text = objExpr.Name
  End If

  Set objExpr = Nothing

End Sub

Private Sub RefreshButtons()
  
  Dim bBypassPolicy As Boolean
  
  ' Only enable OK button if username and password expressions are supplied
  bBypassPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)  ' Default - Off
  
  ' SQL Authentication
  If optAuthenticationMethod.Item(0).Value = True Then
    
    lblSQLUserName.Enabled = True
    cmdSQLUserName.Enabled = True
    lblPassword.Enabled = True
    cmdPassword.Enabled = True
    
    mlngWindowsUserNameExprID = 0
    
    lblDomain.Enabled = False
    cboDomain.Enabled = False
    cboDomain.BackColor = vbButtonFace
    lblWindowsUserName.Enabled = False
    txtWindowsUserName.Text = ""
    cmdWindowsUsername.Enabled = False
    If bBypassPolicy And glngSQLVersion >= 9 Then
      chkChangePassword.Value = vbUnchecked
      chkChangePassword.Enabled = False
    Else
      chkChangePassword.Value = vbChecked
      chkChangePassword.Enabled = True
    End If
    cmdOK.Enabled = mlngSQLUserNameExprID > 0 And mlngPasswordExprID > 0
  End If
  
  ' Windows Authentication  (Manual Username)
  If optAuthenticationMethod.Item(1).Value = True Then
    
    lblSQLUserName.Enabled = False
    txtSQLUserName.Text = ""
    cmdSQLUserName.Enabled = False
    lblPassword.Enabled = False
    txtPassword.Text = ""
    cmdPassword.Enabled = False
    
    mlngSQLUserNameExprID = 0
    mlngPasswordExprID = 0
    
    lblDomain.Enabled = True
    cboDomain.Enabled = Not mbDisableDomainBrowsing
    cboDomain.BackColor = IIf(cboDomain.Enabled, vbWhite, vbButtonFace)
    lblWindowsUserName.Enabled = True
    cmdWindowsUsername.Enabled = True
      
    chkChangePassword.Value = vbUnchecked
    chkChangePassword.Enabled = False
    cmdOK.Enabled = mlngWindowsUserNameExprID > 0
  End If
     
  ' Windows Authentication  (Auto map to username)
  If optAuthenticationMethod.Item(2).Value = True Then
    
    lblSQLUserName.Enabled = False
    txtSQLUserName.Text = ""
    cmdSQLUserName.Enabled = False
    lblPassword.Enabled = False
    cmdPassword.Enabled = False
    
    mlngSQLUserNameExprID = 0
    mlngPasswordExprID = 0
    mlngWindowsUserNameExprID = 0
    
    lblDomain.Enabled = False
    cboDomain.Enabled = Not mbDisableDomainBrowsing
    cboDomain.BackColor = IIf(cboDomain.Enabled, vbWhite, vbButtonFace)
    lblWindowsUserName.Enabled = False
    txtWindowsUserName.Text = ""
    cmdWindowsUsername.Enabled = False
    
    cmdOK.Enabled = True
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optAuthenticationMethod_Click(Index As Integer)
  
  miAuthenticationMode = Index
  
  RefreshButtons

End Sub

' Return the create mode
Public Property Get CreateUserMode() As HrProSecurityMgr.CreateUserMode
  CreateUserMode = miAuthenticationMode
End Property

Private Sub ToggleFrameControls(pfrmFrame As Frame, pbEnabled As Boolean)

  'Disable all controls on a frame
  On Local Error Resume Next

  Dim ctlTemp As Control
  For Each ctlTemp In Me.Controls
    If ctlTemp.Container = pfrmFrame Then
    
'    If Not (TypeOf ctlTemp Is Label) And _
'       Not (TypeOf ctlTemp Is Frame) And _
'       Not (TypeOf ctlTemp Is TabStrip) Then
          ctlTemp.Enabled = False
          ctlTemp.BackColor = vbButtonFace
    End If
  Next

End Sub

' Return the selected domain name
Public Property Get DomainName() As String
  DomainName = mstrDomainName
End Property
