VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User Login"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8015
   Icon            =   "frmNewUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmAccount 
      Caption         =   "User Information : "
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   90
      Width           =   4515
      Begin VB.OptionButton optAuthenticationMethod 
         Caption         =   "Windows Authentication"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1215
         Width           =   2445
      End
      Begin VB.ComboBox cboDomain 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1530
         Width           =   2670
      End
      Begin VB.CommandButton cmdGetWindowsLogins 
         Caption         =   "..."
         Height          =   300
         Left            =   3945
         TabIndex        =   8
         Top             =   1950
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton optAuthenticationMethod 
         Caption         =   "SQL Server Authentication"
         Height          =   330
         Index           =   1
         Left            =   270
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   3120
      End
      Begin VB.TextBox txtUserLogin 
         Height          =   315
         Left            =   1590
         TabIndex        =   7
         Top             =   1950
         Width           =   2340
      End
      Begin VB.ComboBox cboUserLogin 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cboUserLogin"
         Top             =   660
         Width           =   2670
      End
      Begin VB.Label lblSQLUser 
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDomain 
         Caption         =   "Domain :"
         Height          =   255
         Left            =   525
         TabIndex        =   4
         Top             =   1605
         Width           =   930
      End
      Begin VB.Label lblWindowsUser 
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Left            =   525
         TabIndex        =   6
         Top             =   1995
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3435
      TabIndex        =   10
      Top             =   2745
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2160
      TabIndex        =   9
      Top             =   2745
      Width           =   1200
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbForcePasswordChange As Boolean
Private mstrDomainName As String
Private mfCancelled As Boolean
Private msUserName As String
Private msUserNamesLoginTypes As String
Private miLoginType As SecurityMgr.LoginType

Public Property Get UserLogin() As String
'  Return the entered user login.

  Dim astrUserNames() As String
  Dim astrLoginTypes() As String
  Dim iCount As Integer
  
  ' Put the domain name into the username string if necessary
  If miLoginType = iUSERTYPE_TRUSTEDUSER And Len(msUserName) > 0 Then
    
    astrUserNames = Split(msUserName, ";")
    astrLoginTypes = Split(msUserNamesLoginTypes, ";")
    ReDim Preserve astrLoginTypes(UBound(astrUserNames))
    
    For iCount = LBound(astrUserNames) To UBound(astrUserNames)
    
      ' If no login type defined set as windows user
      If Len(astrLoginTypes(iCount)) = 0 Then
        astrLoginTypes(iCount) = iUSERTYPE_TRUSTEDUSER
      End If
    
      ' Is this a Windows group instead of a Windows user
      If IsWindowsGroup(astrUserNames(iCount)) Then
        astrLoginTypes(iCount) = iUSERTYPE_TRUSTEDGROUP
      End If
    
      ' Prefix with domain name
      If InStr(1, UCase(astrUserNames(iCount)), UCase(mstrDomainName) & "\") = 0 Then
        'astrUserNames(icount) = UCase(mstrDomainName) & "\" & astrUserNames(icount)
        astrUserNames(iCount) = astrUserNames(iCount)
      End If
        
    Next iCount
  
    UserLogin = Join(astrUserNames, ";")
    msUserNamesLoginTypes = Join(astrLoginTypes, ";")
  Else
    UserLogin = msUserName
  End If

End Property

Public Property Get UserLoginTypes() As String
'  Return the entered user login types.
  UserLoginTypes = msUserNamesLoginTypes
End Property

Private Sub cboDomain_Click()
  RefreshButtons
End Sub

Private Sub cboUserLogin_Change()

  If Len(Me.cboUserLogin.Text) > 50 Then
    Me.cboUserLogin.Text = Left(Me.cboUserLogin.Text, 50)
  End If
  
  RefreshButtons
  
End Sub

Private Sub cboUserLogin_KeyPress(KeyAscii As Integer)

'MsgBox Me.cboUserLogin.SelLength

  'TM20011203 Fault 3241
  If KeyAscii <> vbKeyBack Then
    ' Check that the char is valid
    If (Len(Me.cboUserLogin.Text) >= 50) And (Me.cboUserLogin.SelLength >= 50) Then
      Me.cboUserLogin.Text = vbNullString
      KeyAscii = ValidNameChar(KeyAscii, cboUserLogin.SelStart)
    ElseIf (Len(Me.cboUserLogin.Text) >= 50) Then
      KeyAscii = 0
    Else
      KeyAscii = ValidNameChar(KeyAscii, cboUserLogin.SelStart)
    End If
  End If
  
End Sub

Private Sub cmdCancel_Click()
  ' the user has selected to cancel
  Cancelled = True
  Unload Me

End Sub

Private Sub cmdGetWindowsLogins_Click()

  Dim frmUserList As SecurityMgr.frmNewTrustedUser
  
  Set frmUserList = New SecurityMgr.frmNewTrustedUser
  
  With frmUserList
    .DomainName = cboDomain.Text
    If .Initialise Then
      .Show vbModal
  
      If Not .Cancelled Then
        txtUserLogin.Text = .UsersSelected
        msUserNamesLoginTypes = .UsersSelectedTypes
      End If
      
      RefreshButtons
  
    Else
      Unload frmUserList
      ' AE20080502 Fault #13143
'      MsgBox "Unable to browse the " & Trim(cboDomain.Text) & " domain." & vbCrLf _
'        & "This is mostly likely because the SQL service account does not have sufficient privileges on that domain." & vbCrLf _
'        & "Please contact your system administrator.", vbInformation, Me.Caption
        
      MsgBox "Unable to browse the " & Trim(cboDomain.Text) & " domain." & vbCrLf _
        & "This is mostly likely because the SQL service account does not have sufficient privileges on that domain " _
        & "or the domain is no longer available." & vbCrLf _
        & "Please contact your system administrator.", vbInformation, Me.Caption
    End If
  End With

End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsLogins As New ADODB.Recordset
  Dim strWhereClause As String
  Dim astrWhereClause() As String
  Dim iCount As Integer

  fOK = True
  
  If fOK Then
    ' Check that the current user is 'sa' if a new login is being created.
    'TM20011114 Fault  3125 - retrieve loginname not just name.
    If optAuthenticationMethod(0).Value = True Then
      astrWhereClause = Split(Trim(Replace(txtUserLogin.Text, "'", "''")), ";")
      strWhereClause = ""
      
      For iCount = LBound(astrWhereClause) To UBound(astrWhereClause)
          'strWhereClause = strWhereClause & IIf(LenB(strWhereClause) <> 0, " AND ", "") & " loginname = '" & cboDomain.Text & "\" & astrWhereClause(icount) & "'"
          strWhereClause = strWhereClause & IIf(LenB(strWhereClause) <> 0, " AND ", "") & " loginname = '" & astrWhereClause(iCount) & "'"
      Next iCount
      
      strWhereClause = "WHERE" & strWhereClause
    Else
      strWhereClause = "WHERE loginname = '" & Replace(Trim(cboUserLogin.Text), "'", "''") & "'"
    End If
      
    sSQL = "SELECT loginname " _
      & "FROM master.dbo.syslogins " _
      & strWhereClause
    rsLogins.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    If rsLogins.EOF And rsLogins.BOF And _
      Not gbUserCanManageLogins Then
      
      MsgBox "New logins can only be created by a System or Security Administrator.", vbInformation + vbOKOnly, App.Title
      If cboUserLogin.ListCount > 0 Then
        cboUserLogin.ListIndex = 0
      End If
      cboUserLogin.SetFocus
      fOK = False
    End If
  
    rsLogins.Close
  End If
  
  
  If fOK Then
    ' Check that the user to be created is not 'OpenHR2IIS' as this is a fixed
    ' username created by the 5.1 update script
    If LCase(cboUserLogin.Text) = "openhr2iis" Then
     MsgBox "Cannot create user 'OpenHR2IIS' as this is a reserved user name.", vbInformation + vbOKOnly, App.Title
      If cboUserLogin.ListCount > 0 Then
        cboUserLogin.ListIndex = 0
      End If
      cboUserLogin.SetFocus
      fOK = False
      
    End If
  End If
  
  
TidyUpAndExit:
  Set rsLogins = Nothing
  
  If fOK Then
  
    If optAuthenticationMethod(1).Value = True Then
      msUserName = Trim(cboUserLogin.Text)
    Else
      msUserName = Trim(txtUserLogin.Text)
    End If
    
    mstrDomainName = cboDomain.Text
    
    Cancelled = False
    Unload Me
  End If
  
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  Dim iLoop As Integer
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim astrDomainList() As String
  Dim objNetwork As New SecurityMgr.Net
  Dim rsRecords As New ADODB.Recordset
  Dim iResizeArray As Integer
  
  'ReDim mastrWindowsGroups(0)
  ReDim astrDomainList(0)
            
' Retrieve all logins that do not have any principles in this database
  sSQL = "SELECT l.loginname" & _
    " FROM sys.syslogins l" & _
    " WHERE l.loginname NOT IN (SELECT name COLLATE DATABASE_DEFAULT FROM sys.database_principals WHERE type = 'S')" & _
    " AND l.sysadmin = 0" & _
    " AND LEN(l.loginname) <= 50" & _
    " AND l.IsNTName = 0" & _
    " AND l.loginname NOT LIKE '##MS_%'" & _
    " ORDER BY l.loginname"
    
  FillCombo cboUserLogin, sSQL

  ' Select the first combo item, or disable the control if no items exist.
  If cboUserLogin.ListCount > 0 Then
    cboUserLogin.ListIndex = 0
  End If
     
  
  mbForcePasswordChange = True
  msUserNamesLoginTypes = Str(iUSERTYPE_SQLLOGIN)
  
  ' Set Default authentication method
  If giSQLServerAuthenticationType = iWINDOWSONLY Then
    optAuthenticationMethod(0).Value = True
    optAuthenticationMethod(1).Enabled = False
    optAuthenticationMethod_Click (0)
  Else
    If gbCanUseWindowsAuthentication Then
      If gbUseWindowsAuthentication Then
        optAuthenticationMethod(0).Value = True
        optAuthenticationMethod_Click (0)
      Else
        optAuthenticationMethod(1).Value = True
        optAuthenticationMethod_Click (1)
      End If
    Else
      optAuthenticationMethod(0).Enabled = False
      cboDomain.Enabled = False
      cmdGetWindowsLogins.Enabled = False
      cboDomain.BackColor = vbButtonFace
      lblDomain.ForeColor = vbGrayText
      cboUserLogin.Visible = True
      txtUserLogin.Enabled = False
      txtUserLogin.BackColor = vbButtonFace
      lblWindowsUser.Enabled = False
      
    End If
  End If
  
  RefreshButtons
  
End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
End Property

Public Property Get ForcePasswordChange() As Boolean
  ForcePasswordChange = mbForcePasswordChange
End Property


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancelled = True
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Public Property Get LoginType() As SecurityMgr.LoginType
'  Return the login type.
  LoginType = miLoginType
End Property

Private Sub optAuthenticationMethod_Click(Index As Integer)

  ' Windows authentication / sql authentication
  miLoginType = IIf(optAuthenticationMethod(0).Value, iUSERTYPE_TRUSTEDUSER, iUSERTYPE_SQLLOGIN)
  
  Select Case miLoginType
 
    Case iUSERTYPE_TRUSTEDUSER
      msUserNamesLoginTypes = Str(iUSERTYPE_TRUSTEDUSER)
      cboDomain.Enabled = True
      cmdGetWindowsLogins.Enabled = True
      cboUserLogin.Enabled = False
      txtUserLogin.Enabled = True
    
      ' Populate domain list
      If GetSystemSetting("Misc", "AutoBuildDomainList", 1) = 0 Then
        Screen.MousePointer = vbHourglass
        If cboDomain.Text = "" Then
          FillCombo cboDomain, , GetWindowsDomains
          SetComboText cboDomain, gstrServerDefaultDomain
        End If
        
        'AE20080116 Fault #12704, #13143
        If cboDomain.Text = "" And cboDomain.ListCount > 0 Then
          cboDomain.ListIndex = 0
        End If
        
        If cboDomain.Text <> "" Then
          InitialiseWindowsGroups cboDomain.Text
        End If
      Else
        cboDomain.Enabled = False
        InitialiseWindowsDomains
      End If
    
    Case iUSERTYPE_SQLLOGIN
      msUserNamesLoginTypes = Str(iUSERTYPE_SQLLOGIN)
      cboDomain.Enabled = False
      cmdGetWindowsLogins.Enabled = False
      cboUserLogin.Enabled = True
      txtUserLogin.Enabled = False
      
  End Select

  lblWindowsUser.ForeColor = IIf(miLoginType = iUSERTYPE_TRUSTEDUSER, vbBlack, vbGrayText)
  lblDomain.ForeColor = IIf(miLoginType = iUSERTYPE_TRUSTEDUSER, vbBlack, vbGrayText)
  txtUserLogin.BackColor = IIf(miLoginType = iUSERTYPE_TRUSTEDUSER, vbWhite, vbButtonFace)
  cboDomain.BackColor = IIf(miLoginType = iUSERTYPE_TRUSTEDUSER, vbWhite, vbButtonFace)
  lblSQLUser.ForeColor = IIf(miLoginType = iUSERTYPE_SQLLOGIN, vbBlack, vbGrayText)
  cboUserLogin.BackColor = IIf(miLoginType = iUSERTYPE_SQLLOGIN, vbWhite, vbButtonFace)
  cboDomain.BackColor = IIf(cboDomain.Enabled, vbWhite, vbButtonFace)

  RefreshButtons
  Screen.MousePointer = vbDefault

End Sub

Private Sub RefreshButtons()

  Dim bBypassPolicy As Boolean

  Select Case miLoginType
 
    Case iUSERTYPE_TRUSTEDUSER
      cmdOK.Enabled = (Len(txtUserLogin.Text) > 0)
      cmdGetWindowsLogins.Enabled = (cboDomain.Text <> "")
    
    Case iUSERTYPE_SQLLOGIN
      cmdOK.Enabled = (Len(cboUserLogin.Text) > 0)
    
  End Select

End Sub

Private Sub txtUserLogin_Change()
  RefreshButtons
End Sub



