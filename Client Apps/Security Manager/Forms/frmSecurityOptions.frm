VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmSecurityOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security Options"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8036
   Icon            =   "frmSecurityOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5625
      TabIndex        =   2
      Top             =   4590
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4380
      TabIndex        =   1
      Top             =   4590
      Width           =   1200
   End
   Begin TabDlg.SSTab sstabOptions 
      Height          =   4425
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   7805
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Security Policy"
      TabPicture(0)   =   "frmSecurityOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPolicy"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Advanced"
      TabPicture(1)   =   "frmSecurityOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOrphanedAccounts"
      Tab(1).Control(1)=   "frmWindowsAuth"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Login Maintenance"
      TabPicture(2)   =   "frmSecurityOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSelfServiceLogin"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraSelfServiceLogin 
         Caption         =   "Self Service Logins :"
         Height          =   3825
         Left            =   -74850
         TabIndex        =   32
         Top             =   405
         Width           =   6375
         Begin VB.CheckBox chkEmailWorkAddress 
            Caption         =   "Send login details to wor&k email address"
            Height          =   225
            Left            =   135
            TabIndex        =   31
            Top             =   1815
            Width           =   3975
         End
         Begin VB.CheckBox chkDisableLoginsOnLeaveDate 
            Caption         =   "Disable logins on leavin&g date"
            Height          =   255
            Left            =   135
            TabIndex        =   30
            Top             =   1245
            Width           =   3210
         End
         Begin VB.CheckBox chkAutoAddFromSelfService 
            Caption         =   "Automaticall&y add logins for self service column"
            Height          =   435
            Left            =   135
            TabIndex        =   29
            Top             =   315
            Width           =   4575
         End
         Begin VB.ComboBox cboAutoAddSelfServiceGroup 
            Height          =   315
            ItemData        =   "frmSecurityOptions.frx":0060
            Left            =   2835
            List            =   "frmSecurityOptions.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   795
            Width           =   2805
         End
         Begin VB.Label Label2 
            Caption         =   $"frmSecurityOptions.frx":0064
            ForeColor       =   &H000000FF&
            Height          =   900
            Left            =   405
            TabIndex        =   35
            Top             =   2175
            Width           =   5625
         End
         Begin VB.Label Label1 
            Caption         =   "Default Security Group : "
            Height          =   270
            Left            =   405
            TabIndex        =   34
            Top             =   840
            Width           =   2220
         End
      End
      Begin VB.Frame frmWindowsAuth 
         Caption         =   "Windows Authentication : "
         Height          =   1830
         Left            =   -74850
         TabIndex        =   23
         Top             =   2400
         Width           =   6375
         Begin VB.CheckBox chkDisableDomainListBuilder 
            Caption         =   "Disable Domain Bro&wsing"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   1440
            Width           =   2715
         End
         Begin VB.CheckBox chkDelOrphanedLogins 
            Caption         =   "&Delete Orphan Logins on save"
            Height          =   270
            Left            =   180
            TabIndex        =   17
            Top             =   300
            Width           =   3900
         End
         Begin VB.Label lblDeleteOrphanWarning 
            Caption         =   $"frmSecurityOptions.frx":0142
            ForeColor       =   &H000000FF&
            Height          =   600
            Left            =   435
            TabIndex        =   24
            Top             =   645
            Width           =   5505
         End
      End
      Begin VB.Frame fraOrphanedAccounts 
         Caption         =   "Options : "
         Height          =   1830
         Left            =   -74850
         TabIndex        =   19
         Top             =   405
         Width           =   6375
         Begin VB.CheckBox chkSecManBypass 
            Caption         =   "&Bypass Domain policy when creating new users"
            Height          =   255
            Left            =   195
            TabIndex        =   16
            Top             =   330
            Width           =   4500
         End
         Begin VB.Label lblBypassWarning 
            Caption         =   $"frmSecurityOptions.frx":01EE
            ForeColor       =   &H000000FF&
            Height          =   900
            Left            =   465
            TabIndex        =   25
            Top             =   690
            Width           =   5505
         End
      End
      Begin VB.Frame fraPolicy 
         Caption         =   "Policy :"
         Height          =   3825
         Left            =   150
         TabIndex        =   18
         Top             =   405
         Width           =   6375
         Begin VB.TextBox txtPasswordsRemembered 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2955
            MaxLength       =   4
            TabIndex        =   27
            Top             =   1470
            Width           =   405
         End
         Begin VB.CheckBox chkPasswordsRemembered 
            Caption         =   "Password History :                       remembered"
            Enabled         =   0   'False
            Height          =   240
            Left            =   195
            TabIndex        =   26
            Top             =   1515
            Width           =   4755
         End
         Begin VB.ComboBox cboMinPasswordAge 
            Height          =   315
            Left            =   3390
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   660
            Width           =   1185
         End
         Begin VB.TextBox txtMimimumPasswordAge 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2955
            TabIndex        =   6
            Top             =   660
            Width           =   420
         End
         Begin VB.CheckBox chkMinimumPasswordAge 
            Caption         =   "Minimum password age :"
            Height          =   240
            Left            =   195
            TabIndex        =   5
            Top             =   735
            Width           =   3060
         End
         Begin VB.CheckBox chkPCLockout 
            Caption         =   "Enable Lockout"
            Height          =   255
            Left            =   195
            TabIndex        =   12
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txtPasswordLength 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2955
            MaxLength       =   4
            TabIndex        =   4
            Top             =   270
            Width           =   405
         End
         Begin VB.CheckBox chkComplexity 
            Caption         =   "Passwords must meet complexity requirements"
            Enabled         =   0   'False
            Height          =   300
            Left            =   195
            TabIndex        =   11
            Top             =   1875
            Width           =   4470
         End
         Begin VB.ComboBox cboChangePeriod 
            Height          =   315
            Left            =   3390
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1065
            Width           =   1185
         End
         Begin VB.TextBox txtChangeFrequency 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2955
            MaxLength       =   3
            TabIndex        =   9
            Top             =   1065
            Width           =   420
         End
         Begin VB.CheckBox chkMinimumPasswordLength 
            Caption         =   "Minimum password length : "
            Height          =   255
            Left            =   195
            TabIndex        =   3
            Top             =   330
            Width           =   3165
         End
         Begin VB.CheckBox chkMaximumPasswordAge 
            Caption         =   "Maximum password age : "
            Height          =   270
            Left            =   195
            TabIndex        =   8
            Top             =   1110
            Width           =   2550
         End
         Begin COASpinner.COA_Spinner spnAttempts 
            Height          =   315
            Left            =   2895
            TabIndex        =   13
            Top             =   2610
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
         Begin COASpinner.COA_Spinner spnLockoutDuration 
            Height          =   315
            Left            =   2895
            TabIndex        =   15
            Top             =   3390
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
         Begin COASpinner.COA_Spinner spnResetTime 
            Height          =   315
            Left            =   2895
            TabIndex        =   14
            Top             =   3000
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
         Begin VB.Label lblResetTime 
            Caption         =   "Counter reset (minutes)"
            Height          =   255
            Left            =   450
            TabIndex        =   22
            Top             =   3060
            Width           =   2265
         End
         Begin VB.Label lblLockoutDuration 
            Caption         =   "Lockout duration (minutes)"
            Height          =   255
            Left            =   450
            TabIndex        =   21
            Top             =   3450
            Width           =   2370
         End
         Begin VB.Label lblLockoutAfter 
            Caption         =   "Failed logon attempts"
            Height          =   255
            Left            =   450
            TabIndex        =   20
            Top             =   2670
            Width           =   2235
         End
      End
   End
End
Attribute VB_Name = "frmSecurityOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnChanged As Boolean
Private mblnReadOnly As Boolean

Public Property Let Changed(pblnNewValue As Boolean)
  mblnChanged = pblnNewValue
  cmdOK.Enabled = pblnNewValue
End Property


Public Function Initialise() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim pblnOK As Boolean
  
  Screen.MousePointer = vbHourglass
  
  pblnOK = True
  
  mblnReadOnly = (Application.AccessMode <> accFull)
  
  LoadCombos
  
  If pblnOK Then
    pblnOK = LoadSecuritySettings()
  End If
    
  RefreshButtons
   
  lblBypassWarning.Visible = chkSecManBypass.Enabled
  lblDeleteOrphanWarning.Visible = chkDelOrphanedLogins.Enabled
   
   
  Changed = False
  
  Initialise = pblnOK

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Exit Function
  
ErrorTrap:
  MsgBox "Error initialising the security options form." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  Initialise = False
  GoTo TidyUpAndExit
  
End Function
Private Function LoadSecuritySettings() As Boolean

  On Error GoTo ErrorTrap
  
  Dim blnPCLockout As Boolean
  Dim intAttempts As Integer
  Dim lngResetTime As Long
  Dim lngLockoutDuration As Long
  Dim lMinimumLength As Long      ' The minimum length for passwords
  Dim lMinPasswordAge As Long
  Dim lChangeFrequency As Long    ' How often passwords must be changed
  Dim sChangePeriod As String     ' How often passwords must be changed
  Dim cmdPolicy As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim iResult As Integer
  Dim iComplexity As Integer

  ' Get the policy settings from the domain
  If glngSQLVersion > 8 Then
  
    blnPCLockout = gblnDomainPCLockout
    intAttempts = gintDomainAttempts
    lngResetTime = glngDomainResetTime
    lngLockoutDuration = glngDomainLockoutDuration
    lMinimumLength = glngDomainMinimumLength
    lChangeFrequency = glngDomainChangeFrequency
    sChangePeriod = gstrDomainChangePeriod
    iComplexity = giDomainComplexity
    lMinPasswordAge = glngDomainMinPasswordAge
  
  Else
  
    ' Password options
    lMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)
    lMinPasswordAge = GetSystemSetting("Password", "Minimum Age", 0)
    lChangeFrequency = GetSystemSetting("Password", "Change Frequency", 0)
    sChangePeriod = GetSystemSetting("Password", "Change Period", "")
    iComplexity = GetSystemSetting("Password", "Use Complexity", 0)
    
     ' Lockout settings
    intAttempts = GetSystemSetting("Misc", "CFG_BA", 3)
    blnPCLockout = GetSystemSetting("Misc", "CFG_PCL", True)
    lngResetTime = GetSystemSetting("Misc", "CFG_RT", 3600)
    lngLockoutDuration = GetSystemSetting("Misc", "CFG_LD", 300)
    
  End If
  
  'Set the lockout option
  If blnPCLockout Then
    chkPCLockout.Value = vbChecked
    spnAttempts.Value = intAttempts
    spnResetTime.Value = (lngResetTime / 60)
    spnLockoutDuration.Value = (lngLockoutDuration / 60)
  Else
    chkPCLockout.Value = vbUnchecked
    spnAttempts.Value = 0
    spnResetTime.Value = 0
    spnLockoutDuration.Value = 0
  End If
  
  
  ' Set the password length options
  chkMinimumPasswordLength.Value = IIf(lMinimumLength > 0, vbChecked, vbUnchecked)
  txtPasswordLength.Text = CStr(lMinimumLength)
  chkMinimumPasswordLength_Click
  
  ' Minimum password age
  chkMinimumPasswordAge.Value = IIf(lMinPasswordAge > 0, vbChecked, vbUnchecked)
  txtMimimumPasswordAge.Text = CStr(lMinPasswordAge)
  chkMinimumPasswordAge_Click
  
  ' Complexity settings
  chkComplexity.Value = IIf(iComplexity Mod 2 = 1, vbChecked, vbUnchecked)
  
  ' Set the password changing options
  If lChangeFrequency = 0 Or sChangePeriod = vbNullString Then
    chkMaximumPasswordAge.Value = vbUnchecked
    chkMaximumPasswordAge_Click
  Else
    chkMaximumPasswordAge.Value = vbChecked
    txtChangeFrequency.Text = CStr(lChangeFrequency)
    
    Select Case sChangePeriod
    Case "W": SetComboText Me.cboChangePeriod, "Week(s)"
    Case "M": SetComboText Me.cboChangePeriod, "Month(s)"
    Case "Y": SetComboText Me.cboChangePeriod, "Year(s)"
    Case Else: SetComboText Me.cboChangePeriod, "Day(s)"
    End Select
  End If
  
  ' Password history
  chkPasswordsRemembered.Value = IIf(giDomainPasswordsRemembered > 0, vbChecked, vbUnchecked)
  txtPasswordsRemembered.Text = CStr(giDomainPasswordsRemembered)
  chkPasswordsRemembered_Click
  
  ' Options
  chkSecManBypass.Value = GetSystemSetting("Policy", "Sec Man Bypass", 0)                 ' Default - Off
  chkDelOrphanedLogins.Value = IIf(gbDeleteOrphanWindowsLogins, vbChecked, vbUnchecked)   ' Default - Off
  chkDisableDomainListBuilder.Value = GetSystemSetting("Misc", "AutoBuildDomainList", 0)  ' Default - Off
  
  If Not gbUserCanManageLogins Then
    chkDelOrphanedLogins.Enabled = False
    'AE20071221 Fault #12725
    chkDisableDomainListBuilder.Enabled = False
  End If
  
  ' Login maintenance
  chkAutoAddFromSelfService.Value = IIf(gbLoginMaintAutoAdd, vbChecked, vbUnchecked)
  SetComboText cboAutoAddSelfServiceGroup, gstrLoginMaintAutoAddGroup
  chkDisableLoginsOnLeaveDate.Value = IIf(gbLoginMaintDisableOnLeave, vbChecked, vbUnchecked)
  chkEmailWorkAddress.Value = IIf(gbLoginMaintSendEmail, vbChecked, vbUnchecked)
  
  RefreshButtons
  LoadSecuritySettings = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error initialising the Security Settings form." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  LoadSecuritySettings = False
  GoTo TidyUpAndExit
  
End Function

Private Function SaveSecuritySettings() As Boolean

  On Error GoTo ErrorTrap

  If glngSQLVersion < 9 Then
    If (chkPCLockout.Value = vbUnchecked) Then
      'LC_SaveSecuritySettings False, 1, 1
      SaveSystemSetting "Misc", "CFG_PCL", "0"
      SaveSystemSetting "Misc", "CFG_BA", "1"
      SaveSystemSetting "Misc", "CFG_RT", "1"
      SaveSystemSetting "Misc", "CFG_LD", "1"
    Else
      'LC_SaveSecuritySettings True, spnAttempts.Value, spnLockoutDuration.Value
      SaveSystemSetting "Misc", "CFG_PCL", "1"
      SaveSystemSetting "Misc", "CFG_BA", CStr(spnAttempts.Value)
      SaveSystemSetting "Misc", "CFG_RT", CStr((spnResetTime.Value * 60))
      SaveSystemSetting "Misc", "CFG_LD", CStr((spnLockoutDuration.Value * 60))
    End If
  
    If chkMinimumPasswordLength.Value = vbChecked Then
      SaveSystemSetting "Password", "Minimum Length", txtPasswordLength.Text
    Else
      SaveSystemSetting "Password", "Minimum Length", 0
    End If
    
  
    ' Password Options
    If chkMaximumPasswordAge.Value = vbChecked Then
      SaveSystemSetting "Password", "Change Frequency", txtChangeFrequency.Text
      SaveSystemSetting "Password", "Change Period", Left(UCase(Me.cboChangePeriod.Text), 1)
    Else
      SaveSystemSetting "Password", "Change Frequency", 0
      SaveSystemSetting "Password", "Change Period", ""
    End If
  
    ' Re-initialse the policy settings for rest of application
    glngDomainMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)
    glngDomainMinPasswordAge = GetSystemSetting("Password", "Minimum Age", 0)
    glngDomainChangeFrequency = GetSystemSetting("Password", "Change Frequency", 0)
    gstrDomainChangePeriod = GetSystemSetting("Password", "Change Period", "")
    giDomainComplexity = GetSystemSetting("Password", "Use Complexity", 0)
    gintDomainAttempts = GetSystemSetting("Misc", "CFG_BA", 3)
    gblnDomainPCLockout = GetSystemSetting("Misc", "CFG_PCL", True)
    glngDomainResetTime = GetSystemSetting("Misc", "CFG_RT", 3600)
    glngDomainLockoutDuration = GetSystemSetting("Misc", "CFG_LD", 300)
  
    SaveSystemSetting "Password", "Use Complexity", IIf(chkComplexity.Value = vbChecked, 1, 0)
  End If

  ' Options
  SaveSystemSetting "Policy", "Sec Man Bypass", IIf(chkSecManBypass.Value = vbChecked, 1, 0)
  SaveSystemSetting "Misc", "AutoBuildDomainList", IIf(chkDisableDomainListBuilder.Value = vbChecked, 1, 0)
   
  gbDeleteOrphanWindowsLogins = IIf(chkDelOrphanedLogins.Value = vbChecked, True, False)
  SaveSystemSetting "Misc", "CFG_DELETEORPHANLOGINS", gbDeleteOrphanWindowsLogins
  SaveSystemSetting "Misc", "CFG_DELETEORPHANUSERS", gbDeleteOrphanUsers

  ' Login Maintenance
  gbLoginMaintAutoAdd = IIf(chkAutoAddFromSelfService.Value = vbChecked, True, False)
  SaveSystemSetting "LoginMaintenance", "AUTOADD", gbLoginMaintAutoAdd
  
  gstrLoginMaintAutoAddGroup = cboAutoAddSelfServiceGroup.Text
  SaveSystemSetting "LoginMaintenance", "AUTOADDGROUP", gstrLoginMaintAutoAddGroup
  
  gbLoginMaintDisableOnLeave = IIf(chkDisableLoginsOnLeaveDate.Value = vbChecked, True, False)
  SaveSystemSetting "LoginMaintenance", "DISABLEONLEAVE", gbLoginMaintDisableOnLeave
  
  gbLoginMaintSendEmail = IIf(chkEmailWorkAddress.Value = vbChecked, True, False)
  SaveSystemSetting "LoginMaintenance", "SENDEMAIL", gbLoginMaintSendEmail

  ApplyChanges_LoginMaintenance

  SaveSecuritySettings = True
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  MsgBox "Error saving the Security Settings." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  SaveSecuritySettings = False
  GoTo TidyUpAndExit
  
End Function

Private Function ValidateSecuritySettings() As Boolean
  
  ' AE20080423 Fault #13123 - Don't validate from SQL2005 onwards as the settings come from the domain
  If glngSQLVersion < 9 Then
    If (spnResetTime.Value < spnLockoutDuration.Value) Then
      MsgBox "The reset count must be greater than the lockout duration.", vbExclamation + vbOKOnly, App.Title
      ValidateSecuritySettings = False
      spnResetTime.SetFocus
      Exit Function
    End If
    
    ' Only validate if user wishes to enforce minimum length or change
    If chkMaximumPasswordAge.Value = vbChecked Or chkMinimumPasswordLength.Value = vbChecked Then
    
      ' Validate minimum password length
      If chkMinimumPasswordLength.Value = vbChecked Then
        
        If Not IsNumeric(Me.txtPasswordLength.Text) Then
          MsgBox "Minimum password length must be numeric.", vbExclamation + vbOKOnly, App.Title
          ValidateSecuritySettings = False
          Exit Function
        ElseIf Me.txtPasswordLength.Text = 0 Then
          MsgBox "Minimum password length must be greater than zero.", vbExclamation + vbOKOnly, App.Title
          ValidateSecuritySettings = False
          Exit Function
        End If
      
      End If
      
      ' Validate password change freq/period
      If chkMaximumPasswordAge.Value = vbChecked Then
        
        If Not IsNumeric(Me.txtChangeFrequency.Text) Then
          MsgBox "Maximum password age must be numeric.", vbExclamation + vbOKOnly, App.Title
          ValidateSecuritySettings = False
          Exit Function
        ElseIf Me.txtChangeFrequency.Text = 0 Then
          MsgBox "Maximum password age must be greater than zero.", vbExclamation + vbOKOnly, App.Title
          ValidateSecuritySettings = False
          Exit Function
        ElseIf Len(Me.cboChangePeriod.Text) = 0 Then
          MsgBox "You must select a period for changing password.", vbExclamation + vbOKOnly, App.Title
          ValidateSecuritySettings = False
          Exit Function
        'TM20030122 Fault 4951 - restrict period to less that 200 years.
        ElseIf (Me.cboChangePeriod.ListIndex = 3) And (CLng(Me.txtChangeFrequency.Text) > 200) Then
          MsgBox "You cannot select a period of greater than 200 years.", vbExclamation + vbOKOnly, App.Title
          ValidateSecuritySettings = False
          Exit Function
        End If
      
      End If
    
    End If
  End If
  
  ValidateSecuritySettings = True
  
End Function

Private Sub cboAutoAddSelfServiceGroup_Click()
  Changed = True
End Sub

Private Sub cboChangePeriod_Change()
  Changed = True
End Sub

Private Sub chkAutoAddFromSelfService_Click()
  Changed = True
End Sub

Private Sub chkComplexity_Click()
  Changed = True
End Sub

Private Sub chkDelOrphanedLogins_Click()
  Changed = True
End Sub

Private Sub chkDisableDomainListBuilder_Click()
  Changed = True
End Sub

Private Sub chkDisableLoginsOnLeaveDate_Click()
  Changed = True
End Sub

Private Sub chkEmailWorkAddress_Click()
  Changed = True
End Sub

Private Sub chkMaximumPasswordAge_Click()

  ' En/Disable the related controls depending on the option selected
  If chkMaximumPasswordAge.Value = vbUnchecked Or mblnReadOnly Then
    Me.txtChangeFrequency.Text = 0
    Me.txtChangeFrequency.Enabled = False
    Me.txtChangeFrequency.BackColor = vbButtonFace
    Me.cboChangePeriod.ListIndex = 0
    Me.cboChangePeriod.Enabled = False
    Me.cboChangePeriod.BackColor = vbButtonFace
  Else
    'Me.txtChangeFrequency.Text = ""
    Me.txtChangeFrequency.Enabled = True
    Me.txtChangeFrequency.BackColor = vbWindowBackground
    Me.cboChangePeriod.ListIndex = 0
    Me.cboChangePeriod.Enabled = True
    Me.cboChangePeriod.BackColor = vbWindowBackground
    If Me.Visible Then Me.txtChangeFrequency.SetFocus ' visible clause HAS TO BE HERE !
  End If

  Changed = True

End Sub

Private Sub chkMinimumPasswordAge_Click()

  If chkMinimumPasswordAge.Value = vbUnchecked Or mblnReadOnly Then
    Me.txtMimimumPasswordAge.Text = 0
    Me.txtMimimumPasswordAge.Enabled = False
    Me.txtMimimumPasswordAge.BackColor = vbButtonFace
    Me.cboMinPasswordAge.ListIndex = 0
    Me.cboMinPasswordAge.Enabled = False
    Me.cboMinPasswordAge.BackColor = vbButtonFace
  Else
    Me.txtMimimumPasswordAge.Enabled = True
    Me.txtMimimumPasswordAge.BackColor = vbWindowBackground
    Me.cboMinPasswordAge.ListIndex = 0
    Me.cboMinPasswordAge.Enabled = True
    Me.cboMinPasswordAge.BackColor = vbWindowBackground
    If Me.Visible Then Me.txtMimimumPasswordAge.SetFocus ' visible clause HAS TO BE HERE !
  
  End If

End Sub

Private Sub chkMinimumPasswordLength_Click()
  
  ' En/Disable the related controls depending on the option selected
  If chkMinimumPasswordLength.Value = vbUnchecked Or mblnReadOnly Then
    Me.txtPasswordLength.Text = 0
    Me.txtPasswordLength.Enabled = False
    Me.txtPasswordLength.BackColor = vbButtonFace
  Else
    'Me.txtPasswordLength.Text = ""
    Me.txtPasswordLength.Enabled = True
    Me.txtPasswordLength.BackColor = vbWindowBackground
    If Me.Visible Then Me.txtPasswordLength.SetFocus  ' visible clause HAS TO BE HERE !
  End If
  
  Changed = True
  
End Sub

Private Sub chkPasswordsRemembered_Click()

  ' En/Disable the related controls depending on the option selected
  If chkPasswordsRemembered.Value = vbUnchecked Or mblnReadOnly Then
    Me.txtPasswordsRemembered.Text = 0
    Me.txtPasswordsRemembered.Enabled = False
    Me.txtPasswordsRemembered.BackColor = vbButtonFace
  Else
    Me.txtPasswordsRemembered.Enabled = True
    Me.txtPasswordsRemembered.BackColor = vbWindowBackground
    If Me.Visible Then Me.txtPasswordsRemembered.SetFocus  ' visible clause HAS TO BE HERE !
  End If
  
  Changed = True

End Sub

Private Sub chkPCLockout_Click()
  RefreshButtons
  Changed = True
End Sub

Private Sub chkUseDomainPolicy_Click()
  Changed = True
  RefreshButtons
End Sub

Private Sub chkSecManBypass_Click()
  Changed = True
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If ValidateSecuritySettings Then
    If SaveSecuritySettings Then
      Changed = False
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
  Dim pintAnswer As Integer
    
  If mblnChanged = True Then
    
    pintAnswer = MsgBox("You have changed the current security options. Save changes?", vbQuestion + vbYesNoCancel, App.Title)
      
    If pintAnswer = vbYes Then
      cmdOK_Click
      Cancel = False
      Exit Sub
    ElseIf pintAnswer = vbCancel Then
      Cancel = True
      Exit Sub
    End If
  
  End If

End Sub

Private Sub Form_Resize()
  DisplayApplication
End Sub

Private Sub spnAttempts_Change()
  Changed = True
End Sub

Private Sub spnLockoutDuration_Change()
  Changed = True
End Sub

Private Sub spnResetTime_Change()
  Changed = True
End Sub

Private Function LoadCombos() As Boolean

  On Error GoTo Load_ERROR
  
  ' Populate the combo box with the required frequencies
  With cboChangePeriod
    .AddItem "Day(s)"
    .AddItem "Week(s)"
    .AddItem "Month(s)"
    .AddItem "Year(s)"
  End With

  With cboMinPasswordAge
    .AddItem "Day(s)"
    .AddItem "Week(s)"
    .AddItem "Month(s)"
    .AddItem "Year(s)"
  End With
  
  Dim objGroup As SecurityGroup
  For Each objGroup In gObjGroups
   cboAutoAddSelfServiceGroup.AddItem (objGroup.Name)
  Next objGroup

  LoadCombos = True
  Exit Function
  
Load_ERROR:
  
  MsgBox "Error initialising the security options." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
         
  LoadCombos = False
  
End Function

Private Sub txtChangeFrequency_Change()
  Changed = True
End Sub

Private Sub txtPasswordLength_Change()
  Changed = True
End Sub

Private Sub RefreshButtons()

  If mblnReadOnly Then
    ControlsDisableAll Me
  Else
    
    ' AE20080218 Fault #12876
    'If glngSQLVersion = 9 Then
    If glngSQLVersion >= 9 Then
      ControlsDisableAll fraPolicy, False
      chkSecManBypass.Enabled = gbUserCanManageLogins
      chkDelOrphanedLogins.Enabled = True
    Else
      chkMinimumPasswordLength.Enabled = True
      chkMaximumPasswordAge.Enabled = True
      chkComplexity.Enabled = False
      chkMinimumPasswordAge.Enabled = False
      chkPCLockout.Enabled = True
      chkSecManBypass.Enabled = False
      chkDelOrphanedLogins.Enabled = gbCanUseWindowsAuthentication
      
      spnAttempts.Enabled = (chkPCLockout.Value = vbChecked)
      spnResetTime.Enabled = (chkPCLockout.Value = vbChecked)
      spnLockoutDuration.Enabled = (chkPCLockout.Value = vbChecked)
      
      spnAttempts.BackColor = IIf(chkPCLockout.Enabled, vbWindowBackground, vbButtonFace)
      spnResetTime.BackColor = IIf(chkPCLockout.Enabled, vbWindowBackground, vbButtonFace)
      spnLockoutDuration.BackColor = IIf(chkPCLockout.Enabled, vbWindowBackground, vbButtonFace)
      
      If (chkPCLockout.Value = vbUnchecked) Then
        spnAttempts.Value = 1
        spnResetTime.Value = 1
        spnLockoutDuration.Value = 1
      End If
      
    End If
  
    chkDelOrphanedLogins.Enabled = gbCanUseWindowsAuthentication
    lblDeleteOrphanWarning.Visible = chkDelOrphanedLogins.Enabled
  
    ' PC lockout options
    lblLockoutAfter.ForeColor = IIf(chkPCLockout.Enabled, vbBlack, vbGrayText)
    lblResetTime.ForeColor = IIf(chkPCLockout.Enabled, vbBlack, vbGrayText)
    lblLockoutDuration.ForeColor = IIf(chkPCLockout.Enabled, vbBlack, vbGrayText)
    spnAttempts.BackColor = IIf(chkPCLockout.Enabled, vbWindowBackground, vbButtonFace)
    spnResetTime.BackColor = IIf(chkPCLockout.Enabled, vbWindowBackground, vbButtonFace)
    spnLockoutDuration.BackColor = IIf(chkPCLockout.Enabled, vbWindowBackground, vbButtonFace)
  
  End If
  
End Sub

