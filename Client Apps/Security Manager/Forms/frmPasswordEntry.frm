VERSION 5.00
Begin VB.Form frmPasswordEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Login Password"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8017
   Icon            =   "frmPasswordEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkChangePassword 
      Caption         =   "&User must change password at next login"
      Height          =   270
      Left            =   195
      TabIndex        =   4
      Top             =   1215
      Value           =   1  'Checked
      Width           =   3870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2355
      TabIndex        =   5
      Top             =   1605
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3615
      TabIndex        =   6
      Top             =   1605
      Width           =   1200
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1890
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   700
      Width           =   2955
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1890
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   200
      Width           =   2955
   End
   Begin VB.Label lblConfirmpassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password :"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   765
      Width           =   1755
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   255
      Width           =   1020
   End
End
Attribute VB_Name = "frmPasswordEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private msPassword As String
Private mstrUserName As String
Private mbForcePasswordChange As Boolean
Private mbCheckPolicy As Boolean

Public Property Let UserName(ByVal pstrNewValue As String)
  mstrUserName = pstrNewValue
End Property

Public Property Get Password() As String
  Password = msPassword
End Property

Public Property Get ForcePasswordChange() As Boolean
  ForcePasswordChange = mbForcePasswordChange
End Property

Public Property Get CheckPolicy() As Boolean
  CheckPolicy = mbCheckPolicy
End Property

Private Sub chkChangePassword_Click()
  mbForcePasswordChange = IIf(chkChangePassword.Value = vbChecked, True, False)
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  Unload Me
  
End Sub


Private Sub cmdOK_Click()
  ' Validate the password.
 
  'Dim bBypassPolicy As Boolean
  'bBypassPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)   ' Default - Off
  ' AE20080519 Fault #13164
  ' mbCheckPolicy = Not (GetSystemSetting("Policy", "Sec Man Bypass", 0))
  mbCheckPolicy = Not (GetSystemSetting("Policy", "Sec Man Bypass", 0) = 1)
  
  If (Trim(txtPassword.Text) <> Trim(txtConfirmPassword.Text)) Then
    MsgBox "The confirmation password is not correct.", vbExclamation + vbOKOnly, App.ProductName
    txtConfirmPassword.SetFocus
    Exit Sub
  End If
  
  'If Not bBypassPolicy Then
  If mbCheckPolicy Then
  
    If CheckOverMaxLength = False Then
      txtPassword.SetFocus
      Exit Sub
    End If
  
    ' Complex enough
    If CheckPasswordComplexity(mstrUserName, txtPassword.Text) = False Then
      txtPassword.Text = ""
      txtConfirmPassword.Text = ""
      txtPassword.SetFocus
      MsgBox "The password does not meet the policy requirement because it is not complicated enough.", vbExclamation + vbOKOnly, App.Title
      Exit Sub
    End If
  
  End If
  
  msPassword = Trim(txtPassword.Text)
  Cancelled = False
  Unload Me
End Sub


Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
  
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  
  Dim bBypassPolicy As Boolean
  
  bBypassPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)  ' Default - Off
  
  If glngSQLVersion >= 9 And bBypassPolicy Then
    mbForcePasswordChange = False
    chkChangePassword.Value = vbUnchecked
    chkChangePassword.Enabled = False
'  ElseIf glngSQLVersion >= 9 Then
'    mbForcePasswordChange = True
'    chkChangePassword.Enabled = False
  Else
    mbForcePasswordChange = True
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancelled = True
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub txtConfirmPassword_GotFocus()
  UI.txtSelText

End Sub


Private Sub txtPassword_GotFocus()
  UI.txtSelText

End Sub

Private Function CheckOverMaxLength() As Boolean

  If glngDomainMinimumLength = 0 Then
    CheckOverMaxLength = True
  ElseIf Len(Me.txtPassword.Text) >= glngDomainMinimumLength Then
    CheckOverMaxLength = True
  Else
    MsgBox "The new password must be at least " & glngDomainMinimumLength & " characters long.", vbExclamation + vbOKOnly, App.Title
    CheckOverMaxLength = False
  End If
  
End Function
