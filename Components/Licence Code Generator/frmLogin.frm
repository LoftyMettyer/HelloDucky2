VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1260
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   744.45
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   720
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   720
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   195
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

  ' Only show the login form if we arent in VB mode
  If Not vbCompiled Then
    Unload Me
    frmHRProLicence.Show
  End If
  
End Sub

Private Sub cmdCancel_Click()
  
  Unload Me

End Sub

Private Sub cmdOK_Click()
    
  Dim sTodaysPassword As String
  
  'check for correct password

  sTodaysPassword = Format(LTrim(Str(Day(Date))), "00") _
                  + Format(LTrim(Str(Month(Date))), "00") _
                  + Format(LTrim(Str(Day(Date) + 10)), "00") _
                  + Format(LTrim(Str(Month(Date) + 10)), "00")

  If txtPassword.Text = sTodaysPassword Then
    Unload Me
    frmHRProLicence.Show
  Else
    MsgBox "Invalid Password", , "Login"
    txtPassword.SetFocus
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
  End If

End Sub

Private Function vbCompiled() As Boolean

  On Local Error Resume Next
  Err.Clear
  Debug.Print 1 / 0
  vbCompiled = (Err.Number = 0)

End Function
