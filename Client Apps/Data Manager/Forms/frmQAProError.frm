VERSION 5.00
Begin VB.Form frmQAProError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Error"
   ClientHeight    =   3735
   ClientLeft      =   4530
   ClientTop       =   6240
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQAProError.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtError 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmQAProError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdDetails_Click()
    Me.Height = 4140
    txtError.SetFocus
    cmdDetails.Enabled = False
End Sub

Private Sub cmdOK_Click()
    Unload Me
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
    Me.Height = 1500
    SetWindowPos Me.hWnd, conHwndTopmost, (Me.Left / 15.5), (Me.Top / 15.5), (Me.Width / 15.5) + 13, (Me.Height / 15.5), conSwpNoActivate Or conSwpShowWindow
    Me.Visible = False
End Sub



