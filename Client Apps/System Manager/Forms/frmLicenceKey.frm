VERSION 5.00
Begin VB.Form frmLicenceKey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Key"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8041
   Icon            =   "frmLicenceKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   4
      Top             =   960
      Width           =   6255
      Begin VB.TextBox txtLicence 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
      Begin VB.TextBox txtLicence 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2200
         MaxLength       =   6
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
      Begin VB.TextBox txtLicence 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
      Begin VB.TextBox txtLicence 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   45
         MaxLength       =   6
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
      Begin VB.TextBox txtLicence 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4260
         MaxLength       =   6
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
      Begin VB.TextBox txtLicence 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   5235
         MaxLength       =   6
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   3087
         TabIndex        =   15
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2007
         TabIndex        =   14
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   917
         TabIndex        =   13
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4110
         TabIndex        =   12
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   5115
         TabIndex        =   11
         Top             =   45
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2070
      TabIndex        =   1
      Top             =   1620
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3360
      TabIndex        =   2
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label lblSupportTel 
      AutoSize        =   -1  'True
      Caption         =   "(XXXX) XXXXXX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3915
      TabIndex        =   3
      Top             =   240
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "Please call OpenHR Customer Services on"
      Height          =   615
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   5910
   End
End
Attribute VB_Name = "frmLicenceKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean
Private mstrAllowedInputCharacters As String

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Get LicenceKey() As String
  LicenceKey = txtLicence(0).Text & "-" & txtLicence(1).Text & "-" & _
               txtLicence(2).Text & "-" & txtLicence(3).Text & "-" & txtLicence(4).Text & "-" & txtLicence(5).Text
End Property

Private Function GenerateAlphaString() As String

  Dim strOutput As String
  Dim lngCount As Long
  Dim lngLoop As Long

  'Only allow these characters...
  strOutput = vbNullString

  For lngCount = Asc("A") + lngLoop To Asc("Z")
    strOutput = strOutput & Chr(lngCount)
  Next

  For lngCount = Asc("0") + lngLoop To Asc("9")
    strOutput = strOutput & Chr(lngCount)
  Next

  GenerateAlphaString = strOutput

End Function


Private Sub cmdApply_Click()
  mblnCancelled = False
  Me.Hide
End Sub

Private Sub cmdCancel_Click()
  mblnCancelled = True
  Me.Hide
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

  mblnCancelled = True

  Frame1.BackColor = Me.BackColor
  Label1.Caption = _
    "Please call OpenHR Customer Services on" & vbCrLf & _
    "quoting your existing licence number and licence amendments."
  lblSupportTel = GetSystemSetting("Support", "Telephone No", "")

  mstrAllowedInputCharacters = GenerateAlphaString

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub txtLicence_Change(Index As Integer)

  If Len(txtLicence(Index).Text) >= 4 And txtLicence(Index).SelStart = 5 Then
    If Index < txtLicence.UBound Then
      txtLicence(Index + 1).SetFocus
    End If
  End If

End Sub

Private Sub txtLicence_GotFocus(Index As Integer)
  With txtLicence(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLicence_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  'Check if a user is trying to paste in a whole licence key
  'If they are, then separate it into each text box.
  If KeyCode = vbKeyV And (Shift And vbCtrlMask) Then
    If Clipboard.GetText Like "??????-??????-??????-??????-??????-??????" Then
      txtLicence(0).Text = Mid(Clipboard.GetText, 1, 6)
      txtLicence(1).Text = Mid(Clipboard.GetText, 8, 6)
      txtLicence(2).Text = Mid(Clipboard.GetText, 15, 6)
      txtLicence(3).Text = Mid(Clipboard.GetText, 22, 6)
      txtLicence(4).Text = Mid(Clipboard.GetText, 29, 6)
      txtLicence(5).Text = Mid(Clipboard.GetText, 36, 6)
      KeyCode = 0
      Shift = 0
    End If
  End If

End Sub

Private Sub txtLicence_KeyPress(Index As Integer, KeyAscii As Integer)

  Dim strChar As String
  
  'Allow control characters...
  If KeyAscii > 31 Then
  
    strChar = UCase(Chr(KeyAscii))
    If InStr(mstrAllowedInputCharacters, strChar) > 0 Then
      KeyAscii = Asc(strChar)
    Else
      KeyAscii = 0
    End If
  
  End If

End Sub

Private Sub txtLicence_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

  Dim objLicence As New clsLicence
  objLicence.ValidateCreationDate = True
  objLicence.LicenceKey = txtLicence(0).Text & "-" & txtLicence(1).Text & "-" & _
               txtLicence(2).Text & "-" & txtLicence(3).Text & "-" & txtLicence(4).Text & "-" & txtLicence(5).Text

  ' Validate licence key
  cmdApply.Enabled = objLicence.IsValid

End Sub
