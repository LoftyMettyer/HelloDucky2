VERSION 5.00
Begin VB.Form frmSupportMode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Support Mode"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1048
   Icon            =   "frmSupportMode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   4320
      Begin VB.TextBox txtLicence 
         Height          =   315
         Index           =   3
         Left            =   2655
         MaxLength       =   4
         TabIndex        =   8
         Top             =   0
         Width           =   750
      End
      Begin VB.TextBox txtLicence 
         Height          =   315
         Index           =   2
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   7
         Top             =   0
         Width           =   750
      End
      Begin VB.TextBox txtLicence 
         Height          =   315
         Index           =   1
         Left            =   885
         MaxLength       =   4
         TabIndex        =   6
         Top             =   0
         Width           =   750
      End
      Begin VB.TextBox txtLicence 
         Height          =   315
         Index           =   0
         Left            =   0
         MaxLength       =   4
         TabIndex        =   5
         Top             =   0
         Width           =   750
      End
      Begin VB.TextBox txtLicence 
         Height          =   315
         Index           =   4
         Left            =   3540
         MaxLength       =   4
         TabIndex        =   4
         Top             =   0
         Width           =   750
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2550
         TabIndex        =   12
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1665
         TabIndex        =   11
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   780
         TabIndex        =   10
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3435
         TabIndex        =   9
         Top             =   45
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1290
      TabIndex        =   0
      Top             =   2160
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2625
      TabIndex        =   1
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label lblOutputKey 
      Alignment       =   2  'Center
      Caption         =   "lblOutputKey"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   848
      TabIndex        =   13
      Top             =   1080
      Width           =   3375
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
      Left            =   2970
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Please call HR Pro Support on"
      Height          =   615
      Left            =   345
      TabIndex        =   14
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmSupportMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjSupport As COALicence.clsSupport
Private mstrAllowedInputCharacters As String
Private mblnCancelled As Boolean

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Get LicenceKey() As String
  LicenceKey = txtLicence(0).Text & "-" & txtLicence(1).Text & "-" & _
               txtLicence(2).Text & "-" & txtLicence(3).Text & "-" & txtLicence(4).Text
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

  If mobjSupport.CheckSupportInputString2(LicenceKey) Then

    If Application.AccessMode = accLimited Then
      ' JPD20030211 Fault 5046
      MsgBox "Support mode successfully enabled." & vbCrLf & vbCrLf & _
             "NOTE: After you have made the required changes you will be " & vbCrLf & _
             "asked to enter an additional code prior to saving changes.", _
             vbInformation, "Support Mode"
      Application.AccessMode = accSupportMode
      frmSysMgr.SetCaption
    End If

    mblnCancelled = False
    Me.Hide

  Else
    MsgBox "Invalid authorisation code", vbExclamation, "Support Mode"

  End If

End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub Form_Load()

  lblSupportTel.Visible = (Application.AccessMode = accLimited)
  
  If Application.AccessMode = accLimited Then
    Label1.Caption = _
      "Please call HR Pro Support on" & vbCrLf & _
      "quoting the key shown below to obtain an authorisation code."
    lblSupportTel = GetSystemSetting("Support", "Telephone No", "")

  Else
    lblSupportTel.Visible = False
    Label1.Caption = "In order to proceed with saving changes please enter an authorisation code based on the following key:"

  End If
  
  
  mblnCancelled = True
  Frame1.BackColor = Me.BackColor
  
  mstrAllowedInputCharacters = GenerateAlphaString
  
  Set mobjSupport = New COALicence.clsSupport
  lblOutputKey.Caption = mobjSupport.GetSupportString2

  Screen.MousePointer = vbNormal

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Me.Visible Then
    mblnCancelled = True
    Cancel = True
    Me.Hide
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mobjSupport = Nothing
End Sub

Private Sub lblOutputKey_DblClick()
  Clipboard.SetText lblOutputKey.Caption
End Sub

Private Sub txtLicence_Change(Index As Integer)

  If Len(txtLicence(Index).Text) >= 4 And txtLicence(Index).SelStart = 4 Then
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
    If Clipboard.GetText Like "????-????-????-????-????" Then
      txtLicence(0).Text = Mid(Clipboard.GetText, 1, 4)
      txtLicence(1).Text = Mid(Clipboard.GetText, 6, 4)
      txtLicence(2).Text = Mid(Clipboard.GetText, 11, 4)
      txtLicence(3).Text = Mid(Clipboard.GetText, 16, 4)
      txtLicence(4).Text = Mid(Clipboard.GetText, 21, 4)
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

