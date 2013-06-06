VERSION 5.00
Begin VB.Form frmSSIntranetLinkSeparator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Link Separator"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5087
   Icon            =   "frmSSIIntranetLinkSeparator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLinkSeparator 
      Caption         =   "Link Separator :"
      Height          =   870
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   4380
      Begin VB.TextBox txtLinkSeparator 
         Height          =   315
         Left            =   1035
         MaxLength       =   500
         TabIndex        =   2
         Top             =   300
         Width           =   3195
      End
      Begin VB.Label lblLinkSeparator 
         Caption         =   "Text :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   2600
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   5
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   135
         TabIndex        =   4
         Top             =   0
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmSSIntranetLinkSeparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean
Private mfChanged As Boolean
Private mblnRefreshing As Boolean
Private mlngTableID As Long
Private mlngViewID As Long

Private mblnReadOnly As Boolean

Public Sub Initialize(piType As SSINTRANETLINKTYPES, _
                      psText As String, _
                      plngTableID As Long, _
                      plngViewID As Long, _
                      pfCopy As Boolean)
  
  mlngTableID = plngTableID
  mlngViewID = plngViewID
  
  Select Case piType
    Case SSINTLINK_HYPERTEXT
      fraLinkSeparator.Caption = "Hypertext Link Separator :"
    Case SSINTLINK_BUTTON
      fraLinkSeparator.Caption = "Button Link Separator :"
  End Select
  
  FormatScreen
  
  Text = psText
  
  mfChanged = False
  If pfCopy Then mfChanged = True
  RefreshControls
  
End Sub

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property
Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Private Sub FormatScreen()

  Const GAPBETWEENTEXTBOXES = 85
  Const GAPABOVEBUTTONS = 150
  Const GAPUNDERBUTTONS = 600
  Const LEFTGAP = 200
  Const GAPUNDERLASTCONTROL = 200
  Const GAPUNDERRADIOBUTTON = -15

  ' Position the OK/Cancel buttons
  fraOKCancel.Top = fraLinkSeparator.Top + fraLinkSeparator.Height + GAPABOVEBUTTONS
  
  ' Redimension the form.
  Me.Height = fraOKCancel.Top + fraOKCancel.Height + GAPUNDERBUTTONS

End Sub

Private Function ValidateLink() As Boolean

  ' Return FALSE if the link definition is invalid.
  Dim fValid As Boolean
  
  fValid = True

  ValidateLink = fValid
  
End Function

Private Sub cmdOK_Click()

  If ValidateLink Then
    Cancelled = False
    Me.Hide
  End If

End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  UnLoad Me
End Sub

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Get ViewID() As Long
  ViewID = mlngViewID
End Property

Private Sub RefreshControls()

  Dim sUtilityMessage As String
  
  If mblnRefreshing Then Exit Sub
 
  ' Disable the OK button as required.
  cmdOK.Enabled = True
  
End Sub

Public Property Get Text() As String
  Text = txtLinkSeparator.Text
End Property
Public Property Let Text(ByVal psNewValue As String)
  txtLinkSeparator.Text = psNewValue
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

  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  fraOKCancel.BorderStyle = vbBSNone

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancelled = True
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOK_Click
        Cancel = True
    End Select
  End If

End Sub


Private Sub txtLinkSeparator_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtLinkSeparator_GotFocus()
  UI.txtSelText
End Sub

