VERSION 5.00
Begin VB.Form frmSaveChangesPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Changes ?"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1026
   Icon            =   "frmSaveChangesPrompt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOKCancel 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   1080
      TabIndex        =   7
      Top             =   1275
      Width           =   1725
      Begin VB.CommandButton cmdOKCancel_Cancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   885
         TabIndex        =   4
         Top             =   0
         Width           =   825
      End
      Begin VB.CommandButton cmdOKCancel_OK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.Frame fraYesNoCancel 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   585
      TabIndex        =   6
      Top             =   825
      Width           =   2520
      Begin VB.CommandButton cmdYesNoCancel_No 
         Caption         =   "&No"
         Height          =   350
         Left            =   885
         TabIndex        =   1
         Top             =   0
         Width           =   780
      End
      Begin VB.CommandButton cmdYesNoCancel_Yes 
         Caption         =   "&Yes"
         Height          =   350
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   780
      End
      Begin VB.CommandButton cmdYesNoCancel_Cancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   1740
         TabIndex        =   2
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmSaveChangesPrompt.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblSaveChanges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save all changes ?"
      Height          =   195
      Left            =   870
      TabIndex        =   5
      Top             =   150
      Width           =   1320
   End
End
Attribute VB_Name = "frmSaveChangesPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miChoice As Integer
Private mfRefreshDatabase As Boolean
Private miButtons As Integer






Private Sub cmdOKCancel_Cancel_Click()
  miChoice = vbCancel
  UnLoad Me

End Sub

Private Sub cmdOKCancel_OK_Click()
  miChoice = vbOK
  UnLoad Me

End Sub

Private Sub cmdYesNoCancel_Cancel_Click()
  miChoice = vbCancel
  UnLoad Me

End Sub

Private Sub cmdYesNoCancel_No_Click()
  miChoice = vbNo
  UnLoad Me

End Sub

Private Sub cmdYesNoCancel_Yes_Click()
  miChoice = vbYes
  UnLoad Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If ((Shift And vbShiftMask) > 0) Then
    mfRefreshDatabase = True
  End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    mfRefreshDatabase = False

End Sub


Private Sub Form_Load()
'  Const iBUTTONSTOP = 550
  
  'NHRD - 17042003 - Fault 2603
  With frmSysMgr
    If .WindowState = vbMinimized Then .WindowState = vbMaximized
  End With
      
  
  With fraYesNoCancel
    .BackColor = Me.BackColor
'    .Top = iBUTTONSTOP
'    .Left = 200
  End With
  
  With fraOKCancel
    .BackColor = Me.BackColor
'    .Top = iBUTTONSTOP
'    .Left = 650
  End With
  
'  Me.Height = 1400
'  Me.Width = 2700
  
  If gfRefreshStoredProcedures Then
    lblSaveChanges.Caption = "Changes need to be saved to update the database to the latest version." & vbCrLf & vbCrLf & _
      "Save changes ?"
  Else
    lblSaveChanges.Caption = "Save all changes ?"
  End If
  
  FormatForm vbYesNoCancel
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    miChoice = vbCancel
  End If

End Sub



Public Property Get Choice() As Integer
  Choice = miChoice
  
End Property

Public Property Let Choice(ByVal piNewValue As Integer)
  miChoice = piNewValue
  
End Property

Public Property Get RefreshDatabase() As Boolean
  RefreshDatabase = mfRefreshDatabase
  
End Property

Public Property Let RefreshDatabase(ByVal pfNewValue As Boolean)
  mfRefreshDatabase = pfNewValue
  
End Property

Public Property Get Buttons() As Integer
  Buttons = miButtons
  
End Property

Public Property Let Buttons(ByVal piNewValue As Integer)
  miButtons = piNewValue
  FormatForm piNewValue
  
End Property

Private Sub FormatForm(piButtons As Integer)
  ' Format the form.
  Dim lngXExtent As Long
  Dim lngYExtent As Long
  
  Const XGAP = 200
  Const YGAP = 200
  
  fraYesNoCancel.Visible = (piButtons = vbYesNoCancel)
  fraOKCancel.Visible = (piButtons = vbOKCancel)
  
  ' Ensure the form is wide enough for all displayed controls.
  lngXExtent = lblSaveChanges.Left + lblSaveChanges.Width
  If (piButtons = vbYesNoCancel) And _
    (lngXExtent < (fraYesNoCancel.Left + fraYesNoCancel.Width)) Then
    lngXExtent = fraYesNoCancel.Left + fraYesNoCancel.Width
  End If
  If (piButtons = vbOKCancel) And _
    (lngXExtent < (fraOKCancel.Left + fraOKCancel.Width)) Then
    lngXExtent = fraOKCancel.Left + fraOKCancel.Width
  End If
  
  Me.Width = lngXExtent + XGAP
  
  ' Ensure the buttons are correctly positioned.
  ' Ensure the form is tall enough.
  'NHRD24042002 Fault 3740 Added an question mark image icon so used this as the
  'benchmark for determing the height of the form
  'lngYExtent = lblSaveChanges.Top + lblSaveChanges.Height
  lngYExtent = Image1.Top + Image1.Height
  If (piButtons = vbYesNoCancel) Then
    fraYesNoCancel.Top = lngYExtent + YGAP
    fraYesNoCancel.Left = (Me.Width - fraYesNoCancel.Width) / 2
    Me.Height = fraYesNoCancel.Top + fraYesNoCancel.Height + YGAP + UI.CaptionHeight + (2 * UI.YBorder) + (2 * UI.YFrame)
  End If
  If (piButtons = vbOKCancel) Then
    fraOKCancel.Top = lngYExtent + YGAP
    Me.Height = fraOKCancel.Top + fraOKCancel.Height + YGAP + UI.CaptionHeight + (2 * UI.YBorder) + (2 * UI.YFrame)
    fraOKCancel.Left = (Me.Width - fraOKCancel.Width) / 2
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


