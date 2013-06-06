VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmWorkflowEditOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resize Workflow Canvas"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5076
   Icon            =   "frmWorkflowEditOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3300
      TabIndex        =   7
      Top             =   1100
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2000
      TabIndex        =   6
      Top             =   1100
      Width           =   1200
   End
   Begin COASpinner.COA_Spinner asrNewWidth 
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaximumValue    =   245752
      MinimumValue    =   1
      Text            =   "1"
   End
   Begin COASpinner.COA_Spinner asrNewHeight 
      Height          =   315
      Left            =   3300
      TabIndex        =   5
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaximumValue    =   245752
      MinimumValue    =   1
      Text            =   "1"
   End
   Begin VB.Label lblNewHeight 
      Caption         =   "New Height :"
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblNewWidth 
      Caption         =   "New Width :"
      Height          =   195
      Left            =   2100
      TabIndex        =   1
      Top             =   255
      Width           =   1035
   End
   Begin VB.Label lblCurrentHeight 
      AutoSize        =   -1  'True
      Caption         =   "Current Height : 0"
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   660
      Width           =   1560
   End
   Begin VB.Label lblCurrentWidth 
      AutoSize        =   -1  'True
      Caption         =   "Current Width : 0"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   255
      Width           =   1500
   End
End
Attribute VB_Name = "frmWorkflowEditOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LABELPREFIX_CURRENTWIDTH = "Current Width : "
Private Const LABELPREFIX_CURRENTHEIGHT = "Current Height : "

Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mlngWidth As Long
Private mlngHeight As Long


Public Property Let Changed(ByVal pfNewValue As Boolean)
  mfChanged = pfNewValue
  RefreshScreen
  
End Property
Private Sub RefreshScreen()
  ' Refresh the screen controls.
  Dim fReadOnly As Boolean
  
  fReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)
  
  If fReadOnly Then
    ControlsDisableAll Me
  End If

  cmdOK.Enabled = mfChanged And (Not fReadOnly)

End Sub



Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property
Private Sub FormatForm()
  Const GAP_CurrentToNew = 500
  Const GAP_LabelToSpinner = 100
  Const GAP_ButtonToButton = 100
  
  lblNewWidth.Left = Maximum(lblCurrentWidth.Left + lblCurrentWidth.Width, _
    lblCurrentHeight.Left + lblCurrentHeight.Width) + GAP_CurrentToNew
  lblNewHeight.Left = lblNewWidth.Left
  
  asrNewWidth.Left = Maximum(lblNewWidth.Left + lblNewWidth.Width, _
    lblNewHeight.Left + lblNewHeight.Width) + GAP_LabelToSpinner
  asrNewHeight.Left = asrNewWidth.Left
    
  cmdCancel.Left = asrNewWidth.Left + asrNewWidth.Width - cmdCancel.Width
  cmdOK.Left = cmdCancel.Left - cmdOK.Width - GAP_ButtonToButton
  
  Me.Width = asrNewWidth.Left + asrNewWidth.Width + lblCurrentWidth.Left
    
End Sub

Public Property Get CanvasHeight() As Long
  ' Return the canvas height
  CanvasHeight = mlngHeight
  
End Property

Public Property Get CanvasWidth() As Long
  ' Return the canvas width
  CanvasWidth = mlngWidth
  
End Property


Public Property Let CanvasHeight(ByVal plngNewValue As Long)
  ' Set the canvas height.
  asrNewHeight.value = plngNewValue
  lblCurrentHeight.Caption = LABELPREFIX_CURRENTHEIGHT & CStr(plngNewValue)
  Changed = False
  
  FormatForm

End Property
Public Property Let CanvasWidth(ByVal plngNewValue As Long)
  ' Set the canvas width.
  asrNewWidth.value = plngNewValue
  lblCurrentWidth.Caption = LABELPREFIX_CURRENTWIDTH & CStr(plngNewValue)

  Changed = False
  FormatForm
  
End Property

Private Sub asrNewHeight_Change()
  mlngHeight = asrNewHeight.value
  Changed = True

End Sub

Private Sub asrNewWidth_Change()
  mlngWidth = asrNewWidth.value
  Changed = True

End Sub


Private Sub cmdCancel_Click()
  mfCancelled = True
  UnLoad Me

End Sub

Private Sub cmdOK_Click()
  mfCancelled = False
  UnLoad Me

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
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iAnswer As Integer
  
  If UnloadMode <> vbFormCode Then

    'Check if any changes have been made.
    If mfChanged Then
      iAnswer = MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If iAnswer = vbYes Then
        Call cmdOK_Click
        If mfCancelled Then Cancel = 1
      ElseIf iAnswer = vbNo Then
        mfCancelled = True
      ElseIf iAnswer = vbCancel Then
        Cancel = 1
      End If
    Else
      mfCancelled = True
    End If
  End If
  
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


