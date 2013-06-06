VERSION 5.00
Begin VB.Form frmWorkflowPrintOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Print Options"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5068
   Icon            =   "frmWorkflowPrintOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   200
      TabIndex        =   2
      Top             =   1100
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1600
      TabIndex        =   3
      Top             =   1100
      Width           =   1200
   End
   Begin VB.CheckBox chkPrintDetails 
      Caption         =   "Print &Details (text)"
      Height          =   315
      Left            =   200
      TabIndex        =   1
      Top             =   600
      Width           =   2145
   End
   Begin VB.CheckBox chkPrintOverview 
      Caption         =   "&Print Overview (graphic)"
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Value           =   1  'Checked
      Width           =   2550
   End
End
Attribute VB_Name = "frmWorkflowPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mfPrintDetails As Boolean
Private mfPrintOverview As Boolean

Private Sub RefreshButtons()
  mfPrintDetails = (chkPrintDetails.value = vbChecked)
  mfPrintOverview = (chkPrintOverview.value = vbChecked)
  cmdOK.Enabled = mfPrintDetails Or mfPrintOverview
  
End Sub

Private Sub chkPrintDetails_Click()
  RefreshButtons
  
End Sub


Private Sub chkPrintOverview_Click()
  RefreshButtons

End Sub



Public Property Get PrintDetails() As Boolean
  PrintDetails = mfPrintDetails
  
End Property
Public Property Get PrintOverview() As Boolean
  PrintOverview = mfPrintOverview
  
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property

Private Sub cmdCancel_Click()
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
  mfCancelled = True
  RefreshButtons
  
End Sub


