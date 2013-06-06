VERSION 5.00
Begin VB.Form frmAuditOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Audit Log"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1008
   Icon            =   "frmAuditOpen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1590
      TabIndex        =   5
      Top             =   2100
      Width           =   1200
   End
   Begin VB.Frame fraSelectAuditType 
      Caption         =   "Select Audit Log :"
      Height          =   1875
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   2700
      Begin VB.OptionButton optAccess 
         Caption         =   "User &Access"
         Height          =   315
         Left            =   195
         TabIndex        =   3
         Top             =   1470
         Width           =   1800
      End
      Begin VB.OptionButton optRecords 
         Caption         =   "&Data Records"
         Height          =   315
         Left            =   200
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optPermissions 
         Caption         =   "Data &Permissions"
         Height          =   315
         Left            =   200
         TabIndex        =   1
         Top             =   700
         Width           =   2300
      End
      Begin VB.OptionButton optUsers 
         Caption         =   "&User Maintenance"
         Height          =   315
         Left            =   200
         TabIndex        =   2
         Top             =   1100
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   255
      TabIndex        =   4
      Top             =   2100
      Width           =   1200
   End
End
Attribute VB_Name = "frmAuditOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property

Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide

End Sub

Private Sub cmdOK_Click()
  mfCancelled = False
  Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
    Cancel = True
    Me.Hide
  End If

End Sub

Public Property Get OpenType() As audType
  
  If optRecords Then
    OpenType = audRecords
  ElseIf optPermissions Then
    OpenType = audPermissions
  ElseIf Me.optUsers Then
    OpenType = audGroups
  ElseIf Me.optAccess Then
    OpenType = audAccess
  End If

End Property

Public Sub Initialise(AuditType As audType)
  
  Select Case AuditType
    Case audRecords
      optRecords.Value = True
    Case audPermissions
      optPermissions.Value = True
    Case audGroups
      optUsers.Value = True
    Case audAccess
      optAccess.Value = True
  End Select

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


