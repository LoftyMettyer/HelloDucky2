VERSION 5.00
Begin VB.Form frmSecurityPrintOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Group Details"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8024
   Icon            =   "frmSecurityPrintOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2295
      TabIndex        =   5
      Top             =   2970
      Width           =   1200
   End
   Begin VB.Frame frmContents 
      Caption         =   "Items :"
      Height          =   1950
      Left            =   135
      TabIndex        =   6
      Top             =   885
      Width           =   3390
      Begin VB.CheckBox chkWithColPermisions 
         Caption         =   "Column &Permissions"
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   2100
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "&User Logins"
         Height          =   315
         Left            =   200
         TabIndex        =   0
         Top             =   255
         Width           =   1500
      End
      Begin VB.CheckBox chkTablesViews 
         Caption         =   "&Data Permissions"
         Height          =   315
         Left            =   200
         TabIndex        =   1
         Top             =   675
         Width           =   1900
      End
      Begin VB.CheckBox chkSysPerms 
         Caption         =   "&System Permissions"
         Height          =   315
         Left            =   200
         TabIndex        =   3
         Top             =   1470
         Width           =   2100
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1020
      TabIndex        =   4
      Top             =   2970
      Width           =   1200
   End
   Begin VB.Frame frmSelection 
      Caption         =   "Selection :"
      Height          =   735
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   3390
      Begin VB.CheckBox chkBlankVersion 
         Caption         =   "&Blank Version of Selected Items"
         Height          =   405
         Left            =   200
         TabIndex        =   8
         Top             =   225
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmSecurityPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mbPrintUserLogins As Boolean
Private mbPrintDataPerms As Boolean
Private mbPrintColumnPerms As Boolean
Private mbPrintSystemPerms As Boolean

Private Sub RefreshControls()
  'MH20040130 Fault 7985
  chkWithColPermisions.Enabled = (chkTablesViews.Value = vbChecked)

  'JPD 20030515 - Set values as required
  If (chkTablesViews.Value = vbUnchecked) Then
    chkWithColPermisions.Value = vbUnchecked
  End If
  
  'NHRD09022004 Fault 8042
  If chkBlankVersion Then
    chkUser.Enabled = False
    chkUser.Value = vbUnchecked
  Else
    chkUser.Enabled = True
  '  chkUser.Value = vbChecked
  End If
  
  'JPD 20040308 Fault 8098
  'JPD 20030515 - Only enable the OK button if the user
  ' has selected something to print.
  'NHRD28012004 Fault 6922, 6953
  cmdOK.Enabled = (chkUser.Value = vbChecked) Or _
    (chkTablesViews.Value = vbChecked) Or _
    (chkSysPerms.Value = vbChecked)

End Sub

Private Sub chkBlankVersion_Click()
  RefreshControls
End Sub

Private Sub chkSysPerms_Click()
  RefreshControls
End Sub

Private Sub chkTablesViews_Click()
  RefreshControls
End Sub

Private Sub chkUser_Click()
  RefreshControls
End Sub

Private Sub chkWithColPermisions_Click()
  RefreshControls
End Sub

Private Sub cmdCancel_Click()
  mfCancelled = True
  Unload Me
End Sub


Private Sub cmdOK_Click()
    
  gasPrintOptions(1).PrintLPaneUSERS = (chkUser.Value = vbChecked)
  gasPrintOptions(1).PrintLPaneTABLESVIEWS = (chkTablesViews.Value = vbChecked)
  gasPrintOptions(1).PrintLPaneTABLE = (chkWithColPermisions.Value = vbChecked)
  gasPrintOptions(1).PrintLPaneSYSTEM = (chkSysPerms.Value = vbChecked)
  '
  gasPrintOptions(1).PrintRPaneUSERS = (chkUser.Value = vbChecked)
  gasPrintOptions(1).PrintRPaneTABLESVIEWS = (chkTablesViews.Value = vbChecked)
  gasPrintOptions(1).PrintRPaneTABLE = (chkWithColPermisions.Value = vbChecked)
  gasPrintOptions(1).PrintRPaneSYSTEM = (chkSysPerms.Value = vbChecked)
  '
  gasPrintOptions(1).PrintBlankVersion = (chkBlankVersion.Value = vbChecked)

  Unload Me
  
End Sub

Private Sub Form_Initialize()
  mbPrintUserLogins = True
  mbPrintDataPerms = True
  mbPrintColumnPerms = False
  mbPrintSystemPerms = True
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
  
  Call ResetPrintArray(1, False)
  
  'Some need to be set as true as default
  gasPrintOptions(1).PrintLPaneUSERS = mbPrintUserLogins
  gasPrintOptions(1).PrintLPaneTABLESVIEWS = mbPrintDataPerms
  gasPrintOptions(1).PrintLPaneTABLE = mbPrintColumnPerms
  gasPrintOptions(1).PrintLPaneSYSTEM = mbPrintSystemPerms
  '
  gasPrintOptions(1).PrintRPaneUSERS = mbPrintUserLogins
  gasPrintOptions(1).PrintRPaneTABLESVIEWS = mbPrintDataPerms
  gasPrintOptions(1).PrintRPaneTABLE = mbPrintColumnPerms
  gasPrintOptions(1).PrintRPaneSYSTEM = mbPrintSystemPerms

  chkUser.Value = IIf(mbPrintUserLogins = True, vbChecked, vbUnchecked)
  chkTablesViews.Value = IIf(mbPrintDataPerms = True, vbChecked, vbUnchecked)
  chkWithColPermisions.Value = IIf(mbPrintColumnPerms = True, vbChecked, vbUnchecked)
  chkSysPerms.Value = IIf(mbPrintSystemPerms = True, vbChecked, vbUnchecked)

End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    mfCancelled = True
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Public Sub SetDefaultOptions(pbUserLogins As Boolean, pbDataPermissions As Boolean, pbColumnPermissions As Boolean, pbSystemPermissions As Boolean)
  mbPrintUserLogins = pbUserLogins
  mbPrintDataPerms = pbDataPermissions
  mbPrintColumnPerms = pbColumnPermissions
  mbPrintSystemPerms = pbSystemPermissions
End Sub

