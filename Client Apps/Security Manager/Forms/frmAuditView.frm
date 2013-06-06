VERSION 5.00
Begin VB.Form frmAuditView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audit Log Columns"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8011
   Icon            =   "frmAuditView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAccess 
      Caption         =   "Audit Columns :"
      Height          =   2775
      Left            =   180
      TabIndex        =   24
      Top             =   4170
      Width           =   2550
      Begin VB.CheckBox chkAction 
         Caption         =   "&Action"
         Height          =   315
         Index           =   3
         Left            =   300
         TabIndex        =   30
         Tag             =   "A5"
         Top             =   2265
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkAction 
         Caption         =   "&OpenHR Module"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   29
         Tag             =   "A4"
         Top             =   1900
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.CheckBox chkLogin 
         Caption         =   "Co&mputer Name"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   28
         Tag             =   "A3"
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "User &Name"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   27
         Tag             =   "A2"
         Top             =   1100
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "User &Group"
         Height          =   315
         Left            =   300
         TabIndex        =   26
         Tag             =   "A1"
         Top             =   700
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "&Date / Time"
         Height          =   315
         Index           =   3
         Left            =   300
         TabIndex        =   25
         Tag             =   "A0"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Frame fraGroups 
      Caption         =   "Audit Columns :"
      Height          =   2400
      Left            =   5800
      TabIndex        =   23
      Top             =   100
      Width           =   2550
      Begin VB.CheckBox chkUser 
         Caption         =   "&User"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Tag             =   "G0"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1100
      End
      Begin VB.CheckBox Check4 
         Caption         =   "&Date / Time"
         Height          =   315
         Left            =   300
         TabIndex        =   15
         Tag             =   "G1"
         Top             =   700
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "User &Group"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   16
         Tag             =   "G2"
         Top             =   1100
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox chkLogin 
         Caption         =   "User &Login"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   17
         Tag             =   "G3"
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkAction 
         Caption         =   "&Action"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   18
         Tag             =   "G4"
         Top             =   1900
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   150
      TabIndex        =   19
      Top             =   3500
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1500
      TabIndex        =   20
      Top             =   3500
      Width           =   1200
   End
   Begin VB.Frame fraData 
      Caption         =   "Audit Columns :"
      Height          =   3200
      Left            =   3000
      TabIndex        =   21
      Top             =   100
      Width           =   2550
      Begin VB.CheckBox chkDesc 
         Caption         =   "&Record Description"
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Tag             =   "D6"
         Top             =   2700
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin VB.CheckBox chkNewValue 
         Caption         =   "&New Value"
         Height          =   315
         Left            =   300
         TabIndex        =   12
         Tag             =   "D5"
         Top             =   2300
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkOldValue 
         Caption         =   "O&ld Value"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Tag             =   "D4"
         Top             =   1900
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Colu&mn"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Tag             =   "D3"
         Top             =   1485
         Value           =   1  'Checked
         Width           =   1300
      End
      Begin VB.CheckBox chkTable 
         Caption         =   "&Table"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Tag             =   "D2"
         Top             =   1100
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox chkDateTime 
         Caption         =   "&Date / Time"
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Tag             =   "D1"
         Top             =   700
         Value           =   1  'Checked
         Width           =   1470
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "&User"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Tag             =   "D0"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1100
      End
   End
   Begin VB.Frame fraPermissions 
      Caption         =   "Audit Columns :"
      Height          =   3200
      Left            =   150
      TabIndex        =   22
      Top             =   100
      Width           =   2550
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Permission"
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Tag             =   "P6"
         Top             =   2700
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.CheckBox chkAction 
         Caption         =   "&Action"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Tag             =   "P5"
         Top             =   2300
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "&User"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   0
         Tag             =   "P0"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1100
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "&Date / Time"
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Tag             =   "P1"
         Top             =   700
         Value           =   1  'Checked
         Width           =   1470
      End
      Begin VB.CheckBox chkTable 
         Caption         =   "&Table"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   3
         Tag             =   "P3"
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Colu&mn"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   4
         Tag             =   "P4"
         Top             =   1900
         Value           =   1  'Checked
         Width           =   1300
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "User &Group"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Tag             =   "P2"
         Top             =   1100
         Value           =   1  'Checked
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmAuditView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private maudType As audType
Private mbCancelled As Boolean

Public Sub Initialise(AuditType As audType, avView() As Boolean)
  ' Display the appropriate column selection controls.
  Dim dblCurrentHeight As Double
  
  Const iFRAMETOP = 100
  Const iFRAMELEFT = 150
  Const iYGAP = 200
  Const iFORMWIDTH = 2925
  
  With fraData
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audRecords)
    
    If (AuditType = audRecords) Then
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  With fraPermissions
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audPermissions)
    
    If (AuditType = audPermissions) Then
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  With fraGroups
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audGroups)
    
    If (AuditType = audGroups) Then
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  With fraAccess
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audAccess)
    
    If (AuditType = audAccess) Then
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  cmdCancel.Top = dblCurrentHeight + iYGAP
  cmdOK.Top = cmdCancel.Top
    
  Me.Height = cmdCancel.Top + cmdCancel.Height + iYGAP + UI.CaptionHeight + (2 * UI.YFrame)
  Me.Width = iFORMWIDTH
  
  maudType = AuditType
  SetupDetails avView()
  
End Sub

Private Sub cmdCancel_Click()

    mbCancelled = True
    Me.Hide

End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled

End Property

Private Sub cmdOK_Click()

    mbCancelled = False
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        mbCancelled = True
        Cancel = True
        Me.Hide
    End If

End Sub

Public Sub GetDetails(mbView() As Boolean)
  Dim ctlTemp As Control
  Dim sType As String
  Dim lngCount As Long
  
  Select Case maudType
    Case audRecords
      sType = "D"
    Case audPermissions
      sType = "P"
    Case audGroups
      sType = "G"
    Case audAccess
      sType = "A"
  End Select


  'MH20010823 Fault 2272 "ColumnID" column was appearing!
  'make sure that all columns are not visible (unless specified)
  For lngCount = LBound(mbView) To UBound(mbView)
    mbView(lngCount) = False
  Next
  
  For Each ctlTemp In Controls
    If TypeOf ctlTemp Is CheckBox Then
      If Left(ctlTemp.Tag, 1) = sType Then
        mbView(CLng(Mid$(CStr(ctlTemp.Tag), 2, Len(ctlTemp.Tag)))) = (ctlTemp.Value = 1)
      End If
    End If
  Next

End Sub

Private Sub SetupDetails(avView() As Boolean)
  Dim ctlTemp As Control
  Dim sType As String
  
  Select Case maudType
    Case audRecords
      sType = "D"
    Case audPermissions
      sType = "P"
    Case audGroups
      sType = "G"
    Case audAccess
      sType = "A"
  End Select
  
  For Each ctlTemp In Controls
    If TypeOf ctlTemp Is CheckBox Then
      If Left(ctlTemp.Tag, 1) = sType Then
        ctlTemp.Value = IIf(avView(CLng(Mid$(ctlTemp.Tag, 2, Len(ctlTemp.Tag)))), 1, 0)
      End If
    End If
  Next

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


