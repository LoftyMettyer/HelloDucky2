VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLabelsPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Label start position"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1081
   Icon            =   "frmLabelsPrompt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2115
      TabIndex        =   1
      Top             =   1140
      Width           =   1200
   End
   Begin VB.TextBox txtStartColumn 
      Height          =   315
      Left            =   2655
      TabIndex        =   3
      Text            =   "1"
      Top             =   645
      Width           =   420
   End
   Begin MSComCtl2.UpDown upnStartColumn 
      Height          =   315
      Left            =   3075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   645
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtStartColumn"
      BuddyDispid     =   196610
      OrigLeft        =   2850
      OrigTop         =   690
      OrigRight       =   3105
      OrigBottom      =   1005
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown upnStartRow 
      Height          =   315
      Left            =   3075
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   195
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtStartRow"
      BuddyDispid     =   196614
      OrigLeft        =   2520
      OrigTop         =   120
      OrigRight       =   2775
      OrigBottom      =   435
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   810
      TabIndex        =   0
      Top             =   1140
      Width           =   1200
   End
   Begin VB.TextBox txtStartRow 
      Height          =   315
      Left            =   2655
      TabIndex        =   2
      Text            =   "1"
      Top             =   195
      Width           =   420
   End
   Begin VB.Label lblStartColumn 
      Caption         =   "Start position at column :"
      Height          =   210
      Left            =   255
      TabIndex        =   5
      Top             =   705
      Width           =   2265
   End
   Begin VB.Label lblStartingRow 
      Caption         =   "Start position at row :"
      Height          =   225
      Left            =   255
      TabIndex        =   4
      Top             =   240
      Width           =   1905
   End
End
Attribute VB_Name = "frmLabelsPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim miStartRow As Integer
Dim iStartColumn As Integer
Dim mbCancel As Boolean

Public Property Get StartRow() As Long
  StartRow = Val(txtStartRow.Text)
End Property

Public Property Get StartColumn() As Long
  StartColumn = Val(txtStartColumn.Text)
End Property

Public Property Let MaximumRows(piMaximumRows As Integer)
  upnStartRow.Max = piMaximumRows
End Property

Public Property Let MaximumColumns(piMaximumColumns As Integer)
  upnStartColumn.Max = piMaximumColumns
End Property

Private Sub cmdCancel_Click()
  mbCancel = True
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  mbCancel = False
  Me.Hide
End Sub

Private Sub Form_Load()
  mbCancel = False
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub txtStartColumn_GotFocus()

  With txtStartColumn
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub

Private Sub txtStartColumn_KeyPress(KeyAscii As Integer)

  If Not IsNumeric(Chr(KeyAscii)) Then
    KeyAscii = 0
  End If

End Sub

Private Sub txtStartColumn_LostFocus()

  If Val(txtStartColumn.Text) > upnStartColumn.Max Then
    txtStartColumn.Text = upnStartColumn.Max
  End If

  If Val(txtStartColumn.Text) < upnStartColumn.Min Then
    txtStartColumn.Text = upnStartColumn.Min
  End If

End Sub

Private Sub txtStartRow_GotFocus()

  With txtStartRow
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub

Private Sub txtStartRow_KeyPress(KeyAscii As Integer)

  If Not IsNumeric(Chr(KeyAscii)) Then
    KeyAscii = 0
  End If

End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancel
End Property

Private Sub txtStartRow_LostFocus()

  If Val(txtStartRow.Text) > upnStartRow.Max Then
    txtStartRow.Text = upnStartRow.Max
  End If

  If Val(txtStartRow.Text) < upnStartRow.Min Then
    txtStartRow.Text = upnStartRow.Min
  End If

End Sub

