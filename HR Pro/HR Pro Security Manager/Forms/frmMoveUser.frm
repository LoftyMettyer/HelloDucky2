VERSION 5.00
Begin VB.Form frmMoveUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Users"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8028
   Icon            =   "frmMoveUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1335
      TabIndex        =   3
      Top             =   1275
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2595
      TabIndex        =   2
      Top             =   1275
      Width           =   1200
   End
   Begin VB.ComboBox cboSecurityGroups 
      Height          =   315
      Left            =   180
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   690
      Width           =   3645
   End
   Begin VB.Label Label1 
      Caption         =   "Move selected user(s) to security group..."
      Height          =   270
      Left            =   195
      TabIndex        =   1
      Top             =   225
      Width           =   3675
   End
End
Attribute VB_Name = "frmMoveUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mbCancelled As Boolean
Private mstrMoveToGroup As String

Private Sub cboSecurityGroups_Click()

  mstrMoveToGroup = cboSecurityGroups.Text

End Sub

Private Sub cmdCancel_Click()

  mbCancelled = True
  Unload Me

End Sub

Private Sub cmdOK_Click()

  mbCancelled = False
  Unload Me

End Sub

Private Sub Form_Activate()
'NHRD07032003 Fault 3378
  For i = 0 To cboSecurityGroups.ListCount
    If cboSecurityGroups.List(i) = mstrMoveToGroup Then
      cboSecurityGroups.RemoveItem (i)
    End If
  Next
  
  cboSecurityGroups.ListIndex = 0
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

  mbCancelled = True

  ' Load the security combo box
  PopulateSecurityGroupsCombo

End Sub

Private Sub PopulateSecurityGroupsCombo()

  Dim objSecurityGroup As SecurityGroup
  
  cboSecurityGroups.Clear
  For Each objSecurityGroup In gObjGroups
    'TM20030116 Fault 4717 - don't add to list if it has been deleted.
    
    If Not objSecurityGroup.DeleteGroup Then
        cboSecurityGroups.AddItem objSecurityGroup.Name
    End If
  Next objSecurityGroup

End Sub

Public Property Get Cancelled() As Boolean

  Cancelled = mbCancelled

End Property

Public Property Get MoveToGroupName() As String

  MoveToGroupName = mstrMoveToGroup

End Property

Public Property Let MoveToGroupName(ByVal pstrNewValue As String)
  mstrMoveToGroup = pstrNewValue
  cboSecurityGroups.Text = pstrNewValue
End Property

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


