VERSION 5.00
Begin VB.Form frmFindUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find User"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8064
   Icon            =   "frmFindUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   2940
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Locate"
      Height          =   400
      Left            =   1650
      TabIndex        =   1
      Top             =   765
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2910
      TabIndex        =   2
      Top             =   765
      Width           =   1200
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login :"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "frmFindUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbCancelled As Boolean
Private mstrSelectedUser As String

Public Function Initialise() As Boolean

  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser

  Screen.MousePointer = vbHourglass
  
  For Each objGroup In gObjGroups
  
    'MH20061207 Fault 11768
    If Not objGroup.DeleteGroup Then
    
      ' Load the users into collections...may take a while
      If Not gObjGroups(objGroup.Name).Users_Initialised Then
        InitialiseUsersCollection gObjGroups(objGroup.Name)
      End If
    
      For Each objUser In gObjGroups(objGroup.Name).Users
        If (Not objUser.DeleteUser) And _
          (objUser.MovedUserTo = "") Then
            cboUser.AddItem objUser.UserName
        End If
      Next objUser
  
    End If
  
  Next objGroup

  Screen.MousePointer = vbNormal

  If cboUser.ListCount > 0 Then
    cboUser.ListIndex = 0
    Initialise = True
  Else
    MsgBox "There are no valid HR Pro users in the database.", vbExclamation + vbOKOnly, App.Title
    Initialise = False
  End If

End Function

Public Property Get SelectedUser() As String
  SelectedUser = mstrSelectedUser
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Private Sub cboUser_Change()
  mstrSelectedUser = cboUser.Text
End Sub

Private Sub cboUser_Click()
  mstrSelectedUser = cboUser.Text
End Sub

Private Sub cmdCancel_Click()
  mbCancelled = True
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  mbCancelled = False
  Unload Me
End Sub

Private Sub Form_Load()
  mbCancelled = True
End Sub
