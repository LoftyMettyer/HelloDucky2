VERSION 5.00
Begin VB.Form frmLockMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lock Message"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1046
   Icon            =   "frmLockMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMessage 
      Height          =   1500
      Left            =   200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   500
      Width           =   4000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   1600
      TabIndex        =   1
      Top             =   2150
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3000
      TabIndex        =   2
      Top             =   2150
      Width           =   1200
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "Message :"
      Height          =   195
      Left            =   200
      TabIndex        =   3
      Top             =   200
      Width           =   735
   End
End
Attribute VB_Name = "frmLockMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private mfCancelled As Boolean
Private msMessage As String

Private Sub cmdCancel_Click()
  ' Flag that the message has been cancelled..
  mfCancelled = True
  
  ' Unload the form.
  Unload Me

End Sub

Private Sub cmdOK_Click()
  ' Flag that the message has been confirmed..
  mfCancelled = False
  
  ' Trimm message to maximum of 200 otherwise it breaks the savesystemsetting
  msMessage = Left(Trim(txtMessage.Text), 200)
  
  ' Replace single quotation marks with two single quotation marks.
  SaveSystemSetting "Messaging", "LockMessage", Replace(msMessage, "'", "''")
  
  ' Unload the form.
  Unload Me

End Sub

Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = mfCancelled
End Property
  
Public Property Get Message() As String
  Message = msMessage
End Property
  
Private Sub Form_Activate()
  mfCancelled = True  'capture closing form with "X"
End Sub

Private Sub Form_Load()
  txtMessage.Text = GetSystemSetting("Messaging", "LockMessage", "")
  txtMessage.SelStart = 0
  txtMessage.SelLength = Len(txtMessage.Text)
  
End Sub

Private Sub Form_Resize()
  DisplayApplication
End Sub


