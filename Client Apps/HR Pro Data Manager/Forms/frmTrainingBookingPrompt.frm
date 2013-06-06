VERSION 5.00
Begin VB.Form frmTrainingBookingPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Course"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1010
   Icon            =   "frmTrainingBookingPrompt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optStatus 
      Caption         =   "&Provisionally booked"
      Height          =   345
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   900
      Width           =   2160
   End
   Begin VB.OptionButton optStatus 
      Caption         =   "&Booked"
      Height          =   345
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   500
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1455
      TabIndex        =   3
      Top             =   1395
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   195
      TabIndex        =   2
      Top             =   1395
      Width           =   1200
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the required booking status :"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   2565
   End
End
Attribute VB_Name = "frmTrainingBookingPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide
  
End Sub


Private Sub cmdOK_Click()
  mfCancelled = False
  Me.Hide

End Sub


Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property
Public Property Get Booked() As Boolean
  Booked = (optStatus(0).Value = True)
  
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



