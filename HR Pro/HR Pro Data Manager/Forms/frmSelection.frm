VERSION 5.00
Begin VB.Form frmSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      Picture         =   "frmSelection.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   135
      Width           =   480
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2100
      Left            =   555
      TabIndex        =   9
      Top             =   660
      Width           =   5535
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Cancel"
         Height          =   400
         Index           =   0
         Left            =   2625
         TabIndex        =   5
         Top             =   1560
         Width           =   1200
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Delete"
         Default         =   -1  'True
         Height          =   400
         Index           =   1
         Left            =   1290
         TabIndex        =   4
         Top             =   1560
         Width           =   1200
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "All entries (that the current user has permission to see)"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   3
         Top             =   1110
         Width           =   5205
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "All entries currently displayed"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   810
         Width           =   4755
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Only the currently highlighted row(s)"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   4305
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "The highlighted row(s) and any copies of these events"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   1
         Top             =   510
         Width           =   5070
      End
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You have opted to delete entries from the "
      Height          =   195
      Left            =   735
      TabIndex        =   8
      Top             =   150
      Width           =   3075
   End
   Begin VB.Label lblQuestion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please make a selection from the options below :"
      Height          =   195
      Left            =   735
      TabIndex        =   7
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintAnswer As Integer
Private lngOptionCount As Long

Private Sub cmdAction_Click(Index As Integer)

  Dim intCount As Integer
  
  mintAnswer = -1
  If Index = 1 Then
    For intCount = optSelection.LBound To optSelection.UBound
      If optSelection(intCount).Value = True Then
        mintAnswer = intCount
        Exit For
      End If
    Next
  End If

  Me.Hide

End Sub

Public Property Get Answer() As Variant
  Answer = mintAnswer
End Property


Public Property Let OptionCount(ByVal lngNewValue As Long)
  lngOptionCount = lngNewValue
End Property


Private Sub Form_Activate()

'MH20041202 Fault 9580
'''MH20040129 Fault 8002
  If lngOptionCount <> 4 Then
    Me.Height = Me.Height - 300
    optSelection(1).Top = optSelection(1).Top - 300
    optSelection(2).Top = optSelection(2).Top - 300
    optSelection(3).Visible = False
    cmdAction(0).Top = cmdAction(0).Top - 300
    cmdAction(1).Top = cmdAction(1).Top - 300
  End If

End Sub

Public Property Let Source(ByVal strNewValue As String)
  Me.Caption = strNewValue & " Selection"
  lblSource = "You have opted to delete entries from the " & strNewValue
End Property

Private Sub Form_Load()
  mintAnswer = -1
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



