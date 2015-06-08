VERSION 5.00
Begin VB.Form frmExportHeaderFooter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Header"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportHeaderFooter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHeaderFooter 
      Height          =   3240
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmExportHeaderFooter.frx":000C
      Top             =   120
      Width           =   9270
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6975
      TabIndex        =   1
      Top             =   4965
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8235
      TabIndex        =   0
      Top             =   4965
      Width           =   1035
   End
   Begin VB.Label lblCodes 
      Caption         =   "{BLOCKCOUNT}   Number of records in current block"
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   7
      Top             =   4665
      Width           =   5130
   End
   Begin VB.Label lblCodes 
      Caption         =   "{BLOCKNUMBER}   File Block Number"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   4395
      Width           =   3555
   End
   Begin VB.Label lblCodes 
      Caption         =   "{DATETIME}    Date and time of export"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   4125
      Width           =   2925
   End
   Begin VB.Label lblCodes 
      Caption         =   "{DATE}    Date of export"
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3855
      Width           =   2925
   End
   Begin VB.Label lblCodes 
      Caption         =   "{TOTALCOUNT}    Total Record Count"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   3585
      Width           =   3690
   End
End
Attribute VB_Name = "frmExportHeaderFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancelled As Boolean
Public Text As String
Public IsHeader As Boolean

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mbCancelled = bCancel
End Property

Public Sub Initialise()
  txtHeaderFooter.Text = Text
  
  If IsHeader Then
    Me.Caption = "Custom Header"
  Else
    Me.Caption = "Custom Footer"
  End If
  
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Text = txtHeaderFooter.Text
  Unload Me
End Sub

Private Sub txtHeaderFooter_GotFocus()
  ' When Text2 gets the focus, clear all TabStop properties on all
  ' controls on the form. Ignore all errors, in case a control does
  ' not have the TabStop property.
  On Error Resume Next
  Dim i As Integer
  For i = 0 To Controls.Count - 1   ' Use the Controls collection
    Controls(i).TabStop = False
  Next
End Sub

Private Sub txtHeaderFooter_LostFocus()
  ' When Text2 loses the focus, make the TabStop property True for all
  ' controls on the form. That restores the ability to tab between
  ' controls. Ignore all errors, in case a control does not have the
  ' TabStop property.
  On Error Resume Next
  Dim i As Integer
  For i = 0 To Controls.Count - 1   ' Use the Controls collection
    Controls(i).TabStop = True
  Next
End Sub
