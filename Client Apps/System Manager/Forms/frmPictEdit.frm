VERSION 5.00
Begin VB.Form frmPictEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picture Properties"
   ClientHeight    =   2655
   ClientLeft      =   825
   ClientTop       =   4890
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5021
   Icon            =   "frmPictEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWidth 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1000
      MaxLength       =   50
      TabIndex        =   10
      Text            =   "txtWidth"
      Top             =   1700
      Width           =   750
   End
   Begin VB.TextBox txtHeight 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1000
      MaxLength       =   50
      TabIndex        =   9
      Text            =   "txtHeight"
      Top             =   1200
      Width           =   750
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1000
      MaxLength       =   50
      TabIndex        =   8
      Text            =   "txtType"
      Top             =   700
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Height          =   1500
      Left            =   4300
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   200
      Width           =   1500
      Begin VB.Image Image1 
         Height          =   1410
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4600
      TabIndex        =   2
      Top             =   2100
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3200
      TabIndex        =   1
      Top             =   2100
      Width           =   1200
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1000
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "txtName"
      Top             =   200
      Width           =   3000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width :"
      Height          =   195
      Left            =   200
      TabIndex        =   6
      Top             =   1760
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height :"
      Height          =   195
      Left            =   200
      TabIndex        =   5
      Top             =   1260
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      Height          =   195
      Left            =   200
      TabIndex        =   4
      Top             =   760
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   200
      TabIndex        =   3
      Top             =   260
      Width           =   510
   End
End
Attribute VB_Name = "frmPictEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarCancelled As Boolean
Private mvarName As String

'Private strFileName As String

Public Property Get Cancelled() As Boolean
  Cancelled = mvarCancelled
End Property
Public Sub EditMenu(ByVal psMenuOption As String)
End Sub

Public Property Set PictureObj(NewPicture As Object)
  Set Image1.Picture = NewPicture
  SizeImage Image1
  
  Image1.Top = (Picture1.ScaleHeight - Image1.Height) \ 2
  Image1.Left = (Picture1.ScaleWidth - Image1.Width) \ 2
  
  txtType.Text = GetPictureType(Image1.Picture)
  txtHeight.Text = Trim(Str(Int(Me.ScaleY(Image1.Picture.Height, vbHimetric, vbPixels))))
  txtWidth.Text = Trim(Str(Int(Me.ScaleX(Image1.Picture.Width, vbHimetric, vbPixels))))
End Property

Public Property Get PictureName() As String
  PictureName = mvarName
End Property

Public Property Let PictureName(NewName As String)
  mvarName = NewName
  If txtName.Text <> NewName Then
    txtName.Text = NewName
  End If
End Property

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdOK_Click()
  mvarCancelled = False
  
  UnLoad Me
End Sub

Private Sub Form_Activate()

  ' Set focus on the first control.
  txtName.SetFocus

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
  
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  ' Initialize variables.
  mvarCancelled = True
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr
  
  ' Position the contained controls.
  Image1.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
  
End Sub



Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub txtName_Change()
  
  PictureName = txtName.Text

End Sub

Private Sub txtName_DblClick()
  
  ' Select all of the text in the textbox.
  UI.txtSelText

End Sub

Private Sub txtName_GotFocus()

  ' Select all of the text in the textbox.
  UI.txtSelText

End Sub


