VERSION 5.00
Begin VB.UserControl COAWF_LinkArrow 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MaskColor       =   &H000000FF&
   ScaleHeight     =   1170
   ScaleWidth      =   675
   ToolboxBitmap   =   "COAWF_LinkArrow.ctx":0000
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   3
      Left            =   360
      Picture         =   "COAWF_LinkArrow.ctx":0312
      ScaleHeight     =   120
      ScaleWidth      =   135
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   360
      Picture         =   "COAWF_LinkArrow.ctx":0434
      ScaleHeight     =   135
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   360
      Picture         =   "COAWF_LinkArrow.ctx":054E
      ScaleHeight     =   135
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   0
      Left            =   360
      Picture         =   "COAWF_LinkArrow.ctx":0668
      ScaleHeight     =   120
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   3
      Left            =   120
      Picture         =   "COAWF_LinkArrow.ctx":078A
      ScaleHeight     =   120
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   120
      Picture         =   "COAWF_LinkArrow.ctx":08AC
      ScaleHeight     =   135
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   120
      Picture         =   "COAWF_LinkArrow.ctx":09C6
      ScaleHeight     =   135
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   0
      Left            =   120
      Picture         =   "COAWF_LinkArrow.ctx":0AE0
      ScaleHeight     =   120
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "COAWF_LinkArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ArrowDirection
  arrowDirection_Down = 0
  arrowDirection_Left = 1
  arrowDirection_Right = 2
  arrowDirection_Up = 3
End Enum

Private miArrowDirection As ArrowDirection

' Declare public events.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Property Get ArrowDirection() As ArrowDirection
  ' Return the current ArrowDirection
  ArrowDirection = miArrowDirection
  
End Property

Public Property Get ArrowPicture() As StdPicture
  ' Return the current Arrow picture
  Set ArrowPicture = picPicture(miArrowDirection).Picture
  
End Property


Public Property Let ArrowDirection(ByVal piNewValue As ArrowDirection)
  ' Set the direction.
  miArrowDirection = piNewValue
  PropertyChanged "ArrowDirection"

  ' Display the appropriate picture.
  DrawArrow

End Property
Private Sub DrawArrow()
  ' Display the appropriate picture.
    
  With UserControl
    .Picture = picPicture(miArrowDirection).Picture
    .MaskPicture = picMask(miArrowDirection).Picture
  End With
  
  ResizeControl
  
End Sub

Private Sub ResizeControl()
  ' Resize to the current picture.
  With UserControl
    .Height = picPicture(miArrowDirection).Height
    .Width = picPicture(miArrowDirection).Width
  End With
    
End Sub


Public Sub About()
Attribute About.VB_UserMemId = -552
  ' Display the 'About' box.
  MsgBox App.ProductName & " - " & App.FileDescription & _
    vbCr & vbCr & App.LegalCopyright, _
    vbOKOnly, "About " & App.ProductName
    
End Sub

Private Sub picMask_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub


Private Sub picPicture_DblClick(Index As Integer)
  ' Pass the DblClick event to the parent form.
  RaiseEvent DblClick

End Sub

Private Sub picPicture_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub


Private Sub picPicture_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub picPicture_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub picPicture_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub UserControl_DblClick()
  ' Pass the DblClick event to the parent form.
  RaiseEvent DblClick

End Sub

Private Sub UserControl_InitProperties()
  ' Initialise the properties.
  On Error Resume Next
  
  ArrowDirection = arrowDirection_Down

End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass the KeyDown event to the parent form.
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Load property values from storage.
  On Error Resume Next

  ' Read the previous set of properties.
  ArrowDirection = PropBag.ReadProperty("ArrowDirection", arrowDirection_Down)

End Sub


Private Sub UserControl_Resize()
  ResizeControl

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  
  ' Save the current set of properties.
  Call PropBag.WriteProperty("ArrowDirection", miArrowDirection, arrowDirection_Down)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

End Sub


