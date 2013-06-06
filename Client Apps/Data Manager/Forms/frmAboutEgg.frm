VERSION 5.00
Begin VB.Form frmAboutEgg 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   7125
   ClientLeft      =   75
   ClientTop       =   -495
   ClientWidth     =   11655
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLives 
      Height          =   375
      Left            =   9600
      MaxLength       =   2
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtLevel 
      Height          =   375
      Left            =   9600
      MaxLength       =   2
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEnemyMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   10560
      Picture         =   "frmAboutEgg.frx":0000
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picEnemyGrey 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11280
      Picture         =   "frmAboutEgg.frx":0DC2
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picEnemyYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11280
      Picture         =   "frmAboutEgg.frx":0F14
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picEnemyBlue 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10560
      Picture         =   "frmAboutEgg.frx":1211
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2490
      Left            =   10680
      Picture         =   "frmAboutEgg.frx":153D
      ScaleHeight     =   162
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   13
      Top             =   3000
      Width           =   1650
   End
   Begin VB.PictureBox picEnemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10560
      Picture         =   "frmAboutEgg.frx":2BD6
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   11160
      Picture         =   "frmAboutEgg.frx":2ED9
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox picBonusMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   11520
      Picture         =   "frmAboutEgg.frx":31EB
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox picBackGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8040
      Left            =   10800
      ScaleHeight     =   532
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   564
      TabIndex        =   9
      Top             =   5280
      Width           =   8520
   End
   Begin VB.PictureBox picPlayerMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10560
      Picture         =   "frmAboutEgg.frx":34FD
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox picPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   10560
      Picture         =   "frmAboutEgg.frx":40AB
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.PictureBox picBallMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   10800
      Picture         =   "frmAboutEgg.frx":4DB1
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   180
   End
   Begin VB.PictureBox picBall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   10560
      Picture         =   "frmAboutEgg.frx":4FA3
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   180
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   8760
      Top             =   1800
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6960
      Left            =   0
      ScaleHeight     =   464
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   0
      Top             =   0
      Width           =   8160
   End
   Begin VB.Label lblCheatLives 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lives:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   8520
      TabIndex        =   27
      Top             =   5160
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblCheatLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   8520
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASRkaniod"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   10120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   255
      Left            =   9840
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9900
      TabIndex        =   21
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLevelText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L e v e l"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   8640
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblNewGame 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8580
      TabIndex        =   14
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   495
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblLives 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lives :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CA8826&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      FillColor       =   &H00C0FFFF&
      Height          =   6975
      Left            =   8190
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmAboutEgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
 
Private Type PlayerMoveDirection
  MoveLeft As Boolean
  MoveRight As Boolean
End Type

Private Type Rect
  Left As Long
  Top As Long
  Width As Long
  Height As Long
End Type

Private Type EnemyType
  Bonus As BonusEnum
  Alive As Boolean
  EnemyRect As Rect
  Color As EnemyColor
End Type

Private Enum EnemyColor
  Red = 0
  Blue = 1
  Yellow = 2
  Grey = 3
End Enum

Private Enum BonusEnum
  None = 0
  Large = 1
  Live = 2
  FireBall = 3
  'Glue = 2
End Enum

Private Type TypeBonus
  BonusRect As Rect
  BonusType As BonusEnum
  BonusNotExist As Boolean
End Type

Private Const PlayerStep = 15
Private Const BonusStep = 5

Private PlayerDirection As PlayerMoveDirection
Private PlayerRect As Rect
Private BallRect As Rect
Private Enemy() As EnemyType
Private EnemyCount As Integer
Private Bonuses() As TypeBonus
Private BonusCount As Integer

Private FormWidth As Single
Private FormHeight As Single

Private BallVelocityX As Long
Private BallVelocityY As Long

Private ENEMY_WIDTH As Single
Private ENEMY_HEIGHT As Single

Private StartGame As Boolean
Private LiveNumber As Integer
Private LevelNumber As Integer
Private Score As Long

Private Sub DrawBitmap(PrimaryPicture As PictureBox, MaskPictureHDC As Long, BackGroungPictureHDC As Long, ByVal LeftPos As Long, ByVal TopPos As Long)
  Dim lngReturn As Long
  lngReturn = BitBlt(BackGroungPictureHDC, LeftPos, TopPos, PrimaryPicture.ScaleWidth, PrimaryPicture.ScaleHeight, MaskPictureHDC, 0, 0, vbSrcAnd)
  lngReturn = BitBlt(BackGroungPictureHDC, LeftPos, TopPos, PrimaryPicture.ScaleWidth, PrimaryPicture.ScaleHeight, PrimaryPicture.hDC, 0, 0, vbSrcPaint)
End Sub

Private Sub PlaySound(ByVal Index As Long)
  Dim sSoundBuffer As String
  Dim Ret As Long
  sSoundBuffer = StrConv(LoadResData(Index, "SOUND"), vbUnicode)
  Ret = sndPlaySound(sSoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_NOSTOP)
End Sub

Private Function Collision(Rect1 As Rect, Rect2 As Rect) As Boolean
  Collision = False
  If Rect1.Left + Rect1.Width > Rect2.Left Then
    If Rect1.Left < Rect2.Left + Rect2.Width Then
      If Rect1.Top + Rect1.Height > Rect2.Top Then
        If Rect1.Top < Rect2.Top + Rect2.Height Then
          Collision = True
        End If
      End If
    End If
  End If
End Function

Private Function BallPlayerCollision(BallRect1 As Rect, PlayerRect2 As Rect) As Boolean
  BallPlayerCollision = False
  If BallRect1.Top + BallRect1.Height + 1 > PlayerRect2.Top Then
    If BallRect1.Left + BallRect1.Width > PlayerRect2.Left Then
      If BallRect1.Left < PlayerRect2.Left + PlayerRect2.Width Then
        BallPlayerCollision = True
      End If
    End If
  End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft Then
    PlayerDirection.MoveLeft = True
  ElseIf KeyCode = vbKeyRight Then
    PlayerDirection.MoveRight = True
  ElseIf KeyCode = vbKeyN Then
    Call mnuNewGame_Click
  ElseIf KeyCode = vbKeyEscape Then
    Unload Me
  ElseIf KeyCode = vbKeyP Then
    If StartGame = True Then
     tmrMove.Enabled = Not tmrMove.Enabled
     lblPause.Visible = Not tmrMove.Enabled
    End If
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft Then
    PlayerDirection.MoveLeft = False
  ElseIf KeyCode = vbKeyRight Then
    PlayerDirection.MoveRight = False
  ElseIf KeyCode = vbKeySpace Then
    StartGame = True
  End If
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim J As Integer
  
  Me.Width = picGame.Width + Shape1.Width + 40
  Me.Height = picGame.Height - 40
  Shape1.Top = 0
  Shape1.Height = Me.Height
  FormWidth = picGame.ScaleWidth
  FormHeight = picGame.ScaleHeight
  BallRect.Width = picBall.ScaleWidth
  BallRect.Height = picBall.ScaleHeight
  ENEMY_WIDTH = picEnemy.ScaleWidth
  ENEMY_HEIGHT = picEnemy.ScaleHeight
  
  For i = 0 To 5
    For J = 0 To 4
      BitBlt picBackGround.hDC, picBack.ScaleWidth * i, picBack.ScaleHeight * J, picBack.ScaleWidth, picBack.ScaleHeight, picBack.hDC, 0, 0, vbSrcCopy
    Next J
  Next i
  
  picBackGround.Refresh
  Call GameBackGroundRefresh

  'Shape2.Move 0, 0, (Me.ScaleWidth / Screen.TwipsPerPixelY), (Me.ScaleHeight / Screen.TwipsPerPixelX)
  'Shape2.ZOrder 0

End Sub

Private Sub PlayerMove()
  If PlayerDirection.MoveLeft = True Then
    If PlayerRect.Left - PlayerStep >= 0 Then
      PlayerRect.Left = PlayerRect.Left - PlayerStep
    Else
      PlayerRect.Left = 0
    End If
  ElseIf PlayerDirection.MoveRight = True Then
    If PlayerRect.Left + PlayerRect.Width + PlayerStep <= FormWidth Then
      PlayerRect.Left = PlayerRect.Left + PlayerStep
    Else
      PlayerRect.Left = FormWidth - PlayerRect.Width
    End If
  End If
  Call DrawBitmap(picPlayer, picPlayerMask.hDC, picGame.hDC, PlayerRect.Left, PlayerRect.Top)
End Sub

Private Sub BallMove()
  Dim i As Long
  BallRect.Left = BallRect.Left + BallVelocityX
  If BallRect.Left <= 0 Then
    BallRect.Left = 0
    BallVelocityX = -BallVelocityX
  ElseIf BallRect.Left + BallRect.Width >= FormWidth Then
    BallRect.Left = FormWidth - BallRect.Width
    BallVelocityX = -BallVelocityX
  End If
  
  BallRect.Top = BallRect.Top + BallVelocityY
  If BallRect.Top <= 0 Then
    BallRect.Top = 0
    BallVelocityY = -BallVelocityY
  ElseIf BallRect.Top + BallRect.Height >= FormHeight Then
    Beep
    PlayerDirection.MoveLeft = False
    PlayerDirection.MoveRight = False
    LiveNumber = LiveNumber - 1
    lblLives.Caption = LiveNumber
    If LiveNumber = 0 Then
      If COAMsgBox("Game Over.  Play again?", vbExclamation + vbYesNo, "ASRkanoid") = vbYes Then
        Call mnuNewGame_Click
      Else
        Unload Me
      End If
    Else
      Call SetStartPosition
    End If
  End If
  
  BallRect.Top = BallRect.Top + Abs(BallVelocityY)
  If BallPlayerCollision(BallRect, PlayerRect) = True Then
    BallRect.Top = PlayerRect.Top - BallRect.Height - 1
    BallVelocityX = BallVelocityX + ((BallRect.Left + (BallRect.Width / 2)) - (PlayerRect.Left + (PlayerRect.Width / 2))) / 3
    BallVelocityY = -BallVelocityY

    'Make the ball get gradually faster...
    BallVelocityX = BallVelocityX * 1.02
    BallVelocityY = BallVelocityY * 1.02

    Call PlaySound(2)
  Else
    BallRect.Top = BallRect.Top - Abs(BallVelocityY)
  End If
  
  For i = 0 To EnemyCount - 1
    If Enemy(i).Alive = True Then
      If Collision(Enemy(i).EnemyRect, BallRect) = True Then
        If Enemy(i).Color <> Grey Then
          If picBall.Tag <> "fast" Then
            Call ChangeVellosity(Enemy(i).EnemyRect)
          End If
          
          Score = Score + 20
          lblScore.Caption = Score
          Enemy(i).Alive = False
          If Enemy(i).Bonus <> None Then
            BonusCount = BonusCount + 1
            ReDim Preserve Bonuses(BonusCount)
            Bonuses(BonusCount).BonusRect.Width = picBonus.ScaleWidth
            Bonuses(BonusCount).BonusRect.Height = picBonus.ScaleHeight
            Bonuses(BonusCount).BonusRect.Top = Enemy(i).EnemyRect.Top + Enemy(i).EnemyRect.Height + 1
            Bonuses(BonusCount).BonusRect.Left = Enemy(i).EnemyRect.Left + Enemy(i).EnemyRect.Width / 2 - Bonuses(BonusCount).BonusRect.Width / 2
            Bonuses(BonusCount).BonusNotExist = False
            Randomize
            Bonuses(BonusCount).BonusType = Enemy(i).Bonus
          End If
          Call PlaySound(1)
        Else 'is gray
          Call ChangeVellosity(Enemy(i).EnemyRect)
        End If
        
        Exit For
      End If
    End If
  Next i
  Call DrawBitmap(picBall, picBallMask.hDC, picGame.hDC, BallRect.Left, BallRect.Top)
End Sub

Private Sub mnuNewGame_Click()
  Call SetStartPosition
  Score = 0
  lblScore.Caption = Score
  LiveNumber = 3
  lblLives.Caption = LiveNumber
  lblLevelText.Visible = True
  tmrMove.Enabled = True
  picGame.SetFocus
  LevelNumber = 1
  Call LoadLevels(LevelNumber)
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Label3_DblClick()
  lblCheatLevel.Visible = Not lblCheatLevel.Visible
  txtLevel.Visible = Not txtLevel.Visible
  txtLevel.Text = lblLevel.Caption
  lblCheatLives.Visible = Not lblCheatLives.Visible
  txtLives.Visible = Not txtLives.Visible
  txtLives.Text = lblLives.Caption
End Sub

Private Sub lblExit_Click()
  Unload Me
End Sub

Private Sub lblNewGame_Click()
  Call mnuNewGame_Click
End Sub

Private Sub txtLevel_Change()
  Dim lngLevel As Integer
  lngLevel = Val(txtLevel.Text)
  If lngLevel > 0 And lngLevel < 9 Then
    PlayerDirection.MoveLeft = False
    PlayerDirection.MoveRight = False
    Call SetStartPosition
    LevelNumber = lngLevel
    Call LoadLevels(lngLevel)
  End If
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtLives_Change()
  If Val(txtLives) > 0 Then
    lblLives.Caption = Val(txtLives)
    LiveNumber = Val(txtLives)
  End If
End Sub

Private Sub tmrMove_Timer()
  If StartGame = True Then
    Call GameBackGroundRefresh
    Call PlayerMove
    Call BallMove
    Call DrawEnemis
    If BonusCount > 0 Then
      Call BonusMove
    End If
    picGame.Refresh
  Else
    Call GameBackGroundRefresh
    Call PlayerMove
    Call MoveBallWithPlayer
    Call DrawEnemis
    picGame.Refresh
  End If
End Sub

Private Sub MoveBallWithPlayer()
  BallRect.Left = PlayerRect.Left + PlayerRect.Width / 2 - BallRect.Width / 2
  BallRect.Top = PlayerRect.Top - BallRect.Height - 1
  Call DrawBitmap(picBall, picBallMask.hDC, picGame.hDC, BallRect.Left, BallRect.Top)
End Sub

Private Sub GameBackGroundRefresh()
  Dim lngResponce As Long
  lngResponce = BitBlt(picGame.hDC, 1, 1, FormWidth, FormHeight - 5, picBackGround.hDC, 0, 0, vbSrcCopy)
End Sub

Private Sub SetStartPosition()
  Set picPlayer.Picture = LoadResPicture("Player", 0)
  Set picPlayerMask.Picture = LoadResPicture("PlayerMask", 0)
  PlayerRect.Width = picPlayer.ScaleWidth
  PlayerRect.Height = picPlayer.ScaleHeight
  PlayerRect.Left = picGame.ScaleWidth / 2 - PlayerRect.Width / 2
  PlayerRect.Top = picGame.ScaleHeight - PlayerRect.Height - 5
  BallRect.Left = PlayerRect.Left + PlayerRect.Width / 2 - BallRect.Width / 2
  BallRect.Top = PlayerRect.Top - BallRect.Height
  BallVelocityX = 10
  BallVelocityY = 10
  BonusCount = 0
  picBall.Tag = ""
  StartGame = False
End Sub

Private Sub DrawEnemis()
  Dim i As Integer
  Dim MissionComplete As Boolean
  MissionComplete = True
  For i = 0 To EnemyCount - 1
    If Enemy(i).Alive = True Then
      Select Case Enemy(i).Color
      Case Red: Call DrawBitmap(picEnemy, picEnemyMask.hDC, picGame.hDC, Enemy(i).EnemyRect.Left, Enemy(i).EnemyRect.Top)
                MissionComplete = False
      Case Blue: Call DrawBitmap(picEnemyBlue, picEnemyMask.hDC, picGame.hDC, Enemy(i).EnemyRect.Left, Enemy(i).EnemyRect.Top)
                 MissionComplete = False
      Case Yellow: Call DrawBitmap(picEnemyYellow, picEnemyMask.hDC, picGame.hDC, Enemy(i).EnemyRect.Left, Enemy(i).EnemyRect.Top)
                   MissionComplete = False
      Case Grey: Call DrawBitmap(picEnemyGrey, picEnemyMask.hDC, picGame.hDC, Enemy(i).EnemyRect.Left, Enemy(i).EnemyRect.Top)
      'Case Else: Call DrawBitmap(picEnemyGrey, picEnemyMask.hDC, picGame.hDC, Enemy(I).EnemyRect.Left, Enemy(I).EnemyRect.Top)
      End Select
    End If
  Next i
  If MissionComplete = True Then
    PlayerDirection.MoveLeft = False
    PlayerDirection.MoveRight = False
    Call SetStartPosition
    LevelNumber = LevelNumber + 1
    Call LoadLevels(LevelNumber)
  End If
End Sub

Private Sub BonusMove()
  Dim i As Integer
  Dim MoreBonuses As Boolean
  MoreBonuses = False
  For i = 1 To BonusCount
    If Bonuses(i).BonusNotExist = False Then
      Bonuses(i).BonusRect.Top = Bonuses(i).BonusRect.Top + BonusStep
      If Collision(Bonuses(i).BonusRect, PlayerRect) = True Then
        Bonuses(i).BonusNotExist = True
        Call CheckBonusType(Bonuses(i).BonusType)
      Else
        MoreBonuses = True
        Call DrawBitmap(picBonus, picBonusMask.hDC, picGame.hDC, Bonuses(i).BonusRect.Left, Bonuses(i).BonusRect.Top)
      End If
    End If
  Next i
  If MoreBonuses = False Then
    BonusCount = 0
  End If
End Sub

Private Sub CheckBonusType(ByVal BonusType As BonusEnum)
  
  Set picPlayer.Picture = LoadResPicture("Player", 0)
  Set picPlayerMask.Picture = LoadResPicture("PlayerMask", 0)
  PlayerRect.Width = picPlayer.ScaleWidth
  PlayerRect.Height = picPlayer.ScaleHeight
  
  If picBall.Tag = "fast" Then
    BallVelocityX = BallVelocityX / 2
    BallVelocityY = BallVelocityY / 2
  End If
  picBall.Tag = ""
  Select Case BonusType
  Case Large: Set picPlayer.Picture = LoadResPicture("PlayerLarge", 0)
              Set picPlayerMask.Picture = LoadResPicture("PlayerlargeMask", 0)
              PlayerRect.Width = picPlayer.ScaleWidth
              PlayerRect.Height = picPlayer.ScaleHeight
  Case Live: LiveNumber = LiveNumber + 1
             lblLives.Caption = LiveNumber
  Case FireBall: BallVelocityX = BallVelocityX * 2
                 BallVelocityY = BallVelocityY * 2
                 picBall.Tag = "fast"
  End Select
End Sub

Private Sub ChangeVellosity(EnemyRect As Rect)
  
  Dim OldX As Long
  Dim OldY As Long
  Dim J As Long
  
  OldY = BallRect.Top - BallVelocityY
  OldX = BallRect.Left - BallVelocityX
  If BallVelocityY > 0 Then
    If BallVelocityX > 0 Then
      If EnemyRect.Left - OldX <= EnemyRect.Top - OldY Then 'y>0 x>0
        BallVelocityY = -BallVelocityY
        BallRect.Top = EnemyRect.Top - BallRect.Height - 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left - BallRect.Width - 1
              BallRect.Top = EnemyRect.Top
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      Else
        BallVelocityX = -BallVelocityX
        BallRect.Left = EnemyRect.Left - BallRect.Width - 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left
              BallRect.Top = EnemyRect.Top - BallRect.Height - 1
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      End If
    Else
      If OldX - (EnemyRect.Left + EnemyRect.Width) <= EnemyRect.Top - OldY Then 'y>0 x<0
        BallVelocityY = -BallVelocityY
        BallRect.Top = EnemyRect.Top - BallRect.Height - 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
              BallRect.Top = EnemyRect.Top
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      Else
        BallVelocityX = -BallVelocityX
        BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left + EnemyRect.Width - BallRect.Width
              BallRect.Top = EnemyRect.Top - BallRect.Height - 1
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      End If
    End If
  Else
    If BallVelocityX > 0 Then
      If EnemyRect.Left - OldX <= OldY - (EnemyRect.Top + EnemyRect.Height) Then 'y<0 x>0
        BallVelocityY = -BallVelocityY
        BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left - BallRect.Width - 1
              BallRect.Top = EnemyRect.Top + EnemyRect.Height - BallRect.Height
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      Else
        BallVelocityX = -BallVelocityX
        BallRect.Left = EnemyRect.Left - BallRect.Width - 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left
              BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      End If
    Else
      If OldX - (EnemyRect.Left + EnemyRect.Width) <= OldY - (EnemyRect.Top + EnemyRect.Height) Then 'y<0 x<0
        BallVelocityY = -BallVelocityY
        BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
              BallRect.Top = EnemyRect.Top + EnemyRect.Height - BallRect.Height
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      Else
        BallVelocityX = -BallVelocityX
        BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
        For J = 1 To EnemyCount - 1
          If Enemy(J).Alive = True Then
            If Collision(BallRect, Enemy(J).EnemyRect) = True Then
              BallRect.Left = EnemyRect.Left + EnemyRect.Width - BallRect.Width
              BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
              BallVelocityX = -BallVelocityX
              BallVelocityY = -BallVelocityY
              Exit For
            End If
          End If
        Next J
      End If
    End If
  End If
End Sub

Private Sub LoadLevels(LevelNum As Integer)

  Dim strLineInput As String
  Dim lngCount As Long

  'LevelNum =

  Select Case LevelNum Mod 10
  Case 1
    strLineInput = "004900992000490126000049015300009100990000490180000133009900024500990102870099010133012600013301530001330180000091015300013302071000490207200203012601024501530102870180010245020711020302071102870153010203015301020300990102870207110357009902035701260203570153020357018002035702070204060153020448018002044802071204480126020406009902"
  Case 2
    strLineInput = "000000360000420036000084003600012600361001680036000210003600025200360002940036000336003600037800360004200036000462003600050400360000840090010084011701004200900100420117010168009001016801170102100090010210011701029400900102940117010336009001033601170104200090010420011701046200900104620117010000017100004201710000840171000126017100016801710002100171100252017100029401710003360171000378017100042001710004620171200504017100000000360000000063030126006303025200630303780063030504006303000001440301260144030252014403037801440305040144030000003600"
  Case 3
    strLineInput = "00000036210042006301008400360101260063010168003601021000630102520036010294006301033600360103780063010420003631046200630105040036010084009002008401440200000090020042011702016800902203360144320126011702021001170202520090020294011702033600900204200144020420009002037801170205040090020462011702000001440200420171110000019801012601710101680144020210017101025201440202940171010084019801037801710101680198010462017101050401440202520198010336019801042001980105040198010000003601"
  Case 4
    strLineInput = "000000360100000306010084003601016802520101680036010336025201025200360104200252010336003601050403060104200036010252025201050400360100840090320084014402000000900202100144000168009002033601440202940144000462014400025200900203780090000336009002042001440204200090020378014400050400900204620090200000014402000002520100000198010084025201016801440202520306010252014402033603060100840198010504025201016801980104200306010504014402025201980103360198110420019801050401980100840306010168030601012600900000420090000042014400021000900001260144000294009000"
  Case 5
    strLineInput = "02310009000189000900027300090001470036000231011721023100360202730036020315003600010500630000630090000021011700006301440001050171000147019800018902250002310225000273022500031501980003570171000357006300039900900003990144000441011700018900360201890063000147006302014700900001470117020147014400014701710201050090020105011712010501440200630117020189009002023100900202310063020273006300027300900201890144020189011702018901710001890198020231014402023101710202310198320273011702027301440202730171000273019802031501170203570117020399011702031500900003150063020357009002031501440003150171020357014412"
  Case 6
    strLineInput = "00000252000042025200008402520001260252000168025200021002523002520252000294025200033602520003780252200420025200046202520005040252000000027900004202793000840279000126027900016802790002100279000252027900029402790003360279000378027900042002790004620279000504027900004202250104620225010042019801046201980100840198010084022501042001980104200225010084017101012601710100840144010126014401012601170101680117010126009001016800900104200171010378017101037801440104200144010378011731037800900103360090010336011701016800630101680036010210003601021000633103360063010336003601029400630102940036010252003601025200630102520207020252012602025201533202520180220294015302021001530202100180020294018002"
  Case 7
    strLineInput = "0217002702025900270202170054020259005402021700810202590081020217010802025901080202170135020217016202021701890202170216320217024302021702700202590135020259016202025901890202590216020259024312025902700204270027020091005402042700540204270081020427010802042701353204270162020427018902042702160204270243020385027002038502430203850216020385018902038501620203850135020385010802038500810203850054120385002702042702700200910081020091002722009101080200910135020091016202009101892200910216020091024302009102700200490270020049024302004902163200490189020049016202004901350200490108020049008102004900540200490027020154013503032201350304900135030004013503"
  Case 8
    strLineInput = "0217002702025900270202170054020259005402021700810202590081020217010802025901080202170135020217016202021701890202170216320217024302021702700202590135020259016202025901890202590216020259024312025902700204270027020091005402042700540204270081020427010802042701353204270162020427018902042702160204270243020385027002038502430203850216020385018902038501620203850135020385010802038500810203850054120385002702042702700200910081020091002722009101080200910135020091016202009101892200910216020091024302009102700200490270020049024302004902163200490189020049016202004901350200490108020049008102004900540200490027020154013503032201350304900135030004013503"
  Case 9
    strLineInput = "011900900002310072000252016200018901260001680180010189023401033602430103290153010329008101"
  Case 0
    strLineInput = "0014003602001400630200140090020014011732001401440200560036020056009032009800630200980117020133014402001401710200560171020098017102009800360201890171020189014402018901170201890090020189006302023100360202730063020273009002027301170202730144020273017102023101171203290036020329006302032900901203290117020329014402032901710203710171020427003602042701710204270144220427011702042700900204270063020469017102"
  End Select

  EnemyCount = Len(strLineInput) / 10
  ReDim Enemy(EnemyCount - 1)
  'I = 0
  For lngCount = 1 To EnemyCount
    With Enemy(lngCount - 1)
      .EnemyRect.Left = Val(Mid(strLineInput, 1, 4))
      .EnemyRect.Top = Val(Mid(strLineInput, 5, 4))
      .EnemyRect.Width = ENEMY_WIDTH
      .EnemyRect.Height = ENEMY_HEIGHT
      .Alive = True
      .Bonus = Mid(strLineInput, 9, 1)
      .Color = Mid(strLineInput, 10, 1)
    End With
    strLineInput = Mid(strLineInput, 11)
  Next

  lblLevel.Caption = LevelNum
  txtLevel.Text = CStr(LevelNum)
  Exit Sub
LoadLevelsError:
  COAMsgBox Err.Description, vbCritical, "ASRkanoid"
  Unload Me
End Sub


