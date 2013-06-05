VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "CODEJO~2.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00FCF7F9&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmProgress.frx":058A
   ScaleHeight     =   3810
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraProgress1 
      BackColor       =   &H00FCF7F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1010
      Left            =   180
      TabIndex        =   8
      Top             =   1110
      Width           =   5685
      Begin VB.PictureBox picProgress1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   5625
         TabIndex        =   9
         Top             =   550
         Width           =   5685
         Begin MSComctlLib.ProgressBar pbrProgress1 
            Height          =   165
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   5620
            _ExtentX        =   9922
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.Label lblBar1Percent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00663333&
         Height          =   195
         Left            =   5220
         TabIndex        =   11
         Top             =   105
         Width           =   465
      End
      Begin VB.Label lblProgress1 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgress1"
         ForeColor       =   &H00663333&
         Height          =   420
         Left            =   0
         TabIndex        =   12
         Top             =   105
         Width           =   5235
      End
   End
   Begin VB.Frame fraProgress2 
      BackColor       =   &H00FCF7F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1010
      Left            =   165
      TabIndex        =   3
      Top             =   2135
      Width           =   5685
      Begin VB.PictureBox picProgress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   5625
         TabIndex        =   4
         Top             =   550
         Width           =   5685
         Begin MSComctlLib.ProgressBar pbrProgress2 
            Height          =   165
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   5620
            _ExtentX        =   9922
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.Label lblBar2Percent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00663333&
         Height          =   195
         Left            =   5220
         TabIndex        =   6
         Top             =   105
         Width           =   465
      End
      Begin VB.Label lblProgress2 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgress2"
         ForeColor       =   &H00663333&
         Height          =   420
         Left            =   0
         TabIndex        =   7
         Top             =   105
         Width           =   5235
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4650
      TabIndex        =   2
      Top             =   3205
      Width           =   1200
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   900
      Left            =   3780
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1588
      _Version        =   327681
      AutoPlay        =   -1  'True
      BackStyle       =   1
      FullWidth       =   150
      FullHeight      =   60
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      ForeColor       =   &H00663333&
      Height          =   195
      Left            =   165
      TabIndex        =   13
      Top             =   3205
      Width           =   420
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5520
      Top             =   315
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblMainCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00663333&
      Height          =   270
      Left            =   210
      TabIndex        =   1
      Top             =   330
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   -2760
      Picture         =   "frmProgress.frx":2C77
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_DLGMODALFRAME As Long = &H1

Private Const hWnd_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Const SC_CLOSE As Long = &HF060&
Private Const xSC_CLOSE As Long = -10&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Const WM_USER = &H400&
Private Const ACM_OPEN = WM_USER + 100&

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Event Cancelled()

Private mstrStyleResource As String
Private mstrStyleIni As String

Private Sub cmdCancel_Click()
  RaiseEvent Cancelled
End Sub


Public Function FormSetTopMost() As Boolean
  
  On Local Error GoTo LocalError

  FormSetTopMost = False
  
  Call SetWindowPos(Me.hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
  Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

  FormSetTopMost = True

Exit Function


LocalError:
  FormSetTopMost = False

End Function


Public Function FormEnableCloseButton(blnEnable As Boolean) As Boolean
    
  Dim hMenu As Long
  Dim MII As MENUITEMINFO
  Dim lngMenuID As Long
  
  
  On Local Error GoTo LocalError
    
  FormEnableCloseButton = False
    
    
  If IsWindow(Me.hWnd) = 0 Then
    Exit Function
  End If
    
  hMenu = GetSystemMenu(Me.hWnd, 0)
    
  MII.cbSize = Len(MII)
  MII.dwTypeData = String$(80, 0)
  MII.cch = Len(MII.dwTypeData)
  MII.fMask = MIIM_STATE
    
  If blnEnable Then
    MII.wID = xSC_CLOSE
  Else
    MII.wID = SC_CLOSE
  End If
    
  If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
    Exit Function
  End If
    
    
  lngMenuID = MII.wID
    
  If blnEnable Then
    MII.wID = SC_CLOSE
  Else
    MII.wID = xSC_CLOSE
  End If
    
  MII.fMask = MIIM_ID
  If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then
    Exit Function
  End If
    
  If blnEnable Then
    MII.fState = MII.fState And Not MFS_GRAYED
  Else
    MII.fState = MII.fState Or MFS_GRAYED
  End If
    
  MII.fMask = MIIM_STATE
  If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
    Exit Function
  End If
    
  SendMessage Me.hWnd, WM_NCACTIVATE, True, 0
    
  FormEnableCloseButton = True

Exit Function


LocalError:
  FormEnableCloseButton = False

End Function

Public Property Get StyleResource() As String
  StyleResource = mstrStyleResource
End Property

Public Property Let StyleResource(ByVal sNewValue As String)
  mstrStyleResource = sNewValue
End Property

Public Property Get StyleIni() As String
  StyleIni = mstrStyleIni
End Property

Public Property Let StyleIni(ByVal sNewValue As String)
  mstrStyleIni = sNewValue
End Property

Private Sub Form_Load()

  LoadSkin Me, Me.SkinFramework1, mstrStyleResource, mstrStyleIni

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    If cmdCancel.Visible Then
        cmdCancel_Click
        Cancel = True
    End If
  End If
End Sub


Public Sub SetCaption(strCaption As String)
  Me.Caption = strCaption
  Me.Icon = Nothing
  'SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME
  SetWindowLong Me.hWnd, GWL_EXSTYLE, (WS_EX_WINDOWEDGE Or WS_EX_DLGMODALFRAME) And Not WS_EX_APPWINDOW
End Sub


Public Sub SetAVI(lngAVI As Long)

  On Error GoTo LocalErr
  Animation1.Visible = False

  If lngAVI > 0 Then
    SendMessage Animation1.hWnd, ACM_OPEN, ByVal App.hInstance, ByVal lngAVI
  End If
  Animation1.Visible = (lngAVI > 0)
    
Exit Sub

LocalErr:

End Sub
