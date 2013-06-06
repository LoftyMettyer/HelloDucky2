VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmToolbox 
   Caption         =   "Toolbox"
   ClientHeight    =   8460
   ClientLeft      =   375
   ClientTop       =   2205
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5035
   Icon            =   "frmToolbox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   3390
   Visible         =   0   'False
   Begin VB.PictureBox PicDragIcon_ColourPicker 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   36
      Top             =   8115
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicDragIcon_Navigation 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":0596
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   7800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_ColumnDrag 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":0920
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   7500
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_PageTab 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":0C2A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Frame 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":11B4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Date 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":12FE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Textbox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":1888
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   300
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_CheckBox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":1E12
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   5700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Button 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":239C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Image 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":2726
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Numeric 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":2CB0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   3300
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Grid 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":323A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   2100
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Label 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":37C4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Line 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":3D4E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Column 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":40D8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_ComboBox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":4662
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   6300
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_WorkingPattern 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":4BEC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_WebForm 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":4F76
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Link 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":5878
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   4500
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Lookup 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":5E02
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   2700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Photo 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":638C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_OLE 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":6916
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Radio 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":6EA0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   1500
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Spinner 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":742A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   6900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_WorkFlow 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":79B4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Properties 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":7F3E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Table 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":84C8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   5100
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_ToolBox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmToolbox.frx":8A52
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   3900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   120
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   1890
      Width           =   2940
   End
   Begin ComctlLib.TreeView trvStandardControls 
      DragIcon        =   "frmToolbox.frx":8FDC
      Height          =   1890
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3334
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraColumns 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   265
      Left            =   0
      TabIndex        =   5
      Top             =   2025
      Width           =   3060
      Begin VB.Label lblTitleBar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "Columns"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   45
         Width           =   1245
      End
      Begin VB.Image imgTitleBar 
         Height          =   255
         Left            =   45
         Picture         =   "frmToolbox.frx":92E6
         Stretch         =   -1  'True
         Top             =   10
         Width           =   255
      End
   End
   Begin VB.Frame fraTable 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   45
      TabIndex        =   2
      Top             =   2205
      Width           =   2940
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   45
         TabIndex        =   4
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lblTableName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTableName"
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   150
         Width           =   2160
      End
   End
   Begin ComctlLib.TreeView trvColumns 
      DragIcon        =   "frmToolbox.frx":96A2
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   2715
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3625
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmToolbox.frx":9C2C
      Height          =   2670
      Left            =   60
      TabIndex        =   32
      Top             =   5430
      Width           =   2805
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   1230
      Top             =   4785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   26
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":9D1D
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":A26F
            Key             =   "IMG_BUTTON"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":A7C1
            Key             =   "IMG_WORKINGPATTERN"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":AD13
            Key             =   "IMG_WEBFORM"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":B065
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":B3B7
            Key             =   "IMG_GRID"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":B909
            Key             =   "IMG_COLUMN"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":BE5B
            Key             =   "IMG_COMBOBOX"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":C3AD
            Key             =   "IMG_DATE"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":C8FF
            Key             =   "IMG_LINE"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":CE51
            Key             =   "IMG_IMAGE"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":D3A3
            Key             =   "IMG_FRAME"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":D8F5
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":DE47
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":E399
            Key             =   "IMG_CHECKBOX"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":E8EB
            Key             =   "IMG_LOOKUP"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":EE3D
            Key             =   "IMG_NUMERIC"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":F38F
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":F8E1
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":FE33
            Key             =   "IMG_PROPERTIES"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":10385
            Key             =   "IMG_RADIO"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":108D7
            Key             =   "IMG_SPINNER"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":10E29
            Key             =   "IMG_TEXTBOX"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":1137B
            Key             =   "IMG_TOOLBOX"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":118CD
            Key             =   "IMG_WORKFLOW"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":11E1F
            Key             =   "IMG_COLOURPICKER"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   660
      Top             =   4785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":12171
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":126C3
            Key             =   "IMG_NAVIGATION"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":12A15
            Key             =   "IMG_IMAGE"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":12F67
            Key             =   "IMG_PAGETAB"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":134B9
            Key             =   "IMG_FRAME"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":13A0B
            Key             =   "IMG_TOOLS"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbox.frx":13F5D
            Key             =   "IMG_LINE"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local variables to hold property values
Private gFrmScreen As frmScrDesigner2
Private giScreenCount As Integer

'Declare local variables
Dim gSngSplitStartY As Single
Dim gfSplitMoving As Boolean

Private Const MIN_FORM_WIDTH = 2000
Private Const MIN_FORM_HEIGHT = 2600

Public Sub EditMenu(ByVal psMenuOption As String)
  
  ' Pass any menu events onto the current screens
  ' 'frmScrDesigner2' form to handle.
  CurrentScreen.EditMenu psMenuOption

End Sub


Public Property Get CurrentScreen() As frmScrDesigner2
  
  ' Set the current screen property
  Set CurrentScreen = gFrmScreen

End Property

Public Property Set CurrentScreen(pScrForm As frmScrDesigner2)

  ' Set the current screen property
  Set gFrmScreen = pScrForm
  
  ' Populate the 'columns' treeview with the columns of the databases
  ' associated with the current screen.
  RefreshControls

End Property

Public Sub RefreshControls()
  RefreshStandardControlsTreeView
  RefreshColumnsTreeView
  
End Sub

Private Sub Form_Activate()

  'Refresh menus
  frmSysMgr.RefreshMenu
  
  ' Change the colours of some of the controls to
  ' indicate that this form is now active.
  fraColumns.BackColor = vbActiveTitleBar
  lblTitleBar.BackColor = vbActiveTitleBar
  lblTitleBar.ForeColor = vbTitleBarText

End Sub

Private Sub Form_Deactivate()

  ' Change the colours of some of the controls to indicate
  ' that this screen is now inactive.
  fraColumns.BackColor = vbInactiveTitleBar
  lblTitleBar.BackColor = vbInactiveTitleBar
  lblTitleBar.ForeColor = vbInactiveCaptionText

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'NHRD Checked in by NHRD: Code inserted by JPD28112006 Fault 11228
  Dim bHandled As Boolean
  
  Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  End Select

  bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If
'Pass any keystrokes onto the toolbar in the frmSysMgr form.
'frmSysMgr.ActiveBar1.OnKeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()
  Dim iCYFrame As Integer
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  ' Get then dimension of windows borders
  iCYFrame = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
  
  ' Position and size form
  Me.Move 0, 0, 2800, Forms(0).ScaleHeight
  
  ' Position and size form controls
  fraColumns.Left = 0
  imgTitleBar.Move 20, 10, 16 * Screen.TwipsPerPixelX, 16 * Screen.TwipsPerPixelY
  
  'NPG20091029 Fault HRPRO-531
  imgTitleBar.Visible = False
  lblTitleBar.Move 20, 30, 820, 195
  
  fraTable.Left = 0
  trvColumns.Left = 0
  fraSplit.Left = 0
  fraSplit.Height = iCYFrame
  
  Form_Resize
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, piUnloadMode As Integer)
  
  ' Only unload the form if really required.
  If piUnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
  ElseIf piUnloadMode <> vbFormCode Then
    If Me.ScreenCount > 0 Then
      Cancel = True
    End If
  End If

End Sub


Private Sub Form_Resize()

  ' If this form is not already minimised then ensure that all controls on this
  ' form are resized accordingly.
  If Me.WindowState <> vbMinimized Then
    
    fraSplit.Width = Me.ScaleWidth
     
    fraColumns.Width = Me.ScaleWidth
    fraTable.Move 0, fraColumns.Top + fraColumns.Height, Me.ScaleWidth, 365
    lblTable.Move 45, 70
    lblTableName.Move 700, 40, Me.ScaleWidth - 630, 285
    trvColumns.Width = Me.ScaleWidth
    trvStandardControls.Width = Me.ScaleWidth
  
    SplitMove
  End If

  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME

End Sub


Private Sub SplitMove()
  ' Resize the standard controls treeview, and the columns treeview
  ' as the user has dragged the split between the two treeviews.
  
  ' Limit the minimum size of the standard controls treeview.
  If fraSplit.Top < 810 Then
    fraSplit.Top = 810
  Else
    ' Update the display of the columns treeview if it is now too small to
    ' see properly.
    If fraSplit.Top > Me.ScaleHeight - (fraColumns.Height + 500) Then
      fraTable.Visible = False
      trvColumns.Visible = False
      fraSplit.Top = Me.ScaleHeight - (fraColumns.Height + 30)
      fraColumns.BackColor = vbInactiveTitleBar
      lblTitleBar.BackColor = vbInactiveTitleBar
      lblTitleBar.ForeColor = vbInactiveCaptionText
    Else
      fraColumns.BackColor = vbActiveTitleBar
      lblTitleBar.BackColor = vbActiveTitleBar
      lblTitleBar.ForeColor = vbTitleBarText
      fraTable.Visible = True
      trvColumns.Visible = True
    End If
  End If

  ' Resize the two treeviews.
  trvStandardControls.Height = fraSplit.Top - trvStandardControls.Top
  fraColumns.Width = Me.ScaleWidth
  fraColumns.Top = fraSplit.Top + fraSplit.Height
  If trvColumns.Visible Then
    fraTable.Top = fraColumns.Top + fraColumns.Height
    trvColumns.Top = fraTable.Top + fraTable.Height
    trvColumns.Height = Me.ScaleHeight - trvColumns.Top
  End If

  gfSplitMoving = False

End Sub
Private Sub RefreshColumnsTreeView()
  ' Populate the treeview of database column controls.
  Dim objTable As HRProSystemMgr.Table
  Dim objNode As ComctlLib.Node
  Dim rsColumns As DAO.Recordset
  Dim fPopulated As Boolean
  Dim sSQL As String
  Dim sIconKey As String
  Static lngTableID As Long
  Static datUpdated As Date
  
  fPopulated = False
  
  ' If there is a current screen ...
  If Not gFrmScreen Is Nothing Then
    ' If there is a primary table associated with the current screen ...
    If gFrmScreen.TableID > 0 Then
      ' Get primary table details
      Set objTable = New HRProSystemMgr.Table
      objTable.TableID = gFrmScreen.TableID
      
      If objTable.ReadTable Then
        lblTableName.Caption = objTable.TableName
        lngTableID = objTable.TableID
        datUpdated = objTable.LastUpdated
        
        Set objTable = Nothing
        Set objTable = New HRProSystemMgr.Table
        
        ' Remove any existing nodes from the treeview.
        trvColumns.Nodes.Clear
          
        ' Get column details for the primary table.
        ' NPG20091013 Fault HRPRO-478 (added dataType to the select)
        sSQL = "SELECT tableID, columnID, columnName, columnType, deleted, controlType, dataType " & _
          " FROM tmpColumns" & _
          " WHERE tableID = " & Trim(Str(lngTableID)) & _
          " AND columnType <>" & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
          " AND deleted = FALSE"

'        If gFrmScreen.IsSSIntranetScreen Then
'          sSQL = sSQL & _
'            " AND controlType <>" & CStr(giCTRL_OLE) & _
'            " AND controlType <>" & CStr(giCTRL_PHOTO)
'        End If

        sSQL = sSQL & _
          " ORDER BY columnName"

        Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
        
        ' Loop though columns and populate treeview
        Do While Not rsColumns.EOF
          ' Add a node to the treeview for any parent tables.
          If rsColumns!TableID <> objTable.TableID Then
            objTable.TableID = rsColumns!TableID
            
            If objTable.ReadTable Then
              Set objNode = trvColumns.Nodes.Add(, tvwChild, _
                "T" & objTable.TableID, objTable.TableName, 1, 1)
              objNode.Sorted = True
              objNode.Expanded = True
              Set objNode = Nothing
            End If
          End If
              
          ' Get the correct icon for the current column.
          ' NPG20091013 Fault HRPRO-478 added the datatype parameter for numeric columns
          sIconKey = GetColumnIcon(rsColumns!ControlType, rsColumns!DataType)

          'Add column to TreeView
          Set objNode = trvColumns.Nodes.Add("T" & rsColumns!TableID, _
            tvwChild, "C" & rsColumns!ColumnID & "T" & rsColumns!TableID, _
            rsColumns!ColumnName, sIconKey, sIconKey)
          objNode.Tag = rsColumns!ColumnID
          Set objNode = Nothing
        
          rsColumns.MoveNext
        Loop
        
        rsColumns.Close
        Set rsColumns = Nothing
      
        ' Get column details for the primary table's parent tables.
        sSQL = "SELECT tableID, columnID, columnName, columnType, deleted, controlType, dataType " & _
          " FROM tmpColumns, tmpRelations" & _
          " WHERE parentID = tableID" & _
          " AND childID = " & Trim(Str(lngTableID)) & _
          " AND columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
          " AND columnType <> " & Trim(Str(giCOLUMNTYPE_LINK)) & _
          " AND deleted = FALSE"

'        If gFrmScreen.IsSSIntranetScreen Then
'          sSQL = sSQL & _
'            " AND controlType <>" & CStr(giCTRL_OLE) & _
'            " AND controlType <>" & CStr(giCTRL_PHOTO)
'        End If

        sSQL = sSQL & _
          " ORDER BY tableID, columnName"
        Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
        
        ' Loop though columns and populate treeview
        Do While Not rsColumns.EOF
          ' Add a node to the treeview for any parent tables.
          If rsColumns!TableID <> objTable.TableID Then
            objTable.TableID = rsColumns!TableID
            
            If objTable.ReadTable Then
              Set objNode = trvColumns.Nodes.Add(, tvwChild, _
                "T" & objTable.TableID, objTable.TableName, 1, 1)
              objNode.Sorted = True
              objNode.Expanded = True
              Set objNode = Nothing
            End If
          End If
              
          ' Get the correct icon for the current column.
          ' NPG20091013 Fault HRPRO-478 added the datatype parameter for numeric columns
          sIconKey = GetColumnIcon(rsColumns!ControlType, rsColumns!DataType)

          'Add column to TreeView
          Set objNode = trvColumns.Nodes.Add("T" & rsColumns!TableID, _
            tvwChild, "C" & rsColumns!ColumnID & "T" & rsColumns!TableID, _
            rsColumns!ColumnName, sIconKey, sIconKey)
          objNode.Tag = rsColumns!ColumnID
          Set objNode = Nothing
            
          rsColumns.MoveNext
        Loop
        
        rsColumns.Close
        Set rsColumns = Nothing
          
        fPopulated = True
      End If
      
      ' Disassociate object variables.
      Set objTable = Nothing
    End If
  End If
  
  If Not fPopulated Then
    trvColumns.Nodes.Clear
  End If
End Sub

Private Sub RefreshStandardControlsTreeView()
  '
  ' Populate the treeview of standard controls.
  '
  Dim objNode As ComctlLib.Node
  
  ' Clear the treeview
  trvStandardControls.Nodes.Clear
  
  Set objNode = trvStandardControls.Nodes.Add(, , "STDROOT", "Standard Controls", "IMG_TOOLS", "IMG_TOOLS")
  objNode.Expanded = True
    
  ' Add the standard controls to the tree view.
  ' Add the Label control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "LABELCTRL", "Label", "IMG_LABEL", "IMG_LABEL")
  objNode.Expanded = True
  
  ' Add the Frame control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "FRAMECTRL", "Frame", "IMG_FRAME", "IMG_FRAME")
  objNode.Expanded = True
  
  If Not gFrmScreen.IsSSIntranetScreen Then
    ' Add the Image control node.
    Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "IMAGECTRL", "Image", "IMG_IMAGE", "IMG_IMAGE")
    objNode.Expanded = True
  End If
  
  ' Add the PageTab control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "PAGETABCTRL", "Page Tab", "IMG_PAGETAB", "IMG_PAGETAB")
  objNode.Expanded = True
  
  ' Add the Line control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "LINECTRL", "Line", "IMG_LINE", "IMG_LINE")
  objNode.Expanded = True
   
  ' Add the Fire Button control
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "NAVIGATIONCTRL", "Navigation", "IMG_NAVIGATION", "IMG_NAVIGATION")
  objNode.Expanded = True
  
  ' Disassociate the objNode variable.
  Set objNode = Nothing
  
End Sub



Private Sub Form_Unload(Cancel As Integer)
  
  ' Disassociate global variables and the form itself.
  Set gFrmScreen = Nothing
  Set frmToolbox = Nothing
  
End Sub

Private Sub fraColumns_Click()

  ' Set focus onto the correct treeview.
  If trvColumns.Visible Then
    trvColumns.SetFocus
  Else
    trvStandardControls.SetFocus
  End If

End Sub

Private Sub fraColumns_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  fraColumns_Click

End Sub

Private Sub fraSplit_MouseDown(piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)
  
  ' Record the original split start Y co-ordinate.
  gSngSplitStartY = pSngY
  gfSplitMoving = True
  
End Sub


Private Sub fraSplit_MouseMove(piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)
  
  ' Move the split between the standard controls and the columns tree views.
  If gfSplitMoving Then
    fraSplit.Top = fraSplit.Top + (pSngY - gSngSplitStartY)
  End If

End Sub


Private Sub fraSplit_MouseUp(piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)
  
  ' If we a moving the split the resize the neighbouring tree views.
  If gfSplitMoving Then
    SplitMove
  End If

End Sub


Private Sub trvColumns_Click()

  ' Mark that there is a currently selected item.
  If Not trvColumns.SelectedItem Is Nothing Then
    trvColumns.SelectedItem.Selected = True
  End If
  
End Sub


Private Sub trvColumns_MouseDown(piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)
  Dim ThisNode As ComctlLib.Node
  
  ' If the left mouse button is pressed ...
  If piButton = vbLeftButton Then
  
    'Get the node at the mouse position
    Set ThisNode = trvColumns.HitTest(pSngX, pSngY)
    
    ' If we have selected a valid node ...
    If Not ThisNode Is Nothing Then
    
      ' Only allow columns to be selected (ie. not table name nodes).
      If Left(ThisNode.key, 1) = "C" Then
      
        ' Ensure this node is not the selected node.
        If Not ThisNode Is trvColumns.SelectedItem Then
          Set trvColumns.SelectedItem = ThisNode
        End If
        
        'NHRD20092006 Fault 10990
        'Added all of the possible Icon so far to futureproof it a bit
        Select Case ThisNode.SelectedImage
            Case "IMG_FRAME"
                trvColumns.DragIcon = picDragIcon_Frame.Picture
            Case "IMG_DATE"
                trvColumns.DragIcon = picDragIcon_Date.Picture
            Case "IMG_BUTTON"
                trvColumns.DragIcon = picDragIcon_Button.Picture
            Case "IMG_IMAGE"
                trvColumns.DragIcon = picDragIcon_Image.Picture
            Case "IMG_TEXTBOX"
                trvColumns.DragIcon = picDragIcon_Textbox.Picture
            Case "IMG_CHECKBOX"
                trvColumns.DragIcon = picDragIcon_CheckBox.Picture
            Case "IMG_NUMERIC"
                trvColumns.DragIcon = picDragIcon_Numeric.Picture
            Case "IMG_GRID"
                trvColumns.DragIcon = picDragIcon_Grid.Picture
            Case "IMG_LABEL"
                trvColumns.DragIcon = picDragIcon_Label.Picture
            Case "IMG_LINE"
                trvColumns.DragIcon = picDragIcon_Line.Picture
            Case "IMG_WORKINGPATTERN"
                trvColumns.DragIcon = picDragIcon_WorkingPattern.Picture
            Case "IMG_WEBFORM"
                trvColumns.DragIcon = picDragIcon_WebForm.Picture
            Case "IMG_COLUMN"
                trvColumns.DragIcon = picDragIcon_Column.Picture
            Case "IMG_COMBOBOX"
                trvColumns.DragIcon = picDragIcon_ComboBox.Picture
            Case "IMG_LINK"
                trvColumns.DragIcon = picDragIcon_Link.Picture
            Case "IMG_LOOKUP"
                trvColumns.DragIcon = picDragIcon_Lookup.Picture
            Case "IMG_PHOTO"
                trvColumns.DragIcon = picDragIcon_Photo.Picture
            Case "IMG_OLE"
                trvColumns.DragIcon = picDragIcon_OLE.Picture
            Case "IMG_WORKFLOW"
                trvColumns.DragIcon = picDragIcon_WorkFlow.Picture
            Case "IMG_PROPERTIES"
                trvColumns.DragIcon = picDragIcon_Properties.Picture
            Case "IMG_RADIO"
                trvColumns.DragIcon = picDragIcon_Radio.Picture
            Case "IMG_SPINNER"
                trvColumns.DragIcon = picDragIcon_Spinner.Picture
            Case "IMG_TABLE"
                trvColumns.DragIcon = picDragIcon_Table.Picture
            Case "IMG_TOOLBOX"
                trvColumns.DragIcon = picDragIcon_ToolBox.Picture
            Case "IMG_PAGETAB"
                trvColumns.DragIcon = picDragIcon_PageTab.Picture
            Case "IMG_FILEUPLOAD"
                'trvColumns.DragIcon = picDragIcon_FileUpload.Picture
            Case "IMG_NAVIGATION"
                trvColumns.DragIcon = PicDragIcon_Navigation
            Case "IMG_COLOURPICKER"
                trvColumns.DragIcon = PicDragIcon_ColourPicker
            Case Else
                trvColumns.DragIcon = picDragIcon_ColumnDrag.Picture
        End Select
        'Begin drag
        trvColumns.Drag vbBeginDrag
      End If
    End If
    ' Disassociate object variables.
    Set ThisNode = Nothing
  End If
End Sub


Private Sub trvColumns_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Signal that the drag operation has ended.
  trvColumns.Drag vbEndDrag

End Sub



Public Property Get ScreenCount() As Integer
  
  ' Return the current screen count
  ScreenCount = giScreenCount

End Property

Public Property Let ScreenCount(piCount As Integer)

  ' Set the current screen count
  giScreenCount = piCount
  
  ' If there are no more screens then unload the toolbox.
  If giScreenCount < 1 Then
    UnLoad Me
  End If
  
End Property

Private Sub trvStandardControls_Click()

  ' Mark that there is a currently selected item.
  If Not trvStandardControls.SelectedItem Is Nothing Then
    trvStandardControls.SelectedItem.Selected = True
  End If

End Sub


Private Sub trvStandardControls_MouseDown(piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)
  Dim ThisNode As ComctlLib.Node
    
  ' If the left mouse button is pressed ...
  If piButton = vbLeftButton Then
    
    'Get the node at the mouse position
    Set ThisNode = trvStandardControls.HitTest(pSngX, pSngY)
    
    ' If we have selected a valid node ...
    If Not ThisNode Is Nothing Then
    
      ' Ensure this node is not the selected node.
      If Not ThisNode Is trvStandardControls.SelectedItem Then
        Set trvStandardControls.SelectedItem = ThisNode
      End If
      
      'NHRD20092006 Fault 10990
      'Added all of the possible Icon so far to futureproof it a bit
      Select Case ThisNode.SelectedImage
          Case "IMG_FRAME"
              trvStandardControls.DragIcon = picDragIcon_Frame.Picture
          Case "IMG_DATE"
              trvStandardControls.DragIcon = picDragIcon_Date.Picture
          Case "IMG_BUTTON"
              trvStandardControls.DragIcon = picDragIcon_Button.Picture
          Case "IMG_IMAGE"
              trvStandardControls.DragIcon = picDragIcon_Image.Picture
          Case "IMG_TEXTBOX"
              trvStandardControls.DragIcon = picDragIcon_Textbox.Picture
          Case "IMG_CHECKBOX"
              trvStandardControls.DragIcon = picDragIcon_CheckBox.Picture
          Case "IMG_NUMERIC"
              trvStandardControls.DragIcon = picDragIcon_Numeric.Picture
          Case "IMG_GRID"
              trvStandardControls.DragIcon = picDragIcon_Grid.Picture
          Case "IMG_LABEL"
              trvStandardControls.DragIcon = picDragIcon_Label.Picture
          Case "IMG_LINE"
              trvStandardControls.DragIcon = picDragIcon_Line.Picture
          Case "IMG_WORKINGPATTERN"
              trvStandardControls.DragIcon = picDragIcon_WorkingPattern.Picture
          Case "IMG_WEBFORM"
              trvStandardControls.DragIcon = picDragIcon_WebForm.Picture
          Case "IMG_COLUMN"
              trvStandardControls.DragIcon = picDragIcon_Column.Picture
          Case "IMG_COMBOBOX"
              trvStandardControls.DragIcon = picDragIcon_ComboBox.Picture
          Case "IMG_LINK"
              trvStandardControls.DragIcon = picDragIcon_Link.Picture
          Case "IMG_LOOKUP"
              trvStandardControls.DragIcon = picDragIcon_Lookup.Picture
          Case "IMG_PHOTO"
              trvStandardControls.DragIcon = picDragIcon_Photo.Picture
          Case "IMG_OLE"
              trvStandardControls.DragIcon = picDragIcon_OLE.Picture
          Case "IMG_WORKFLOW"
              trvStandardControls.DragIcon = picDragIcon_WorkFlow.Picture
          Case "IMG_PROPERTIES"
              trvStandardControls.DragIcon = picDragIcon_Properties.Picture
          Case "IMG_RADIO"
              trvStandardControls.DragIcon = picDragIcon_Radio.Picture
          Case "IMG_SPINNER"
              trvStandardControls.DragIcon = picDragIcon_Spinner.Picture
          Case "IMG_TABLE"
              trvStandardControls.DragIcon = picDragIcon_Table.Picture
          Case "IMG_TOOLBOX"
              trvStandardControls.DragIcon = picDragIcon_ToolBox.Picture
          Case "IMG_PAGETAB"
              trvStandardControls.DragIcon = picDragIcon_PageTab.Picture
          Case "IMG_FILEUPLOAD"
              ' trvStandardControls.DragIcon = picDragIcon_FileUpload.Picture
          Case "IMG_NAVIGATION"
              trvStandardControls.DragIcon = PicDragIcon_Navigation
          Case "IMG_COLOURPICKER"
              trvStandardControls.DragIcon = PicDragIcon_ColourPicker
          Case Else
              trvStandardControls.DragIcon = picDragIcon_ColumnDrag.Picture
      End Select
      'Begin drag
      trvStandardControls.Drag vbBeginDrag
    End If
    Set ThisNode = Nothing
  End If
End Sub


Private Sub trvStandardControls_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Cancel the control drag operation.
  trvColumns.Drag vbEndDrag

End Sub



Private Function GetColumnIcon(piColumnType As Long, piDataType As Integer) As String
  'Dim iLoop As Integer
  
  ' Depending on the control type of the current column, determine the
  ' associated icon key, as defined in the imageList2 control.
  ' NPG20091013 Fault HRPRO-478 added the datatype parameter for numeric columns
  Select Case piColumnType
    Case giCTRL_TEXTBOX
      If piDataType = sqlNumeric Or piDataType = sqlInteger Then  '11
        GetColumnIcon = "IMG_NUMERIC"
      Else
        GetColumnIcon = "IMG_TEXTBOX"
      End If
    Case giCTRL_CHECKBOX
      GetColumnIcon = "IMG_CHECKBOX"
    Case giCTRL_OPTIONGROUP
      GetColumnIcon = "IMG_RADIO"
    Case giCTRL_OLE
      GetColumnIcon = "IMG_OLE"
    Case giCTRL_PHOTO
      GetColumnIcon = "IMG_PHOTO"
    Case giCTRL_COMBOBOX
      GetColumnIcon = "IMG_COMBOBOX"
    Case giCTRL_SPINNER
      GetColumnIcon = "IMG_SPINNER"
    Case giCTRL_LINK
      GetColumnIcon = "IMG_LINK"
    Case giCTRL_WORKINGPATTERN
      GetColumnIcon = "IMG_WORKINGPATTERN"
    Case giCTRL_NAVIGATION
      'NHRD We are using the textbox icon for Navigation control
      GetColumnIcon = "IMG_TEXTBOX"
      
    Case giCTRL_COLOURPICKER
      GetColumnIcon = "IMG_COLOURPICKER"
    
    Case Else
      GetColumnIcon = "IMG_COLUMN"
    End Select
  
End Function
