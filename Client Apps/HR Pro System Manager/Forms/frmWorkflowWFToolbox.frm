VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmWorkflowWFToolbox 
   Caption         =   "Toolbox"
   ClientHeight    =   9180
   ClientLeft      =   375
   ClientTop       =   2205
   ClientWidth     =   3360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1074
   Icon            =   "frmWorkflowWFToolbox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   3360
   Visible         =   0   'False
   Begin VB.PictureBox picDragIcon_FileDownload 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   40
      Top             =   8220
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_FileUpload 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":0596
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   39
      Top             =   7920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_ColumnDrag 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":0B20
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   38
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_PageTab 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":0E2A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   37
      Top             =   7320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_ToolBox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":13B4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   3420
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Table 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":193E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   4620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Properties 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":1EC8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   3120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_WorkFlow 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":2452
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Spinner 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":29DC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   6420
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Radio 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":2F66
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_OLE 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":34F0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   6120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Photo 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":3A7A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Lookup 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":4004
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   27
      Top             =   2820
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Link 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":458E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   4020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_WebForm 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":4918
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_WorkingPattern 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":521A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_ComboBox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":55A4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   5820
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Column 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":5B2E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   1020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Line 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":60B8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Label 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":6442
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Grid 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":69CC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   2220
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Numeric 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":6F56
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   7020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Image 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":74E0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Button 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":7A6A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_CheckBox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":7DF4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   5220
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Textbox 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":837E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Date 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":8908
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picDragIcon_Frame 
      Height          =   300
      Left            =   3000
      Picture         =   "frmWorkflowWFToolbox.frx":8E92
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picToolbox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   120
      Picture         =   "frmWorkflowWFToolbox.frx":8FDC
      ScaleHeight     =   7695
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.Frame fraToolboxSplit 
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   11
         Top             =   5880
         Width           =   2730
      End
      Begin VB.Frame fraToolboxSplit 
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   10
         Top             =   3960
         Width           =   2730
      End
      Begin VB.Frame fraToolboxTitle 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000009&
         Height          =   265
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   4080
         Width           =   2655
         Begin VB.Image imgWorkflowValue 
            Height          =   255
            Left            =   45
            Picture         =   "frmWorkflowWFToolbox.frx":9996
            Stretch         =   -1  'True
            Top             =   10
            Width           =   255
         End
         Begin VB.Label lblToolboxLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            Caption         =   "Workflow Values"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   9
            Top             =   45
            Width           =   1425
         End
      End
      Begin VB.Frame fraToolboxTitle 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000009&
         Height          =   265
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   6000
         Width           =   2655
         Begin VB.Image imgColumns 
            Height          =   255
            Left            =   45
            Picture         =   "frmWorkflowWFToolbox.frx":9D9F
            Stretch         =   -1  'True
            Top             =   10
            Width           =   255
         End
         Begin VB.Label lblToolboxLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            Caption         =   "Database Columns"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   7
            Top             =   45
            Width           =   1620
         End
      End
      Begin VB.Frame fraToolboxTitle 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   265
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2655
         Begin VB.Label lblToolboxLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            Caption         =   "Standard Element Items"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   5
            Top             =   0
            Width           =   2085
         End
         Begin VB.Image imgToolboxImage 
            Height          =   255
            Left            =   45
            Picture         =   "frmWorkflowWFToolbox.frx":A15B
            Stretch         =   -1  'True
            Top             =   10
            Width           =   255
         End
      End
      Begin ComctlLib.TreeView trvStandardControls 
         DragIcon        =   "frmWorkflowWFToolbox.frx":A51C
         Height          =   960
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1693
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   0
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
      Begin ComctlLib.TreeView trvColumns 
         DragIcon        =   "frmWorkflowWFToolbox.frx":A826
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   6480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1720
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   0
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
      Begin ComctlLib.TreeView trvWorkflowValue 
         DragIcon        =   "frmWorkflowWFToolbox.frx":AB30
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   4440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1931
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   0
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
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmWorkflowWFToolbox.frx":AE3A
      Height          =   1215
      Left            =   120
      TabIndex        =   36
      Top             =   7920
      Width           =   2760
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   27
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":AEEF
            Key             =   "IMG_WORKINGPATTERN"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":B441
            Key             =   "IMG_WEBFORM"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":B993
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":BCE5
            Key             =   "IMG_BUTTON"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":C237
            Key             =   "IMG_GRID"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":C789
            Key             =   "IMG_COLUMN"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":CCDB
            Key             =   "IMG_COMBOBOX"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":D22D
            Key             =   "IMG_DATE"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":D77F
            Key             =   "IMG_LINE"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":DCD1
            Key             =   "IMG_IMAGE"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":E223
            Key             =   "IMG_FRAME"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":E775
            Key             =   "IMG_LABEL"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":ECC7
            Key             =   "IMG_LINK"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":F219
            Key             =   "IMG_CHECKBOX"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":F76B
            Key             =   "IMG_LOOKUP"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":FCBD
            Key             =   "IMG_NUMERIC"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":1020F
            Key             =   "IMG_OLE"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":10761
            Key             =   "IMG_PHOTO"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":10CB3
            Key             =   "IMG_PROPERTIES"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":11205
            Key             =   "IMG_FILEUPLOAD"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":11757
            Key             =   "IMG_FILEDOWNLOAD"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":11CA9
            Key             =   "IMG_RADIO"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":121FB
            Key             =   "IMG_SPINNER"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":1274D
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":12C9F
            Key             =   "IMG_TEXTBOX"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":131F1
            Key             =   "IMG_TOOLBOX"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWorkflowWFToolbox.frx":13743
            Key             =   "IMG_WORKFLOW"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmWorkflowWFToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local variables to hold property values
Private mfrmWebFormDes As frmWorkflowWFDesigner

'Declare local variables
Dim mSngSplitStartY As Single
Dim mfSplitMoving As Boolean

Private Const iXGAP = 200
Private Const iYGAP = 200
Private Const iXFRAMEGAP = 150
Private Const iYFRAMEGAP = 100
Private Const iCOMPONENTFRAMEWIDTH = 2000
Private Const iFRAMEWIDTH = 5400
Private Const iFRAMEHEIGHT = 3900
Private Const iMINTOOLBOXHEIGHT = 600

Private Const MIN_FORM_HEIGHT = 3000
Private Const MIN_FORM_WIDTH = 2600

Private mlngPersonnelTableID As Long
Private miInitiationType As WorkflowInitiationTypes
Private mlngBaseTableID As Long



Public Sub EditMenu(ByVal psMenuOption As String)
  ' Pass any menu events onto the current screens
  ' 'frmWorkflowWFDesigner' form to handle.
  CurrentWebForm.EditMenu psMenuOption
End Sub

Public Property Get CurrentWebForm() As frmWorkflowWFDesigner
  ' Set the current web form designer property
  Set CurrentWebForm = mfrmWebFormDes
End Property

Public Property Set CurrentWebForm(pWebForm As frmWorkflowWFDesigner)
  ' Set the current web form designer property
  Set mfrmWebFormDes = pWebForm
  
  mlngBaseTableID = pWebForm.BaseTable
  miInitiationType = pWebForm.InitiationType
  
  If Not mfrmWebFormDes.Loading Then
    ' Populate the 'columns' treeview with the columns of the databases
    ' associated with the current screen.
    RefreshControls
  End If
  
End Property

Public Sub RefreshControls()
  RefreshStandardControlsTreeView
  RefreshWebFormValueTreeView
  RefreshColumnsTreeView
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

End Sub

Private Sub Form_Load()
  Dim iCYFrame As Integer
 
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
 
  ' Get then dimension of windows borders
  iCYFrame = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
  
  ' Position and size form
  Me.Move 0, 0, Screen.Width / 4, Forms(0).ScaleHeight
  
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  
  ' NPG20091118 Fault HRPRO-531
  imgToolboxImage.Visible = False
  imgColumns.Visible = False
  imgWorkflowValue.Visible = False
  lblToolboxLabel(0).Move 45, 45, 2070, 315
  lblToolboxLabel(1).Move 45, 45, 2070, 195
  lblToolboxLabel(2).Move 45, 45, 2070, 195
  
  Form_Resize
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, piUnloadMode As Integer)
  
  ' Only unload the form if really required.
  If piUnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
  ElseIf piUnloadMode <> vbFormCode Then
    Cancel = True
  End If

End Sub

Private Sub Form_Resize()

  ' If this form is not already minimised then ensure that all controls on this
  ' form are resized accordingly.
  If Me.WindowState <> vbMinimized Then
    
    With picToolbox
      .Left = 0
      .Top = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight
    End With
 
    SplitMove (0)
    SplitMove (1)
  End If

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

Private Sub picToolbox_Resize()

  Dim i As Integer
  
  For i = 0 To fraToolboxSplit.Count - 1
    With fraToolboxSplit(i)
      .Left = 0
      .Width = picToolbox.Width
    End With
  Next i
  
  For i = 0 To fraToolboxTitle.Count - 1
    With fraToolboxTitle(i)
      .Left = 0
      .Width = picToolbox.Width
    End With
  Next i
  
  With trvStandardControls
    .Left = iXFRAMEGAP
    .Width = picToolbox.Width - .Left
  End With
  
  With trvColumns
    .Left = iXFRAMEGAP
    .Width = picToolbox.Width - .Left
  End With
  
  With trvWorkflowValue
    .Left = iXFRAMEGAP
    .Width = picToolbox.Width - .Left
  End With
    
End Sub

Private Sub SplitMove(piSplitterIndex As Integer)

  Select Case piSplitterIndex
    Case 0
      ' Limit the minimum size of the tree/grid.
      If fraToolboxSplit(0).Top < iMINTOOLBOXHEIGHT Then
        fraToolboxSplit(0).Top = iMINTOOLBOXHEIGHT
      ElseIf (fraToolboxSplit(1).Top - fraToolboxSplit(0).Top) < iMINTOOLBOXHEIGHT Then
        fraToolboxSplit(0).Top = fraToolboxSplit(1).Top - iMINTOOLBOXHEIGHT
      End If
      
      trvStandardControls.Top = fraToolboxTitle(0).Top + fraToolboxTitle(0).Height
      trvStandardControls.Height = fraToolboxSplit(0).Top - trvStandardControls.Top
      fraToolboxTitle(2).Top = fraToolboxSplit(0).Top + fraToolboxSplit(0).Height
      trvWorkflowValue.Top = fraToolboxTitle(2).Top + fraToolboxTitle(2).Height
      trvWorkflowValue.Height = fraToolboxSplit(1).Top - trvWorkflowValue.Top
      
    Case 1
      ' Limit the minimum size of the tree/grid.
      If fraToolboxSplit(1).Top < (fraToolboxSplit(0).Top + iMINTOOLBOXHEIGHT) Then
        fraToolboxSplit(1).Top = (fraToolboxSplit(0).Top + iMINTOOLBOXHEIGHT)
      ElseIf (picToolbox.Height - fraToolboxSplit(1).Top) < (2 * iMINTOOLBOXHEIGHT) Then
        fraToolboxSplit(1).Top = picToolbox.Height - (2 * iMINTOOLBOXHEIGHT)
      End If
      
      trvWorkflowValue.Height = fraToolboxSplit(1).Top - trvWorkflowValue.Top
      fraToolboxTitle(1).Top = fraToolboxSplit(1).Top + fraToolboxSplit(1).Height
      trvColumns.Top = fraToolboxTitle(1).Top + fraToolboxTitle(1).Height
      trvColumns.Height = picToolbox.Height - trvColumns.Top
   
  End Select

  ' Flag that the split move has ended.
  mfSplitMoving = False

End Sub

Private Function GetUsedTables() As String
  ' Get a list of tableIDs used in the preceding web forms.
  ' Include the ascendant table records of those identified.
  ' eg. If we've identified an Absence record, we can use this to identifiy the related Personnel record.
  Dim aWFPrecedingElements() As VB.Control
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sTableList As String
  Dim alngValidTables() As Long
  
  ReDim aWFPrecedingElements(0)
  
  mfrmWebFormDes.PrecedingElements aWFPrecedingElements
 
  'Check if there are any preceding record identifying elements (StoredData or WebForms with RecordSelectors).
  For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfTemp = aWFPrecedingElements(iLoop)
    
    If wfTemp.ElementType = elem_WebForm Then
      asItems = wfTemp.Items
  
      For iLoop2 = 1 To UBound(asItems, 2)
        If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then

          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants CLng(asItems(44, iLoop2)), alngValidTables

          For lngLoop = 1 To UBound(alngValidTables)
            sTableList = sTableList & IIf(Len(sTableList) > 0, ",", vbNullString) & CStr(alngValidTables(lngLoop))
          Next lngLoop
        End If
      Next iLoop2
    ElseIf wfTemp.ElementType = elem_StoredData Then
      ReDim alngValidTables(0)
      TableAscendants wfTemp.DataTableID, alngValidTables

      'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
      'If wfTemp.DataAction = DATAACTION_DELETE Then
      '  ' Cannot do anything with a Deleted record, but can use its ascendants.
      '  ' Remove the table itself from the array of valid tables.
      '  alngValidTables(1) = 0
      'End If

      For lngLoop = 1 To UBound(alngValidTables)
        sTableList = sTableList & IIf(Len(sTableList) > 0, ",", vbNullString) & CStr(alngValidTables(lngLoop))
      Next lngLoop
    End If
  
    Set wfTemp = Nothing
  Next iLoop
  
  GetUsedTables = sTableList
  
End Function

Private Sub RefreshColumnsTreeView()
  ' Populate the treeview of database column controls.
  Dim objNode As ComctlLib.Node
  Dim rsColumns As dao.Recordset
  Dim fPopulated As Boolean
  Dim sSQL As String
  Dim sIconKey As String
  Dim sTableList As String
  Dim lngLastTableID As Long
  Dim alngValidTables() As Long
  Dim lngLoop As Long
  Dim fNoValidColumns As Boolean
  
  fPopulated = False
  lngLastTableID = -1
  fNoValidColumns = True
  
  ' Remove any existing nodes from the treeview.
  trvColumns.Nodes.Clear

  sTableList = GetUsedTables
  If Len(sTableList) = 0 Then
    sTableList = "0"
  End If
  
  ReDim alngValidTables(0)
  If miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
    TableAscendants mlngPersonnelTableID, alngValidTables
  ElseIf miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
    TableAscendants mlngBaseTableID, alngValidTables
  End If
  
  For lngLoop = 1 To UBound(alngValidTables)
    sTableList = sTableList & IIf(Len(sTableList) > 0, ",", vbNullString) & CStr(alngValidTables(lngLoop))
  Next lngLoop
    
  ' Get column details for the primary table.
  sSQL = "SELECT tmpColumns.tableID, tmpColumns.columnID, tmpColumns.datatype, tmpColumns.columnName, tmpColumns.columnType, tmpColumns.deleted, tmpColumns.controlType, tmpTables.tableName " & _
    " FROM tmpColumns, tmpTables " & _
    " WHERE (tmpColumns.tableID IN (" & sTableList & "))" & _
    " AND tmpColumns.deleted = FALSE" & _
    " AND tmpColumns.columnType <>" & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
    " AND tmpColumns.columnType <>" & Trim(Str(giCOLUMNTYPE_LINK)) & _
    " AND tmpColumns.TableID = tmpTables.TableID "
  sSQL = sSQL & _
    " ORDER BY tmpTables.tableName, tmpColumns.columnName"

  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
  
  ' Loop though columns and populate treeview
  Do While Not rsColumns.EOF
    fNoValidColumns = False
    
    ' Add a node to the treeview for any parent tables.
    If rsColumns!TableID <> lngLastTableID Then
      lngLastTableID = rsColumns!TableID
      
      Set objNode = trvColumns.Nodes.Add(, tvwChild, _
        "T" & rsColumns!TableID, rsColumns!TableName, "IMG_TABLE", "IMG_TABLE")
      objNode.Sorted = True
      objNode.Expanded = False
      Set objNode = Nothing
    End If
                 
    If rsColumns!DataType = SQLDataType.sqlNumeric Then
      sIconKey = "IMG_NUMERIC"
    Else
      ' Get the correct icon for the current column.
      sIconKey = GetColumnIcon(rsColumns!ControlType)
    End If
    
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
   
  If fNoValidColumns Then
    Set objNode = trvColumns.Nodes.Add(, , , "<No identifiable column values>")
    trvColumns.LineStyle = tvwTreeLines
    objNode.Expanded = True
  End If
   
  fPopulated = True
  
End Sub

Private Sub RefreshStandardControlsTreeView()
  '
  ' Populate the treeview of standard controls.
  '
  Dim objNode As ComctlLib.Node
  
  ' Clear the treeview
  trvStandardControls.Nodes.Clear
  
  Set objNode = trvStandardControls.Nodes.Add(, , "STDROOT", "Standard Controls", "IMG_TOOLBOX", "IMG_TOOLBOX")
  objNode.Expanded = True
    
  ' Add the standard controls to the tree view.
  
  ' Add the Button control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "BUTTON", "Button", "IMG_BUTTON", "IMG_BUTTON")
  objNode.Expanded = True
  
  ' Add the Frame control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "FRAMECTRL", "Frame", "IMG_FRAME", "IMG_FRAME")
  objNode.Expanded = True
  
  ' Add the Image control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "IMAGECTRL", "Image", "IMG_IMAGE", "IMG_IMAGE")
  objNode.Expanded = True
  
  '----------------------------------------------------------------------------
  ' Add the Input Value node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "INPUT", "Input", "IMG_TEXTBOX", "IMG_TEXTBOX")
  objNode.Expanded = True

    ' Add the Input Value - Character node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_CHARACTER", "Character", "IMG_TEXTBOX", "IMG_TEXTBOX")
    objNode.Expanded = True
  
    ' Add the Input Value - Date node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_DATE", "Date", "IMG_DATE", "IMG_DATE")
    objNode.Expanded = True

    ' Add the Input Value - Dropdown List.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_DROPDOWN", "Dropdown", "IMG_COMBOBOX", "IMG_COMBOBOX")
    objNode.Expanded = True

    ' Add the Input Value - File.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_FILEUPLOAD", "File Upload", "IMG_FILEUPLOAD", "IMG_FILEUPLOAD")
    objNode.Expanded = True

    ' Add the Input Value - Logic node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_LOGIC", "Logic", "IMG_CHECKBOX", "IMG_CHECKBOX")
    objNode.Expanded = True
  
    ' Add the Input Value - Lookup node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_LOOKUP", "Lookup", "IMG_LOOKUP", "IMG_LOOKUP")
    objNode.Expanded = True
  
    ' Add the Input Value - Numeric node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_NUMERIC", "Numeric", "IMG_NUMERIC", "IMG_NUMERIC")
    objNode.Expanded = True
    
    ' Add the Input Value - Option Group node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_OPTIONGROUP", "Option Group", "IMG_RADIO", "IMG_RADIO")
    objNode.Expanded = True
    
    ' Add the Input Value - Grid node.
    Set objNode = trvStandardControls.Nodes.Add("INPUT", tvwChild, "INPUT_GRID", "Record Selector", "IMG_GRID", "IMG_GRID")
    objNode.Expanded = True
  
  '----------------------------------------------------------------------------
  
  ' Add the Label control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "LABELCTRL", "Label", "IMG_LABEL", "IMG_LABEL")
  objNode.Expanded = True
  
  ' Add the Line control node.
  Set objNode = trvStandardControls.Nodes.Add("STDROOT", tvwChild, "LINECTRL", "Line", "IMG_LINE", "IMG_LINE")
  objNode.Expanded = True
  
  ' Disassociate the objNode variable.
  Set objNode = Nothing
  
  ' Default standard controls toolbox treeview to show all nodes.
  Do While trvStandardControls.GetVisibleCount < trvStandardControls.Nodes.Count
    trvStandardControls.Height = trvStandardControls.Height + 50
  Loop
  fraToolboxSplit(0).Top = trvStandardControls.Top + trvStandardControls.Height

End Sub

Private Sub trvWorkflowValue_Click()
  ' Mark that there is a currently selected item.
  If Not trvWorkflowValue.SelectedItem Is Nothing Then
    trvWorkflowValue.SelectedItem.Selected = True
  End If
End Sub

Private Sub trvWorkflowValue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ThisNode As ComctlLib.Node
    
  ' If the left mouse button is pressed ...
  If Button = vbLeftButton Then
    
    'Get the node at the mouse position
    Set ThisNode = trvWorkflowValue.HitTest(x, y)
    
    ' If we have selected a valid node ...
    If Not ThisNode Is Nothing Then
      
      If Not ThisNode.Parent Is Nothing Then
        ' Ensure this node is not the selected node.
        If Not ThisNode Is trvWorkflowValue.SelectedItem Then
          Set trvWorkflowValue.SelectedItem = ThisNode
        End If
        
        'NHRD20092006 Fault 10990
        'Added all of the possible Icon so far to futureproof it a bit
        Select Case ThisNode.SelectedImage
           Case "IMG_FRAME"
               trvWorkflowValue.DragIcon = picDragIcon_Frame.Picture
           Case "IMG_DATE"
               trvWorkflowValue.DragIcon = picDragIcon_Date.Picture
           Case "IMG_BUTTON"
               trvWorkflowValue.DragIcon = picDragIcon_Button.Picture
           Case "IMG_IMAGE"
               trvWorkflowValue.DragIcon = picDragIcon_Image.Picture
           Case "IMG_TEXTBOX"
               trvWorkflowValue.DragIcon = picDragIcon_Textbox.Picture
           Case "IMG_CHECKBOX"
               trvWorkflowValue.DragIcon = picDragIcon_CheckBox.Picture
           Case "IMG_NUMERIC"
               trvWorkflowValue.DragIcon = picDragIcon_Numeric.Picture
           Case "IMG_GRID"
               trvWorkflowValue.DragIcon = picDragIcon_Grid.Picture
           Case "IMG_LABEL"
               trvWorkflowValue.DragIcon = picDragIcon_Label.Picture
           Case "IMG_LINE"
               trvWorkflowValue.DragIcon = picDragIcon_Line.Picture
           Case "IMG_WORKINGPATTERN"
               trvWorkflowValue.DragIcon = picDragIcon_WorkingPattern.Picture
           Case "IMG_WEBFORM"
               trvWorkflowValue.DragIcon = picDragIcon_WebForm.Picture
           Case "IMG_COLUMN"
               trvWorkflowValue.DragIcon = picDragIcon_Column.Picture
           Case "IMG_COMBOBOX"
               trvWorkflowValue.DragIcon = picDragIcon_ComboBox.Picture
           Case "IMG_LINK"
               trvWorkflowValue.DragIcon = picDragIcon_Link.Picture
           Case "IMG_LOOKUP"
               trvWorkflowValue.DragIcon = picDragIcon_Lookup.Picture
           Case "IMG_PHOTO"
               trvWorkflowValue.DragIcon = picDragIcon_Photo.Picture
           Case "IMG_OLE"
               trvWorkflowValue.DragIcon = picDragIcon_OLE.Picture
           Case "IMG_WORKFLOW"
               trvWorkflowValue.DragIcon = picDragIcon_WorkFlow.Picture
           Case "IMG_PROPERTIES"
               trvWorkflowValue.DragIcon = picDragIcon_Properties.Picture
           Case "IMG_RADIO"
               trvWorkflowValue.DragIcon = picDragIcon_Radio.Picture
           Case "IMG_SPINNER"
               trvWorkflowValue.DragIcon = picDragIcon_Spinner.Picture
           Case "IMG_TABLE"
               trvWorkflowValue.DragIcon = picDragIcon_Table.Picture
           Case "IMG_TOOLBOX"
               trvWorkflowValue.DragIcon = picDragIcon_ToolBox.Picture
           Case "IMG_PAGETAB"
               trvWorkflowValue.DragIcon = picDragIcon_PageTab.Picture
           Case "IMG_FILEUPLOAD"
               trvWorkflowValue.DragIcon = picDragIcon_FileUpload.Picture
           Case "IMG_FILEDOWNLOAD"
               trvWorkflowValue.DragIcon = picDragIcon_FileDownload.Picture
           Case Else
               trvWorkflowValue.DragIcon = picDragIcon_ColumnDrag.Picture
        End Select
        'Begin drag
        trvWorkflowValue.Drag vbBeginDrag
      End If
    End If
    Set ThisNode = Nothing
  End If
End Sub

Private Sub trvWorkflowValue_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Cancel the control drag operation.
  trvWorkflowValue.Drag vbEndDrag
End Sub

Private Sub RefreshWebFormValueTreeView()

  ' Populate the WF WebForm Treeview and
  ' select the current webform if it is still valid.
  Dim lngIndex As Long
  Dim aWFPrecedingElements() As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sImage As String
  Dim objNode As ComctlLib.Node
  Dim objRootNode As ComctlLib.Node
  Dim fRootNodeDone As Boolean
  Dim fNoValidWebForms As Boolean
  Dim avToolboxItems() As Variant
  
  fNoValidWebForms = True
  
  ' Clear the current contents of the combo.
  trvWorkflowValue.Nodes.Clear
  
  ReDim aWFPrecedingElements(0)
  mfrmWebFormDes.PrecedingElements aWFPrecedingElements
 
  'Check if there are any preceding INPUT elements.
  For iLoop = 2 To UBound(aWFPrecedingElements)
    If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
      fRootNodeDone = False
      
      Set wfTemp = aWFPrecedingElements(iLoop)
      asItems = wfTemp.Items
  
      ' Create an array of the items to be put in the toolbox.
      ' This array array is then sorted into alphabtic order before
      ' being used to populate the treeview.
      ' Column 0 = identifier (display text)
      ' Column 1 = asItems index
      ReDim avToolboxItems(1, 0)
      
      For iLoop2 = 1 To UBound(asItems, 2)
        If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_CHAR) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_DATE) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOGIC) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) _
          Or (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Then
          
          If Not fRootNodeDone Then
            Set objRootNode = trvWorkflowValue.Nodes.Add(, , "E" & aWFPrecedingElements(iLoop).Identifier, aWFPrecedingElements(iLoop).Identifier, "IMG_WEBFORM", "IMG_WEBFORM")
            fRootNodeDone = True
            fNoValidWebForms = False
          End If
          
          ReDim Preserve avToolboxItems(1, UBound(avToolboxItems, 2) + 1)
          avToolboxItems(0, UBound(avToolboxItems, 2)) = UCase(asItems(9, iLoop2))
          avToolboxItems(1, UBound(avToolboxItems, 2)) = iLoop2
        End If
      Next iLoop2
          
      ' Sort the array of toolbox items into alphabetic order.
      ShellSortArray avToolboxItems

      ' Populate the toolbox treeview
      For iLoop2 = 1 To UBound(avToolboxItems, 2)
        lngIndex = avToolboxItems(1, iLoop2)
        
        Select Case CInt(asItems(2, lngIndex))
          Case giWFFORMITEM_INPUTVALUE_CHAR
            sImage = "IMG_TEXTBOX"
          Case giWFFORMITEM_INPUTVALUE_DATE
            sImage = "IMG_DATE"
          Case giWFFORMITEM_INPUTVALUE_NUMERIC
            sImage = "IMG_NUMERIC"
          Case giWFFORMITEM_INPUTVALUE_LOGIC
            sImage = "IMG_CHECKBOX"
          Case giWFFORMITEM_INPUTVALUE_DROPDOWN
            sImage = "IMG_COMBOBOX"
          Case giWFFORMITEM_INPUTVALUE_LOOKUP
            sImage = "IMG_LOOKUP"
          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
            sImage = "IMG_RADIO"
          Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
            sImage = "IMG_FILEDOWNLOAD"
          Case Else
            sImage = ""
        End Select

        If Len(sImage) > 0 Then
          Set objNode = trvWorkflowValue.Nodes.Add("E" & aWFPrecedingElements(iLoop).Identifier, _
            tvwChild, _
            "E" & aWFPrecedingElements(iLoop).Identifier & " " & asItems(9, lngIndex), _
            asItems(9, lngIndex), _
            sImage, _
            sImage)
        End If
          
        If Not objNode Is Nothing Then objNode.Expanded = True
      Next iLoop2
    End If
  Next iLoop
 
  If fNoValidWebForms Then
    Set objNode = trvWorkflowValue.Nodes.Add(, , , "<No preceding workflow values>")
    trvWorkflowValue.LineStyle = tvwTreeLines
    objNode.Expanded = True
  End If
    
  Set objNode = Nothing
  Set objRootNode = Nothing
  
End Sub

Private Sub ShellSortArray(vArray As Variant)
  ' Sort the given array by column 0.
  ' Assumes column 0 is a string.
  ' Assumes column 1 is a long.
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim sTemp As String
  Dim lngTemp As Long

  lHValue = LBound(vArray, 2)
  
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray, 2)
  
  Do
    lHValue = lHValue / 3
    
    For lLoop1 = lHValue + LBound(vArray, 2) To UBound(vArray, 2)
      sTemp = vArray(0, lLoop1)
      lngTemp = vArray(1, lLoop1)
      
      lHold = lLoop1
      
      Do While vArray(0, lHold - lHValue) > sTemp

        vArray(0, lHold) = vArray(0, lHold - lHValue)
        vArray(1, lHold) = vArray(1, lHold - lHValue)
        
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      
      vArray(0, lHold) = sTemp
      vArray(1, lHold) = lngTemp
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Disassociate global variables and the form itself.
  Set mfrmWebFormDes = Nothing
  Set frmWorkflowWFToolbox = Nothing
  
  Unhook Me.hWnd
End Sub

Private Sub fraToolboxSplit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Record the split move start position.
  mSngSplitStartY = y
  ' Flag that the split is being moved.
  mfSplitMoving = True
End Sub

Private Sub fraToolboxSplit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' If we are moving the split then move it.
  If mfSplitMoving Then
    fraToolboxSplit(Index).Top = fraToolboxSplit(Index).Top + (y - mSngSplitStartY)
  End If
End Sub

Private Sub fraToolboxSplit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' If the split is being moved then call the routine that resizes the
  ' tree and list views accordingly.
  If mfSplitMoving Then
    SplitMove (Index)
  End If
End Sub

Private Sub trvColumns_Click()
  ' Mark that there is a currently selected item.
  If Not trvColumns.SelectedItem Is Nothing Then
    trvColumns.SelectedItem.Selected = True
  End If

End Sub

Private Sub trvColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ThisNode As ComctlLib.Node
  
  ' If the left mouse button is pressed ...
  If Button = vbLeftButton Then
  
    'Get the node at the mouse position
    Set ThisNode = trvColumns.HitTest(x, y)
    
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
               trvColumns.DragIcon = picDragIcon_FileUpload.Picture
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

Private Sub trvColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Signal that the drag operation has ended.
  trvColumns.Drag vbEndDrag
End Sub

Private Sub trvStandardControls_Click()
  ' Mark that there is a currently selected item.
  If Not trvStandardControls.SelectedItem Is Nothing Then
    trvStandardControls.SelectedItem.Selected = True
  End If
End Sub

Private Sub trvStandardControls_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ThisNode As ComctlLib.Node
    
  ' If the left mouse button is pressed ...
  If Button = vbLeftButton Then
    
    'Get the node at the mouse position
    Set ThisNode = trvStandardControls.HitTest(x, y)
    
    ' If we have selected a valid node ...
    If Not ThisNode Is Nothing Then
   
      ' Only allow columns to be selected (ie. not table name nodes).
      If (Not ThisNode.Parent Is Nothing) And (ThisNode.key <> "INPUT") Then
   
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
                trvStandardControls.DragIcon = picDragIcon_FileUpload.Picture
            Case Else
                trvStandardControls.DragIcon = picDragIcon_ColumnDrag.Picture
        End Select
        'Begin drag
        trvStandardControls.Drag vbBeginDrag
      End If
    End If
    Set ThisNode = Nothing
  End If
End Sub

Private Sub trvStandardControls_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Cancel the control drag operation.
  trvColumns.Drag vbEndDrag
End Sub

Private Function GetColumnIcon(piColumnType As Integer) As String
  Dim iLoop As Integer
  
  ' Depending on the control type of the current column, determine the
  ' associated icon key, as defined in the imageList2 control.
  Select Case piColumnType
    Case giCTRL_TEXTBOX
      GetColumnIcon = "IMG_TEXTBOX"
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
    Case Else
      GetColumnIcon = "IMG_COLUMN"
    End Select
  
End Function
