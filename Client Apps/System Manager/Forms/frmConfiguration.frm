VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "COA_ColourPicker.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmConfiguration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Configuration"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5009
   Icon            =   "frmConfiguration.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin COAColourPicker.COA_ColourPicker ColorPicker 
      Left            =   720
      Top             =   5760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   4380
      TabIndex        =   21
      Top             =   5800
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5640
      TabIndex        =   22
      Top             =   5800
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   120
      Top             =   5805
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   150
      TabIndex        =   23
      Top             =   105
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "&Email"
      TabPicture(0)   =   "frmConfiguration.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTestSQLMail"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEmailOptions"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraEmailSetup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Processing"
      TabPicture(1)   =   "frmConfiguration.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSQL2005"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Display"
      TabPicture(2)   =   "frmConfiguration.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmExpressions"
      Tab(2).Control(1)=   "fraGeneral"
      Tab(2).Control(2)=   "frmBackgrounds"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Dev"
      TabPicture(3)   =   "frmConfiguration.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraQuickAddress"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "frmDeveloperAFD"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&Advanced"
      TabPicture(4)   =   "frmConfiguration.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraAdvancedSettings"
      Tab(4).Control(1)=   "fraOutlookCalendar"
      Tab(4).Control(2)=   "fraTime"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Desi&gners"
      TabPicture(5)   =   "frmConfiguration.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Web"
      TabPicture(6)   =   "frmConfiguration.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Web Information :"
         Height          =   5020
         Left            =   -74850
         TabIndex        =   88
         Top             =   400
         Width           =   6465
         Begin VB.TextBox txtWebSiteAddress 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1950
            TabIndex        =   89
            Top             =   300
            Width           =   4275
         End
         Begin VB.Label lblWebSiteAddress 
            Caption         =   "Web Site Address :"
            Height          =   240
            Left            =   195
            TabIndex        =   90
            Top             =   360
            Width           =   1665
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Defaults :"
         Height          =   5020
         Left            =   -74850
         TabIndex        =   84
         Top             =   400
         Width           =   6465
         Begin VB.CommandButton cmdDefaultScreenFont 
            Caption         =   "..."
            Height          =   315
            Left            =   3750
            TabIndex        =   87
            ToolTipText     =   "Select Path"
            Top             =   285
            Width           =   330
         End
         Begin VB.TextBox txtDefaultScreenFont 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1065
            TabIndex        =   86
            Top             =   300
            Width           =   2700
         End
         Begin VB.Label lblDefaultScreenFont 
            Caption         =   "Font :"
            Height          =   255
            Left            =   195
            TabIndex        =   85
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame fraAdvancedSettings 
         Caption         =   "Database Settings : "
         Height          =   2055
         Left            =   -74850
         TabIndex        =   77
         Top             =   3375
         Width           =   6465
         Begin VB.CheckBox chkDisableSpecialFunctionAutoUpdate 
            Caption         =   "Di&sable immediate update of columns using the following functions :"
            Height          =   285
            Left            =   180
            TabIndex        =   78
            Top             =   315
            Width           =   6180
         End
         Begin VB.CheckBox chkRecursionLevelManual 
            Caption         =   "Manual Recursion Level :"
            Height          =   285
            Left            =   180
            TabIndex        =   80
            Top             =   1515
            Width           =   2445
         End
         Begin COASpinner.COA_Spinner spnTriggelLevel 
            Height          =   315
            Left            =   2700
            TabIndex        =   81
            Top             =   1515
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   30
            MinimumValue    =   1
            Text            =   "8"
         End
         Begin VB.Label chkSpecialFunctions 
            Caption         =   $"frmConfiguration.frx":00D0
            Height          =   795
            Left            =   615
            TabIndex        =   79
            Top             =   615
            Width           =   3000
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraOutlookCalendar 
         Caption         =   "Outlook Calendar :"
         Height          =   1485
         Left            =   -74850
         TabIndex        =   64
         Top             =   400
         Width           =   6465
         Begin TDBMask6Ctl.TDBMask TDBAMStartTime 
            Height          =   300
            Left            =   1650
            TabIndex        =   66
            Top             =   300
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   529
            Caption         =   "frmConfiguration.frx":0161
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmConfiguration.frx":01C6
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "0"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "00:00"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TDBAMEndTime 
            Height          =   300
            Left            =   4335
            TabIndex        =   68
            Top             =   300
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   529
            Caption         =   "frmConfiguration.frx":0208
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmConfiguration.frx":026D
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "0"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "00:00"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TDBPMStartTime 
            Height          =   300
            Left            =   1650
            TabIndex        =   70
            Top             =   705
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   529
            Caption         =   "frmConfiguration.frx":02AF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmConfiguration.frx":0314
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "0"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "00:00"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TDBPMEndTime 
            Height          =   300
            Left            =   4335
            TabIndex        =   72
            Top             =   705
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   529
            Caption         =   "frmConfiguration.frx":0356
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmConfiguration.frx":03BB
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "0"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "00:00"
            Value           =   ""
         End
         Begin VB.Label lblPMEndTime 
            AutoSize        =   -1  'True
            Caption         =   "PM End Time :"
            Height          =   195
            Left            =   3000
            TabIndex        =   71
            Top             =   765
            Width           =   1365
         End
         Begin VB.Label lblPMStartTime 
            AutoSize        =   -1  'True
            Caption         =   "PM Start Time :"
            Height          =   195
            Left            =   195
            TabIndex        =   69
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label lblAMEndTime 
            AutoSize        =   -1  'True
            Caption         =   "AM End Time :"
            Height          =   195
            Left            =   3000
            TabIndex        =   67
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label lblAMStartTime 
            AutoSize        =   -1  'True
            Caption         =   "AM Start Time :"
            Height          =   195
            Left            =   195
            TabIndex        =   65
            Top             =   360
            Width           =   1470
         End
      End
      Begin VB.Frame fraTime 
         Caption         =   "Overnight Processing :"
         Height          =   1365
         Left            =   -74850
         TabIndex        =   73
         Top             =   1920
         Width           =   6465
         Begin VB.CheckBox chkReorganiseIndexes 
            Caption         =   "De&fragment indexes in overnight job"
            Height          =   285
            Left            =   225
            TabIndex        =   76
            Top             =   675
            Width           =   3660
         End
         Begin TDBMask6Ctl.TDBMask TDBMaskTime 
            Height          =   300
            Left            =   1560
            TabIndex        =   75
            Top             =   300
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   529
            Caption         =   "frmConfiguration.frx":03FD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmConfiguration.frx":0462
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99:99"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "0"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "00:00:00"
            Value           =   ""
         End
         Begin VB.Label lblOccurs 
            Caption         =   "Start Time : "
            Height          =   255
            Left            =   195
            TabIndex        =   74
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frmDeveloperAFD 
         Caption         =   "AFD Evaluation :"
         Height          =   1250
         Left            =   -74850
         TabIndex        =   52
         Top             =   400
         Width           =   6465
         Begin VB.TextBox txtDeveloperAFDPlus 
            Height          =   285
            Left            =   1980
            TabIndex        =   57
            Top             =   825
            Width           =   1575
         End
         Begin VB.TextBox txtDeveloperAFDNormal 
            Height          =   285
            Left            =   200
            TabIndex        =   55
            Top             =   825
            Width           =   1575
         End
         Begin VB.TextBox txtDeveloperAFDNamesNumbers 
            Height          =   285
            Left            =   3825
            TabIndex        =   59
            Top             =   825
            Width           =   1710
         End
         Begin VB.CheckBox chkAllowAFDEvaluation 
            Caption         =   "Allow AFD Evaluation Software to be used"
            Height          =   195
            Left            =   200
            TabIndex        =   53
            Top             =   360
            Width           =   4035
         End
         Begin VB.Label lblDeveloperAFDPlus 
            Caption         =   "Plus:"
            Height          =   255
            Left            =   1995
            TabIndex        =   56
            Top             =   615
            Width           =   1185
         End
         Begin VB.Label lblDeveloperAFDNormal 
            Caption         =   "Normal:"
            Height          =   195
            Left            =   195
            TabIndex        =   54
            Top             =   615
            Width           =   1305
         End
         Begin VB.Label lblDeveloperAFDNamesNumbers 
            Caption         =   "Names && Numbers:"
            Height          =   240
            Left            =   3840
            TabIndex        =   58
            Top             =   615
            Width           =   2250
         End
      End
      Begin VB.Frame frmExpressions 
         Caption         =   "Filters / Calculations :"
         Height          =   1260
         Left            =   -74850
         TabIndex        =   44
         Top             =   2400
         Width           =   6465
         Begin VB.ComboBox cboNodeSize 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":04A4
            Left            =   1920
            List            =   "frmConfiguration.frx":04A6
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   700
            Width           =   2790
         End
         Begin VB.ComboBox cboColours 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":04A8
            Left            =   1920
            List            =   "frmConfiguration.frx":04AA
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   300
            Width           =   2790
         End
         Begin VB.Label lblExpandNodes 
            AutoSize        =   -1  'True
            Caption         =   "Expand Nodes :"
            Height          =   195
            Left            =   195
            TabIndex        =   47
            Top             =   765
            Width           =   1635
         End
         Begin VB.Label lblViewInColour 
            AutoSize        =   -1  'True
            Caption         =   "View In Colour :"
            Height          =   195
            Left            =   195
            TabIndex        =   45
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.Frame fraGeneral 
         Caption         =   "General :"
         Height          =   1710
         Left            =   -74850
         TabIndex        =   49
         Top             =   3720
         Width           =   6465
         Begin VB.CheckBox chkMaximize 
            Caption         =   "Ma&ximize Screens On Entry"
            Height          =   195
            Left            =   195
            TabIndex        =   50
            Top             =   360
            Width           =   3135
         End
         Begin VB.CheckBox chkRememberDBColumns 
            Caption         =   "&Remember Database Columns View"
            Height          =   195
            Left            =   195
            TabIndex        =   51
            Top             =   760
            Width           =   3840
         End
      End
      Begin VB.Frame fraQuickAddress 
         Caption         =   "Quick Address "
         Height          =   3740
         Left            =   -74850
         TabIndex        =   60
         Top             =   1700
         Width           =   6465
         Begin VB.CheckBox chkAllowQAddressEvaluation 
            Caption         =   "Allow Quick Address Evaluation Software to be used"
            Height          =   255
            Left            =   200
            TabIndex        =   61
            Top             =   315
            Width           =   4890
         End
         Begin VB.TextBox txtQAddressSeedValue 
            Height          =   300
            Left            =   210
            TabIndex        =   63
            Top             =   900
            Width           =   1560
         End
         Begin VB.Label lblDeveloperQAddressNormal 
            Caption         =   "Normal:"
            Height          =   180
            Left            =   240
            TabIndex        =   62
            Top             =   660
            Width           =   765
         End
      End
      Begin VB.Frame fraSQL2005 
         Caption         =   "Processing Account :"
         Height          =   5020
         Left            =   -74865
         TabIndex        =   24
         Top             =   400
         Width           =   6465
         Begin VB.ComboBox cboProcessMethod 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":04AC
            Left            =   2010
            List            =   "frmConfiguration.frx":04AE
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   300
            Width           =   4275
         End
         Begin VB.TextBox txtLogin 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2010
            TabIndex        =   27
            Top             =   720
            Width           =   4275
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2010
            PasswordChar    =   "*"
            TabIndex        =   29
            Top             =   1110
            Width           =   4275
         End
         Begin VB.TextBox txtConfirmPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2010
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   1500
            Width           =   4275
         End
         Begin VB.CommandButton cmdTestLogon 
            Caption         =   "&Test Login"
            Height          =   400
            Left            =   5070
            TabIndex        =   32
            Top             =   1905
            Width           =   1200
         End
         Begin VB.Label lblProcessAdminWarning 
            Caption         =   "You do not have permission to assign Process Admin priviledge to logins. Please contact your system administrator."
            ForeColor       =   &H000000FF&
            Height          =   915
            Left            =   225
            TabIndex        =   83
            Top             =   1980
            Visible         =   0   'False
            Width           =   4650
         End
         Begin VB.Label lblProcessMethod 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Method :"
            Height          =   195
            Left            =   195
            TabIndex        =   82
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lblConfirmPassword 
            Caption         =   "Confirm Password : "
            Height          =   300
            Left            =   180
            TabIndex        =   30
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblLogin 
            Caption         =   "Login :"
            Height          =   240
            Left            =   180
            TabIndex        =   26
            Top             =   765
            Width           =   1320
         End
         Begin VB.Label lblPassword 
            Caption         =   "Password :"
            Height          =   240
            Left            =   180
            TabIndex        =   28
            Top             =   1170
            Width           =   1455
         End
      End
      Begin VB.Frame frmBackgrounds 
         Caption         =   "Background :"
         Height          =   1950
         Left            =   -74850
         TabIndex        =   33
         Top             =   400
         Width           =   6465
         Begin VB.CommandButton cmdColourPicker 
            Caption         =   "..."
            Height          =   315
            Left            =   4380
            TabIndex        =   36
            ToolTipText     =   "Select Colour"
            Top             =   300
            Width           =   330
         End
         Begin VB.PictureBox picHolder 
            Height          =   1470
            Left            =   4875
            ScaleHeight     =   1410
            ScaleWidth      =   1410
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   300
            Width           =   1470
            Begin VB.Image picWork 
               Height          =   855
               Left            =   255
               Stretch         =   -1  'True
               Top             =   0
               Width           =   930
            End
         End
         Begin VB.ComboBox cboBitmapLocation 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1100
            Width           =   2790
         End
         Begin VB.CommandButton cmdPictureSelect 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   39
            ToolTipText     =   "Select Path"
            Top             =   700
            Width           =   330
         End
         Begin VB.TextBox txtDeskTopBitmapName 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   700
            Width           =   2160
         End
         Begin VB.CommandButton cmdPictureClear 
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   20.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4380
            MaskColor       =   &H000000FF&
            TabIndex        =   40
            ToolTipText     =   "Clear Path"
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.Label lblBackColour 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1920
            TabIndex        =   35
            Top             =   300
            Width           =   2460
         End
         Begin VB.Label Label1 
            Caption         =   "Desktop Colour : "
            Height          =   255
            Left            =   195
            TabIndex        =   34
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label lblBitmapLocation 
            Caption         =   "Location : "
            Height          =   270
            Left            =   200
            TabIndex        =   41
            Top             =   1160
            Width           =   1245
         End
         Begin VB.Label lblDesktopBitmapName 
            Caption         =   "Desktop Bitmap : "
            Height          =   255
            Left            =   195
            TabIndex        =   37
            Top             =   765
            Width           =   1605
         End
      End
      Begin VB.Frame fraEmailSetup 
         Caption         =   "Setup :"
         Height          =   2535
         Left            =   150
         TabIndex        =   0
         Top             =   400
         Width           =   6465
         Begin VB.CommandButton cmdEmailTest 
            Caption         =   "&Test Email"
            Height          =   400
            Left            =   5085
            TabIndex        =   9
            Top             =   1920
            Width           =   1200
         End
         Begin VB.ComboBox cboEmailMethod 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":04B0
            Left            =   1875
            List            =   "frmConfiguration.frx":04B2
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   270
            Width           =   4410
         End
         Begin VB.ComboBox cboEmailProfile 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":04B4
            Left            =   1875
            List            =   "frmConfiguration.frx":04B6
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   700
            Width           =   4410
         End
         Begin VB.TextBox txtEmailAccount 
            Height          =   315
            Left            =   1875
            TabIndex        =   8
            Top             =   1500
            Width           =   4410
         End
         Begin VB.TextBox txtEmailServer 
            Height          =   315
            Left            =   1875
            TabIndex        =   6
            Top             =   1100
            Width           =   4410
         End
         Begin VB.Label lblEmailMethod 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Method :"
            Height          =   195
            Left            =   195
            TabIndex        =   1
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lblEmailProfile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profile :"
            Height          =   195
            Left            =   195
            TabIndex        =   3
            Top             =   760
            Width           =   555
         End
         Begin VB.Label lblEmailAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account :"
            Height          =   195
            Left            =   195
            TabIndex        =   7
            Top             =   1560
            Width           =   690
         End
         Begin VB.Label lblEmailServer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server :"
            Height          =   195
            Left            =   195
            TabIndex        =   5
            Top             =   1160
            Width           =   585
         End
      End
      Begin VB.Frame fraEmailOptions 
         Caption         =   "Options :"
         Height          =   1190
         Left            =   150
         TabIndex        =   14
         Top             =   4240
         Width           =   6465
         Begin VB.CommandButton cmdAttachmentsPath 
            Caption         =   "..."
            Height          =   315
            Left            =   5685
            TabIndex        =   19
            ToolTipText     =   "Select Path"
            Top             =   700
            Width           =   330
         End
         Begin VB.TextBox txtAttachmentsPath 
            BackColor       =   &H8000000F&
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   1875
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   700
            Width           =   3810
         End
         Begin VB.CommandButton cmdAttachmentsPathClear 
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   20.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5985
            MaskColor       =   &H000000FF&
            TabIndex        =   20
            ToolTipText     =   "Clear Path"
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboEmailDateFormat 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":04B8
            Left            =   1875
            List            =   "frmConfiguration.frx":04BA
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label lblAttachmentsPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attachments Path :"
            Height          =   195
            Left            =   195
            TabIndex        =   17
            Top             =   765
            Width           =   1710
         End
         Begin VB.Label lblDateFormat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Format :"
            Height          =   195
            Left            =   195
            TabIndex        =   15
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame fraTestSQLMail 
         Caption         =   "Test :"
         Height          =   1190
         Left            =   150
         TabIndex        =   10
         Top             =   3000
         Width           =   6465
         Begin VB.CheckBox chkTestEmail 
            Caption         =   "Te&st server email during Data Manager logon"
            Height          =   195
            Left            =   200
            TabIndex        =   11
            Top             =   300
            Value           =   1  'Checked
            Width           =   5000
         End
         Begin VB.TextBox txtTestEmailAddr 
            Height          =   315
            Left            =   1875
            TabIndex        =   13
            Text            =   "hrpro@hrpro.com"
            Top             =   660
            Width           =   4410
         End
         Begin VB.Label lblTestEmailAddr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address :"
            Height          =   195
            Left            =   495
            TabIndex        =   12
            Top             =   720
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeskTopBitmapID As Long
Private mbCMGExportUseCSV As Boolean
Private mbMaximizeScreens As Boolean
Private mbRememberDBColumnsView As Boolean

' Functions to display/tile the background image
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal lDC As Long, ByVal hObject As Long) As Long

Private mblnLoading As Boolean
Private mfChanged As Boolean
Private mstrTimeSeparator As String
Private mlngDefaultEmailProfile As Long
'Private sRealPassword As String
'Private sRealConfirmPassword As String
'Private mblnEmailProfilesHadError As Boolean

Private Enum ProcessAccountStatus
  iPROCESS_ERROR = 0
  iPROCESS_HUNKYDOREY = 1
  iPROCESS_NOTPROCESSADMIN = 2
  iPROCESS_PASSWORDSNOMATCH = 3
  iPROCESS_INVALIDLOGIN = 4
  iPROCESS_NOUSERNAME = 5
End Enum

Private Function IsServer64Bit() As Boolean

  Dim rsTemp As ADODB.Recordset
  
  Set rsTemp = New ADODB.Recordset
  rsTemp.Open "SELECT dbo.udfASRIsServer64Bit()", gADOCon, adOpenForwardOnly, adLockReadOnly
  IsServer64Bit = (rsTemp.Fields(0).value = 1)
  rsTemp.Close
  Set rsTemp = Nothing

End Function

Private Sub cboBitmapLocation_Populate()

  ' Stuff the values into the location combo
  cboBitmapLocation.Clear
  AddItemToComboBox cboBitmapLocation, "Top Left", 0
  AddItemToComboBox cboBitmapLocation, "Top Right", 1
  AddItemToComboBox cboBitmapLocation, "Centre", 2
  AddItemToComboBox cboBitmapLocation, "Left Tile", 3
  AddItemToComboBox cboBitmapLocation, "Right Tile", 4
  AddItemToComboBox cboBitmapLocation, "Top Tile", 5
  AddItemToComboBox cboBitmapLocation, "Bottom Tile", 6
  AddItemToComboBox cboBitmapLocation, "Tile", 7

End Sub

Private Sub cboBitmapLocation_Click()
If Not mblnLoading Then Changed = True
End Sub

Private Sub cboColours_Click()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub cboEmailDateFormat_Click()
  If (Not mblnLoading) Then Changed = True
End Sub

Private Sub cboEmailMethod_Click()
  
  Dim lngMethod As Long
  
  With cboEmailMethod
    If .ListIndex >= 0 Then
      lngMethod = .ItemData(.ListIndex)
    End If
  End With
  
  cboEmailProfile.Enabled = (lngMethod = 2)
  cboEmailProfile.BackColor = IIf(lngMethod = 2, vbWindowBackground, vbButtonFace)
  If lngMethod <> 2 Then
    cboEmailProfile.ListIndex = -1
  End If

  txtEmailServer.Enabled = (lngMethod = 3)
  txtEmailServer.BackColor = IIf(lngMethod = 3, vbWindowBackground, vbButtonFace)
  txtEmailAccount.Enabled = (lngMethod = 3)
  txtEmailAccount.BackColor = IIf(lngMethod = 3, vbWindowBackground, vbButtonFace)


  cmdEmailTest.Enabled = (lngMethod > 0)
  chkTestEmail.Enabled = (lngMethod > 0)
  
  If lngMethod = 0 Then
    chkTestEmail.value = False
  End If

  Select Case lngMethod
  Case 0, 1
    txtEmailServer.Text = vbNullString
    txtEmailAccount.Text = vbNullString
  
  Case 2  'Database Mail
    If glngSQLVersion >= 9 Then
      cboEmailProfile_Populate
    End If
  Case 3  'Thorpe Software
    txtEmailServer.Text = gstrEmailServer
    txtEmailAccount.Text = gstrEmailAccount
  End Select

  If (Not mblnLoading) And cboEmailMethod.DataChanged Then Changed = True
  
End Sub

Private Function cboEmailProfile_Populate()

  Const strDEFAULTPROFILE = "<Use Default Profile>"

  Dim rsProfiles As ADODB.Recordset
  Dim rsProfileAccount As ADODB.Recordset
  Dim sSQL As String
  Dim lngNewIndex As Long
  
  On Error GoTo LocalErr
  
  mlngDefaultEmailProfile = 0

  With cboEmailProfile
    .Clear
    AddItemToComboBox cboEmailProfile, strDEFAULTPROFILE, 0

    sSQL = "exec msdb..sysmail_help_principalprofile_sp"
    Set rsProfiles = New ADODB.Recordset
    rsProfiles.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    Do While Not rsProfiles.EOF

      lngNewIndex = AddItemToComboBox(cboEmailProfile, rsProfiles!profile_name, rsProfiles!profile_id)
      
      If rsProfiles!profile_name = gstrEmailProfile Then
        .ListIndex = lngNewIndex
      End If

      If rsProfiles!is_default Then
        mlngDefaultEmailProfile = lngNewIndex
      End If

      rsProfiles.MoveNext
    Loop

    rsProfiles.Close

    If mlngDefaultEmailProfile = 0 And .ListCount > 1 Then
      .RemoveItem 0
    End If

    If .ListIndex < 0 And .ListCount > 0 Then
      .ListIndex = 0
      If mlngDefaultEmailProfile = 0 Then
        SetComboItem cboEmailProfile, mlngDefaultEmailProfile
      End If
    End If

  End With

TidyAndExit:
  If Not rsProfiles Is Nothing Then
    If rsProfiles.State = adStateOpen Then
      rsProfiles.Close
    End If
    Set rsProfiles = Nothing
  End If

Exit Function

LocalErr:
  With cboEmailProfile
    If gstrEmailProfile = vbNullString Then
      .ListIndex = 0
    Else
      .AddItem gstrEmailProfile
      .ListIndex = lngNewIndex
    End If
    .Enabled = False
    .BackColor = vbButtonFace
  End With
  
  Resume TidyAndExit

End Function

Private Sub cboEmailProfile_Click()
        
  Dim rsProfileAccount As ADODB.Recordset
  Dim rsAccount As ADODB.Recordset
  Dim sSQL As String
  Dim lngProfileAccount As Long
  Dim lngProfileID As Long


  On Local Error GoTo LocalErr

  txtEmailServer.Text = vbNullString
  txtEmailAccount.Text = vbNullString

  With cboEmailProfile
    lngProfileID = mlngDefaultEmailProfile
    If .ListIndex >= 0 Then
      If .ItemData(.ListIndex) > 0 Then
        lngProfileID = .ItemData(.ListIndex)
      End If
    End If
        
    If lngProfileID > 0 Then
      sSQL = "exec msdb..sysmail_help_profileaccount_sp " & CStr(lngProfileID)
      Set rsProfileAccount = New ADODB.Recordset
      rsProfileAccount.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      If Not rsProfileAccount.EOF Then
        lngProfileAccount = rsProfileAccount!account_id
      End If
      rsProfileAccount.Close
      Set rsProfileAccount = Nothing
    End If

    If lngProfileAccount > 0 Then
      sSQL = "exec msdb..sysmail_help_account_sp " & CStr(lngProfileAccount)
      Set rsAccount = New ADODB.Recordset
      rsAccount.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      If Not rsAccount.EOF Then
        txtEmailServer.Text = rsAccount!ServerName
        txtEmailAccount.Text = rsAccount!Email_Address
      End If
      rsAccount.Close
      Set rsAccount = Nothing
    End If

  End With

  If (Not mblnLoading) Then Changed = True
  
Exit Sub

LocalErr:
  If Not rsProfileAccount Is Nothing Then
    If rsProfileAccount.State = adStateOpen Then
      rsProfileAccount.Close
    End If
    Set rsProfileAccount = Nothing
  End If

  If Not rsAccount Is Nothing Then
    If rsAccount.State = adStateOpen Then
      rsAccount.Close
    End If
    Set rsAccount = Nothing
  End If

  'txtEmailStatus.Text = "Error: " & Err.Description

End Sub

Private Sub cboNodeSize_Click()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub cboProcessMethod_Click()

  Dim mbEnable As Boolean
  
  mbEnable = IIf(cboProcessMethod.ListIndex = iPROCESSADMIN_SQLACCOUNT, True, False)
  ControlsDisableAll txtLogin, mbEnable
  ControlsDisableAll txtPassword, mbEnable
  ControlsDisableAll txtConfirmPassword, mbEnable
  
  ControlsDisableAll cmdTestLogon, mbEnable

  ' We're not important enough to be able to do this...
  If Not gbIsUserSystemAdmin Then
    lblProcessAdminWarning.Visible = cboProcessMethod.ListIndex = iPROCESSADMIN_EVERYONE
  End If

  ' Clear out existing settings
  If cboProcessMethod.ListIndex <> iPROCESSADMIN_SQLACCOUNT Then
    txtLogin.Text = ""
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""
  End If


  ' Enable/Disable test
  cmdTestLogon.Enabled = (cboProcessMethod.ListIndex = iPROCESSADMIN_SERVICEACCOUNT _
    Or cboProcessMethod.ListIndex = iPROCESSADMIN_SQLACCOUNT)

  If Not mblnLoading Then Changed = True

End Sub

Private Sub chkAllowAFDEvaluation_Click()
If Not mblnLoading Then Changed = True
End Sub

Private Sub chkAllowQAddressEvaluation_Click()
If Not mblnLoading Then Changed = True
End Sub

Private Sub chkDisableSpecialFunctionAutoUpdate_Click()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub chkMaximize_Click()

  mbMaximizeScreens = IIf(chkMaximize.value = vbChecked, True, False)
  If Not mblnLoading Then Changed = True
End Sub

Private Sub chkRecursionLevelManual_Click()

  Dim mbEnable As Boolean
  
  mbEnable = IIf(chkRecursionLevelManual.value = vbChecked, True, False)
  spnTriggelLevel.value = IIf(mbEnable, giDefaultRecursionLevel, giManualRecursionLevel)
  ControlsDisableAll spnTriggelLevel, mbEnable
  If Not mblnLoading Then Changed = True
End Sub

Private Sub chkRememberDBColumns_Click()

  mbRememberDBColumnsView = IIf(chkRememberDBColumns.value = vbChecked, True, False)
  If Not mblnLoading Then Changed = True
End Sub

Private Sub chkReorganiseIndexes_Click()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub chkTestEmail_Click()

  Dim blnTestEmail As Boolean

  blnTestEmail = (chkTestEmail.value = vbChecked)

  txtTestEmailAddr.Enabled = blnTestEmail
  txtTestEmailAddr.BackColor = IIf(blnTestEmail, vbWindowBackground, vbButtonFace)
  lblTestEmailAddr.Enabled = blnTestEmail
  If blnTestEmail = False Then
    txtTestEmailAddr.Text = vbNullString
  End If
  If (Not mblnLoading) And (chkTestEmail.DataChanged) Then Changed = True
End Sub


Private Sub chkUpdateStatsOvernight_Click()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub cmdAttachmentsPathClear_Click()

  If MsgBox("Are you sure you want to clear the attachment path?", vbQuestion + vbYesNoCancel, "Email Options") = vbYes Then
    txtAttachmentsPath.Text = vbNullString
    Changed = True
  End If
  cmdAttachmentsPathClear.Enabled = (txtAttachmentsPath.Text <> vbNullString)

End Sub

Private Sub cmdCancel_Click()
'UnLoad Me
Dim pintAnswer As Integer
If Changed = True Or cmdOK.Enabled Then
  pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
  If pintAnswer = vbYes Then
    'AE20071108 Fault #12551
    'Using Me.MousePointer = vbNormal forces the form to be reloaded
    'after its been unloaded in cmdOK_Click, changed to Screen.MousePointer
    'Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    cmdOK_Click 'This is just like saving
    Screen.MousePointer = vbDefault
    'Me.MousePointer = vbNormal
    Exit Sub
  ElseIf pintAnswer = vbCancel Then
    Exit Sub
  End If
End If
TidyUpAndExit:
  UnLoad Me
End Sub

Private Sub cmdColourPicker_Click()

  On Error GoTo ErrorTrap

'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = glngDeskTopColour
'    .ShowColor
'    glngDeskTopColour = .Color
'  End With
  
  ' AE20080331 Fault #13052
  With ColorPicker
    .Color = glngDeskTopColour
    .ShowPalette
    glngDeskTopColour = .Color
  End With

  If (lblBackColour.BackColor <> glngDeskTopColour) And (Not mblnLoading) Then Changed = True

  ' Display the selected colour
  lblBackColour.BackColor = glngDeskTopColour

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  ' User pressed cancel.
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub cmdDefaultScreenFont_Click()

  On Error GoTo ErrorTrap
  
  With comDlgBox
    .FontName = gobjDefaultScreenFont.Name
    .FontSize = gobjDefaultScreenFont.Size
    .FontBold = gobjDefaultScreenFont.Bold
    .FontItalic = gobjDefaultScreenFont.Italic
    .FontUnderline = gobjDefaultScreenFont.Underline
    .FontStrikethru = gobjDefaultScreenFont.Strikethrough
    .Color = txtDefaultScreenFont.ForeColor
       
    .Flags = cdlCFScreenFonts Or cdlCFEffects
    .ShowFont
      
    If gobjDefaultScreenFont.Name <> .FontName _
      Or gobjDefaultScreenFont.Size <> .FontSize _
      Or gobjDefaultScreenFont.Bold <> .FontBold _
      Or gobjDefaultScreenFont.Italic <> .FontItalic _
      Or gobjDefaultScreenFont.Underline <> .FontUnderline _
      Or gobjDefaultScreenFont.Strikethrough <> .FontStrikethru _
      Or txtDefaultScreenFont.ForeColor <> .Color Then
      
      gobjDefaultScreenFont.Name = .FontName
      gobjDefaultScreenFont.Size = .FontSize
      gobjDefaultScreenFont.Bold = .FontBold
      gobjDefaultScreenFont.Italic = .FontItalic
      gobjDefaultScreenFont.Underline = .FontUnderline
      gobjDefaultScreenFont.Strikethrough = .FontStrikethru
      glngDefaultScreenForeColor = .Color
      
      txtDefaultScreenFont.Text = GetFontDescription(gobjDefaultScreenFont)
      txtDefaultScreenFont.ForeColor = glngDefaultScreenForeColor
      
      Changed = True

    End If
  End With

ErrorTrap:
  Err = False

End Sub

Private Sub cmdOK_Click()

  Dim lngMethod As Long
  Dim iTestProcess As ProcessAccountStatus
  
  lngMethod = cboEmailMethod.ItemData(cboEmailMethod.ListIndex)
  If Not ValidateEmailMethod(lngMethod) Then
    Exit Sub
  End If


  If Me.chkTestEmail.value = vbChecked And txtTestEmailAddr.Text = vbNullString Then
    SSTab1.Tab = 0
    MsgBox "No test email address entered.", vbExclamation, "Configuration"
    txtTestEmailAddr.SetFocus
    Exit Sub
  End If

  If Not IsValidTime(TDBAMStartTime.Text) Then
    SSTab1.Tab = 5
    MsgBox "Invalid AM start time entered.", vbExclamation, Me.Caption
    TDBAMStartTime.SetFocus
    Exit Sub
  End If

  If Not IsValidTime(TDBAMEndTime.Text) Then
    SSTab1.Tab = 5
    MsgBox "Invalid AM end time entered.", vbExclamation, Me.Caption
    TDBAMEndTime.SetFocus
    Exit Sub
  End If
  
  If Not IsValidTime(TDBPMStartTime.Text) Then
    SSTab1.Tab = 5
    MsgBox "Invalid PM start time entered.", vbExclamation, Me.Caption
    TDBPMStartTime.SetFocus
    Exit Sub
  End If
  
  If Not IsValidTime(TDBPMEndTime.Text) Then
    SSTab1.Tab = 5
    MsgBox "Invalid PM end time entered.", vbExclamation, Me.Caption
    TDBPMEndTime.SetFocus
    Exit Sub
  End If
   
   
  ' Process Account stuff
  If glngSQLVersion >= 9 Then
    iTestProcess = TestProcessLogon(False)
    If iTestProcess <> iPROCESS_HUNKYDOREY Then
      SSTab1.Tab = 1
      
      If iTestProcess = iPROCESS_NOUSERNAME Then
        txtLogin.SetFocus
      End If
      
      If iTestProcess = iPROCESS_PASSWORDSNOMATCH Then
        txtPassword.SetFocus
      End If
      
      Exit Sub
    End If
  End If
  
   
  'MH20080118 Fault 12777 - If Database Mail and can't read profiles then don't save config.
  ' Save the email settings
  With cboEmailMethod
    'If .ItemData(.ListIndex) <> 2 Or Not mblnEmailProfilesHadError Then
      If .ListIndex >= 0 Then
        glngEmailMethod = .ItemData(.ListIndex)
      Else
        glngEmailMethod = 0
      End If
      gstrEmailProfile = cboEmailProfile.Text
      gstrEmailServer = txtEmailServer.Text
      gstrEmailAccount = txtEmailAccount.Text
    'End If
  End With
    


  glngEmailDateFormat = cboEmailDateFormat.ItemData(cboEmailDateFormat.ListIndex)
  gstrEmailAttachmentPath = txtAttachmentsPath.Text
  gstrEmailTestAddr = txtTestEmailAddr.Text
'  gstrEmailEventLogToAddr = Trim(txtSendEventLogTo.Text)
  
  ' Save the desktop bitmap settings
  glngDesktopBitmapID = mlngDeskTopBitmapID
  glngDesktopBitmapLocation = cboBitmapLocation.ItemData(cboBitmapLocation.ListIndex)
  glngDeskTopColour = lblBackColour.BackColor

  'Apply the changes
  frmSysMgr.SetBackground (False)
  
  ' Expression defaults
  glngExpressionViewColours = cboColours.ItemData(cboColours.ListIndex)
  glngExpressionViewNodes = cboNodeSize.ItemData(cboNodeSize.ListIndex)
  
  gbMaximizeScreens = mbMaximizeScreens
  gbRememberDBColumnsView = mbRememberDBColumnsView
   
  ' Advanced Settings
  gbManualRecursionLevel = chkRecursionLevelManual.value
  giManualRecursionLevel = spnTriggelLevel.value
  gbDisableSpecialFunctionAutoUpdate = chkDisableSpecialFunctionAutoUpdate.value
  gbReorganiseIndexesInOvernightJob = chkReorganiseIndexes.value
  
  ' Save development options (No need to run save code for this work. Saves time eh?)
'  SaveSystemSetting "Development", "EventLog_Email_Enable", chkAllowEmailToDevelopers.Value
'  SaveSystemSetting "Development", "EventLog_Email_1", txtEventLogEmail(0).Text
'  SaveSystemSetting "Development", "EventLog_Email_2", txtEventLogEmail(1).Text
'  SaveSystemSetting "Development", "EventLog_Email_3", txtEventLogEmail(2).Text
'  SaveSystemSetting "Development", "EventLog_Email_4", txtEventLogEmail(3).Text
  SaveSystemSetting "Development", "AFD_Evaluation_Enable", chkAllowAFDEvaluation.value
  SaveSystemSetting "Development", "AFD_Evaluation_Seed_Normal", txtDeveloperAFDNormal.Text
  SaveSystemSetting "Development", "AFD_Evaluation_Seed_Plus", txtDeveloperAFDPlus.Text
  SaveSystemSetting "Development", "AFD_Evaluation_Seed_NN", txtDeveloperAFDNamesNumbers.Text
  
  SaveSystemSetting "Development", "QAddress_Evaluation_Enable", chkAllowQAddressEvaluation.value
  SaveSystemSetting "Development", "QAddress_Evaluation_Seed", txtQAddressSeedValue.Text
  
  glngOvernightJobTime = ConvertStringToTime(TDBMaskTime.Text)
  
  
  'MH20040301
  glngAMStartTime = ConvertStringToTime(TDBAMStartTime.Text)
  glngAMEndTime = ConvertStringToTime(TDBAMEndTime.Text)
  glngPMStartTime = ConvertStringToTime(TDBPMStartTime.Text)
  glngPMEndTime = ConvertStringToTime(TDBPMEndTime.Text)


  ' Save SQL process admin settings
  With cboProcessMethod
    If .ListIndex >= 0 Then
      glngProcessMethod = .ItemData(.ListIndex)
    Else
      glngProcessMethod = iPROCESSADMIN_DISABLED
    End If
  End With


  ' SQL 2005 Process account
  With recModuleSetup
    .Index = "idxModuleParameter"
      
    ' Save the Login name.
    .Seek "=", gsMODULEKEY_SQL, gsPARAMETERKEY_LOGINDETAILS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_SQL
      !parameterkey = gsPARAMETERKEY_LOGINDETAILS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_ENCYPTED
    
    If cboProcessMethod.ListIndex = iPROCESSADMIN_SQLACCOUNT Or cboProcessMethod.ListIndex = iPROCESSADMIN_SERVICEACCOUNT Then
      !parametervalue = EncryptLogonDetails(txtLogin.Text, txtPassword.Tag, gsDatabaseName, gsServerName)
    Else
      !parametervalue = ""
    End If
    
    .Update
  End With

  ' Web Information
  gstrWebSiteAddress = txtWebSiteAddress.Text

  ' Done
  Application.Changed = True
  UnLoad Me

End Sub

Private Function ConvertStringToTime(pstrTimeString As String) As Long

  ConvertStringToTime = CLng(Replace(pstrTimeString, mstrTimeSeparator, ""))
  
End Function

Private Function ConvertTimeToString(plngTime As Long, lngLen As Long) As String
  
  Dim strTemp As String
  
  strTemp = CStr(plngTime)
  strTemp = String((lngLen - Len(strTemp)), "0") + strTemp
  
  strTemp = Left(strTemp, 2) & mstrTimeSeparator & Mid(strTemp, 3, 2) & mstrTimeSeparator & Mid(strTemp, 5, 2)
  
  ConvertTimeToString = strTemp

End Function

Private Function ValidateTime(pstrTime As String) As Boolean

  Dim strTemp As String
  
  strTemp = pstrTime
  
  ValidateTime = True
  
  If (CInt(Left(strTemp, 2)) > 23) _
    Or (CInt(Mid(strTemp, 4, 2)) > 59) _
    Or (CInt(Mid(strTemp, 7, 2)) > 59) Then
    ValidateTime = False
    Exit Function
  End If
    
End Function

Private Sub cmdPictureClear_Click()

  mlngDeskTopBitmapID = 0
  txtDeskTopBitmapName.Text = ""

  cboBitmapLocation.Enabled = False
  cmdPictureClear.Enabled = False
  picHolder.Visible = False

  Changed = True
End Sub

Private Sub cmdPictureSelect_Click()

  Dim lngPictureID As Long
  Dim sFileName As String

  frmPictSel.PictureType = vbPicTypeBitmap
  frmPictSel.SelectedPicture = mlngDeskTopBitmapID
  frmPictSel.ExcludedExtensions = ".gif"
  frmPictSel.Show vbModal

If (frmPictSel.SelectedPicture <> mlngDeskTopBitmapID) And (Not mblnLoading) Then Changed = True

  If frmPictSel.SelectedPicture > 0 Then
    With recPictEdit
      .Index = "idxID"
      .Seek "=", frmPictSel.SelectedPicture
      If Not .NoMatch Then
        mlngDeskTopBitmapID = !PictureID
        txtDeskTopBitmapName.Text = !Name
        cboBitmapLocation.Enabled = True
        cmdPictureClear.Enabled = True
        sFileName = ReadPicture
        picWork.Picture = LoadPicture(sFileName)
        picWork.Move 0, 0, picHolder.ScaleWidth, picHolder.ScaleHeight
        SizeImage picWork
        picWork.Top = (picHolder.ScaleHeight - picWork.Height) \ 2
        picWork.Left = (picHolder.ScaleWidth - picWork.Width) \ 2
        Kill sFileName
        picHolder.Visible = True
      Else
        cboBitmapLocation.Enabled = False
        cmdPictureClear.Enabled = False
        picHolder.Visible = False
      End If
    End With

  End If

End Sub

Private Sub cmdEmailTest_Click()

  Dim frmTestEmail As frmConfigurationTestEmail
  Dim lngMethod As Long
  
  lngMethod = cboEmailMethod.ItemData(cboEmailMethod.ListIndex)
  
  If ValidateEmailMethod(lngMethod) Then
    Set frmTestEmail = New frmConfigurationTestEmail
    frmTestEmail.Initialise lngMethod, cboEmailProfile.Text, txtEmailServer.Text, txtEmailAccount.Text
    frmTestEmail.Show vbModal
    Set frmTestEmail = Nothing
  End If

End Sub

Private Function ValidateEmailMethod(lngMethod As Long) As Boolean

  ValidateEmailMethod = True

  If lngMethod = 3 Then
    
    If Trim(txtEmailServer.Text) = vbNullString Then
      SSTab1.Tab = 0
      MsgBox "Please enter the name or IP address of your mail server.", vbCritical, Me.Caption
      txtEmailServer.SetFocus
      ValidateEmailMethod = False
    
    ElseIf Trim(txtEmailAccount.Text) = vbNullString Then
      SSTab1.Tab = 0
      MsgBox "Please enter the name of the email account allocated to the SQL server.", vbCritical, Me.Caption
      txtEmailAccount.SetFocus
      ValidateEmailMethod = False
    
    End If
  
  End If

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()

  Dim blnReadonly As Boolean
  Dim blnIsSAUser As Boolean
  Dim sFileName As String
  Dim strTimeFormat As String

  mblnLoading = True
  Changed = False
  
  mstrTimeSeparator = UI.GetSystemTimeSeparator
  strTimeFormat = "99" & mstrTimeSeparator & "99"
  TDBAMStartTime.Format = strTimeFormat
  TDBAMEndTime.Format = strTimeFormat
  TDBPMStartTime.Format = strTimeFormat
  TDBPMEndTime.Format = strTimeFormat

  strTimeFormat = "99" & mstrTimeSeparator & "99" & mstrTimeSeparator & "99"
  TDBMaskTime.Format = strTimeFormat

  ' AE20080327 Fault #13014
  chkTestEmail.value = IIf(gstrEmailTestAddr <> vbNullString, vbChecked, vbUnchecked)
  txtTestEmailAddr.Text = gstrEmailTestAddr
  
  cboEmailMethod_Populate
  
  cboEmailDateFormat_Populate
  SetComboItem cboEmailDateFormat, glngEmailDateFormat
  txtAttachmentsPath.Text = gstrEmailAttachmentPath
  picHolder.BorderStyle = 0

  ' AE20080327 Fault #13014
'  chkTestEmail.Value = IIf(gstrEmailTestAddr <> vbNullString, vbChecked, vbUnchecked)
'  txtTestEmailAddr.Text = gstrEmailTestAddr

'  ' Default email event log to
'  txtSendEventLogTo.Text = gstrEmailEventLogToAddr

  'Desktop Bitmap
  With recPictEdit
    .Index = "idxID"
    .Seek "=", glngDesktopBitmapID
    If Not .NoMatch Then
      mlngDeskTopBitmapID = !PictureID
      txtDeskTopBitmapName.Text = !Name
      cboBitmapLocation.Enabled = True
      cmdPictureClear.Enabled = True
      sFileName = ReadPicture
      picWork.Picture = LoadPicture(sFileName)
      picWork.Move 0, 0, picHolder.ScaleWidth, picHolder.ScaleHeight
      SizeImage picWork
      picWork.Top = (picHolder.ScaleHeight - picWork.Height) \ 2
      picWork.Left = (picHolder.ScaleWidth - picWork.Width) \ 2
      Kill sFileName
    Else
      cboBitmapLocation.Enabled = False
      cmdPictureClear.Enabled = False
      picHolder.Visible = False
    End If
  End With

  ' Bitmap Screen Location
  cboBitmapLocation_Populate
  SetComboItem cboBitmapLocation, glngDesktopBitmapLocation

  ' Display the selected colour
  lblBackColour.BackColor = glngDeskTopColour
  
  ' Load the default expression types
  cboExpressions_Populate
  SetComboItem cboColours, glngExpressionViewColours
  SetComboItem cboNodeSize, glngExpressionViewNodes
  
  blnReadonly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)
  
  ' AE20090624 Fault HRPRO-55
  'blnIsSAUser = gbCurrentUserIsSysSecMgr
  blnIsSAUser = gbIsUserSystemAdmin
  
  ' Moved to seperate module setup screen
  If glngSQLVersion < 9 Then
    SSTab1.TabVisible(1) = False
  Else
    cboProcessMethod_Populate
    If glngProcessMethod = iPROCESSADMIN_SQLACCOUNT Then
      LoadProcessAccount
    Else
      SetComboItem cboProcessMethod, glngProcessMethod
    End If
  End If
  
  'Development options
  SSTab1.TabVisible(3) = ASRDEVELOPMENT
  'SSTab1.TabVisible(6) = ASRDEVELOPMENT
  chkRecursionLevelManual.Visible = ASRDEVELOPMENT
  spnTriggelLevel.Visible = ASRDEVELOPMENT
  
'  chkAllowEmailToDevelopers.Value = GetSystemSetting("Development", "EventLog_Email_Enable", False)
'  txtEventLogEmail(0).Text = GetSystemSetting("Development", "EventLog_Email_1", "")
'  txtEventLogEmail(1).Text = GetSystemSetting("Development", "EventLog_Email_2", "")
'  txtEventLogEmail(2).Text = GetSystemSetting("Development", "EventLog_Email_3", "")
'  txtEventLogEmail(3).Text = GetSystemSetting("Development", "EventLog_Email_4", "")

  ' AFD Settings
  chkAllowAFDEvaluation.value = GetSystemSetting("Development", "AFD_Evaluation_Enable", False)
  txtDeveloperAFDNormal.Text = GetSystemSetting("Development", "AFD_Evaluation_Seed_Normal", "B13 9JD")
  txtDeveloperAFDPlus.Text = GetSystemSetting("Development", "AFD_Evaluation_Seed_Plus", "B45 9AA")
  txtDeveloperAFDNamesNumbers.Text = GetSystemSetting("Development", "AFD_Evaluation_Seed_NN", "IS2 9NY")
  
  ' Quick Address Settings
  chkAllowQAddressEvaluation.value = GetSystemSetting("Development", "QAddress_Evaluation_Enable", False)
  txtQAddressSeedValue.Text = GetSystemSetting("Development", "QAddress_Evaluation_Seed", "AL1 5ST")

  'Overnight Job Schedule Tab
  fraTime.Enabled = ((Not blnReadonly) And (blnIsSAUser))
  lblOccurs.Enabled = ((Not blnReadonly) And (blnIsSAUser))
  TDBMaskTime.Enabled = ((Not blnReadonly) And (blnIsSAUser))
  TDBMaskTime.Text = ConvertTimeToString(glngOvernightJobTime, 6)
  chkReorganiseIndexes.value = IIf(gbReorganiseIndexesInOvernightJob, vbChecked, vbUnchecked)
  
  'MH20040301
  TDBAMStartTime.Text = ConvertTimeToString(glngAMStartTime, 4)
  TDBAMEndTime.Text = ConvertTimeToString(glngAMEndTime, 4)
  TDBPMStartTime.Text = ConvertTimeToString(glngPMStartTime, 4)
  TDBPMEndTime.Text = ConvertTimeToString(glngPMEndTime, 4)

  ' Load general display settings
  mbMaximizeScreens = gbMaximizeScreens
  chkMaximize.value = IIf(mbMaximizeScreens, vbChecked, vbUnchecked)

  mbRememberDBColumnsView = gbRememberDBColumnsView
  chkRememberDBColumns.value = IIf(mbRememberDBColumnsView, vbChecked, vbUnchecked)

  ' Advanced Settings
  chkRecursionLevelManual.value = IIf(gbManualRecursionLevel, vbChecked, vbUnchecked)
  spnTriggelLevel.value = IIf(gbManualRecursionLevel, giManualRecursionLevel, giDefaultRecursionLevel)
  ControlsDisableAll spnTriggelLevel, gbManualRecursionLevel
  
  chkDisableSpecialFunctionAutoUpdate.value = IIf(gbDisableSpecialFunctionAutoUpdate, vbChecked, vbUnchecked)
  
  'Read Only settings
  If blnReadonly Then
    ControlsDisableAll Me
  Else
    'cmdAttachmentsPathClear.Enabled = ((txtAttachmentsPath.Text <> vbNullString) And (blnIsSAUser))
    'cmdAttachmentsPath.Enabled = (blnIsSAUser)
    ' AE20090624 Fault HRPRO-55
    cmdAttachmentsPathClear.Enabled = ((txtAttachmentsPath.Text <> vbNullString) And (gbCurrentUserIsSysSecMgr))
    cmdAttachmentsPath.Enabled = (gbCurrentUserIsSysSecMgr)
  End If
  
  txtDefaultScreenFont.Text = GetFontDescription(gobjDefaultScreenFont)
  txtDefaultScreenFont.ForeColor = glngDefaultScreenForeColor
  
  ' Web Information
  txtWebSiteAddress.Text = gstrWebSiteAddress
  
  'Disable other non used frames to stop elements being tabbed to.
  SSTab1_Click (1)
  'Set to the first tab - handy when you forget to do it at design time
  SSTab1.Tab = 0
  
  mblnLoading = False
End Sub

Private Sub cboEmailDateFormat_Populate()

  cboEmailDateFormat.Clear
  AddItemToComboBox cboEmailDateFormat, "dd/mm/yyyy", 103
  AddItemToComboBox cboEmailDateFormat, "dd-mm-yyyy", 105
  AddItemToComboBox cboEmailDateFormat, "dd.mm.yyyy", 104
  AddItemToComboBox cboEmailDateFormat, "dd mon yyyy", 106
  AddItemToComboBox cboEmailDateFormat, "mm/dd/yyyy", 101
  AddItemToComboBox cboEmailDateFormat, "mm-dd-yyyy", 110
  AddItemToComboBox cboEmailDateFormat, "mon dd, yyyy", 107
  AddItemToComboBox cboEmailDateFormat, "yyyy/mm/dd", 111
  AddItemToComboBox cboEmailDateFormat, "yyyy.mm.dd", 102
  AddItemToComboBox cboEmailDateFormat, "yyyymmdd", 112

End Sub

Private Sub cboEmailMethod_Populate()

  cboEmailMethod.Clear
  AddItemToComboBox cboEmailMethod, "<Disable Emails>", 0
  
  If Not IsServer64Bit And glngSQLVersion < 11 Then
    AddItemToComboBox cboEmailMethod, "SQL Mail", 1
  End If

  If glngSQLVersion >= 9 Then
    AddItemToComboBox cboEmailMethod, "Database Mail", 2
  End If

  If ProcedureExists("master", "xp_SMTPSendMail80") Or glngEmailMethod = 3 Then
    AddItemToComboBox cboEmailMethod, "Thorpe Software (xp_SMTPSendMail80)", 3
  End If

  SetComboItem cboEmailMethod, glngEmailMethod

End Sub

Private Sub cboProcessMethod_Populate()

  cboProcessMethod.Clear
  AddItemToComboBox cboProcessMethod, "Public has been granted view server state", 0
  AddItemToComboBox cboProcessMethod, "Trusted SQL Service Account", 1
  AddItemToComboBox cboProcessMethod, "Specified SQL Account", 2
  SetComboItem cboProcessMethod, glngProcessMethod

End Sub

Private Function ProcedureExists(strDatabase As String, strName As String) As Boolean

  Dim rsProcs As ADODB.Recordset
  Dim sSQL As String

  sSQL = "SELECT Name FROM [" & strDatabase & "].[dbo].[sysObjects] WHERE name = '" & strName & "'"
  
  Set rsProcs = New ADODB.Recordset
  rsProcs.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  ProcedureExists = (Not rsProcs.EOF)
  
  rsProcs.Close
  Set rsProcs = Nothing

End Function

Private Sub cmdAttachmentsPath_Click()
  
  Dim frmPathSel As frmConfigurationPathSel
  
  On Error GoTo LocalErr
  
  Set frmPathSel = New frmConfigurationPathSel
  With frmPathSel
    If .Initialise Then
      .AttachmentPath = txtAttachmentsPath.Text
      .Show vbModal
      If Not .Cancelled Then
        'MH20040127 Fault 7519 ASRSysSettings can only handle 200 characters
        'If Len(.AttachmentPath) > 255 Then
        '  MsgBox "Selected path cannot be longer than 255 characters", vbExclamation, "Email Options"
        If Len(.AttachmentPath) > 200 Then
          MsgBox "Selected path cannot be longer than 200 characters", vbExclamation, "Email Options"
        Else
          
          If txtAttachmentsPath.Text <> .AttachmentPath And _
            Trim(txtAttachmentsPath.Text) <> vbNullString Then
              If MsgBox("Changing the email attachment path could mean that your existing email links are sent without attachments." & vbCrLf & _
                        "Do you wish to continue ?", vbExclamation + vbYesNo, "Email Options") = vbYes Then
                txtAttachmentsPath.Text = .AttachmentPath
                Changed = True
              End If
          Else
            txtAttachmentsPath.Text = .AttachmentPath
          End If
          
        End If
      End If
    End If
  End With
  Set frmPathSel = Nothing
  
  Dim blnReadonly As Boolean
  Dim blnIsSAUser As Boolean
  
  blnReadonly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)
  blnIsSAUser = (LCase(Trim(gsUserName)) = "sa")

  cmdAttachmentsPathClear.Enabled = ((txtAttachmentsPath.Text <> vbNullString) And (blnIsSAUser))

Exit Sub

LocalErr:
  Set frmConfigurationPathSel = Nothing
  Set frmPathSel = Nothing
  MsgBox "Error selecting Email Attachment Path", vbCritical, "Configuration"

End Sub

Private Sub cboExpressions_Populate()

  ' Colour options
  cboColours.Clear
  AddItemToComboBox cboColours, "Black", EXPRESSIONBUILDER_COLOUROFF
  AddItemToComboBox cboColours, "Colour Levels", EXPRESSIONBUILDER_COLOURON
  
  ' Node statuses
  cboNodeSize.Clear
  AddItemToComboBox cboNodeSize, "Minimized", EXPRESSIONBUILDER_NODESMINIMIZE
  AddItemToComboBox cboNodeSize, "Expand All", EXPRESSIONBUILDER_NODESEXPAND
  AddItemToComboBox cboNodeSize, "Expand Top Level", EXPRESSIONBUILDER_NODESTOPLEVEL

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub spnTriggelLevel_Click()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
      
    fraEmailSetup.Enabled = (SSTab1.Tab = 0)
    fraEmailOptions.Enabled = (SSTab1.Tab = 0)
    fraTestSQLMail.Enabled = (SSTab1.Tab = 0)
    
    fraSQL2005.Enabled = (SSTab1.Tab = 1)
    
    frmBackgrounds.Enabled = (SSTab1.Tab = 2)
    fraGeneral.Enabled = (SSTab1.Tab = 2)
    frmExpressions.Enabled = (SSTab1.Tab = 2)
    
    frmDeveloperAFD.Enabled = (SSTab1.Tab = 3)
    fraQuickAddress.Enabled = (SSTab1.Tab = 3)
    
    fraTime.Enabled = (SSTab1.Tab = 4)
    fraOutlookCalendar.Enabled = (SSTab1.Tab = 4)
    fraAdvancedSettings.Enabled = (SSTab1.Tab = 4)

End Sub

Private Sub TDBAMEndTime_Change()
If Not mblnLoading Then Changed = True
End Sub

Private Sub TDBAMStartTime_Change()
    If Not mblnLoading Then Changed = True
End Sub

Private Sub TDBAMStartTime_GotFocus()
  With TDBAMStartTime
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub TDBAMEndTime_GotFocus()
  With TDBAMEndTime
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub TDBPMEndTime_Change()
If Not mblnLoading Then Changed = True
End Sub

Private Sub TDBPMStartTime_Change()
If Not mblnLoading Then Changed = True
End Sub

Private Sub TDBPMStartTime_GotFocus()
  With TDBPMStartTime
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub TDBPMEndTime_GotFocus()
  With TDBPMEndTime
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub TDBMaskTime_GotFocus()
  With TDBMaskTime
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub


Private Sub TDBMaskTime_Change()
  If Not mblnLoading Then
    Application.ChangedOvernightJobSchedule = True
    Changed = True
  End If
End Sub

Private Sub TDBMaskTime_LostFocus()

  If Not ValidateTime(TDBMaskTime.Text) Then
    MsgBox "Invalid Time.", vbOKOnly + vbExclamation, App.Title
    TDBMaskTime.SetFocus
  End If
  
End Sub

Private Sub txtAttachmentsPath_Change()
  If (Not mblnLoading) Then Changed = True
End Sub

Private Sub txtConfirmPassword_Change()
  'sRealConfirmPassword = txtConfirmPassword.Text
  If (Not mblnLoading) Then
    Changed = True
    txtConfirmPassword.Tag = txtConfirmPassword.Text
  End If
End Sub

Private Sub txtConfirmPassword_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText
End Sub

Private Sub txtDeveloperAFDNamesNumbers_Change()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub txtDeveloperAFDNormal_Change()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub txtDeveloperAFDPlus_Change()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub txtEmailAccount_Change()
  If (Not mblnLoading) Then Changed = True
End Sub

Private Sub txtEmailServer_Change()
  If (Not mblnLoading) Then Changed = True
End Sub

Private Sub txtLogin_Change()
  If (Not mblnLoading) Then Changed = True
End Sub

Private Sub txtLogin_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText
End Sub

Private Sub txtPassword_Change()
  If (Not mblnLoading) Then
    Changed = True
    txtPassword.Tag = txtPassword.Text
  End If
End Sub

Private Sub txtPassword_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText
End Sub

Private Sub txtQAddressSeedValue_Change()
  If Not mblnLoading Then Changed = True
End Sub

Private Sub txtTestEmailAddr_Change()
  If (Not mblnLoading) Then Changed = True
End Sub

Private Sub txtTestEmailAddr_GotFocus()
  With txtTestEmailAddr
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Function IsValidTime(strInput As String) As Boolean
  IsValidTime = False
  If strInput Like "??" & mstrTimeSeparator & "??" Then
    IsValidTime = (val(Left(strInput, 2)) < 24 And val(Right(strInput, 2)) < 60)
  End If
End Function


Private Sub cmdTestLogon_Click()

  Dim iProcessStatus As ProcessAccountStatus
  iProcessStatus = TestProcessLogon(True)

End Sub

Private Function TestProcessLogon(ByRef bNotifyIfSuccessful As Boolean) As ProcessAccountStatus

  On Error GoTo ErrorTrap

  Dim iProcessStatus As ProcessAccountStatus
  Dim sConnect As String
  Dim sSQL As String
  Dim objTestConn As ADODB.Connection
  Dim rsTest As ADODB.Recordset
  Dim bOK As Boolean
  Dim strEncrypted As String
     
  ' Test the encrypted logon
  If cboProcessMethod.ListIndex = iPROCESSADMIN_DISABLED Then
    iProcessStatus = iPROCESS_HUNKYDOREY
  Else
  
    If cboProcessMethod.ListIndex = iPROCESSADMIN_SERVICEACCOUNT Then
      strEncrypted = EncryptLogonDetails("", "", gsDatabaseName, gsServerName)
    Else
    
      If Trim(txtLogin.Text) = vbNullString Then
        iProcessStatus = iPROCESS_NOUSERNAME
        GoTo TidyUpAndExit
      End If
      
      If txtPassword.Tag <> txtConfirmPassword.Tag Then
        iProcessStatus = iPROCESS_PASSWORDSNOMATCH
        GoTo TidyUpAndExit
      End If
    
      strEncrypted = EncryptLogonDetails(Replace(txtLogin.Text, ";", ""), Replace(txtPassword.Tag, ";", ""), gsDatabaseName, gsServerName)
    End If
     
    Screen.MousePointer = vbHourglass
        
    If cboProcessMethod.ListIndex = iPROCESSADMIN_SQLACCOUNT Then
    
      On Error GoTo ErrorLogin
  
      ' AE20080414 Fault #13089
'      sConnect = "Driver=SQL Server;" & _
'           "Server=" & gsServerName & ";" & _
'           "UID=" & Replace(txtLogin.Text, ";", "") & ";" & _
'           "PWD=" & Replace(txtPassword.Tag, ";", "") & ";" & _
'           "Database=" & gsDatabaseName & ";"
           
      sConnect = "Driver=SQL Server;" & _
                  "Server=" & gsServerName & ";" & _
                  "UID=" & Replace(txtLogin.Text, ";", "") & ";" & _
                  "PWD=" & Replace(txtPassword.Tag, ";", "") & ";" & _
                  "Database=" & gsDatabaseName & ";" & _
                  "Pooling=False;" & _
                  "App=Test OpenHR Config;"
                 
      Set objTestConn = New ADODB.Connection
      With objTestConn
        .ConnectionString = sConnect
        .Provider = "SQLOLEDB"
        .CommandTimeout = 10
        .ConnectionTimeout = 30
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .Properties("Packet Size") = 32767
        .Open
        .Close
      End With
       
      On Error GoTo ErrorTrap
       
    End If
       
    Set rsTest = New ADODB.Recordset
    rsTest.Open "SELECT dbo.udfASRNetIsProcessValid('" & Replace(strEncrypted, "'", "''") & "')", gADOCon, adOpenForwardOnly, adLockReadOnly
    bOK = (rsTest.Fields(0).value = True)
    rsTest.Close
  
    If bOK Then
      iProcessStatus = iPROCESS_HUNKYDOREY
    Else
      iProcessStatus = iPROCESS_NOTPROCESSADMIN
    End If
    
  End If
   
TidyUpAndExit:

  Set rsTest = Nothing
  Screen.MousePointer = vbDefault

  ' User notification
  Select Case iProcessStatus
  
    Case iPROCESS_HUNKYDOREY
      If bNotifyIfSuccessful Then
        MsgBox "Test completed successfully.", vbInformation, "Test Process Login"
      End If
    
    Case iPROCESS_NOTPROCESSADMIN
      SSTab1.Tab = 1
      MsgBox "Selected account needs to be defined as a process admin." & vbNewLine & vbNewLine _
        & "Please see your SQL administrator.", vbInformation, "Test Process Login"
    
    Case iPROCESS_PASSWORDSNOMATCH
      SSTab1.Tab = 1
      MsgBox "The passwords you typed do not match. Type the correct password in both textboxes.", vbInformation, "Test Login"
    
    Case iPROCESS_INVALIDLOGIN
      SSTab1.Tab = 1
      MsgBox "Unable to login using the credentials you supplied." & vbNewLine & vbNewLine _
        & "Please ensure that the username and password are correct.", vbInformation, "Test Process Login"
    
    Case iPROCESS_NOUSERNAME
      SSTab1.Tab = 1
      MsgBox "You must enter a SQL login.", vbInformation, "Test Process Login"
    
    Case iPROCESS_ERROR
      SSTab1.Tab = 1
      MsgBox "Unable to test process administrator login." & vbNewLine & vbNewLine _
        & "Please see your SQL administrator.", vbInformation, "Test Process Login"
    
  End Select

  Set objTestConn = Nothing

  TestProcessLogon = iProcessStatus
  Exit Function
  
ErrorLogin:
  iProcessStatus = iPROCESS_INVALIDLOGIN
  GoTo TidyUpAndExit
  
ErrorTrap:
  iProcessStatus = iPROCESS_ERROR
  GoTo TidyUpAndExit

End Function


Private Sub LoadProcessAccount()

  Dim strEncrypted As String
  Dim sUserName As String
  Dim sPassword As String
  Dim sDatabase As String
  Dim sServer As String

  ' Load process login info
  With recModuleSetup
    .Index = "idxModuleParameter"
      
    ' Get the Login Name.
    .Seek "=", gsMODULEKEY_SQL, gsPARAMETERKEY_LOGINDETAILS
    If .NoMatch Then
      ' Get the Personnel module Personnel table ID.
      .Seek "=", gsMODULEKEY_SQL, gsPARAMETERKEY_LOGINDETAILS
      If .NoMatch Then
        strEncrypted = ""
      Else
        strEncrypted = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
      End If
    Else
      strEncrypted = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    DecryptLogonDetails strEncrypted, sUserName, sPassword, sDatabase, sServer
    
    If sUserName = "" Then
      EnableControl txtLogin, False
      EnableControl txtPassword, False
      EnableControl txtConfirmPassword, False
      EnableControl cmdTestLogon, False
    Else
      txtLogin.Text = sUserName
      txtPassword.Tag = sPassword
      txtConfirmPassword.Tag = sPassword
      'sRealPassword = sPassword
      'sRealConfirmPassword = sPassword
'      txtPassword.Text = sPassword
'      txtConfirmPassword.Text = sPassword
      'NHRD11122006 Fault 11720
      txtPassword.Text = Space(20)
      txtConfirmPassword.Text = Space(20)
    End If
  End With

End Sub

Private Function ADOConError(objTestConn As ADODB.Connection) As String

  Dim strErrorDesc As String
  Dim lngCount As Long

  strErrorDesc = vbNullString
  If Not objTestConn Is Nothing Then
    If Not objTestConn.Errors Is Nothing Then
      For lngCount = 0 To objTestConn.Errors.Count - 1
        strErrorDesc = objTestConn.Errors(lngCount).Description
      Next
      strErrorDesc = Mid(strErrorDesc, InStrRev(strErrorDesc, "]") + 1)
    End If
  End If

  ADOConError = strErrorDesc

End Function

Private Property Get Changed() As Boolean
  mfChanged = Changed
End Property

Private Property Let Changed(ByVal fNewValue As Boolean)
  mfChanged = fNewValue
  cmdOK.Enabled = mfChanged
End Property

Private Sub txtWebSiteAddress_Change()
  If Not mblnLoading Then Changed = True
End Sub
