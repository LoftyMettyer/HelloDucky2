VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Object = "{F563792C-3E4B-4D13-A0C5-81DA6B7B314B}#1.0#0"; "COA_CalRepDates.ocx"
Begin VB.Form frmCalendarReportPreview 
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1069
   Icon            =   "frmCalendarReportPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDateNav 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2280
      TabIndex        =   24
      Top             =   360
      Width           =   4680
      Begin VB.CommandButton cmdToday 
         DisabledPicture =   "frmCalendarReportPreview.frx":000C
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         Picture         =   "frmCalendarReportPreview.frx":03E6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Current Month"
         Top             =   0
         Width           =   345
      End
      Begin VB.ComboBox cboMonth 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCalendarReportPreview.frx":07F0
         Left            =   840
         List            =   "frmCalendarReportPreview.frx":07F2
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   1740
      End
      Begin VB.CommandButton cmdFirstMonth 
         DisabledPicture =   "frmCalendarReportPreview.frx":07F4
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         Picture         =   "frmCalendarReportPreview.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "First Month"
         Top             =   0
         Width           =   405
      End
      Begin VB.CommandButton cmdLastMonth 
         DisabledPicture =   "frmCalendarReportPreview.frx":0EFA
         Enabled         =   0   'False
         Height          =   315
         Left            =   4275
         Picture         =   "frmCalendarReportPreview.frx":12C0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Last Month"
         Top             =   0
         Width           =   405
      End
      Begin VB.CommandButton cmdNextMonth 
         DisabledPicture =   "frmCalendarReportPreview.frx":1604
         Enabled         =   0   'False
         Height          =   315
         Left            =   3915
         Picture         =   "frmCalendarReportPreview.frx":19B7
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Next Month"
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdPrevMonth 
         DisabledPicture =   "frmCalendarReportPreview.frx":1CFB
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         Picture         =   "frmCalendarReportPreview.frx":20AA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Previous Month"
         Top             =   0
         Width           =   285
      End
      Begin COASpinner.COA_Spinner spnYear 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2595
         TabIndex        =   5
         Top             =   0
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaximumValue    =   3000
         Text            =   "2002"
      End
   End
   Begin MSComCtl2.FlatScrollBar VScrollCalendar 
      Height          =   4815
      Left            =   8280
      TabIndex        =   18
      Top             =   1695
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8493
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1179648
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "&Output"
      Height          =   400
      Left            =   6120
      TabIndex        =   0
      Top             =   8520
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   400
      Left            =   7440
      TabIndex        =   1
      Top             =   8520
      Width           =   1200
   End
   Begin VB.Frame fraOptionsShade 
      Caption         =   "Options :"
      Height          =   1635
      Left            =   5880
      TabIndex        =   20
      Top             =   6720
      Width           =   2655
      Begin VB.CheckBox chkIncludeBHols 
         Caption         =   "Include &Bank Holidays"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   2205
      End
      Begin VB.CheckBox chkIncludeWorkingDaysOnly 
         Caption         =   "&Working Days Only"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   510
         Width           =   2280
      End
      Begin VB.CheckBox chkShadeWeekends 
         Caption         =   "Show Wee&kends"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   1320
         Width           =   2010
      End
      Begin VB.CheckBox chkShadeBHols 
         Caption         =   "Show Bank &Holidays"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   780
         Width           =   2145
      End
      Begin VB.CheckBox chkCaptions 
         Caption         =   "Show Calendar Ca&ptions"
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   1050
         Width           =   2445
      End
   End
   Begin VB.Frame fraLegend 
      Caption         =   "Key : "
      Height          =   1635
      Left            =   120
      TabIndex        =   21
      Top             =   6720
      Width           =   5655
      Begin VB.PictureBox picLegendScroll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         Height          =   1150
         Left            =   200
         ScaleHeight     =   1095
         ScaleWidth      =   4935
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   5000
         Begin VB.PictureBox picLegend 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   0
            ScaleHeight     =   735
            ScaleWidth      =   4995
            TabIndex        =   15
            Top             =   0
            Width           =   5000
            Begin VB.Label lblLegend 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "E"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   120
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label lblEventName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Event Name"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   16
               Top             =   120
               Visible         =   0   'False
               Width           =   1785
            End
         End
      End
      Begin MSComCtl2.FlatScrollBar VScrollLegend 
         Height          =   1155
         Left            =   5200
         TabIndex        =   17
         Top             =   300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1179648
      End
   End
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   8175
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   840
      Width           =   8175
      Begin VB.PictureBox picDates 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1920
         ScaleHeight     =   480
         ScaleWidth      =   5865
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Width           =   5895
         Begin VB.Line VerticalDateLine 
            BorderColor     =   &H00000000&
            Index           =   0
            Visible         =   0   'False
            X1              =   480
            X2              =   480
            Y1              =   120
            Y2              =   320
         End
         Begin VB.Line HorizontalDateLine 
            BorderColor     =   &H00000000&
            Index           =   0
            Visible         =   0   'False
            X1              =   720
            X2              =   920
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   33
            Tag             =   "2"
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Tag             =   "2"
            Top             =   255
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picScroll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4200
         Left            =   0
         ScaleHeight     =   4170
         ScaleWidth      =   7965
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   8000
         Begin VB.PictureBox picCalendar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            ForeColor       =   &H80000008&
            Height          =   3840
            Left            =   2040
            ScaleHeight     =   3810
            ScaleWidth      =   5145
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   120
            Width           =   5175
            Begin COACalRepDates.COA_CalRepDates ctlCalDates 
               Height          =   525
               Index           =   0
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   2400
               Visible         =   0   'False
               Width           =   9075
               _ExtentX        =   16007
               _ExtentY        =   926
            End
            Begin VB.Label lblRangeDisabled 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00969696&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1200
               TabIndex        =   36
               Tag             =   "2"
               Top             =   1680
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label lblCalDates 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   35
               Top             =   1680
               Width           =   255
            End
            Begin VB.Label lblDisabled 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   480
               TabIndex        =   30
               Tag             =   "2"
               Top             =   1680
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label lblWeekend 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2040
               TabIndex        =   29
               Tag             =   "2"
               Top             =   1680
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.PictureBox picBase 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   3855
            Left            =   0
            ScaleHeight     =   3855
            ScaleWidth      =   1575
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   1572
            Begin VB.Line HorizontalBaseLine 
               BorderColor     =   &H00000000&
               Index           =   0
               Visible         =   0   'False
               X1              =   120
               X2              =   1320
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label lblBaseDesc 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000006&
               Height          =   510
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Visible         =   0   'False
               Width           =   1500
            End
         End
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdOutput 
      Height          =   840
      Left            =   360
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8520
      Visible         =   0   'False
      Width           =   3075
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      GroupHeadLines  =   0
      DividerType     =   0
      BevelColorFrame =   0
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      RowSelectionStyle=   2
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      PictureRecordSelectors=   "frmCalendarReportPreview.frx":23EE
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ExtraHeight     =   185
      Columns(0).Width=   3413
      Columns(0).Name =   "ButtonColumn"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).Style=   4
      Columns(0).ButtonsAlways=   -1  'True
      TabNavigation   =   1
      _ExtentX        =   5424
      _ExtentY        =   1482
      _StockProps     =   79
      BackColor       =   -2147483643
      Enabled         =   0   'False
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
Attribute VB_Name = "frmCalendarReportPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const KEY_FONTSIZE_SMALL = 7
Private Const KEY_FONTSIZE_MEDIUM = 8
Private Const KEY_FONTSIZE_LARGE = 9

Private mstrSQLSelect_RegInfoRegion As String
Private mstrSQLSelect_BankHolDate As String
Private mstrSQLSelect_BankHolDesc As String

Private mstrSQLSelect_PersonnelStaticRegion As String
Private mstrSQLSelect_PersonnelHRegion As String
Private mstrSQLSelect_PersonnelHDate As String
Private mstrSQLSelect_PersonnelStaticWP As String

Private mstrBaseTableName As String

Private mvarTableViews() As Variant
Private mobjTableView As CTablePrivilege
Private mobjColumnPrivileges As CColumnPrivileges
Private mstrTempRealSource As String
Private mstrRealSource As String
Private mstrViews() As String

'********************************************************************
'new formatting constants
Private Const NAVFRAME_HEIGHT = 315
Private Const NAVFRAME_WIDTH = 4680 '4965 '4240
Private Const NAVCONTROL_OFFSET = 75

Private Const CONTROL_OFFSET = 120

Private Const LEGENDFRAME_HEIGHT = 1650
Private Const LEGENDFRAME_MINWIDTH = 5655

Private Const OPTIONSFRAME_HEIGHT = 1650
Private Const OPTIONSFRAME_WIDTH = 2655 '2535

Private Const SCROLLBAR_WIDTH = 255

Private Const BASEPIC_LEFT = 0

Private Const FORM_MINHEIGHT = 6200

Private mintLegendCount As Integer
Private mintLegendLeft As Integer
Private mintLegendRight As Integer

Private Const FORM_STARTHEIGHT = 7020
Private Const FORM_STARTWIDTH = 8790

Private Const BASE_BOXWIDTH = 1500
Private Const BASE_BOXHEIGHT = 520
Private Const BASE_BOXSTARTX = 0
Private Const BASE_BOXSTARTY = 0

Private Const DAY_BOXWIDTH = 260
Private Const DAY_BOXHEIGHT = 260
Private Const DAY_BOXSTARTX = 0
Private Const DAY_BOXSTARTY = 0

Private Const DATES_BOXWIDTH = 260
Private Const DATES_BOXHEIGHT = 260
Private Const DATES_BOXSTARTX = 0
Private Const DATES_BOXSTARTY = DAY_BOXHEIGHT

Private Const CALDATES_BOXWIDTH = 260

Private Const CALDATES_BOXHEIGHT = 260
Private Const CALDATES_BOXSTARTX = 0
Private Const CALDATES_BOXSTARTY = 0

Private Const DAY_CONTROL_COUNT = 37

Private Const LEGEND_BOXWIDTH = 255
Private Const LEGEND_BOXHEIGHT = 255
Private Const LEGEND_BOXSTARTX = 120
Private Const LEGEND_BOXSTARTY = 120
Private Const LEGEND_BOXOFFSETY = 305

Private Const LEGENDDESC_BOXWIDTH = 3700
Private Const LEGENDDESC_BOXHEIGHT = 255
Private Const LEGENDDESC_BOXSTARTX = 480
Private Const LEGENDDESC_BOXSTARTY = 120
Private Const LEGENDDESC_BOXOFFSETY = 305

Private mblnLoading As Boolean
Private mblnChangingDate As Boolean
'Private gblnBatchMode As Boolean
Private mblnUserCancelled As Boolean
Private mstrErrorMessage As String

Private mblnPersonnelBase As Boolean
Private mstrStaticRegionColumn As String
Private mlngStaticRegionColumnID As Long
Private mstrBaseTableRealSource As String
Private mstrStaticRegionRealSource As String

Private mlngBaseTableID As Long
Private mstrSQLIDs As String

Private mblnRegions As Boolean
Private mblnWorkingPatterns As Boolean

Private mfDefaultToSystemDate As Boolean
Private mdtSystemStartDate As Date
Private mdtSystemEndDate As Date

Private mdtReportStartDate As Date
Private mdtReportEndDate As Date
Private mlngMonth As Long
Private mlngYear As Long

Private mdtVisibleStartDate As Date
Private mdtVisibleEndDate As Date

Private mstrBaseIDColumn As String
Private mstrEventIDColumn As String
Private mlngCurrentRecordID As Long
Private mstrBaseRecDesc As String
Private mstrConvertedBaseRecDesc As String
Private mstrEventToolTip As String
Private mstrCurrentEventKey As String

Private mintCurrentBaseIndex As Integer
Private mstrCurrentBaseRegion As String
Private mintBaseRecordCount As Integer

Private mintFirstDayOfMonth As Integer
Private mintDaysInMonth As Integer

Private mblnGroupByDesc As Boolean

Private mblnShowBankHols As Boolean
Private mblnShowCaptions As Boolean
Private mblnShowWeekends As Boolean
Private mblnStartOnCurrentMonth As Boolean

Private mrsEvents As ADODB.Recordset
Private mrsBase As ADODB.Recordset

Private mavLegend() As Variant
Private mstrArray() As String
Private mstrLegend() As String
Private mavAvailableColours() As Variant

Private mcolStaticBankHolidays As Collection
Private mcolHistoricBankHolidays As Collection
Private mcolStaticWorkingPatterns As Collection
Private mcolHistoricWorkingPatterns As Collection
Private mcolBaseDescIndex As Collection
Private mcolDateControlEvents As Collection

Private mblnDisableRegions As Boolean
Private mblnDisableWPs As Boolean

'-----------------------------------------------
'Event Break-Down Information
Private mstrEventName_BD As String
Private mstrBaseDescription_BD As String
Private mdtEventStartDate_BD As Date
Private mstrEventStartSession_BD As String
Private mdtEventEndDate_BD As Date
Private mstrEventEndSession_BD As String
Private mstrDuration_BD As String
Private mstrDesc1ColumnName_BD As String
Private mstrDesc1Value_BD As String
Private mstrDesc2ColumnName_BD As String
Private mstrDesc2Value_BD As String
Private mstrEventLegend_BD As String
Private mstrCurrentRegion_BD As String
Private mstrCurrentWorkingPattern_BD As String
'-----------------------------------------------

Private mstrDescriptionSeparator As String
Private mlngDescription1ID As Long
Private mlngDescription2ID As Long
Private mlngDescriptionExprID As Long

'New Default Output Variables
Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
'Private mlngOutputSaveFormat As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
'Private mlngOutputEmailFileFormat As Long
Private mstrOutputFileName As String

Private mobjOutput As clsOutputRun

'default output colours
Private mlngBC_Data As Long
Private mlngFC_Data As Long
Private mlngBC_Heading As Long
Private mlngFC_Heading As Long
Private mlngBC_DataOutput As Long

'****************************************************
'variables for outputting
Private mavOutputDateIndex() As Variant

Private mintFirstDayOfMonth_Output As Integer
Private mintDaysInMonth_Output As Integer

Private mintRangeStartIndex_Output As Integer
Private mintRangeEndIndex_Output As Integer

Private mdtVisibleStartDate_Output As Date
Private mdtVisibleEndDate_Output As Date

Private mdtEventStartDate_Output As Date
Private mstrEventStartSession_Output As String
Private mdtEventEndDate_Output As Date
Private mstrEventEndSession_Output As String
Private mstrDuration_Output As String
Private mstrEventLegend_Output As String

Private mlngMonth_Output As Long
Private mlngYear_Output As Long

Private mintCurrentBaseIndex_Output As Integer
Private mintBaseCount_Output As Integer
Private mstrBaseRecDesc_Output As String
Private mintBaseRecordCount_Output As Integer

Private mcolBaseDescIndex_Output As Collection
'****************************************************

Private mstrCalendarReportName As String

Private mintCurrentMonth As Integer
Private mlngCurrentYear As Long

Private mblnBaseDesc1IsDate As Boolean
Private mblnBaseDesc2IsDate As Boolean
Private mblnBaseDescExprIsDate As Boolean

Private mstrExcludedColours As String


Private Type typMemoryStatus
    lngLength As Long
    lngMemoryLoad As Long
    lngTotalPhys As Long
    lngAvailPhys As Long
    lngTotalPageFile As Long
    lngAvailPageFile As Long
    lngTotalVirtual As Long
    lngAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" _
 (lpBuffer As typMemoryStatus)
Private ms As typMemoryStatus

Private mavCareerRanges() As Variant

'****************************************************
'variables for checking for multiple events
Private mavLegendDateIndex() As Variant

Private mintFirstDayOfMonth_Legend As Integer
Private mintDaysInMonth_Legend As Integer

Private mintRangeStartIndex_Legend As Integer
Private mintRangeEndIndex_Legend As Integer

Private mdtVisibleStartDate_Legend As Date
Private mdtVisibleEndDate_Legend As Date

Private mdtEventStartDate_Legend As Date
Private mstrEventStartSession_Legend As String
Private mdtEventEndDate_Legend As Date
Private mstrEventEndSession_Legend As String
Private mstrDuration_Legend As String
Private mstrEventLegend_Legend As String

Private mlngMonth_Legend As Long
Private mlngYear_Legend As Long

Private mintCurrentBaseIndex_Legend As Integer
Private mintBaseCount_Legend As Integer
Private mstrBaseRecDesc_Legend As String
Private mintBaseRecordCount_Legend As Integer

Private mcolBaseDescIndex_Legend As Collection

Private mblnHasMultipleEvents As Boolean
'****************************************************

Private mblnOutputFromPreview As Boolean

Private Const CALREP_DATEFORMAT = "dd/mm/yyyy"

Private mintType_BaseDesc1 As Integer
Private mintType_BaseDesc2 As Integer
Private mintType_BaseDescExpr As Integer
Private mstrFormat_BaseDesc1 As String
Private mstrFormat_BaseDesc2 As String
Private mstrDateFormat As String

Private mblnEnableMouseWheel As Boolean

Private mlngScrollBarMultiplier As Long
Private mdblScrollBarMultiplier As Double



Private Function CheckPermission_Columns(plngTableID As Long, pstrTableName As String, _
                                        pstrColumnName As String, strSQLRef As String) As Boolean

  'This function checks if the current user has read(select) permissions
  'on this column. If the user only has access through views then the
  'relevent views are added to the mvarTableViews() array which in turn
  'are used to create the join part of the query.

  Dim lngTempTableID As Long
  Dim strTempTableName As String
  Dim strTempColumnName As String
  Dim blnColumnOK As Boolean
  Dim blnFound As Boolean
  Dim blnNoSelect As Boolean
  Dim iLoop1 As Integer
  Dim intLoop As Integer
  Dim strColumnCode As String
  Dim strSource As String
  Dim intNextIndex As Integer
  Dim blnOK As Boolean
  Dim strTable As String
  Dim strColumn As String
  
  Dim pintNextIndex  As Integer
  
  ' Set flags with their starting values
  blnOK = True
  blnNoSelect = False

  strTable = vbNullString
  strColumn = vbNullString
 
  ' Load the temp variables
  lngTempTableID = plngTableID
  strTempTableName = pstrTableName
  strTempColumnName = pstrColumnName

  ' Check permission on that column
  Set mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
  mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

  blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

  If blnColumnOK Then
    blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
  End If

  If blnColumnOK Then
    ' this column can be read direct from the tbl/view or from a parent table
    strTable = mstrRealSource
    strColumn = strTempColumnName
    
'    ' If the table isnt the base table (or its realsource) then
'    ' Check if it has already been added to the array. If not, add it.
'    If lngTempTableID <> mlngCalendarReportsBaseTable Then
      blnFound = False
      For intNextIndex = 1 To UBound(mvarTableViews, 2)
        If mvarTableViews(1, intNextIndex) = 0 And _
        mvarTableViews(2, intNextIndex) = lngTempTableID Then
        blnFound = True
          Exit For
        End If
      Next intNextIndex

      If Not blnFound Then
        intNextIndex = UBound(mvarTableViews, 2) + 1
        ReDim Preserve mvarTableViews(3, intNextIndex)
        mvarTableViews(1, intNextIndex) = 0
        mvarTableViews(2, intNextIndex) = lngTempTableID
      End If
'    End If
  
    strSQLRef = strTable & "." & strColumn
  Else

    ' this column cannot be read direct. If its from a parent, try parent views
    ' Loop thru the views on the table, seeing if any have read permis for the column

    ReDim mstrViews(0)
    For Each mobjTableView In gcoTablePrivileges.Collection
      If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = lngTempTableID) And _
          (mobjTableView.AllowSelect) Then

        strSource = mobjTableView.ViewName
        mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

        ' Get the column permission for the view
        Set mobjColumnPrivileges = GetColumnPrivileges(strSource)

        ' If we can see the column from this view
        If mobjColumnPrivileges.IsValid(strTempColumnName) Then
          If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then
            
            ReDim Preserve mstrViews(UBound(mstrViews) + 1)
            mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

            ' Check if view has already been added to the array
            blnFound = False
            For intNextIndex = 0 To UBound(mvarTableViews, 2)
              If mvarTableViews(1, intNextIndex) = 1 And _
              mvarTableViews(2, intNextIndex) = mobjTableView.ViewID Then
                blnFound = True
                Exit For
              End If
            Next intNextIndex

            If Not blnFound Then
              ' View hasnt yet been added, so add it !
              intNextIndex = UBound(mvarTableViews, 2) + 1
              ReDim Preserve mvarTableViews(3, intNextIndex)
              mvarTableViews(0, intNextIndex) = mobjTableView.TableID
              mvarTableViews(1, intNextIndex) = 1
              mvarTableViews(2, intNextIndex) = mobjTableView.ViewID
              mvarTableViews(3, intNextIndex) = mobjTableView.ViewName
            End If
            
          End If
        End If
      End If

    Next mobjTableView
    Set mobjTableView = Nothing

    ' Does the user have select permission thru ANY views ?
    If UBound(mstrViews()) = 0 Then
      blnNoSelect = True
    Else
      strSQLRef = ""
      For pintNextIndex = 1 To UBound(mstrViews)
        If pintNextIndex = 1 Then
          strSQLRef = "CASE"
        End If
        
        strSQLRef = strSQLRef & _
        " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & strTempColumnName
      Next pintNextIndex
      
      If Len(strSQLRef) > 0 Then
        strSQLRef = strSQLRef & _
        " ELSE NULL" & _
        " END "
      End If

    End If

    ' If we cant see a column, then get outta here
    If blnNoSelect Then
      strSQLRef = vbNullString
      CheckPermission_Columns = False
      Exit Function
    End If

    If Not blnOK Then
      strSQLRef = vbNullString
      CheckPermission_Columns = False
      Exit Function
    End If

  End If

'  'TM01042004 Fault 8428
'  If mblnCheckingRegionColumn = True Then
'    mstrRegionColumnRealSource = mstrRealSource
'  End If

  CheckPermission_Columns = True
  
End Function

Public Property Let BaseDesc1IsDate(pblnIsDate As Boolean)
  mblnBaseDesc1IsDate = pblnIsDate
End Property

Public Property Let BaseDescExprIsDate(pblnIsDate As Boolean)
  mblnBaseDescExprIsDate = pblnIsDate
End Property

Public Property Let BaseDesc2IsDate(pblnIsDate As Boolean)
  mblnBaseDesc2IsDate = pblnIsDate
End Property

Private Function ConvertCalendarDateToDateFormat(pstrDateString As String) As Date

  Dim dtTemp As Date
  Dim strDateFormat As String
  Dim lngDay_CR As Long
  Dim lngMonth_CR As Long
  Dim lngYear_CR As Long
  
  Dim blnDateComplete As Boolean
  Dim blnMonthDone As Boolean
  Dim blnDayDone As Boolean
  Dim blnYearDone As Boolean
  
  Dim strShortDate As String
  
  Dim strDateSeparator As String
  
  Dim i As Integer
  
  ' eg. DateFormat = "mm/dd/yyyy"
  '     Calendar   = "dd/mm/yyyy"
  '     DateString = "06/02/2000"
  '     Compare to = 02/06/2000
  
  strDateFormat = DateFormat

  strDateSeparator = UI.GetSystemDateSeparator
  
  blnDateComplete = False
  blnMonthDone = False
  blnDayDone = False
  blnYearDone = False
  
  lngDay_CR = CLng(Mid(pstrDateString, 1, 2))
  lngMonth_CR = CLng(Mid(pstrDateString, 4, 2))
  lngYear_CR = CLng(Mid(pstrDateString, 7, 4))
  
  strShortDate = vbNullString
  
  For i = 1 To Len(strDateFormat) Step 1
    
    If (LCase(Mid(strDateFormat, i, 1)) = "d") And (Not blnDayDone) Then
      strShortDate = strShortDate + LCase(Mid(strDateFormat, i, 1))
      blnDayDone = True
    End If
    
    If (LCase(Mid(strDateFormat, i, 1)) = "m") And (Not blnMonthDone) Then
      strShortDate = strShortDate + LCase(Mid(strDateFormat, i, 1))
      blnMonthDone = True
    End If
    
    If (LCase(Mid(strDateFormat, i, 1)) = "y") And (Not blnYearDone) Then
      strShortDate = strShortDate + LCase(Mid(strDateFormat, i, 1))
      blnYearDone = True
    End If

    If blnDayDone And blnMonthDone And blnYearDone Then
      blnDateComplete = True
      Exit For
    End If
  
  Next i
  
  Select Case strShortDate
    Case "dmy": dtTemp = CDate(lngDay_CR & strDateSeparator & lngMonth_CR & strDateSeparator & lngYear_CR)
    Case "mdy": dtTemp = CDate(lngMonth_CR & strDateSeparator & lngDay_CR & strDateSeparator & lngYear_CR)
    Case "ydm": dtTemp = CDate(lngYear_CR & strDateSeparator & lngDay_CR & strDateSeparator & lngMonth_CR)
    Case "myd": dtTemp = CDate(lngMonth_CR & strDateSeparator & lngYear_CR & strDateSeparator & lngDay_CR)
    Case "ymd": dtTemp = CDate(lngYear_CR & strDateSeparator & lngMonth_CR & strDateSeparator & lngDay_CR)
  End Select
  
  ConvertCalendarDateToDateFormat = dtTemp

End Function

Public Property Let HasMultipleEvents(pblnNewValue As Boolean)
  mblnHasMultipleEvents = pblnNewValue
End Property

Private Function EnableDisableNavigation()

  Dim dtShownStart As Date
  Dim dtShownEnd As Date
  Dim dtSystemMonth As Date
  Dim dtReportStartMonth As Date
  Dim dtReportEndMonth As Date
  Dim intDaysInMonth As Integer
  
  Dim blnNextYear As Boolean
  Dim blnPrevYear As Boolean

  dtShownStart = mdtVisibleStartDate
  dtShownEnd = mdtVisibleEndDate
  dtSystemMonth = Now()
  
  If dtShownStart <= mdtReportStartDate Then
    cmdPrevMonth.Enabled = False
    cmdFirstMonth.Enabled = False
  Else
    cmdPrevMonth.Enabled = True
    cmdFirstMonth.Enabled = True
  End If
  
  If dtShownEnd >= mdtReportEndDate Then
    cmdNextMonth.Enabled = False
    cmdLastMonth.Enabled = False
  Else
    cmdNextMonth.Enabled = True
    cmdLastMonth.Enabled = True
  End If
  
  cboMonth.Enabled = ((cmdNextMonth.Enabled) Or (cmdPrevMonth.Enabled))
  cboMonth.BackColor = (IIf(cboMonth.Enabled, vbWindowBackground, vbButtonFace))
  spnYear.Enabled = (cboMonth.Enabled)
  spnYear.BackColor = (IIf(spnYear.Enabled, vbWindowBackground, vbButtonFace))
  
  dtReportStartMonth = DateAdd("d", CDbl(-(Day(mdtReportStartDate) - 1)), mdtReportStartDate)
  intDaysInMonth = DaysInMonth(mdtReportEndDate)
  dtReportEndMonth = DateAdd("d", CDbl(intDaysInMonth - Day(mdtReportEndDate)), mdtReportEndDate)

  If (Month(dtShownStart) = Month(dtSystemMonth)) And _
    (Year(dtShownStart) = Year(dtSystemMonth)) Then
    cmdToday.Enabled = False
  ElseIf ((dtSystemMonth < dtReportEndMonth) And (dtSystemMonth > dtReportStartMonth)) Then
    cmdToday.Enabled = True
  Else
    cmdToday.Enabled = False
  End If
    
  mintCurrentMonth = Month(dtShownStart)
  mlngCurrentYear = Year(dtShownStart)
  
End Function

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property

Private Function GetAvailableColours(pstrExcludedColours As String) As Boolean

  On Error GoTo ErrorTrap

  Dim rsColours As ADODB.Recordset
  
  Dim intColourCount As Integer
  Dim intNextIndex As Integer
  
  intColourCount = 0
  intNextIndex = 0
  ReDim mavAvailableColours(3, intNextIndex)
  
  Dim strSQL As String
  
  strSQL = vbNullString
  strSQL = strSQL & "SELECT ASRSysColours.ColOrder, ASRSysColours.ColValue, "
  strSQL = strSQL & "       ASRSysColours.ColDesc, ASRSysColours.WordColourIndex, "
  strSQL = strSQL & "       ASRSysColours.CalendarLegendColour "
  strSQL = strSQL & "FROM ASRSysColours "
  strSQL = strSQL & "WHERE (CalendarLegendColour = 1) "
  strSQL = strSQL & "  AND (ASRSysColours.ColValue NOT IN ( " & pstrExcludedColours & ")) "
  strSQL = strSQL & "ORDER BY ASRSysColours.ColOrder "
  
  Set rsColours = datGeneral.GetRecords(strSQL)
  
  With rsColours
    If .BOF And .EOF Then
      GetAvailableColours = False
      GoTo TidyUpAndExit
    End If
    
    .MoveFirst
    Do While Not .EOF
      ReDim Preserve mavAvailableColours(3, intNextIndex)
      
      mavAvailableColours(0, intNextIndex) = !ColValue
      mavAvailableColours(1, intNextIndex) = HexValue(CLng(!ColValue))
      mavAvailableColours(2, intNextIndex) = !ColDesc
      mavAvailableColours(3, intNextIndex) = !WordColourIndex

      intNextIndex = UBound(mavAvailableColours, 2) + 1
      
      .MoveNext
    Loop
    
  End With
  rsColours.Close
  
  GetAvailableColours = True
  
TidyUpAndExit:
  Set rsColours = Nothing
  Exit Function
  
ErrorTrap:
  GetAvailableColours = False
  GoTo TidyUpAndExit

End Function

Private Function GetCurrentRegion(plngBaseRecordID As Long, pdtDate As Date) As String

  Dim intCount As Integer
  
  On Error GoTo ErrorTrap
  
  For intCount = 1 To UBound(mavCareerRanges, 2) Step 1
    If plngBaseRecordID = CLng(mavCareerRanges(0, intCount)) Then
      If mavCareerRanges(2, intCount) <> "" Then
        'has a career change in the past
        If (pdtDate >= CDate(mavCareerRanges(1, intCount))) And (pdtDate < CDate(mavCareerRanges(2, intCount))) Then
          GetCurrentRegion = mavCareerRanges(3, intCount)
          Exit Function
        End If
      Else
        'has a effective start date but has no end date. (most recent career change)
        If (pdtDate >= CDate(mavCareerRanges(1, intCount))) Then
          GetCurrentRegion = mavCareerRanges(3, intCount)
          Exit Function
        End If
      End If
    End If
  Next intCount
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  GetCurrentRegion = vbNullString
  GoTo TidyUpAndExit
  
End Function

Private Function HexValue(plngColour As Long) As String

  Dim strHEX As String
  
  strHEX = Hex(plngColour)
  
  If Len(strHEX) < 6 Then
    strHEX = String(6 - Len(strHEX), "0") & strHEX
  End If
    
  HexValue = "&H" & strHEX

End Function


Private Function OutputArray_GetLegendArray() As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim i As Long
  Dim iLegendCount As Long
  Dim iNewIndex As Long
  
  fOK = True
  
  iLegendCount = 0
  
  mobjOutput.AddColumn "Event Name", sqlVarChar, 0
  mobjOutput.AddColumn "", sqlVarChar, 0
  mobjOutput.AddColumn "Key", sqlVarChar, 0
  
  'add the header row for the Key page
  ReDim mstrLegend(2, 0)
  mstrLegend(0, 0) = "Event Name"
  mstrLegend(1, 0) = "    "
  mstrLegend(2, 0) = " "
  
  ReDim Preserve mstrLegend(2, 1)
  mstrLegend(0, 1) = ""
  mstrLegend(1, 1) = "    "
  mstrLegend(2, 1) = " "
  
  For i = 1 To (UBound(mavLegend, 2) * 2) Step 2
    iLegendCount = iLegendCount + 1
    iNewIndex = UBound(mstrLegend, 2) + 1
    ReDim Preserve mstrLegend(2, iNewIndex)
    mstrLegend(0, iNewIndex) = mavLegend(1, iLegendCount)
    mstrLegend(1, iNewIndex) = "    "
    mstrLegend(2, iNewIndex) = Replace(lblLegend(iLegendCount).Caption, "&&", "&") 'mavLegend(2, iLegendCount)
    
    iNewIndex = UBound(mstrLegend, 2) + 1
    ReDim Preserve mstrLegend(2, iNewIndex)
    mstrLegend(0, iNewIndex) = ""
    mstrLegend(1, iNewIndex) = "    "
    mstrLegend(2, iNewIndex) = ""
  Next i
  
  iLegendCount = 0
  For i = 1 To (UBound(mavLegend, 2) * 2) Step 2
    iLegendCount = iLegendCount + 1
    mobjOutput.AddStyle "", 2, (i + 1), 2, (i + 1), CLng(lblLegend(iLegendCount).BackColor), CLng(lblLegend(iLegendCount).ForeColor), False, False, True
  Next i
 ' CLng(mavLegend(3, iLegendCount))
  OutputArray_GetLegendArray = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_GetLegendArray = False
  GoTo TidyUpAndExit

End Function

Public Property Let OutputFormat(plngOutputFormat As Long)
  mlngOutputFormat = plngOutputFormat
End Property

Public Property Let OutputScreen(pblnOutputScreen As Boolean)
  mblnOutputScreen = pblnOutputScreen
End Property

Public Property Let OutputPrinter(pblnOutputPrinter As Boolean)
  mblnOutputPrinter = pblnOutputPrinter
End Property

Public Property Let OutputPrinterName(pstrOutputPrinterName As String)
  mstrOutputPrinterName = pstrOutputPrinterName
End Property

Public Property Let OutputSave(pblnOutputSave As Boolean)
  mblnOutputSave = pblnOutputSave
End Property

Public Property Let OutputSaveExisting(plngOutputSaveExisting As Long)
  mlngOutputSaveExisting = plngOutputSaveExisting
End Property

'Public Property Let OutputSaveFormat(plngOutputSaveFormat As Long)
'  mlngOutputSaveFormat = plngOutputSaveFormat
'End Property

Public Property Let OutputEmail(pblnOutputEmail As Boolean)
  mblnOutputEmail = pblnOutputEmail
End Property

Public Property Let OutputEmailAddr(plngOutputEmailAddr As Long)
  mlngOutputEmailAddr = plngOutputEmailAddr
End Property

Public Property Let OutputEmailSubject(pstrOutputEmailSubject As String)
  mstrOutputEmailSubject = pstrOutputEmailSubject
End Property

Public Property Let OutputEmailAttachAs(pstrOutputEmailAttachAs As String)
  mstrOutputEmailAttachAs = pstrOutputEmailAttachAs
End Property

'Public Property Let OutputEmailFileFormat(plngOutputEmailFileFormat As String)
'  mlngOutputEmailFileFormat = plngOutputEmailFileFormat
'End Property

Public Property Let OutputFilename(pstrOutputFilename As String)
  mstrOutputFileName = pstrOutputFilename
End Property

Private Function AddDataToArray(pintRow As Integer, pintCol As Integer, _
                                pstrValue As String) As Boolean
  'adds a single row of data to the data array
    
  Dim intNewIndex As Integer
  
  If pintRow > UBound(mstrArray, 2) Then
    intNewIndex = UBound(mstrArray, 2) + 1
    ReDim Preserve mstrArray(37, intNewIndex)
  End If
  
  mstrArray(pintCol, pintRow) = pstrValue
  
  AddDataToArray = True
  
End Function

Private Function RefreshLegend(pblnShowCaptions As Boolean) As Boolean
  
  Dim iCount As Integer
  
  For iCount = 1 To lblLegend.UBound Step 1
    If pblnShowCaptions Then
      lblLegend(iCount).ForeColor = lblLegend(0).ForeColor
    Else
      lblLegend(iCount).ForeColor = lblLegend(iCount).BackColor
    End If
  Next iCount
  
End Function

Private Function SetSystemDateValues() As Boolean

  'Set the initial month on calendar preview to the current system month if within report range.

  Dim intDaysInMonth As Integer
  Dim dtSystemDate As Date
  Dim intSystemMonth As Integer
  Dim intSystemYear As Integer
  Dim intStartMonth As Integer
  Dim intStartYear As Integer
  Dim intEndMonth As Integer
  Dim intEndYear As Integer
  
  dtSystemDate = Now()
  intDaysInMonth = DaysInMonth(dtSystemDate)

  intSystemMonth = Month(dtSystemDate)
  intSystemYear = Year(dtSystemDate)
  intStartMonth = Month(mdtReportStartDate)
  intStartYear = Year(mdtReportStartDate)
  intEndMonth = Month(mdtReportEndDate)
  intEndYear = Year(mdtReportEndDate)
 
  'TM20070514 - Fault 12235 fixed.
  'Got the logic right this time...I think!?!?
  Dim fStartMonthBeforeSysMonth As Boolean
  Dim fStartMonthIsSysMonth As Boolean
  Dim fEndMonthAfterSysMonth As Boolean
  Dim fEndMonthIsSysMonth As Boolean
  
  fStartMonthBeforeSysMonth = ((intStartYear = intSystemYear) And (intStartMonth < intSystemMonth)) _
                            Or (intStartYear < intSystemYear)
  fStartMonthIsSysMonth = (intStartYear = intSystemYear) And (intStartMonth = intSystemMonth)
  
  fEndMonthAfterSysMonth = ((intEndYear = intSystemYear) And (intEndMonth > intSystemMonth)) _
                          Or (intEndYear > intSystemYear)
  fEndMonthIsSysMonth = (intEndYear = intSystemYear) And (intEndMonth = intSystemMonth)
  
  mfDefaultToSystemDate = ((fStartMonthBeforeSysMonth And fEndMonthAfterSysMonth) _
                           Or (fStartMonthBeforeSysMonth And fEndMonthIsSysMonth) _
                           Or (fEndMonthAfterSysMonth And fStartMonthIsSysMonth)) _
                        And mblnStartOnCurrentMonth
                    
  If mfDefaultToSystemDate Then
    mlngMonth = Month(dtSystemDate)
    mlngYear = Year(dtSystemDate)
  Else
    mlngMonth = Month(mdtReportStartDate)
    mlngYear = Year(mdtReportStartDate)
  End If

  mdtSystemEndDate = DateAdd("d", CDbl(intDaysInMonth - Day(dtSystemDate)), dtSystemDate)
  mdtSystemStartDate = DateAdd("d", CDbl(-(intDaysInMonth - 1)), mdtSystemEndDate)

  'TM20070509 - Fault 12218 fixed.
  'Need to format these variables to remove the time part of the datetime value.
  mdtSystemEndDate = CDate(Format(mdtSystemEndDate, DateFormat))
  mdtSystemStartDate = CDate(Format(mdtSystemStartDate, DateFormat))

  SetSystemDateValues = True

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  SetSystemDateValues = False
  GoTo TidyUpAndExit

End Function

Public Function ShowCalendar() As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Set the loading flag
  mblnLoading = True

  If fOK Then fOK = SetSystemDateValues
  
  If fOK Then fOK = ArrangeControls
  
  If fOK Then fOK = PopulateMonthCombo
  
  If fOK Then fOK = SetYear
  
  If fOK Then fOK = Load_Controls
  
  'format scrollbars
  If fOK Then
    picBase.Height = lblBaseDesc(lblBaseDesc.UBound).Top + BASE_BOXHEIGHT
    picCalendar.Height = picBase.Height
  End If
  
  If fOK Then FillGridWithEvents
  
  If fOK Then EnableDisableNavigation
  
  If fOK Then
    mblnChangingDate = True
    fOK = RefreshDateSpecifics
    mblnChangingDate = False
  End If

  If fOK Then Form_Resize

  HookFormSizes

  Screen.MousePointer = vbDefault

  mblnLoading = False

  ShowCalendar = True
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  mblnLoading = False
  ShowCalendar = False
  GoTo TidyUpAndExit
  
End Function

Private Function SetOutputStyles() As Boolean

  Dim intBaseRowCount As Integer
  
  intBaseRowCount = mintBaseRecordCount_Output
 
  'add merge for the empty top left cells
  mobjOutput.AddMerge 0, 0, 0, 1

  '******************************************************************************
  'add style for the weekend ranges if required
  
  If chkShadeWeekends.Value Then
    'first Sunday column (Sunday only)
    mobjOutput.AddStyle "", 1, 2, _
                          1, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
    
    'first Sat, second Sunday
    mobjOutput.AddStyle "", 7, 2, _
                          8, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
    
    'second Sat, third Sunday
    mobjOutput.AddStyle "", 14, 2, _
                          15, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
    
    'third Sat, fourth Sunday
    mobjOutput.AddStyle "", 21, 2, _
                          22, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
    
    'fourth Sat, fifth Sunday
    mobjOutput.AddStyle "", 28, 2, _
                          29, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
    
    'fifth Sat, sixth Sunday
    mobjOutput.AddStyle "", 35, 2, _
                          36, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
  End If
  
  '******************************************************************************
  
  
  'add style for the outside of report date boundaries
  'first out of range (if required)
  If (mintRangeStartIndex_Output > 0) Then
    mobjOutput.AddStyle "", 1, 2, _
                         CLng(mintRangeStartIndex_Output), CLng((2 * intBaseRowCount) + 1), _
                           CLng(lblRangeDisabled.BackColor), lblRangeDisabled.ForeColor, False, False, True
  End If
  
  'second out of range (if required)
  'TM12092003 Fault 6964 - there might not be an out of range, if the last day of the month
  'falls on the last column of the report
  If (mintRangeEndIndex_Output > 0) And (mintRangeEndIndex_Output < 38) Then
    mobjOutput.AddStyle "", CLng(mintRangeEndIndex_Output), 2, _
                         37, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblRangeDisabled.BackColor), lblRangeDisabled.ForeColor, False, False, True
  End If
  
  
  'add style for the disabled ranges
  'first disabled range (if required)
  If (mintFirstDayOfMonth_Output > 1) Then
    mobjOutput.AddStyle "", 1, 2, _
                         (mintFirstDayOfMonth_Output - 1), CLng((2 * intBaseRowCount) + 1), _
                           CLng(lblDisabled.BackColor), lblDisabled.ForeColor, False, False, True
  End If
  
  'second disabled range (if required)
  If ((mintFirstDayOfMonth_Output + mintDaysInMonth_Output) <= 37) Then
    mobjOutput.AddStyle "", (mintFirstDayOfMonth_Output + mintDaysInMonth_Output), 2, _
                         37, CLng((2 * intBaseRowCount) + 1), _
                            CLng(lblDisabled.BackColor), lblDisabled.ForeColor, False, False, True
  End If
  
End Function

Private Function SortLegend(pavLegend As Variant, pintIndex As Integer) As Boolean

  On Error GoTo ErrorTrap
  
  Dim lngCount As Long
  Dim lngRestOfArray As Long
  Dim lngRowIndex As Long
  Dim intStrComp As Integer
  Dim i As Integer
  
  Dim varTemp As Variant
  
  For lngCount = 1 To UBound(pavLegend, 2) Step 1
    lngRowIndex = lngCount
    
    For lngRestOfArray = (lngCount + 1) To UBound(pavLegend, 2) Step 1
      intStrComp = StrComp(pavLegend(pintIndex, lngRowIndex), pavLegend(pintIndex, lngRestOfArray), vbTextCompare)
      If intStrComp = 1 Then
        lngRowIndex = lngRestOfArray
      End If
    Next lngRestOfArray
    
    'put the new lowest in position
    For i = 0 To UBound(pavLegend) Step 1
      varTemp = pavLegend(i, lngRowIndex)
      pavLegend(i, lngRowIndex) = pavLegend(i, lngCount)
      pavLegend(i, lngCount) = varTemp
    Next i
  Next lngCount
  
  SortLegend = True

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  SortLegend = False
  GoTo TidyUpAndExit

End Function

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Private Function ArrangeControls() As Boolean

  With picCalendar
    .Top = 0
  End With
  
  ArrangeControls = True

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  ArrangeControls = False
  GoTo TidyUpAndExit
  
End Function

Public Property Let BaseIDColumn(pstrBaseIDColumn As String)
  mstrBaseIDColumn = pstrBaseIDColumn
End Property

Public Property Let SQLIDs(pstrSQLIDs As String)
  mstrSQLIDs = pstrSQLIDs
End Property

Public Property Let BaseTableRealSource(pstrBaseTableRealSource As String)
  mstrBaseTableRealSource = pstrBaseTableRealSource
End Property

Public Property Let StaticRegionRealSource(pstrStaticRegionRealSource As String)
  mstrStaticRegionRealSource = pstrStaticRegionRealSource
End Property

Public Property Let BaseTableID(pstrBaseTableID As Long)
  mlngBaseTableID = pstrBaseTableID
End Property

Public Property Let BaseTableName(pstrBaseTableName As String)
  mstrBaseTableName = pstrBaseTableName
End Property

Public Property Let CalendarReportName(pstrCalendarReportName As String)
  Me.Caption = "Calendar Report - " & pstrCalendarReportName
  mstrCalendarReportName = pstrCalendarReportName
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
End Property

Private Function EventToolTipText(pdtStartDate As Date, pstrStartSession As String, _
                                    pdtEndDate As Date, pstrEndSession As String) As String
                                    
  Dim strToolTip As String
  
  strToolTip = vbNullString
  strToolTip = strToolTip & "Start Date: " & Format(pdtStartDate, "dd-mmm-yyyy ")
  strToolTip = strToolTip & LCase(pstrStartSession)
  strToolTip = strToolTip & "  --->  "
  strToolTip = strToolTip & "End Date: " & Format(pdtEndDate, "dd-mmm-yyyy ")
  strToolTip = strToolTip & LCase(pstrEndSession)

  EventToolTipText = strToolTip
  
End Function

Private Sub DateChange()
  mblnChangingDate = True
  If Not mblnLoading Then
    Screen.MousePointer = vbHourglass
    Set mcolDateControlEvents = Nothing
    Set mcolDateControlEvents = New Collection
    RefreshCalendar
    FillGridWithEvents
    RefreshDateSpecifics
    Screen.MousePointer = vbDefault
  End If
  
  EnableDisableNavigation

  mblnChangingDate = False
End Sub

Public Property Let EventIDColumn(pstrEventIDColumn As String)
  mstrEventIDColumn = pstrEventIDColumn
End Property

Public Property Let EventsRecordset(prsEventsRecordset As ADODB.Recordset)
  Set mrsEvents = prsEventsRecordset
End Property

Public Property Let BaseRecordset(prsBaseRecordset As ADODB.Recordset)
  Set mrsBase = prsBaseRecordset
End Property

Public Property Let GroupByDescription(pblnGroupByDesc As Boolean)
  mblnGroupByDesc = pblnGroupByDesc
End Property

'Public Property Let BatchMode(pgblnBatchMode As Boolean)
'  gblnBatchMode = pgblnBatchMode
'End Property

Private Function IsBankHoliday(pdtDate As Date, plngBaseID As Long, pstrRegion As String) As Boolean

  On Error GoTo ErrorTrap
  
  Dim colBankHolidays As clsBankHolidays
  Dim objBankHoliday As clsBankHoliday

  If mblnPersonnelBase _
    And (grtRegionType = rtHistoricRegion) _
    And (Not mblnGroupByDesc) _
    And (mlngStaticRegionColumnID < 1) Then
    
    'Need to get the current region from the previously populated.
    'NB. cant get the region from the collection as the current region is required even
    'when the date is NOT a bank holiday
    pstrRegion = GetCurrentRegion(plngBaseID, pdtDate)
    
    'Historic Region Bank Holidays
    Set colBankHolidays = mcolHistoricBankHolidays.Item(CStr(plngBaseID))

    For Each objBankHoliday In colBankHolidays.Collection
      With objBankHoliday
        If pdtDate = .HolidayDate Then
          'pstrRegion = .Region
          IsBankHoliday = True
          GoTo TidyUpAndExit
        End If
      End With
    Next objBankHoliday
    
  ElseIf ((mlngStaticRegionColumnID > 0) _
          Or (mblnPersonnelBase _
              And (grtRegionType = rtStaticRegion))) _
    And (Not mblnGroupByDesc) Then
    
    'Static Region Bank Holidays
    Set colBankHolidays = mcolStaticBankHolidays(CStr(plngBaseID))

    For Each objBankHoliday In colBankHolidays.Collection
      With objBankHoliday
        If pdtDate = .HolidayDate Then
          pstrRegion = .Region
          IsBankHoliday = True
          GoTo TidyUpAndExit
        End If
      End With
    Next objBankHoliday

  End If
  
  IsBankHoliday = False

TidyUpAndExit:
  Set objBankHoliday = Nothing
  Set colBankHolidays = Nothing
  Exit Function
  
ErrorTrap:
  IsBankHoliday = False
  GoTo TidyUpAndExit
  
End Function

Private Function IsWorkingDay(pdtDate As Date, plngBaseID As Long, _
                              pstrSession As String, pstrWorkingPattern As String) As Boolean

  On Error GoTo ErrorTrap
  
  Dim colWorkingPatterns As clsCalendarEvents
  Dim objWorkingPattern As clsCalendarEvent

  Dim strWorkingPattern As String
  Dim intWeekDay As String

  Const WORKINGPATTERN_LENGTH = 14

  strWorkingPattern = "              " 'empty working pattern
  intWeekDay = Weekday(pdtDate, vbSunday)
  
  If mblnPersonnelBase _
    And (gwptWorkingPatternType = wptHistoricWPattern) _
    And (Not mblnGroupByDesc) Then
    
    'Historic Working Pattern

    Set colWorkingPatterns = mcolHistoricWorkingPatterns.Item(CStr(plngBaseID))
    For Each objWorkingPattern In colWorkingPatterns.Collection
      With objWorkingPattern
        
        'TM02072004 Fault 8851 - Force the working pattern length to be 14 characters!
        If Len(.WorkingPattern) < WORKINGPATTERN_LENGTH Then
          .WorkingPattern = .WorkingPattern & String((WORKINGPATTERN_LENGTH - Len(.WorkingPattern)), " ")
        ElseIf Len(.WorkingPattern) > WORKINGPATTERN_LENGTH Then
          .WorkingPattern = Left(.WorkingPattern, WORKINGPATTERN_LENGTH)
        End If
        
        If (.EndDateName <> vbNullString) Then
          If (pdtDate >= CDate(.StartDateName)) And (pdtDate < CDate(.EndDateName)) Then
            Select Case UCase(pstrSession)
              Case "AM"
                If Mid(.WorkingPattern, (intWeekDay * 2) - 1, 1) = " " Then
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = False
                  GoTo TidyUpAndExit
                Else
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = True
                  GoTo TidyUpAndExit
                End If
              Case "PM"
                If Mid(.WorkingPattern, (intWeekDay * 2), 1) = " " Then
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = False
                  GoTo TidyUpAndExit
                Else
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = True
                  GoTo TidyUpAndExit
                End If
            End Select
          End If
        Else
          If (pdtDate >= CDate(.StartDateName)) Then
            Select Case UCase(pstrSession)
              Case "AM"
                If Mid(.WorkingPattern, (intWeekDay * 2) - 1, 1) = " " Then
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = False
                  GoTo TidyUpAndExit
                Else
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = True
                  GoTo TidyUpAndExit
                End If
              Case "PM"
                If Mid(.WorkingPattern, (intWeekDay * 2), 1) = " " Then
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = False
                  GoTo TidyUpAndExit
                Else
                  pstrWorkingPattern = .WorkingPattern
                  IsWorkingDay = True
                  GoTo TidyUpAndExit
                End If
            End Select
          End If
        End If
      End With
    Next objWorkingPattern
    
  ElseIf mblnPersonnelBase _
    And (gwptWorkingPatternType = wptStaticWPattern) _
    And (Not mblnGroupByDesc) Then
    
    'Static Working Pattern
    
    Set colWorkingPatterns = mcolStaticWorkingPatterns.Item(CStr(plngBaseID))
    For Each objWorkingPattern In colWorkingPatterns.Collection
      With objWorkingPattern
        
        'TM02072004 Fault 8851 - Force the working pattern length to be 14 characters!
        If Len(.WorkingPattern) < WORKINGPATTERN_LENGTH Then
          .WorkingPattern = .WorkingPattern & String((WORKINGPATTERN_LENGTH - Len(.WorkingPattern)), " ")
        ElseIf Len(.WorkingPattern) > WORKINGPATTERN_LENGTH Then
          .WorkingPattern = Left(.WorkingPattern, WORKINGPATTERN_LENGTH)
        End If
        
        strWorkingPattern = .WorkingPattern

        Select Case UCase(pstrSession)
          Case "AM"
            If Mid(strWorkingPattern, (intWeekDay * 2) - 1, 1) = " " Then
              pstrWorkingPattern = strWorkingPattern
              IsWorkingDay = False
              GoTo TidyUpAndExit
            Else
              pstrWorkingPattern = strWorkingPattern
              IsWorkingDay = True
              GoTo TidyUpAndExit
            End If
          Case "PM"
            If Mid(strWorkingPattern, (intWeekDay * 2), 1) = " " Then
              pstrWorkingPattern = strWorkingPattern
              IsWorkingDay = False
              GoTo TidyUpAndExit
            Else
              pstrWorkingPattern = strWorkingPattern
              IsWorkingDay = True
              GoTo TidyUpAndExit
            End If
        End Select
      End With
    Next objWorkingPattern
  End If
  
  pstrWorkingPattern = "              "
  IsWorkingDay = False

TidyUpAndExit:
  Set objWorkingPattern = Nothing
  Set colWorkingPatterns = Nothing
  Exit Function
  
ErrorTrap:
  pstrWorkingPattern = "              "
  IsWorkingDay = False
  GoTo TidyUpAndExit
  
End Function

Private Function IsWeekend(pdtDate As Date) As Boolean
  If (Weekday(pdtDate, vbSunday) = vbSaturday) Or (Weekday(pdtDate, vbSunday) = vbSunday) Then
    IsWeekend = True
  Else
    IsWeekend = False
  End If
End Function

Private Function Load_Controls() As Boolean

  Dim fOK As Boolean
  Dim blnNewBaseRecord As Boolean
  Dim intNextArrayIndex As Integer
  Dim intStartIndex As Integer
  Dim intEndIndex As Integer
  Dim intNewIndex As Integer
  Dim strTempRecordDesc As String
  Dim intDescEmpty As Integer
  Dim blnDescEmpty As Boolean
  Dim strBaseDescription1, strBaseDescription2, strBaseDescriptionExpr As String
  Dim iDecimals As Integer
  Dim lngRecordNo As Long
  
  fOK = True
  
  mintBaseRecordCount = 0
  mlngCurrentRecordID = -1
  mstrBaseRecDesc = vbNullString
  mblnBaseDesc1IsDate = False
  mblnBaseDesc2IsDate = False
  mblnBaseDescExprIsDate = False
  mstrConvertedBaseRecDesc = False
  mintCurrentBaseIndex = False
  Set mcolBaseDescIndex = New Collection
  
  mintType_BaseDesc1 = -1
  mintType_BaseDesc2 = -1
  mintType_BaseDescExpr = -1
  
  If fOK Then
    fOK = Load_Days
  End If

  If fOK Then
    fOK = Load_Dates()
  End If
  
  If fOK Then
    fOK = Load_DateVerticalLines
  End If
  
  If fOK Then
    picDates.Width = (lblDate(lblDate.UBound).Left + lblDate(lblDate.UBound).Width + 30)
    picDates.Height = (lblDate(lblDate.UBound).Top + lblDate(lblDate.UBound).Height)
    picCalendar.Width = picDates.Width
  End If
  
  With mrsBase
    If .BOF And .EOF Then
      Load_Controls = False
      Exit Function
    End If
    
    'get the datatype/properties for the desc1 column
    If (mlngDescription1ID > 0) Then
      If datGeneral.DoesColumnUseSeparators(mlngDescription1ID) Then
        mintType_BaseDesc1 = 3
        iDecimals = datGeneral.GetDecimalsSize(mlngDescription1ID)
        mstrFormat_BaseDesc1 = "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "#"), "")
      ElseIf datGeneral.BitColumn("C", mlngBaseTableID, mlngDescription1ID) Then
        mintType_BaseDesc1 = 2
      ElseIf datGeneral.DateColumn("C", mlngBaseTableID, mlngDescription1ID) Then
        mintType_BaseDesc1 = 1
      Else
        mintType_BaseDesc1 = 0
      End If
    End If
    'get the datatype/properties for the desc2 column
    If (mlngDescription2ID > 0) Then
      If datGeneral.DoesColumnUseSeparators(mlngDescription2ID) Then
        mintType_BaseDesc2 = 3
        iDecimals = datGeneral.GetDecimalsSize(mlngDescription2ID)
        mstrFormat_BaseDesc2 = "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "#"), "")
      ElseIf datGeneral.BitColumn("C", mlngBaseTableID, mlngDescription2ID) Then
        mintType_BaseDesc2 = 2
      ElseIf datGeneral.DateColumn("C", mlngBaseTableID, mlngDescription2ID) Then
        mintType_BaseDesc2 = 1
      Else
        mintType_BaseDesc2 = 0
      End If
    End If
    'get the datatype/properties for the descexpr column
    If (mlngDescriptionExprID > 0) Then
      If datGeneral.BitColumn("X", mlngBaseTableID, mlngDescriptionExprID) Then
        mintType_BaseDescExpr = 2
      ElseIf datGeneral.DateColumn("X", mlngBaseTableID, mlngDescriptionExprID) Then
        mintType_BaseDescExpr = 1
      Else
        mintType_BaseDescExpr = 0
      End If
    End If

    .MoveFirst
    lngRecordNo = 0
    Do While Not .EOF
    
      If gobjProgress.Visible Then
        ' Update the progress bar
        gobjProgress.UpdateProgress gblnBatchMode
      End If
      
      ' If user cancels the report, abort
      lngRecordNo = lngRecordNo + 1
      
      'Update the progress bar
      If lngRecordNo Mod 50 = 0 Then
        If gobjProgress.Cancelled Then
          mblnUserCancelled = True
          Load_Controls = False
          Exit Function
        End If
      End If
      
      ' JDM - 05/08/03 Fault 5605 - Put separators in
      ' Get base description 1
      If Not IsNull(.Fields("Description1").Value) Then
        Select Case mintType_BaseDesc1
          Case 3: strBaseDescription1 = Format(.Fields("Description1").Value, mstrFormat_BaseDesc1)
          Case 2: strBaseDescription1 = IIf(.Fields("Description1").Value, "Y", "N")
          Case 1: strBaseDescription1 = Format(.Fields("Description1").Value, mstrDateFormat)
          Case 0: strBaseDescription1 = .Fields("Description1").Value
        End Select
      Else
        strBaseDescription1 = vbNullString
      End If
      ' Get base description 2
      If Not IsNull(.Fields("Description2").Value) Then
        Select Case mintType_BaseDesc2
          Case 3: strBaseDescription2 = Format(.Fields("Description2").Value, mstrFormat_BaseDesc2)
          Case 2: strBaseDescription2 = IIf(.Fields("Description2").Value, "Y", "N")
          Case 1: strBaseDescription2 = Format(.Fields("Description2").Value, mstrDateFormat)
          Case 0: strBaseDescription2 = .Fields("Description2").Value
        End Select
      Else
        strBaseDescription2 = vbNullString
      End If
      ' Get base description expression
      If Not IsNull(.Fields("DescriptionExpr").Value) Then
        Select Case mintType_BaseDescExpr
          Case 2: strBaseDescriptionExpr = IIf(.Fields("DescriptionExpr").Value, "Y", "N")
          Case 1: strBaseDescriptionExpr = Format(.Fields("DescriptionExpr").Value, mstrDateFormat)
          Case 0: strBaseDescriptionExpr = .Fields("DescriptionExpr").Value
        End Select
      Else
        strBaseDescriptionExpr = vbNullString
      End If
      
      strTempRecordDesc = strBaseDescription1
      strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescription2)), mstrDescriptionSeparator, "") & strBaseDescription2
      strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescriptionExpr) > 0), mstrDescriptionSeparator, "") & strBaseDescriptionExpr
      
      'TM20030521 Fault 5736
      blnDescEmpty = (strTempRecordDesc = vbNullString)
      If blnDescEmpty Then
        intDescEmpty = intDescEmpty + 1
      Else
        intDescEmpty = 0
      End If
      
      If mblnGroupByDesc Then
        If ((strTempRecordDesc) <> mstrBaseRecDesc) Or (blnDescEmpty And Int(intDescEmpty = 1)) Then
          blnNewBaseRecord = True
          blnDescEmpty = False
                              
          mstrBaseRecDesc = strTempRecordDesc
          
          If Len(Trim(mstrStaticRegionColumn)) > 0 Then
            mstrCurrentBaseRegion = IIf(IsNull(.Fields("Region").Value), "", .Fields("Region").Value)
          End If
          mintBaseRecordCount = mintBaseRecordCount + 1
        End If
        mlngCurrentRecordID = .Fields(mstrBaseIDColumn).Value
        
      Else
        If .Fields(mstrBaseIDColumn).Value <> mlngCurrentRecordID Then
          blnNewBaseRecord = True
          
          mstrBaseRecDesc = strTempRecordDesc
         
          mlngCurrentRecordID = .Fields(mstrBaseIDColumn).Value
          If Len(Trim(mstrStaticRegionColumn)) > 0 Then
            mstrCurrentBaseRegion = IIf(IsNull(.Fields("Region").Value), "", .Fields("Region").Value)
          End If
          mintBaseRecordCount = mintBaseRecordCount + 1
        End If
        
      End If
      
      If fOK And blnNewBaseRecord Then

        fOK = Load_Description(mlngCurrentRecordID, mstrBaseRecDesc)
        
        If fOK Then
          fOK = Load_Calendar
        End If
        
        If fOK Then
          fOK = Load_BaseHorizontalLines
        End If
      
      End If

      mcolBaseDescIndex.Add mintCurrentBaseIndex, CStr(mlngCurrentRecordID)

      blnNewBaseRecord = False
      
      .MoveNext
    Loop
  End With
  
  Load_Controls = fOK
    
End Function

Private Function FillGridWithEvents() As Boolean
  
  On Error Resume Next
  
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim lngCurrentBaseID As Long
  Dim intBaseRecordIndex As Integer
  Dim iDecimals As Integer
  Dim strFormat As String
  Dim strBaseDescription1, strBaseDescription2, strBaseDescriptionExpr As String
  
  Dim fOK As Boolean
  
  Dim sSQL As String

  fOK = True
  
  lngCurrentBaseID = -1
  intBaseRecordIndex = -1
  strBaseDescription1 = vbNullString
  strBaseDescription2 = vbNullString
  strBaseDescriptionExpr = vbNullString
  mstrBaseDescription_BD = vbNullString
  mstrDesc1ColumnName_BD = vbNullString
  mstrDesc1Value_BD = vbNullString
  mstrDesc2ColumnName_BD = vbNullString
  mstrDesc2Value_BD = vbNullString
  mstrCurrentEventKey = vbNullString
  mstrEventName_BD = vbNullString
  mstrEventLegend_BD = vbNullString
  
  With mrsEvents
  
    ' If there are no event records, skip this bit
    ' this bit (but still show the form)
    If .BOF And .EOF Then
      Exit Function
    End If
      
    .MoveFirst
    ' Loop through the events recordset
    Do Until .EOF

      If gobjProgress.Visible = True Then
        ' Update the progress bar
        gobjProgress.UpdateProgress gblnBatchMode
      End If
      
      ' If user cancels the report, abort
      If gobjProgress.Cancelled Then
        mblnUserCancelled = True
        FillGridWithEvents = False
        Exit Function
      End If

      lngCurrentBaseID = mrsEvents.Fields(mstrBaseIDColumn)
      
      intBaseRecordIndex = mcolBaseDescIndex.Item(CStr(lngCurrentBaseID))
      
      ' Get base description 1
      If Not IsNull(.Fields("Description1").Value) Then
        Select Case mintType_BaseDesc1
          Case 3: strBaseDescription1 = Format(.Fields("Description1").Value, mstrFormat_BaseDesc1)
          Case 2: strBaseDescription1 = IIf(.Fields("Description1").Value, "Y", "N")
          Case 1: strBaseDescription1 = Format(.Fields("Description1").Value, mstrDateFormat)
          Case 0: strBaseDescription1 = .Fields("Description1").Value
        End Select
      Else
        strBaseDescription1 = vbNullString
      End If
      ' Get base description 2
      If Not IsNull(.Fields("Description2").Value) Then
        Select Case mintType_BaseDesc2
          Case 3: strBaseDescription2 = Format(.Fields("Description2").Value, mstrFormat_BaseDesc2)
          Case 2: strBaseDescription2 = IIf(.Fields("Description2").Value, "Y", "N")
          Case 1: strBaseDescription2 = Format(.Fields("Description2").Value, mstrDateFormat)
          Case 0: strBaseDescription2 = .Fields("Description2").Value
        End Select
      Else
        strBaseDescription2 = vbNullString
      End If
      ' Get base description expression
      If Not IsNull(.Fields("DescriptionExpr").Value) Then
        Select Case mintType_BaseDescExpr
          Case 2: strBaseDescriptionExpr = IIf(.Fields("DescriptionExpr").Value, "Y", "N")
          Case 1: strBaseDescriptionExpr = Format(.Fields("DescriptionExpr").Value, mstrDateFormat)
          Case 0: strBaseDescriptionExpr = .Fields("DescriptionExpr").Value
        End Select
      Else
        strBaseDescriptionExpr = vbNullString
      End If
      
      'Concat base expression using the Description Separator.
      mstrBaseDescription_BD = strBaseDescription1
      mstrBaseDescription_BD = mstrBaseDescription_BD & IIf((Len(mstrBaseDescription_BD) > 0) And (Len(strBaseDescription2) > 0), mstrDescriptionSeparator, "") & strBaseDescription2
      mstrBaseDescription_BD = mstrBaseDescription_BD & IIf((Len(mstrBaseDescription_BD) > 0) And (Len(strBaseDescriptionExpr) > 0), mstrDescriptionSeparator, "") & strBaseDescriptionExpr
      
      'Get Event Description 1
      mstrDesc1ColumnName_BD = CStr(IIf(IsNull(.Fields("EventDescription1Column").Value), "", .Fields("EventDescription1Column").Value))
      If Not IsNull(.Fields("EventDescription1ColumnID").Value) Then
        If datGeneral.DoesColumnUseSeparators(.Fields("EventDescription1ColumnID").Value) Then
          iDecimals = datGeneral.GetDecimalsSize(.Fields("EventDescription1ColumnID").Value)
          strFormat = "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "#"), "")
          mstrDesc1Value_BD = Format(.Fields("EventDescription1").Value, strFormat)
        ElseIf datGeneral.GetColumnDataType(.Fields("EventDescription1ColumnID").Value) = SQLDataType.sqlDate Then
          mstrDesc1Value_BD = Format(.Fields("EventDescription1").Value, mstrDateFormat)
        Else
          mstrDesc1Value_BD = IIf(IsNull(.Fields("EventDescription1").Value), "", .Fields("EventDescription1").Value)
        End If
      Else
        mstrDesc1Value_BD = vbNullString
      End If
      
      'Get Event Description 2
      mstrDesc2ColumnName_BD = CStr(IIf(IsNull(.Fields("EventDescription2Column").Value), "", .Fields("EventDescription2Column").Value))
      If Not IsNull(.Fields("EventDescription2Column").Value) Then
        If datGeneral.DoesColumnUseSeparators(.Fields("EventDescription2ColumnID").Value) Then
          iDecimals = datGeneral.GetDecimalsSize(.Fields("EventDescription2ColumnID").Value)
          strFormat = "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "#"), "")
          mstrDesc2Value_BD = Format(.Fields("EventDescription2").Value, strFormat)
        ElseIf datGeneral.GetColumnDataType(.Fields("EventDescription2ColumnID").Value) = SQLDataType.sqlDate Then
          mstrDesc2Value_BD = Format(.Fields("EventDescription2").Value, mstrDateFormat)
        Else
          mstrDesc2Value_BD = IIf(IsNull(.Fields("EventDescription2").Value), "", .Fields("EventDescription2").Value)
        End If
      Else
        mstrDesc2Value_BD = vbNullString
      End If
      
         
      ' Load each event record data into variables
      ' (has to be done because start/end dates may be modified by code to fill grid correctly)
      mstrCurrentEventKey = .Fields(mstrEventIDColumn)
      
      mstrEventName_BD = .Fields("Name").Value
      
      mstrEventLegend_BD = IIf(IsNull(.Fields("Legend")), "", Left(.Fields("Legend").Value, 2))
     
      '****************************************************************************
      mdtEventStartDate_BD = Format(.Fields("StartDate"), mstrDateFormat)
    
      If IsNull(.Fields("EndDate")) Then
        mdtEventEndDate_BD = mdtEventStartDate_BD
      Else
        mdtEventEndDate_BD = Format(.Fields("EndDate"), mstrDateFormat)
      End If
  
      If IsNull(.Fields("StartSession")) And IsNull(.Fields("EndSession")) Then
        mstrEventStartSession_BD = "AM"
        mstrEventEndSession_BD = "PM"
      ElseIf IsNull(.Fields("EndSession")) Then
        mstrEventEndSession_BD = mstrEventStartSession_BD
      Else
        mstrEventStartSession_BD = UCase(.Fields("StartSession"))
        mstrEventEndSession_BD = UCase(.Fields("EndSession"))
      End If

      mstrEventToolTip = EventToolTipText(mdtEventStartDate_BD, mstrEventStartSession_BD, _
                                        mdtEventEndDate_BD, mstrEventEndSession_BD)
    
'      'Force the Start & End Dates to be between the Report Start and End dates.
'      If mdtEventStartDate_BD < mdtReportStartDate Then
'        mdtEventStartDate_BD = mdtReportStartDate
'      End If
'
'      If mdtEventEndDate_BD > mdtReportEndDate Then
'        mdtEventEndDate_BD = mdtReportEndDate
'      End If

      'Force the Start & End Dates to be between the Visible Start and End dates.
      If mdtEventStartDate_BD < mdtVisibleStartDate Then
        mdtEventStartDate_BD = mdtReportStartDate
      End If

      If mdtEventEndDate_BD > mdtReportEndDate Then
        mdtEventEndDate_BD = mdtReportEndDate
      End If

      mstrDuration_BD = .Fields("Duration")
      
      '****************************************************************************
      
      ' If the event start date is after the event end date, ignore the record
      If (mdtEventStartDate_BD > mdtEventEndDate_BD) Then
      
      ' if the event is totally before the currently viewed timespan then do nothing
      ElseIf (mdtEventStartDate_BD < mdtVisibleStartDate) _
        And (mdtEventEndDate_BD < mdtVisibleStartDate) Then
      
      ' if the event is totally after the currently viewed timespan then do nothing
      ElseIf (mdtEventStartDate_BD > mdtVisibleEndDate) _
        And (mdtEventEndDate_BD > mdtVisibleEndDate) Then
      
      ' if the event starts before currently viewed timespan, but ends in the timspan then
      ElseIf (mdtEventStartDate_BD < mdtVisibleStartDate) _
        And (mdtEventEndDate_BD <= mdtVisibleEndDate) Then
        
        mdtEventStartDate_BD = mdtVisibleStartDate
        mstrEventStartSession_BD = "AM"
        
        lngStart = GetCalLabelIndex(intBaseRecordIndex, mdtEventStartDate_BD, IIf(mstrEventStartSession_BD = "AM", False, True))
        lngEnd = GetCalLabelIndex(intBaseRecordIndex, mdtEventEndDate_BD, IIf(mstrEventEndSession_BD = "AM", False, True))
  
        fOK = FillEventCalBoxes(intBaseRecordIndex, lngStart, lngEnd)
        
      ' if the event starts in the currently viewed timespan, but ends after it then
      ElseIf (mdtEventStartDate_BD >= mdtVisibleStartDate) _
        And (mdtEventEndDate_BD > mdtVisibleEndDate) Then
        
        mdtEventEndDate_BD = mdtVisibleEndDate
        mstrEventEndSession_BD = "PM"
        
        lngStart = GetCalLabelIndex(intBaseRecordIndex, mdtEventStartDate_BD, IIf(mstrEventStartSession_BD = "AM", False, True))
        lngEnd = GetCalLabelIndex(intBaseRecordIndex, mdtEventEndDate_BD, IIf(mstrEventEndSession_BD = "AM", False, True))
  
        fOK = FillEventCalBoxes(intBaseRecordIndex, lngStart, lngEnd)
  
      ' if the event is enclosed within viewed timespan, and months are equal then
      ElseIf (mdtEventStartDate_BD >= mdtVisibleStartDate) _
        And (mdtEventEndDate_BD <= mdtVisibleEndDate) _
        And (Month(mdtEventStartDate_BD) = Month(mdtEventEndDate_BD)) _
          Then
  
        lngStart = GetCalLabelIndex(intBaseRecordIndex, mdtEventStartDate_BD, IIf(mstrEventStartSession_BD = "AM", False, True))
        lngEnd = GetCalLabelIndex(intBaseRecordIndex, mdtEventEndDate_BD, IIf(mstrEventEndSession_BD = "AM", False, True))
  
        fOK = FillEventCalBoxes(intBaseRecordIndex, lngStart, lngEnd)
      
      ' if the event starts before the the viewed timespan and ends after the viewed timespan then
      ElseIf (mdtEventStartDate_BD < mdtVisibleStartDate) _
        And (mdtEventEndDate_BD > mdtVisibleEndDate) Then
        
        mdtEventStartDate_BD = mdtVisibleStartDate
        mstrEventStartSession_BD = "AM"
        
        mdtEventEndDate_BD = mdtVisibleEndDate
        mstrEventEndSession_BD = "PM"
                
        lngStart = GetCalLabelIndex(intBaseRecordIndex, mdtEventStartDate_BD, IIf(mstrEventStartSession_BD = "AM", False, True))
        lngEnd = GetCalLabelIndex(intBaseRecordIndex, mdtEventEndDate_BD, IIf(mstrEventEndSession_BD = "AM", False, True))
  
        fOK = FillEventCalBoxes(intBaseRecordIndex, lngStart, lngEnd)
            
      End If
      
      If fOK = False Then
        Exit Do
      End If
    
      .MoveNext
    
    Loop
  End With
  
  If fOK = False Then
    COAMsgBox "An Error Has Occurred Whilst Filling The Cal Labels:" & vbNewLine & Err.Number & " - " & Err.Description
  End If

End Function

Private Function Load_Days() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim iNewIndex As Integer
  Dim iDayCount As Integer
  Dim sDay As String
  
  iDayCount = 1
  sDay = vbNullString
  iNewIndex = 0

'  mobjOutput.AddColumn "", sqlVarChar, 0
'  AddDataToArray 0, 0, ""
    
  Do While iNewIndex < DAY_CONTROL_COUNT
  
    iNewIndex = lblDay().UBound + 1
    
    Load lblDay(iNewIndex)
    
    With lblDay(iNewIndex)
      
'      Select Case iDayCount
'        Case DayConstants.mvwSunday: sDay = "S"
'        Case DayConstants.mvwMonday: sDay = "M"
'        Case DayConstants.mvwTuesday: sDay = "T"
'        Case DayConstants.mvwWednesday: sDay = "W"
'        Case DayConstants.mvwThursday: sDay = "T"
'        Case DayConstants.mvwFriday: sDay = "F"
'        Case DayConstants.mvwSaturday: sDay = "S"
'      End Select
      
      sDay = Left(WeekdayName(CLng(iDayCount), True, vbSunday), 1)
      
      .Tag = iDayCount
      .Caption = sDay
      .Top = DAY_BOXSTARTY
      .Left = (DAY_BOXSTARTX + ((DAY_BOXWIDTH - 15) * (iNewIndex - 1)))
      .Visible = True
      
'      mobjOutput.AddColumn .Caption, sqlVarChar, 0
'      AddDataToArray 0, iNewIndex, .Caption
      
      If iDayCount = DayConstants.mvwSaturday Then
        iDayCount = 0
      End If
      iDayCount = iDayCount + 1
    End With
    
  Loop
  
  Load_Days = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_Days = False
  GoTo TidyUpAndExit

End Function

Private Function OutputArray_AddDays() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim iDayCount As Integer
  Dim sDay As String
  Dim intCount As Integer
  
  iDayCount = 1
  sDay = vbNullString
  intCount = 0
  
  mobjOutput.AddColumn "", sqlVarChar, 0
  AddDataToArray 0, 0, ""
    
  For intCount = 1 To DAY_CONTROL_COUNT Step 1
  
'    Select Case iDayCount
'      Case DayConstants.mvwSunday: sDay = "S"
'      Case DayConstants.mvwMonday: sDay = "M"
'      Case DayConstants.mvwTuesday: sDay = "T"
'      Case DayConstants.mvwWednesday: sDay = "W"
'      Case DayConstants.mvwThursday: sDay = "T"
'      Case DayConstants.mvwFriday: sDay = "F"
'      Case DayConstants.mvwSaturday: sDay = "S"
'    End Select
    
    sDay = Left(WeekdayName(CLng(iDayCount), True, vbSunday), 1)

    mobjOutput.AddColumn sDay, sqlVarChar, 0
    AddDataToArray 0, intCount, sDay
    
    If iDayCount = DayConstants.mvwSaturday Then
      iDayCount = 0
    End If
    iDayCount = iDayCount + 1
    
  Next intCount
  
  OutputArray_AddDays = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_AddDays = False
  GoTo TidyUpAndExit

End Function

Private Function OutputArray_GetArray() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If fOK Then fOK = OutputArray_AddDays()
  If Not gblnBatchMode Then gobjProgress.UpdateProgress gblnBatchMode
  
  If fOK Then fOK = OutputArray_AddDates()
  If Not gblnBatchMode Then gobjProgress.UpdateProgress gblnBatchMode
  
  If fOK Then fOK = OutputArray_AddCalendar()
  If Not gblnBatchMode Then gobjProgress.UpdateProgress gblnBatchMode
  
  If fOK Then fOK = OutputArray_AddEvents()
  If Not gblnBatchMode Then gobjProgress.UpdateProgress gblnBatchMode
  
  If fOK Then fOK = OutputArray_RefreshDateSpecifics()
  If Not gblnBatchMode Then gobjProgress.UpdateProgress gblnBatchMode
    
  OutputArray_GetArray = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_GetArray = False
  GoTo TidyUpAndExit

End Function

Private Function OutputArray_RefreshDateSpecifics() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim intCount As Integer
  
  'following variables used to establish required back & fore color for the label
  Dim blnIsWeekend As Boolean
  Dim blnIsBankHoliday As Boolean
  Dim blnIsWorkingDay As Boolean
  Dim blnIncBankHoliday As Boolean
  Dim blnIncWorkingDays As Boolean
  Dim blnShadeBankHolidays As Boolean
  Dim blnShadeWeekends As Boolean
  Dim blnHasEvent As Boolean
  Dim blnShowCaption As Boolean
  Dim intDefinedColourStyle As Integer
  
  Dim strColour As String
  Dim intThisStartCount As Integer
  Dim intThisEndCount As Integer
  Dim intNextStartCount As Integer
  Dim intNext2StartCount As Integer
  Dim intIndexModulus As Integer
  Dim intCurrentStartCount As Integer
  Dim intCurrentEndCount As Integer
  Dim intBaseCount As Integer
  
  Dim strSession As String

  Dim blnNextHasEvent As Boolean
  Dim blnNext2HasEvent As Boolean
  Dim blnPrevHasEvent As Boolean

  Dim intSessionCount As Integer
  
  Dim varTempArray As Variant
  
  Dim strBaseDesc As String
  Dim strBackColour As String
  Dim strForeColour As String
  Dim strCaption As String
  
  Dim dtConvertedDate As Date
 
  intSessionCount = 0
  
  If mintBaseRecordCount_Output < 1 Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  
  blnIncBankHoliday = chkIncludeBHols.Value
  blnIncWorkingDays = chkIncludeWorkingDaysOnly.Value
  blnShadeBankHolidays = chkShadeBHols.Value
  blnShadeWeekends = chkShadeWeekends.Value

  SetOutputStyles
  
  For intBaseCount = 1 To mintBaseRecordCount_Output Step 1
    
    strBaseDesc = mavOutputDateIndex(1, intBaseCount)
    
    'add the Description values to the array
    AddDataToArray (intBaseCount * 2), 0, strBaseDesc
    AddDataToArray ((intBaseCount * 2) + 1), 0, ""
    
    mobjOutput.AddMerge 0, (intBaseCount * 2), 0, ((intBaseCount * 2) + 1)
    
    varTempArray = mavOutputDateIndex(2, intBaseCount)
    
    For intCount = 1 To 74 Step 1
    
      intSessionCount = intSessionCount + 1
    
      If varTempArray(1, intCount) = "  /  /    " Then
        varTempArray(8, intCount) = HexValue(lblDisabled.BackColor)
        varTempArray(7, intCount) = ""
        
        If intSessionCount = 2 Then
          intSessionCount = 0
        End If

      Else
        dtConvertedDate = ConvertCalendarDateToDateFormat(CStr(varTempArray(1, intCount)))
        If (dtConvertedDate >= mdtReportStartDate) And (dtConvertedDate <= mdtReportEndDate) Then
        
          blnIsBankHoliday = IIf(varTempArray(3, intCount) = "1", True, False)
          blnIsWeekend = IIf(varTempArray(4, intCount) = "1", True, False)
          strColour = varTempArray(8, intCount)
          blnHasEvent = IIf(varTempArray(6, intCount) > 0, True, False)
          blnIsWorkingDay = IIf(varTempArray(5, intCount) = "1", True, False)
          
          intDefinedColourStyle = 0   'Default Colour
'          intDefinedColourStyle = 1   'Weekend/Bank Holiday Colour
'          intDefinedColourStyle = 2   'Event Key Colour
          
          If blnHasEvent Then
            'Event
            intDefinedColourStyle = 2
            
            If (blnIsWorkingDay) Then
              'Event + Working Day
              
              If (blnIsBankHoliday) And (Not blnIncBankHoliday) And (Not blnShadeBankHolidays) And (Not ((blnIsWeekend) And (blnShadeWeekends))) Then
                'Event + Working Day + ((Bank Holiday + Not Inc. Working Days Only + Not Shade Bank Holidays)))
                intDefinedColourStyle = 0
              ElseIf (blnIsBankHoliday) And (Not blnIncBankHoliday) And ((blnShadeBankHolidays) Or ((blnIsWeekend) And (blnShadeWeekends))) Then
                'Event + Working Day + ((Bank Holiday + Not Inc. Working Days Only + Shade Bank Holidays)))
                intDefinedColourStyle = 1
              End If
              
            Else
              'Event + Not Working Day
              
              If (blnIncWorkingDays) And ((blnIsBankHoliday And Not blnIncBankHoliday) Or (Not blnIsBankHoliday)) And ((blnIsWeekend And Not blnShadeWeekends) Or (Not blnIsWeekend)) Then
                'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only) || Not Bank Holiday) + ((Weekend + Not Show Weekends) || Not Weekend))
                intDefinedColourStyle = 0
              End If
              
              If (blnIsBankHoliday) And (blnShadeBankHolidays) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                'Event + Not Working Day + Bank Holiday + Shade Bank Holidays + Inc. Working Days Only + Not Inc. Bank Holidays
                intDefinedColourStyle = 1
              ElseIf (blnIsWeekend) And (blnShadeWeekends) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                'Event + Not Working Day + Weekend + Show Weekends + Inc. Working Days Only + Not Inc. Bank Holidays
                intDefinedColourStyle = 1
              ElseIf (blnIsWeekend) And (Not blnIsBankHoliday) And (blnShadeWeekends) And (blnIncWorkingDays) And (blnIncBankHoliday) Then
                'Event + Not Working Day + Weekend + Show Weekends + Inc. Working Days Only + Inc. Bank Holidays
                intDefinedColourStyle = 1
              End If

              If (blnIsBankHoliday) And (blnIsWeekend) And (blnShadeWeekends) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                'Event + Not Working Day + Bank Holiday + Weekend + Show Weekends + Inc. Working Days Only + Not Inc. Bank Holidays
                intDefinedColourStyle = 1
              End If
              
              If (blnIsBankHoliday) And (Not blnIncBankHoliday) And (Not blnShadeBankHolidays) And (Not ((blnIsWeekend) And (blnShadeWeekends))) Then
                'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only + Not Shade Bank Holidays)))
                intDefinedColourStyle = 0
              ElseIf (blnIsBankHoliday) And (Not blnIncBankHoliday) And ((blnShadeBankHolidays) Or ((blnIsWeekend) And (blnShadeWeekends))) Then
                'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only + Shade Bank Holidays)))
                intDefinedColourStyle = 1
              End If

            End If
            
          Else
            'Not Event
            intDefinedColourStyle = 0
            
            If (blnIsWeekend) And (blnShadeWeekends) Then
              'Not Event + Weekend + Show Weekends
              intDefinedColourStyle = 1
            End If
              
            If (blnIsBankHoliday) And (blnShadeBankHolidays) Then
              'Not Event + Bank Holiday + Show Bank Holidays
              intDefinedColourStyle = 1
            End If
                           
          End If
            
            
          Select Case intDefinedColourStyle
            Case 0:
              'Show the default colour
              varTempArray(8, intCount) = HexValue(mlngBC_DataOutput)
              varTempArray(9, intCount) = HexValue(mlngBC_DataOutput)
              strBackColour = varTempArray(8, intCount)
              strForeColour = varTempArray(9, intCount)
              blnShowCaption = False
              
            Case 1:
              'Show the Weekend/Bank Holiday colour
              varTempArray(8, intCount) = HexValue(lblWeekend.BackColor)
              varTempArray(9, intCount) = HexValue(lblWeekend.BackColor)
              strBackColour = varTempArray(8, intCount)
              strForeColour = varTempArray(9, intCount)
              blnShowCaption = False
            
            Case 2:
              'Show the colour from the Event Key!
              varTempArray(8, intCount) = strColour
              varTempArray(9, intCount) = HexValue(lblCalDates(0).ForeColor)
              strBackColour = varTempArray(8, intCount)
              strForeColour = varTempArray(9, intCount)
              blnShowCaption = True
            
          End Select
          
          'set key character OR NOT.
          'TM17122003 Faults 7818 & 7819 fixed.
          'if the caption is not to be shown then set the caption to null string
          'rather than hide by making the forecolor the same as the backcolor.
          If ((chkCaptions.Value = vbChecked) And (blnShowCaption)) Then
            varTempArray(9, intCount) = HexValue(lblCalDates(0).ForeColor)
            strForeColour = varTempArray(9, intCount)
            strCaption = varTempArray(7, intCount)
          Else
            varTempArray(9, intCount) = varTempArray(8, intCount)
            strForeColour = varTempArray(9, intCount)
            strCaption = vbNullString
          End If
      
'          strCaption = varTempArray(7, intCount)
  
          If intSessionCount = 1 Then
            
            If blnHasEvent Or ((blnIsBankHoliday) And (blnShadeBankHolidays)) Then
              mobjOutput.AddStyle "", CLng((intCount + 1) / 2), CLng(intBaseCount * 2), _
                                    CLng((intCount + 1) / 2), CLng(intBaseCount * 2), _
                                    CLng(varTempArray(8, intCount)), CLng(varTempArray(9, intCount)), False, False, True
            End If
            
            AddDataToArray CInt(intBaseCount * 2), CInt((intCount + 1) / 2), strCaption
            
          ElseIf intSessionCount = 2 Then
            
            If blnHasEvent Or ((blnIsBankHoliday) And (blnShadeBankHolidays)) Then
              mobjOutput.AddStyle "", CLng(intCount / 2), CLng((intBaseCount * 2) + 1), _
                                    CLng(intCount / 2), CLng((intBaseCount * 2) + 1), _
                                    CLng(varTempArray(8, intCount)), CLng(varTempArray(9, intCount)), False, False, True
            End If
            
            AddDataToArray CInt((intBaseCount * 2) + 1), CInt(intCount / 2), strCaption
            
            intSessionCount = 0
            
          End If
        Else
          If intSessionCount = 2 Then
            intSessionCount = 0
          End If
        End If
      End If
      
    Next intCount
  Next intBaseCount
  
  OutputArray_RefreshDateSpecifics = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_RefreshDateSpecifics = False
  GoTo TidyUpAndExit

End Function

Private Function OutputArray_AddDates() As Boolean

  On Error GoTo ErrorTrap
  
  Dim intControlCount As Integer
  Dim intDateCount As Integer
  
'  mintFirstDayOfMonth_Output = Weekday("01/" & plngMonth & "/" + CStr(plngYear), vbSunday)
'
'  mintDaysInMonth_Output = DaysInMonth(CDate("01/" & plngMonth & "/" + CStr(plngYear)))
  
  AddDataToArray 1, 0, ""
  
  For intControlCount = 1 To DAY_CONTROL_COUNT Step 1
    
    If (intControlCount >= mintFirstDayOfMonth_Output) And _
      (intControlCount < (mintFirstDayOfMonth_Output + mintDaysInMonth_Output)) Then
      intDateCount = intDateCount + 1
      AddDataToArray 1, intControlCount, CStr(intDateCount)
    Else
      'Add a blank date box
      AddDataToArray 1, intControlCount, ""
    End If
  
  Next intControlCount
  
'  'Define the current visible Start and End Dates.
'  mdtVisibleStartDate_Output = CDate("01/" & CStr(plngMonth) & "/" + CStr(plngYear))
'  mdtVisibleEndDate_Output = CDate(mintDaysInMonth_Output & "/" & CStr(plngMonth) & "/" + CStr(plngYear))
  
  OutputArray_AddDates = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_AddDates = False
  GoTo TidyUpAndExit
  
End Function

Private Function OutputArray_AddCalendar() As Boolean

  On Error GoTo ErrorTrap
  
  Dim iNewIndex As Integer
  Dim intDateValue As Integer
  Dim intDateCount As Integer
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim intCurrentIndex As Integer
  Dim intControlCount As Integer
  Dim intSessionCount As Integer
  Dim intNextIndex As Integer
  
  Dim dtLabelsDate As Date
  
  Dim lngBaseID As Long
  Dim strDate As String
  Dim strSession As String
  Dim strIsBankHoliday As String
  Dim strIsWeekend As String
  Dim strIsWorkingDay As String
  Dim intHasEvent As Integer
  Dim strCaption As String
  Dim strBackColour As String
  Dim strForeColour As String
  
  Dim strRegion As String
  Dim strWorkingPattern As String
  
  Dim varTempArray() As Variant
 
  Dim blnNewBaseRecord As Boolean
  Dim strTempRecordDesc As String
  Dim intDescEmpty As Integer
  Dim blnDescEmpty As Boolean

  Dim strBaseDescription1, strBaseDescription2, strBaseDescriptionExpr As String
  Dim iDecimals As Integer
  Dim strFormat As String

  intDateCount = 0
  mstrBaseRecDesc_Output = vbNullString
  mintBaseRecordCount_Output = 0
  mstrBaseRecDesc = vbNullString
  mintCurrentBaseIndex_Output = 0
  blnNewBaseRecord = True
  mintRangeStartIndex_Output = 0
  mintRangeEndIndex_Output = 0
  mlngCurrentRecordID = -1
  
  mintType_BaseDesc1 = -1
  mintType_BaseDesc2 = -1
  mintType_BaseDescExpr = -1

  With mrsBase
    If .BOF And .EOF Then
      OutputArray_AddCalendar = False
      GoTo TidyUpAndExit
    End If
    
    mintBaseCount_Output = .RecordCount
    ReDim mavOutputDateIndex(2, 0)
    
    'get the datatype/properties for the desc1 column
    If (mlngDescription1ID > 0) Then
      If datGeneral.DoesColumnUseSeparators(mlngDescription1ID) Then
        mintType_BaseDesc1 = 3
        iDecimals = datGeneral.GetDecimalsSize(mlngDescription1ID)
        mstrFormat_BaseDesc1 = "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "#"), "")
      ElseIf datGeneral.BitColumn("C", mlngBaseTableID, mlngDescription1ID) Then
        mintType_BaseDesc1 = 2
      ElseIf datGeneral.DateColumn("C", mlngBaseTableID, mlngDescription1ID) Then
        mintType_BaseDesc1 = 1
      Else
        mintType_BaseDesc1 = 0
      End If
    End If
    'get the datatype/properties for the desc2 column
    If (mlngDescription2ID > 0) Then
      If datGeneral.DoesColumnUseSeparators(mlngDescription2ID) Then
        mintType_BaseDesc2 = 3
        iDecimals = datGeneral.GetDecimalsSize(mlngDescription2ID)
        mstrFormat_BaseDesc2 = "#,0" & IIf(iDecimals > 0, "." & String(iDecimals, "#"), "")
      ElseIf datGeneral.BitColumn("C", mlngBaseTableID, mlngDescription2ID) Then
        mintType_BaseDesc2 = 2
      ElseIf datGeneral.DateColumn("C", mlngBaseTableID, mlngDescription2ID) Then
        mintType_BaseDesc2 = 1
      Else
        mintType_BaseDesc2 = 0
      End If
    End If
    'get the datatype/properties for the descexpr column
    If (mlngDescriptionExprID > 0) Then
      If datGeneral.BitColumn("X", mlngBaseTableID, mlngDescriptionExprID) Then
        mintType_BaseDescExpr = 2
      ElseIf datGeneral.DateColumn("X", mlngBaseTableID, mlngDescriptionExprID) Then
        mintType_BaseDescExpr = 1
      Else
        mintType_BaseDescExpr = 0
      End If
    End If
    
    .MoveFirst
    Do While Not .EOF
    
      ' Get base description 1
      If Not IsNull(.Fields("Description1").Value) Then
        Select Case mintType_BaseDesc1
          Case 3: strBaseDescription1 = Format(.Fields("Description1").Value, mstrFormat_BaseDesc1)
          Case 2: strBaseDescription1 = IIf(.Fields("Description1").Value, "Y", "N")
          Case 1: strBaseDescription1 = Format(.Fields("Description1").Value, mstrDateFormat)
          Case 0: strBaseDescription1 = .Fields("Description1").Value
        End Select
      Else
        strBaseDescription1 = vbNullString
      End If
      ' Get base description 2
      If Not IsNull(.Fields("Description2").Value) Then
        Select Case mintType_BaseDesc2
          Case 3: strBaseDescription2 = Format(.Fields("Description2").Value, mstrFormat_BaseDesc2)
          Case 2: strBaseDescription2 = IIf(.Fields("Description2").Value, "Y", "N")
          Case 1: strBaseDescription2 = Format(.Fields("Description2").Value, mstrDateFormat)
          Case 0: strBaseDescription2 = .Fields("Description2").Value
        End Select
      Else
        strBaseDescription2 = vbNullString
      End If
      ' Get base description expression
      If Not IsNull(.Fields("DescriptionExpr").Value) Then
        Select Case mintType_BaseDescExpr
          Case 2: strBaseDescriptionExpr = IIf(.Fields("DescriptionExpr").Value, "Y", "N")
          Case 1: strBaseDescriptionExpr = Format(.Fields("DescriptionExpr").Value, mstrDateFormat)
          Case 0: strBaseDescriptionExpr = .Fields("DescriptionExpr").Value
        End Select
      Else
        strBaseDescriptionExpr = vbNullString
      End If
      
      strTempRecordDesc = strBaseDescription1
      strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescription2) > 0), mstrDescriptionSeparator, "") & strBaseDescription2
      strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescriptionExpr) > 0), mstrDescriptionSeparator, "") & strBaseDescriptionExpr

      blnDescEmpty = (strTempRecordDesc = vbNullString)
      If blnDescEmpty Then
        intDescEmpty = intDescEmpty + 1
      Else
        intDescEmpty = 0
      End If

      If mblnGroupByDesc Then
        If ((strTempRecordDesc) <> mstrBaseRecDesc) Or (blnDescEmpty And Int(intDescEmpty = 1)) Then
          blnNewBaseRecord = True
          blnDescEmpty = False
                              
          mstrBaseRecDesc = strTempRecordDesc
          
          If Len(Trim(mstrStaticRegionColumn)) > 0 Then
            mstrCurrentBaseRegion = IIf(IsNull(.Fields("Region").Value), "", .Fields("Region").Value)
          End If
          mintBaseRecordCount_Output = mintBaseRecordCount_Output + 1
        End If
        mlngCurrentRecordID = .Fields(mstrBaseIDColumn).Value
        
      Else
        If .Fields(mstrBaseIDColumn).Value <> mlngCurrentRecordID Then
          blnNewBaseRecord = True
          
          mstrBaseRecDesc = strTempRecordDesc
         
          mlngCurrentRecordID = .Fields(mstrBaseIDColumn).Value
          If Len(Trim(mstrStaticRegionColumn)) > 0 Then
            mstrCurrentBaseRegion = IIf(IsNull(.Fields("Region").Value), "", .Fields("Region").Value)
          End If
          mintBaseRecordCount_Output = mintBaseRecordCount_Output + 1
        End If
        
      End If
      
      intSessionCount = 0
      mintCurrentBaseIndex_Output = mintCurrentBaseIndex_Output + 1
      
      ReDim Preserve mavOutputDateIndex(2, mintBaseRecordCount_Output)
      ReDim varTempArray(9, 74)
      
      If blnNewBaseRecord Then
      
        For iCount2 = 1 To 74 Step 1
          intSessionCount = intSessionCount + 1
          
          If intSessionCount = 1 Then
            intControlCount = intControlCount + 1
          End If
    
          lngBaseID = CLng(.Fields(mstrBaseIDColumn).Value)
          
          If (intControlCount >= mintFirstDayOfMonth_Output) And (intControlCount < (mintFirstDayOfMonth_Output + mintDaysInMonth_Output)) Then
            strSession = IIf(intSessionCount = 2, " PM", " AM")
            If Trim(strSession) = "AM" Then
              intDateCount = intDateCount + 1
            End If
'            dtLabelsDate = CDate(intDateCount & "/" & mlngMonth_Output & "/" & CStr(mlngYear_Output))
            dtLabelsDate = DateAdd("d", CDbl(intDateCount - 1), mdtVisibleStartDate_Output)
            
            'calculate the indices of the out of report range bounaries.
            If dtLabelsDate < mdtReportStartDate Then
              mintRangeStartIndex_Output = intControlCount
            End If
            If dtLabelsDate = mdtReportEndDate Then
              mintRangeEndIndex_Output = intControlCount + 1
            End If
                        
            strDate = Format(dtLabelsDate, CALREP_DATEFORMAT)
            strBackColour = HexValue(lblCalDates(0).BackColor)
            strCaption = vbNullString
            
          Else
            strDate = "  /  /    "
            strSession = vbNullString
            strBackColour = HexValue(lblDisabled.BackColor)
            strCaption = vbNullString
            
          End If
            
          If Trim(strSession) <> vbNullString Then
            strIsBankHoliday = IIf(IsBankHoliday(dtLabelsDate, lngBaseID, strRegion), "1", "0")
  
            'flag if the date is a weekend
            strIsWeekend = IIf(IsWeekend(dtLabelsDate), "1", "0")
  
            'flag if the date & session is in the current personnel's working pattern.
            strIsWorkingDay = IIf(IsWorkingDay(dtLabelsDate, lngBaseID, Trim(strSession), strWorkingPattern), "1", "0")
  
          Else
            strIsBankHoliday = "0"
            strIsWeekend = "0"
            strIsWorkingDay = "0"
            
          End If
          
          'Add values to Date Index array
          varTempArray(0, iCount2) = lngBaseID
          varTempArray(1, iCount2) = strDate
          varTempArray(2, iCount2) = strSession
          varTempArray(3, iCount2) = strIsBankHoliday
          varTempArray(4, iCount2) = strIsWeekend
          varTempArray(5, iCount2) = strIsWorkingDay
          varTempArray(6, iCount2) = 0
          varTempArray(7, iCount2) = strCaption
          varTempArray(8, iCount2) = strBackColour
          varTempArray(9, iCount2) = HexValue(vbBlack)
          
          If intSessionCount = 2 Then
            intSessionCount = 0
          End If
          
        Next iCount2
        
        mavOutputDateIndex(0, mintBaseRecordCount_Output) = .Fields(mstrBaseIDColumn).Value
        mavOutputDateIndex(1, mintBaseRecordCount_Output) = mstrBaseRecDesc
        mavOutputDateIndex(2, mintBaseRecordCount_Output) = varTempArray()
      
      End If
      
      mcolBaseDescIndex_Output.Add mintBaseRecordCount_Output, CStr(.Fields(mstrBaseIDColumn).Value)
        
      ReDim varTempArray(9, 0)
      
      intControlCount = 0
      intDateCount = 0
    
      blnNewBaseRecord = False
      
      .MoveNext
    Loop
    
  End With
  
  OutputArray_AddCalendar = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_AddCalendar = False
  GoTo TidyUpAndExit
  
End Function

Private Function OutputArray_AddEvents() As Boolean

  On Error GoTo ErrorTrap
  
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim lngCurrentBaseID As Long
  Dim intBaseRecordIndex As Integer
  
  Dim fOK As Boolean
  
  Dim sSQL As String

  fOK = True
  
  With mrsEvents
  
    ' If there are no event records, skip this bit
    ' this bit (but still show the form)
    If .BOF And .EOF Then
      Exit Function
    End If
      
    .MoveFirst
    ' Loop through the events recordset
    Do Until .EOF

      lngCurrentBaseID = mrsEvents.Fields(mstrBaseIDColumn)
      
      intBaseRecordIndex = mcolBaseDescIndex_Output.Item(CStr(lngCurrentBaseID))
      
      ' Load each event record data into variables
      ' (has to be done because start/end dates may be modified by code to fill grid correctly)
      mstrCurrentEventKey = .Fields(mstrEventIDColumn)
     
      mstrEventLegend_Output = IIf(IsNull(.Fields("Legend")), "", Left(.Fields("Legend").Value, 2))
     
      '****************************************************************************
      mdtEventStartDate_Output = Format(.Fields("StartDate"), mstrDateFormat)
    
      If IsNull(.Fields("EndDate")) Then
        mdtEventEndDate_Output = mdtEventStartDate_Output
      Else
        mdtEventEndDate_Output = Format(.Fields("EndDate"), mstrDateFormat)
      End If
    
      If IsNull(.Fields("StartSession")) And IsNull(.Fields("EndSession")) Then
        mstrEventStartSession_Output = "AM"
        mstrEventEndSession_Output = "PM"
      ElseIf IsNull(.Fields("EndSession")) Then
        mstrEventEndSession_Output = mstrEventStartSession_Output
      Else
        mstrEventStartSession_Output = UCase(.Fields("StartSession"))
        mstrEventEndSession_Output = UCase(.Fields("EndSession"))
      End If
      
      mstrDuration_Output = .Fields("Duration")
      
      '****************************************************************************
      
      ' If the event start date is after the event end date, ignore the record
      If (mdtEventStartDate_Output > mdtEventEndDate_Output) Then
      
      ' if the event is totally before the currently viewed timespan then do nothing
      ElseIf (mdtEventStartDate_Output < mdtVisibleStartDate_Output) _
        And (mdtEventEndDate_Output < mdtVisibleStartDate_Output) Then
      
      ' if the event is totally after the currently viewed timespan then do nothing
      ElseIf (mdtEventStartDate_Output > mdtVisibleEndDate_Output) _
        And (mdtEventEndDate_Output > mdtVisibleEndDate_Output) Then
      
      ' if the event starts before currently viewed timespan, but ends in the timspan then
      ElseIf (mdtEventStartDate_Output < mdtVisibleStartDate_Output) _
        And (mdtEventEndDate_Output <= mdtVisibleEndDate_Output) Then
        
        mdtEventStartDate_Output = mdtVisibleStartDate_Output
        mstrEventStartSession_Output = "AM"
        
        lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
        lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))
  
        fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)
        
      ' if the event starts in the currently viewed timespan, but ends after it then
      ElseIf (mdtEventStartDate_Output >= mdtVisibleStartDate_Output) _
        And (mdtEventEndDate_Output > mdtVisibleEndDate_Output) Then
        
        mdtEventEndDate_Output = mdtVisibleEndDate_Output
        mstrEventEndSession_Output = "PM"
        
        lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
        lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))
  
        fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)
  
      ' if the event is enclosed within viewed timespan, and months are equal then
      ElseIf (mdtEventStartDate_Output >= mdtVisibleStartDate_Output) _
        And (mdtEventEndDate_Output <= mdtVisibleEndDate_Output) _
        And (Month(mdtEventStartDate_Output) = Month(mdtEventEndDate_Output)) Then
  
        lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
        lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))
  
        fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)
      
      ' if the event starts before the the viewed timespan and ends after the viewed timespan then
      ElseIf (mdtEventStartDate_Output < mdtVisibleStartDate_Output) _
        And (mdtEventEndDate_Output > mdtVisibleEndDate_Output) Then
        
        mdtEventStartDate_Output = mdtVisibleStartDate_Output
        mstrEventStartSession_Output = "AM"
        
        mdtEventEndDate_Output = mdtVisibleEndDate_Output
        mstrEventEndSession_Output = "PM"
                
        lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
        lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))
  
        fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)
            
      End If
      
      If fOK = False Then
        Exit Do
      End If
    
      .MoveNext
    Loop
  End With
  
  If fOK = False Then
    COAMsgBox "An Error Has Occurred Whilst Filling The Cal Labels:" & vbNewLine & Err.Number & " - " & Err.Description
  End If
  
  OutputArray_AddEvents = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_AddEvents = False
  GoTo TidyUpAndExit
  
End Function

Private Function OutputArray_FillEvents(plngCalDatIndex As Integer, plngStart As Long, plngEnd As Long) As Boolean

  ' This function actually fills the cal boxes between the indexes specified
  ' according to the options selected by the user.
  
  On Error GoTo ErrorTrap
  
  Dim colEvents As clsCalendarEvents
  
  Dim intCount As Integer
  
  Dim strCurrentRegion_BD As String
  Dim strCurrentWorkingPattern_BD As String
  
  Dim varTempArray() As Variant
  
  Dim intStartCount As Integer
  Dim intEndCount As Integer
  
  varTempArray = mavOutputDateIndex(2, plngCalDatIndex)
  
  ' Loop through the indexes as specified.
  For intCount = plngStart To plngEnd Step 1
    
    If varTempArray(6, intCount) = 0 Then
      'Date & Session clear
      varTempArray(6, intCount) = 1
      varTempArray(7, intCount) = mstrEventLegend_Output
      varTempArray(8, intCount) = GetLegendColour(mstrCurrentEventKey)
      varTempArray(9, intCount) = HexValue(vbBlack)
     
    Else
      'Date & Session already has an event, set it as Multiple.
      varTempArray(6, intCount) = 2
      varTempArray(7, intCount) = "."
      varTempArray(8, intCount) = HexValue(vbWhite)
      varTempArray(9, intCount) = HexValue(vbBlack)
      
    End If
    
  Next intCount
  
  mavOutputDateIndex(2, plngCalDatIndex) = varTempArray()

  OutputArray_FillEvents = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  OutputArray_FillEvents = False
  GoTo TidyUpAndExit
  
End Function

Public Property Let PersonnelBase(pblnPersonnelBase As Boolean)
  mblnPersonnelBase = pblnPersonnelBase
End Property

Private Function Get_HistoricWorkingPatterns()

  On Error GoTo ErrorTrap
 
  Dim rsCC As ADODB.Recordset   'career change data for base records
  Dim colWorkingPatterns As clsCalendarEvents
  
  Dim strSQLCC As String    'sql for retieving career change data
  
  Dim dtStartDate As Date
  Dim dtEndDate As Date
  
  Dim avCareerRanges() As String
  Dim intNextIndex As Integer
  
  Dim blnNewBaseRecord As Boolean
  Dim lngBaseRecordID As Long
  
  Dim intCount As Integer
  
  ReDim avCareerRanges(4, 0)
  
  strSQLCC = vbNullString
  strSQLCC = strSQLCC & "SELECT " & vbNewLine
  strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngBaseTableID & "," & vbNewLine
  strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & ", " & vbNewLine
  strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & ", " & vbNewLine
  strSQLCC = strSQLCC & "     (SELECT COUNT(B.ID) FROM " & gsPersonnelHWorkingPatternTableRealSource & " B WHERE B.ID_" & mlngBaseTableID & " = " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngBaseTableID & " AND B." & gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL) AS 'CareerChanges' " & vbNewLine
  strSQLCC = strSQLCC & "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & vbNewLine
  If Len(Trim(mstrSQLIDs)) > 0 Then
    strSQLCC = strSQLCC & "WHERE " & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngBaseTableID & " IN (" & mstrSQLIDs & ") " & vbNewLine
    strSQLCC = strSQLCC & " AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL " & vbNewLine
  Else
    strSQLCC = strSQLCC & "WHERE " & vbNewLine
    strSQLCC = strSQLCC & "      " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL " & vbNewLine
  End If
  strSQLCC = strSQLCC & "ORDER BY "
  strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngBaseTableID & ", "
  strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " "
  
  Set rsCC = datGeneral.GetRecords(strSQLCC)
  
  lngBaseRecordID = -1
  blnNewBaseRecord = False
  
  '******************************************************************************
  'Create an array containing the ranges of career change period
  With rsCC
  
    If Not (.BOF And .EOF) Then
        
      Do While Not .EOF
        intNextIndex = UBound(avCareerRanges, 2) + 1
        ReDim Preserve avCareerRanges(4, intNextIndex)
        
        If lngBaseRecordID <> .Fields("ID_" & CStr(mlngBaseTableID)).Value Then
          lngBaseRecordID = .Fields("ID_" & CStr(mlngBaseTableID)).Value
          blnNewBaseRecord = True
          dtStartDate = .Fields(gsPersonnelHWorkingPatternDateColumnName).Value
          
          avCareerRanges(0, intNextIndex) = lngBaseRecordID     'BaseRecordID
          avCareerRanges(1, intNextIndex) = dtStartDate         'Start Date
          avCareerRanges(2, intNextIndex) = ""                  'End Date
          avCareerRanges(3, intNextIndex) = IIf(IsNull(.Fields(gsPersonnelHWorkingPatternColumnName).Value), "", .Fields(gsPersonnelHWorkingPatternColumnName).Value)       'Working Pattern???
          avCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value   'Career Change Count
          
        Else
          dtStartDate = .Fields(gsPersonnelHWorkingPatternDateColumnName).Value
          dtEndDate = dtStartDate
          avCareerRanges(2, intNextIndex - 1) = dtEndDate       'End Date
          
          avCareerRanges(0, intNextIndex) = lngBaseRecordID     'BaseRecordID
          avCareerRanges(1, intNextIndex) = dtStartDate         'Start Date
          avCareerRanges(2, intNextIndex) = ""                  'End Date
          avCareerRanges(3, intNextIndex) = IIf(IsNull(.Fields(gsPersonnelHWorkingPatternColumnName).Value), "", .Fields(gsPersonnelHWorkingPatternColumnName).Value)       'Working Pattern???
          avCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value   'Career Change Count
          
        End If
        
        blnNewBaseRecord = False
        .MoveNext
      Loop
    
    Else
      Get_HistoricWorkingPatterns = True
      GoTo TidyUpAndExit

    End If
  
  End With
  '******************************************************************************
 
  lngBaseRecordID = -1
  blnNewBaseRecord = False
  
  '##############################################################################
  'populate collections with new data
  
  For intCount = 1 To UBound(avCareerRanges, 2) Step 1
    If lngBaseRecordID <> CLng(avCareerRanges(0, intCount)) Then
       If Not (colWorkingPatterns Is Nothing) Then
         mcolHistoricWorkingPatterns.Add colWorkingPatterns, CStr(lngBaseRecordID)
         Set colWorkingPatterns = Nothing
       End If
       Set colWorkingPatterns = New clsCalendarEvents
      
      lngBaseRecordID = avCareerRanges(0, intCount)
      blnNewBaseRecord = True
    End If
     
    colWorkingPatterns.Add CStr(colWorkingPatterns.Count), CStr(lngBaseRecordID), _
                          , , CInt(avCareerRanges(4, intCount)), , avCareerRanges(1, intCount), , , , _
                          avCareerRanges(2, intCount), , , , , , , , _
                          , , , , , , , , , , , , , avCareerRanges(3, intCount)

    If (intCount = UBound(avCareerRanges, 2)) And Not (colWorkingPatterns Is Nothing) Then
      mcolHistoricWorkingPatterns.Add colWorkingPatterns, CStr(lngBaseRecordID)
      Set colWorkingPatterns = Nothing
    End If

    blnNewBaseRecord = False
  Next intCount

  '##############################################################################
  
  Get_HistoricWorkingPatterns = True

TidyUpAndExit:
  Set rsCC = Nothing
  Set colWorkingPatterns = Nothing
  Exit Function

ErrorTrap:
  Get_HistoricWorkingPatterns = False
  GoTo TidyUpAndExit
  
End Function

Private Function Get_HistoricBankHolidays()

  On Error GoTo ErrorTrap
  
  Dim rsCC As ADODB.Recordset   'career change data for base records
  Dim rsPersonnelBHols As ADODB.Recordset
  Dim colBankHolidays As clsBankHolidays
  
  Dim strSQLCC As String    'sql for retieving career change data
  Dim strSQLAllBHols As String
  Dim strSQLSelect As String
  Dim strSQLWhere As String
  Dim strSQLDateRegion As String
  Dim strSQLOrder As String
  
  Dim dtStartDate As Date
  Dim dtEndDate As Date
  
  Dim intNextIndex As Integer
  
  Dim blnNewBaseRecord As Boolean
  Dim lngBaseRecordID As Long
  
  Dim lng100Counter As Long
  Dim lngBaseRowCount As Long
  Dim lngMainBaseCounter As Long
  Dim lngTotalCareerChanges As Long
    
  Dim intCount As Integer
  Dim intBHolCount As Integer
  Dim fFinalCareerChange As Boolean
  
  Dim iViewCount As Integer
  Dim strTempRealSource As String
  Dim lngCount As Long
  
  ReDim mavCareerRanges(4, 0)
  
  strSQLCC = "SELECT " & vbNewLine _
    & "     " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngBaseTableID & "," & vbNewLine _
    & "     " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & ", " & vbNewLine _
    & "     " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & ", " & vbNewLine _
    & "     (SELECT COUNT(B.ID) FROM " & gsPersonnelHRegionTableRealSource & " B WHERE B.ID_" & mlngBaseTableID & " = " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngBaseTableID & " AND B." & gsPersonnelHRegionDateColumnName & " IS NOT NULL) AS 'CareerChanges' " & vbNewLine _
    & "FROM " & gsPersonnelHRegionTableRealSource & " " & vbNewLine
  
  If LenB(Trim(mstrSQLIDs)) > 0 Then
    strSQLCC = strSQLCC & "WHERE " & vbNewLine _
      & "     " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngBaseTableID & " IN (" & mstrSQLIDs & ") " & vbNewLine _
      & " AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " IS NOT NULL " & vbNewLine
  Else
    strSQLCC = strSQLCC & "WHERE " & vbNewLine _
      & "      " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " IS NOT NULL " & vbNewLine
  End If
  
  strSQLCC = strSQLCC & "ORDER BY " _
    & "     " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngBaseTableID & ", " _
    & "     " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " "
 
  Set rsCC = datGeneral.GetRecords(strSQLCC)
 
  lngBaseRecordID = -1
  blnNewBaseRecord = False
  lng100Counter = 0
  lngMainBaseCounter = 0
 
  '******************************************************************************
  'Create an array containing the ranges of career change period
  
  With rsCC
 
    If Not (.BOF And .EOF) Then
       
      Do While Not .EOF
        intNextIndex = UBound(mavCareerRanges, 2) + 1
        ReDim Preserve mavCareerRanges(4, intNextIndex)
        
        If lngBaseRecordID <> .Fields("ID_" & CStr(mlngBaseTableID)).Value Then
          lngBaseRecordID = .Fields("ID_" & CStr(mlngBaseTableID)).Value
          blnNewBaseRecord = True
          
          lngBaseRowCount = lngBaseRowCount + 1
'          dtStartDate = Format(.Fields(gsPersonnelHRegionDateColumnName).Value, "mm/dd/yyyy")
          dtStartDate = .Fields(gsPersonnelHRegionDateColumnName).Value
          
          mavCareerRanges(0, intNextIndex) = lngBaseRecordID     'BaseRecordID
          mavCareerRanges(1, intNextIndex) = dtStartDate         'Start Date
          mavCareerRanges(2, intNextIndex) = ""                  'End Date
          mavCareerRanges(3, intNextIndex) = IIf(IsNull(.Fields(gsPersonnelHRegionColumnName).Value), "", .Fields(gsPersonnelHRegionColumnName).Value)       'Region
          mavCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value   'Career Change Count
          
        Else
'          dtStartDate = Format(.Fields(gsPersonnelHRegionDateColumnName).Value, "mm/dd/yyyy")
          dtStartDate = .Fields(gsPersonnelHRegionDateColumnName).Value

          dtEndDate = dtStartDate
          mavCareerRanges(2, intNextIndex - 1) = dtEndDate       'End Date
          
          mavCareerRanges(0, intNextIndex) = lngBaseRecordID     'BaseRecordID
          mavCareerRanges(1, intNextIndex) = dtStartDate         'Start Date
          mavCareerRanges(2, intNextIndex) = ""                  'End Date
          mavCareerRanges(3, intNextIndex) = IIf(IsNull(.Fields(gsPersonnelHRegionColumnName).Value), "", .Fields(gsPersonnelHRegionColumnName).Value)       'Region
          mavCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value   'Career Change Count
          
        End If
        
        blnNewBaseRecord = False
        .MoveNext
      Loop
  
    Else
      Get_HistoricBankHolidays = True
      GoTo TidyUpAndExit
  
    End If
 
  End With
  
  lngTotalCareerChanges = UBound(mavCareerRanges, 2)
  
  '******************************************************************************
  lngBaseRecordID = -1
  blnNewBaseRecord = False
   
  'Fault 12358
  '------------------------------------------------------------------------------
  'Create and execute sql strings in batches of 100 base records, that ultimately return the bank holidays
  'for all the selcted base table records.
  
  strSQLAllBHols = vbNullString
  strSQLSelect = vbNullString
  strSQLWhere = vbNullString
  strSQLDateRegion = vbNullString
   
  For intCount = 1 To UBound(mavCareerRanges, 2) Step 1
   
    If lngBaseRecordID <> mavCareerRanges(0, intCount) Then
      lngBaseRecordID = mavCareerRanges(0, intCount)
      blnNewBaseRecord = True
      lng100Counter = lng100Counter + 1
      lngMainBaseCounter = lngMainBaseCounter + 1
      strSQLSelect = vbNullString
      strSQLDateRegion = "         ( " & vbNewLine
     
      strSQLWhere = "WHERE " & vbNewLine
      
      intBHolCount = 0
    End If
     
    intBHolCount = intBHolCount + 1
    
    strSQLSelect = vbNewLine & "SELECT  '" & mavCareerRanges(0, intCount) & "' AS 'ID' , " & vbNewLine
    strSQLSelect = strSQLSelect & "       " & mstrSQLSelect_RegInfoRegion & " AS 'Region', " & vbNewLine
    strSQLSelect = strSQLSelect & "       " & mstrSQLSelect_BankHolDate & " , " & vbNewLine
    strSQLSelect = strSQLSelect & "       " & mstrSQLSelect_BankHolDesc & " " & vbNewLine
    strSQLSelect = strSQLSelect & "FROM " & gsBHolRegionTableName & " " & vbNewLine
  
    For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
      '<REGIONAL CODE>
      If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
        strSQLSelect = strSQLSelect & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
        strSQLSelect = strSQLSelect & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
      End If
    Next lngCount
   
    strSQLSelect = strSQLSelect & "           INNER JOIN " & gsBHolTableRealSource & vbNewLine
    strSQLSelect = strSQLSelect & "           ON  " & gsBHolRegionTableName & ".ID = " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & vbNewLine
   
    If intBHolCount > 1 Then
      strSQLDateRegion = strSQLDateRegion & " OR " & vbNewLine
    End If
    
    fFinalCareerChange = (intBHolCount = CInt(mavCareerRanges(4, intCount)))
    
    If fFinalCareerChange Then
      strSQLDateRegion = strSQLDateRegion & "( " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & "(" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= CONVERT(datetime, '" & Replace(Format(mavCareerRanges(1, intCount), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')) " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & Replace(Format(mdtReportStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & Replace(Format(mdtReportEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & " AND " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & "(" & mstrSQLSelect_RegInfoRegion & " = '" & mavCareerRanges(3, intCount) & "') " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & ") " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & ") " & vbNewLine
    Else
      strSQLDateRegion = strSQLDateRegion & "( " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & "(" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= CONVERT(datetime, '" & Replace(Format(mavCareerRanges(1, intCount), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine & _
                                             " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " < CONVERT(datetime, '" & Replace(Format(mavCareerRanges(1, (intCount + 1)), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))) " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & Replace(Format(mdtReportStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & Replace(Format(mdtReportEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & " AND (" & mstrSQLSelect_RegInfoRegion & " = '" & mavCareerRanges(3, intCount) & "') " & vbNewLine
      strSQLDateRegion = strSQLDateRegion & ") " & vbNewLine
    End If
    
     
    If fFinalCareerChange Then
      strSQLAllBHols = strSQLAllBHols & strSQLSelect & vbNewLine
      strSQLAllBHols = strSQLAllBHols & strSQLWhere & vbNewLine
      strSQLAllBHols = strSQLAllBHols & strSQLDateRegion & vbNewLine
      strSQLAllBHols = strSQLAllBHols & " UNION "
      strSQLWhere = vbNullString
      strSQLDateRegion = vbNullString
    End If
     
    
    'Fault 12358
    'Send the query to SQL Server in batches of approximately 100, to avoid 256(260) Table/Views limit.
    'Do not split base records in to more than one batch!
    If ((lng100Counter = lngBaseRowCount) And fFinalCareerChange) _
      Or ((lng100Counter > 100) And fFinalCareerChange) _
      Or ((lngMainBaseCounter = lngBaseRowCount) And fFinalCareerChange) Then
      
      strSQLAllBHols = Left(strSQLAllBHols, (Len(strSQLAllBHols) - 8))
      strSQLOrder = " ORDER BY 'ID', 'Region' " & vbNewLine
      strSQLAllBHols = strSQLAllBHols & strSQLOrder
      
'Open App.Path & "\calrep.txt" For Output As #1
'Print #1, strSQLAllBHols
'Close #1

      Set rsPersonnelBHols = datGeneral.GetRecords(strSQLAllBHols)

      lngBaseRecordID = -1
      blnNewBaseRecord = False
      
      '##############################################################################
      'populate collections with new data
      With rsPersonnelBHols
     
        If Not (.BOF And .EOF) Then
           
          Do While Not .EOF
            If lngBaseRecordID <> CLng(.Fields("ID").Value) Then
              
              If Not (colBankHolidays Is Nothing) Then
                mcolHistoricBankHolidays.Add colBankHolidays, CStr(lngBaseRecordID)
                Set colBankHolidays = Nothing
              End If
              Set colBankHolidays = New clsBankHolidays
              
              lngBaseRecordID = CLng(.Fields("ID").Value)
              blnNewBaseRecord = True
            
            End If
           
            colBankHolidays.Add IIf(IsNull(.Fields("Region").Value), "", .Fields("Region").Value), _
                                 IIf(IsNull(.Fields(gsBHolDescriptionColumnName).Value), "", .Fields(gsBHolDescriptionColumnName).Value), _
                                 IIf(IsNull(.Fields(gsBHolDateColumnName).Value), "", .Fields(gsBHolDateColumnName).Value)
            
            blnNewBaseRecord = False
            
            .MoveNext
            
            If .EOF And Not (colBankHolidays Is Nothing) Then
              mcolHistoricBankHolidays.Add colBankHolidays, CStr(lngBaseRecordID)
              Set colBankHolidays = Nothing
            End If
                 
          Loop
   
        End If
     
      End With
     '##############################################################################

      'Reset SQL string variables ready for next batch to be created.
      strSQLAllBHols = vbNullString
      strSQLSelect = vbNullString
      strSQLWhere = vbNullString
      strSQLDateRegion = vbNullString
      lng100Counter = 0
      lngBaseRecordID = -1
    End If
    
    blnNewBaseRecord = False
     
  Next intCount

  Get_HistoricBankHolidays = True

TidyUpAndExit:
  Set rsCC = Nothing
  Set rsPersonnelBHols = Nothing
  Set colBankHolidays = Nothing
  Exit Function

ErrorTrap:
  Get_HistoricBankHolidays = False
  GoTo TidyUpAndExit
  
End Function

Private Function Get_StaticBankHolidays()

  On Error GoTo ErrorTrap
 
  Dim rsPersonnelBHols As ADODB.Recordset
  Dim colBankHolidays As clsBankHolidays
  
  Dim strSQLAllBHols As String
  
  Dim blnNewBaseRecord As Boolean
  Dim lngBaseRecordID As Long
  
  Dim intCount As Integer
  Dim intBHolCount As Integer
  Dim lngCount As Long
  Dim lngView As Long
  
  strSQLAllBHols = vbNullString
  strSQLAllBHols = strSQLAllBHols & "SELECT DISTINCT  [Base].ID, " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "                 [RegionInfo].Region, " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "                 [RegionInfo].Holiday_Date, " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "                 [RegionInfo].Description " & vbNewLine

  'gsBHolTableRealSource
  'gsBHolRegionTableName
  strSQLAllBHols = strSQLAllBHols & "FROM (SELECT DISTINCT " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "             " & mstrBaseTableName & ".ID AS 'ID', " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "             " & mstrSQLSelect_PersonnelStaticRegion & " AS 'Region' " & vbNewLine
  
  If mlngStaticRegionColumnID > 0 Then
    strSQLAllBHols = strSQLAllBHols & "      FROM " & mstrBaseTableName & vbNewLine
  Else
    strSQLAllBHols = strSQLAllBHols & "      FROM " & gsPersonnelTableName & vbNewLine
  End If
  
  For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
    If mlngStaticRegionColumnID > 0 Then
      If mvarTableViews(0, lngCount) = mlngBaseTableID Then
        strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
        strSQLAllBHols = strSQLAllBHols & "           ON  " & mstrBaseTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
      End If
    Else
      If mvarTableViews(0, lngCount) = mlngBaseTableID Then
        strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
        strSQLAllBHols = strSQLAllBHols & "           ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
      End If
    End If
  Next lngCount
  
  If Len(Trim(mstrSQLIDs)) > 0 Then
    strSQLAllBHols = strSQLAllBHols & "      WHERE " & mstrBaseTableName & ".ID IN (" & mstrSQLIDs & ") " & vbNewLine
  End If
  
  strSQLAllBHols = strSQLAllBHols & "      ) AS [Base] " & vbNewLine
  
  strSQLAllBHols = strSQLAllBHols & "   INNER JOIN " & vbNewLine
  
  strSQLAllBHols = strSQLAllBHols & "   (SELECT DISTINCT " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "   " & gsBHolRegionTableName & ".ID AS 'ID', " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "   " & mstrSQLSelect_RegInfoRegion & " AS 'Region', " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "   " & mstrSQLSelect_BankHolDate & ", " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "   " & mstrSQLSelect_BankHolDesc & " " & vbNewLine
  
  strSQLAllBHols = strSQLAllBHols & "      FROM " & gsBHolRegionTableName & vbNewLine
       
  For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
    '<REGIONAL CODE>
    If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
      strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
      strSQLAllBHols = strSQLAllBHols & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
    End If
  Next lngCount
       
  strSQLAllBHols = strSQLAllBHols & "           INNER JOIN " & gsBHolTableRealSource & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "           ON  " & gsBHolRegionTableName & ".ID = " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & vbNewLine
       
  If Len(Trim(mstrSQLIDs)) > 0 Then
    strSQLAllBHols = strSQLAllBHols & "     WHERE (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & Replace(Format(mdtReportStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "         AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & Replace(Format(mdtReportEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
  Else
    strSQLAllBHols = strSQLAllBHols & "     WHERE (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & Replace(Format(mdtReportStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "         AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & Replace(Format(mdtReportEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') " & vbNewLine
  End If
  
  strSQLAllBHols = strSQLAllBHols & "    ) AS [RegionInfo] " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "    ON [Base].Region = [RegionInfo].Region " & vbNewLine
  strSQLAllBHols = strSQLAllBHols & "ORDER BY [Base].ID " & vbNewLine

  Set rsPersonnelBHols = datGeneral.GetRecords(strSQLAllBHols)
  
  lngBaseRecordID = -1
  blnNewBaseRecord = False
  
  '##############################################################################
  'populate collections with new data
  With rsPersonnelBHols
  
    If Not (.BOF And .EOF) Then
        
      Do While Not .EOF
        If lngBaseRecordID <> .Fields("ID").Value Then
          
          If Not (colBankHolidays Is Nothing) Then
            mcolStaticBankHolidays.Add colBankHolidays, CStr(lngBaseRecordID)
            Set colBankHolidays = Nothing
          End If
          Set colBankHolidays = New clsBankHolidays
          
          lngBaseRecordID = .Fields("ID").Value
          blnNewBaseRecord = True

        End If
        
        colBankHolidays.Add IIf(IsNull(.Fields("Region").Value), "", .Fields("Region").Value), _
                            IIf(IsNull(.Fields(gsBHolDescriptionColumnName).Value), "", .Fields(gsBHolDescriptionColumnName).Value), _
                            IIf(IsNull(.Fields(gsBHolDateColumnName).Value), "", .Fields(gsBHolDateColumnName).Value)

        blnNewBaseRecord = False
        
        .MoveNext
        
        If .EOF And Not (colBankHolidays Is Nothing) Then
          mcolStaticBankHolidays.Add colBankHolidays, CStr(lngBaseRecordID)
          Set colBankHolidays = Nothing
        End If
        
      Loop
    
    End If
    
  End With
  '##############################################################################
  
  Get_StaticBankHolidays = True

TidyUpAndExit:
  Set rsPersonnelBHols = Nothing
  Set colBankHolidays = Nothing
  Exit Function

ErrorTrap:
  Get_StaticBankHolidays = False
  GoTo TidyUpAndExit
  
End Function

Private Function Get_StaticWorkingPatterns()

  On Error GoTo ErrorTrap
 
  Dim colWorkingPatterns As clsCalendarEvents
  
  Dim strSQLAllBHols As String
  
  Dim blnNewBaseRecord As Boolean
  Dim lngBaseRecordID As Long
  
  Dim intCount As Integer
  Dim intBHolCount As Integer
  
  lngBaseRecordID = -1
  blnNewBaseRecord = False
  
  '##############################################################################
  'populate collections with new data
  With mrsBase
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
        
      Do While Not .EOF
        If lngBaseRecordID <> .Fields(mstrBaseIDColumn).Value Then
          
          If Not (colWorkingPatterns Is Nothing) Then
            mcolStaticWorkingPatterns.Add colWorkingPatterns, CStr(lngBaseRecordID)
            Set colWorkingPatterns = Nothing
          End If
          Set colWorkingPatterns = New clsCalendarEvents
          
          lngBaseRecordID = .Fields(mstrBaseIDColumn).Value
          blnNewBaseRecord = True

        End If
        
        colWorkingPatterns.Add CStr(colWorkingPatterns.Count), CStr(lngBaseRecordID), _
                              , , , , , , , , _
                              , , , , , , , , _
                              , , , , , , , , , , , , , IIf(IsNull(.Fields(gsPersonnelWorkingPatternColumnName).Value), "              ", .Fields(gsPersonnelWorkingPatternColumnName).Value)

        blnNewBaseRecord = False
        
        .MoveNext
        
        If .EOF And Not (colWorkingPatterns Is Nothing) Then
          mcolStaticWorkingPatterns.Add colWorkingPatterns, CStr(lngBaseRecordID)
          Set colWorkingPatterns = Nothing
        End If
        
      Loop
    
    Else
      Get_StaticWorkingPatterns = True
      GoTo TidyUpAndExit

    End If
    
  End With
  '##############################################################################
  
  Get_StaticWorkingPatterns = True

TidyUpAndExit:
  Set colWorkingPatterns = Nothing
  Exit Function

ErrorTrap:
  Get_StaticWorkingPatterns = False
  GoTo TidyUpAndExit
  
End Function

Private Function PopulateMonthCombo() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  
  For iCount = 1 To 12
    cboMonth.AddItem StrConv(MonthName(iCount), vbProperCase)
    cboMonth.ItemData(cboMonth.NewIndex) = iCount
    If iCount = mlngMonth Then
      cboMonth.ListIndex = cboMonth.NewIndex
    End If
  Next iCount
  
  PopulateMonthCombo = True

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  PopulateMonthCombo = False
  GoTo TidyUpAndExit
  
End Function

Private Function SetYear() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  
  spnYear.Value = mlngYear
  
  SetYear = True

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  SetYear = False
  GoTo TidyUpAndExit
  
End Function

Private Function RefreshDateSpecifics() As Boolean

  'Shades the CalDates label controls depending on the Show Bank Hols & Show Weekend options.
  'Shows/Hides the legend character depending on the 'Show Captions' option.
  
  On Error GoTo ErrorTrap
  
  Dim intCount As Integer
  
  'following variables used to establish required back & fore color for the label
  Dim blnIsWeekend As Boolean
  Dim blnIsBankHoliday As Boolean
  Dim blnIsWorkingDay As Boolean
  Dim blnIncBankHoliday As Boolean
  Dim blnIncWorkingDays As Boolean
  Dim blnShadeBankHolidays As Boolean
  Dim blnShadeWeekends As Boolean
  Dim blnHasEvent As Boolean
  Dim blnShowCaption As Boolean
  Dim intDefinedColourStyle As Integer
  
  Dim strColour As String
  Dim intThisStartCount As Integer
  Dim intThisEndCount As Integer
  Dim intNextStartCount As Integer
  Dim intNext2StartCount As Integer
  Dim intIndexModulus As Integer
  Dim intCurrentStartCount As Integer
  Dim intCurrentEndCount As Integer
  Dim intBaseCount As Integer
  
  Dim lblTemp As VB.Label
  Dim lblTempNext As VB.Label
  Dim lblTempNext2 As VB.Label
  Dim lblTempPrev As VB.Label
  
  Dim strSession As String

  Dim blnNextHasEvent As Boolean
  Dim blnNext2HasEvent As Boolean
  Dim blnPrevHasEvent As Boolean

  Dim dtConvertedDate As Date
  
  Dim intSessionCount As Integer
  intSessionCount = 0
  
  If mintBaseRecordCount < 1 Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  
  blnIncBankHoliday = chkIncludeBHols.Value
  blnIncWorkingDays = chkIncludeWorkingDaysOnly.Value
  blnShadeBankHolidays = chkShadeBHols.Value
  blnShadeWeekends = chkShadeWeekends.Value
  
  For intBaseCount = 1 To mintBaseRecordCount Step 1
    
    For intCount = 1 To 74 Step 1
    
      intSessionCount = intSessionCount + 1
    
      Set lblTemp = ctlCalDates(intBaseCount).CalDate(intCount)
      
      With lblTemp

        If TagInfo_Get(.Tag, "DATE") = "  /  /    " Then
          .BackColor = lblDisabled.BackColor
          .Caption = ""
          
          If intSessionCount = 2 Then
            intSessionCount = 0
          End If

        Else
          dtConvertedDate = ConvertCalendarDateToDateFormat(TagInfo_Get(.Tag, "DATE"))
          If (dtConvertedDate >= mdtReportStartDate) And (dtConvertedDate <= mdtReportEndDate) Then
          
            blnIsBankHoliday = IIf(TagInfo_Get(.Tag, "BANK_HOLIDAY") = "1", True, False)
            blnIsWeekend = IIf(TagInfo_Get(.Tag, "WEEKEND") = "1", True, False)
            strColour = TagInfo_Get(.Tag, "COLOUR")
            blnHasEvent = IIf(CInt(TagInfo_Get(.Tag, "HAS_EVENT")) > 0, True, False)
            blnIsWorkingDay = IIf(TagInfo_Get(.Tag, "WORKING_DAY") = "1", True, False)
            
            intDefinedColourStyle = 0   'Default Colour
'            intDefinedColourStyle = 1   'Weekend/Bank Holiday Colour
'            intDefinedColourStyle = 2   'Event Key Colour
            
            If blnHasEvent Then
              'Event
              intDefinedColourStyle = 2
              
              If (blnIsWorkingDay) Then
                'Event + Working Day
                
                If (blnIsBankHoliday) And (Not blnIncBankHoliday) And (Not blnShadeBankHolidays) And (Not ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Working Day + ((Bank Holiday + Not Inc. Working Days Only + Not Shade Bank Holidays)))
                  intDefinedColourStyle = 0
                ElseIf (blnIsBankHoliday) And (Not blnIncBankHoliday) And ((blnShadeBankHolidays) Or ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Working Day + ((Bank Holiday + Not Inc. Working Days Only + Shade Bank Holidays)))
                  intDefinedColourStyle = 1
                End If
                
              Else
                'Event + Not Working Day
                
                If (blnIncWorkingDays) And ((blnIsBankHoliday And Not blnIncBankHoliday) Or (Not blnIsBankHoliday)) And ((blnIsWeekend And Not blnShadeWeekends) Or (Not blnIsWeekend)) Then
                  'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only) || Not Bank Holiday) + ((Weekend + Not Show Weekends) || Not Weekend))
                  intDefinedColourStyle = 0
                End If
                
                If (blnIsBankHoliday) And (blnShadeBankHolidays) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                  'Event + Not Working Day + Bank Holiday + Shade Bank Holidays + Inc. Working Days Only + Not Inc. Bank Holidays
                  intDefinedColourStyle = 1
                ElseIf (blnIsWeekend) And (blnShadeWeekends) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                  'Event + Not Working Day + Weekend + Show Weekends + Inc. Working Days Only + Not Inc. Bank Holidays
                  intDefinedColourStyle = 1
                ElseIf (blnIsWeekend) And (Not blnIsBankHoliday) And (blnShadeWeekends) And (blnIncWorkingDays) And (blnIncBankHoliday) Then
                  'Event + Not Working Day + Weekend + Show Weekends + Inc. Working Days Only + Inc. Bank Holidays
                  intDefinedColourStyle = 1
                End If

                If (blnIsBankHoliday) And (blnIsWeekend) And (blnShadeWeekends) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                  'Event + Not Working Day + Bank Holiday + Weekend + Show Weekends + Inc. Working Days Only + Not Inc. Bank Holidays
                  intDefinedColourStyle = 1
                End If
                
                If (blnIsBankHoliday) And (Not blnIncBankHoliday) And (Not blnShadeBankHolidays) And (Not ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only + Not Shade Bank Holidays)))
                  intDefinedColourStyle = 0
                ElseIf (blnIsBankHoliday) And (Not blnIncBankHoliday) And ((blnShadeBankHolidays) Or ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only + Shade Bank Holidays)))
                  intDefinedColourStyle = 1
                End If

              End If
              
            Else
              'Not Event
              intDefinedColourStyle = 0
              
              If (blnIsWeekend) And (blnShadeWeekends) Then
                'Not Event + Weekend + Show Weekends
                intDefinedColourStyle = 1
              End If
                
              If (blnIsBankHoliday) And (blnShadeBankHolidays) Then
                'Not Event + Bank Holiday + Show Bank Holidays
                intDefinedColourStyle = 1
              End If
                             
            End If
            
            
            Select Case intDefinedColourStyle
              Case 0:
                'Show the default colour
                .BackColor = lblCalDates(0).BackColor
                .ForeColor = .BackColor
                blnShowCaption = False
                
              Case 1:
                'Show the Weekend/Bank Holiday colour
                .BackColor = lblWeekend.BackColor
                .ForeColor = .BackColor
                blnShowCaption = False
              
              Case 2:
                'Show the colour from the Event Key!
                .BackColor = strColour
                .ForeColor = lblCalDates(0).ForeColor
                'lblLegend.Item(1).ForeColor = lblCalDates(0).ForeColor
                blnShowCaption = True
              
            End Select
            
            'set key character OR NOT.
            If ((chkCaptions.Value = vbChecked) And (blnShowCaption)) Then
              .ForeColor = lblCalDates(0).ForeColor
            Else
              .ForeColor = .BackColor
            End If
            
            If mblnChangingDate Then
  
              '**************************************************************************
              ' setup variables for bold line separators
              intIndexModulus = (intCount Mod 74)
              If intIndexModulus = 0 Then
                intCurrentStartCount = 0
                intCurrentEndCount = 0
              End If
  
              intThisStartCount = CInt(TagInfo_Get(.Tag, "EVENT_START"))
              intThisEndCount = CInt(TagInfo_Get(.Tag, "EVENT_END"))
              intCurrentStartCount = intCurrentStartCount + intThisStartCount
              intCurrentEndCount = intCurrentEndCount + intThisEndCount
  
              If (intIndexModulus > 0) Then
                Set lblTempNext = ctlCalDates(intBaseCount).CalDate(intCount + 1)
                If (TagInfo_Get(lblTempNext.Tag, "DATE") <> "  /  /    ") Then
                  intNextStartCount = CInt(TagInfo_Get(lblTempNext.Tag, "EVENT_START"))
                End If
              Else
                intNextStartCount = 0
              End If
  
              If ((intIndexModulus > 0) And (intIndexModulus < 73)) Then
                Set lblTempNext2 = ctlCalDates(intBaseCount).CalDate(intCount + 2)
                If (TagInfo_Get(lblTempNext2.Tag, "DATE") <> "  /  /    ") Then
                  intNext2StartCount = CInt(TagInfo_Get(lblTempNext2.Tag, "EVENT_START"))
                End If
              Else
                intNext2StartCount = 0
              End If
              
              strSession = UCase(TagInfo_Get(.Tag, "SESSION"))
  
              blnNextHasEvent = IIf(CInt(TagInfo_Get(lblTempNext.Tag, "HAS_EVENT")) > 0, True, False)
  
              If (intIndexModulus > 0) And (TagInfo_Get(lblTempNext2.Tag, "DATE") <> "  /  /    ") Then
                blnNext2HasEvent = IIf(CInt(TagInfo_Get(lblTempNext2.Tag, "HAS_EVENT")) > 0, True, False)
              Else
                blnNext2HasEvent = False
              End If
              
              Set lblTempPrev = ctlCalDates(intBaseCount).CalDate(intCount - 1)
              If (intIndexModulus > 1) And (TagInfo_Get(lblTempPrev.Tag, "DATE") <> "  /  /    ") Then
                blnPrevHasEvent = IIf(CInt(TagInfo_Get(lblTempPrev.Tag, "HAS_EVENT")) > 0, True, False)
              Else
                blnPrevHasEvent = False
              End If
  
              'separate the event - positions a bold line separating events if required.
              If (intThisEndCount > 0) _
                  And ((intNextStartCount > 0) Or (intNext2StartCount > 0)) _
                  And Not (intCurrentStartCount > intCurrentEndCount) Then
                'Add Date separator
                ctlCalDates(intBaseCount).AddEventSeparator intCount, strSession, blnHasEvent, _
                                                            blnNextHasEvent, blnNext2HasEvent, _
                                                            blnPrevHasEvent
              End If
              '**************************************************************************
              
            End If
            
            If intSessionCount = 2 Then
              intSessionCount = 0
            End If
          
          Else
            .BackColor = lblRangeDisabled.BackColor
            .ForeColor = .BackColor
          End If
          
        End If
      
      End With

    Next intCount
  Next intBaseCount
  
  RefreshDateSpecifics = True
  
TidyUpAndExit:
  Set lblTemp = Nothing
  Screen.MousePointer = vbDefault
  Exit Function
  
ErrorTrap:
  RefreshDateSpecifics = False
  GoTo TidyUpAndExit
End Function

Private Function SetScrollBarValues() As Boolean
  
  '******************************************************************************
  'format, resize & set values for the calendar scroll bar.
 
  If ((picCalendar.Height <= picScroll.Height) And _
    (picBase.Height <= picScroll.Height)) Then
    VScrollCalendar.Value = 0
    VScrollCalendar.Visible = False
    VScrollCalendar.Enabled = False
  Else
    VScrollCalendar.Visible = True
    VScrollCalendar.Enabled = True

    'TM - if the Windows Display Option "Show window contents while dragging" is checked then an error is caused
    'to be expedient I have fudged this.
    On Error Resume Next
    mlngScrollBarMultiplier = Fix((picCalendar.Height - picScroll.Height) / 32767) + 1

    VScrollCalendar.Max = ((picCalendar.Height - picScroll.Height) / mlngScrollBarMultiplier)
    
    VScrollCalendar.SmallChange = ((VScrollCalendar.Max) / (mintBaseRecordCount))
    VScrollCalendar.LargeChange = (5 * VScrollCalendar.SmallChange)
  
  End If

  With VScrollCalendar
    .Top = (picPrint.Top + picDates.Height - 10)
    .Height = (picPrint.Height - picDates.Height)
    .Width = SCROLLBAR_WIDTH
  End With
  
  '******************************************************************************
  'format, resize & set values for the legend scroll bar.
  VScrollLegend.Max = picLegend.Height - picLegendScroll.Height
  VScrollLegend.SmallChange = IIf((LEGEND_BOXOFFSETY > VScrollLegend.Max), (LEGEND_BOXOFFSETY / 2), LEGEND_BOXOFFSETY)
  VScrollLegend.LargeChange = IIf((LEGEND_BOXOFFSETY * 2) > VScrollLegend.Max, VScrollLegend.SmallChange, (LEGEND_BOXOFFSETY * 2))
  
  If picLegend.Height <= picLegendScroll.Height Then
    VScrollLegend.Visible = False
    VScrollLegend.Enabled = False
  Else
    VScrollLegend.Visible = True
    VScrollLegend.Enabled = True
  End If
  
  'resize & position the controls encapsulated in the fraLegend container
  With picLegend
    .Left = 0
  End With
 
  With VScrollLegend
    If .Enabled Then
      picLegendScroll.Width = fraLegend.Width - (2 * picLegendScroll.Left) - .Width
    Else
      picLegendScroll.Width = fraLegend.Width - (2 * picLegendScroll.Left)
    End If
    picLegend.Width = picLegendScroll.Width
    
    .Left = picLegendScroll.Left + picLegendScroll.Width - 15
    .Top = picLegendScroll.Top
    .Height = picLegendScroll.Height
    .Width = SCROLLBAR_WIDTH
  End With
 
End Function

Public Property Let ShowBankHolidays(pblnShowBankHols As Boolean)
  If (Not mblnGroupByDesc) _
    And (Not mblnDisableRegions) _
    And (((mblnPersonnelBase) And (Len(Trim(gsPersonnelRegionColumnName)) > 0) And (glngBHolRegionID > 0)) _
      Or ((mblnPersonnelBase) And (Len(Trim(gsPersonnelHRegionColumnName)) > 0) And (glngBHolRegionID > 0)) _
      Or (mlngStaticRegionColumnID > 0)) Then
      
    chkShadeBHols.Enabled = True
    If pblnShowBankHols Then
      chkShadeBHols.Value = vbChecked
    Else
      chkShadeBHols.Value = vbUnchecked
    End If
  Else
    chkShadeBHols.Value = vbUnchecked
    chkShadeBHols.Enabled = False
  End If
End Property

Public Property Let IncludeWorkingDaysOnly(pblnIncludeWorkingDaysOnly As Boolean)
  If (Not mblnGroupByDesc) _
    And (Not mblnDisableWPs) _
    And (((mblnPersonnelBase) And (Len(Trim(gsPersonnelWorkingPatternColumnName)) > 0)) _
      Or ((mblnPersonnelBase) And (Len(Trim(gsPersonnelHWorkingPatternColumnName)) > 0))) Then
      
    chkIncludeWorkingDaysOnly.Enabled = True
    If pblnIncludeWorkingDaysOnly Then
      chkIncludeWorkingDaysOnly.Value = vbChecked
    Else
      chkIncludeWorkingDaysOnly.Value = vbUnchecked
    End If
  Else
    chkIncludeWorkingDaysOnly.Value = vbUnchecked
    chkIncludeWorkingDaysOnly.Enabled = False
  End If
End Property

Public Property Let IncludeBankHolidays(pblnIncludeBankHolidays As Boolean)
  If (Not mblnGroupByDesc) _
    And (Not mblnDisableRegions) _
    And (((mblnPersonnelBase) And (Len(Trim(gsPersonnelRegionColumnName)) > 0) And (glngBHolRegionID > 0)) _
      Or ((mblnPersonnelBase) And (Len(Trim(gsPersonnelHRegionColumnName)) > 0) And (glngBHolRegionID > 0)) _
      Or (mlngStaticRegionColumnID > 0)) Then
      
    chkIncludeBHols.Enabled = True
    If pblnIncludeBankHolidays Then
      chkIncludeBHols.Value = vbChecked
    Else
      chkIncludeBHols.Value = vbUnchecked
    End If
  Else
    chkIncludeBHols.Value = vbUnchecked
    chkIncludeBHols.Enabled = False
  End If
End Property

Public Property Let ShowCaptions(pblnShowCaptions As Boolean)
  chkCaptions.Value = IIf(pblnShowCaptions, vbChecked, vbUnchecked)
End Property

Public Property Let ShowWeekends(pblnShowWeekends As Boolean)
  chkShadeWeekends.Value = IIf(pblnShowWeekends, vbChecked, vbUnchecked)
End Property

Public Function Initialise() As Boolean
  
  Dim fOK As Boolean
  Dim blnRegionEnabled As Boolean
  Dim blnWorkingPatternEnabled As Boolean
  
  fOK = True
  
  mstrDateFormat = DateFormat
  
  blnRegionEnabled = False
  blnWorkingPatternEnabled = False
  
  Screen.MousePointer = vbHourglass

  ' Set the loading flag
  mblnLoading = True
  
  If (fOK And mblnPersonnelBase _
      And (grtRegionType = rtHistoricRegion) _
      And (Not mblnGroupByDesc) _
      And (mlngStaticRegionColumnID < 1)) _
    Or _
      (fOK And ((mlngStaticRegionColumnID > 0) _
          Or (mblnPersonnelBase _
              And (grtRegionType = rtStaticRegion))) _
              And (Not mblnGroupByDesc)) Then

    blnRegionEnabled = CheckPermission_RegionInfo
  End If

  If blnRegionEnabled Then
    If fOK And mblnPersonnelBase _
      And (grtRegionType = rtHistoricRegion) _
      And (Not mblnGroupByDesc) _
      And (mlngStaticRegionColumnID < 1) Then

      'get historical bank holidays
      fOK = Get_HistoricBankHolidays

      If fOK Then mblnRegions = True

    ElseIf fOK And ((mlngStaticRegionColumnID > 0) _
            Or (mblnPersonnelBase _
                And (grtRegionType = rtStaticRegion))) _
      And (Not mblnGroupByDesc) Then

      'get static bank holidays collection
      fOK = Get_StaticBankHolidays

      If fOK Then mblnRegions = True

    Else
      ShowBankHolidays = False

    End If
  End If

  If (fOK And mblnPersonnelBase _
      And (gwptWorkingPatternType = wptHistoricWPattern) _
      And (Not mblnGroupByDesc)) Or _
      (fOK And (mblnPersonnelBase _
      And (gwptWorkingPatternType = wptStaticWPattern) _
      And (Not mblnGroupByDesc))) Then

    blnWorkingPatternEnabled = CheckPermission_WPInfo
  End If
  
  If blnWorkingPatternEnabled Then
    If fOK And mblnPersonnelBase _
      And (gwptWorkingPatternType = wptHistoricWPattern) _
      And (Not mblnGroupByDesc) _
      Then
      
      'get historical working patterns
      fOK = Get_HistoricWorkingPatterns
      
      If fOK Then mblnWorkingPatterns = True
    
    ElseIf fOK And (mblnPersonnelBase _
      And (gwptWorkingPatternType = wptStaticWPattern) _
      And (Not mblnGroupByDesc)) _
       Then
    
      'get static working patterns
      fOK = Get_StaticWorkingPatterns
      
      If fOK Then mblnWorkingPatterns = True
    
    Else
      IncludeWorkingDaysOnly = False
      
    End If
  End If
  
  If fOK Then fOK = GetAvailableColours(mstrExcludedColours)
  
  If fOK Then fOK = Load_Legend
    
  mblnLoading = False
  
  Initialise = fOK
  
End Function


Private Function Load_Description(plngBaseRecordID As Long, pstrBaseRecordDesc As String) As Boolean

  Dim lngNewIndex As Long
  
  lngNewIndex = lblBaseDesc().UBound + 1
  
  Load lblBaseDesc(lngNewIndex)
  
  With lblBaseDesc(lngNewIndex)
    .Tag = IIf(mblnGroupByDesc, "-1", plngBaseRecordID)
    .Caption = Replace(pstrBaseRecordDesc, "&", "&&")
    .Height = BASE_BOXHEIGHT
    .Width = BASE_BOXWIDTH
    .Left = BASE_BOXSTARTX
    .Top = BASE_BOXSTARTY + ((BASE_BOXHEIGHT) * (lngNewIndex - 1))
    .Visible = True
  End With
  
  mintCurrentBaseIndex = lblBaseDesc().UBound
  
  Load_Description = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_Description = False
  GoTo TidyUpAndExit
  
End Function

Private Function Load_Dates() As Boolean

  Dim iNewIndex As Integer
  Dim intDateValue As Integer
  Dim intControlCount As Integer
  Dim intDateCount As Integer

  'Define the current visible Start and End Dates.
  If mfDefaultToSystemDate Then
    mintDaysInMonth = DaysInMonth(mdtSystemStartDate)
    mdtVisibleEndDate = DateAdd("d", CDbl(mintDaysInMonth - Day(mdtSystemStartDate)), mdtSystemStartDate)
    mdtVisibleStartDate = DateAdd("d", CDbl(-(mintDaysInMonth - 1)), mdtSystemEndDate)
  Else
    mintDaysInMonth = DaysInMonth(mdtReportStartDate)
    mdtVisibleEndDate = DateAdd("d", CDbl(mintDaysInMonth - Day(mdtReportStartDate)), mdtReportStartDate)
    mdtVisibleStartDate = DateAdd("d", CDbl(-(mintDaysInMonth - 1)), mdtVisibleEndDate)
  End If
  
'  mintFirstDayOfMonth = Weekday("01/" & CStr(pintMonth) & "/" + CStr(plngYear), vbSunday)
  mintFirstDayOfMonth = Weekday(mdtVisibleStartDate, vbSunday)
  
  For intControlCount = 1 To DAY_CONTROL_COUNT Step 1
    
    iNewIndex = lblDate().UBound + 1
    Load lblDate(iNewIndex)
    
    With lblDate(iNewIndex)
      If (intControlCount >= mintFirstDayOfMonth) And _
        (intControlCount < (mintFirstDayOfMonth + mintDaysInMonth)) Then
        intDateCount = intDateCount + 1
        .Tag = intDateCount
        .Caption = intDateCount
      Else
        'Add a blank date box
        .Tag = ""
        .Caption = ""
      End If
      
      .Width = DATES_BOXWIDTH
      .Height = DATES_BOXHEIGHT
      .Top = DATES_BOXSTARTY
      .Left = (DATES_BOXSTARTX + ((DATES_BOXWIDTH - 15) * (intControlCount - 1)))
      .Visible = True
    End With
  
  Next intControlCount
  
  Load_Dates = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_Dates = False
  GoTo TidyUpAndExit
  
End Function

Public Function FillEventCalBoxes(plngCalDatIndex As Integer, plngStart As Long, plngEnd As Long) As Boolean

  ' This function actually fills the cal boxes between the indexes specified
  ' according to the options selected by the user.
  
  On Error GoTo ErrorTrap
  
  Dim colEvents As clsCalendarEvents
  
  Dim intCount As Integer
  
  Dim strCurrentRegion_BD As String
  Dim strCurrentWorkingPattern_BD As String
  
  Dim lblTemp As VB.Label

  Dim intStartCount As Integer
  Dim intEndCount As Integer
  
  ' Loop through the indexes as specified.
  For intCount = plngStart To plngEnd Step 1
    Set lblTemp = ctlCalDates(CInt(plngCalDatIndex)).CalDate(intCount)
    
    If intCount = plngStart Then

      If Trim(TagInfo_Get(lblTemp.Tag, "EVENT_START")) = "" Then
        TagInfo_Store lblTemp.Tag, "EVENT_START", "0"
      End If
      
      intStartCount = CInt(TagInfo_Get(lblTemp.Tag, "EVENT_START")) + 1

      lblTemp.Tag = TagInfo_Store(lblTemp.Tag, "EVENT_START", CStr(intStartCount))

    End If
    
    If intCount = plngEnd Then
      intEndCount = CInt(TagInfo_Get(lblTemp.Tag, "EVENT_END")) + 1
      lblTemp.Tag = TagInfo_Store(lblTemp.Tag, "EVENT_END", CStr(intEndCount))
    End If
    
    If CInt(TagInfo_Get(lblTemp.Tag, "HAS_EVENT")) = 0 Then
      'Date & Session clear
        
      With lblTemp
        .Caption = Replace(Replace(mstrEventLegend_BD, "&", "&&"), " ", "")
        If Len(.Caption) > 1 Then
          .Font.Size = KEY_FONTSIZE_SMALL
        Else
          .Font.Size = KEY_FONTSIZE_MEDIUM
        End If
        
        .BackColor = GetLegendColour(mstrCurrentEventKey)
        .ForeColor = vbBlack
        
        .Tag = TagInfo_Store(.Tag, "CAPTION", mstrEventLegend_BD)

        .Tag = TagInfo_Store(.Tag, "COLOUR", HexValue(.BackColor))
        .Tag = TagInfo_Store(.Tag, "HAS_EVENT", "1")
        
        .ToolTipText = mstrEventToolTip
        
        
        strCurrentRegion_BD = TagInfo_Get(.Tag, "REGION")
        
        strCurrentWorkingPattern_BD = TagInfo_Get(.Tag, "WORKING_PATTERN")
      End With
      
      '--------------------------------------------------------------------------
      'Add event to colEvents --> add colEvents to mcolDateControlEvents for the CalDate control.
      'NOTE: Use values directly from the recordset as these might have been changed in FillGridWithEvents.
      Set colEvents = New clsCalendarEvents
      
      colEvents.Add CStr(colEvents.Count), mstrEventName_BD, , , , , CStr(mrsEvents.Fields("StartDate").Value), , _
                    mrsEvents.Fields("StartSession").Value, , CStr(mrsEvents.Fields("EndDate").Value), , _
                    mrsEvents.Fields("EndSession").Value, , Format(mrsEvents.Fields("Duration").Value, "###0.0"), , mstrEventLegend_BD, _
                    , , , , , , , , IIf(IsNull(mrsEvents.Fields("EventDescription1ColumnID").Value), 0, mrsEvents.Fields("EventDescription1ColumnID").Value), CStr(IIf(IsNull(mrsEvents.Fields("EventDescription1Column").Value), "", _
                    mrsEvents.Fields("EventDescription1Column").Value)), IIf(IsNull(mrsEvents.Fields("EventDescription2ColumnID").Value), 0, mrsEvents.Fields("EventDescription2ColumnID").Value), _
                    CStr(IIf(IsNull(mrsEvents.Fields("EventDescription2Column").Value), "", _
                    mrsEvents.Fields("EventDescription2Column").Value)), mstrBaseDescription_BD, strCurrentRegion_BD, _
                    strCurrentWorkingPattern_BD, mstrDesc1Value_BD, mstrDesc2Value_BD
      mcolDateControlEvents.Add colEvents, CStr(plngCalDatIndex & "_CALDATEINDEX_" & intCount)
      Set colEvents = Nothing
      '--------------------------------------------------------------------------
      
    Else
      'Date & Session already has an event, set it as Multiple.
      
      With lblTemp
        .FontSize = KEY_FONTSIZE_MEDIUM
        .Caption = "."
        .BackColor = vbWhite
        .ForeColor = vbBlack
      
        .Tag = TagInfo_Store(.Tag, "CAPTION", mstrEventLegend_BD)
        .Tag = TagInfo_Store(.Tag, "COLOUR", HexValue(.BackColor))
        .Tag = TagInfo_Store(.Tag, "HAS_EVENT", "2")

        .ToolTipText = "Multiple Events - Click for details"
        
        strCurrentRegion_BD = TagInfo_Get(.Tag, "REGION")
        strCurrentWorkingPattern_BD = TagInfo_Get(.Tag, "WORKING_PATTERN")
      End With
      
      '--------------------------------------------------------------------------
      'Add event to colEvents --> add colEvents to mcolDateControlEvents for the CalDate control.
      Set colEvents = mcolDateControlEvents.Item(CStr(plngCalDatIndex & "_CALDATEINDEX_" & intCount))
      colEvents.Add CStr(colEvents.Count), mstrEventName_BD, , , , , CStr(mrsEvents.Fields("StartDate").Value), , _
                    mrsEvents.Fields("StartSession").Value, , CStr(mrsEvents.Fields("EndDate").Value), , _
                    mrsEvents.Fields("EndSession").Value, , Format(mrsEvents.Fields("Duration").Value, "###0.0"), , mstrEventLegend_BD, _
                    , , , , , , , , IIf(IsNull(mrsEvents.Fields("EventDescription1ColumnID").Value), 0, mrsEvents.Fields("EventDescription1ColumnID").Value), CStr(IIf(IsNull(mrsEvents.Fields("EventDescription1Column").Value), "", _
                    mrsEvents.Fields("EventDescription1Column").Value)), IIf(IsNull(mrsEvents.Fields("EventDescription2ColumnID").Value), 0, mrsEvents.Fields("EventDescription2ColumnID").Value), _
                    CStr(IIf(IsNull(mrsEvents.Fields("EventDescription2Column").Value), "", _
                    mrsEvents.Fields("EventDescription2Column").Value)), mstrBaseDescription_BD, strCurrentRegion_BD, _
                    strCurrentWorkingPattern_BD, mstrDesc1Value_BD, mstrDesc2Value_BD
      Set colEvents = Nothing
      '--------------------------------------------------------------------------
      
    End If
    
  Next intCount
  
  FillEventCalBoxes = True
  
TidyUpAndExit:
  Set lblTemp = Nothing
  Exit Function
  
ErrorTrap:
  COAMsgBox "An Error Has Occurred Whilst Filling The Calendar:" & vbNewLine & Err.Number & " - " & Err.Description
  FillEventCalBoxes = False
  GoTo TidyUpAndExit
  
End Function

Private Function GetLegendColour(pstrEventKey As String) As String
  
  Dim ctl As Control

  For Each ctl In lblLegend
    If UCase(RTrim(ctl.Tag)) = UCase(RTrim(pstrEventKey)) Then
      GetLegendColour = HexValue(ctl.BackColor)
      Exit Function
    End If
  Next ctl
  
  GetLegendColour = vbBlack
  
End Function

Private Function Load_Legend() As Boolean

  On Error GoTo ErrorTrap
  
  Dim intNewIndex As Integer
  Dim intCount As Integer
  Dim lngWidth As Long
  
  Dim strEventID As String
  
  Dim blnNewEvent As Boolean
  
  Dim intColourIndex As Integer
  Dim intColourMax As Integer
  
  Dim lngFC_Data As Long
  Dim lngBD_Data As Long
  Dim lngFC_Header As Long
  Dim lngBC_Header As Long
  
  Const LEGEND_COLS = 2
  
  strEventID = vbNullString

  ReDim mavLegend(3, 0)
  
  mintLegendCount = 0
  
  intColourMax = UBound(mavAvailableColours, 2)
  
  With mrsEvents
    If Not (.BOF And .EOF) Then
      
      .MoveFirst
      Do While Not .EOF
        If strEventID <> .Fields(mstrEventIDColumn).Value Then
          strEventID = .Fields(mstrEventIDColumn).Value
          
          blnNewEvent = True
          For intCount = 1 To UBound(mavLegend, 2) Step 1
            If mavLegend(0, intCount) = strEventID Then
              blnNewEvent = False
            End If
          Next intCount
          
          If blnNewEvent Then
            intNewIndex = UBound(mavLegend, 2) + 1
            
            ReDim Preserve mavLegend(3, intNewIndex)
            mavLegend(0, intNewIndex) = strEventID
            mavLegend(1, intNewIndex) = Left(mrsEvents.Fields("Name").Value, 50)
            mavLegend(2, intNewIndex) = Left(mrsEvents.Fields("Legend").Value, 2)
          
            intColourIndex = (intNewIndex - 1) Mod intColourMax
            mavLegend(3, intNewIndex) = mavAvailableColours(1, intColourIndex)
            
          End If
        End If

        .MoveNext
      Loop
      
      ' Sort the Array here - then add the Multiple events item to the end.
      SortLegend mavLegend, 1
      
      If mblnHasMultipleEvents Then
        intNewIndex = UBound(mavLegend, 2) + 1
        ReDim Preserve mavLegend(3, intNewIndex)
        mavLegend(0, intNewIndex) = "EVENT_MULTIPLE"
        mavLegend(1, intNewIndex) = "Multiple Events"
        mavLegend(2, intNewIndex) = "."
        mavLegend(3, intNewIndex) = "&HFFFFFF"
      End If
      
      'calculate how many items are in each of the columns
      mintLegendCount = UBound(mavLegend, 2)
      If (mintLegendCount Mod LEGEND_COLS) = 0 Then
        mintLegendLeft = (mintLegendCount / LEGEND_COLS)
      ElseIf mintLegendCount = 1 Then
        mintLegendLeft = 1
      Else
        mintLegendLeft = Fix(mintLegendCount / LEGEND_COLS) + 1
      End If
      mintLegendRight = (mintLegendCount - mintLegendLeft)
      
      lngWidth = ((picLegend.ScaleWidth / 2) - (3 * CONTROL_OFFSET) - LEGEND_BOXWIDTH)

      For intCount = 1 To mintLegendLeft Step 1
        intNewIndex = lblLegend.UBound + 1
        
        Load lblLegend(intNewIndex)
        With lblLegend(intNewIndex)
          .Tag = mavLegend(0, intNewIndex)
          .Top = (LEGEND_BOXSTARTY + (LEGEND_BOXOFFSETY * (intCount - 1)))
          .Left = CONTROL_OFFSET
          .Height = LEGEND_BOXHEIGHT
          .Width = LEGEND_BOXWIDTH
          .Visible = True
          .Caption = Replace(IIf(IsNull(mavLegend(2, intNewIndex)), "", mavLegend(2, intNewIndex)), "&", "&&")
          
          If Len(mavLegend(2, intNewIndex)) > 1 Then
            .Font.Size = KEY_FONTSIZE_SMALL
          Else
            .Font.Size = KEY_FONTSIZE_MEDIUM
          End If
          
          .BackColor = mavLegend(3, intNewIndex)
          .ToolTipText = mavLegend(1, intNewIndex)
        End With
       
        Load lblEventName(intNewIndex)
        With lblEventName(intNewIndex)
          .Tag = mavLegend(0, intNewIndex)
          .Top = (lblLegend(intNewIndex).Top)
          .Left = (lblLegend(intNewIndex).Left + lblLegend(intNewIndex).Width + CONTROL_OFFSET)
          .Height = LEGENDDESC_BOXHEIGHT
          .Width = lngWidth
          .Visible = True
          .Caption = Replace(IIf(IsNull(mavLegend(1, intNewIndex)), "", mavLegend(1, intNewIndex)), "&", "&&")
          .ForeColor = &HFF0000
          .ToolTipText = mavLegend(1, intNewIndex)
        End With
        
      Next intCount
      
      For intCount = 1 To mintLegendRight Step 1
        intNewIndex = lblLegend.UBound + 1
        
        Load lblLegend(intNewIndex)
        With lblLegend(intNewIndex)
          .Tag = mavLegend(0, intNewIndex)
          .Top = (LEGEND_BOXSTARTY + (LEGEND_BOXOFFSETY * (intCount - 1)))
          .Left = ((picLegend.Width / 2) + CONTROL_OFFSET)
          .Height = LEGEND_BOXHEIGHT
          .Width = LEGEND_BOXWIDTH
          .Visible = True
          .Caption = Replace(IIf(IsNull(mavLegend(2, intNewIndex)), "", mavLegend(2, intNewIndex)), "&", "&&")
          
          If Len(mavLegend(2, intNewIndex)) > 1 Then
            .Font.Size = KEY_FONTSIZE_SMALL
          Else
            .Font.Size = KEY_FONTSIZE_MEDIUM
          End If
          
          .BackColor = mavLegend(3, intNewIndex)
          .ToolTipText = mavLegend(1, intNewIndex)
        End With
       
        Load lblEventName(intNewIndex)
        With lblEventName(intNewIndex)
          .Tag = mavLegend(0, intNewIndex)
          .Top = (lblLegend(intNewIndex).Top)
          .Left = (lblLegend(intNewIndex).Left + lblLegend(intNewIndex).Width + CONTROL_OFFSET)
          .Height = LEGENDDESC_BOXHEIGHT
          .Width = lngWidth
          .Visible = True
          .Caption = Replace(IIf(IsNull(mavLegend(1, intNewIndex)), "", mavLegend(1, intNewIndex)), "&", "&&")
          .ForeColor = &HFF0000
          .ToolTipText = mavLegend(1, intNewIndex)
        End With
        
      Next intCount

      picLegend.Height = (lblLegend(mintLegendLeft).Top + LEGEND_BOXOFFSETY + CONTROL_OFFSET)
  
    Else
      Load_Legend = False
      GoTo TidyUpAndExit
      
    End If
  End With
  
  Load_Legend = True
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  mstrErrorMessage = "Error creating Calendar Report Key."
  Load_Legend = False
  GoTo TidyUpAndExit
  
End Function

Public Function GetCalLabelIndex(pintBaseRecordIndex As Integer, pdtDate As Date, pblnSession As Boolean) As Integer

  ' This function returns the index value of the cal label for the specified date.
  ' and session.
  
  Dim dtFirstDate As Date
  Dim dtLastDate As Date
  
  Dim iCount As Integer
  
  Dim lblTemp As VB.Label
  
'  dtFirstDate = CDate("01/" & cboMonth.ItemData(cboMonth.ListIndex) & "/" & spnYear.Value)
'  dtLastDate = CDate(DaysInMonth(dtFirstDate) & "/" & cboMonth.ItemData(cboMonth.ListIndex) & "/" & spnYear.Value)
  dtFirstDate = mdtVisibleStartDate
  dtLastDate = mdtVisibleEndDate
  
  If (pdtDate < dtFirstDate) Or (pdtDate > dtLastDate) Then
    GetCalLabelIndex = -1
    Exit Function
  End If
  
  For iCount = 1 To 74 Step 2
    Set lblTemp = ctlCalDates(pintBaseRecordIndex).CalDate(iCount)
    If TagInfo_Get(lblTemp.Tag, "DATE") <> "  /  /    " Then
      If (TagInfo_Get(lblTemp.Tag, "DATE") = Format(pdtDate, CALREP_DATEFORMAT)) Then
        GetCalLabelIndex = IIf(pblnSession, iCount + 1, iCount)
        Set lblTemp = Nothing
        Exit Function
      End If
    End If
  Next iCount
  
  GetCalLabelIndex = -1
  
End Function

Public Function Output_GetCalArrayIndex(pintBaseRecordIndex As Integer, pdtDate As Date, pblnSession As Boolean) As Integer

  ' This function returns the index value for the specified date and session.

  Dim dtFirstDate As Date
  Dim dtLastDate As Date

  Dim iCount As Integer
  Dim varTempArray As Variant
  
'  dtFirstDate = CDate("01/" & mlngMonth_Output & "/" & mlngYear_Output)
'  dtLastDate = CDate(DaysInMonth(dtFirstDate) & "/" & mlngMonth_Output & "/" & mlngYear_Output)
  dtFirstDate = mdtVisibleStartDate_Output
  dtLastDate = mdtVisibleEndDate_Output
  
  If (pdtDate < dtFirstDate) Or (pdtDate > dtLastDate) Then
    Output_GetCalArrayIndex = -1
    Exit Function
  End If

  varTempArray = mavOutputDateIndex(2, pintBaseRecordIndex)

  For iCount = 1 To 74 Step 2
    
    If varTempArray(1, iCount) <> "  /  /    " Then
      If (varTempArray(1, iCount) = Format(pdtDate, CALREP_DATEFORMAT)) Then
        Output_GetCalArrayIndex = IIf(pblnSession, iCount + 1, iCount)
        Exit Function
      End If
    End If
    
  Next iCount

  Output_GetCalArrayIndex = -1
  
End Function

Private Function Load_Calendar() As Boolean

  On Error GoTo ErrorTrap
  
  Dim lngCount_X As Long
  Dim lngCount_Y As Long
  Dim lngCount As Long
  Dim intNewIndex As Integer
  Dim lngLeft As Long
  Dim lngTop As Long
  Dim lngHeight As Long
  Dim lngWidth As Long
  Dim lngControlCount As Long
  Dim intDateCount As Integer
  Dim lngNextIndex As Long
  Dim intSessionCount As Integer
  Dim intCurrentIndex As Integer
  Dim iCount2 As Integer
  Dim intControlCount As Integer
  Dim iCount As Integer
  Dim lngBaseID As Long
  
  Dim lblTemp As VB.Label
  
  Dim strIsWeekend As String
  Dim strIsBankHoliday As String
  Dim strRegion As String
  Dim strIsWorkingDay As String
  Dim strWorkingPattern As String
  
  Dim dtLabelsDate As Date
  
  Dim strSession As String
  
  intDateCount = 0
  
  intNewIndex = CLng(ctlCalDates.UBound) + 1
  DoEvents
  Load ctlCalDates(intNewIndex)
  
  lngTop = (lblBaseDesc(mintCurrentBaseIndex).Top - 5)
  With ctlCalDates(intNewIndex)
    .BoxSize = 255
    .Load_Labels
    .Visible = True
    .Move 0, lngTop
    '.Tag = mlngCurrentRecordID
    .Tag = mstrCurrentBaseRegion
  End With

  intSessionCount = 0
  
  For iCount2 = 1 To 74 Step 1
    intSessionCount = intSessionCount + 1
    
    If intSessionCount = 1 Then
      intControlCount = intControlCount + 1
    End If

    Set lblTemp = ctlCalDates(intNewIndex).CalDate(iCount2)
    
    With lblTemp
      
      .Visible = True
      
'      lngBaseID = ctlCalDates(intNewIndex).Tag
      lngBaseID = CLng(lblBaseDesc(intNewIndex).Tag)
      
      If (intControlCount >= mintFirstDayOfMonth) And (intControlCount < (mintFirstDayOfMonth + mintDaysInMonth)) Then
        strSession = IIf(intSessionCount = 2, " PM", " AM")
        If Trim(strSession) = "AM" Then
          intDateCount = intDateCount + 1
        End If
'        dtLabelsDate = CDate(intDateCount & "/" & CStr(cboMonth.ItemData(cboMonth.ListIndex)) & "/" & CStr(spnYear.Value))
        dtLabelsDate = DateAdd("d", CDbl(intDateCount - 1), mdtVisibleStartDate)

        .Caption = vbNullString
        .BackColor = lblCalDates(0).BackColor
        'populate control tag with information on calDate label control eg. is weekend, bankhol, used.
        
        .Tag = TagInfo_Store(.Tag, "DATE", Format(dtLabelsDate, CALREP_DATEFORMAT))
        .Tag = TagInfo_Store(.Tag, "SESSION", UCase$(Trim$(strSession)))

      Else
        .Caption = vbNullString
        .BackColor = lblDisabled.BackColor
        strSession = vbNullString
        'populate control tag with information on calDate label control eg. is weekend, bankhol, used.
        .Tag = TagInfo_Store(.Tag, "DATE", "  /  /    ")
        .Tag = TagInfo_Store(.Tag, "SESSION", "  ")
        
      End If
      
      If Trim(strSession) <> vbNullString Then
        strRegion = ctlCalDates(intNewIndex).Tag
        strIsBankHoliday = IIf(IsBankHoliday(dtLabelsDate, lngBaseID, strRegion), "1", "0")
        .Tag = TagInfo_Store(.Tag, "BANK_HOLIDAY", strIsBankHoliday)      'Is Bank Holiday
        .Tag = TagInfo_Store(.Tag, "REGION", strRegion)                   'Store Working Pattern
         
        'flag if the date is a weekend
        strIsWeekend = IIf(IsWeekend(dtLabelsDate), "1", "0")
        .Tag = TagInfo_Store(.Tag, "WEEKEND", strIsWeekend)               'Is Weekend
        
        'flag if the date & session is in the current personnel's working pattern.
        strIsWorkingDay = IIf(IsWorkingDay(dtLabelsDate, lngBaseID, Trim(strSession), strWorkingPattern), "1", "0")
        strWorkingPattern = IIf(Trim(strWorkingPattern) = vbNullString, "              ", strWorkingPattern)
        .Tag = TagInfo_Store(.Tag, "WORKING_DAY", strIsWorkingDay)        'Is Working Day
        .Tag = TagInfo_Store(.Tag, "WORKING_PATTERN", strWorkingPattern)  'Store Working Pattern

      Else
        .Tag = TagInfo_Store(.Tag, "BANK_HOLIDAY", "0")                   'Is Bank Holiday
        .Tag = TagInfo_Store(.Tag, "REGION", "<None>")                    'Store Working Pattern
        .Tag = TagInfo_Store(.Tag, "WEEKEND", "0")                        'Is Weekend
        .Tag = TagInfo_Store(.Tag, "WORKING_DAY", "0")                    'Is Working Day
        .Tag = TagInfo_Store(.Tag, "WORKING_PATTERN", "              ")   'Store Working Pattern
      End If
          
    End With
    
    If intSessionCount = 2 Then
      intSessionCount = 0
    End If

  Next iCount2
  intControlCount = 0
  intDateCount = 0
  
  Load_Calendar = True

TidyUpAndExit:
  Set lblTemp = Nothing
  Exit Function

ErrorTrap:
  mstrErrorMessage = Err.Description
  
  If Err.Number = 7 Then
    mstrErrorMessage = mstrErrorMessage & vbNewLine & "Try reducing the number of base table records in the Record Profile."
  End If
  
  Load_Calendar = False
  GoTo TidyUpAndExit
  
End Function

Private Function Load_DateVerticalLines() As Boolean

  Dim count_vert As Integer
  
  For count_vert = 1 To DAY_CONTROL_COUNT Step 1
    Load VerticalDateLine(count_vert)
    VerticalDateLine(count_vert).Visible = True
    VerticalDateLine(count_vert).X1 = lblDate(count_vert).Left
    VerticalDateLine(count_vert).X2 = VerticalDateLine(count_vert).X1
    VerticalDateLine(count_vert).Y1 = 0
    VerticalDateLine(count_vert).Y2 = (2 * DATES_BOXHEIGHT)
    VerticalDateLine(count_vert).ZOrder 0
  Next count_vert
  
  count_vert = (DAY_CONTROL_COUNT + 1)
  Load VerticalDateLine(count_vert)
  VerticalDateLine(count_vert).Visible = True
  VerticalDateLine(count_vert).X1 = (lblDate(lblDate.UBound).Left + lblDate(lblDate.UBound).Width - 15)
  VerticalDateLine(count_vert).X2 = VerticalDateLine(count_vert).X1
  VerticalDateLine(count_vert).Y1 = 0
  VerticalDateLine(count_vert).Y2 = (2 * DATES_BOXHEIGHT)
  VerticalDateLine(count_vert).BorderWidth = 1
  VerticalDateLine(count_vert).ZOrder 0

  Load_DateVerticalLines = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_DateVerticalLines = False
  GoTo TidyUpAndExit
  
End Function

Private Function Load_BaseHorizontalLines() As Boolean

  On Error GoTo ErrorTrap
  
  Dim intNewIndex As Integer
  
  intNewIndex = HorizontalBaseLine.UBound + 1
  
  Load HorizontalBaseLine(intNewIndex)
  With HorizontalBaseLine(intNewIndex)
    .X1 = -10
    .X2 = picBase.Width + 20
    .Y1 = lblBaseDesc(mintCurrentBaseIndex).Top + BASE_BOXHEIGHT
    .Y2 = .Y1
    .BorderWidth = 1
    .Visible = True
    .ZOrder 0
  End With

  Load_BaseHorizontalLines = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Load_BaseHorizontalLines = False
  GoTo TidyUpAndExit
  
End Function

Private Function DaysInMonth(pdtMonth As Date) As Integer
  
  'Return the number of days in the month
 
  Dim dtNextMonth As Date
  
  dtNextMonth = DateAdd("m", 1, pdtMonth)
  DaysInMonth = Day(DateAdd("d", Day(dtNextMonth) * -1, dtNextMonth))

End Function

Public Function RefreshMonthLayout() As Boolean
  
  'clear the calendar dates and display a new month/year

  On Error GoTo ErrorTrap
  
  Dim iNewIndex As Integer
  Dim intDateValue As Integer
  Dim intDateCount As Integer
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim lngMonth As Long
  Dim lngYear As Long
  Dim intStartIndex As Integer
  Dim intEndIndex As Integer
  Dim intCurrentIndex As Integer
  Dim intControlCount As Integer
  Dim dtMonth As Date
  
  'clear all the dates
  ClearDates
  
  lngMonth = cboMonth.ItemData(cboMonth.ListIndex)
  lngYear = spnYear.Value
  
  dtMonth = DateAdd("yyyy", CDbl(lngYear - Year(mdtReportStartDate)), mdtReportStartDate)
  dtMonth = DateAdd("m", CDbl(lngMonth - Month(mdtReportStartDate)), dtMonth)
  
  mintDaysInMonth = DaysInMonth(dtMonth)

  'Define the current visible Start and End Dates.
  mdtVisibleEndDate = DateAdd("d", CDbl(mintDaysInMonth - Day(dtMonth)), dtMonth)
  mdtVisibleStartDate = DateAdd("d", CDbl(-(mintDaysInMonth - 1)), mdtVisibleEndDate)
  
  mintFirstDayOfMonth = Weekday(mdtVisibleStartDate, vbSunday)

  AddDataToArray 1, 0, ""

  For iCount2 = 1 To DAY_CONTROL_COUNT Step 1
    intControlCount = intControlCount + 1
    With lblDate(iCount2)
      If (intDateCount < mintDaysInMonth) And (intControlCount >= mintFirstDayOfMonth) And _
        (intControlCount < (mintFirstDayOfMonth + mintDaysInMonth)) Then
        intDateCount = intDateCount + 1
        .Tag = intDateCount
        .Caption = intDateCount
      Else
        'Add a blank date box
        .Tag = ""
        .Caption = ""
      End If
      
      AddDataToArray 1, intControlCount, .Caption

    End With
  Next iCount2
  
'  'Define the current visible Start and End Dates.
'  mdtVisibleStartDate = CDate("01/" & Str(lngMonth) & "/" + Str(lngYear))
'  mdtVisibleEndDate = CDate(mintDaysInMonth & "/" & Str(lngMonth) & "/" + Str(lngYear))

  RefreshMonthLayout = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  RefreshMonthLayout = False
  GoTo TidyUpAndExit

End Function

Public Function RefreshCalendar() As Boolean
  
  'clear the calendar dates and display a new month/year

  On Error GoTo ErrorTrap
  
  Dim iNewIndex As Integer
  Dim intDateValue As Integer
  Dim intDateCount As Integer
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim intCurrentIndex As Integer
  Dim intControlCount As Integer
  Dim intSessionCount As Integer
  Dim intNextIndex As Integer
  
  Dim lblTemp As VB.Label
  
  Dim lngBaseID As Long
  
  Dim dtLabelsDate As Date
  
  Dim strIsWeekend As String
  Dim strIsBankHoliday As String
  Dim strRegion As String
  Dim strIsWorkingDay As String
  Dim strWorkingPattern As String
  
  Dim strSession As String
  
  intDateCount = 0
  
  If Not RefreshMonthLayout Then
    RefreshCalendar = False
    Exit Function
  End If
  
  For iCount = 1 To mintBaseRecordCount Step 1
    
    ctlCalDates(iCount).HideSeparators
    
    intSessionCount = 0
    
    For iCount2 = 1 To 74 Step 1
      intSessionCount = intSessionCount + 1
      
      If intSessionCount = 1 Then
        intControlCount = intControlCount + 1
      End If

      Set lblTemp = ctlCalDates(iCount).CalDate(iCount2)
      
      With lblTemp
        .Visible = True
        
        lngBaseID = CLng(lblBaseDesc(iCount).Tag)
        
        If (intControlCount >= mintFirstDayOfMonth) And (intControlCount < (mintFirstDayOfMonth + mintDaysInMonth)) Then
          strSession = IIf(intSessionCount = 2, " PM", " AM")
          If Trim(strSession) = "AM" Then
            intDateCount = intDateCount + 1
          End If
          'dtLabelsDate = CDate(intDateCount & "/" & CStr(cboMonth.ItemData(cboMonth.ListIndex)) & "/" & CStr(spnYear.Value))
          dtLabelsDate = DateAdd("d", CDbl(intDateCount - 1), mdtVisibleStartDate)

          .Caption = vbNullString
          .BackColor = lblCalDates(0).BackColor
          'populate control tag with information on calDate label control eg. is weekend, bankhol, used.
          .Tag = TagInfo_Store(.Tag, "DATE", Format(dtLabelsDate, CALREP_DATEFORMAT))
          .Tag = TagInfo_Store(.Tag, "SESSION", UCase(Trim(strSession)))

        Else
          .Caption = vbNullString
          .BackColor = lblDisabled.BackColor
          strSession = vbNullString
          'populate control tag with information on calDate label control eg. is weekend, bankhol, used.
          .Tag = TagInfo_Store(.Tag, "DATE", "  /  /    ")
          .Tag = TagInfo_Store(.Tag, "SESSION", UCase("  "))
          
        End If
        
        If Trim(strSession) <> vbNullString Then
          strRegion = ctlCalDates(iCount).Tag
          strIsBankHoliday = IIf(IsBankHoliday(dtLabelsDate, lngBaseID, strRegion), "1", "0")
          .Tag = TagInfo_Store(.Tag, "BANK_HOLIDAY", strIsBankHoliday)    'Is Bank Holiday
          .Tag = TagInfo_Store(.Tag, "REGION", strRegion)                         'Store Working Pattern
          
          'flag if the date is a weekend
          strIsWeekend = IIf(IsWeekend(dtLabelsDate), "1", "0")
          .Tag = TagInfo_Store(.Tag, "WEEKEND", strIsWeekend)    'Is Weekend
          
          'flag if the date & session is in the current personnel's working pattern.
          strIsWorkingDay = IIf(IsWorkingDay(dtLabelsDate, lngBaseID, Trim(strSession), strWorkingPattern), "1", "0")
          strWorkingPattern = IIf(Trim(strWorkingPattern) = vbNullString, "              ", strWorkingPattern)
          .Tag = TagInfo_Store(.Tag, "WORKING_DAY", strIsWorkingDay)      'Is Working Day
          .Tag = TagInfo_Store(.Tag, "WORKING_PATTERN", strWorkingPattern)   'Store Working Pattern

        Else
          .Tag = TagInfo_Store(.Tag, "BANK_HOLIDAY", "0")           'Is Bank Holiday
          .Tag = TagInfo_Store(.Tag, "REGION", "<None>")                    'Store Working Pattern
          .Tag = TagInfo_Store(.Tag, "WEEKEND", "0")                     'Is Weekend
          .Tag = TagInfo_Store(.Tag, "WORKING_DAY", "0")                   'Is Working Day
          .Tag = TagInfo_Store(.Tag, "WORKING_PATTERN", "              ")  'Store Working Pattern
        End If

      End With
      
      If intSessionCount = 2 Then
        intSessionCount = 0
      End If
      
    Next iCount2
    intControlCount = 0
    intDateCount = 0
  Next iCount
  
  RefreshCalendar = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  RefreshCalendar = False
  GoTo TidyUpAndExit

End Function

Public Function ClearDates() As Boolean

  ' This function clears the caldates and cal boxes, ready to display a new month
  
  On Error GoTo ClearAll_ERROR
  
  Dim iCount As Integer
  Dim iBaseCount As Integer
  Dim lblTemp As VB.Label
  
  For iBaseCount = 1 To mintBaseRecordCount Step 1
    For iCount = 1 To 74 Step 1
      Set lblTemp = ctlCalDates(iBaseCount).CalDate(iCount)
      With lblTemp
        .Caption = vbNullString
        .Tag = TagInfo_Store(.Tag, "DEFAULT", "")
        .BackColor = lblCalDates(0).BackColor
        .ToolTipText = vbNullString
      End With
    Next iCount
  Next iBaseCount
  
  ClearDates = True

TidyUpAndExit:
  Set lblTemp = Nothing
  Exit Function
  
ClearAll_ERROR:
  ClearDates = False
  GoTo TidyUpAndExit
  
End Function

Public Property Let StartOnCurrentMonth(New_Value As Boolean)
  mblnStartOnCurrentMonth = New_Value
End Property

Public Property Let ReportStartDate(pdtReportStartDate As Date)
  mdtReportStartDate = pdtReportStartDate
  mlngMonth = Month(pdtReportStartDate)
  mlngYear = Year(pdtReportStartDate)
End Property

Public Property Let ReportEndDate(pdtReportEndDate As Date)
  mdtReportEndDate = pdtReportEndDate
End Property

Public Property Let StaticRegionColumn(pstrStaticRegionColumn As String)
  mstrStaticRegionColumn = pstrStaticRegionColumn
End Property

Public Property Let StaticRegionColumnID(plngStaticRegionColumnID As Long)
  mlngStaticRegionColumnID = plngStaticRegionColumnID
End Property

Private Function TagInfo_Store(pstrTagValue As String, pstrParameter As String, pstrValue As String) As String

  Dim strCount As String
  Dim intStrLen As Integer
  
  Dim DEFAULT_VALUE As String
  
  DEFAULT_VALUE = "  /  /    AM0  00              " & HexValue(lblCalDates(0).BackColor) & "0000000<None>"
  
  Const WORKINGPATTERN_LENGTH = 14
  Const MAXCAPTION_LENGTH = 2

  If Trim(pstrTagValue) = vbNullString Then
    pstrTagValue = DEFAULT_VALUE
  End If
    
  Select Case UCase(pstrParameter)
    Case "DEFAULT": TagInfo_Store = DEFAULT_VALUE

    Case "DATE": TagInfo_Store = pstrValue & Mid(pstrTagValue, 11)
        
    Case "SESSION": TagInfo_Store = Left(pstrTagValue, 10) & pstrValue & Mid(pstrTagValue, 13)
    Case "BANK_HOLIDAY": TagInfo_Store = Left(pstrTagValue, 12) & pstrValue & Mid(pstrTagValue, 14)
    Case "CAPTION":
      intStrLen = Len(pstrValue)
      If Len(Trim(pstrValue)) < 1 Then
        pstrValue = String((MAXCAPTION_LENGTH), " ")
      ElseIf Len(Trim(pstrValue)) > MAXCAPTION_LENGTH Then
        pstrValue = Left(Trim(pstrValue), MAXCAPTION_LENGTH)
      Else
        pstrValue = pstrValue & String((MAXCAPTION_LENGTH - intStrLen), " ")
      End If
      TagInfo_Store = Left(pstrTagValue, 13) & pstrValue & Mid(pstrTagValue, 16)
      
    Case "WEEKEND": TagInfo_Store = Left(pstrTagValue, 15) & pstrValue & Mid(pstrTagValue, 17)
    Case "WORKING_DAY": TagInfo_Store = Left(pstrTagValue, 16) & pstrValue & Mid(pstrTagValue, 18)
    Case "WORKING_PATTERN":
      intStrLen = Len(pstrValue)
      If intStrLen < WORKINGPATTERN_LENGTH Then
        pstrValue = pstrValue & String((WORKINGPATTERN_LENGTH - intStrLen), " ")
      ElseIf intStrLen > WORKINGPATTERN_LENGTH Then
        pstrValue = Left(pstrValue, WORKINGPATTERN_LENGTH)
      End If
      TagInfo_Store = Left(pstrTagValue, 17) & pstrValue & Mid(pstrTagValue, 32)

    Case "COLOUR": TagInfo_Store = Left(pstrTagValue, 31) & pstrValue & Mid(pstrTagValue, 40)
    Case "HAS_EVENT": TagInfo_Store = Left(pstrTagValue, 39) & pstrValue & Mid(pstrTagValue, 41)

    Case "EVENT_START":
      strCount = String((3 - Len(pstrValue)), "0") & pstrValue
      TagInfo_Store = Left(pstrTagValue, 40) & strCount & Mid(pstrTagValue, 44)

    Case "EVENT_END":
      strCount = String((3 - Len(pstrValue)), "0") & pstrValue
      TagInfo_Store = Left(pstrTagValue, 43) & strCount & Mid(pstrTagValue, 47)

    Case "REGION": TagInfo_Store = Left(pstrTagValue, 46) & pstrValue

    Case Else: TagInfo_Store = pstrTagValue
  End Select

End Function

Private Function TagInfo_Get(pstrTagValue As String, pstrParameter As String) As String

  Select Case UCase(pstrParameter)
    Case "DATE": TagInfo_Get = Mid(pstrTagValue, 1, 10)
    Case "SESSION": TagInfo_Get = Mid(pstrTagValue, 11, 2)
    Case "BANK_HOLIDAY": TagInfo_Get = Mid(pstrTagValue, 13, 1)
    Case "CAPTION": TagInfo_Get = Mid(pstrTagValue, 14, 2)
    Case "WEEKEND": TagInfo_Get = Mid(pstrTagValue, 16, 1)
    Case "WORKING_DAY": TagInfo_Get = Mid(pstrTagValue, 17, 1)
    Case "WORKING_PATTERN": TagInfo_Get = Mid(pstrTagValue, 18, 14)
    Case "COLOUR": TagInfo_Get = Mid(pstrTagValue, 32, 8)
    Case "HAS_EVENT": TagInfo_Get = Mid(pstrTagValue, 40, 1)
    Case "EVENT_START": TagInfo_Get = Mid(pstrTagValue, 41, 3)
    Case "EVENT_END": TagInfo_Get = Mid(pstrTagValue, 44, 3)
    Case "REGION": TagInfo_Get = Mid(pstrTagValue, 47)

    Case Else: TagInfo_Get = ""
  End Select

End Function


Private Sub cboMonth_Click()
  
  Dim dtShownStart As Date
  Dim dtShownEnd As Date
  Dim blnShowMSG As Boolean
  Dim dtMonth As Date
  Dim lngMonth As Long
  Dim lngYear As Long
  Dim intTempDaysInMonth As Integer
  
  Const strMessage = "The selected date is outside of the report date boundaries."
  
  blnShowMSG = False
  
  If Not mblnLoading Then
   
'    dtShownStart = CDate(CStr("01/" & cboMonth.ItemData(cboMonth.ListIndex) & "/" & spnYear.Value))
'    dtShownEnd = CDate(CStr(DaysInMonth(dtShownStart) & "/" & Month(dtShownStart) & "/" & Year(dtShownStart)))
    lngMonth = cboMonth.ItemData(cboMonth.ListIndex)
    lngYear = spnYear.Value
    
    dtMonth = DateAdd("yyyy", CDbl(lngYear - Year(mdtReportStartDate)), mdtReportStartDate)
    dtMonth = DateAdd("m", CDbl(lngMonth - Month(mdtReportStartDate)), dtMonth)
    
    intTempDaysInMonth = DaysInMonth(dtMonth)
  
    'Define the current visible Start and End Dates.
    dtShownEnd = DateAdd("d", CDbl(intTempDaysInMonth - Day(dtMonth)), dtMonth)
    dtShownStart = DateAdd("d", CDbl(-(intTempDaysInMonth - 1)), dtShownEnd)
  
    If (dtShownStart > mdtReportEndDate) Or (dtShownEnd < mdtReportStartDate) Then
      blnShowMSG = True
      mblnLoading = True
      cboMonth.ListIndex = (mintCurrentMonth - 1)
      spnYear.Value = mlngCurrentYear
      mblnLoading = False
    End If
    
    If blnShowMSG Then
      COAMsgBox strMessage, vbExclamation + vbOKOnly, "Calendar Reports"
    Else
      DateChange
    End If
  End If
  
End Sub

Private Sub chkCaptions_Click()
  If Not mblnLoading Then
    'NHRD15092004 Fault 7043
    RefreshLegend CBool(chkCaptions.Value)
    RefreshDateSpecifics
  End If
End Sub

Private Sub chkIncludeBHols_Click()
  If Not mblnLoading Then
    RefreshDateSpecifics
  End If
End Sub

Private Sub chkIncludeWorkingDaysOnly_Click()
  If Not mblnLoading Then
    RefreshDateSpecifics
  End If
End Sub

Private Sub chkShadeBHols_Click()
  If Not mblnLoading Then
    RefreshDateSpecifics
  End If
End Sub

Private Sub chkShadeWeekends_Click()
  If Not mblnLoading Then
    RefreshDateSpecifics
  End If
End Sub

Private Sub cmdToday_Click()
  
  Dim dtSystemDate As Date
  
  dtSystemDate = Now()
  
  mblnLoading = True
  cboMonth.ListIndex = Month(dtSystemDate) - 1
  spnYear.Value = Year(dtSystemDate)
  mblnLoading = False
  DateChange

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirstMonth_Click()
 
  Dim intListIndex As Integer
  Dim lngYearValue As Long
  Dim lngMonthValue As Long
  
  lngMonthValue = Month(mdtReportStartDate)
  lngYearValue = Year(mdtReportStartDate)
  
  mblnLoading = True
  cboMonth.ListIndex = lngMonthValue - 1
  spnYear.Value = lngYearValue
  mblnLoading = False
  DateChange
  
End Sub

Private Sub cmdLastMonth_Click()
  
  Dim intListIndex As Integer
  Dim lngYearValue As Long
  Dim lngMonthValue As Long
  
  lngMonthValue = Month(mdtReportEndDate)
  lngYearValue = Year(mdtReportEndDate)
  
  mblnLoading = True
  cboMonth.ListIndex = lngMonthValue - 1
  spnYear.Value = lngYearValue
  mblnLoading = False
  DateChange

End Sub

Private Sub cmdNextMonth_Click()
  
  Dim intListIndex As Integer
  Dim lngYearValue As Long
  
  intListIndex = cboMonth.ListIndex
  
  If intListIndex = 11 Then
    lngYearValue = spnYear.Value
    mblnLoading = True
    cboMonth.ListIndex = 0
    spnYear.Value = lngYearValue + 1
    mblnLoading = False
    DateChange
  Else
    cboMonth.ListIndex = intListIndex + 1
  End If
  
End Sub

Private Sub cmdOutput_Click()

  mblnOutputFromPreview = True
  
  OutputReport (True)
  
  Screen.MousePointer = vbDefault

End Sub

Public Function OutputReport(blnPrompt As Boolean) As Boolean
 
  Dim intMonth As Integer
  Dim intMonthCount As Integer
  Dim dtMonth As Date
  Dim fOK As Boolean

  ReDim mstrArray(37, 0)
  ReDim mavOutputDataIndex(2, 0)

  Set mobjOutput = New clsOutputRun
  
  mobjOutput.ShowFormats True, False, True, True, True, False, False
    
  If mobjOutput.SetOptions _
      (blnPrompt, mlngOutputFormat, mblnOutputScreen, _
      mblnOutputPrinter, mstrOutputPrinterName, _
      mblnOutputSave, mlngOutputSaveExisting, _
      mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, _
      mstrOutputEmailAttachAs, mstrOutputFileName) Then

    If mobjOutput.GetFile Then
    
      mobjOutput.HeaderRows = 1
      mobjOutput.HeaderCols = 0
      mobjOutput.SizeColumnsIndependently = True
      
      intMonthCount = DateDiff("m", mdtReportStartDate, mdtReportEndDate)
     
      If Not gblnBatchMode Then
        With gobjProgress
          
        Select Case mlngOutputFormat
          Case fmtExcelChart, fmtExcelPivotTable, fmtExcelWorksheet
            '.AviFile = App.Path & "\videos\excel.avi"
            .AVI = dbExcel
            .MainCaption = "Output to Excel"
          Case fmtWordDoc
            '.AviFile = App.Path & "\videos\word.avi"
            .AVI = dbWord
            .MainCaption = "Output to Word"
          Case fmtHTML
            '.AviFile = App.Path & "\videos\internet.avi"
            .AVI = dbInternet
            .MainCaption = "Output to HTML"
          Case Else
            '.AviFile = App.Path & "\videos\report.avi"
            .AVI = dbText
            .MainCaption = "Output to File"
          End Select
      
          If Not gblnBatchMode Then
            .NumberOfBars = 1
            .Caption = "Calendar Reports"
            .Time = False
            .Cancel = True
            .Bar1Caption = "Outputting " & Me.Caption
            .Bar1MaxValue = ((intMonthCount + 1) * 6)
            .OpenProgress
          End If
          
        End With
      End If
      
      '***********************************************************
      'send an array to the output classes for the legend
      'TM 07/06/06 Fault 10816
      If mobjOutput.Format = fmtExcelWorksheet Or mobjOutput.Format = fmtDataOnly Then
        mobjOutput.AddPage mstrCalendarReportName & " - Key", "Key"
      Else
        mobjOutput.AddPage mstrCalendarReportName, "Key"
      End If
      OutputArray_GetLegendArray
      mobjOutput.DataArrayToGrid mstrLegend(), grdOutput
      
      'if user cancels the report, abort
      If (gobjProgress.Cancelled) Or (mobjOutput.UserCancelled) Then
        If Not gblnBatchMode And (mblnOutputFromPreview = True) Then
          gobjProgress.CloseProgress
          COAMsgBox "Calendar Report Output cancelled by user.", vbExclamation + vbOKOnly, "Calendar Reports Output"
        End If
        
        mblnUserCancelled = True
        
        OutputReport = False
        If Not mobjOutput Is Nothing Then mobjOutput.ClearUp
        Set mobjOutput = Nothing
        Exit Function
      End If
      
      mobjOutput.ResetStyles
      mobjOutput.ResetColumns
      mobjOutput.ResetMerges
      '***********************************************************
      
      mobjOutput.HeaderRows = 2
      mobjOutput.HeaderCols = 1
      mobjOutput.SizeColumnsIndependently = True
  
      For intMonth = 0 To intMonthCount Step 1
  
        'if user cancels the report, abort
        If gobjProgress.Cancelled Or (mobjOutput.UserCancelled) Then
          If Not gblnBatchMode And (mblnOutputFromPreview = True) Then
            gobjProgress.CloseProgress
            COAMsgBox "Calendar Report Output cancelled by user.", vbExclamation + vbOKOnly, "Calendar Reports Output"
          End If
          
          mblnUserCancelled = True
          
          OutputReport = False
          If Not mobjOutput Is Nothing Then mobjOutput.ClearUp
          Set mobjOutput = Nothing
          Exit Function
        End If
  
        Set mcolBaseDescIndex_Output = New Collection
  
'***
        dtMonth = DateAdd("m", intMonth, mdtReportStartDate)
        mlngYear_Output = Year(dtMonth)
        mlngMonth_Output = Month(dtMonth)
  
        mintDaysInMonth_Output = DaysInMonth(dtMonth)
        
        'Define the current visible Start and End Dates.
        mdtVisibleEndDate_Output = DateAdd("d", CDbl(mintDaysInMonth_Output - Day(dtMonth)), dtMonth)
        mdtVisibleStartDate_Output = DateAdd("d", CDbl(-(mintDaysInMonth_Output - 1)), mdtVisibleEndDate_Output)
    
        mintFirstDayOfMonth_Output = Weekday(mdtVisibleStartDate_Output, vbSunday)
'***

        mobjOutput.AddPage mstrCalendarReportName & " - " & MonthName(mlngMonth_Output) & " " & mlngYear_Output, MonthName(mlngMonth_Output) & " " & mlngYear_Output
  
        OutputArray_GetArray
  
        'if user cancels the report, abort
        If gobjProgress.Cancelled Or (mobjOutput.UserCancelled) Then
          If Not gblnBatchMode And (mblnOutputFromPreview = True) Then
            gobjProgress.CloseProgress
            COAMsgBox "Calendar Report Output cancelled by user.", vbExclamation + vbOKOnly, "Calendar Reports Output"
          End If
          
          mblnUserCancelled = True
          
          OutputReport = False
          If Not mobjOutput Is Nothing Then mobjOutput.ClearUp
          Set mobjOutput = Nothing
          Exit Function
        End If
  
        mobjOutput.DataArrayToGrid mstrArray(), grdOutput
  
        ReDim mstrArray(37, 0)
        mobjOutput.ResetColumns
        mobjOutput.ResetStyles
        mobjOutput.ResetMerges
        Set mcolBaseDescIndex_Output = Nothing
  
        If Not gblnBatchMode Then gobjProgress.UpdateProgress gblnBatchMode
  
      Next intMonth

    End If

    'if user cancels the report, abort
    If (gobjProgress.Cancelled) Or (mobjOutput.UserCancelled) Then
      If Not gblnBatchMode And (mblnOutputFromPreview = True) Then
        gobjProgress.CloseProgress
        COAMsgBox "Calendar Report Output cancelled by user.", vbExclamation + vbOKOnly, "Calendar Reports Output"
      End If
          
      mblnUserCancelled = True
          
      OutputReport = False
      If Not mobjOutput Is Nothing Then mobjOutput.ClearUp
      Set mobjOutput = Nothing
      Exit Function
    End If

    mobjOutput.Complete

    If Not gblnBatchMode Then
      If gobjProgress.Visible Then gobjProgress.CloseProgress
    End If

    mstrErrorMessage = mobjOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)
  
  Else
    blnPrompt = (blnPrompt And Not mobjOutput.UserCancelled)
    mstrErrorMessage = mobjOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)
  
  End If

  If blnPrompt Then
    gobjProgress.CloseProgress
    If fOK Then
      COAMsgBox "Calendar Report: '" & mstrCalendarReportName & "' output complete.", _
          vbInformation, "Calendar Report"
    Else
      COAMsgBox "Calendar Report: '" & mstrCalendarReportName & "' output failed." & vbNewLine & vbNewLine & mstrErrorMessage, _
          vbExclamation, "Calendar Report"
    End If
  End If

  If Not mobjOutput Is Nothing Then mobjOutput.ClearUp
  Set mobjOutput = Nothing

  OutputReport = fOK

End Function

Private Sub cmdPrevMonth_Click()
  
  Dim intListIndex As Integer
  Dim lngYearValue As Long
  
  intListIndex = cboMonth.ListIndex
  
  If intListIndex = 0 Then
    lngYearValue = spnYear.Value
    mblnLoading = True
    cboMonth.ListIndex = 11
    spnYear.Value = lngYearValue - 1
    mblnLoading = False
    DateChange
  Else
    cboMonth.ListIndex = intListIndex - 1
  End If

End Sub

Private Sub ctlCalDates_Click(Index As Integer, pvarLabel As Variant)
  
  Dim frmCalendarEventDetails As New frmCalendarReportEventDetails
  Dim colEvents As clsCalendarEvents
  Dim strKey As String
  Dim lngOriginalLeft As Long
  Dim lngOriginalTop As Long
  Dim strDate As String
  Dim strSession As String
  Dim lblTemp As VB.Label
  Dim sngOrigFontSize As Single
  
  Set lblTemp = pvarLabel
  
  If CInt(TagInfo_Get(lblTemp.Tag, "HAS_EVENT")) > 0 Then
    strKey = CStr(pvarLabel.Index)
    Set colEvents = mcolDateControlEvents(CStr(Index & "_CALDATEINDEX_" & strKey))

    frmCalendarEventDetails.ShowRegion = mblnRegions
    frmCalendarEventDetails.ShowWorkingPattern = mblnWorkingPatterns

    If frmCalendarEventDetails.Initialse(colEvents) Then
      
      With lblTemp

        strDate = TagInfo_Get(.Tag, "DATE")
        strSession = TagInfo_Get(.Tag, "SESSION")

        .BorderStyle = vbFixedSingle
        .ZOrder 0

        .Width = CALDATES_BOXWIDTH + 50
        .Height = CALDATES_BOXHEIGHT + 50
        sngOrigFontSize = .Font.Size
        .Font.Size = KEY_FONTSIZE_LARGE
        .Font.Bold = True

        lngOriginalLeft = .Left
        lngOriginalTop = .Top
        
        If UCase(strSession) = "PM" Then
          .Top = lngOriginalTop - 50
        Else
          .Top = lngOriginalTop
        End If
        
        If (pvarLabel.Index = 1) Or (pvarLabel.Index = 2) Then
          .Left = lngOriginalLeft
        ElseIf (pvarLabel.Index = 73) Or (pvarLabel.Index = 74) Then
          .Left = lngOriginalLeft - 50
        Else
          .Left = lngOriginalLeft - 25
        End If
        
        frmCalendarEventDetails.BreakdownCaption = "Calendar Report Breakdown - " & strDate & " " & LCase(strSession)
        frmCalendarEventDetails.Show vbModal

        .Left = lngOriginalLeft
        .Top = lngOriginalTop

        .Font.Size = sngOrigFontSize
        .Font.Bold = False

        .Width = CALDATES_BOXWIDTH
        .Height = CALDATES_BOXHEIGHT

        .BorderStyle = vbBSNone
        .ZOrder 1
      End With

    End If

    Set lblTemp = Nothing
    Set frmCalendarEventDetails = Nothing
  End If

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

  mblnLoading = True
  Me.Width = FORM_STARTWIDTH
  Me.Height = FORM_STARTHEIGHT
  mblnLoading = False
    
  Hook Me.hWnd, FORM_STARTWIDTH, FORM_MINHEIGHT
  
  Set mcolHistoricBankHolidays = New Collection
  Set mcolStaticBankHolidays = New Collection
  Set mcolHistoricWorkingPatterns = New Collection
  Set mcolStaticWorkingPatterns = New Collection
  Set mcolBaseDescIndex = New Collection
  Set mcolDateControlEvents = New Collection
  
  picBase.BackColor = vbButtonFace
  picCalendar.BackColor = vbButtonFace
  picDates.BackColor = vbButtonFace
  picLegend.BackColor = vbWindowBackground
  picLegendScroll.BackColor = vbWindowBackground
  picPrint.BackColor = vbButtonFace
  picScroll.BackColor = vbButtonFace
    
  'output options class, data array and default colours
  Set mobjOutput = New clsOutputRun
  
  ReDim mstrArray(37, 0)
    
  ReDim mvarTableViews(3, 0)
  ReDim mstrViews(0)

  ReDim mavOutputDateIndex(2, 0)
  ReDim mavLegendDateIndex(2, 0)
  
'TM20030407 Fault 5246 - use default output option colors.
  mlngBC_Data = 13434879
  mlngFC_Data = 0
  mlngBC_Heading = 13395456
  mlngFC_Heading = 16777215
  
  lblCalDates(0).BackColor = HexValue(mlngBC_Data)
  lblCalDates(0).ForeColor = HexValue(mlngFC_Data)
  lblDay(0).BackColor = HexValue(mlngBC_Heading)
  lblDay(0).ForeColor = HexValue(mlngFC_Heading)
  lblDate(0).BackColor = HexValue(mlngBC_Heading)
  lblDate(0).ForeColor = HexValue(mlngFC_Heading)

  lblBaseDesc(0).BackColor = lblDay(0).BackColor
  lblBaseDesc(0).ForeColor = lblDay(0).ForeColor
  
  mstrExcludedColours = CStr(mlngBC_Data)
  mlngBC_Data = GetUserSetting("output", "databackcolour", 13434879)
  mstrExcludedColours = mstrExcludedColours & ", " & CStr(mlngBC_Data)
  mlngBC_DataOutput = GetUserSetting("output", "databackcolour", 13434879)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If (y >= Me.picPrint.Top) And (y <= (Me.picPrint.Top + Me.picPrint.Height)) Then
  
  Else
    
  End If
End Sub

Private Sub Form_Resize()

  Dim blnScrollCalendar As Boolean
  Dim blnScrollLegend As Boolean
  Dim lngTop As Long
  Dim lngMinWidth As Long
  Dim lngMaxHeight As Long
  Dim intBaseCount As Integer
  Dim lngBaseWidth As Long
  Dim lngPrintWidth As Long
  Dim intCount As Integer
  Dim lngLeft As Long
  Dim lngWidth As Long
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  If (Me.WindowState = vbMinimized) Or (mblnLoading) Then
    Exit Sub
  End If
  
  lngMinWidth = (BASE_BOXWIDTH + picDates.Width + SCROLLBAR_WIDTH + (2 * CONTROL_OFFSET))
  If (Me.Width < lngMinWidth) Then
    mblnLoading = True
    Me.Width = lngMinWidth
    mblnLoading = False
  End If

  If (Me.Height < FORM_MINHEIGHT) Then
    mblnLoading = True
    Me.Height = FORM_MINHEIGHT
    mblnLoading = False
  End If

  'position the buttons
  With cmdClose
    .Left = (Me.ScaleWidth - CONTROL_OFFSET - .Width)
    .Top = (Me.ScaleHeight - (.Height + CONTROL_OFFSET))
  End With
  
  With cmdOutput
    .Left = (cmdClose.Left - CONTROL_OFFSET - .Width)
    .Top = cmdClose.Top
  End With

  'resize & position options frame
  With fraOptionsShade
    .Height = OPTIONSFRAME_HEIGHT
    .Width = OPTIONSFRAME_WIDTH
    .Left = (Me.ScaleWidth - .Width - CONTROL_OFFSET)
    .Top = (cmdClose.Top - CONTROL_OFFSET - OPTIONSFRAME_HEIGHT)
  End With

  'resize & position the legend frame
  With fraLegend
    .Left = CONTROL_OFFSET
    .Height = fraOptionsShade.Height
    .Top = fraOptionsShade.Top
    If (Me.ScaleWidth - fraOptionsShade.Width - (3 * CONTROL_OFFSET)) < LEGENDFRAME_MINWIDTH Then
      .Width = LEGENDFRAME_MINWIDTH
    Else
      .Width = (Me.ScaleWidth - fraOptionsShade.Width - (3 * CONTROL_OFFSET))
    End If
  End With

  'position the date navigation frame
  With fraDateNav
    .Width = NAVFRAME_WIDTH
    .Height = NAVFRAME_HEIGHT
    .Top = CONTROL_OFFSET
  End With
  
  'resize & position the picPrint & VScrollCalendar controls
  With picPrint
    .Left = CONTROL_OFFSET
    .Top = (fraDateNav.Top + fraDateNav.Height + CONTROL_OFFSET)
    .Height = (fraLegend.Top - .Top - CONTROL_OFFSET)
  End With
  
  'resize & position the controls encapsulated in the picPrint container
  With picDates
    .Top = 0
  End With
  
  With picScroll
    .Height = (picPrint.ScaleHeight - picDates.Height)
    .Top = (picDates.Top + picDates.Height)
    .Left = 0
  End With
  
  With picBase
    .Left = BASEPIC_LEFT
  End With
  
  With picCalendar
    .Left = picDates.Left
    .Width = picDates.Width
  End With
  
  SetScrollBarValues
  
  If (VScrollCalendar.Value = VScrollCalendar.Max) Then
    VScrollCalendar_Change
  End If
  
  If VScrollCalendar.Visible Then
    lngPrintWidth = (Me.ScaleWidth - SCROLLBAR_WIDTH - (2 * CONTROL_OFFSET))
  Else
    lngPrintWidth = (Me.ScaleWidth - (2 * CONTROL_OFFSET))
  End If
  picPrint.Width = lngPrintWidth
  VScrollCalendar.Left = (picPrint.Left + picPrint.Width)
  
  picScroll.Width = picPrint.Width
  
  If (picPrint.ScaleWidth - picDates.Width) >= 0 Then
    lngBaseWidth = (picPrint.ScaleWidth - picDates.Width)
  End If
  
  picBase.Width = lngBaseWidth
  picDates.Left = (picPrint.ScaleWidth - picDates.Width + 15)
  picCalendar.Left = picBase.Width
  For intBaseCount = 1 To lblBaseDesc.UBound Step 1
    lblBaseDesc(intBaseCount).Width = lngBaseWidth
    HorizontalBaseLine(intBaseCount).X2 = lngBaseWidth + 20
    With ctlCalDates(intBaseCount)
      .Width = picCalendar.Width
      .Height = lblBaseDesc(intBaseCount).Height + 20
    End With
    
  Next intBaseCount
  HorizontalBaseLine(HorizontalBaseLine.UBound).X2 = lngBaseWidth
  
  fraDateNav.Left = ((picDates.Left + CONTROL_OFFSET) + ((picDates.Width - fraDateNav.Width) / 2))
  
  lngLeft = ((picLegend.ScaleWidth / 2) + CONTROL_OFFSET)
  lngWidth = ((picLegend.ScaleWidth / 2) - (3 * CONTROL_OFFSET) - LEGEND_BOXWIDTH)
  For intCount = 1 To mintLegendCount Step 1
    lblEventName(intCount).Width = lngWidth
    If (intCount > mintLegendLeft) Then
      lblLegend(intCount).Left = lngLeft
      lblEventName(intCount).Left = (lblLegend(intCount).Left + lblLegend(intCount).Width + CONTROL_OFFSET)
    End If
  Next intCount
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mcolHistoricBankHolidays = Nothing
  Set mcolStaticBankHolidays = Nothing
  Set mcolHistoricWorkingPatterns = Nothing
  Set mcolStaticWorkingPatterns = Nothing
  Set mcolBaseDescIndex = Nothing
  Set mcolDateControlEvents = Nothing
  
  If Not mobjOutput Is Nothing Then mobjOutput.ClearUp
  Set mobjOutput = Nothing
  
  Unhook Me.hWnd
End Sub

Private Sub spnYear_Change()
  
  Dim dtShownStart As Date
  Dim dtShownEnd As Date
  Dim blnShowMSG As Boolean
  Dim dtMonth As Date
  Dim lngMonth As Long
  Dim lngYear As Long
  Dim intTempDaysInMonth As Integer
  
  Const strMessage = "The selected date is outside of the report date boundaries."
  
  blnShowMSG = False
  
  If Not mblnLoading Then
    If (spnYear.Value >= 1900) And (spnYear.Value <= 3000) Then
      
'      dtShownStart = CDate(CStr("01/" & cboMonth.ItemData(cboMonth.ListIndex) & "/" & spnYear.Value))
'      dtShownEnd = CDate(CStr(DaysInMonth(dtShownStart) & "/" & Month(dtShownStart) & "/" & Year(dtShownStart)))
      lngMonth = cboMonth.ItemData(cboMonth.ListIndex)
      lngYear = spnYear.Value
      
      dtMonth = DateAdd("yyyy", CDbl(lngYear - Year(mdtReportStartDate)), mdtReportStartDate)
      dtMonth = DateAdd("m", CDbl(lngMonth - Month(mdtReportStartDate)), dtMonth)
      
      intTempDaysInMonth = DaysInMonth(dtMonth)
    
      'Define the current visible Start and End Dates.
      dtShownEnd = DateAdd("d", CDbl(intTempDaysInMonth - Day(dtMonth)), dtMonth)
      dtShownStart = DateAdd("d", CDbl(-(intTempDaysInMonth - 1)), dtShownEnd)
    
      If (dtShownStart > mdtReportEndDate) Or (dtShownEnd < mdtReportStartDate) Then
        blnShowMSG = True
        mblnLoading = True
        cboMonth.ListIndex = (mintCurrentMonth - 1)
        spnYear.Value = mlngCurrentYear
        mblnLoading = False
      End If
      
      If blnShowMSG Then
        COAMsgBox strMessage, vbExclamation + vbOKOnly, "Calendar Reports"
      Else
        DateChange
      End If
    
    End If
  End If
  
End Sub

Private Sub VScrollCalendar_Change()
  
'  If (CLng(VScrollCalendar.Value) * mlngScrollBarMultiplier) > (picCalendar.Height - picScroll.Height) Then
'    picCalendar.Top = -(picCalendar.Height - picScroll.Height)
'    picBase.Top = -(picBase.Height - picScroll.Height)
'  Else
    picCalendar.Top = (CLng(VScrollCalendar.Value) * -mlngScrollBarMultiplier)
    picBase.Top = (CLng(VScrollCalendar.Value) * -mlngScrollBarMultiplier)
'  End If
End Sub

Private Function CheckPermission_RegionInfo() As Boolean

  Dim strTableColumn As String
  
  
  'Check the  Bank Holiday Region Table - Region Table
  '           Bank Holiday Region Table - Region Column
  '           Bank Holidays Table - Bank Holiday Table
  '           Bank Holidays Table - Date Column
  '           Bank Holidays Table - Descripiton Column
  '...Bank Holiday module setup information.
  'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
  If gsBHolRegionTableName = "" Or _
     gsBHolRegionColumnName = "" Or _
     gsBHolTableName = "" Or _
     gsBHolDateColumnName = "" Or _
     gsBHolDescriptionColumnName = "" Then
     
    GoTo DisableRegions
  End If
   
  'Check the  Career Change Region - Static Region Column
  '           Career Change Region - Historic Region Table
  '           Career Change Region - Historic Region Column
  '           Career Change Region - Historic Region Effective Date Column
  '...Personnel - Career Change module setup information.
  'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
  If gsPersonnelRegionColumnName = "" Then
    If gsPersonnelHRegionTableName = "" Or _
       gsPersonnelHRegionColumnName = "" Or _
       gsPersonnelHRegionDateColumnName = "" Then
       
      GoTo DisableRegions
    End If
  End If




  '*******************************************************************
  ' All Region module information is setup                           *
  ' Now check the permissions on the Region module setup information *
  '*******************************************************************
  'Bank Holiday Region Table - Region Table (Regional Information)
  'Bank Holiday Region Table - Region Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngBHolRegionTableID, gsBHolRegionTableName, _
                            gsBHolRegionColumnName, strTableColumn) Then
    mstrSQLSelect_RegInfoRegion = strTableColumn
    strTableColumn = vbNullString
  Else
    GoTo DisableRegions
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
 
  'Bank Holidays Table - Bank Holiday Table (Region History)
  'Bank Holidays Table - Date Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, _
                            gsBHolDateColumnName, strTableColumn) Then
    mstrSQLSelect_BankHolDate = strTableColumn
    strTableColumn = vbNullString
  Else
    GoTo DisableRegions
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  
  'Bank Holidays Table - Bank Holiday Table (Region History)
  'Bank Holidays Table - Descripiton Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, _
                            gsBHolDescriptionColumnName, strTableColumn) Then
    mstrSQLSelect_BankHolDesc = strTableColumn
    strTableColumn = vbNullString
  Else
    GoTo DisableRegions
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\




  '*******************************************************************
  ' Permission granted on all Region module information.             *
  ' Now check the permissions on the                                 *
  ' Personnel Career Change Region module setup information          *
  '*******************************************************************
  If mlngStaticRegionColumnID > 0 Then
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(mlngBaseTableID, mstrBaseTableName, _
                              mstrStaticRegionColumn, strTableColumn) Then
      mstrSQLSelect_PersonnelStaticRegion = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

  Else
    'Check Career Change Region access
    If gsPersonnelRegionColumnName <> "" Then
      'Personnel Table
      'Career Change Region - Static Region Column
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, _
                                gsPersonnelRegionColumnName, strTableColumn) Then
        mstrSQLSelect_PersonnelStaticRegion = strTableColumn
        strTableColumn = vbNullString
      Else
        GoTo DisableRegions
      End If
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    Else
      'Career Change Region - Historic Region Table
      'Career Change Region - Historic Region Column
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, _
                                gsPersonnelHRegionColumnName, strTableColumn) Then
        mstrSQLSelect_PersonnelHRegion = strTableColumn
        strTableColumn = vbNullString
      Else
        GoTo DisableRegions
      End If
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
      'Career Change Region - Historic Region Table
      'Career Change Region - Historic Region Effective Date Column
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, _
                                gsPersonnelHRegionDateColumnName, strTableColumn) Then
        mstrSQLSelect_PersonnelHDate = strTableColumn
        strTableColumn = vbNullString
      Else
        GoTo DisableRegions
      End If
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
    End If
  End If
  CheckPermission_RegionInfo = True

TidyUpAndExit:
  Exit Function

DisableRegions:
  mblnDisableRegions = True
  ShowBankHolidays = False
  IncludeBankHolidays = False
  mblnShowBankHols = False
  mblnRegions = False
  CheckPermission_RegionInfo = False
  GoTo TidyUpAndExit
  
End Function
Private Function CheckPermission_WPInfo() As Boolean
 
  Dim objTable As CTablePrivilege
  Dim objColumn As CColumnPrivileges
  Dim pblnColumnOK As Boolean
  Dim strTableColumn As String
  
  'Check the  Career Change Working Pattern - Static Working Pattern Column
  '           Career Change Working Pattern - Historic Working Pattern Table
  '           Career Change Working Pattern - Historic Working Pattern Column
  '           Career Change Working Pattern - Historic Working Pattern Effective Date Column
  '...Personnel - Career Change module setup information.
  'If any are blank then we need to allow the report to run, but disable the Working Dys Display Option.
  If gsPersonnelWorkingPatternColumnName = "" Then
    If gsPersonnelHWorkingPatternTableName = "" Or _
       gsPersonnelHWorkingPatternColumnName = "" Or _
       gsPersonnelHWorkingPatternDateColumnName = "" Then
       
      GoTo DisableWPs
    End If
  End If
   
  '****************************************************************************
  ' All Working Pattern module information is setup                           *
  ' Now check the permissions on the Working Pattern module setup information *
  '****************************************************************************
  
  'Check Career Change Working Pattern access
  If gsPersonnelWorkingPatternColumnName <> "" Then
    'Career Change Working Pattern - Static Working Pattern Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, _
                              gsPersonnelWorkingPatternColumnName, strTableColumn) Then
      mstrSQLSelect_PersonnelStaticWP = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableWPs
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
  Else
    'Career Change Working Pattern - Historic Working Pattern Table
    Set objColumn = GetColumnPrivileges(gsPersonnelHWorkingPatternTableName)
    
    'Career Change Working Pattern - Historic Working Pattern Column
    pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternColumnName)
    If pblnColumnOK Then
      pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternColumnName).AllowSelect
    End If
    If pblnColumnOK = False Then
      GoTo DisableWPs
    End If

    'Career Change Working Pattern - Historic Working Pattern Effective Date Column
    pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternDateColumnName)
    If pblnColumnOK Then
      pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternDateColumnName).AllowSelect
    End If
    If pblnColumnOK = False Then
      GoTo DisableWPs
    End If

  End If

  CheckPermission_WPInfo = True
  
TidyUpAndExit:
  Set objTable = Nothing
  Set objColumn = Nothing
  Exit Function

DisableWPs:
  mblnDisableWPs = True
  IncludeWorkingDaysOnly = False
  mblnWorkingPatterns = False
  CheckPermission_WPInfo = False
  GoTo TidyUpAndExit

End Function

Private Sub VScrollLegend_Change()
  picLegend.Top = VScrollLegend.Value * -1
End Sub

Public Property Let DescriptionSeparator(pstrNewValue As String)
  mstrDescriptionSeparator = pstrNewValue
End Property

Public Property Let Description1ID(plngNewValue As Long)
  mlngDescription1ID = plngNewValue
End Property



Public Property Let Description2ID(plngNewValue As Long)
  mlngDescription2ID = plngNewValue
End Property

Public Property Let DescriptionExprID(plngNewValue As Long)
  mlngDescriptionExprID = plngNewValue
End Property

Public Sub HookFormSizes()
  Dim lngMinWidth As Long
  lngMinWidth = (BASE_BOXWIDTH + picDates.Width + SCROLLBAR_WIDTH + (2 * CONTROL_OFFSET))
  Unhook Me.hWnd
  Hook Me.hWnd, lngMinWidth, FORM_MINHEIGHT
End Sub

