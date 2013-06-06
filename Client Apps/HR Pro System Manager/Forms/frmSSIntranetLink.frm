VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmSSIntranetLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Link"
   ClientHeight    =   12660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5041
   Icon            =   "frmSSIntranetLink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12660
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraChartLink 
      Caption         =   "Chart :"
      Height          =   3315
      Left            =   2880
      TabIndex        =   61
      Top             =   8310
      Width           =   6300
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2505
         Left            =   2730
         OleObjectBlob   =   "frmSSIntranetLink.frx":000C
         TabIndex        =   69
         Top             =   555
         Width           =   3330
      End
      Begin VB.CheckBox chkShowValues 
         Caption         =   "S&how Values"
         Height          =   210
         Left            =   195
         TabIndex        =   66
         Top             =   1695
         Width           =   1665
      End
      Begin VB.CommandButton cmdChartData 
         Caption         =   "Data..."
         Height          =   375
         Left            =   180
         TabIndex        =   68
         Top             =   2355
         Width           =   1200
      End
      Begin VB.CheckBox chkStackSeries 
         Caption         =   "S&tack Series"
         Height          =   210
         Left            =   210
         TabIndex        =   67
         Top             =   2040
         Width           =   1665
      End
      Begin VB.CheckBox chkDottedGridlines 
         Caption         =   "Dotted &Gridlines"
         Height          =   195
         Left            =   195
         TabIndex        =   65
         Top             =   1350
         Width           =   1980
      End
      Begin VB.CheckBox chkShowLegend 
         Caption         =   "Show &Legend"
         Height          =   240
         Left            =   195
         TabIndex        =   64
         Top             =   990
         Width           =   1710
      End
      Begin VB.ComboBox cboChartType 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":24FC
         Left            =   195
         List            =   "frmSSIntranetLink.frx":24FE
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   555
         Width           =   2205
      End
      Begin VB.Label lblChartyType 
         AutoSize        =   -1  'True
         Caption         =   "Chart Type :"
         Height          =   195
         Left            =   195
         TabIndex        =   62
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraLinkSeparator 
      Caption         =   "Separator :"
      Height          =   1875
      Left            =   2880
      TabIndex        =   55
      Top             =   7440
      Width           =   6300
      Begin VB.CommandButton cmdIcon 
         Caption         =   "..."
         Height          =   315
         Left            =   4830
         TabIndex        =   58
         Top             =   315
         Width           =   315
      End
      Begin VB.TextBox txtIcon 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1050
         TabIndex        =   57
         Top             =   300
         Width           =   3765
      End
      Begin VB.CommandButton cmdIconClear 
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
         Left            =   5160
         MaskColor       =   &H000000FF&
         TabIndex        =   59
         ToolTipText     =   "Clear Path"
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chkNewColumn 
         Caption         =   "Column &break"
         Height          =   255
         Left            =   1050
         TabIndex        =   60
         Top             =   690
         Width           =   2040
      End
      Begin VB.Label lblNoOptions 
         AutoSize        =   -1  'True
         Caption         =   "There are no configurable options for this link type"
         Height          =   195
         Left            =   210
         TabIndex        =   70
         Top             =   1110
         Visible         =   0   'False
         Width           =   4350
      End
      Begin VB.Label lblIcon 
         Caption         =   "Icon :"
         Height          =   195
         Left            =   210
         TabIndex        =   56
         Top             =   345
         Width           =   615
      End
      Begin VB.Image imgIcon 
         Height          =   495
         Left            =   5565
         Stretch         =   -1  'True
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.Frame fraHRProUtilityLink 
      Caption         =   "HR Pro Report / Utility :"
      Height          =   1485
      Left            =   2880
      TabIndex        =   29
      Top             =   6180
      Width           =   6300
      Begin VB.ComboBox cboHRProUtility 
         Height          =   315
         Left            =   1400
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   700
         Width           =   4700
      End
      Begin VB.ComboBox cboHRProUtilityType 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":2500
         Left            =   1400
         List            =   "frmSSIntranetLink.frx":2502
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   300
         Width           =   4700
      End
      Begin VB.Label lblHRProUtilityMessage 
         AutoSize        =   -1  'True
         Caption         =   "<message>"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1395
         TabIndex        =   34
         Top             =   1160
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHRProUtility 
         Caption         =   "Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   32
         Top             =   765
         Width           =   780
      End
      Begin VB.Label lblHRProUtilityType 
         Caption         =   "Type :"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame fraEmailLink 
      Caption         =   "Email Link :"
      Height          =   1245
      Left            =   2880
      TabIndex        =   45
      Top             =   5310
      Width           =   6300
      Begin VB.TextBox txtEmailSubject 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   49
         Top             =   700
         Width           =   4515
      End
      Begin VB.TextBox txtEmailAddress 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   47
         Top             =   300
         Width           =   4515
      End
      Begin VB.Label lblEmailSubject 
         AutoSize        =   -1  'True
         Caption         =   "Email Subject :"
         Height          =   195
         Left            =   195
         TabIndex        =   48
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label lblEMailAddress 
         AutoSize        =   -1  'True
         Caption         =   "Email Address :"
         Height          =   195
         Left            =   195
         TabIndex        =   46
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Frame fraURLLink 
      Caption         =   "URL :"
      Height          =   1125
      Left            =   2880
      TabIndex        =   35
      Top             =   4470
      Width           =   6300
      Begin VB.TextBox txtURL 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   37
         Top             =   300
         Width           =   4515
      End
      Begin VB.CheckBox chkNewWindow 
         Caption         =   "&Display in new window"
         Height          =   330
         Left            =   1575
         TabIndex        =   38
         Top             =   690
         Width           =   2685
      End
      Begin VB.Label lblURL 
         Caption         =   "URL :"
         Height          =   195
         Left            =   195
         TabIndex        =   36
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame fraDocument 
      Caption         =   "Document :"
      Height          =   1125
      Left            =   165
      TabIndex        =   50
      Top             =   5895
      Width           =   9060
      Begin VB.TextBox txtDocumentFilePath 
         Height          =   315
         Left            =   1400
         MaxLength       =   500
         TabIndex        =   52
         Top             =   300
         Width           =   7365
      End
      Begin VB.CheckBox chkDisplayDocumentHyperlink 
         Caption         =   "Displa&y hyperlink to document"
         Height          =   330
         Left            =   1395
         TabIndex        =   53
         Top             =   690
         Width           =   3720
      End
      Begin VB.Label lblDocumentFilePath 
         AutoSize        =   -1  'True
         Caption         =   "URL :"
         Height          =   195
         Left            =   195
         TabIndex        =   51
         Top             =   360
         Width           =   390
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraApplicationLink 
      Caption         =   "Application :"
      Height          =   1245
      Left            =   2880
      TabIndex        =   39
      Top             =   3585
      Width           =   6300
      Begin VB.CommandButton cmdAppFilePathSel 
         Caption         =   "..."
         Height          =   315
         Left            =   5760
         TabIndex        =   42
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtAppFilePath 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   41
         Top             =   300
         Width           =   4185
      End
      Begin VB.TextBox txtAppParameters 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   44
         Top             =   700
         Width           =   4515
      End
      Begin VB.Label lblAppFilePath 
         AutoSize        =   -1  'True
         Caption         =   "File Path :"
         Height          =   195
         Left            =   195
         TabIndex        =   40
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblAppParameters 
         AutoSize        =   -1  'True
         Caption         =   "Parameters :"
         Height          =   195
         Left            =   195
         TabIndex        =   43
         Top             =   765
         Width           =   930
      End
   End
   Begin VB.Frame fraLinkType 
      Caption         =   "Link Type :"
      Height          =   3645
      Left            =   150
      TabIndex        =   9
      Top             =   1920
      Width           =   2500
      Begin VB.OptionButton optLink 
         Caption         =   "&Database Value"
         Height          =   450
         Index           =   7
         Left            =   195
         TabIndex        =   17
         Top             =   2685
         Width           =   2235
      End
      Begin VB.OptionButton optLink 
         Caption         =   "Pending &Workflows"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   18
         Top             =   3150
         Width           =   2235
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Chart"
         Height          =   315
         Index           =   6
         Left            =   195
         TabIndex        =   16
         Top             =   2385
         Width           =   1305
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Separator"
         Height          =   315
         Index           =   5
         Left            =   195
         TabIndex        =   15
         Top             =   2040
         Width           =   1305
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&On-screen Document Display"
         Height          =   450
         Index           =   9
         Left            =   200
         TabIndex        =   19
         Top             =   3420
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Application"
         Height          =   315
         Index           =   4
         Left            =   200
         TabIndex        =   14
         Top             =   1700
         Width           =   1545
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Email Link"
         Height          =   315
         Index           =   3
         Left            =   200
         TabIndex        =   13
         Top             =   1350
         Width           =   1305
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&HR Pro Screen"
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&URL"
         Height          =   315
         Index           =   1
         Left            =   200
         TabIndex        =   12
         Top             =   1000
         Width           =   960
      End
      Begin VB.OptionButton optLink 
         Caption         =   "HR &Pro Report / Utility"
         Height          =   315
         Index           =   2
         Left            =   200
         TabIndex        =   11
         Top             =   650
         Width           =   2265
      End
   End
   Begin VB.Frame fraHRProScreenLink 
      Caption         =   "HR Pro Screen :"
      Height          =   3645
      Left            =   2880
      TabIndex        =   20
      Top             =   1920
      Width           =   6300
      Begin VB.TextBox txtPageTitle 
         Height          =   315
         Left            =   1575
         MaxLength       =   100
         TabIndex        =   26
         Top             =   1100
         Width           =   4515
      End
      Begin VB.ComboBox cboHRProScreen 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   700
         Width           =   4515
      End
      Begin VB.ComboBox cboHRProTable 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   4515
      End
      Begin VB.ComboBox cboStartMode 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":2504
         Left            =   1575
         List            =   "frmSSIntranetLink.frx":2506
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1500
         Width           =   4515
      End
      Begin VB.Label lblPageTitle 
         AutoSize        =   -1  'True
         Caption         =   "Page Title :"
         Height          =   195
         Left            =   200
         TabIndex        =   25
         Top             =   1160
         Width           =   810
      End
      Begin VB.Label lblHRProScreen 
         Caption         =   "Screen :"
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblHRProTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblStartMode 
         AutoSize        =   -1  'True
         Caption         =   "Start Mode :"
         Height          =   195
         Left            =   200
         TabIndex        =   27
         Top             =   1560
         Width           =   900
      End
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   6600
      TabIndex        =   54
      Top             =   12150
      Width           =   2600
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   135
         TabIndex        =   83
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   84
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraLink 
      Caption         =   "Link :"
      Height          =   1710
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   9000
      Begin VB.ComboBox cboTableView 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":2508
         Left            =   1485
         List            =   "frmSSIntranetLink.frx":250A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1100
         Width           =   3030
      End
      Begin VB.TextBox txtText 
         Height          =   315
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   4
         Top             =   700
         Width           =   3030
      End
      Begin VB.TextBox txtPrompt 
         Height          =   315
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   2
         Top             =   300
         Width           =   3030
      End
      Begin SSDataWidgets_B.SSDBGrid grdAccess 
         Height          =   1230
         Left            =   5595
         TabIndex        =   8
         Top             =   300
         Width           =   3180
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   2
         stylesets(0).Name=   "SysSecMgr"
         stylesets(0).ForeColor=   -2147483631
         stylesets(0).BackColor=   -2147483633
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmSSIntranetLink.frx":250C
         stylesets(1).Name=   "ReadOnly"
         stylesets(1).ForeColor=   -2147483631
         stylesets(1).BackColor=   -2147483633
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "frmSSIntranetLink.frx":2528
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   3889
         Columns(0).Caption=   "User Group"
         Columns(0).Name =   "GroupName"
         Columns(0).AllowSizing=   0   'False
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   1244
         Columns(1).Caption=   "Visible"
         Columns(1).Name =   "Access"
         Columns(1).CaptionAlignment=   2
         Columns(1).AllowSizing=   0   'False
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   11
         Columns(1).FieldLen=   256
         Columns(1).Style=   2
         TabNavigation   =   1
         _ExtentX        =   5609
         _ExtentY        =   2170
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblAccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visibility :"
         Height          =   195
         Left            =   4665
         TabIndex        =   7
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblTableView 
         Caption         =   "Table (View) :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   1155
         Width           =   1245
      End
      Begin VB.Label lblText 
         Caption         =   "Text :"
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   760
         Width           =   615
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Prompt :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame fraDBValue 
      Caption         =   "Database Value :"
      Height          =   1875
      Left            =   2880
      TabIndex        =   71
      Top             =   9990
      Width           =   6300
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   1395
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   705
         Width           =   3930
      End
      Begin VB.ComboBox cboParents 
         Height          =   315
         Left            =   1395
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   315
         Width           =   3930
      End
      Begin VB.OptionButton optAggregateType 
         Caption         =   "Count"
         Height          =   285
         Index           =   0
         Left            =   2265
         TabIndex        =   81
         Top             =   1545
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAggregateType 
         Caption         =   "Total"
         Height          =   285
         Index           =   1
         Left            =   3390
         TabIndex        =   82
         Top             =   1545
         Width           =   765
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   4650
         TabIndex        =   78
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtFilter 
         Height          =   330
         Left            =   1395
         TabIndex        =   77
         Top             =   1080
         Width           =   3225
      End
      Begin VB.CommandButton cmdFilterClear 
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
         Left            =   4965
         MaskColor       =   &H000000FF&
         TabIndex        =   79
         ToolTipText     =   "Clear Path"
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label lblParents 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
         Height          =   195
         Left            =   210
         TabIndex        =   72
         Top             =   330
         Width           =   600
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         Caption         =   "Column :"
         Height          =   195
         Left            =   210
         TabIndex        =   74
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter :"
         Height          =   195
         Left            =   210
         TabIndex        =   76
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label lblAggregateType 
         AutoSize        =   -1  'True
         Caption         =   "Aggregate Function :"
         Height          =   195
         Left            =   210
         TabIndex        =   80
         Top             =   1575
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmSSIntranetLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SSINTRANETSCREENTYPES
  SSINTLINKSCREEN_HRPRO = 0
  SSINTLINKSCREEN_URL = 1
  SSINTLINKSCREEN_UTILITY = 2
  'NPG20080125 Fault 12873
  SSINTLINKSCREEN_EMAIL = 3
  SSINTLINKSCREEN_APPLICATION = 4
  'NPG Dashboard
  SSINTLINKSEPARATOR = 5
  SSINTLINKCHART = 6
  SSINTLINKDB_VALUE = 7
  SSINTLINKPWFSTEPS = 8
  SSINTLINKSCREEN_DOCUMENT = 9
End Enum

Private mblnCancelled As Boolean
Private miLinkType As SSINTRANETLINKTYPES
Private mfChanged As Boolean
Private mfLoading As Boolean
'Private mlngPersonnelTableID As Long
Private mblnRefreshing As Boolean
Private mlngTableID As Long
Private mlngColumnID As Long
Private mlngViewID As Long
Private msTableViewName As String
Private glngPictureID As Long
Private miSeparatorOrientation As Integer
Private mfNewWindow As Boolean
Private miChartType As Integer
Private miChartViewID As Long
Private miChartFilterID As Long
Private miChartTableID As Long
Private miChartColumnID As Long
Private miChartAggregateType As Integer
Private miElementType As Integer
Private msCombinedHiddenGroups As String
Private mblnReadOnly As Boolean

Private mcolSSITableViews As clsSSITableViews

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Private Sub FormatScreen()

  Const GAPBETWEENTEXTBOXES = 85
  Const GAPABOVEBUTTONS = 150
  Const GAPUNDERBUTTONS = 600
  Const LEFTGAP = 200
  Const GAPUNDERLASTCONTROL = 200
  Const GAPUNDERRADIOBUTTON = -15
  
  Select Case miLinkType
    Case SSINTLINK_BUTTON
      fraLink.Caption = "Dashboard Link :"
    Case SSINTLINK_DROPDOWNLIST
      fraLink.Caption = "Dropdown List Link :"
    Case SSINTLINK_HYPERTEXT
      fraLink.Caption = "Hypertext Link :"
    Case SSINTLINK_DOCUMENT
      fraLink.Caption = "On-screen Document Display :"
  End Select
  
  ' Prompt only required for Button Links.
  lblPrompt.Visible = (miLinkType = SSINTLINK_BUTTON)
  txtPrompt.Visible = lblPrompt.Visible

  ' Reposition the Text controls if required.
  If (miLinkType <> SSINTLINK_BUTTON) Then
    lblText.Top = lblPrompt.Top
    txtText.Top = txtPrompt.Top
  End If
  
  cboTableView.Top = txtText.Top + txtText.Height + GAPBETWEENTEXTBOXES
  lblTableView.Top = cboTableView.Top + (lblText.Top - txtText.Top)
 
  ' HR Pro screen links only required for Button or Dropdown List Links
  optLink(SSINTLINKSCREEN_HRPRO).Enabled = (miLinkType <> SSINTLINK_HYPERTEXT) And (miLinkType <> SSINTLINK_DOCUMENT)
  
  If miLinkType = SSINTLINK_DOCUMENT Then
    fraLinkType.Visible = False
    fraDocument.Visible = True
    fraDocument.Top = fraLinkType.Top
    fraDocument.Left = fraLinkType.Left
    fraDocument.Width = fraLink.Width
  Else
    With fraHRProScreenLink
      fraHRProUtilityLink.Top = .Top
      fraHRProUtilityLink.Height = .Height
      fraHRProUtilityLink.Left = .Left
      
      fraURLLink.Top = .Top
      fraURLLink.Height = .Height
      fraURLLink.Left = .Left
      
      fraEmailLink.Top = .Top
      fraEmailLink.Height = .Height
      fraEmailLink.Left = .Left
      
      fraApplicationLink.Top = .Top
      fraApplicationLink.Height = .Height
      fraApplicationLink.Left = .Left
      
      fraLinkSeparator.Top = .Top
      fraLinkSeparator.Height = .Height
      fraLinkSeparator.Left = .Left
      
      fraChartLink.Top = .Top
      fraChartLink.Height = .Height
      fraChartLink.Left = .Left
      
      fraDBValue.Top = .Top
      fraDBValue.Height = .Height
      fraDBValue.Left = .Left
    End With
  End If

  ' Position the OK/Cancel buttons
  If (miLinkType = SSINTLINK_DOCUMENT) Then
    fraOKCancel.Top = fraDocument.Top + fraDocument.Height + GAPABOVEBUTTONS
  Else
    fraOKCancel.Top = fraLinkType.Top + fraLinkType.Height + GAPABOVEBUTTONS
  End If
  
  ' Redimension the form.
  Me.Height = fraOKCancel.Top + fraOKCancel.Height + GAPUNDERBUTTONS
  
End Sub


Private Sub SetChartTypes()

  ' Populate the Chart Types combo.
  Dim iDefaultItem As Integer
  
  iDefaultItem = 0
  
  With cboChartType
    .Clear
      
    .AddItem "3D Bar"
    .ItemData(.NewIndex) = 0
        
    .AddItem "2D Bar"
    .ItemData(.NewIndex) = 1
    
'    .AddItem "3D Line"
'    .ItemData(.NewIndex) = 2
        
'    .AddItem "2D Line"
'    .ItemData(.NewIndex) = 3
        
    .AddItem "3D Area"
    .ItemData(.NewIndex) = 4
        
'    .AddItem "2D Area"
'    .ItemData(.NewIndex) = 5
        
    .AddItem "3D Step"
    .ItemData(.NewIndex) = 6
        
    .AddItem "2D Step"
    .ItemData(.NewIndex) = 7
        
    .AddItem "2D Pie"
    .ItemData(.NewIndex) = 14
        
'    .AddItem "2D XY"
'    .ItemData(.NewIndex) = 16
        
    .ListIndex = iDefaultItem
  End With
 
End Sub


Private Sub GetHRProUtilityTypes()

  ' Populate the Utility Types combo.
  Dim iDefaultItem As Integer
  
  iDefaultItem = 0
  
  With cboHRProUtilityType
    .Clear
      
    .AddItem "Calendar Report"
    .ItemData(.NewIndex) = utlCalendarReport
        
    .AddItem "Custom Report"
    .ItemData(.NewIndex) = utlCustomReport
    
    .AddItem "Mail Merge"
    .ItemData(.NewIndex) = utlMailMerge
    
    If ASRDEVELOPMENT Or Application.WorkflowModule Then
      .AddItem "Workflow"
      .ItemData(.NewIndex) = utlWorkflow
    End If
    
    .ListIndex = iDefaultItem
  End With
 
End Sub

Private Sub GetHRProTables()

  ' Populate the tables combo.
  Dim sSQL As String
  Dim rsTables As dao.Recordset
  Dim iDefaultItem As Integer
  
  iDefaultItem = 0
  
  If (miLinkType <> SSINTLINK_HYPERTEXT) And (miLinkType <> SSINTLINK_DOCUMENT) Then
    cboHRProTable.Clear
      
    If mlngTableID > 0 Then
      ' Add the table and its children (not grand children).
      sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
        " FROM tmpTables" & _
        " WHERE (tmpTables.deleted = FALSE)" & _
        " AND ((tmpTables.tableID = " & CStr(mlngTableID) & ")" & _
        " OR (tmpTables.tableID IN (SELECT childID FROM tmpRelations WHERE parentID =" & CStr(mlngTableID) & ")))" & _
        " ORDER BY tmpTables.tableName"
      Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsTables.EOF
        cboHRProTable.AddItem rsTables!TableName
        cboHRProTable.ItemData(cboHRProTable.NewIndex) = rsTables!TableID
        
        If mlngTableID = rsTables!TableID Then
          iDefaultItem = cboHRProTable.NewIndex
        End If
        
        rsTables.MoveNext
      Wend
      rsTables.Close
      Set rsTables = Nothing
    End If
    
    If cboHRProTable.ListCount = 0 Then
      optLink(SSINTLINKSCREEN_UTILITY).value = True
      optLink(SSINTLINKSCREEN_HRPRO).Enabled = False
      GetHRProScreens
    Else
      cboHRProTable.ListIndex = iDefaultItem
    End If
  
    GetStartModes
  End If
  
End Sub

Private Sub GetHRProUtilities(pUtilityType As UtilityType)

  ' Populate the utilities combo.
  Dim sSQL As String
  Dim sWhereSQL As String
  Dim rsUtilities As New ADODB.Recordset
  Dim rsLocalUtilities As dao.Recordset
  Dim sTableName As String
  Dim sIDColumnName As String
  Dim fLocalTable As Boolean
  
  fLocalTable = False
  
  cboHRProUtility.Clear

  Select Case pUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sIDColumnName = "ID"
      
    Case utlCalendarReport
      sTableName = "ASRSysCalendarReports"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sIDColumnName = "CrossTabID"
    
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sIDColumnName = "ID"
    
    Case utlDataTransfer
      sTableName = "ASRSysDataTransferName"
      sIDColumnName = "DataTransferID"
      
    Case utlExport
      sTableName = "ASRSysExportName"
      sIDColumnName = "ID"
      
    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTableName = "ASRSysGlobalFunctions"
      sIDColumnName = "functionID"

    Case utlImport
      sTableName = "ASRSysImportName"
      sIDColumnName = "ID"

    Case utlLabel
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
      sWhereSQL = "ASRSysMailMergeName.IsLabel = 1 "
    
    Case utlMailMerge
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
      sWhereSQL = "ASRSysMailMergeName.IsLabel = 0 "

    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sIDColumnName = "recordProfileID"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTableName = "ASRSysMatchReportName"
      sIDColumnName = "matchReportID"
  
    Case utlWorkflow
      fLocalTable = True
      sTableName = "tmpWorkflows"
      sIDColumnName = "ID"
      sWhereSQL = "tmpWorkflows.initiationType = " & CStr(WORKFLOWINITIATIONTYPE_MANUAL) & _
        " OR tmpWorkflows.initiationType is null"
  End Select
  
  If Len(sTableName) > 0 Then
    ' Get the available utilities of the given type.
    If fLocalTable Then
      sSQL = "SELECT " & sTableName & "." & sIDColumnName & " AS [ID], " & sTableName & ".name" & _
        " FROM " & sTableName & _
        " WHERE (" & sTableName & ".deleted = FALSE)"
      If Len(sWhereSQL) > 0 Then
        sSQL = sSQL & _
          " AND (" & sWhereSQL & ")"
      End If
      sSQL = sSQL & _
        " ORDER BY " & sTableName & ".name"
      Set rsLocalUtilities = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsLocalUtilities.EOF
        cboHRProUtility.AddItem rsLocalUtilities!Name
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsLocalUtilities!id

        rsLocalUtilities.MoveNext
      Wend
      rsLocalUtilities.Close
      Set rsLocalUtilities = Nothing
    Else
      sSQL = "SELECT" & _
        "  " & sTableName & ".name," & _
        "  " & sTableName & "." & sIDColumnName & " AS [ID]" & _
        "  FROM " & sTableName
      If Len(sWhereSQL) > 0 Then
        sSQL = sSQL & _
          " WHERE (" & sWhereSQL & ")"
      End If
      sSQL = sSQL & _
        "  ORDER BY " & sTableName & ".name"
      
      rsUtilities.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      While Not rsUtilities.EOF
        cboHRProUtility.AddItem rsUtilities!Name
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsUtilities!id
  
        rsUtilities.MoveNext
      Wend
      rsUtilities.Close
      Set rsUtilities = Nothing
    End If
  End If
  
  If cboHRProUtility.ListCount > 0 Then
    cboHRProUtility.ListIndex = 0
  End If

End Sub

Private Sub GetHRProScreens()

  ' Populate the screens combo.
  Dim sSQL As String
  Dim rsScreens As dao.Recordset

  If miLinkType <> SSINTLINK_HYPERTEXT Then
    cboHRProScreen.Clear

    If cboHRProTable.ListIndex >= 0 Then
      ' Add any SS Int screens for the seledcted table.
      sSQL = "SELECT tmpScreens.screenID, tmpScreens.name" & _
        " FROM tmpScreens" & _
        " WHERE (tmpScreens.deleted = FALSE)" & _
        " AND (tmpScreens.ssIntranet = TRUE)" & _
        " AND (tmpScreens.tableID = " & CStr(cboHRProTable.ItemData(cboHRProTable.ListIndex)) & ")" & _
        " ORDER BY tmpScreens.name"
      Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsScreens.EOF
        cboHRProScreen.AddItem rsScreens!Name
        cboHRProScreen.ItemData(cboHRProScreen.NewIndex) = rsScreens!ScreenID

        rsScreens.MoveNext
      Wend
      rsScreens.Close
      Set rsScreens = Nothing
    End If
  
    If cboHRProScreen.ListCount > 0 Then
      cboHRProScreen.ListIndex = 0
    End If
  End If
  
End Sub

Private Sub GetTablesViews()

  ' Populate the table(views) combo with the selected tables (views)
  ' (as passed in the ssi table views collection)
  
  Dim oSSITableView As clsSSITableView
  Dim iIndex As Integer
  Dim iLoop As Integer
  
  cboTableView.Clear
  
  For Each oSSITableView In mcolSSITableViews.Collection
    cboTableView.AddItem (oSSITableView.TableViewName)
  Next oSSITableView
  
  If cboTableView.ListCount > 0 Then
    iIndex = 0
    For iLoop = 0 To cboTableView.ListCount - 1
      If cboTableView.List(iLoop) = msTableViewName Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    cboTableView.ListIndex = iIndex
  End If
  
End Sub

Public Sub Initialize(piType As SSINTRANETLINKTYPES, _
                      psPrompt As String, _
                      psText As String, _
                      psHRProScreenID As String, _
                      psPageTitle As String, _
                      psURL As String, _
                      plngTableID As Long, _
                      psStartMode As String, _
                      plngViewID As Long, _
                      psUtilityType As String, psUtilityID As String, _
                      pfCopy As Boolean, psHiddenGroups As String, _
                      psTableViewName As String, _
                      pfNewWindow As Boolean, _
                      psEMailAddress As String, psEMailSubject As String, _
                      psAppFilePath As String, psAppParameters As String, _
                      psDocumentFilePath As String, pfDisplayDocumentHyperlink As Boolean, _
                      piElement_Type As Integer, piSeparatorOrientation As Integer, plngPictureID As Long, _
                      pfChartShowLegend As Boolean, piChartType As Integer, pfChartShowGrid As Boolean, _
                      pfChartStackSeries As Boolean, plngChartViewID As Long, miChartTableID As Long, _
                      plngChartColumnID As Long, plngChartFilterID As Long, piChartAggregateType As Integer, _
                      pfChartShowValues As Boolean, psCombinedHiddenGroups As String, _
                      ByRef pcolSSITableViews As clsSSITableViews)
  
  Set mcolSSITableViews = pcolSSITableViews
  
  mfLoading = True
  
  miLinkType = piType
  'mlngPersonnelTableID = plngPersonnelTableID
  mlngTableID = plngTableID
  mlngViewID = plngViewID
  msTableViewName = psTableViewName
  msCombinedHiddenGroups = psCombinedHiddenGroups
  
  FormatScreen
  
  GetTablesViews
    
  'NPG20080128 Fault 12873     ' If Len(psURL) > 0 Then
  If Len(psURL) > 0 And Len(Trim(psEMailAddress)) = 0 Then
    optLink(SSINTLINKSCREEN_URL).value = True
  End If
  
  'NPG20080125 Fault 12873
  If Len(psEMailAddress) > 0 Then
    optLink(SSINTLINKSCREEN_EMAIL).value = True
  End If

  If Len(psAppFilePath) > 0 Then
    optLink(SSINTLINKSCREEN_APPLICATION).value = True
  End If
  
  If (Len(psDocumentFilePath) > 0) Or (miLinkType = SSINTLINK_DOCUMENT) Then
    optLink(SSINTLINKSCREEN_DOCUMENT).value = True
  End If
  
  If Len(psUtilityID) > 0 And piChartType = 0 Then
    If CLng(psUtilityID) > 0 Then
      optLink(SSINTLINKSCREEN_UTILITY).value = True
    End If
  End If
  
  If miLinkType = SSINTLINK_HYPERTEXT And _
    Len(psURL) = 0 And _
    Len(psEMailAddress) = 0 And _
    Len(psAppFilePath) = 0 Then
    
    optLink(SSINTLINKSCREEN_UTILITY).value = True
    
  
  End If
  
  If miLinkType = SSINTLINK_DOCUMENT Then
    optLink(SSINTLINKSCREEN_DOCUMENT).value = True
  End If
  
  If miLinkType <> SSINTLINK_BUTTON Then
    ' disable the irrelevant options
    optLink(SSINTLINKPWFSTEPS).Enabled = False
    optLink(SSINTLINKCHART).Enabled = False
    optLink(SSINTLINKDB_VALUE).Enabled = False
  End If
    
    
  
  If piElement_Type = 1 Then
    optLink(SSINTLINKSEPARATOR).value = True
    chkNewColumn.value = IIf(piSeparatorOrientation > 0, 1, 0)
  ElseIf piElement_Type = 2 Then
    optLink(SSINTLINKCHART).value = True
  ElseIf piElement_Type = 3 Then
    optLink(SSINTLINKPWFSTEPS).value = True
  ElseIf piElement_Type = 4 Then
    optLink(SSINTLINKDB_VALUE).value = True
  End If
  

  
  GetHRProTables
  GetHRProUtilityTypes
  UtilityType = psUtilityType
  SetChartTypes
    
  Prompt = psPrompt
  Text = psText
  HRProScreenID = psHRProScreenID
  PageTitle = psPageTitle
  'NPG20080125 Fault 12873
  EMailAddress = psEMailAddress
  EMailSubject = psEMailSubject
  URL = psURL

  StartMode = psStartMode
  UtilityType = psUtilityType
  UtilityID = psUtilityID
  NewWindow = pfNewWindow
  
  AppFilePath = psAppFilePath
  AppParameters = psAppParameters
  
  DocumentFilePath = psDocumentFilePath
  DisplayDocumentHyperlink = pfDisplayDocumentHyperlink

  'NPG Dashboard
  ' Separator...
  glngPictureID = plngPictureID
  miSeparatorOrientation = piSeparatorOrientation
  imgIcon_Refresh
  
  'Chart Controls...
  ChartType = piChartType
  ChartViewID = plngChartViewID
  ChartFilterID = plngChartFilterID
  ChartTableID = miChartTableID
  ChartColumnID = plngChartColumnID
  ChartAggregateType = piChartAggregateType
  ChartShowLegend = pfChartShowLegend
  ChartShowGridlines = pfChartShowGrid
  ChartStackSeries = pfChartStackSeries
  ChartShowValues = pfChartShowValues
  
  ' Set up 'Database Value' combos...
  PopulateParentsCombo (miChartTableID) ' populate and set default value
  PopulateColumnsCombo (cboParents.ItemData(cboParents.ListIndex))
  optAggregateType(0).value = IIf(ChartAggregateType = 0, True, False)
  optAggregateType(1).value = IIf(ChartAggregateType = 1, True, False)
  txtFilter.Tag = miChartFilterID
  txtFilter.Text = GetExpressionName(txtFilter.Tag)
  
  txtFilter.Enabled = False
  txtFilter.BackColor = vbButtonFace
  
  PopulateAccessGrid psHiddenGroups

  mfChanged = False
  
  mfLoading = False
  
  If pfCopy Then mfChanged = True
  RefreshControls
  
End Sub

Private Sub RefreshControls()

  Dim sUtilityMessage As String
  
  If mblnRefreshing Then Exit Sub
  
  sUtilityMessage = ""
  
  fraHRProScreenLink.Visible = optLink(SSINTLINKSCREEN_HRPRO).value
  fraHRProUtilityLink.Visible = optLink(SSINTLINKSCREEN_UTILITY).value
  fraURLLink.Visible = optLink(SSINTLINKSCREEN_URL).value
  fraEmailLink.Visible = optLink(SSINTLINKSCREEN_EMAIL).value
  fraApplicationLink.Visible = optLink(SSINTLINKSCREEN_APPLICATION).value
  fraDocument.Visible = optLink(SSINTLINKSCREEN_DOCUMENT).value
  fraLinkSeparator.Visible = (optLink(SSINTLINKSEPARATOR).value Or optLink(SSINTLINKPWFSTEPS).value)
  fraChartLink.Visible = optLink(SSINTLINKCHART).value
  fraDBValue.Visible = optLink(SSINTLINKDB_VALUE).value
      
  ' Disable the HR Pro screen controls as required.
  cboHRProTable.Enabled = (optLink(SSINTLINKSCREEN_HRPRO).value) And (cboHRProTable.ListCount > 0)
  cboHRProTable.BackColor = IIf(cboHRProTable.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProTable.Enabled = cboHRProTable.Enabled
  cboHRProScreen.Enabled = (optLink(SSINTLINKSCREEN_HRPRO).value) And (cboHRProScreen.ListCount > 0)
  cboHRProScreen.BackColor = IIf(cboHRProScreen.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProScreen.Enabled = cboHRProTable.Enabled
  
  txtPageTitle.Enabled = cboHRProTable.Enabled
  txtPageTitle.BackColor = cboHRProTable.BackColor
  lblPageTitle.Enabled = cboHRProTable.Enabled
  If Not optLink(SSINTLINKSCREEN_HRPRO).value Then
    cboHRProTable.Clear
    cboHRProScreen.Clear
    cboStartMode.Clear
    txtPageTitle.Text = ""
  End If
  
  cboStartMode.Enabled = (cboStartMode.ListCount > 1) _
      And (optLink(SSINTLINKSCREEN_HRPRO).value)
  cboStartMode.BackColor = IIf(cboStartMode.Enabled, vbWindowBackground, vbButtonFace)
  lblStartMode.Enabled = cboHRProTable.Enabled
  
  ' Disable the UTILITY controls as required.
  If Not optLink(SSINTLINKSCREEN_UTILITY).value Then
    cboHRProUtilityType.Clear
    cboHRProUtility.Clear
  Else
    ' For Workflows only, check if the selected Workflow is enabled.
    ' Display a message as required.
    'JPD 20060714 Fault 11226
    If cboHRProUtilityType.ListIndex >= 0 Then
      If cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex) = utlWorkflow Then
        If cboHRProUtility.ListCount > 0 Then
          recWorkflowEdit.Index = "idxWorkflowID"
          recWorkflowEdit.Seek "=", cboHRProUtility.ItemData(cboHRProUtility.ListIndex)
    
          If Not recWorkflowEdit.NoMatch Then
            If (Not recWorkflowEdit.Fields("enabled").value) Then
              sUtilityMessage = "This Workflow is not currently enabled."
            End If
          End If
        End If
      End If
    End If
  End If
  
  cboHRProUtilityType.Enabled = optLink(SSINTLINKSCREEN_UTILITY).value
  cboHRProUtilityType.BackColor = IIf(cboHRProUtilityType.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProUtilityType.Enabled = cboHRProUtilityType.Enabled
  
  cboHRProUtility.Enabled = optLink(SSINTLINKSCREEN_UTILITY).value And (cboHRProUtility.ListCount > 0)
  cboHRProUtility.BackColor = IIf(cboHRProUtility.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProUtility.Enabled = cboHRProUtility.Enabled
  
  ' Disable the URL controls as required.
  txtURL.Enabled = optLink(SSINTLINKSCREEN_URL).value
  txtURL.BackColor = IIf(txtURL.Enabled, vbWindowBackground, vbButtonFace)
  lblURL.Enabled = txtURL.Enabled
  chkNewWindow.Enabled = txtURL.Enabled
  ' 'NPG20080128 Fault 12873 - If Not txtURL.Enabled Then
  If Not txtURL.Enabled And Not optLink(SSINTLINKSCREEN_EMAIL).value Then
    txtURL.Text = ""
  End If
  
  'NPG20080125 Fault 12873
  ' Disable the EMail controls as required.
  txtEmailAddress.Enabled = optLink(SSINTLINKSCREEN_EMAIL).value
  txtEmailAddress.BackColor = IIf(txtEmailAddress.Enabled, vbWindowBackground, vbButtonFace)
  lblEMailAddress.Enabled = txtEmailAddress.Enabled
  txtEmailSubject.Enabled = optLink(SSINTLINKSCREEN_EMAIL).value
  txtEmailSubject.BackColor = IIf(txtEmailAddress.Enabled, vbWindowBackground, vbButtonFace)
  lblEmailSubject.Enabled = txtEmailAddress.Enabled
  If Not txtEmailAddress.Enabled Then
    txtEmailAddress.Text = ""
    txtEmailSubject.Text = ""
  End If

  ' Disable the Application link controls as required.
  txtAppFilePath.Enabled = optLink(SSINTLINKSCREEN_APPLICATION).value
  txtAppFilePath.BackColor = IIf(txtAppFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblAppFilePath.Enabled = txtAppFilePath.Enabled
  txtAppParameters.Enabled = optLink(SSINTLINKSCREEN_APPLICATION).value
  txtAppParameters.BackColor = IIf(txtAppFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblAppParameters.Enabled = txtAppFilePath.Enabled
  If Not txtAppFilePath.Enabled Then
    txtAppFilePath.Text = ""
    txtAppParameters.Text = ""
  End If

  ' Disable the Report Link controls as required.
  txtDocumentFilePath.Enabled = optLink(SSINTLINKSCREEN_DOCUMENT).value
  txtDocumentFilePath.BackColor = IIf(txtDocumentFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblDocumentFilePath.Enabled = txtDocumentFilePath.Enabled
  chkDisplayDocumentHyperlink.Enabled = txtDocumentFilePath.Enabled
  If Not txtDocumentFilePath.Enabled Then
    txtDocumentFilePath.Text = ""
    chkDisplayDocumentHyperlink.value = vbUnchecked
  End If
  
  ' NPG Dashboard
  
  txtIcon.Enabled = False
  txtIcon.BackColor = vbButtonFace

  
  If optLink(SSINTLINKSEPARATOR).value Or optLink(SSINTLINKCHART) _
      Or optLink(SSINTLINKPWFSTEPS).value Or optLink(SSINTLINKDB_VALUE).value Then
    txtPrompt.Enabled = False
    txtPrompt.BackColor = vbButtonFace
    
    If optLink(SSINTLINKPWFSTEPS).value Then
      txtText.Enabled = False
      txtText.BackColor = vbButtonFace
    Else
      txtText.Enabled = True
      txtText.BackColor = vbWindowBackground
    End If
        
    If optLink(SSINTLINKSEPARATOR).value Then txtPrompt.Text = "<SEPARATOR>"
    If optLink(SSINTLINKCHART).value Then txtPrompt.Text = "<CHART>"
    If optLink(SSINTLINKPWFSTEPS).value Then txtPrompt.Text = "<PENDING WORKFLOWS>"
    If optLink(SSINTLINKDB_VALUE).value Then txtPrompt.Text = "<DATABASE VALUE>"
    
    If optLink(SSINTLINKCHART).value Then
      MSChart1.RowCount = 1
    End If
  Else
    ' NPG20100427 Fault HRPRO-888
    If Not txtPrompt.Enabled Then txtPrompt.Text = ""
    
    txtPrompt.Enabled = True
    txtPrompt.BackColor = vbWindowBackground
    ' NPG20100427 Fault HRPRO-910
    txtText.Enabled = True
    txtText.BackColor = vbWindowBackground
  End If
  
  If (optLink(SSINTLINKSEPARATOR).value And miLinkType = SSINTLINK_HYPERTEXT) Or optLink(SSINTLINKPWFSTEPS).value Or miLinkType = SSINTLINK_DROPDOWNLIST Then
    ' Disable the icon and new column options for hypertext link separators...
    chkNewColumn.Visible = False
    lblIcon.Visible = False
    txtIcon.Visible = False
    cmdIcon.Visible = False
    cmdIconClear.Visible = False
    imgIcon.Visible = False
    lblNoOptions.Visible = True
    lblNoOptions.Top = 345
    
  ElseIf (optLink(SSINTLINKSEPARATOR).value And miLinkType <> SSINTLINK_HYPERTEXT) Then
    ' Enable the icon and new column options for dashboard link separators...
    chkNewColumn.Visible = True
    lblIcon.Visible = True
    txtIcon.Visible = True
    cmdIcon.Visible = True
    cmdIconClear.Visible = True
    imgIcon.Visible = True
    lblNoOptions.Visible = False
    lblNoOptions.Top = 345
  Else
    lblNoOptions.Visible = False
    lblNoOptions.Top = 345
  End If
  
  If optLink(SSINTLINKPWFSTEPS).value Then
    fraLinkSeparator.Caption = "Pending Workflow Steps :"
  Else
    fraLinkSeparator.Caption = "Separator :"
  End If
  
  If optLink(SSINTLINKCHART).value And cboChartType.ListIndex >= 0 Then
  
    chkStackSeries.Visible = False
    chkDottedGridlines.Enabled = True
    
    Select Case cboChartType.ItemData(cboChartType.ListIndex)
      Case 0  '3D Bar
        
      Case 1  '2D Bar
            
      Case 4  '3D Area
        
      Case 6  '3D Step
    
      Case 7  '2D Step
        
      Case 14 '2D Pie
        ' disable dotted gridlines option
        chkDottedGridlines.value = 0
        chkDottedGridlines.Enabled = False
    End Select
  
  End If
  
  mblnRefreshing = True
  GetStartModes
  mblnRefreshing = False
  
  lblHRProUtilityMessage.Caption = sUtilityMessage
  
  ' Disable the OK button as required.
  cmdOK.Enabled = mfChanged
  

End Sub

Private Sub GetStartModes()

  Dim iIndex As Integer
  Dim iOriginalStartMode As Integer
  Dim fPersonnelLink As Boolean
  
  iOriginalStartMode = 0
  iIndex = 0
  
  If (cboHRProTable.ListIndex >= 0) _
    And (cboTableView.ListIndex >= 0) Then
    fPersonnelLink = (cboHRProTable.ItemData(cboHRProTable.ListIndex) = mlngTableID)
  End If
  
  With cboStartMode
    If .ListIndex >= 0 Then
      iOriginalStartMode = .ItemData(.ListIndex)
    End If
    
    .Clear
    
    If optLink(SSINTLINKSCREEN_HRPRO).value Then
      If Not fPersonnelLink Then
        .AddItem "Find Window"
        .ItemData(.NewIndex) = 3
        If iOriginalStartMode = 3 Then
          iIndex = .NewIndex
        End If
      End If
      
      .AddItem "First Record"
      .ItemData(.NewIndex) = 2
      If iOriginalStartMode = 2 Then
        iIndex = .NewIndex
      End If
      
      If Not fPersonnelLink Then
        .AddItem "New Record"
        .ItemData(.NewIndex) = 1
        If iOriginalStartMode = 1 Then
          iIndex = .NewIndex
        End If
      End If
      
      .ListIndex = iIndex
    End If
    
    .Enabled = (.ListCount > 1) _
      And (optLink(SSINTLINKSCREEN_HRPRO).value)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With
  
End Sub


Private Sub RefreshChart()
  
  Dim numSeries As Integer, iCount As Integer
  Dim iLoop As Integer
  
  ' Set the chart type from the combo
  MSChart1.ChartType = cboChartType.ItemData(cboChartType.ListIndex)

  ' Display Legend
  MSChart1.ShowLegend = chkShowLegend
  
  ' Display Dotted Gridlines
  If chkDottedGridlines Then
    With MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid
       .MajorPen.VtColor.Set 195, 195, 195
       .MajorPen.Width = 1
       .MajorPen.Style = VtPenStyleDotted
       .MinorPen.Style = VtPenStyleNull
    End With
    With MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid
       .MajorPen.VtColor.Set 195, 195, 195
       .MajorPen.Width = 1
       .MajorPen.Style = VtPenStyleDotted
       .MinorPen.Style = VtPenStyleNull
    End With
    With MSChart1.Plot.Axis(VtChAxisIdZ).AxisGrid
       .MajorPen.VtColor.Set 195, 195, 195
       .MajorPen.Width = 1
       .MajorPen.Style = VtPenStyleDotted
       .MinorPen.Style = VtPenStyleNull
    End With
  Else
    With MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid
       .MajorPen.VtColor.Set 195, 195, 195
       .MajorPen.Width = 0
       .MajorPen.Style = VtPenStyleNull
    End With
    With MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid
       .MajorPen.VtColor.Set 195, 195, 195
       .MajorPen.Width = 0
       .MajorPen.Style = VtPenStyleNull
    End With
    With MSChart1.Plot.Axis(VtChAxisIdZ).AxisGrid
       .MajorPen.VtColor.Set 195, 195, 195
       .MajorPen.Width = 0
       .MajorPen.Style = VtPenStyleNull
    End With
  End If
  
  ' Stack Series
  MSChart1.Stacking = chkStackSeries
  
  ' set the colours to the new set
  If MSChart1.ColumnCount > 0 Then MSChart1.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 166, 206, 227
  If MSChart1.ColumnCount > 1 Then MSChart1.Plot.SeriesCollection(2).DataPoints(-1).Brush.FillColor.Set 178, 223, 138
  If MSChart1.ColumnCount > 2 Then MSChart1.Plot.SeriesCollection(3).DataPoints(-1).Brush.FillColor.Set 251, 154, 153
  If MSChart1.ColumnCount > 3 Then MSChart1.Plot.SeriesCollection(4).DataPoints(-1).Brush.FillColor.Set 253, 191, 111
  
  ' Show Values
  With MSChart1
  numSeries = .Plot.SeriesCollection.Count
    For iCount = 1 To numSeries
      .Plot.SeriesCollection(iCount).DataPoints(-1).EdgePen.VtColor.Set 0, 0, 0
      
      If chkShowValues Then
        .Plot.SeriesCollection.Item(iCount).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
      Else
        .Plot.SeriesCollection.Item(iCount).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeNone
      End If
    Next iCount
  End With
  
  ' set the angle of the 3d chart to improve clarity...
  If MSChart1.Chart3d = True Then MSChart1.Plot.View3d.Set 165, 5
  
  'Set the font text of the axis for clarity
  For iLoop = 1 To MSChart1.Plot.Axis(1).Labels.Count
    MSChart1.Plot.Axis(1).Labels(iLoop).VtFont.Size = IIf(MSChart1.Chart3d, 10, 8)
  Next
    
    
  MSChart1.Plot.Axis(0).AxisScale.Hide = True
  MSChart1.Plot.Axis(2).AxisScale.Hide = True
  MSChart1.Plot.Axis(3).AxisScale.Hide = True
    
  MSChart1.Legend.Location.LocationType = VtChLocationTypeRight
                     
  ' Set X Y coordinates for bottom left corner
  MSChart1.Legend.Location.Rect.Min.X = MSChart1.Width - 60
  MSChart1.Legend.Location.Rect.Min.Y = 0
  ' Set X Y coordinates for top right corner
  MSChart1.Legend.Location.Rect.Max.X = MSChart1.Width
  MSChart1.Legend.Location.Rect.Max.Y = MSChart1.Height
                 
  MSChart1.Plot.LocationRect.Min.X = 0
  MSChart1.Plot.LocationRect.Min.Y = 0
  '1200 twips
  MSChart1.Plot.LocationRect.Max.X = MSChart1.Width - 1000
  MSChart1.Plot.LocationRect.Max.Y = MSChart1.Height
  
End Sub


Private Function ValidateLink() As Boolean

  ' Return FALSE if the link definition is invalid.
  Dim fValid As Boolean
  Dim iLoop As Integer
  Dim pSelectedGroup As String
  Dim psDuplicateGroups As String
  
  fValid = True
  
  ' Check that a prompt has been entered (if required)
  'JPD 20070424 Fault 12168
  'If (miLinkType = SSINTLINK_BUTTON) And _
  '  (Len(txtPrompt.Text) = 0) Then
  '  fValid = False
  '  MsgBox "No prompt has been entered.", vbOKOnly + vbExclamation, Application.Name
  '  txtPrompt.SetFocus
  'End If
  
  ' Check that text has been entered
  If fValid Then
    If (Len(txtText.Text) = 0) And Not optLink(SSINTLINKPWFSTEPS).value And Not optLink(SSINTLINKSEPARATOR).value Then
      fValid = False
      MsgBox "No text has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtText.SetFocus
    End If
  End If
  
  ' Check that the HR Pro screen has been selected (if required)
  If fValid Then
    If (miLinkType <> SSINTLINK_HYPERTEXT) And _
      (optLink(SSINTLINKSCREEN_HRPRO).value) Then
      If cboHRProScreen.ListIndex < 0 Then
        fValid = False
        MsgBox "No HR Pro screen has been selected.", vbOKOnly + vbExclamation, Application.Name
        cboHRProTable.SetFocus
      End If
    End If
  End If
  
  ' Check that the HR Pro page title been entered
  If fValid Then
    If (miLinkType <> SSINTLINK_HYPERTEXT) And _
      (optLink(SSINTLINKSCREEN_HRPRO).value) Then
      If (Len(txtPageTitle.Text) = 0) Then
        fValid = False
        MsgBox "No page title has been entered.", vbOKOnly + vbExclamation, Application.Name
        txtPageTitle.SetFocus
      End If
    End If
  End If
  
  ' Check that a URL has been entered (if required)
  If fValid Then
    If optLink(SSINTLINKSCREEN_URL).value And _
      (Len(txtURL.Text) = 0) Then
      fValid = False
      MsgBox "No URL has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtURL.SetFocus
    End If
  End If
    
  'NPG20080125 Fault 12873
  ' Check that an Email Address has been entered (if required)
  If fValid Then
    If optLink(SSINTLINKSCREEN_EMAIL).value And _
      ((Len(txtEmailAddress.Text) = 0) Or InStr(1, txtEmailAddress.Text, "@", 1) = 0) Then
      fValid = False
      MsgBox "No Email Address has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtEmailAddress.SetFocus
    End If
  End If

  ' Check that a utility has been entered (if required)
  If fValid Then
    If (optLink(SSINTLINKSCREEN_UTILITY).value) Then
      If cboHRProUtility.ListIndex < 0 Then
        fValid = False
        MsgBox "No " & cboHRProUtilityType.List(cboHRProUtilityType.ListIndex) & " has been selected.", vbOKOnly + vbExclamation, Application.Name
        cboHRProUtilityType.SetFocus
      End If
    End If
  End If

  ' Check that an Application File Path has been entered (if required)
  If fValid Then
    If optLink(SSINTLINKSCREEN_APPLICATION).value And _
      (Len(txtAppFilePath.Text) = 0) Then
      fValid = False
      MsgBox "No application file path has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtAppFilePath.SetFocus
    End If
  End If

'TM20090224 - Fault 13557, remove the restriction to allow links to shortcuts and other executable types.
'  If fValid Then
'    If optLink(SSINTLINKSCREEN_APPLICATION).Value And _
'      (LCase(Right(txtAppFilePath.Text, 4)) <> ".exe") Then
'      fValid = False
'      MsgBox "Please enter a valid executable file path.", vbOKOnly + vbExclamation, Application.Name
'      txtAppFilePath.SetFocus
'    End If
'  End If

  If fValid Then
    If optLink(SSINTLINKSCREEN_DOCUMENT).value And _
      (Len(txtDocumentFilePath.Text) = 0) Then
      fValid = False
      MsgBox "No Document URL has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtDocumentFilePath.SetFocus
    End If
  End If

  If fValid Then
    If optLink(SSINTLINKSCREEN_DOCUMENT).value And _
      ((LCase(Left(txtDocumentFilePath.Text, 7)) <> "http://") And (LCase(Left(txtDocumentFilePath.Text, 8)) <> "https://")) Then
      fValid = False
      MsgBox "Please enter a valid URL path.", vbOKOnly + vbExclamation, Application.Name
      txtDocumentFilePath.SetFocus
    End If
  End If

  If fValid Then
    If optLink(SSINTLINKCHART).value And _
      ChartColumnID = 0 Then
      fValid = False
      MsgBox "Please define chart data using the 'Data' button.", vbOKOnly + vbExclamation, Application.Name
      cmdChartData.SetFocus
    End If
  End If

  ' Only one Pending Workflow Steps per security group...
  If fValid Then
    If optLink(SSINTLINKPWFSTEPS).value And Len(msCombinedHiddenGroups) > 0 Then
      ' loop through the chosen security groups and check they're in the combined string
      
      psDuplicateGroups = ""
      pSelectedGroup = ""
      
      With grdAccess
        For iLoop = 0 To (.Rows - 1)
          .Bookmark = .AddItemBookmark(iLoop)
          If .Columns("Access").value Then
            pSelectedGroup = vbTab & .Columns("GroupName").Text & vbTab
            If InStr(msCombinedHiddenGroups, pSelectedGroup) = 0 Then
              fValid = False
              psDuplicateGroups = psDuplicateGroups & vbCrLf & Replace(pSelectedGroup, vbTab, "")
            End If
          End If
        Next iLoop
      .MoveFirst
      End With
      
      If Not fValid Then
        MsgBox "'Pending Workflows' can only be defined once per user group." & vbCrLf & _
                "It has already been defined for the following groups:" & vbCrLf & _
                psDuplicateGroups, vbOKOnly + vbExclamation, Application.Name
        grdAccess.SetFocus
      End If
    End If
  End If

  ValidateLink = fValid
  
End Function

Private Function imgIcon_Refresh() As Boolean
  Dim strFileName As String
  
  If PictureID > 0 Then
    With recPictEdit
      .Index = "idxID"
      .Seek "=", PictureID
      If Not .NoMatch Then
        strFileName = ReadPicture
        Set imgIcon.Picture = LoadPicture(strFileName)
        Kill strFileName
        
        txtIcon.Text = .Fields("Name")
      End If
    End With
  Else
    Set imgIcon.Picture = LoadPicture(vbNullString)
    txtIcon.Text = vbNullString
  End If
End Function


Private Sub PopulateParentsCombo(plngDefaultID As Long)
  
  Dim i As Integer
  ' Clear the contents of the combo.
  cboParents.Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If !TableType <> iTabLookup And Not !Deleted Then
        cboParents.AddItem !TableName
        cboParents.ItemData(cboParents.NewIndex) = !TableID
      End If

      .MoveNext
    Loop
  End With

  ' Set the correct item as default
  If plngDefaultID = 0 Then
    cboParents.ListIndex = 0
  Else
    For i = 0 To cboParents.ListCount - 1
      If cboParents.ItemData(i) = plngDefaultID Then
        cboParents.ListIndex = i
        Exit For
      End If
    Next
  End If

End Sub

Private Function PopulateColumnsCombo(plngTableID As Long) As Boolean

  Dim i As Integer
  
  ' Clear the contents of the combo
  cboColumns.Clear

  ' Add the table's columns to the view definition in the local database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  recColEdit.Index = "idxTableID"
  recColEdit.Seek "=", plngTableID

  fOK = Not recColEdit.NoMatch

  If fOK Then

    Do While Not recColEdit.EOF

      ' If no more columns for this table exit loop
      If recColEdit!TableID <> plngTableID Then
        Exit Do
      End If

      ' Don't add deleted or system columns
      If recColEdit!Deleted <> True And recColEdit!columntype <> giCOLUMNTYPE_SYSTEM Then
        ' Add the column to the combo
        ' Making sure it isn't ole, photo, wp or link...
        If recColEdit!DataType <> dtlongvarchar And _
          recColEdit!DataType <> dtBINARY And _
          recColEdit!DataType <> dtVARBINARY And _
          recColEdit!DataType <> dtLONGVARBINARY Then
            cboColumns.AddItem recColEdit.Fields("ColumnName")
            cboColumns.ItemData(cboColumns.NewIndex) = recColEdit.Fields("ColumnID")
        End If
      End If

      recColEdit.MoveNext
    Loop


    ' Set the correct item as default

    For i = 0 To cboColumns.ListCount - 1
      If cboColumns.ItemData(i) = ChartColumnID Then
        cboColumns.ListIndex = i
        Exit For
      End If
    Next

    If cboColumns.ListIndex < 0 Then cboColumns.ListIndex = 0
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Private Sub cboChartType_Click()
  mfChanged = True
  ' Display new chart details
  RefreshChart
  RefreshControls
End Sub

Private Sub cboColumns_Click()
  Dim piColumnDataType As Integer
  Dim lngColumnID As Long

  mfChanged = True

  miChartColumnID = cboColumns.ItemData(cboColumns.ListIndex)
  
  lngColumnID = cboColumns.ItemData(cboColumns.ListIndex)
  
  piColumnDataType = GetColumnDataType(lngColumnID)
  
  ' Disable 'total' option if not numeric or integer
  If piColumnDataType <> dtinteger And piColumnDataType <> dtNUMERIC Then
    optAggregateType(0).value = True
    optAggregateType(1).Enabled = False
    optAggregateType(1).ForeColor = vbButtonFace
  Else
    optAggregateType(1).Enabled = True
    optAggregateType(1).ForeColor = vbWindowBackground
  End If
  
  RefreshControls
End Sub

Private Sub cboParents_Click()

  mfChanged = True

  miChartTableID = cboParents.ItemData(cboParents.ListIndex)
  PopulateColumnsCombo (miChartTableID)
  
  ' Check if the selected expression is for the current table.
  With recExprEdit
    .Index = "idxExprID"
    .Seek "=", txtFilter.Tag, False
    
    If Not .NoMatch Then
      If (!TableID <> miChartTableID) Then
        txtFilter.Tag = 0
        txtFilter.Text = ""
      End If
    Else
      txtFilter.Tag = 0
      txtFilter.Text = ""
    End If
  End With
  
  
  RefreshControls
End Sub

Private Sub cmdFilterClear_Click()
  txtFilter.Text = vbNullString
  txtFilter.Tag = 0
  miChartFilterID = 0
  mfChanged = True
  
  RefreshControls
End Sub

Private Sub optAggregateType_Click(Index As Integer)

  mfChanged = True

  miChartAggregateType = IIf(optAggregateType(0).value, 0, 1)
  
  RefreshControls
End Sub

Private Sub cmdFilter_Click()

  ' Display the 'Where Clause' expression selection form.
  'On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise miChartTableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      
      miChartFilterID = .ExpressionID
      mfChanged = True
      
      RefreshControls
      
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
      End If
    End If

  End With


TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub
Private Sub chkDottedGridlines_Click()
  mfChanged = True
  ' refresh the chart
  RefreshChart
  
  RefreshControls
End Sub

Private Sub chkNewColumn_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkShowLegend_Click()
  mfChanged = True
  ' refresh the chart
  RefreshChart
  
  RefreshControls
End Sub

Private Sub chkStackSeries_Click()
  mfChanged = True
  ' refresh the chart
  RefreshChart
  RefreshControls
End Sub

Private Sub chkShowValues_Click()
  mfChanged = True
  ' refresh the chart
  RefreshChart
  RefreshControls
End Sub


Private Sub cmdChartData_Click()
  
  
  Dim frmSSIChart As New frmSSIntranetChart

  With frmSSIChart
    .Initialize 1, ChartTableID, ChartColumnID, ChartFilterID, ChartAggregateType
    
    .Show vbModal
    
    If Not .Cancelled Then
      ' miChartViewID = .cboTableView.ItemData(.cboTableView.ListIndex)
      ChartTableID = .cboParents.ItemData(.cboParents.ListIndex)
      ChartColumnID = .cboColumns.ItemData(.cboColumns.ListIndex)
      ChartAggregateType = IIf(.optAggregateType(0).value, 0, 1)
      ChartFilterID = .txtFilter.Tag
      
      mfChanged = True
      
      RefreshControls
      
    End If
  
    UnLoad frmSSIChart
    Set frmSSIChart = Nothing
    
  End With
End Sub

Private Sub cmdIconClear_Click()
  glngPictureID = frmPictSel.SelectedPicture
  imgIcon_Refresh
  mfChanged = True
  RefreshControls
End Sub


Private Sub cboHRProScreen_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProTable_Click()
  GetHRProScreens
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProUtility_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProUtilityType_Click()
  GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboStartMode_Click()

  'JPD 20050810 Fault 10241
  If Not mblnRefreshing Then
    mfChanged = True
    RefreshControls
  End If
  
End Sub

Private Sub cboTableView_Click()
  
  'TM20090512 - Fault 13680
  'Only refresh the table combo if the base table changes.
  
  Dim lngNewTableID As Long
  
  msTableViewName = cboTableView.List(cboTableView.ListIndex)
  lngNewTableID = GetTableIDFromCollection(mcolSSITableViews, msTableViewName)
  
  If mlngTableID <> lngNewTableID Then
    mlngTableID = lngNewTableID
    mlngViewID = GetViewIDFromCollection(mcolSSITableViews, msTableViewName)
  
    GetHRProTables
'TM20090512 - Fault 13680
'    GetHRProUtilityTypes
  Else
    mlngTableID = lngNewTableID
    mlngViewID = GetViewIDFromCollection(mcolSSITableViews, msTableViewName)
    
  End If
  
  mfChanged = True
  RefreshControls

End Sub

Private Sub chkDisplayDocumentHyperlink_Click()

  Dim fValid As Boolean
  
  fValid = True

  If optLink(SSINTLINKSCREEN_DOCUMENT).value And (Len(txtDocumentFilePath.Text) = 0) And chkDisplayDocumentHyperlink.value Then
    fValid = False
    MsgBox "No Document URL has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtDocumentFilePath.SetFocus
  End If
  
  If fValid Then
    mfChanged = True
    RefreshControls
  End If

End Sub

Private Sub chkNewWindow_Click()

  Dim fValid As Boolean
  
  fValid = True

  If optLink(SSINTLINKSCREEN_URL).value And (Len(txtURL.Text) = 0) And chkNewWindow.value Then
    fValid = False
    MsgBox "No URL has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtURL.SetFocus
  End If
  
  If fValid Then
    mfChanged = True
    RefreshControls
  End If
  
End Sub

Private Sub cmdAppFilePathSel_Click()

  On Local Error GoTo LocalErr
  
  With CommonDialog1

    .FileName = txtAppFilePath.Text
    If txtAppFilePath.Text = vbNullString Then
      .InitDir = "c:\"
    End If

    .CancelError = True
    .DialogTitle = "Application file path"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
    
    .ShowOpen
    
    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        MsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
      Else
        txtAppFilePath.Text = .FileName
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    MsgBox Err.Description, vbCritical
  End If

End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  UnLoad Me
End Sub

Private Sub cmdIcon_Click()
  ' Display the icon selection form.

  frmPictSel.SelectedPicture = glngPictureID
  frmPictSel.PictureType = vbPicTypeBitmap ' vbPicTypeIcon

  frmPictSel.Show vbModal
  
  If frmPictSel.SelectedPicture > 0 Then
    glngPictureID = frmPictSel.SelectedPicture
    imgIcon_Refresh
  End If
  
  If frmPictSel.Cancelled = False Then
    mfChanged = True
    RefreshControls
  End If
  
  Set frmPictSel = Nothing


End Sub



Private Sub cmdOK_Click()

  If ValidateLink Then
    Cancelled = False
    Me.Hide
  End If

End Sub

Private Sub cmdReportOutputFilePathSel_Click()

  On Local Error GoTo LocalErr
  
  With CommonDialog1

    .FileName = txtDocumentFilePath.Text
    If txtDocumentFilePath.Text = vbNullString Then
      .InitDir = "c:\"
    End If

    .CancelError = True
    .DialogTitle = "Report Output file path"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
    
    .ShowOpen
    
    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        MsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
      Else
        txtDocumentFilePath.Text = .FileName
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    MsgBox Err.Description, vbCritical
  End If

End Sub

Private Sub Form_Initialize()
  mblnReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)
End Sub

Private Sub Form_Load()

  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  fraOKCancel.BorderStyle = vbBSNone

  grdAccess.RowHeight = 239

End Sub

Private Sub PopulateAccessGrid(psHiddenGroups As String)

  ' Populate the access grid.
  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset
  Dim sVisibility As String
  Dim fAllVisible As Boolean
  
  fAllVisible = True

  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With

  ' Get the recordset of user groups and their access on this definition.
  sSQL = "SELECT name FROM sysusers" & _
    " WHERE gid = uid AND gid > 0" & _
    "   AND not (name like 'ASRSys%') AND not (name like 'db[_]%')" & _
    " ORDER BY name"
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsGroups
    Do While Not .EOF
      ' Add the user groups and their access on this definition to the access grid.
      If InStr(vbTab & UCase(psHiddenGroups) & vbTab, vbTab & UCase(Trim(!Name)) & vbTab) > 0 Then
        sVisibility = "False"
        fAllVisible = False
      Else
        sVisibility = "True"
      End If
            
      grdAccess.AddItem Trim(!Name) & vbTab & sVisibility

      .MoveNext
    Loop

    .Close
  End With
  Set rsGroups = Nothing

  With grdAccess
    .MoveFirst
    .Columns("Access").value = fAllVisible
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Cancelled = True
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOK_Click
        Cancel = True   'MH20021105 Fault 4694
    End Select
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' Save the form position to registry.
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
  End If

End Sub

Private Sub grdAccess_Change()

  Dim iLoop As Integer
  Dim varFirstRow As Variant
  Dim varCurrentRow As Variant
  Dim fNewValue As Boolean
  Dim fAllVisible As Boolean
  
  UI.LockWindow grdAccess.hWnd
  
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) Then
    ' The 'All Groups' access has changed. Apply the selection to all other groups.
    With grdAccess
      .MoveFirst

      For iLoop = 0 To (.Rows - 1)
        If iLoop = 0 Then
          fNewValue = .Columns("Access").value
        Else
          .Columns("Access").value = fNewValue
        End If

        .MoveNext
      Next iLoop

      .MoveFirst
    End With
  Else
    fAllVisible = True
    
    With grdAccess
      varFirstRow = .FirstRow
      varCurrentRow = .Bookmark
      .MoveLast
      
      For iLoop = (.Rows - 1) To 0 Step -1
        If iLoop = 0 Then
          .Columns("Access").value = fAllVisible
        Else
          If Not .Columns("Access").value Then fAllVisible = False
        End If
        
        .MovePrevious
      Next iLoop

      .MoveFirst
    
      .FirstRow = varFirstRow
      .Bookmark = varCurrentRow
    End With
  End If
    
  UI.UnlockWindow

  grdAccess.col = 1

  mfChanged = True
  RefreshControls

End Sub

Private Sub optLink_Click(Index As Integer)

  GetHRProTables
  GetHRProUtilityTypes
  UtilityType = CStr(utlCalendarReport)
  
  'dashboard
  If optLink(SSINTLINKSEPARATOR).value Then
    ElementType = 1
  ElseIf optLink(SSINTLINKCHART).value Then
    ElementType = 2
  ElseIf optLink(SSINTLINKPWFSTEPS).value Then
    ElementType = 3
  ElseIf optLink(SSINTLINKDB_VALUE).value Then
    ElementType = 4
  Else
    ElementType = 0
  End If
  
  mfChanged = True
  RefreshControls

End Sub

Private Sub txtDocumentFilePath_Change()
  mfChanged = True
  RefreshControls
End Sub
Private Sub txtDocumentFilePath_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtAppFilePath_Change()
  mfChanged = True
  RefreshControls
End Sub
Private Sub txtAppFilePath_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtAppParameters_Change()
  mfChanged = True
  RefreshControls
End Sub
Private Sub txtAppParameters_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtEmailAddress_Change()
  If Len(txtEmailAddress.Text) > 0 Then
    txtURL.Text = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  End If
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtEmailAddress_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtEmailSubject_Change()
  If Len(txtEmailAddress.Text) > 0 Then
    txtURL.Text = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  End If
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtEmailSubject_GotFocus()
  UI.txtSelText
End Sub

'Private Sub txtLinkSeparator_Change()
'  mfChanged = True
'  If miLinkType = SSINTLINK_BUTTON Then
'    Prompt = txtLinkSeparator.Text
'  ElseIf miLinkType = SSINTLINK_HYPERTEXT Then
'    Text = txtLinkSeparator.Text
'  End If
'End Sub

Private Sub txtPageTitle_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtPageTitle_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtPrompt_Change()
'  If optLink(SSINTLINKSEPARATOR).value Then txtLinkSeparator.Text = Prompt  'keep them in sync
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtPrompt_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtText_Change()
  'txtLinkSeparator.Text = Text
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtText_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtURL_Change()
  mfChanged = True
  RefreshControls
End Sub

Public Property Get Text() As String
  Text = txtText.Text
End Property

Public Property Let Text(ByVal psNewValue As String)
  txtText.Text = psNewValue
End Property

Public Property Get NewWindow() As Boolean
  NewWindow = chkNewWindow.value
End Property

Public Property Let NewWindow(ByVal pfNewValue As Boolean)
  chkNewWindow.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get HiddenGroups() As String

  Dim iLoop As Integer
  Dim sHiddenGroups As String
  
  sHiddenGroups = ""
  
  With grdAccess
    For iLoop = 0 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      If Not .Columns("Access").value Then
        sHiddenGroups = sHiddenGroups & .Columns("GroupName").Text & vbTab
      End If
    Next iLoop
    
    .MoveFirst
  End With

  If Len(sHiddenGroups) > 0 Then
    sHiddenGroups = vbTab & sHiddenGroups
  End If
  
  HiddenGroups = sHiddenGroups
  
End Property

Public Property Get URL() As String
  URL = IIf(optLink(SSINTLINKSCREEN_URL).value Or optLink(SSINTLINKSCREEN_EMAIL).value, txtURL.Text, "")
End Property

Public Property Get EMailAddress() As String
  If optLink(SSINTLINKSCREEN_EMAIL).value Then
    EMailAddress = txtEmailAddress.Text
    URL = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  Else
    EMailAddress = ""
    URL = ""
  End If
End Property

Public Property Get EMailSubject() As String
  If optLink(SSINTLINKSCREEN_EMAIL).value Then
    EMailSubject = txtEmailSubject.Text
    URL = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  Else
    EMailSubject = ""
    URL = ""
  End If
End Property

Public Property Get AppFilePath() As String
  If optLink(SSINTLINKSCREEN_APPLICATION).value Then
    AppFilePath = txtAppFilePath.Text
  Else
    AppFilePath = ""
  End If
End Property

Public Property Get AppParameters() As String
  If optLink(SSINTLINKSCREEN_APPLICATION).value Then
    AppParameters = txtAppParameters.Text
  Else
    AppParameters = ""
  End If
End Property

Public Property Get DocumentFilePath() As String
  If optLink(SSINTLINKSCREEN_DOCUMENT).value Then
    DocumentFilePath = txtDocumentFilePath.Text
  Else
    DocumentFilePath = ""
  End If
End Property

Public Property Get DisplayDocumentHyperlink() As Boolean
  If optLink(SSINTLINKSCREEN_DOCUMENT).value Then
    DisplayDocumentHyperlink = chkDisplayDocumentHyperlink.value
  Else
    DisplayDocumentHyperlink = False
  End If
End Property

Public Property Get UtilityType() As String

  If (cboHRProUtility.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_UTILITY).value) Then
    UtilityType = ""
  Else
    UtilityType = CStr(cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex))
  End If

End Property

Public Property Get UtilityID() As String

  If (cboHRProUtility.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_UTILITY).value) Then
    
    UtilityID = ""
  Else
    UtilityID = CStr(cboHRProUtility.ItemData(cboHRProUtility.ListIndex))
  End If

End Property

Public Property Let URL(ByVal psNewValue As String)
  txtURL.Text = IIf(optLink(SSINTLINKSCREEN_URL).value Or optLink(SSINTLINKSCREEN_EMAIL).value, psNewValue, "")
End Property


Public Property Let EMailAddress(ByVal psNewValue As String)
  txtEmailAddress.Text = IIf(optLink(SSINTLINKSCREEN_EMAIL).value, psNewValue, "")
End Property

Public Property Let EMailSubject(ByVal psNewValue As String)
  txtEmailSubject.Text = IIf(optLink(SSINTLINKSCREEN_EMAIL).value, psNewValue, "")
End Property

Public Property Let AppFilePath(ByVal psNewValue As String)
  txtAppFilePath.Text = IIf(optLink(SSINTLINKSCREEN_APPLICATION).value, psNewValue, "")
End Property

Public Property Let AppParameters(ByVal psNewValue As String)
  txtAppParameters.Text = IIf(optLink(SSINTLINKSCREEN_APPLICATION).value, psNewValue, "")
End Property

Public Property Let DocumentFilePath(ByVal psNewValue As String)
  txtDocumentFilePath.Text = IIf(optLink(SSINTLINKSCREEN_DOCUMENT).value, psNewValue, "")
End Property

Public Property Let DisplayDocumentHyperlink(ByVal pbNewValue As Boolean)
  chkDisplayDocumentHyperlink.value = IIf(optLink(SSINTLINKSCREEN_DOCUMENT).value, IIf(pbNewValue, vbChecked, vbUnchecked), vbUnchecked)
End Property

Private Sub txtURL_GotFocus()
  UI.txtSelText
End Sub

Public Property Get Prompt() As String
  Prompt = IIf(miLinkType = SSINTLINK_BUTTON, txtPrompt.Text, "")
End Property

Public Property Let Prompt(ByVal psNewValue As String)
  txtPrompt.Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
'  txtLinkSeparator.Text = psNewValue
End Property

Public Property Get HRProScreenID() As String

  If (cboHRProScreen.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_HRPRO).value) Then
    
    HRProScreenID = ""
  Else
    HRProScreenID = CStr(cboHRProScreen.ItemData(cboHRProScreen.ListIndex))
  End If

End Property

Public Property Get StartMode() As String

  If (cboHRProScreen.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_HRPRO).value) Then
    
    StartMode = ""
  Else
    StartMode = CStr(cboStartMode.ItemData(cboStartMode.ListIndex))
  End If

End Property

Public Property Let HRProScreenID(ByVal psNewValue As String)

  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim sSQL As String
  Dim rsScreens As dao.Recordset
  
  If (miLinkType <> SSINTLINK_HYPERTEXT) And _
    (optLink(SSINTLINKSCREEN_HRPRO).value) And _
    (Len(psNewValue) > 0) Then
    ' Get the given screen's table.
    sSQL = "SELECT tmpScreens.tableID" & _
      " FROM tmpScreens" & _
      " WHERE (tmpScreens.deleted = FALSE)" & _
      " AND (tmpScreens.screenID = " & psNewValue & ")"
    Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    If Not (rsScreens.EOF And rsScreens.BOF) Then
      For iLoop = 0 To cboHRProTable.ListCount - 1
        If cboHRProTable.ItemData(iLoop) = CLng(rsScreens!TableID) Then
          cboHRProTable.ListIndex = iLoop
          
          For iLoop2 = 0 To cboHRProScreen.ListCount - 1
            If cboHRProScreen.ItemData(iLoop2) = CLng(psNewValue) Then
              cboHRProScreen.ListIndex = iLoop2
              Exit For
            End If
          Next iLoop2
          
          Exit For
        End If
      Next iLoop
    End If
    rsScreens.Close
    Set rsScreens = Nothing
  End If
  
End Property

Public Property Let StartMode(ByVal psNewValue As String)
  Dim iLoop As Integer
  
  If (miLinkType <> SSINTLINK_HYPERTEXT) And _
    (optLink(SSINTLINKSCREEN_HRPRO).value) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboStartMode.ListCount - 1
      If cboStartMode.ItemData(iLoop) = CLng(psNewValue) Then
        cboStartMode.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
End Property

Public Property Let UtilityID(ByVal psNewValue As String)

  Dim iLoop As Integer
  
  If (optLink(SSINTLINKSCREEN_UTILITY).value) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtility.ListCount - 1
      If cboHRProUtility.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtility.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
End Property

Public Property Let UtilityType(ByVal psNewValue As String)

  Dim iLoop As Integer

  If (optLink(SSINTLINKSCREEN_UTILITY).value) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtilityType.ListCount - 1
      If cboHRProUtilityType.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtilityType.ListIndex = iLoop
        Exit For
      End If
    Next iLoop

    GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
  End If
    
End Property

Public Property Get PageTitle() As String
  PageTitle = IIf(optLink(SSINTLINKSCREEN_HRPRO).value, _
    txtPageTitle.Text, "")
End Property

Public Property Let PageTitle(ByVal psNewValue As String)
  txtPageTitle.Text = IIf(optLink(SSINTLINKSCREEN_HRPRO).value, _
    psNewValue, "")
End Property
 
Public Property Get PictureID() As Long
  ' Return the selected picture's ID.
  PictureID = glngPictureID
End Property
 
Public Property Get SeparatorOrientation() As Integer
  ' Return the selected separator orientation.
  SeparatorOrientation = miSeparatorOrientation
End Property
 
Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Get ViewID() As Long
  ViewID = mlngViewID
End Property

Public Property Get TableViewName() As String
  TableViewName = msTableViewName
End Property

Public Property Get ChartType() As Integer
  ChartType = miChartType
End Property

Public Property Let ChartType(ByVal piNewValue As Integer)
  miChartType = piNewValue
  
  Dim iLoop As Integer

  If (optLink(SSINTLINKCHART).value) And piNewValue > 0 Then

    For iLoop = 0 To cboChartType.ListCount - 1
      If cboChartType.ItemData(iLoop) = piNewValue Then
        cboChartType.ListIndex = iLoop
        Exit For
      End If
    Next iLoop

  End If
  
End Property

Public Property Get ChartViewID() As Long
  ChartViewID = miChartViewID
End Property

Public Property Let ChartViewID(ByVal plngNewValue As Long)
  miChartViewID = plngNewValue
End Property

Public Property Get ChartFilterID() As Long
  ChartFilterID = miChartFilterID
End Property

Public Property Let ChartFilterID(ByVal plngNewValue As Long)
  miChartFilterID = plngNewValue
End Property

Public Property Get ChartTableID() As Long
  ChartTableID = miChartTableID
End Property

Public Property Let ChartTableID(ByVal plngNewValue As Long)
  miChartTableID = plngNewValue
End Property

Public Property Get ChartColumnID() As Long
  ChartColumnID = miChartColumnID
End Property

Public Property Let ChartColumnID(ByVal plngNewValue As Long)
  miChartColumnID = plngNewValue
End Property

Public Property Get ChartAggregateType() As Integer
  ChartAggregateType = miChartAggregateType
End Property

Public Property Let ChartAggregateType(ByVal piNewValue As Integer)
  miChartAggregateType = piNewValue
End Property

Public Property Get ElementType() As Integer
  ElementType = miElementType
End Property

Public Property Let ElementType(ByVal piNewValue As Integer)
  miElementType = piNewValue
End Property

Public Property Get ChartShowLegend() As Boolean
  ChartShowLegend = chkShowLegend.value
End Property

Public Property Let ChartShowLegend(ByVal pfNewValue As Boolean)
  chkShowLegend.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get ChartShowGridlines() As Boolean
  ChartShowGridlines = chkDottedGridlines.value
End Property

Public Property Let ChartShowGridlines(ByVal pfNewValue As Boolean)
  chkDottedGridlines.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get ChartStackSeries() As Boolean
  ChartStackSeries = chkStackSeries.value
End Property

Public Property Let ChartStackSeries(ByVal pfNewValue As Boolean)
  chkStackSeries.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get ChartShowValues() As Boolean
  ChartShowValues = chkShowValues.value
End Property

Public Property Let ChartShowValues(ByVal pfNewValue As Boolean)
  chkShowValues.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property











