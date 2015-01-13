VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "actbar.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "coa_spinner.ocx"
Object = "{96E404DC-B217-4A2D-A891-C73A92A628CC}#1.0#0"; "coa_workingpattern.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "coa_line.ocx"
Object = "{E28F058B-8430-42F1-9D74-4BEDD2F27CCE}#1.0#0"; "COA_OptionGroup.ocx"
Object = "{4FD0EB05-F124-4460-A61D-CB587234FB75}#1.0#0"; "COA_Image.ocx"
Object = "{EDB7B7A8-7908-4896-B964-57CE7262666E}#1.0#0"; "COA_OLE.ocx"
Object = "{A48C54F8-25F4-4F50-9112-A9A3B0DBAD63}#1.0#0"; "coa_label.ocx"
Object = "{3389D561-C8E1-4CB0-A73E-77582EA68D3C}#1.1#0"; "COA_Lookup.ocx"
Object = "{AD837810-DD1E-44E0-97C5-854390EA7D3A}#3.2#0"; "coa_navigation.ocx"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "coa_colourpicker.ocx"
Object = "{19400013-2704-42FE-AAA4-45D1A725A895}#1.0#0"; "coa_colourselector.ocx"
Begin VB.Form frmRecEdit4 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1054
   Icon            =   "frmRecEdit4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   6465
   Begin COAColourPicker.COA_ColourPicker ColourPicker1 
      Left            =   4815
      Top             =   45
      _ExtentX        =   820
      _ExtentY        =   820
      ShowSysColorButton=   0   'False
   End
   Begin VB.Frame fraTabPage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4500
      Index           =   0
      Left            =   210
      TabIndex        =   2
      Top             =   795
      Visible         =   0   'False
      Width           =   6000
      Begin COAColourSelector.COA_ColourSelector ColourSelector1 
         Height          =   315
         Index           =   0
         Left            =   2205
         TabIndex        =   21
         Top             =   3285
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
      End
      Begin COANavigation.COA_Navigation COA_Navigation1 
         Height          =   645
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1530
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1138
         Caption         =   "Navigate..."
         DisplayType     =   1
         NavigateIn      =   0
         NavigateTo      =   ""
         InScreenDesigner=   0   'False
         ColumnID        =   0
         ColumnName      =   ""
         Selected        =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontSize        =   8.25
         FontStrikethrough=   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   0
         BackColor       =   -2147483633
         NavigateOnSave  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox Frame1 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   585
         Width           =   1500
         _Version        =   851969
         _ExtentX        =   2646
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "GroupBox1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin TDBMask6Ctl.TDBMask TDBMask1 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   1800
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   582
         Caption         =   "frmRecEdit4.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmRecEdit4.frx":0071
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
         Format          =   "&&&&&&&&&&"
         HighlightText   =   2
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
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "          "
         Value           =   ""
      End
      Begin TDBText6Ctl.TDBText TDBText1 
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   17
         Top             =   990
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   503
         Caption         =   "frmRecEdit4.frx":00B3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRecEdit4.frx":0118
         Key             =   "frmRecEdit4.frx":0136
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   -1
         ScrollBars      =   2
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin XtremeSuiteControls.CheckBox Check1 
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   15
         Top             =   3015
         Width           =   1455
         _Version        =   851969
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "CheckBox1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   14
         Top             =   2595
         Visible         =   0   'False
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmRecEdit4.frx":017A
         Caption         =   "frmRecEdit4.frx":019A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRecEdit4.frx":01FF
         Keys            =   "frmRecEdit4.frx":021D
         Spin            =   "frmRecEdit4.frx":026D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0; -########0;0;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########; -########"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   -99999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1919877125
         MinValueVT      =   1685389317
      End
      Begin COALine.COA_Line ASRLine1 
         Height          =   30
         Index           =   0
         Left            =   2535
         Top             =   2895
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   53
      End
      Begin VB.Frame OLEFrame 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   915
         Index           =   0
         Left            =   4140
         TabIndex        =   12
         Top             =   135
         Visible         =   0   'False
         Width           =   1140
         Begin COAOLE.COA_OLE OLE1 
            Height          =   750
            Index           =   0
            Left            =   90
            TabIndex        =   13
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1323
            OLEType         =   0
         End
      End
      Begin GTMaskDate.GTMaskDate GTMaskDate1 
         CausesValidation=   0   'False
         Height          =   300
         Index           =   0
         Left            =   4050
         TabIndex        =   11
         Top             =   1755
         Width           =   1560
         _Version        =   65537
         _ExtentX        =   2752
         _ExtentY        =   529
         _StockProps     =   77
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
         CaretPicture    =   "frmRecEdit4.frx":0295
         NullText        =   "__/__/____"
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelect      =   -1  'True
         MaskCentury     =   2
         ValidateRangeHigh=   "31/12/9999"
         ValidateRangeLow=   "01/01/1753"
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTips        =   0   'False
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin COAOptionGroup.COA_OptionGroup OptionGroup1 
         Height          =   645
         Index           =   0
         Left            =   2115
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1138
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4000
         TabIndex        =   5
         Top             =   3420
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   4
         Top             =   2200
         Visible         =   0   'False
         Width           =   1500
      End
      Begin COAImage.COA_Image ASRUserImage1 
         Height          =   315
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   1380
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
      End
      Begin COALookup.COA_Lookup ctlNewLookup1 
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   6
         Top             =   3800
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
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
         FontSize        =   8.25
         BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin COALabel.COA_Label Label1 
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "Label1"
         FontSize        =   8.25
      End
      Begin COASpinner.COA_Spinner Spinner1 
         Height          =   315
         Index           =   0
         Left            =   4000
         TabIndex        =   8
         Top             =   1400
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
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
         MaximumValue    =   99999
      End
      Begin COAWorkingPattern.COA_WorkingPattern ASRWorkingPattern1 
         Height          =   765
         Index           =   0
         Left            =   4005
         TabIndex        =   9
         Top             =   2205
         Visible         =   0   'False
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   1349
      End
      Begin XtremeSuiteControls.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2475
         TabIndex        =   16
         Top             =   4050
         Width           =   1545
         _Version        =   851969
         _ExtentX        =   2725
         _ExtentY        =   556
         _StockProps     =   77
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
         Style           =   2
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4995
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8811
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5550
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   4005
      Tag             =   "BAND_RECORDEDIT"
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmRecEdit4.frx":05B7
   End
   Begin VB.Menu mnuOLE 
      Caption         =   "OLE"
      Visible         =   0   'False
      Begin VB.Menu mnuOLEInsert 
         Caption         =   "&Insert"
      End
      Begin VB.Menu mnuOLEEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuOLEDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuOLESep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOLEPaste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "frmRecEdit4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MH20010108
'Search for the comment "MH20010108" for
'fixes for international numeric columns

'Private mobjBorders As clsBorders

Public Event OLEClick(plngColumnID As Long, psFile As String, pfOLEOnServer As Boolean)

' Form variables.
Private mlngFormID As Long
Private mlngParentFormID As Long
Private mfrmFind As frmFind2
Private mfrmSummary As frmFind2
Private mfParentUnload As Boolean

' Screen variables.
Private mlngScreenID As Long
Private miScreenType As ScreenType
Private msScreenName As String
Private mfUseTab As Boolean
Public mobjScreenControls As clsScreenControls

' Table/view variables.
Private msTableViewName As String
Private mobjTableView As CTablePrivilege
Private mcolColumnPrivileges As CColumnPrivileges
Private malngLinkRecordIDs() As Long
Private mlngRecDescID As Long
Private mfRequiresLocalCursor As Boolean

' Parent table/view variables.
Private mlngParentTableID As Long
Private mlngParentViewID As Long
Private mlngParentRecordID As Long

' Recordset variables.
Private mrsRecords As ADODB.Recordset
Private mlngOrderID As Long
Private mavFilterCriteria() As Variant
Private msHistorySQL As String
Private msCountSQL As String
Private mlngRecordCount As Long
Private mfFirstOrderColumnAscending As Boolean
Private mlngFirstOrderColumnID As Long
' JPD 30/8/00 Multi-user correction.
Private mlngTimeStamp As Long
Private mlngRecordID As Long

' General form handling variables.
Private mfLoading As Boolean
Private mfCancelled As Boolean
Private mfTableEntry As Boolean
Private mvOldValue As Variant
Private mfControlProcessing As Boolean
Private mfDataChanged As Boolean
Private mfUnloading As Boolean
Private mbResendingToAccord As Boolean

' Utility classes.
Private mdatRecEdit As clsRecEdit
Private ODBC As New ODBC

' Lookup control variables.
Private mfLookup As Boolean
Private mfLeaveLookup As Boolean
Private mfLookupLoading As Boolean

'JPD20010815 Fault 2239 Training Booking specific variables.
Private msTBOriginalStatus As String
Private mlngTBOriginalEmpID As Long
Private mlngTBOriginalCourseID As Long

' JPD20021007 Fault 4498
Private malngChangedOLEPhotos() As Long

Private msFindCaption As String
Private msStatusCaption As String
Private msFindStatusCaption As String
Private msFindPrintHeader As String
Private mlngScreenIndex As Long

' JPD20021206 Fault 4854
Private mfSavingInProgress As Boolean
' JPD20030311 Fault 5142
Private mfAddingNewInProgress As Boolean

Private mlngUpdatedMultiLineControl As Long

Private mblnScreenHasAutoUpdate As Boolean

Private mbDisableAURefresh As Boolean 'flag used to disable the refreshing of the auto update screens.
Private mlngPictureID As Long

Public OriginalRecordID As Long

Public Property Get RequiresLocalCursor() As Boolean
  RequiresLocalCursor = mfRequiresLocalCursor
End Property


Public Function SaveFromChild() As Boolean
  ' JPD20021206 Fault 4863
  Dim fOK As Boolean
  
  fOK = True
  
  ' JPD20030206 Fault 5027
  If Not mrsRecords Is Nothing Then
    If mrsRecords.State <> adStateClosed Then
      If mrsRecords.EditMode = adEditAdd Then
        mfDataChanged = True
        fOK = SaveChanges(False, False, True)
        If Not fOK Then
          ' Save changes cancelled, or invalid. Do not deactivate the form.
          If Me.Visible And Me.Enabled Then
            Me.SetFocus
            frmMain.RefreshMainForm Me, False
          End If
        Else
          fOK = SaveAscendants
        End If
      End If
    End If
  End If

  SaveFromChild = fOK
  
End Function



Public Function SaveAscendants() As Boolean
  ' JPD20021206 Fault 4863
  ' Prompt the user to save changes, eben if no changes have been made,
  ' if the current record is a new one (ie. added or copied) and the user
  ' is trying to activate a history screen.
  Dim frmForm As Form
  Dim fOK As Boolean
  
  fOK = True
  
  If mlngParentFormID > 0 Then
    For Each frmForm In Forms
      With frmForm
        If .Name = "frmRecEdit4" Then
          If .FormID = mlngParentFormID Then
            fOK = .SaveFromChild
            
            Exit For
          End If
        End If
      End With
    Next frmForm
    Set frmForm = Nothing
  End If

  SaveAscendants = fOK

End Function



Public Sub SetLookupValue(plngColumnID As Long, psNewValue As String)
  Dim objControl As Control
  Dim sTag As String
  Dim sFormat As String
  
  For Each objControl In Me.Controls
    With objControl
      ' Get the control's tag.
      sTag = .Tag
        
      'JPD 20030610
      If TypeOf objControl Is ActiveBar Then
        sTag = ""
      End If
      
      ' Check if the control is tagged to a column.
      If Len(sTag) > 0 Then
        If mobjScreenControls.Item(sTag).ColumnID = plngColumnID Then
          'JPD 20050810 Fault 10165
          If (mobjScreenControls.Item(sTag).DataType = sqlNumeric) Then
            sFormat = "0"
            If mobjScreenControls.Item(sTag).Use1000Separator Then
              sFormat = "#,0"
            End If
            If mobjScreenControls.Item(sTag).Decimals > 0 Then
              sFormat = sFormat & "." & String(mobjScreenControls.Item(sTag).Decimals, "0")
            End If
                        
            .Text = Format(psNewValue, sFormat)
          Else
            .Text = psNewValue
          End If
        End If
      End If
    End With
  Next objControl
  Set objControl = Nothing
  
End Sub

Public Sub AddNewCopyOf()
  ' Copies a record.
  Dim objOLEControl As COA_OLE
  Dim objImageControl As COA_Image
  Dim sTag As String
  Dim objControl As Control
  Dim fResetControl As Boolean
  
  ' Save changes if required.
  If SaveChanges Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If

    ' Check the user has permission to add new records to the table/view.
    If mobjTableView.AllowInsert Then
      mrsRecords.CancelUpdate
      mrsRecords.AddNew
      
      'NHRD15072003 Fault 4719
      Me.Caption = "Copy of " + Me.Caption
      RemoveIcon Me

      ' JDM - Fault 8597 - 02/06/2004 - Copying OLEs that haven't been yet loaded
      For Each objOLEControl In OLE1
        sTag = objOLEControl.Tag
        If Len(sTag) > 0 Then
          If Not objOLEControl.EmbeddedStream.State = adStateOpen Then
            ReadStream objOLEControl, mobjScreenControls.Item(sTag), False
          End If
        End If
      Next objOLEControl

      OriginalRecordID = mlngRecordID

      ' JPD20021025 Fault 4647
      mlngTimeStamp = 0
      mlngRecordID = 0
      
      mbDisableAURefresh = True
      
      ' JPD20030311 Fault 5142
      mfAddingNewInProgress = True
      
      'Loop through all screen controls resetting the unique column values.
      For Each objControl In Me.Controls
        With objControl
          sTag = .Tag
          
          If TypeOf objControl Is ActiveBar Then
            sTag = ""
          End If
          
          If Len(sTag) > 0 Then
            If mobjScreenControls.Item(sTag).ColumnID > 0 Then
                          
              fResetControl = mobjScreenControls.Item(sTag).UniqueCheck
              
              If Not fResetControl Then
                If Not mcolColumnPrivileges.FindColumnID(mobjScreenControls.Item(sTag).ColumnID) Is Nothing Then
                  fResetControl = Not mcolColumnPrivileges.FindColumnID(mobjScreenControls.Item(sTag).ColumnID).AllowUpdate
                End If
              End If
              
              If fResetControl Then
                SetControlDefaults .Tag, objControl, mlngParentTableID, mlngParentRecordID
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
            
      UpdateChildren
      mfAddingNewInProgress = False

      mfDataChanged = True
      
      ' JPD20021007 Fault 4498
      For Each objImageControl In ASRUserImage1
        sTag = objImageControl.Tag
        
        If Len(sTag) > 0 Then
          ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
          malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = mobjScreenControls.Item(sTag).ColumnID
        End If
      Next objImageControl
      Set objImageControl = Nothing
      
      For Each objOLEControl In OLE1
        sTag = objOLEControl.Tag
        
        If Len(sTag) > 0 Then
          ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
          malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = mobjScreenControls.Item(sTag).ColumnID
        End If
      Next objOLEControl
      Set objOLEControl = Nothing
    Else
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      'COAMsgBox "You do not have 'new' permission on this " & IIf(mobjTableView.ViewID > 0, "view.", "table."), vbExclamation, "Security"
      COAMsgBox "You do not have 'new' permission on this " & IIf(Me.ViewID > 0, "view.", "table."), vbExclamation, "Security"
      Exit Sub
    End If
  End If
  
  ' Highlight the first object
  FocusFirstControl
  
  ' Refresh the main menu.
  frmMain.RefreshMainForm Me
  
End Sub


Private Sub FocusFirstControl()

  'Highlights the first control on the form

  On Error Resume Next
  
  Dim objControl As Control
  
  'JDM - 22/10/01 - Fault 2933 - Only setfocus if form is not minimised
  If WindowState <> vbMinimized Then
    For Each objControl In Me.Controls
      If Not TypeOf objControl Is ActiveBar And Not TypeOf objControl Is Menu Then
        If Len(objControl.Tag) > 0 Then
          If objControl.Visible Then
            If mobjScreenControls.Collection(objControl.Tag).TabIndex = 1 Then
              objControl.SelStart = 0
              'TM20011126 Fault 3168
              If TypeOf objControl Is GTMaskDate.GTMaskDate Then
                objControl.SelLength = Len(objControl.InputText)
              Else
                objControl.SelLength = Len(objControl.Text)
              End If
              objControl.SetFocus
              Exit For
            End If
          End If
        End If
      End If
    Next objControl
  End If

End Sub

Private Sub FocusCurrentControl()

  'Highlights the current control on the form

  On Error Resume Next

  ActiveControl.SelStart = 0
  ActiveControl.SelLength = Len(ActiveControl.Text)
  ActiveControl.SetFocus

End Sub


Public Property Get TableName() As String
  ' Return the name of the screen's associated table.
  TableName = msTableViewName

End Property


Public Sub EnableMe()
  ' Enable the record edit screen and its find/summary window.
  Me.Enabled = True
  
  If Not mfrmFind Is Nothing Then
    mfrmFind.Enabled = True
  End If
  
  If Not mfrmSummary Is Nothing Then
    mfrmSummary.Enabled = True
  End If

End Sub


Public Property Get NewTableEntry() As Boolean
  NewTableEntry = mfTableEntry

End Property

Public Property Let NewTableEntry(ByVal pfNewValue As Boolean)
  mfTableEntry = pfNewValue

End Property

Public Sub UpdateAll()
  ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
  ' UpdateChildren call as the summary fields that are updated in
  ' UpdateChildren may be dependent on the parent recordset being
  ' refreshed first (in UpdateParentWindow).
  UpdateParentWindow
  UpdateControls
  UpdateChildren
  UpdateFindWindow
End Sub

Private Sub ActiveBar1_Click(ByVal pTool As ActiveBarLibraryCtl.Tool)
  ' Perform the selected menu option action.


  'MH20001025 Fault 1080
  'Need to check if other RecEdit screen have changed
  'This is due to a problem which only happens in the EXE
  'and this is the only way that I found of getting around it.
  If mobjTableView.TableType <> tabLookup Then
    Dim f As Form
    For Each f In Forms
      If TypeOf f Is frmRecEdit4 Then
        If Me.TableID <> f.TableID Or Me.ViewID <> f.ViewID Then
          If f.RecordChanged Then
            Exit Sub
          End If
        End If
      End If
    Next
  End If


  Select Case pTool.Name
    '
    ' <Record> menu.
    '
    ' <New>
    Case "NewRecord"
      AddNew

    ' <Copy this record>
    Case "CopyRecord"
      AddNewCopyOf

    ' <Save>
    Case "SaveRecord"
      If Not Screen.ActiveControl Is Nothing Then
        If LostFocusCheck(Screen.ActiveControl) = True Then
          'COAMsgBox "true"
        Else
          Exit Sub
        End If
      End If
      
      UpdateWithAVI
      
    ' <Delete>
    Case "DeleteRecord"
      DeleteRecord
    
    ' <First Record>
    Case "FirstRecord"
      'NPG20080516 Fault 12973
      EnableNavigation False
      MoveFirst
    
    ' <Previous Record>
    Case "PreviousRecord"
      'NPG20080516 Fault 12973
      EnableNavigation False
      MovePrevious
    
    ' <Next Record>
    Case "NextRecord"
      'NPG20080516 Fault 12973
      EnableNavigation False
      MoveNext
    
    ' <Last Record>
    Case "LastRecord"
      'NPG20080516 Fault 12973
      EnableNavigation False
      MoveLast
    
    ' <Find>
    Case "FindRecord"
      Find
    
    ' <QuickFind>

    Case "QuickFind"
      SelectQuickFind
      
    ' <Refresh>
    Case "Refresh"
      Requery False
    ' <Order>
    Case "Order"
      SelectOrder
    ' <Filter>
    Case "Filter"
      SelectFilter
      
    Case "FilterClear"
      ClearFilter
      
    Case "MailMerge"
      MailMergeClick
    
    Case "LabelsRec"
      LabelsClick
    
    Case "DataTransfer"
      DataTransferClick
    
    Case "AbsenceBreakdownRec"
      AbsenceBreakdownClick

    Case "AbsenceCalendar"
      AbsenceCalendarClick

    Case "BradfordFactorRec"
      BradfordFactorClick

    Case "Email"
      EmailClick
    
    Case "MatchReportRec"
      MatchReportClick mrtNormal
    
    Case "SuccessionRec"
      'CareerSuccessionClick mrtSucession
      MatchReportClick mrtSucession
    
    Case "CareerRec"
      'CareerSuccessionClick mrtCareer
      MatchReportClick mrtCareer

    Case "RecordProfileRec"
      RecordProfileClick
    
    Case "CalendarReportRec"
      CalendarReportClick
      
    ' <Module specifics> menu
    '
    ' <Cancel Course>
    Case "CancelCourse"
        CancelCourse
  
  End Select
  
'  ' JPD20021126 Fault 4676
'  If mrsRecords.State = adStateClosed Then
'    frmMain.RefreshMainForm Me, True
'    ' DoEvents is not enough. Refresh the screen with an API call.
'    UI.RedrawScreen frmMain.hWnd
'  End If
    
End Sub

Private Sub SelectQuickFind()

  ' Display the quick find form.
  If RecordCount = 0 Then
    COAMsgBox "No records exist.", vbExclamation, app.ProductName
    Exit Sub
  ElseIf (RecordCount = 1) And (mrsRecords.EditMode = adEditAdd) Then
    COAMsgBox "No records exist.", vbExclamation, app.ProductName
    Exit Sub
  End If
    
  If Not SaveChanges Then
    Exit Sub
  End If
  
  If (mrsRecords.EditMode = adEditAdd) And _
    (Not Database.Validation) Then
    Exit Sub
  End If
  
  If (mrsRecords.BOF And mrsRecords.EOF) Then
    AddNew
    Exit Sub
  End If

  frmQuickFind.Initialise Me

  ' JPD20030211 Fault 5043
  If Not frmQuickFind.Cancelled Then
    UpdateChildren
  End If

End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  'TM20030430 Fault - need to set this to True NOT False!!
'  ' Do not let the user modify the layout.
'  Cancel = False
  Cancel = True

End Sub

Private Sub ASRUserImage1_LostFocus(Index As Integer)

  ASRUserImage1(Index).ShowSelectionMarkers = False
  
End Sub

Private Sub ASRUserImage1_SpacePressed(Index As Integer)

  ASRUserImage1_Click (Index)
  
End Sub

Private Sub COA_Navigation1_ToolClickRequest(Index As Integer, ByVal Tool As String)

  On Error GoTo ErrorTrap

  Dim objTool As New ActiveBarLibraryCtl.Tool
  Dim sTool As Variant
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim asTools() As String
  Dim bPreviousCloseDefSelAfterRun As Boolean
  Dim sParameter As String
  
  gbCloseDefSelAfterRun = True
  gblnBatchMode = True
  gbJustRunIt = True
  
  gobjProgress.Visible = True
  gobjProgress.Bar1Caption = "Batch Job"
  
  asTools = Split(Tool, ";")
  For Each sTool In asTools
    
    lngStart = InStr(1, sTool, "(")
    lngEnd = InStr(1, sTool, ")")
    If lngStart > 0 And lngEnd > 0 Then
      sParameter = Mid(sTool, lngStart + 1, (lngEnd - lngStart) - 1)
      If IsNumeric(sParameter) Then
        glngBypassDefsel_ID = CLng(sParameter)
      Else
        ' Calculate the parameter code
      End If
      
      'objTool.Tag = Replace(sParameter, """", "")
      objTool.Name = Mid(sTool, 1, lngStart - 1)
    Else
      objTool.Name = sTool
    End If
    
    frmMain.abMain_Click objTool
    
    glngBypassDefsel_ID = 0
  
  Next sTool

TidyUpAndExit:
  gbCloseDefSelAfterRun = bPreviousCloseDefSelAfterRun
  gblnBatchMode = False
  gbJustRunIt = False
  gobjProgress.Visible = False
  
  Exit Sub

ErrorTrap:
  GoTo TidyUpAndExit

End Sub

Private Sub ColourSelector1_Click(Index As Integer)

  On Error GoTo ErrorTrap

  With ColourSelector1(Index)
    ColourPicker1.Color = .BackColor
    ColourPicker1.ShowPalette
    If .BackColor <> ColourPicker1.Color Then
      .BackColor = ColourPicker1.Color
      mfDataChanged = True
      frmMain.RefreshMainForm Me
    End If
  End With

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

  ' RH 31/07/00 - If space bar is pressed, then drop down the combo box
  If KeyAscii = 32 Then
    UI.cboDropDown Combo1(Index).hWnd, True
  End If
  
End Sub

Private Sub command1_GotFocus(Index As Integer)
    'ND20020225 Fault 3533 - Added this code to the GotFocus Method
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck Command1(Index)
End Sub

Private Sub Command2_Click()

ActiveBar1.Customize

End Sub

Private Sub Form_Deactivate()

  Dim iTabPage As Integer
  Dim ctlControl As Control
  Dim frmNewForm As Form
  Dim fOK As Boolean
  Dim intErrorLine As Integer

  On Error GoTo LocalErr
  
  DebugOutput "frmRecEdit4", "Form_Deactivate"
  
  intErrorLine = 10
  
  DoEvents
  
  ' JPD20020926 Fault 4414
  Screen.MousePointer = vbHourglass

  fOK = True
  
'  Set frmNewForm = frmMain.ActiveForm
'  If TypeOf frmNewForm Is frmRecEdit4 Then
'    Set ctlControl = frmNewForm.ActiveControl
'    If frmNewForm.TabStrip1.Visible Then
'      iTabPage = frmNewForm.TabStrip1.SelectedItem.Index
'    End If
'  End If
  
  intErrorLine = 20
  
  ' Prompt to save changes (if required).
  fOK = mfLookup
  If Not fOK Then
    'TM20020607 Fault 3963 - don't try and do the save changes stuff if we have already cancelled.
    If Not Me.Cancelled Then
      intErrorLine = 30
      fOK = SaveChanges(False, False, True)
      intErrorLine = 40
      frmMain.RefreshMainForm Screen.ActiveForm, False
      intErrorLine = 50
    End If
  End If

  If Not fOK Then
    ' Save changes cancelled, or invalid. Do not deactivate the form.
    If Me.Visible And Me.Enabled Then
      intErrorLine = 60
      Me.SetFocus
    End If
    
    'TM20020610 Fault 3966 - reset the cancelled flag.
    mfCancelled = False
    
    intErrorLine = 70
'  Else
'    If TypeOf frmNewForm Is frmRecEdit4 Then
'      If frmNewForm.TabStrip1.Visible Then
'        frmNewForm.TabStrip1.Tabs.Item(iTabPage).Selected = True
'      End If
'
'      If Not ctlControl Is Nothing Then
'        If Not TypeOf ctlControl Is TabStrip Then
'          'If ctlControl.Visible And ctlControl.Enabled Then
'          '  ctlControl.SetFocus
'          'End If
'          ControlSetFocus ctlControl
'        End If
'      End If
'    End If
'
'    'With frmNewForm.ActiveControl
'    '  If .Visible And .Enabled Then
'    '    .SetFocus
'    '  End If
'    'End With
'    ControlSetFocus frmNewForm.ActiveControl
'
'    ' Set up the menus and toolbars
'    frmMain.RefreshMainForm frmNewForm
  End If
  
  intErrorLine = 80
  
  ' JPD20020926 Fault 4414
  Screen.MousePointer = vbDefault

Exit Sub

LocalErr:
  COAMsgBox Err.Description & vbCrLf & "(RecEdit4 - Form_Deactivate " & CStr(intErrorLine) & ")", vbCritical

End Sub

Private Sub GTMaskDate1_Change(Index As Integer)
  
  On Error GoTo Err_Trap

  Dim fGoodDate As Boolean
  Dim iDatePart1 As Integer
  Dim iDatePart2 As Integer
  Dim iDatePart3 As Integer
  Dim dtDate As Date
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sDate As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control

  If Not mfLoading Then

    ' RH BUG 825 - If the value of the date control is the same as the
    '              recordset, then skip this sub, because it enables the
    '              save button !
    If GTMaskDate1(Index).DateValue = mrsRecords.Fields(mobjScreenControls.Item(GTMaskDate1(Index).Tag).ColumnName) Then Exit Sub
    
    fGoodDate = True
    
    iDatePart1 = Val(GTMaskDate1(Index).Text) Mod 100
    iDatePart2 = Val(Mid(GTMaskDate1(Index).Text, InStr(1, GTMaskDate1(Index).Text, UI.GetSystemDateSeparator) + 1)) Mod 100
    iDatePart3 = Val(Mid(GTMaskDate1(Index).Text, InStr(InStr(1, GTMaskDate1(Index).Text, UI.GetSystemDateSeparator) + 1, GTMaskDate1(Index).Text, UI.GetSystemDateSeparator) + 1)) Mod 100
  
    ' Check if the entered text prodcues a valid date.
    If Not IsNull(GTMaskDate1(Index).DateValue) Then
      sDate = CStr(GTMaskDate1(Index).DateValue)
      fGoodDate = (iDatePart1 = Val(sDate) Mod 100) And _
      (iDatePart2 = Val(Mid(sDate, InStr(1, sDate, UI.GetSystemDateSeparator) + 1)) Mod 100) And _
      (iDatePart3 = Val(Mid(sDate, InStr(InStr(1, sDate, UI.GetSystemDateSeparator) + 1, sDate, UI.GetSystemDateSeparator) + 1)) Mod 100)
    End If
  
    ' Get the control's tag.
    sTag = GTMaskDate1(Index).Tag

    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName

      ' Update all other screen control's that represent the same column.
      For Each objControl In GTMaskDate1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag

            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                If .Text <> GTMaskDate1(Index).Text Then
                  mfLoading = True
                  .DateValue = IIf(fGoodDate And Not IsEmpty(GTMaskDate1(Index).DateValue), GTMaskDate1(Index).DateValue, Null)
                  mfLoading = False
                End If
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing

       'Set the 'changed' flag if required.
      If Not mfDataChanged Then
        mfDataChanged = (GTMaskDate1(Index).Text <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName)))
        frmMain.RefreshMainForm Me
      End If
    End If
  End If

  Exit Sub

Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True
    frmMain.RefreshMainForm Me
  End If

End Sub

Private Sub GTMaskDate1_GotFocus(Index As Integer)
  
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck GTMaskDate1(Index)
  
  ' JPD20020920 Fault 4423
  'MH20020813 Fault 4300
  'Need to make sure that the control is enabled before setting focus to it.
  ''TM20020524 Fault 2301
  'Me.GTMaskDate1(Index).SetFocus
  'If Me.GTMaskDate1(Index).Enabled Then
  If Me.GTMaskDate1(Index).Enabled And GTMaskDate1(Index).Visible Then
    ' JPD20020926 Fault 4451
    
    ' JPD20021106 Fault 4691
    If Not frmMain.ActiveForm Is Me Then
      Me.SetFocus
    End If
'Me.GTMaskDate1(Index).SetFocus
  End If

End Sub

Private Sub GTMaskDate1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    GTMaskDate1(Index).DateValue = Date
  End If
End Sub

Private Sub GTMaskDate1_LostFocus(Index As Integer)
  
  ' RH 05/03/01 - To try and keep consistency with the date control.
  '               Bit of a pain in recedit as have to ignore lostfocuscheck
  '               and switch the controls causesvalidation property to false.
  
'  If Not IsDate(GTMaskDate1(Index).DateValue) And _
'     GTMaskDate1(Index).Text <> "  /  /" Then
'
'     GTMaskDate1(Index).ForeColor = vbRed
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     GTMaskDate1(Index).ForeColor = vbWindowText
'     GTMaskDate1(Index).DateValue = Null
'     GTMaskDate1(Index).SetFocus
'     Exit Sub
'  ElseIf GTMaskDate1(Index).DateValue < "01/01/1753" Or GTMaskDate1(Index).DateValue > "31/12/9999" Then
'     GTMaskDate1(Index).ForeColor = vbRed
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     GTMaskDate1(Index).ForeColor = vbWindowText
'     GTMaskDate1(Index).DateValue = Null
'     GTMaskDate1(Index).SetFocus
'     Exit Sub
'  End If

'  'MH20010308 Fixed Roy's shoddy code......
'  With GTMaskDate1(Index)
'    If .Text <> "  /  /" Then
'      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or .DateValue > #12/31/9999# Then
'      If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Then
'
'        .ForeColor = vbRed
'        COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'        .ForeColor = vbWindowText
'        .DateValue = Null
'        If .Visible And .Enabled Then
'          .SetFocus
'        End If
'
'      End If
'    End If
'  End With
  
  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate GTMaskDate1(Index)

End Sub

Private Sub GTMaskDate1_Validate(Index As Integer, Cancel As Boolean)
  
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(GTMaskDate1(Index))

End Sub

Private Sub ASRUserImage1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As COA_Image
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim frmTemp As Form
  Dim bOK As Boolean
  Dim fDataChanged As Boolean

  ' JPD20030115 Fault 4862
  ' Do not action the click event if there are other
  ' recEdit forms that have not had their changes saved.
  ' Set focus to this control to ensure the menu is updated
  ' correctly when the control is right-clicked on, and to fire
  ' the changed recEdit form's 'deactivate' event.
  ' NB. This code is put in the 'click' event rather than 'onFocus'
  ' as right-click fires the 'click' event, but doesn't fire the
  ' 'onFocus' event.
  If Not mfTableEntry Then
    ASRUserImage1(Index).SetFocus
    For Each frmTemp In Forms
      If (TypeOf frmTemp Is frmRecEdit4) Then
        If Not (frmTemp Is Me) Then
          If frmTemp.Changed Then
            Exit Sub
          End If
        End If
      End If
    Next frmTemp
    Set frmTemp = Nothing
  End If
  
  If Not mfLoading Then
  
    ' Get the control's tag.
    sTag = ASRUserImage1(Index).Tag
    
    
  
  
    ' Check that the photo path is defined.
    'TM20011107 Fault 3050 - Need to check that the directory exists.
            
      
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      
      ' Is the photo embedded
      If mobjScreenControls.Item(sTag).OLEType = OLE_EMBEDDED Then
        If Not ASRUserImage1(Index).EmbeddedStream.State = adStateOpen Then
          bOK = ReadStream(ASRUserImage1(Index), mobjScreenControls.Item(sTag), False)
        End If
      End If
      
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName

      If Trim(sThisControlsColumnName) = vbNullString Then
        Exit Sub
      End If
 
      ' Check if it's a OLE control, if so load the select OLE form.
      'If mrsRecords.Fields(sThisControlsColumnName).Type = adVarChar Then
      If ASRUserImage1(Index).OLEType = OLE_UNC Then
        With frmSelectEmbedded
        
          '.Initialise mrsRecords.Fields(sThisControlsColumnName)
          .OLEType = ASRUserImage1(Index).OLEType
          .EmbeddedEnabled = mobjScreenControls.Item(sTag).EmbeddedEnabled
          .MaxOLESize = mobjScreenControls.Item(sTag).MaxOLESize
          .IsPhoto = True
          .Initialise ASRUserImage1(Index).EmbeddedStream
          .Show vbModal
        
          'JPD 20041103 Fault 8898
          ASRUserImage1(Index).Selecting = False
          
          Select Case .Selection
            Case optSelect
              ASRUserImage1(Index).EmbeddedStream = .EmbeddedFile
              Set ASRUserImage1(Index).Picture = LoadPictureFromStream(.EmbeddedFile)
              fDataChanged = True
              
              fFound = False
              For iLoop = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              If Not fFound Then
                ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
              End If
              
            Case optCancel
              fDataChanged = False
            
            Case optNone
            
              fFound = False
              For iLoop = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              If Not fFound Then
                ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
              End If
            
              ASRUserImage1(Index).EmbeddedStream = .EmbeddedFile
              fDataChanged = True

          End Select
        
          Unload frmSelectEmbedded
        
        End With
    
      Else
  
        If Len(gsPhotoPath) = 0 Or (Dir(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString) Then
          COAMsgBox "Unable to edit photo fields." & vbNewLine & _
                  "The photo path has not been defined, or is invalid." & vbNewLine & _
                  "Please set the Photo Path (non-linked) in PC Configuration." _
                   , vbExclamation + vbOKOnly, app.ProductName
        Else

          'Check if it's a photo image box, if so load the select photo form
          If mrsRecords.Fields(sThisControlsColumnName).Type = adVarChar Then
            mfControlProcessing = True
            
            Screen.MousePointer = vbHourglass
            
            ' RH 24/01/01 - Keep focus lines even though control has lost focus
            ASRUserImage1(Index).Selecting = True
            
            With frmSelectPhoto
              .Initialise ASRUserImage1(Index).ASRDataField
              Screen.MousePointer = vbDefault
              .Show vbModal
              
            ' RH 24/01/01 - Reset property as we are no longer selecting
            ASRUserImage1(Index).Selecting = False
              
              Select Case .optPhoto
                Case optSelect
                  ASRUserImage1(Index).ASRDataField = .Photo
                  Set ASRUserImage1(Index).Picture = LoadPicture(.Path & "\" & .Photo)
                  fDataChanged = True
                  
                  ' JPD20021007 Fault 4498
                  fFound = False
                  For iLoop = 1 To UBound(malngChangedOLEPhotos)
                    If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                      fFound = True
                      Exit For
                    End If
                  Next iLoop
                  If Not fFound Then
                    ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                    malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
                  End If
  
                Case optNone
                  ASRUserImage1(Index).ASRDataField = vbNullString
                  Set ASRUserImage1(Index).Picture = Nothing
                  fDataChanged = True
  
                  ' JPD20021007 Fault 4498
                  fFound = False
                  For iLoop = 1 To UBound(malngChangedOLEPhotos)
                    If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                      fFound = True
                      Exit For
                    End If
                  Next iLoop
                  If Not fFound Then
                    ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                    malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
                  End If
                                     
              End Select
              Unload frmSelectPhoto
            End With
          End If
        
          If fDataChanged Then
            For Each objControl In ASRUserImage1
              If objControl.Index <> Index Then
                sTag = objControl.Tag
                
                If Len(sTag) > 0 Then
                  lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
                
                  If lngOtherControlsColumnID = lngThisControlsColumnID Then
                    objControl.ASRDataField = ASRUserImage1(Index).ASRDataField
                    
                    If ASRUserImage1(Index).ASRDataField = vbNullString Then
                      Set objControl.Picture = Nothing
                    Else
                      Set objControl.Picture = LoadPicture(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & ASRUserImage1(Index).ASRDataField)
                    End If
                  End If
                End If
              End If
            Next objControl
            Set objControl = Nothing
          End If
        End If
      End If
    End If
  
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = fDataChanged
    End If
  
    frmMain.RefreshMainForm Me
  
  End If

  Exit Sub
Err_Trap:
  Select Case Err.Number
    ' JPD20020828 Fault 4176
    Case 52, 53, 75
      Resume Next
    Case Else
  End Select

End Sub

Private Sub ASRUserImage1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  
  If Not mfControlProcessing Then
    GotFocusCheck ASRUserImage1(Index)
    ASRUserImage1(Index).ShowSelectionMarkers = True
  End If
  mfControlProcessing = False

End Sub

Private Sub ASRUserImage1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  If Not mfControlProcessing Then
    Cancel = Not LostFocusCheck(ASRUserImage1(Index))
  End If

End Sub

Private Sub ASRWorkingPattern1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
  Dim frmTemp As Form
    
  ' JPD20030115 Fault 4862
  ' Do not action the click event if there are other
  ' recEdit forms that have not had their changes saved.
  ' Set focus to this control to ensure the menu is updated
  ' correctly when the control is right-clicked on, and to fire
  ' the changed recEdit form's 'deactivate' event.
  ' NB. This code is put in the 'click' event rather than 'onFocus'
  ' as right-click fires the 'click' event, but doesn't fire the
  ' 'onFocus' event.
  If Not mfTableEntry Then
    ASRWorkingPattern1(Index).SetFocus
    For Each frmTemp In Forms
      If (TypeOf frmTemp Is frmRecEdit4) Then
        If Not (frmTemp Is Me) Then
          If frmTemp.Changed Then
            ASRWorkingPattern1(Index).Value = mvOldValue
            Exit Sub
          End If
        End If
      End If
    Next frmTemp
    Set frmTemp = Nothing
  End If
  
  If Not mfLoading Then
    ' Get the control's tag.
    sTag = ASRWorkingPattern1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In ASRWorkingPattern1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                If .Value <> ASRWorkingPattern1(Index).Value Then
                  .Value = ASRWorkingPattern1(Index).Value
                End If
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (ASRWorkingPattern1(Index).Value <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))
      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:

End Sub

Private Sub ASRWorkingPattern1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck ASRWorkingPattern1(Index)
  
End Sub

Private Sub ASRWorkingPattern1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'GotFocus' Expression.
  LostFocusCheck ASRWorkingPattern1(Index)

End Sub

Private Sub Check1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If (Not mfLoading) Then
    ' Get the control's tag.
    sTag = Check1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In Check1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .Value = Check1(Index).Value
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (Check1(Index).Value <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))
      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True
    frmMain.RefreshMainForm Me
  End If

End Sub

Private Sub Check1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck Check1(Index)

End Sub

Private Sub Check1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(Check1(Index))

End Sub

Private Sub Combo1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If Not mfLoading Then
    ' Get the control's tag.
    sTag = Combo1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In Combo1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .ListIndex = Combo1(Index).ListIndex
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (Combo1(Index).Text <> mrsRecords(sThisControlsColumnName) & vbNullString)
      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:

End Sub

Private Sub Combo1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck Combo1(Index)

End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If Not mfLoading Then
    ' Get the control's tag.
    sTag = Combo1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In Combo1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .ListIndex = Combo1(Index).ListIndex
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (Combo1(Index).Text <> mrsRecords(sThisControlsColumnName) & vbNullString)
      frmMain.RefreshMainForm Me
    End If
  End If
  
End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(Combo1(Index))

End Sub

Private Sub Command1_Click(Index As Integer)
  Dim fOK As Boolean
  Dim sTag As String
  Dim sSQL As String
  Dim sRealSource As String
  Dim iLoop As Integer
  Dim lngLinkTableID As Long
  Dim lngLinkViewID As Long
  Dim lngLinkOrderID As Long
  Dim frmLink As frmLinkFind
  Dim objLinkTable As CTablePrivilege
  Dim objLinkView As CTablePrivilege
  Dim rsInfo As Recordset
  Dim frmTemp As Form
  Dim sTempString As String
    
  ' JPD20030115 Fault 4862
  ' Do not action the click event if there are other
  ' recEdit forms that have not had their changes saved.
  ' Set focus to this control to ensure the menu is updated
  ' correctly when the control is right-clicked on, and to fire
  ' the changed recEdit form's 'deactivate' event.
  ' NB. This code is put in the 'click' event rather than 'onFocus'
  ' as right-click fires the 'click' event, but doesn't fire the
  ' 'onFocus' event.
  If Not mfTableEntry Then
    Command1(Index).SetFocus
    For Each frmTemp In Forms
      If (TypeOf frmTemp Is frmRecEdit4) Then
        If Not (frmTemp Is Me) Then
          If frmTemp.Changed Then
            Exit Sub
          End If
        End If
      End If
    Next frmTemp
    Set frmTemp = Nothing
  End If
  
  Screen.MousePointer = vbHourglass
  
  sTag = Command1(Index).Tag
  fOK = (Len(sTag) > 0)
  
  If fOK Then
    ' Get the ID of the linked table.
    lngLinkTableID = mobjScreenControls.Item(sTag).LinkTableID
    lngLinkViewID = mobjScreenControls.Item(sTag).LinkViewID
    lngLinkOrderID = mobjScreenControls.Item(sTag).LinkOrderID
    
    Set objLinkTable = gcoTablePrivileges.FindTableID(lngLinkTableID)
    fOK = Not objLinkTable Is Nothing
  End If
  
  If fOK Then
    sRealSource = objLinkTable.RealSource
    
    ' Ensure the user can read from the link table (or views on it).
    If (Not objLinkTable.AllowSelect) And (objLinkTable.TableType = tabTopLevel) Then
      fOK = False
      
      ' The table can't be read. try the views on it.
      For Each objLinkView In gcoTablePrivileges.Collection
        If (Not objLinkView.IsTable) And _
          (objLinkView.TableID = lngLinkTableID) And _
          (objLinkView.AllowSelect) Then
          
          fOK = True
        End If
      Next objLinkView
      Set objLinkView = Nothing
    End If
  End If
    
  Set objLinkTable = Nothing
      
  If fOK Then
    Set frmLink = New frmLinkFind
    With frmLink
      If .Initialise(lngLinkTableID, lngLinkViewID, lngLinkOrderID) Then
        
        'JDM - 27/11/01 - Fault 804 - Automatically position on current link record ID
        For iLoop = 1 To UBound(malngLinkRecordIDs, 2)
          If malngLinkRecordIDs(1, iLoop) = lngLinkTableID Then
            .SetCurrentRecord (malngLinkRecordIDs(2, iLoop))
            Exit For
          End If
        Next iLoop
        
        Screen.MousePointer = vbDefault
        .Show vbModal
  
        If Not .Cancelled Then
          ' Check if the selected parent record still exists.
          ' It may have existed at the time the link find window was populated,
          ' but deleted by another user since.
          sSQL = "SELECT COUNT(id) as recCount" & _
            " FROM " & sRealSource & _
            " WHERE id = " & Trim(Str(.LinkRecordID))
          Set rsInfo = datGeneral.GetRecords(sSQL)
          fOK = (rsInfo!recCount > 0)
          rsInfo.Close
          Set rsInfo = Nothing
            
          If fOK Then
            For iLoop = 1 To UBound(malngLinkRecordIDs, 2)
              If malngLinkRecordIDs(1, iLoop) = lngLinkTableID Then
                malngLinkRecordIDs(2, iLoop) = .LinkRecordID
                Exit For
              End If
            Next iLoop
            
            ' Update any controls that refer to column in the table just linked to.
            UpdateParentControls lngLinkTableID, .LinkRecordID


            'MH20001109 Fault 981
            'After the link has been made, need to repopulate default calcs
            Dim objControl As Control
            For Each objControl In Me.Controls

              If Val(objControl.Tag) > 0 Then
                If mobjScreenControls.Item(objControl.Tag).DfltValueExprID > 0 Then
                  SetControlDefaults objControl.Tag, objControl, lngLinkTableID, .LinkRecordID
                End If
              End If

            Next
            
            mfDataChanged = True
            frmMain.RefreshMainForm Me
          Else
            COAMsgBox "The selected link record has been deleted.", vbExclamation, "Security"
          End If
        End If
      End If
    End With

    Unload frmLink
    Set frmLink = Nothing
  Else
    'NHRD12082004 Fault 8746
    sTempString = Replace(sRealSource, "_", " ")
    COAMsgBox "You do not have permission to link to the " & sTempString & " table.", vbExclamation
  End If
    
  Screen.MousePointer = vbDefault
    
End Sub

Private Sub UpdateParentControls(plngParentTableID As Long, Optional plngParentRecordID As Long)
  ' Update all screen control that are linked to fields in the given parent table
  ' with the appropriate value from the parent table.
  On Error GoTo Err_Trap
  
  Dim fOK As Boolean
  Dim fResetControl As Boolean
  Dim fFound As Boolean
  Dim fFileNameOK As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim lngParentRecordID As Long
  Dim sTag As String
  Dim sSQL As String
  Dim sColumnCode As String
  Dim sColumnName As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sRealSource As String
  Dim rsTemp As Recordset
  Dim objScreenControl As clsScreenControl
  Dim objParentTableColumnPrivileges As CColumnPrivileges
  Dim objParentViewColumnPrivileges As CColumnPrivileges
  Dim objParentTable As CTablePrivilege
  Dim objParentView As CTablePrivilege
  Dim objControl As VB.Control
  Dim fldColumn As ADODB.Field
  Dim asViews() As String
  Dim asJoinViews() As String
  Dim sFormat As String
  
  fOK = True
  
  ' Get the associated parent record ID.
  If plngParentRecordID > 0 Then
    lngParentRecordID = plngParentRecordID
  Else
    lngParentRecordID = IIf(IsNull(mrsRecords.Fields("ID_" & Trim(Str(plngParentTableID)))), _
      0, mrsRecords.Fields("ID_" & Trim(Str(plngParentTableID))))
  End If

  If lngParentRecordID > 0 Then
    ' Get the list of parent columns referred to by controls in the screen.
    sColumnList = ""
    sJoinCode = ""
    ReDim asJoinViews(0)
    
    ' Get the table object.
    Set objParentTable = gcoTablePrivileges.FindTableID(plngParentTableID)
    
    If Not objParentTable Is Nothing Then
      ' Get the table's set of column privileges.
      Set objParentTableColumnPrivileges = GetColumnPrivileges(objParentTable.TableName)

      For Each objScreenControl In mobjScreenControls.Collection
        If (objScreenControl.ColumnID > 0) And _
         (objScreenControl.TableID = plngParentTableID) Then
          ' The current control DOES refer to a column in the parent table.
          ' Check if the user has 'read' permission on the column directly from the table.
          If objParentTableColumnPrivileges.Item(objScreenControl.ColumnName).AllowSelect Then
            ' The column can be read directly from the table so add the column to the column list.
            sColumnList = sColumnList & _
              IIf(Len(sColumnList) > 0, ", ", "") & _
              objParentTable.RealSource & "." & Trim(objScreenControl.ColumnName)
          Else
            ' The column cannot be read from the parent table.
            ' Try to read it from the views on the parent table.
            ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
            ReDim asViews(0)

            For Each objParentView In gcoTablePrivileges.Collection
              If (Not objParentView.IsTable) And _
                (objParentView.TableID = plngParentTableID) And _
                (objParentView.AllowSelect) Then

                ' Get the column permission for the view.
                Set objParentViewColumnPrivileges = GetColumnPrivileges(objParentView.ViewName)

                If objParentViewColumnPrivileges.IsValid(objScreenControl.ColumnName) Then
                  If objParentViewColumnPrivileges.Item(objScreenControl.ColumnName).AllowSelect Then
                    ' The column can be seen through the view.
                    ' Add the view info to an array to be put into the column list or order code below.
                    iNextIndex = UBound(asViews) + 1
                    ReDim Preserve asViews(iNextIndex)
                    asViews(iNextIndex) = objParentView.ViewName
                  End If
                End If
                Set objParentViewColumnPrivileges = Nothing

              End If
            Next objParentView
            Set objParentView = Nothing
          
            sColumnCode = ""
            For iNextIndex = 1 To UBound(asViews)
              ' Add the view to the list of join views if its not already there.
              fFound = False
              For iLoop = 1 To UBound(asJoinViews)
                If UCase(Trim(asViews(iNextIndex))) = UCase(Trim(asJoinViews(iLoop))) Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              
              If Not fFound Then
                iLoop = UBound(asJoinViews) + 1
                ReDim Preserve asJoinViews(iLoop)
                asJoinViews(iLoop) = asViews(iNextIndex)
              End If
              
              If iNextIndex = 1 Then
                sColumnCode = "CASE "
              End If
                
              sColumnCode = sColumnCode & _
                " WHEN NOT " & asViews(iNextIndex) & "." & objScreenControl.ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & objScreenControl.ColumnName
            Next iNextIndex
          
            If Len(sColumnCode) > 0 Then
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END AS " & objScreenControl.ColumnName
              
              sColumnList = sColumnList & _
                IIf(Len(sColumnList) > 0, ", ", "") & _
                sColumnCode
            End If
          End If
        End If
      Next objScreenControl
      Set objScreenControl = Nothing
        
      If Len(sColumnList) > 0 Then
        'MH20030210 Fault 5037
        'If top level then the select needs to hinge off the table itself...
        'If objParentTable.AllowSelect Then
        If objParentTable.AllowSelect Or _
           objParentTable.TableType = tabTopLevel Then
          sJoinCode = ""
          For iNextIndex = 1 To UBound(asJoinViews)
            sJoinCode = sJoinCode & _
              " LEFT OUTER JOIN " & asJoinViews(iNextIndex) & _
              " ON " & objParentTable.RealSource & ".ID = " & asJoinViews(iNextIndex) & ".ID"
          Next iNextIndex
          
          sSQL = "SELECT " & sColumnList & _
          " FROM " & objParentTable.RealSource & _
          " " & sJoinCode & _
          " WHERE " & objParentTable.RealSource & ".ID = " & Trim(Str(lngParentRecordID))
        Else
          sJoinCode = ""
          For iNextIndex = 2 To UBound(asJoinViews)
            sJoinCode = sJoinCode & _
              " FULL OUTER JOIN " & asJoinViews(iNextIndex) & _
              " ON " & asJoinViews(1) & ".ID = " & asJoinViews(iNextIndex) & ".ID"
          Next iNextIndex
          
          sSQL = "SELECT " & sColumnList & _
          " FROM " & asJoinViews(1) & _
          " " & sJoinCode & _
          " WHERE " & asJoinViews(1) & ".ID = " & Trim(Str(lngParentRecordID))
        End If
        Set rsTemp = datGeneral.GetRecords(sSQL)
      End If
    End If
  End If
  
  ' Update the controls with the data.
  mfLoading = True
  For Each objControl In Me.Controls
    sTag = objControl.Tag

    'JPD 20030610
    If TypeOf objControl Is ActiveBar Then
      sTag = ""
    End If
      
    If Len(sTag) > 0 Then
      If (mobjScreenControls.Item(sTag).ColumnID > 0) And _
        (mobjScreenControls.Item(sTag).TableID = plngParentTableID) Then
        
        ' The current control is linked to the given parent table.
        sColumnName = mobjScreenControls.Item(sTag).ColumnName
        
        ' Check if the control's associated column has been read.
        fResetControl = (rsTemp Is Nothing)
        
        If Not fResetControl Then
          fResetControl = (rsTemp.BOF And rsTemp.EOF)
        End If
        
        If Not fResetControl Then
          fFound = False
          For Each fldColumn In rsTemp.Fields
            If UCase(Trim(fldColumn.Name)) = UCase(Trim(sColumnName)) Then
              fFound = True
              Exit For
            End If
          Next fldColumn
          Set fldColumn = Nothing
          
          If Not fFound Then
            ' The control's column is not in the recordset so reset the control.
            fResetControl = True
          End If
        End If
               
        ' Update the control with the value from the parent table.
        With objControl
          If TypeOf objControl Is TDBText6Ctl.TDBText Then
            If fResetControl Then
              .Text = ""
            Else
              .Text = RTrim(rsTemp(sColumnName) & vbNullString)
            End If

          ElseIf TypeOf objControl Is COA_Image Then
            If fResetControl Then
              .Picture = Nothing
              .ASRDataField = ""
            Else
              If rsTemp(sColumnName).Type = adVarChar Then
                fFileNameOK = False
  
                If Not IsNull(rsTemp(sColumnName)) Then
                  If Len(rsTemp(sColumnName)) > 0 Then
                    fFileNameOK = True
                    .Picture = LoadPicture(gsPhotoPath & "\" & rsTemp(sColumnName))
                    .ASRDataField = rsTemp(sColumnName)
                  End If
                End If
  
                If Not fFileNameOK Then
                  .Picture = Nothing
                  .ASRDataField = ""
                End If
              End If
            End If
            
          ElseIf TypeOf objControl Is TDBMask6Ctl.TDBMask Then
            If fResetControl Then
              .Text = ""
            Else
              .Text = RTrim(rsTemp(sColumnName) & vbNullString)
            End If
            
          ElseIf TypeOf objControl Is TextBox Then
            If fResetControl Then
              .Text = ""
            Else
              .Text = RTrim(rsTemp(sColumnName) & vbNullString)
            End If
            
          'JPD 20050302 Fault 9847
          ElseIf (TypeOf objControl Is TDBNumberCtrl.TDBNumber) Or _
            (TypeOf objControl Is TDBNumber6Ctl.TDBNumber) Then
            'TM20060831 Fault ???? - This ClearControl method didn't appear to be doing it's job.
            ' So I called the Ghostbusters and they zapped it by setting the .Value property to "Slimer".
            ' Who ya gonna call???
            If fResetControl Then
              .ClearControl
              .Value = ""   ' GHOSTBUSTERS!!!
            Else
              .Value = rsTemp(sColumnName)
            End If
          
          ElseIf TypeOf objControl Is XtremeSuiteControls.CheckBox Then
            If fResetControl Then
              .Value = 0
            Else
              .Value = IIf(rsTemp(sColumnName), 1, 0)
            End If
            
          ElseIf TypeOf objControl Is XtremeSuiteControls.ComboBox Then
            .Clear
            If fResetControl Then
              If .ListCount > 0 Then
                .ListIndex = 0
              End If
            Else
              ' JPD20021112 Fault 4736
              '.ListIndex = UI.cboSelect(.hWnd, RTrim(rsTemp(sColumnName) & vbNullString))
              'JPD 20030916 Fault 6978
              .AddItem IIf(IsNull(rsTemp(sColumnName)), "", RTrim(rsTemp(sColumnName)))
              .ListIndex = 0
            End If
            
          ElseIf TypeOf objControl Is COA_Lookup Then
            If fResetControl Then
              .Text = ""
            Else
              'JPD 20050810 Fault 10165
              If (mobjScreenControls.Item(objControl.Tag).DataType = sqlNumeric) Then
                sFormat = "0"
                If mobjScreenControls.Item(objControl.Tag).Use1000Separator Then
                  sFormat = "#,0"
                End If
                If mobjScreenControls.Item(objControl.Tag).Decimals > 0 Then
                  sFormat = sFormat & "." & String(mobjScreenControls.Item(objControl.Tag).Decimals, "0")
                End If
                            
                .Text = IIf(IsNull(rsTemp(sColumnName)), "", Format(rsTemp(sColumnName), sFormat))
              Else
                .Text = IIf(IsNull(rsTemp(sColumnName)), "", rsTemp(sColumnName))
              End If
            End If
            
          ElseIf TypeOf objControl Is COA_OptionGroup Then
            If fResetControl Then
              .Text = ""
            Else
              .Text = RTrim(rsTemp(sColumnName) & vbNullString)
            End If
            
          ElseIf TypeOf objControl Is OLE Then
            If fResetControl Then
              .SourceDoc = vbNullString
            Else
              If rsTemp(sColumnName).Type = adVarChar Then
                fFileNameOK = False
                
                If Not IsNull(rsTemp(sColumnName)) Then
                  If Len(rsTemp(sColumnName)) > 0 Then
                    fFileNameOK = True
                    .CreateLink gsOLEPath & "\" & rsTemp(sColumnName)
                  End If
                End If
  
                If Not fFileNameOK Then
                  .SourceDoc = vbNullString
                End If
              End If
            End If
            
          ElseIf TypeOf objControl Is COA_Spinner Then
            If fResetControl Then
              .Text = ""
            Else
              .Text = rsTemp(sColumnName) & vbNullString
            End If
            
          ElseIf TypeOf objControl Is GTMaskDate.GTMaskDate Then
            If fResetControl Then
              .Text = ""
            Else
              If IsNull(rsTemp(sColumnName)) Then
                .Text = ""
              Else
                .Text = Format(DateValue(rsTemp(sColumnName)), DateFormat)
              End If
             End If
             
          ElseIf TypeOf objControl Is COA_WorkingPattern Then
            If fResetControl Then
              .Value = ""
            Else
              .Value = rsTemp(sColumnName) & vbNullString
            End If
              
          ElseIf TypeOf objControl Is COA_Navigation Then
            If fResetControl Then
              .NavigateTo = ""
            Else
              .NavigateTo = rsTemp(sColumnName) & vbNullString
            End If

          ElseIf TypeOf objControl Is COA_ColourSelector Then
            If fResetControl Then
              .BackColor = vbWhite
            Else
              .BackColor = Val(rsTemp(sColumnName))
            End If
          
          End If
        End With
      End If
    End If
  Next objControl
  Set objControl = Nothing
              
  mfLoading = False
        
  If Not rsTemp Is Nothing Then
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  Exit Sub

Err_Trap:
  Select Case Err.Number
    Case 52, 53, 75, 76
      fFileNameOK = False
      Resume Next
  End Select
End Sub

Private Sub ctlNewLookup1_Change(Index As Integer)
  
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If Not mfLoading Then
    ' Get the control's tag.
    sTag = ctlNewLookup1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In ctlNewLookup1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                If .Text <> ctlNewLookup1(Index).Text Then
                  .Text = ctlNewLookup1(Index).Text
                End If
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (ctlNewLookup1(Index).Text <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), 0, mrsRecords(sThisControlsColumnName)))
      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:

End Sub

Private Sub ctlNewLookup1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  RePopulateLookupControl ctlNewLookup1(Index)
    
  Exit Sub
    
Err_Trap:
  COAMsgBox Err.Description & " - ctlNewLookup1_Click", vbCritical

End Sub

Private Sub RePopulateLookupControl(pctlLookupControl As COA_Lookup)
  ' Repopulate the lookup control with the current set of lookup column values.
  Dim iLoop As Integer
  Dim sTag As String
  
  'TM23092004 Fault 9204
  Dim sTag2 As String
  
  Dim sSQL As String
  Dim sLookupColumnName As String
  Dim sSource As String
  Dim rsLookup As ADODB.Recordset
  Dim lngLookUpTableID As Long
  Dim lngLookUpColumnID As Long
  Dim sLookUpColumns As String
  Dim asLookupData() As Variant
  Dim objLookupTable As CTablePrivilege
  Dim objLookupColumns As CColumnPrivileges
  Dim objLookupColumn As CColumnPrivilege
  Dim objLookupView As CTablePrivilege
  Dim fOK As Boolean
  Dim frmLink As frmLinkFind
  
  Dim objColumnPrivilege As DataMgr.CColumnPrivilege
  Dim objControl As Object
  Dim lngLookUpFilterColumnID As Long
  Dim iLookUpFilterOperator As FilterOperators
  Dim lngLookupFilterValueID As Long
  Dim sLookupFilterColumnName As String
  Dim iLookupFilterColumnType As Integer
  Dim sLookupFilterColumnValue As String
  Dim vLookupFilterColumnValue As Variant
  Dim intExtraColumnsAdded As Integer
  Dim sLookupFilterCode As String
  Dim sModifiedFilterValue As String
  Dim fFound As Boolean
  Dim fControlFound As Boolean
  Dim sFormat As String
  Dim sDateFormat As String
  
  Dim aiColumnInfo() As Integer
  Dim lngNextIndex As Long
  
  Set rsLookup = New ADODB.Recordset
  
  ReDim aiColumnInfo(2, 0)
  ' Column 0 = data type
  ' Column 1 = 1000 separator
  ' Column 2 = decimals
  
  ' Get the given control's tag.
  sTag = pctlLookupControl.Tag
  sDateFormat = DateFormat
  
  If Len(sTag) > 0 Then
    ' Get the table and column id of the column to display in the lookup ctl.
    lngLookUpTableID = mobjScreenControls.Item(sTag).LookupTableID
    lngLookUpColumnID = mobjScreenControls.Item(sTag).LookupColumnID
    
    ' Get the lookup table object and the collection of columns.
    Set objLookupTable = gcoTablePrivileges.FindTableID(lngLookUpTableID)
    Set objLookupColumns = GetColumnPrivileges(objLookupTable.TableName)
    
    lngLookUpFilterColumnID = mobjScreenControls.Item(sTag).LookupFilterColumnID
    iLookUpFilterOperator = mobjScreenControls.Item(sTag).LookupFilterOperator
    lngLookupFilterValueID = mobjScreenControls.Item(sTag).LookupFilterValueID
    sLookupFilterColumnValue = ""
    
    ' Get the lookup column name.
    For Each objLookupColumn In objLookupColumns
      If objLookupColumn.ColumnID = lngLookUpColumnID Then
        sLookupColumnName = objLookupColumn.ColumnName
        sLookUpColumns = sLookupColumnName
        
        
        ReDim Preserve aiColumnInfo(2, UBound(aiColumnInfo, 2) + 1)
        aiColumnInfo(0, UBound(aiColumnInfo, 2)) = objLookupColumn.DataType
        aiColumnInfo(1, UBound(aiColumnInfo, 2)) = IIf(objLookupColumn.UseThousandSeparator, 1, 0)
        aiColumnInfo(2, UBound(aiColumnInfo, 2)) = objLookupColumn.Decimals
        
        Exit For
      End If
    Next objLookupColumn
    Set objLookupColumn = Nothing
    
    If lngLookUpFilterColumnID > 0 Then
      ' Get the lookup filter column name.
      For Each objLookupColumn In objLookupColumns
        If objLookupColumn.ColumnID = lngLookUpFilterColumnID Then
          sLookupFilterColumnName = objLookupColumn.ColumnName
          iLookupFilterColumnType = objLookupColumn.DataType
          Exit For
        End If
      Next objLookupColumn
      Set objLookupColumn = Nothing

      ' Get the value type
      fFound = False
      For Each objColumnPrivilege In mcolColumnPrivileges
        If objColumnPrivilege.ColumnID = lngLookupFilterValueID Then
          If objColumnPrivilege.AllowSelect Then
            fFound = True
            
            Select Case iLookupFilterColumnType
              Case sqlBoolean
                sLookupFilterColumnValue = IIf(mrsRecords.Fields(objColumnPrivilege.ColumnName).Value, "1", "0")
              Case Else
                sLookupFilterColumnValue = IIf(IsNull(mrsRecords.Fields(objColumnPrivilege.ColumnName).Value), "", mrsRecords.Fields(objColumnPrivilege.ColumnName).Value)
            End Select
            Exit For
          End If
        End If
      Next objColumnPrivilege

      If Not fFound Then
        lngLookUpFilterColumnID = 0
        iLookUpFilterOperator = 0
        lngLookupFilterValueID = 0
        sLookupFilterColumnValue = vbNullString
        sLookupFilterColumnName = vbNullString
        iLookupFilterColumnType = 0
        
        COAMsgBox "You do not have 'read' permission on the lookup filter value column. No filter will be applied.", vbOKOnly, app.ProductName
      Else
        ' Update the value if it's painted onto current screen
        fControlFound = False
        
        For Each objControl In Me.Controls
          With objControl
            ' get the control's tag.
            sTag2 = .Tag
      
            'JPD 20030610
            If TypeOf objControl Is ActiveBar Then
              sTag2 = ""
            End If
            
            ' Check if the control is tagged to a column.
            If Len(sTag2) > 0 Then
              If mobjScreenControls.Item(sTag2).ColumnID = lngLookupFilterValueID Then
                fControlFound = True
                
                Select Case iLookupFilterColumnType
                  'JPD 20040122 Fault 7963
                  'Case sqlBoolean
                  Case sqlBoolean, sqlLongVarChar
                    sLookupFilterColumnValue = objControl.Value
                  Case Else
                    sLookupFilterColumnValue = objControl.Text
                End Select
              End If
            End If
          End With
        Next objControl
  
        'JPD 20041004 Fault 9255
        ' If we're dealing with a 'new'record, AND the column is not represented on the screen,
        ' then get the defined default value.
        If (Not fControlFound) _
          And (mrsRecords.EditMode = adEditAdd) Then
        
          sLookupFilterColumnValue = GetColumnDefault(lngLookupFilterValueID)
        End If

        Select Case iLookupFilterColumnType
          Case sqlBoolean
            If Len(sLookupFilterColumnValue) = 0 Then
              sLookupFilterCode = vbTab & " IS NOT NULL"
            Else
              sLookupFilterCode = vbTab & " = " & sLookupFilterColumnValue
            End If
          
          Case sqlNumeric, sqlInteger
            If Len(Trim(sLookupFilterColumnValue)) = 0 Then
              sLookupFilterColumnValue = "0"
            End If
            
            Select Case iLookUpFilterOperator
              Case giFILTEROP_EQUALS
                sLookupFilterCode = vbTab & " = " & CStr(sLookupFilterColumnValue)
                If Val(sLookupFilterColumnValue) = 0 Then
                  sLookupFilterCode = sLookupFilterCode & " OR " & vbTab & " IS NULL"
                End If
              
              Case giFILTEROP_NOTEQUALTO
                sLookupFilterCode = vbTab & " <> " & Str(sLookupFilterColumnValue)
                If Val(sLookupFilterColumnValue) = 0 Then
                  sLookupFilterCode = sLookupFilterCode & " AND " & vbTab & " IS NOT NULL"
                End If
          
              Case giFILTEROP_ISATMOST
                sLookupFilterCode = vbTab & " <= " & Str(sLookupFilterColumnValue)
                If Val(sLookupFilterColumnValue) >= 0 Then
                  sLookupFilterCode = sLookupFilterCode & " OR " & vbTab & " IS NULL"
                End If
          
              Case giFILTEROP_ISATLEAST
                sLookupFilterCode = vbTab & " >= " & Str(sLookupFilterColumnValue)
                If Val(sLookupFilterColumnValue) <= 0 Then
                  sLookupFilterCode = sLookupFilterCode & " OR " & vbTab & " IS NULL"
                End If
          
              Case giFILTEROP_ISMORETHAN
                sLookupFilterCode = vbTab & " > " & Str(sLookupFilterColumnValue)
                If Val(sLookupFilterColumnValue) < 0 Then
                  sLookupFilterCode = sLookupFilterCode & " OR " & vbTab & " IS NULL"
                End If
          
              Case giFILTEROP_ISLESSTHAN
                sLookupFilterCode = vbTab & " < " & Str(sLookupFilterColumnValue)
                If Val(sLookupFilterColumnValue) > 0 Then
                  sLookupFilterCode = sLookupFilterCode & " OR " & vbTab & " IS NULL"
                End If
            End Select
          
          Case sqlDate
            Select Case iLookUpFilterOperator
              Case giFILTEROP_ON
                If IsDate(sLookupFilterColumnValue) Then
                  sLookupFilterCode = vbTab & " = CONVERT(datetime, '" & Replace(Format(sLookupFilterColumnValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "')"
                Else
                  sLookupFilterCode = vbTab & " is null"
                End If
  
              Case giFILTEROP_NOTON
                If IsDate(sLookupFilterColumnValue) Then
                  sLookupFilterCode = vbTab & " <> CONVERT(datetime, '" & Replace(Format(sLookupFilterColumnValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "') OR " & vbTab & " is null"
                Else
                  sLookupFilterCode = vbTab & " is not null"
                End If
  
              Case giFILTEROP_ONORBEFORE
                If IsDate(sLookupFilterColumnValue) Then
                  sLookupFilterCode = vbTab & " <= CONVERT(datetime, '" & Replace(Format(sLookupFilterColumnValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "') OR " & vbTab & " is null"
                Else
                  sLookupFilterCode = vbTab & " is null"
                End If
  
              Case giFILTEROP_ONORAFTER
                If IsDate(sLookupFilterColumnValue) Then
                  sLookupFilterCode = vbTab & " >= CONVERT(datetime, '" & Replace(Format(sLookupFilterColumnValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "')"
                Else
                  sLookupFilterCode = vbTab & " is null OR " & vbTab & " IS NOT NULL"
                End If
  
              Case giFILTEROP_AFTER
                If IsDate(sLookupFilterColumnValue) Then
                  sLookupFilterCode = vbTab & " > CONVERT(datetime, '" & Replace(Format(sLookupFilterColumnValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "')"
                Else
                  sLookupFilterCode = vbTab & " IS NOT NULL"
                End If
  
              Case giFILTEROP_BEFORE
                If IsDate(sLookupFilterColumnValue) Then
                  sLookupFilterCode = vbTab & " < CONVERT(datetime, '" & Replace(Format(sLookupFilterColumnValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "') OR " & vbTab & " IS NULL"
                Else
                  sLookupFilterCode = vbTab & " IS NULL AND " & vbTab & " IS NOT NULL"
                End If
            End Select
          
          Case sqlVarChar, sqlVarBinary, sqlLongVarChar
            Select Case iLookUpFilterOperator
              Case giFILTEROP_IS
                If Len(sLookupFilterColumnValue) = 0 Then
                  sLookupFilterCode = vbTab & " = '' OR " & vbTab & " IS NULL"
                Else
                  sLookupFilterCode = vbTab & " = '" & Replace(sLookupFilterColumnValue, "'", "''") & "'"
                End If
  
              Case giFILTEROP_ISNOT
                If Len(sLookupFilterColumnValue) = 0 Then
                  sLookupFilterCode = vbTab & " <> '' AND " & vbTab & " IS NOT NULL"
                Else
                  sLookupFilterCode = vbTab & " <> '" & Replace(sLookupFilterColumnValue, "'", "''") & "'"
                End If
  
              Case giFILTEROP_CONTAINS
                If Len(sLookupFilterColumnValue) = 0 Then
                  sLookupFilterCode = vbTab & " IS NULL OR " & vbTab & " IS NOT NULL"
                Else
                  sLookupFilterCode = vbTab & " LIKE '%" & Replace(sLookupFilterColumnValue, "'", "''") & "%'"
                End If
  
              Case giFILTEROP_DOESNOTCONTAIN
                If Len(sLookupFilterColumnValue) = 0 Then
                  sLookupFilterCode = vbTab & " IS NULL AND " & vbTab & " IS NOT NULL"
                Else
                  sLookupFilterCode = vbTab & " NOT LIKE '%" & Replace(sLookupFilterColumnValue, "'", "''") & "%'"
                End If
            End Select
          Case Else
        End Select
      End If
    End If
    
    'JPD 20031217 Islington changes
    If objLookupTable.TableType = tabLookup Then
      ' Get the list of columns in the default order for the lookup table.
      sSQL = "SELECT ASRSysColumns.columnName, ASRSysColumns.columnID, ASRSysColumns.dataType, ASRSysColumns.use1000Separator, ASRSysColumns.decimals" & _
        " FROM ASRSysOrderItems" & _
        " INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID" & _
        " WHERE orderID = " & Trim(Str(objLookupTable.DefaultOrderID)) & _
        " AND type = 'F'" & _
        " ORDER BY sequence"
    
      Set rsLookup = New ADODB.Recordset
      rsLookup.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

      Do Until rsLookup.EOF
        If rsLookup!ColumnID <> lngLookUpColumnID Then
          sLookUpColumns = sLookUpColumns & ", " & rsLookup!ColumnName
        
          ReDim Preserve aiColumnInfo(2, UBound(aiColumnInfo, 2) + 1)
          aiColumnInfo(0, UBound(aiColumnInfo, 2)) = rsLookup!DataType
          aiColumnInfo(1, UBound(aiColumnInfo, 2)) = IIf(rsLookup!Use1000Separator, 1, 0)
          aiColumnInfo(2, UBound(aiColumnInfo, 2)) = rsLookup!Decimals
        End If
        rsLookup.MoveNext
      Loop
      rsLookup.Close
      Set rsLookup = Nothing
        
      ' Get the lookup table values.
      sSource = "SELECT DISTINCT " & sLookUpColumns _
        & IIf(Len(sLookupFilterColumnName) > 0, "," & sLookupFilterColumnName & " '?ID_FILTERLOOKUPCOLUMN'", "") _
        & " FROM " & objLookupTable.RealSource _
        & IIf(Len(sLookupFilterColumnName) > 0, " WHERE " & Replace(sLookupFilterCode, vbTab, sLookupFilterColumnName), "") _
        & " ORDER BY " & sLookupColumnName
  
      intExtraColumnsAdded = IIf(Len(sLookupFilterColumnName) > 0, 1, 0)
  
      ' Load the lookup control array with the values
      Set rsLookup = New ADODB.Recordset
      rsLookup.Open sSource, gADOCon, adOpenForwardOnly, adLockReadOnly
      ReDim Preserve asLookupData(rsLookup.Fields.Count, 1)
        
      For iLoop = 1 To (rsLookup.Fields.Count - intExtraColumnsAdded)
        asLookupData(iLoop, 1) = RemoveUnderScores(rsLookup.Fields(iLoop - 1).Name)
      Next iLoop
      
      lngNextIndex = 1
      Do Until rsLookup.EOF
      
        lngNextIndex = lngNextIndex + 1
        If lngNextIndex > UBound(asLookupData, 2) Then ReDim Preserve asLookupData(rsLookup.Fields.Count, lngNextIndex + 50)

        'ReDim Preserve asLookupData(rsLookup.Fields.Count, UBound(asLookupData, 2) + 1)
        For iLoop = 1 To (rsLookup.Fields.Count - intExtraColumnsAdded)
          
          ' RH 13/10 FORMAT DATES ACCORDING TO REGIONAL SETTINGS
          ' JPD 9/3/2001 (What year was your fix made in Roy, eh, eh ?)
          ' Use the field type value before the IsData function to avoid formatting sort codes as dates.
          If rsLookup.Fields(iLoop - 1).Type = adDBTimeStamp Then
            If IsDate(rsLookup.Fields(iLoop - 1)) Then
              asLookupData(iLoop, lngNextIndex) = IIf(IsNull(rsLookup.Fields(iLoop - 1).Value), "", Format(DateValue(rsLookup.Fields(iLoop - 1).Value), sDateFormat))
            Else
              asLookupData(iLoop, lngNextIndex) = IIf(IsNull(rsLookup.Fields(iLoop - 1).Value), "", rsLookup.Fields(iLoop - 1).Value)
            End If
          Else
            'JPD 20050810 Fault 10165
            'If (mobjScreenControls.Item(pctlLookupControl.Tag).DataType = sqlNumeric) Then
            If (rsLookup.Fields(iLoop - 1).Type = adNumeric) Then
              sFormat = "0"
              'If mobjScreenControls.Item(pctlLookupControl.Tag).Use1000Separator Then
              If aiColumnInfo(1, iLoop) = 1 Then
                sFormat = "#,0"
              End If
              'If mobjScreenControls.Item(pctlLookupControl.Tag).Decimals > 0 Then
              If aiColumnInfo(2, iLoop) > 0 Then
                sFormat = sFormat & "." & String(aiColumnInfo(2, iLoop), "0")
              End If
                          
              asLookupData(iLoop, lngNextIndex) = IIf(IsNull(rsLookup.Fields(iLoop - 1).Value), "", Format(rsLookup.Fields(iLoop - 1).Value, sFormat))
            Else
              asLookupData(iLoop, lngNextIndex) = IIf(IsNull(rsLookup.Fields(iLoop - 1).Value), "", rsLookup.Fields(iLoop - 1).Value)
            End If
          End If
        Next iLoop
        
        rsLookup.MoveNext
      Loop
    
      ' Size the array to the exact size
      ReDim Preserve asLookupData(rsLookup.Fields.Count, lngNextIndex)
    
      rsLookup.Close
      Set rsLookup = Nothing
    
      With pctlLookupControl
        .Clear
        If UBound(asLookupData, 2) > 0 Then
          .PassArray asLookupData()
        End If
        
        .Mandatory = mobjScreenControls.Item(sTag).Mandatory
        .AllowInsert = objLookupTable.AllowInsert And Not objLookupTable.HideFromMenu
      End With
      
      pctlLookupControl.AllowSelect = True
    Else
      ' Non-lookup table referred to by a lookup column.

      ' Check that the column can be read.
      fOK = False
      If (objLookupTable.AllowSelect) Then
        Set objLookupColumns = GetColumnPrivileges(objLookupTable.TableName)
        If objLookupColumns.IsValid(sLookupColumnName) Then
          If objLookupColumns(sLookupColumnName).AllowSelect Then
            fOK = True
          End If
        End If
        Set objLookupColumns = Nothing
      End If
      
      If (Not fOK) And (objLookupTable.TableType = tabTopLevel) Then
        ' The table can't be read. try the views on it.
        For Each objLookupView In gcoTablePrivileges.Collection
          If (Not objLookupView.IsTable) And _
            (objLookupView.TableID = lngLookUpTableID) And _
            (objLookupView.AllowSelect) Then
            
            Set objLookupColumns = GetColumnPrivileges(objLookupView.ViewName)
            If objLookupColumns.IsValid(sLookupColumnName) Then
              If objLookupColumns(sLookupColumnName).AllowSelect Then
                fOK = True
              End If
            End If
            Set objLookupColumns = Nothing
          End If
        Next objLookupView
        Set objLookupView = Nothing
      End If
    
      If Not fOK Then
        COAMsgBox "Unable to load the lookup values. You do not have 'read' permission on the '" & sLookupColumnName & "' column in the '" & objLookupTable.TableName & "' table.", vbOKOnly, app.ProductName
      Else
        ' User does have some read permission on the table (or views on it).
        Set frmLink = New frmLinkFind
        With frmLink
          If .Initialise(lngLookUpTableID, 0, 0, lngLookUpColumnID, sLookupFilterCode, lngLookUpFilterColumnID) Then
            .SetCurrentRecordValue pctlLookupControl.Text

            Screen.MousePointer = vbDefault
            .Show vbModal

            If Not .Cancelled Then
              pctlLookupControl.Text = .LookupValue

              mfDataChanged = True
              frmMain.RefreshMainForm Me
            End If
          End If
        End With

        Unload frmLink
        Set frmLink = Nothing
      End If
    End If
  
    Set objLookupTable = Nothing
    Set objLookupColumns = Nothing
  End If
  
End Sub

Private Sub ctlNewLookup1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck ctlNewLookup1(Index)

End Sub

Private Sub ctlNewLookup1_NewEntry(Index As Integer)
  ' Call the lookup tables record edit screen.
  Dim fOK As Boolean
  Dim sTag As String
  Dim sLookupColumnName As String
  Dim sLookupTableName As String
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivilege As CColumnPrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim objLUValue As CLookupValue
  
  'MH20010518 Fault 2227 - Fixes problem with nested lookups !
  'Don't clear out the collection every time !
  '' Clear the lookup collection (if it exists).
  ''If Not gcoLookupValues Is Nothing Then
  ''  Set gcoLookupValues = Nothing
  ''End If
  ''Set gcoLookupValues = New CLookupValues
  If gcoLookupValues Is Nothing Then
    Set gcoLookupValues = New CLookupValues
  End If
  
  ' Get the column name.
  sTag = ctlNewLookup1(Index).Tag
  
  If Len(sTag) > 0 Then
    ' Get the lookup table name.
    fOK = False
    For Each objTableView In gcoTablePrivileges.Collection
      If objTableView.TableID = mobjScreenControls.Item(sTag).LookupTableID Then
      'If Me.TableID = mobjScreenControls.Item(sTag).LookupTableID Then
        sLookupTableName = objTableView.TableName
        
        ' Get the lookup tables column collection.
        Set objColumnPrivileges = GetColumnPrivileges(sLookupTableName)
        
        ' Get the lookup column's name.
        For Each objColumnPrivilege In objColumnPrivileges
          If objColumnPrivilege.ColumnID = mobjScreenControls.Item(sTag).LookupColumnID Then
            fOK = True
            sLookupColumnName = objColumnPrivilege.ColumnName
            Exit For
          End If
        Next objColumnPrivilege
        Set objColumnPrivilege = Nothing
        
        Exit For
      End If
    Next objTableView
    Set objTableView = Nothing
    
    If fOK Then
      mfLookup = True
      
      'JPD20020124 Fault 3389
      ' Remove the ClookupValue from the ClookupValues collection if it already exists.
      ' For this fault I've changed the way lookup recEdit screens are called up and handled.
      ' So the change TM made for Fault 2534 has been removed, but his comment has been left in
      ' for tracing purposes.
      ''''TM20010723 Fault 2534
      '''' Only add a new ClookupValue to the ClookupValues collection if it
      '''' does not already exist.
      If gcoLookupValues.Count > 0 Then
        If gcoLookupValues.IsValid(Me.hWnd) Then
          gcoLookupValues.Remove CStr(Me.hWnd)
        End If
      End If
      ' Add the ClookupValue to the ClookupValues collection.
      Set objLUValue = gcoLookupValues.Add(sLookupColumnName, Me.hWnd, 0, mobjScreenControls.Item(ctlNewLookup1(Index).Tag).ColumnID)
  
      ' Disable the current record edit screen, and display the lookup table screen.
      DisableMe
      AddNewTableEntry mobjScreenControls.Item(sTag).LookupTableID
      
      ' Set the child Hwnd value in the ClookupValue object.
      objLUValue.ChildHwnd = frmMain.ActiveForm.hWnd
    End If
  End If

End Sub


Private Sub DisableMe()
  ' Disable the record edit screen and its find/summary window.
  Me.Enabled = False
  
  If Not mfrmFind Is Nothing Then
    mfrmFind.Enabled = False
  End If
  
  If Not mfrmSummary Is Nothing Then
    mfrmSummary.Enabled = False
  End If

End Sub
Private Sub ctlNewLookup1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(ctlNewLookup1(Index))

End Sub


Private Sub Form_Activate()
  ' Check if we are loading the form, if we are then set the loading
  ' property to false so that when activate is called again
  ' it initialises properly.
  Dim objControl As Control
  Dim sTag As String
  
  ' JPD20021209 Fault 4863
  If Not SaveAscendants Then
    Exit Sub
  End If

  DoEvents
  
  If mrsRecords Is Nothing Then
    Exit Sub
  End If
  
  If mrsRecords.State = adStateClosed Then
    Exit Sub
  End If
  
  If mfLookup Then
    mfLookup = False
  End If
  
  If mfLoading Then
    mfLoading = False
  Else
    ' Check for there being no records
    With mrsRecords
      If .EOF And .BOF Then
        AddNew
      End If
    End With
  End If
  
  '#############
  ' Maybe put something in here to reset the activecontrol to be the first control in the
  ' tab order. this is needed if a history recedit scrn is open, focus put onto a control,
  ' the recedit scrn closed, another record selected from the find window.....rather than
  ' the first tab-ordered control having focus, the control that had the focus on the
  ' previous record still has it !  Not sure this is the right place for the code though.
  '#############

  ' Set the first control on this tab to be highlighted
'''  FocusFirstControl
  
  ' Set up the menus and toolbars
  ' JPD20021209 Fault 4863
  'frmMain.RefreshMainForm Me
  frmMain.RefreshMainForm Screen.ActiveForm
  
  ' Refresh any navigation controls because Version One has some teething troubles
  For Each objControl In Me.Controls
    If TypeOf objControl Is COA_Navigation Then
      objControl.RefreshControls
    End If
  Next
  DoEvents
  
  ' Refresh the form icon
  'GetIcon mlngPictureID
    
End Sub

Private Sub Form_Initialize()
  'Set form loading property
  mfLoading = True
  
  Set mdatRecEdit = New clsRecEdit
  Set mrsRecords = New ADODB.Recordset

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  '# RH 26/08/99. To pass shortcut keys thru to the activebar control
  Dim fHandled As Boolean
'
'  If KeyCode <> vbKeyF1 Then
'    fHandled = frmMain.abMain.OnKeyDown(KeyCode, Shift)
'
'    If fHandled Then
'      KeyCode = 0
'      Shift = 0
'    End If
'  End If
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    Case Else
      fHandled = frmMain.abMain.OnKeyDown(KeyCode, Shift)
      If fHandled Then
        KeyCode = 0
        Shift = 0
      End If
      DoEvents
      
  End Select

End Sub


Private Sub Form_Load()
  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl GTMaskDate1(0)
  
  'Remove default tab from tabstrip control
  fraTabPage(0).BackColor = vbButtonFace
  
  With TabStrip1
    .Tabs.Clear
    .TabIndex = 0
  End With

  ' Set the user defined activebar
  OrganiseToolbarControls ActiveBar1
  'NHRD15092006 Fault 11493
  ActiveBar1.Bands(0).Flags = 1 + 256 + 512
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iLoop As Integer
  Dim frmTemp As Form
  Dim afrmForms() As Form
  Dim intErrorLine As Integer
  
  On Local Error GoTo LocalErr
  
  DebugOutput "frmRecEdit4", "Form_QueryUnload"
  
  intErrorLine = 10
  
  DoEvents
  
  intErrorLine = 20
  
  'JPD 20030820 Fault 3048
  For Each frmTemp In Forms
    With frmTemp
      intErrorLine = 30
      If .Name = "frmRecEdit4" Then
        intErrorLine = 40
        If (.FormID = mlngParentFormID) And (.Changed) Then
          intErrorLine = 50
          Cancel = True
          Exit Sub
        End If
        
      'JPD 20040628 Fault 8278
      ElseIf .Name = "frmFind2" Then
        intErrorLine = 60
        If (.ParentFormID = mlngFormID) Then
          intErrorLine = 70
          If .Busy Then
            intErrorLine = 80
            Cancel = True
            Exit Sub
          End If
        End If
      End If
    End With
  Next frmTemp
  Set frmTemp = Nothing
  
  intErrorLine = 90
  mfUnloading = True
  
  If Not mfLoading Then
    ' Unload this form's children.
    
    intErrorLine = 100
    If Me.Visible Then
      'MH20040304 Fault 8212
      intErrorLine = 110
      If Me.Enabled Then
        intErrorLine = 120
        Me.SetFocus
      End If
      intErrorLine = 130
      frmMain.RefreshMainForm Me
      intErrorLine = 140
    End If
    
    For Each frmTemp In Forms
      With frmTemp
        If .Name = "frmRecEdit4" Then
          intErrorLine = 150
          If .ParentFormID = mlngFormID Then
            intErrorLine = 160
            .ParentUnload = True
'MH20001124 Fault 1320
'Message box appearing twice .... sort it out !
'            If Not .SaveChanges Then
            intErrorLine = 170
            If .RecordChanged Then
            
            'TM20010911 Fault 4
            'Need to call the SaveChanges routine to get save prompts when closing the
            'form.
              intErrorLine = 180
              If Not .SaveChanges Then
                intErrorLine = 190
                Cancel = 1
                mfUnloading = False
                intErrorLine = 200
                Exit Sub
              End If
            End If
          End If
        End If
      End With
    Next
    Set frmTemp = Nothing
    
    
    intErrorLine = 210
    
    ' Need to get an array of frmRecEdit so that we do not have to iterate
    ' through the forms collection again as any forms such as the find forms
    ' that will be destroyed when child windows are destroyed will be reinitialised
    ' when their name is checked since they did exist at this point.
    ReDim afrmForms(0)
    For iLoop = 0 To Forms.Count - 1
      intErrorLine = 220
      If Forms(iLoop).Name = "frmRecEdit4" Then
        intErrorLine = 230
        Set afrmForms(UBound(afrmForms)) = Forms(iLoop)
        intErrorLine = 240
        ReDim Preserve afrmForms(UBound(afrmForms) + 1)
      End If
    Next iLoop


    intErrorLine = 250
    
    'MH20040506 Fault 8321
    'Check if save changes to parent BEFORE unloading child screens..
    If (UnloadMode <> vbAppWindows) And _
      (UnloadMode <> vbAppTaskManager) Then
      Cancel = (Not SaveChanges)
    End If
    
    
    intErrorLine = 260
    
    mfSavingInProgress = True
    ReDim Preserve afrmForms(UBound(afrmForms) - 1)
    For iLoop = 0 To UBound(afrmForms)
      intErrorLine = 270
      If afrmForms(iLoop).ParentFormID = mlngFormID Then
        intErrorLine = 280
        Unload afrmForms(iLoop)
      End If
    Next iLoop
    mfSavingInProgress = False


    'MH20040506 Fault 8321
    'Moved to above "child unload" section
    'If (UnloadMode <> vbAppWindows) And _
    '  (UnloadMode <> vbAppTaskManager) Then
    '  Cancel = (Not SaveChanges)
    'End If
    
    
    intErrorLine = 290
    If Not Database.Validation Then
      Cancel = True
      mfUnloading = False
      intErrorLine = 300
      Exit Sub
    End If

    intErrorLine = 310
    
    If Cancel = False Then
      ' Unload any 'find' or 'summary' forms.
      If Not mfrmFind Is Nothing Then
        intErrorLine = 320
        If Not mfrmFind.IsLoading Then
          intErrorLine = 330
          Unload mfrmFind
          Set mfrmFind = Nothing
          intErrorLine = 340
        Else
          intErrorLine = 350
          Cancel = True
          mfUnloading = False
          intErrorLine = 360
          Exit Sub
        End If
        
        intErrorLine = 370
      End If

      intErrorLine = 380
      If Not mfrmSummary Is Nothing Then
        intErrorLine = 390
        If Not mfParentUnload Then
          intErrorLine = 400
          mfrmSummary.Rebind
          intErrorLine = 410
          mfrmSummary.Visible = True
          intErrorLine = 420
          ' JPD20021127 Fault 4218
          mfrmSummary.SetCurrentRecord
          intErrorLine = 430
          Me.Visible = False
          intErrorLine = 440
          Cancel = True
          intErrorLine = 450
          mfUnloading = False
          intErrorLine = 460
          Exit Sub
       Else
          intErrorLine = 470
          Unload mfrmSummary
          intErrorLine = 480
          Set mfrmSummary = Nothing
          intErrorLine = 490
       End If
     End If
    End If
  End If

  intErrorLine = 500
  mfUnloading = False

Exit Sub

LocalErr:
  COAMsgBox Err.Description & vbCrLf & "(RecEdit4 - Form_QueryUnload " & CStr(intErrorLine) & ")", vbCritical

End Sub


Public Property Let ParentUnload(ByVal pfNewValue As Boolean)
  mfParentUnload = pfNewValue
    
End Property


Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo Err_Trap

  Dim fTemp As Form
  Dim objLUValue As CLookupValue
  Dim lngParentHWnd As Long
  Dim sTag As String

  DebugOutput "frmRecEdit4", "Form_Unload"
  
  ' If this is an Add New Table form then enable all the others.
  If mfTableEntry Then
    ' Get the Hwnd value of the parent form.
    For Each objLUValue In gcoLookupValues.Collection
      If objLUValue.ChildHwnd = Me.hWnd Then
        lngParentHWnd = objLUValue.ParentHwnd
        Exit For
      End If
    Next objLUValue
    Set objLUValue = Nothing
    
    ' Get the parent form.
    For Each fTemp In Forms
      If fTemp.hWnd = lngParentHWnd Then
        fTemp.EnableMe
        
        If Not mfLeaveLookup Then
          fTemp.SetLookupValue gcoLookupValues.Item(CStr(lngParentHWnd)).ColumnID, mrsRecords.Fields(gcoLookupValues.Item(CStr(lngParentHWnd)).LookupColName).Value
        Else
          fTemp.SetLookupValue gcoLookupValues.Item(CStr(lngParentHWnd)).ColumnID, ""
        End If
      End If
    Next

    mfLookup = True

    gcoLookupValues.Remove CStr(lngParentHWnd)
  Else
    mfLookup = False
  End If

  'Close database access for this form
  If Filtered Then
    mrsRecords.Filter = ""
  End If


'MH20040218 Fault 8080
'  If mrsRecords.EditMode <> adEditNone Then
'    mrsRecords.CancelUpdate
'  End If
'  If mrsRecords.State <> adStateClosed Then
'    mrsRecords.Close
'  End If
  If mrsRecords.State <> adStateClosed Then
    If mrsRecords.EditMode <> adEditNone Then
      mrsRecords.CancelUpdate
    End If
    mrsRecords.Close
  End If

  Set mrsRecords = Nothing


  frmMain.RefreshMainForm Me, True
  
  'Release internal classes
  Set ODBC = Nothing
  Set mcolColumnPrivileges = Nothing
  Set mdatRecEdit = Nothing
  Set mobjTableView = Nothing

  For Each fTemp In Forms
    If fTemp.hWnd = lngParentHWnd Then
      fTemp.SetFocus
      frmMain.RefreshMainForm fTemp
      Exit For
    End If
  Next
  
  Exit Sub

Err_Trap:

End Sub


Public Sub ShowHistorySummary()

  'MH20040211 Fault 8078
  On Local Error Resume Next

  If mfrmSummary Is Nothing Then
    HistoryInitialise
  Else
    mfrmSummary.SetFocus
  End If

End Sub

Public Sub HistoryInitialise()
  
  Dim fOK As Boolean
 
  If mfrmSummary Is Nothing Then
    Set mfrmSummary = New frmFind2
        
    With mfrmSummary
      .Visible = False
      .ScreenType = screenHistorySummary


      'MH20001019 Faults 1082, 1160
      ' Check result of sub and don't show form if it fails
      '.FindStartFromPrimary(mobjTableView, mlngOrderID, Me, False)
      fOK = .FindStartFromPrimary(mobjTableView, mlngOrderID, Me, False)
      If fOK = False Then
        Unload mfrmSummary
        Set mfrmSummary = Nothing
        Exit Sub
      End If

      If GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & Me.ScreenID, "Width", 0) = 0 Then
        UI.frmAtCenter mfrmSummary
      Else
        .Width = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & Me.ScreenID, "Width", Screen.Width / 2)
        .Height = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & Me.ScreenID, "Height", Screen.Height / 2)
        .Top = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & Me.ScreenID, "Top", (Screen.Height - Me.Height) / 2)
        .Left = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & Me.ScreenID, "Left", (Screen.Width - Me.Width) / 2)
      End If
      
      UI.LockWindow .hWnd
      .Visible = True
      .Show
      .ResizeFindColumns
      .SetFocus
      UI.UnlockWindow
    End With
  Else
    Unload Me
  End If

End Sub


Public Property Get ScreenType() As ScreenType
  ScreenType = miScreenType
  
End Property

Public Property Let ScreenType(piNewValue As ScreenType)
  miScreenType = piNewValue
  
End Property

Public Property Get FormID() As Long
  FormID = mlngFormID
  
End Property

Public Property Let FormID(plngNewValue As Long)
  mlngFormID = plngNewValue
  
End Property

Public Property Get ParentFormID() As Long
  ParentFormID = mlngParentFormID
  
End Property
Public Property Get ParentID() As Long
  GetParentDetails
  ParentID = mlngParentRecordID
    
End Property

Public Property Let ParentFormID(plngNewValue As Long)
  mlngParentFormID = plngNewValue
  
End Property

Public Property Let LookupLoading(pfLoading As Boolean)
  mfLookupLoading = pfLoading

End Property

Public Function LoadScreen(ByVal plngScreenID As Long, ByVal plngViewID As Long) As Boolean
  ' Load a screen from the database and populate it with the correct controls.
  On Error GoTo ErrorTrap_LoadScreen

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iPageNo As Integer
  Dim intMinFormWidth As Integer
  Dim lngOffsetX As Long
  Dim lngOffsetY As Long
  Dim objScreen As clsScreen
  Dim fraFrame As Frame
  Dim blnAddedNew As Boolean
  Dim iExtraHeight As Integer
  Dim iExtraWidth As Integer
  Dim frmForm As Form
  ' NPG20090907 Fault HRPRO-324
  Dim iToolBarHeight As Integer

  mlngScreenID = plngScreenID
  
  'MH20071130 Fault 12666 (v3.5.8)
  mfRequiresLocalCursor = GetSystemSetting("RecEdit", "LocalCursor", False)

  ' Get the screen definition object.
  Set objScreen = mdatRecEdit.GetScreen(mlngScreenID)

  fOK = Not objScreen Is Nothing
  
  ' Get the screen's table/view object.
  If fOK Then
    
    mlngPictureID = objScreen.PictureID
    
    If plngViewID > 0 Then
      Set mobjTableView = gcoTablePrivileges.FindViewID(plngViewID)
      If Not mfRequiresLocalCursor Then
        mfRequiresLocalCursor = datGeneral.ViewRequiresLocalCursor(plngViewID)
      End If
    Else
      Set mobjTableView = gcoTablePrivileges.FindTableID(objScreen.TableID)
      
      'mfRequiresLocalCursor = False
      'JPD 20051018 Fault 10465/Fault 10466
      If (miScreenType = screenHistoryTable) Or _
        (miScreenType = screenHistoryView) Or _
        (miScreenType = screenQuickEntry) Then

        'JPD 20040625 Fault 8545 - Need to use the local cursor if any of the child views parent views
        ' use the hierarchy functions. Originally wejust checked the view that was the parent of the
        ' current record edit screen. This was not enough. We need to check all view parents of
        ' the child view.
        If Not mfRequiresLocalCursor Then
           mfRequiresLocalCursor = datGeneral.ChildViewRequiresLocalCursor(objScreen.TableID)
        End If
'''        ' We are a history table.
'''        ' Get the parent table and view ID
'''        For Each frmForm In Forms
'''          With frmForm
'''            If .Name = "frmRecEdit4" Then
'''              If (.FormID = mlngParentFormID) Then
'''                mfRequiresLocalCursor = .RequiresLocalCursor
'''                Exit For
'''              End If
'''            End If
'''          End With
'''        Next frmForm
'''        Set frmForm = Nothing
      End If
    End If
    
    fOK = Not mobjTableView Is Nothing
  End If

  If fOK Then
  
    Select Case Me.ScreenType
        Case screenParentTable
            Select Case Me.TableID
              Case glngBHolRegionTableID
                Me.HelpContextID = 1120
              Case glngCourseTableID
                Me.HelpContextID = 1131
              Case glngPersonnelTableID
                Me.HelpContextID = 1138
              Case glngPostTableID
                Me.HelpContextID = 1142
            Case Else
                Me.HelpContextID = 0
            End Select
        
        Case screenParentView
            Select Case Me.TableID
              Case glngCourseTableID
                Me.HelpContextID = 1131
            Case Else
                Me.HelpContextID = 0
            End Select
            
        Case screenHistoryTable
            Select Case Me.TableID
              Case 9999
                Me.HelpContextID = 1119
                If Me.ScreenName = "Training Bookings (from Courses)" Then Me.HelpContextID = 1119
              Case Else
                Me.HelpContextID = 0
            End Select
        
        'The rest of these case statements could be used when further detailed
        'Context Sensitive Help is required.  At present decided it was a bit
        'too detailed for development time to implement
        Case screenHistoryView
                Me.HelpContextID = 1119
    
        Case screenLookup
                Me.HelpContextID = 1118
            
        Case screenFind
                Me.HelpContextID = 0
        
        Case screenHistorySummary
                Me.HelpContextID = 0
          
        Case screenQuickEntry
                Me.HelpContextID = 1117
        
        Case screenPickList
                Me.HelpContextID = 0
        Case Else
                Me.HelpContextID = 0
    End Select
  
    With mobjTableView
      ' Get the name of the table/view.
      msTableViewName = IIf(.ViewID > 0, .ViewName, .TableName)

      ' Get the screen's order. Use the table's default order if the screen does not have one defined.
      mlngOrderID = IIf(IsNull(objScreen.OrderID), 0, objScreen.OrderID)
      If mlngOrderID <= 0 Then
        mlngOrderID = .DefaultOrderID
      End If
    
      ' Get the Record Description Expression ID.
      mlngRecDescID = .RecordDescriptionID
    End With
    
    msScreenName = objScreen.ScreenName

    ' Get the column privileges collection.
    SetupColumnPrivileges

    ' Set the screen properties.
    Me.Visible = False
    
    mfUseTab = (UBound(objScreen.TabCaptions) > 0)
    TabStrip1.Visible = mfUseTab

    If mfUseTab Then
      ' Set up the tabstrip font
      With TabStrip1.Font
        .Name = objScreen.FontName
        .Size = objScreen.FontSize
        .Bold = objScreen.FontBold
        .Italic = objScreen.FontItalic
        .Strikethrough = objScreen.FontStrikethru
        .Underline = objScreen.FontUnderline
      End With

      iPageNo = 0
      For iLoop = 1 To UBound(objScreen.TabCaptions)
        iPageNo = iPageNo + 1
        TabStrip1.Tabs.Add

        Load fraTabPage(iPageNo)
        
        'MH20010823 Fault 1883
        'TabStrip1.Tabs(iPageNo).Caption = objScreen.TabCaptions(iLoop)
        TabStrip1.Tabs(iPageNo).Caption = Replace(objScreen.TabCaptions(iLoop), "&", "&&")
      Next iLoop
    End If
    
    If mfUseTab Then
      TabStrip1.Tabs(1).Selected = True
      fraTabPage(1).Visible = True
      fraTabPage(1).ZOrder 0
    End If

    GetParentDetails

    ' Load the controls onto the screen.
    
    UI.LockWindow Me.hWnd
    fOK = LoadControls(objScreen)
    UI.UnlockWindow
  End If

  If fOK Then
    ' Initialise the filter on the current recordset.
    ReDim mavFilterCriteria(3, 0)

    ' Get the records.
    fOK = GetRecords
  End If
  
  If fOK Then
    ' If we are a parenttable or parentview then see if we have insert privileges if no records
    If (mrsRecords.EOF And mrsRecords.BOF) And _
      ((miScreenType = screenParentTable Or _
      miScreenType = screenParentView Or _
      miScreenType = screenLookup Or _
      miScreenType = screenQuickEntry)) And _
      Not mobjTableView.AllowInsert Then
  
      gobjProgress.Visible = False
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      'COAMsgBox "You do not have 'new' permission on this empty " & IIf(mobjTableView.ViewID > 0, "view.", "table."), vbExclamation, "Security"
      COAMsgBox "You do not have 'new' permission on this empty " & IIf(Me.ViewID > 0, "view.", "table."), vbExclamation, "Security"
      fOK = False
    End If
  End If

  If fOK Then
    ' MH20000714 Fault 550 (Quick Entry crash when no records exist)
    ' blnAddedNew flag checks to see if a new record has been added
    ' this avoids the problem to doing an '.addnew' twice
    blnAddedNew = True
    
    ' Add a new record if there are none.
    With mrsRecords
      If (.EOF And .BOF) Then
        .AddNew
      Else
        If .EditMode = adEditAdd Then
          .AddNew
        Else
          mrsRecords.MoveFirst
          blnAddedNew = False
        End If
      End If
    End With
      
    ' Load the controls with the correct values
    If Not mfLookup Then
      
'      If (miScreenType = screenQuickEntry) Then
'        If blnAddedNew = False Then
'          mrsRecords.AddNew
'        End If
'      End If
    
      UpdateControls True

    End If
  End If
  
  ' JPD20020924
  Screen.MousePointer = vbHourglass

ExitLoadScreen:
  'gobjProgress.CloseProgress
  If fOK Then
    
    ' Refresh the toolbar.
    '''LOFTY-Performance-Be more conservative calling the refresh function
    frmMain.RefreshMainForm Me

    ' NPG20090907 Fault HRPRO-324
    ' iToolBarHeight = 90   ' Old toolbar was 16px high, which equated to 90 'somethings'
    iToolBarHeight = 190  ' New toolbar is 24px high

    ' Set the form dimensions.
    '### This needs to be here as toolbar width isnt updated until
    '### the above line of code has been run
    Select Case ActiveBar1.Bands(0).DockingArea
    
      Case giTOOLBAR_NONE
        iExtraHeight = StatusBar1.Height + 450 '+ iToolBarHeight
        iExtraWidth = iToolBarHeight
      Case giTOOLBAR_TOP
        iExtraHeight = StatusBar1.Height + (ActiveBar1.Bands(0).Height * Screen.TwipsPerPixelY) + iToolBarHeight
        iExtraWidth = iToolBarHeight
      Case giTOOLBAR_BOTTOM
        iExtraHeight = StatusBar1.Height + (ActiveBar1.Bands(0).Height * Screen.TwipsPerPixelY) + iToolBarHeight
        iExtraWidth = iToolBarHeight
      Case giTOOLBAR_LEFT
        If (ActiveBar1.Bands(0).MaxHeight * Screen.TwipsPerPixelY) > objScreen.Height - StatusBar1.Height Then
          iExtraHeight = ((ActiveBar1.Bands(0).MaxHeight * Screen.TwipsPerPixelY) - (objScreen.Height - StatusBar1.Height))
        Else
          iExtraHeight = StatusBar1.Height + iToolBarHeight
        End If
        iExtraHeight = iExtraHeight + StatusBar1.Height
        iExtraWidth = ActiveBar1.Bands(0).Width * Screen.TwipsPerPixelY + iToolBarHeight
      Case giTOOLBAR_RIGHT
        If (ActiveBar1.Bands(0).MaxHeight * Screen.TwipsPerPixelY) > objScreen.Height - StatusBar1.Height Then
          iExtraHeight = ((ActiveBar1.Bands(0).MaxHeight * Screen.TwipsPerPixelY) - (objScreen.Height - StatusBar1.Height))
        Else
          iExtraHeight = StatusBar1.Height + iToolBarHeight
        End If
        iExtraHeight = iExtraHeight + StatusBar1.Height
        iExtraWidth = ActiveBar1.Bands(0).Width * Screen.TwipsPerPixelY + iToolBarHeight
    End Select
    
    Me.Width = objScreen.Width + iExtraWidth
    Me.Height = objScreen.Height + iExtraHeight
        
    intMinFormWidth = Screen.TwipsPerPixelX * _
      (Me.ActiveBar1.Bands(0).MaxWidth + UI.GetSystemMetrics(SM_CXFRAME))
    If Me.Width < intMinFormWidth Then
      Me.Width = intMinFormWidth
    End If
    
    ' Size the tabstrip (and constituent frames) to fit the form.
    lngOffsetX = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
    lngOffsetY = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
'    TabStrip1.Move lngOffsetX, lngOffsetY, Me.ScaleWidth - (lngOffsetX * 2), _
'      Me.ScaleHeight - (ActiveBar1.Bands(0).Height * Screen.TwipsPerPixelY)
    TabStrip1.Move lngOffsetX, lngOffsetY, (Me.ScaleWidth - (lngOffsetX * 2)), _
      (Me.ScaleHeight - ((iExtraHeight - iToolBarHeight) - StatusBar1.Height))

    For Each fraFrame In fraTabPage
      With fraFrame
        .Top = TabStrip1.ClientTop
        .Left = TabStrip1.ClientLeft
        .Height = TabStrip1.ClientHeight
        .Width = TabStrip1.ClientWidth
      End With
    Next fraFrame
    Set fraFrame = Nothing
  End If

  Set objScreen = Nothing
  
  LoadScreen = fOK
  Screen.MousePointer = vbDefault
  Exit Function

ErrorTrap_LoadScreen:
  gobjProgress.Visible = False
  COAMsgBox Err.Description & " - LoadScreen", vbCritical
  fOK = False
  Err = False
  Resume ExitLoadScreen

End Function


Public Sub UpdateControls(Optional pfNoWarnings As Boolean)
  ' This function populates the controls with the values from the
  ' database for the current record.
  On Error GoTo Err_Trap
  
  DebugOutput "UpdateControls", "Start"
  
  Dim objFile As New ADODB.Stream
  Dim fFound As Boolean
  Dim fResetControl As Boolean
  Dim fProgBarVisible As Boolean
  Dim iLoop As Integer
  Dim lngParentRecordID As Long
  Dim sDefault As String
  Dim fFileNameOK As Boolean
  Dim sTag As String
  Dim sColumnName As String
  Dim objControl As Control
  Dim objColumn As ADODB.Field
  Dim alngParentTables() As Long
  Dim fldColumn As ADODB.Field
  Dim asWarnings() As String
  Dim sWarning As String
  Dim sFormat As String
  Dim sDateFormat As String
  Dim lngRetryCount As Long
    
  ReDim asWarnings(0)
  lngRetryCount = 0

  ' Do not update the controls if we are unloading the form.
  If mfUnloading Then Exit Sub
  If mrsRecords.State = adStateClosed Then Exit Sub

  Screen.MousePointer = vbHourglass
  'UI.LockWindow Me.hwnd
  
  mfLeaveLookup = False
  mfLoading = True

  mfDataChanged = False
  ' JPD20021007 Fault 4498
  ReDim malngChangedOLEPhotos(0)
  sDateFormat = DateFormat
  
  ' JPD 30/8/00 Multi-user correction.
  ' Remember the original timestamp.
  'mlngTimeStamp = mrsRecords!Timestamp
  'mlngRecordID = mrsRecords!ID
  mlngTimeStamp = IIf(IsNull(mrsRecords!Timestamp), 0, mrsRecords!Timestamp)
  mlngRecordID = IIf(IsNull(mrsRecords!ID), 0, mrsRecords!ID)
  
  ' Update the form caption.
  RefreshFormCaption

  Database.Validation = True

  ' Loop through the controls and use the tags in order to get the
  ' correct column data from the table for each control
  ReDim alngParentTables(0)
  
  For Each objControl In Me.Controls
    With objControl
      ' get the control's tag.
      'sTag = .Tag
      sTag = GetTag(objControl)

      'JPD 20030610
      If TypeOf objControl Is ActiveBar Then
        sTag = ""
      End If
      
      ' Check if the control is tagged to a column.
      If LenB(sTag) > 0 Then
        If mobjScreenControls.Item(sTag).ColumnID > 0 Then
          ' If we are in add mode then get the default's from the table
          ' otherwise use the actual table values
          If (mrsRecords.EditMode = adEditAdd) Or (mfLookupLoading) Then

            If TypeOf objControl Is CommandButton Then
              ' Clear the link record ID.
              For iLoop = 1 To UBound(malngLinkRecordIDs, 2)
                If malngLinkRecordIDs(1, iLoop) = mobjScreenControls.Item(sTag).LinkTableID Then
                  malngLinkRecordIDs(2, iLoop) = 0
                  Exit For
                End If
              Next iLoop
            Else
              'MH20031002 Fault 7082 Reference Property instead of object to trap errors
              'If mobjScreenControls.Item(sTag).TableID = mobjTableView.TableID Then
              If mobjScreenControls.Item(sTag).TableID = Me.TableID Then
                
                'MH20001109 Fault 981 Pass in parent table and record now !
                'SetControlDefaults .Tag, objControl
                SetControlDefaults .Tag, objControl, mlngParentTableID, mlngParentRecordID
              
              Else
                fFound = False
                For iLoop = 1 To UBound(alngParentTables)
                  If alngParentTables(iLoop) = mobjScreenControls.Item(sTag).TableID Then
                    fFound = True
                    Exit For
                  End If
                Next iLoop
                
                If Not fFound Then
                  ' Add the parent table to the array for updating below.
                  iLoop = UBound(alngParentTables) + 1
                  ReDim Preserve alngParentTables(iLoop)
                  alngParentTables(iLoop) = mobjScreenControls.Item(sTag).TableID
                End If
              End If
            End If
          Else
            ' Not adding new.
            ' First check if it's a different table.
            'MH20031002 Fault 7082 Reference Property instead of object to trap errors
            'If mobjScreenControls.Item(sTag).TableID <> mobjTableView.TableID Then
            If mobjScreenControls.Item(sTag).TableID <> Me.TableID Then
              ' Add the table ID to an array. Parent table colums are update below.
              fFound = False
              For iLoop = 1 To UBound(alngParentTables)
                If alngParentTables(iLoop) = mobjScreenControls.Item(sTag).TableID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              
              If Not fFound Then
                ' Add the parent table to the array for updating below.
                iLoop = UBound(alngParentTables) + 1
                ReDim Preserve alngParentTables(iLoop)
                alngParentTables(iLoop) = mobjScreenControls.Item(sTag).TableID
              End If
            Else
              ' Update the controls.
              If TypeOf objControl Is CommandButton Then
                sColumnName = "ID_" & Trim(Str(mobjScreenControls.Item(sTag).LinkTableID))
              Else
                sColumnName = mobjScreenControls.Item(sTag).ColumnName
              End If

              ' Check if the control's associated column has been read.
              fResetControl = True
              For Each fldColumn In mrsRecords.Fields
                If UCase$(Trim$(fldColumn.Name)) = UCase$(Trim$(sColumnName)) Then
                  fResetControl = False
                  Exit For
                End If
              Next fldColumn
              Set fldColumn = Nothing
              
              If TypeOf objControl Is TDBText6Ctl.TDBText Then
                If fResetControl Then
                  .Text = ""
                Else
DebugOutput "UpdateControls", "Before SetTDBText"
                  '.Text = RTrim(mrsRecords(sColumnName).Value & vbNullString)
                  SetTDBText objControl, RTrim(mrsRecords(sColumnName).Value & vbNullString)
DebugOutput "UpdateControls", "After SetTDBText"
                End If

              ElseIf TypeOf objControl Is COA_Image Then
                If fResetControl Then
                  .Picture = Nothing
                  .ASRDataField = ""
                Else
                  
                  ' Embedded/linked image
                  If mobjScreenControls.Item(sTag).OLEType = OLE_EMBEDDED Or mobjScreenControls.Item(sTag).OLEType = OLE_UNC Then
                    .OLEType = OLE_UNC 'mobjScreenControls.Item(sTag).OLEType
                    ReadStream objControl, mobjScreenControls.Item(sTag), False
                    
                    Set .Picture = Nothing
                    If .EmbeddedStream.State = adStateOpen Then
                      If .EmbeddedStream.Size > 0 Then
                        Set .Picture = LoadPictureFromStream(.EmbeddedStream)
                      End If
                    End If
                  
                  Else
                  
                    If mrsRecords(sColumnName).Type = adVarChar Then
                      fFileNameOK = False
  
                      If Not IsNull(mrsRecords(sColumnName).Value) Then
                        'TM20011130 Fault 3050 - Check if path exists before loading img.
                        If Len(mrsRecords(sColumnName).Value) > 0 And (Dir(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & mrsRecords(sColumnName).Value, vbDirectory) <> vbNullString) Then
                          fFileNameOK = True
                          .Picture = LoadPicture(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & mrsRecords(sColumnName).Value)
                          .ASRDataField = mrsRecords(sColumnName).Value
                        End If
                      End If
  
                      If Not fFileNameOK Then
                        .Picture = Nothing
                        .ASRDataField = ""
                      End If
                    End If
                  
                  End If
                End If

              ElseIf TypeOf objControl Is TDBMask6Ctl.TDBMask Then
                If fResetControl Then
                  .Text = ""
                Else
                  .Text = RTrim(mrsRecords(sColumnName).Value & vbNullString)
                End If
 
              ElseIf TypeOf objControl Is TextBox Then
                If fResetControl Then
                  .Text = ""
                Else
                  '.Text = RTrim(mrsRecords(sColumnName) & vbNullString)
                  .Text = mrsRecords(sColumnName).Value & vbNullString
                End If

              'JPD 20050302 Fault 9847
              ElseIf (TypeOf objControl Is TDBNumberCtrl.TDBNumber) Or _
                (TypeOf objControl Is TDBNumber6Ctl.TDBNumber) Then
                .ClearControl
                If Not fResetControl Then
                  'JPD 20050316 Fault 9911
                  '.Value = mrsRecords(sColumnName)
                  .Value = IIf(IsNull(mrsRecords(sColumnName).Value), 0, mrsRecords(sColumnName).Value)
                End If

              ElseIf TypeOf objControl Is XtremeSuiteControls.CheckBox Then
                If fResetControl Then
                  .Value = 0
                Else
                  .Value = IIf(mrsRecords(sColumnName).Value, 1, 0)
                End If

              ElseIf TypeOf objControl Is XtremeSuiteControls.ComboBox Then
                If fResetControl Then
                  If .ListCount > 0 Then
                    .ListIndex = 0
                  End If
                Else
                  .ListIndex = UI.cboSelect(objControl, RTrim(mrsRecords(sColumnName).Value & vbNullString))
                  
                  'JPD 20031016 Fault 7292
                  If (Not IsNull(mrsRecords(sColumnName).Value)) And (Trim(UCase(.Text)) <> Trim(UCase(mrsRecords(sColumnName).Value))) Then
                    'JPD 20030905 Fault 6076
                    ReDim Preserve asWarnings(UBound(asWarnings) + 1)
                    asWarnings(UBound(asWarnings)) = sColumnName
                  
                    SetControlDefaults .Tag, objControl, mlngParentTableID, mlngParentRecordID
                  End If
                End If

              ElseIf TypeOf objControl Is COA_Lookup Then
                If fResetControl Then
                  .Text = ""
                Else
                  'JPD 20050810 Fault 10165
                  If (mobjScreenControls.Item(objControl.Tag).DataType = sqlNumeric) Then
                    sFormat = "0"
                    If mobjScreenControls.Item(objControl.Tag).Use1000Separator Then
                      sFormat = "#,0"
                    End If
                    If mobjScreenControls.Item(objControl.Tag).Decimals > 0 Then
                      sFormat = sFormat & "." & String(mobjScreenControls.Item(objControl.Tag).Decimals, "0")
                    End If
                                
                    .Text = IIf(IsNull(mrsRecords(sColumnName).Value), "", Format(mrsRecords(sColumnName).Value, sFormat))
                  Else
                    .Text = IIf(IsNull(mrsRecords(sColumnName).Value), "", mrsRecords(sColumnName).Value)
                  End If
                End If

              ElseIf TypeOf objControl Is COA_OptionGroup Then
                If fResetControl Then
                  .Text = ""
                Else
                  .Text = RTrim(mrsRecords(sColumnName).Value & vbNullString)
                  
                  'JPD 20031016 Fault 7292
                  If Trim(UCase(.Text)) <> Trim(UCase(mrsRecords(sColumnName).Value)) Then
                    'JPD 20030905 Fault 6076
                    ReDim Preserve asWarnings(UBound(asWarnings) + 1)
                    asWarnings(UBound(asWarnings)) = sColumnName

                    SetControlDefaults .Tag, objControl, mlngParentTableID, mlngParentRecordID
                  End If
                End If

              ElseIf TypeOf objControl Is COA_OLE Then
                If fResetControl Then
                  .FileName = vbNullString
'                  .ToolTipText = vbNullString
                  .OleOnServer = OLEType(mobjScreenControls.Item(sTag).ColumnID)
                Else
                 
                  ' Embedded document
                  If mobjScreenControls.Item(sTag).OLEType = OLE_EMBEDDED Or mobjScreenControls.Item(sTag).OLEType = OLE_UNC Then
                    .OLEType = OLE_UNC ' mobjScreenControls.Item(sTag).OLEType
                  
                    ' Get header information for document
                    ReadStream objControl, mobjScreenControls.Item(sTag), True
                                        
                    If .EmbeddedStream.State = adStateOpen Then
                      If .EmbeddedStream.Size > 0 Then
                        .OLEType = LoadOLETypeFromStream(objControl.EmbeddedStream)
                        .FileName = LoadFileNameFromStream(objControl.EmbeddedStream, (.OLEType = OLE_UNC), False)
                      End If
                    
                      ' Close the stream as we only have the header section (re-read if they click on the OLE button)
                      .EmbeddedStream.Close
                    
                    Else
                      .FileName = ""
                    End If
                 
                  Else
                    fFileNameOK = False

                    If Not IsNull(mrsRecords(sColumnName).Value) Then
                      If Len(mrsRecords(sColumnName).Value) > 0 Then
                        fFileNameOK = True
                        .FileName = mrsRecords(sColumnName).Value
                        If mobjScreenControls.Item(sTag).OLEType = OLE_SERVER Then
                          If Dir(gsOLEPath & IIf(Right(gsOLEPath, 1) = "\", "", "\"), vbDirectory) <> vbNullString Then
                            .FileName = gsOLEPath & IIf(Right(gsOLEPath, 1) = "\", "", "\") & mrsRecords(sColumnName).Value
                          Else
                            .FileName = mrsRecords(sColumnName).Value
                          End If
                          '.ToolTipText = IIf(Dir(gsOLEPath & "\*.*") <> vbNullString, gsOLEPath & "\", vbNullString) & mrsRecords(sColumnName)
                        Else
                          If Dir(gsLocalOLEPath & IIf(Right(gsLocalOLEPath, 1) = "\", "", "\"), vbDirectory) <> vbNullString Then
                            .FileName = gsLocalOLEPath & IIf(Right(gsLocalOLEPath, 1) = "\", "", "\") & mrsRecords(sColumnName).Value
                          Else
                            .FileName = mrsRecords(sColumnName).Value
                          End If
                          '.ToolTipText = IIf(Dir(gsLocalOLEPath & "\*.*") <> vbNullString, gsLocalOLEPath & "\", vbNullString) & mrsRecords(sColumnName)
                        End If
                        .OLEType = mobjScreenControls.Item(sTag).OLEType
                      End If
                    End If

                    If Not fFileNameOK Then
                      .FileName = vbNullString
                      .OLEType = mobjScreenControls.Item(sTag).OLEType
                    End If
                  
                  End If
                End If

              ElseIf TypeOf objControl Is COA_Spinner Then
                If fResetControl Then
                  .Text = ""
                Else
                  .Text = mrsRecords(sColumnName).Value & vbNullString
                End If

              ElseIf TypeOf objControl Is GTMaskDate.GTMaskDate Then
                If fResetControl Then
                  .Text = ""
                Else
                  If IsNull(mrsRecords(sColumnName).Value) Then
                    .Text = ""
                  Else
                    .Text = Format(DateValue(mrsRecords(sColumnName).Value), sDateFormat)
                  End If
                 End If

              ElseIf TypeOf objControl Is COA_WorkingPattern Then
                If fResetControl Then
                  .Value = ""
                Else
                  .Value = mrsRecords(sColumnName).Value & vbNullString
                End If
              
              ElseIf TypeOf objControl Is COA_Navigation Then
                If fResetControl Then
                  .NavigateTo = ""
                Else
                  .NavigateTo = mrsRecords(sColumnName).Value & vbNullString
                End If
              
              ElseIf TypeOf objControl Is COA_ColourSelector Then
                If fResetControl Then
                  .BackColor = vbWhite
                Else
                  .BackColor = Val(mrsRecords(sColumnName).Value)
                End If
              
              ElseIf TypeOf objControl Is CommandButton Then
                lngParentRecordID = 0
                If Not fResetControl Then
                  lngParentRecordID = IIf(IsNull(mrsRecords(sColumnName).Value), 0, mrsRecords(sColumnName).Value)
                End If
                      
                For iLoop = 1 To UBound(malngLinkRecordIDs, 2)
                  If malngLinkRecordIDs(1, iLoop) = mobjScreenControls.Item(sTag).LinkTableID Then
                    malngLinkRecordIDs(2, iLoop) = lngParentRecordID
                    Exit For
                  End If
                Next iLoop
              End If
            End If
          End If
        End If
      End If
    End With
  Next objControl
  Set objControl = Nothing
 
  'JPD 20030905 Fault 6076
  If (UBound(asWarnings) > 0) And (Not pfNoWarnings) Then
    If UBound(asWarnings) = 1 Then
      sWarning = "The '" & asWarnings(1) & "' field has been given a default value as it's value was no longer valid."
    Else
      sWarning = "The following fields have been given default values as their values were no longer valid :" & vbNewLine
  
      For iLoop = 1 To UBound(asWarnings)
        sWarning = sWarning & vbNewLine & vbTab & asWarnings(iLoop)
      Next iLoop
    End If
  
    COAMsgBox sWarning, vbExclamation + vbOKOnly, Application.Name
    mfDataChanged = True
  End If

  'JPD20010815 Fault 2239 Read the original status, employee record ID, and course record ID
  ' for Training Booking records.
  msTBOriginalStatus = ""
  mlngTBOriginalEmpID = 0
  mlngTBOriginalCourseID = 0
  'MH20031002 Fault 7082 Reference Property instead of object to trap errors
  'If (mobjTableView.TableID = glngTrainBookTableID) Then
  If (Me.TableID = glngTrainBookTableID) Then
    ' Remember the original booking status.
    For Each fldColumn In mrsRecords.Fields
      If UCase(Trim(fldColumn.Name)) = UCase(Trim(gsTrainBookStatusColumnName)) Then
        If Not IsNull(mrsRecords(gsTrainBookStatusColumnName)) Then
          msTBOriginalStatus = RTrim(mrsRecords(gsTrainBookStatusColumnName))
        End If
      End If
      If UCase(Trim(fldColumn.Name)) = "ID_" & Trim(Str(glngEmployeeTableID)) Then
        If Not IsNull(mrsRecords("ID_" & Trim(Str(glngEmployeeTableID)))) Then
          mlngTBOriginalEmpID = mrsRecords("ID_" & Trim(Str(glngEmployeeTableID)))
        End If
      End If
      If UCase(Trim(fldColumn.Name)) = "ID_" & Trim(Str(glngCourseTableID)) Then
        If Not IsNull(mrsRecords("ID_" & Trim(Str(glngCourseTableID)))) Then
          mlngTBOriginalCourseID = mrsRecords("ID_" & Trim(Str(glngCourseTableID)))
        End If
      End If
    Next fldColumn
    Set fldColumn = Nothing
  End If
  
  For iLoop = 1 To UBound(alngParentTables)
    UpdateParentControls alngParentTables(iLoop)
  Next iLoop

  If mrsRecords.EditMode = adEditAdd Then
    If GetParentDetails Then
      UpdateParentControls mlngParentTableID, mlngParentRecordID
    End If
  End If
  
  mfLoading = False

  Set objFile = Nothing

'MsgBox "1"
  'UI.UnlockWindow
'MsgBox "2"
  Me.Refresh
  
  Screen.MousePointer = vbDefault
  'Set objFile = Nothing
  
  DebugOutput "UpdateControls", "End"
  
  Exit Sub

Err_Trap:
  DebugOutput "UpdateControls", "Error"
  
  Select Case Err.Number
      Case 28 'Out of Stack Space
        If lngRetryCount < 3 Then
          lngRetryCount = lngRetryCount + 1
          DoEvents
          Resume 0
        Else
          MsgBox Err.Description, vbCritical, "UpdateControls Error"
          Resume Next
        End If
    Case 3021, 91
      Resume Next
    ' JPD20020828 Fault 4176
    Case 52, 53, 75, 76
      fFileNameOK = False
      Resume Next
    Case 380
      ' Mask control was being populated with a value that does not fit the mask.
      If Len(sColumnName) > 0 Then
        fProgBarVisible = gobjProgress.Visible
        gobjProgress.Visible = False
        COAMsgBox "The '" & sColumnName & "' field does not match the defined mask.", vbExclamation + vbOKOnly, Application.Name
        gobjProgress.Visible = fProgBarVisible
      End If
      mfDataChanged = True
      Resume Next
    
    ' Invalid Picture
    Case 481
      Resume Next
    
    Case 31031
'      fProgBarVisible = gobjProgress.Visible
'      gobjProgress.Visible = False
'      COAMsgBox "The '" & sColumnName & "' column contains an OLE object stored locally," & vbNewLine & _
'             "however, the object or its path does not exist on your machine.", vbInformation + vbOKOnly, Application.Name
'      gobjProgress.Visible = fProgBarVisible
      Resume Next
    Case Else
        fProgBarVisible = gobjProgress.Visible
        gobjProgress.Visible = False
        COAMsgBox Err.Description, vbCritical
        gobjProgress.Visible = fProgBarVisible
        ' If the problem was an ole link type error, then still attempt to update the other fields
        If InStr(Err.Description, "link") Then Resume Next
  End Select

End Sub


Private Function SetTDBText(objText As TDBText, strInput As String) As Boolean

  On Local Error GoTo LocalErr
  
  objText.Text = strInput
  SetTDBText = False
  
Exit Function

LocalErr:
  COAMsgBox "Error setting text column" & vbCrLf & "(" & Err.Description & ")", vbCritical
  SetTDBText = False

End Function


Public Function GetTag(objControl As Control) As String

  Dim lngTimeOut As Long
  
  On Local Error GoTo LocalErr
  lngTimeOut = Timer + 3
  GetTag = objControl.Tag

Exit Function

LocalErr:
  If lngTimeOut > Timer Then
    Resume 0
  End If
  GetTag = vbNullString

End Function



Public Function UpdateWithAVI(Optional pfDeactivating As Variant, Optional SaveCaption As String) As Boolean
  ' NPG20090902 Fault HRPRO-219
  ' NHRD12102010 JIRA HRPRO-1107
  ' Show progress bar
  With gobjProgress
    .AVI = dbSaveRec
    If Len(SaveCaption) > 0 Then
      .Caption = SaveCaption
    Else
      .Caption = "Saving changes..."
    End If
    .NumberOfBars = 0
    .Time = False
    .Cancel = False
    .OpenProgress
  End With
  
  UpdateWithAVI = Update(pfDeactivating)

  gobjProgress.CloseProgress
        
End Function



Public Function Update(Optional pfDeactivating As Variant) As Boolean
  ' Stores the current values to the database.
  ' NB. We construct an SQL statement instead of using the ADO recordset 'update' method
  ' as we sometime found problems when using this method with server-side cursors (which we need
  ' for the large amount of data in some systems).
  ' Return TRUE if the record was saved okay.
  On Error GoTo ErrorTrap
    
  Dim fSavedOK As Boolean
  Dim fColumnDone As Boolean
  Dim iLoop As Integer
  Dim iChangeReason As Integer
  Dim iNextIndex As Integer
  Dim lngRecCount As Long
  Dim lngCurrentID As Long
  Dim lngParentTableID As Long
  Dim sColumnName As String
  Dim sSQL As String
  Dim sValueList As String
  Dim sColumnList As String
  Dim objControl As Control
  Dim sTag As String
  Dim asColumns() As String
  Dim objTableView As CTablePrivilege
  
  Dim blnHasLink As Boolean
  Dim sTableName As String
  Dim aryTableNames()
  Dim sMessage As String
  
  Dim bUploadOLEObjects As Boolean
  
  Dim fFound As Boolean
  Dim iLoop2 As Integer
  Dim fDoControl As Boolean
  
  ' JPD20021206 Fault 4854
  mfSavingInProgress = True
  
  bUploadOLEObjects = False
  blnHasLink = False
  sTableName = ""
  
  fSavedOK = True
    
  If IsMissing(pfDeactivating) Then pfDeactivating = False
    
  ' Dimension the array of columns and values to be updated.
  ' NB. Column 1 = column name in uppercase.
  '     Column 2 = column value as it needs to appear in the SQL update/insert string.
  ReDim asColumns(2, 0)

  'MH20040212 Faults 8078 & 8079
  If mrsRecords.State = adStateClosed Then
    Exit Function
  End If


  ' Check if the record has been amended elsewhere.
  If mrsRecords.EditMode <> adEditAdd Then
    'MH20031002 Fault 7082 Reference Property instead of object to trap errors
    'iChangeReason = RecordAmended2(mobjTableView.RealSource, mobjTableView.TableID, mlngRecordID, mlngTimeStamp)
    iChangeReason = datGeneral.RecordAmended(Me.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
    If iChangeReason > 0 Then
      AmendedRecord2 True, iChangeReason
      Update = True

      ' JPD20021206 Fault 4854
      mfSavingInProgress = False
        
      Exit Function
    End If
  End If
  
  lngRecCount = RecordCount
  
  ' Check if the record has changed.
  If Not mbResendingToAccord Then
    If lngRecCount > 0 Then
      'TM20020528 Fault 2895 - only check if the record has changed if there are controls
      'which do do have defaults.
      If Not AllDefaults Then
        If Not RecordChanged Then
          ' Do nothing if the record has not changed.
          Update = True
          
          ' JPD20021206 Fault 4854
          mfSavingInProgress = False
          
          Exit Function
        End If
      End If
    End If
  End If
  
  ' Check if we are a child screen with link controls that we have a parent.
  lngParentTableID = 0
  If (miScreenType = screenQuickEntry) Or _
    (miScreenType = screenHistoryTable) Or _
    (miScreenType = screenHistoryView) Then

    If GetParentDetails Then
      lngParentTableID = mlngParentTableID
    End If

'********************************************************************************
' TM110701 - Fault 2551.                                                        *
'            Verifies that at least one link is made to a Parent table.         *
'********************************************************************************
    If (miScreenType = screenQuickEntry) Then
      ' Check that we have links to all required parent tables.
      For iLoop = 1 To UBound(malngLinkRecordIDs, 2)
        If (malngLinkRecordIDs(2, iLoop) = 0) And _
          (malngLinkRecordIDs(1, iLoop) <> lngParentTableID) Then
          ' Link not made. Inform the user, and stop the update operation.
          For Each objTableView In gcoTablePrivileges.Collection
            If objTableView.TableID = malngLinkRecordIDs(1, iLoop) Then
            
              ReDim Preserve aryTableNames(iLoop + 1)
              sTableName = objTableView.TableName
              aryTableNames(iLoop) = sTableName
              
            End If
          Next objTableView
          Set objTableView = Nothing
          
        Else
          blnHasLink = True
        End If
      Next iLoop
  
      If Not blnHasLink And (UBound(malngLinkRecordIDs, 2) > 0) Then
      
        ' NPG20090902 Fault HRPRO-336
        gobjProgress.CloseProgress
      
      
        sMessage = ""
        For iLoop = 1 To UBound(aryTableNames) - 1 Step 1
        If sMessage = "" Then
          sMessage = aryTableNames(iLoop)
        Else
          sMessage = sMessage & " or " & aryTableNames(iLoop)
        End If
        Next iLoop
        
        'NHRD04062003 Fault 5708
        sMessage = "Unable to save record, a link must be made with the " & sMessage & " table."
        
        If sTableName <> "" Then
          COAMsgBox sMessage, vbExclamation, Me.Caption
        Else
          COAMsgBox "Unable to save record, a link must be made with the parent table.", vbExclamation, Me.Caption
        End If
        Update = False
        
        ' JPD20021206 Fault 4854
        mfSavingInProgress = False
        
        Exit Function
      
      'TM20020219 Fault 3522 - Added to catch if the table has no link columns defined.
      ElseIf (UBound(malngLinkRecordIDs) = 0) Then
        COAMsgBox "Unable to save record, a link must be made with the parent table.", vbExclamation, Me.Caption
        Update = False
        
        ' JPD20021206 Fault 4854
        mfSavingInProgress = False
        
        Exit Function
        
      End If
    End If
  End If
      
  sTableName = ""

'********************************************************************************
' TM110701 - Fault 2551.                                                        *
'            End.
'********************************************************************************

  Screen.MousePointer = vbHourglass
  Database.Validation = True
  
  ' Construct an update/insert string to enter the new values into the database.
  lngCurrentID = mlngRecordID

  ' Loop through the screen controls, creating an array of columns and values with
  ' which we'll construct an insert or update SQL string.
  For Each objControl In Me.Controls
    sTag = objControl.Tag

    'JPD 20030610
    If TypeOf objControl Is ActiveBar Then
      sTag = ""
    End If
      
    ' Check if it is a user editable control.
    If Len(sTag) > 0 Then
      ' Check that the control is associated with a column in the current table/view,
      ' and is updatable.
      If (mobjScreenControls.Item(sTag).ColumnID > 0) Then
      
        'JPD 20040706 Fault 8993
        fDoControl = objControl.Enabled Or mobjScreenControls.Item(sTag).ScreenReadOnly
               
        If fDoControl Then
          
          ' Check we have write permission
          If mcolColumnPrivileges.IsValid(mobjScreenControls.Item(sTag).ColumnName) Then
            fDoControl = mcolColumnPrivileges.Item(mobjScreenControls.Item(sTag).ColumnName).AllowUpdate
          Else
            fDoControl = False
          End If
          
          ' NPG20121114 fault HRPRO-2720 Only allow the control to be added if it hasn't already been excluded
          If fDoControl Then
            If TypeOf objControl Is TDBText6Ctl.TDBText Then
               fDoControl = Not objControl.ReadOnly
            End If
  
            If TypeOf objControl Is COA_Navigation Then
              fDoControl = False
            End If
          End If

        End If
        
        If fDoControl Then
          ' Get the name of the column associated with the current control.
          If TypeOf objControl Is CommandButton Then
            sColumnName = "ID_" & Trim(Str(mobjScreenControls.Item(sTag).LinkTableID))
          Else
            sColumnName = UCase(mobjScreenControls.Item(sTag).ColumnName)
          End If
          
          ' Check if the column's update string has already been constructed.
          fColumnDone = False
          For iNextIndex = 1 To UBound(asColumns, 2)
            If asColumns(1, iNextIndex) = sColumnName Then
              fColumnDone = True
              Exit For
            End If
          Next iNextIndex

          ' JPD20021007 Fault 4498
          ' Do not do photo controls if no path is defined.
          If Not fColumnDone Then
            If (TypeOf objControl Is COA_Image) Or _
              (TypeOf objControl Is COA_OLE) Then
              
              fFound = False
              For iLoop2 = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop2) = mobjScreenControls.Item(sTag).ColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop2
              If Not fFound Then
                fColumnDone = True
              End If
            End If
          End If

          If Not fColumnDone Then
            ' Add the column name to the array of columns that have already been entered in the
            ' SQL update/insert string.
            iNextIndex = UBound(asColumns, 2) + 1
            ReDim Preserve asColumns(2, iNextIndex)
            asColumns(1, iNextIndex) = sColumnName
            asColumns(2, iNextIndex) = ""
            
            ' Construct the SQL update/insert string for the column.
            If TypeOf objControl Is TDBText6Ctl.TDBText Then
              ' Multi-line character field from a masked textbox (CHAR type column). Save the text from the control.
              objControl.Text = CaseConversion(objControl.Text, mobjScreenControls.Item(sTag).ConvertCase)
              If IsNull(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType)) Then
                asColumns(2, iNextIndex) = "''"
              Else
                asColumns(2, iNextIndex) = "'" & Replace(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "'", "''") & "'"
              End If
            
            ElseIf TypeOf objControl Is COA_Image Then
              ' Photo field (CHAR type column). Save the name of the photo file.
              If objControl.OLEType = OLE_EMBEDDED Or objControl.OLEType = OLE_UNC Then
                ' Embedded document
                bUploadOLEObjects = True
                asColumns(2, iNextIndex) = "null"
              Else
                asColumns(2, iNextIndex) = "'" & Replace(objControl.ASRDataField, "'", "''") & "'"
              End If
                           
            ElseIf TypeOf objControl Is TDBMask6Ctl.TDBMask Then
              ' Character field from a masked textbox (CHAR type column). Save the text from the control.
              If Len(objControl.Value) = 0 Then
                asColumns(2, iNextIndex) = "null"
              Else
                objControl.Text = CaseConversion(objControl.Text, mobjScreenControls.Item(sTag).ConvertCase)
                asColumns(2, iNextIndex) = "'" & Replace(objControl.Text, "'", "''") & "'"
              End If
                           
            ElseIf TypeOf objControl Is TextBox Then
              ' Character field from an unmasked textbox (CHAR type column). Save the text from the control.
              objControl.Text = CaseConversion(objControl.Text, mobjScreenControls.Item(sTag).ConvertCase)
              asColumns(2, iNextIndex) = "'" & Replace(objControl.Text, "'", "''") & "'"
    
            'JPD 20050302 Fault 9847
            ElseIf (TypeOf objControl Is TDBNumberCtrl.TDBNumber) Or _
              (TypeOf objControl Is TDBNumber6Ctl.TDBNumber) Then
              ' Integer or Numeric field from a numeric textbox (INT or NUM type column). Save the value from the control.
              
              'JPD 20050309 - changed number control, so now need to handle 'null' values.
              asColumns(2, iNextIndex) = ConvertData(IIf(IsNull(objControl.Value), 0, objControl.Value), mobjScreenControls.Item(sTag).DataType)
            
              'MH20010108
              asColumns(2, iNextIndex) = datGeneral.ConvertNumberForSQL(asColumns(2, iNextIndex))

            
            ElseIf TypeOf objControl Is XtremeSuiteControls.CheckBox Then
              ' Logic field (BIT type column). Save 1 for true, 0 for False.
              asColumns(2, iNextIndex) = IIf(objControl.Value, "1", "0")
              
            ElseIf TypeOf objControl Is XtremeSuiteControls.ComboBox Then
              ' Character field from a combo (CHAR type column). Save the text from the combo.
              asColumns(2, iNextIndex) = "'" & Replace(objControl.Text, "'", "''") & "'"
    
            ElseIf TypeOf objControl Is COA_Lookup Then
              ' Lookup field from a combo (unknown type column). Get the column type and save the appropraite value from the combo.
              Select Case mobjScreenControls.Item(sTag).DataType
                Case sqlVarChar, sqlLongVarChar
                  asColumns(2, iNextIndex) = "'" & Replace(objControl.Text, "'", "''") & "'"
                Case sqlNumeric, sqlInteger
                  'MH20070320 Fault 12052
                  'asColumns(2, iNextIndex) = Str(Val(objControl.Text))
                  asColumns(2, iNextIndex) = CStr(Val(Replace(objControl.Text, UI.GetSystemThousandSeparator, "")))
                Case sqlDate
                  If Len(objControl.Text) > 0 Then
                    asColumns(2, iNextIndex) = "'" & Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
                  Else
                    asColumns(2, iNextIndex) = "null"
                  End If
                Case Else
                  asColumns(1, iNextIndex) = ""
              End Select
    
            ElseIf TypeOf objControl Is COA_OptionGroup Then
              ' Character field from an option group (CHAR type column). Save the text from the option group.
              asColumns(2, iNextIndex) = "'" & Replace(objControl.Text, "'", "''") & "'"
    
            ElseIf TypeOf objControl Is COA_OLE Then
              'asColumns(2, iNextIndex) = "'" & objControl.ID & "'"
              bUploadOLEObjects = True

              If objControl.OLEType = OLE_EMBEDDED Or objControl.OLEType = OLE_UNC Then
                ' Embedded document
                asColumns(2, iNextIndex) = "null"

              Else
                ' Linked document
                If Len(objControl.FileName) > 0 Then
                  asColumns(2, iNextIndex) = "'" & Replace(Mid(objControl.FileName, InStrRev(objControl.FileName, "\") + 1), "'", "''") & "'"
                Else
                  asColumns(2, iNextIndex) = "null"
                End If
              End If
  
            ElseIf TypeOf objControl Is COA_Spinner Then
              ' Integer field from an spinner (INT type column). Save the value from the spinner.
              asColumns(2, iNextIndex) = Trim(Str(Val(objControl.Text)))
    
            ElseIf TypeOf objControl Is GTMaskDate.GTMaskDate Then
              ' Date field from a date control (DATETIME type column). Save the value from the control formatted as 'mm/dd/yyyy' for SQL.
              If Not pfDeactivating Then
                If Len(Trim(Replace(objControl.Text, UI.GetSystemDateSeparator, ""))) <> 0 Then
                  If Not IsDate(objControl.DateValue) Or objControl.DateValue < #1/1/1753# Then
                    objControl.ForeColor = vbRed
                    COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, app.title
                    objControl.ForeColor = vbWindowText
                    objControl.DateValue = Null
                    If objControl.Visible And objControl.Enabled Then
                      objControl.SetFocus
                    End If
                    fSavedOK = False
                    Screen.MousePointer = vbDefault
                    
                    ' JPD20021206 Fault 4854
                    mfSavingInProgress = False
                    
                    ' NPG20090902 Fault HRPRO-219
                    gobjProgress.CloseProgress
                    
                    
                    Exit Function
                  End If
                End If
              End If
              
              If IsNull(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType)) Then
                asColumns(2, iNextIndex) = "null"
              Else
                asColumns(2, iNextIndex) = "'" & Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
              End If
    
            ElseIf TypeOf objControl Is COA_WorkingPattern Then
              ' Working Pattern Field (CHAR type column, len 14).
              asColumns(2, iNextIndex) = "'" & Replace(objControl.Value, "'", "''") & "'"
    
            ElseIf TypeOf objControl Is COA_ColourSelector Then
              asColumns(2, iNextIndex) = CStr(objControl.BackColor)
    
            ElseIf TypeOf objControl Is CommandButton Then
              If mobjScreenControls.Item(sTag).LinkTableID <> lngParentTableID Then
                For iLoop = 1 To UBound(malngLinkRecordIDs, 2)
                  If malngLinkRecordIDs(1, iLoop) = mobjScreenControls.Item(sTag).LinkTableID Then
                    asColumns(2, iNextIndex) = Trim(Str(malngLinkRecordIDs(2, iLoop)))
                    Exit For
                  End If
                Next iLoop
              Else
                asColumns(2, iNextIndex) = Trim(Str(mlngParentRecordID))
              End If
            End If
          End If
        End If
      End If
    End If
  Next objControl
  Set objControl = Nothing
  
  ' See if we are a history screen and if we are save away the id of the parent also
  If GetParentDetails Then
    ' Check if the column's update string has already been constructed.
    fColumnDone = False
    For iNextIndex = 1 To UBound(asColumns, 2)
      If asColumns(1, iNextIndex) = "ID_" & mlngParentTableID Then
        fColumnDone = True
        Exit For
      End If
    Next iNextIndex

    If Not fColumnDone Then
      ' Add the column name to the array of columns that have already been entered in the
      ' SQL update/insert string.
      iNextIndex = UBound(asColumns, 2) + 1
      ReDim Preserve asColumns(2, iNextIndex)
      asColumns(1, iNextIndex) = "ID_" & Trim(Str(mlngParentTableID))
      asColumns(2, iNextIndex) = Trim(Str(mlngParentRecordID))
    End If
  End If
  
  ' Perform module specific validation.
  ' TRAINING BOOKING MODULE SPECIFICS.
  If Not ValidateTrainingBookingRecord(lngCurrentID, asColumns) Then
    Screen.MousePointer = vbDefault
    Update = False

    ' JPD20021206 Fault 4854
    mfSavingInProgress = False
    
    ' NPG20090902 Fault HRPRO-219
    gobjProgress.CloseProgress
    
    Exit Function
  End If
  
  If UBound(asColumns, 2) > 0 Then
       
    ' Create a SQL string to update the record with.
    Select Case mrsRecords.EditMode
      Case adEditAdd
        ' Construct the SQL insert string from the array of columns and values.
        sColumnList = ""
        sValueList = ""
        For iLoop = 1 To UBound(asColumns, 2)
          If Len(asColumns(1, iLoop)) > 0 Then
            sColumnList = sColumnList & IIf(Len(sColumnList) > 0, ", ", "") & asColumns(1, iLoop)
            sValueList = sValueList & IIf(Len(sValueList) > 0, ", ", "") & asColumns(2, iLoop)
          End If
        Next iLoop
        
        sSQL = "INSERT INTO " & mobjTableView.RealSource & " (" & sColumnList & ") VALUES (" & sValueList & ")"
        fSavedOK = datGeneral.InsertTableRecord(sSQL, Me.TableID, lngCurrentID)

        If fSavedOK Then
          fSavedOK = CopyWhenParentRecordIsCopied(Me.TableID, lngCurrentID, OriginalRecordID)
        End If

      Case Else
        ' Construct the SQL update string from the array of columns and values.
        sColumnList = ""
        For iLoop = 1 To UBound(asColumns, 2)
          If Len(asColumns(1, iLoop)) > 0 Then
            sColumnList = sColumnList & IIf(Len(sColumnList) > 0, ", ", "") & asColumns(1, iLoop) & " = " & asColumns(2, iLoop)
          End If
        Next iLoop
          
        If Len(sColumnList) > 0 Then
          sSQL = "UPDATE " & mobjTableView.RealSource & " SET " & sColumnList & " WHERE id = " & Trim(Str(lngCurrentID))

          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'fSavedOK = datGeneral.UpdateTableRecord(sSQL, mobjTableView.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
          fSavedOK = datGeneral.UpdateTableRecord(sSQL, Me.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
        End If
    End Select
    
        ' Upload any OLE objects
    If bUploadOLEObjects And fSavedOK Then
      
      ' Refresh the recordset as we may have copied the record
      'RefreshRecordset
      
      For iLoop = 1 To UBound(malngChangedOLEPhotos)
        
        'datgeneral.GetColumnName(
        For Each objControl In Me.Controls
          
          If TypeOf objControl Is COA_OLE _
            Or TypeOf objControl Is COA_Image Then

              If objControl.OLEType = OLE_EMBEDDED Or objControl.OLEType = OLE_UNC Then
                If objControl.ColumnID = malngChangedOLEPhotos(iLoop) Then
                  SaveStream objControl, mobjScreenControls.Item(objControl.Tag), lngCurrentID
                End If
              End If
          End If

        Next objControl
      Next iLoop
    End If
    
  End If
  
  If fSavedOK Then
       
    ' Refresh the recordset as the current record may no longer be in it.
    RefreshRecordset
        
    ' Try to locate the current record in the recordset.
    ' If the current record is no longer in the recordset then inform the user why we've
    ' moved to the first record.
    If mrsRecords.EditMode <> adEditAdd Then
      LocateRecord lngCurrentID
      
      If mrsRecords!ID <> lngCurrentID Then
        If Filtered Then
          COAMsgBox "The record saved does not satisfy the current filter.", vbExclamation, app.ProductName
        'MH20031002 Fault 7082 Reference Property instead of object to trap errors
        'ElseIf mobjTableView.ViewID > 0 Then
        ElseIf Me.ViewID > 0 Then
          COAMsgBox "The record saved is no longer in the current view.", vbExclamation, app.ProductName
        End If
      End If
    Else
      If Filtered Then
        COAMsgBox "The record saved does not satisfy the current filter.", vbExclamation, app.ProductName
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      'ElseIf mobjTableView.ViewID > 0 Then
      ElseIf Me.ViewID > 0 Then
        COAMsgBox "The record saved is no longer in the current view.", vbExclamation, app.ProductName
      End If
    End If
  
    'MH20070228 Fault 11641
    RefreshFormCaption

    ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
    ' UpdateChildren call as the summary fields that are updated in
    ' UpdateChildren may be dependent on the parent recordset being
    ' refreshed first (in UpdateParentWindow).
    UpdateParentWindow
    UpdateChildren
    UpdateFindWindow
    UpdateSiblingWindows lngCurrentID, False
'    UpdateParentWindow

    'TM05072004 - Not too sure that this should be done as well as the UpdateParentWindow, UpdateChildren etc. stuff.
    'Need to rework this.
    If Not mbDisableAURefresh Then
      Update_AutoUpdateScreens (Me.FormID)
    End If
  
    'MH20040212 Faults 8078 & 8079
    If mrsRecords.State <> adStateClosed Then
      ' Need to relocate to the current record, as the 'UpdateParentWindow' call
      ' may have put us back on the first history record.
      If mrsRecords.EditMode <> adEditAdd Then
        If mrsRecords!ID <> lngCurrentID Then     'MH20071130 Fault 12666
          LocateRecord lngCurrentID
        End If
      End If
    End If
  End If

ExitUpdate:
  Screen.MousePointer = vbDefault

  If fSavedOK Then
    UpdateControls
    mfDataChanged = False
    ' JPD20021007 Fault 4498
    ReDim malngChangedOLEPhotos(0)
    objEmail.SendImmediateEmails
    
    ' Do any bespoke post save navigation code
    ExecutePostSaveCode
    
  Else
    Database.Validation = False
  End If

  'If Not frmMain Is Nothing Then
    frmMain.RefreshMainForm Me

    'MH20040218 Fault 8080
    If Not (gcoTablePrivileges Is Nothing) Then
      ' JPD20021206 Fault 4854
      frmMain.RefreshMainForm Screen.ActiveForm
      frmMain.CheckForNonactiveForms Screen.ActiveForm
    Else
      frmMain.CheckForNonactiveForms Screen.ActiveForm
      Unload frmMain
    End If
  'End If
  mfSavingInProgress = False
  
  ' NPG20090902 Fault HRPRO-219
  gobjProgress.CloseProgress
  
  Update = fSavedOK
  Exit Function

ErrorTrap:
  fSavedOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
  Resume ExitUpdate

End Function

Public Property Get ViewName() As String
  ' Return the view name.
  Dim frmForm As Form
    
  For Each frmForm In Forms
    If (frmForm.Name = "frmRecEdit4") Then
      
      If frmForm.FormID = mlngParentFormID Then
        ViewName = frmForm.ViewName
        Exit Function
      End If
    End If
  Next frmForm

  ViewName = mobjTableView.ViewName

End Property
Public Function AllDefaults() As Boolean

  'TM20020528 Fault 2895 - AllDefaults function checks if all the user updatable controls
  ' have defaults, this is so the save button can be enabled if the all controls have defaults.
  
  Dim objControl As Control
  Dim bAllDef As Boolean
  Dim sTag As String
  
  bAllDef = True
 
  For Each objControl In Me.Controls
    With objControl
      sTag = .Tag
      
      'JPD 20030610
      If TypeOf objControl Is ActiveBar Then
        sTag = ""
      End If
      
      If Len(sTag) > 0 Then
        If mobjScreenControls.Item(sTag).ColumnID > 0 Then
          If (mobjScreenControls.Item(sTag).DfltValueExprID = 0) Then
            If .Enabled Then
              ' Append the result parameter.
              Select Case mobjScreenControls.Item(sTag).DataType
                Case sqlOle ' OLE columns do not have defaults.
                Case sqlBoolean  ' Logic columns automatically have a default
                Case sqlNumeric   ' Numeric columns
                  If mobjScreenControls.Item(sTag).DefaultValue = vbNullString Then bAllDef = False
                Case sqlInteger  ' Integer columns
                  If mobjScreenControls.Item(sTag).DefaultValue = vbNullString Then bAllDef = False
                Case sqlDate
                  If mobjScreenControls.Item(sTag).DefaultValue = vbNullString Then bAllDef = False
                Case sqlVarChar ' Character columns
                  If mobjScreenControls.Item(sTag).DefaultValue = vbNullString Then bAllDef = False
                Case sqlVarBinary ' Photo columns do not have defaults.
                Case sqlLongVarChar ' Working Pattern columns
                  If mobjScreenControls.Item(sTag).DefaultValue = vbNullString Then bAllDef = False
                Case Else
              End Select
              
              'JPD 20030819 Fault 6383
              If mobjScreenControls.Item(sTag).ControlType = ctlCommand Then
                bAllDef = False
              End If
            End If
          End If
        End If
      End If
    End With
  
    If Not bAllDef Then
      Exit For
    End If
  Next objControl

  Set objControl = Nothing
  
  AllDefaults = bAllDef
  
End Function

Private Sub AmendedRecord2(pfShowMessage As Boolean, piChangeReason As Integer)
  ' Tell the user that the current record has been amended by another user.
  ' Try to refresh the recordset.
  Dim sMsg As String
  
  ' Do nothing if the record hasn't changed.
  If piChangeReason = 0 Then
    Exit Sub
  End If
    
  If pfShowMessage Then
    Select Case piChangeReason
      Case 1 ' The record has been amended AND is still in the current table/view.
        sMsg = "The record has been amended by another user and will be refreshed."
      Case 2 ' The record has been amended AND is no longer in the current table/view.
        sMsg = "The record has been amended by another user and is no longer in the current view, screen will be refreshed."
      Case 3 ' The record has been deleted from the table.
        sMsg = "The record has been deleted by another user, screen will be refreshed."
    End Select
    
    COAMsgBox sMsg, vbExclamation, app.ProductName
  End If
  
  Screen.MousePointer = vbHourglass
  
  ' Refresh the recordset.
  If Not RefreshRecordset Then
    Exit Sub
  End If
    
  If mrsRecords.EditMode <> adEditAdd Then
    ' Locate the current record if it is still in the recordset.
    If piChangeReason = 1 Then
      LocateRecord mlngRecordID
    Else
      mrsRecords.MoveFirst
    End If
  End If
  
  ' Update all controls and associated screens.
  ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
  ' UpdateChildren call as the summary fields that are updated in
  ' UpdateChildren may be dependent on the parent recordset being
  ' refreshed first (in UpdateParentWindow).
  UpdateParentWindow
  UpdateControls
  UpdateChildren
  UpdateFindWindow
'  UpdateParentWindow
    
  ' Refresh the menu.
  Screen.MousePointer = vbDefault
  frmMain.RefreshMainForm Me

End Sub







Private Function CaseConversion(psText As String, piCaseConversion As Integer) As String
  ' Perform the required case conversion on the given text.
  Dim lngPos As Long
  Dim sLastCharacter As String

  ' Do nothing if the given text is empty.
  'TM20020107 Fault 3323
'  If Len(Trim(psText)) > 0 Then
  If Len(psText) > 0 Then
  
    ' Do nothing if the given text is numeric.
    If Not IsNumeric(psText) Then
    
      Select Case piCaseConversion
        Case 0      'No conversion
            
        Case 1      'Upper case
          psText = UCase$(psText)
              
        Case 2      'Lower case
          psText = LCase$(psText)
        
        Case 3      'Proper conversion
          ' First LCase everything !
          'TM20020107 Fault 3323
'          psText = LCase(Trim(psText))
          psText = LCase$(psText)
          
          ' Then Ucase first letter
          psText = UCase$(Left$(psText, 1)) & Right$(psText, Len(psText) - 1)

          ' JPD 25/5/00 Corrected the 'propercase' function to handle Jean-Louis O'Sullivan-McNeill.
          ' JPD 22/1/01 Recognises é, â, etc. as alphabetic characters.
          For lngPos = 2 To Len(psText)
            sLastCharacter = Mid(psText, lngPos - 1, 1)
            If ((sLastCharacter < "A") Or (sLastCharacter > "Z")) And _
              ((sLastCharacter < "a") Or (sLastCharacter > "z")) And _
              ((sLastCharacter < "0") Or (sLastCharacter > "9")) And _
              ((sLastCharacter < "À") Or (sLastCharacter > "Ö")) And _
              ((sLastCharacter < "Ù") Or (sLastCharacter > "Ý")) And _
              ((sLastCharacter < "ß") Or (sLastCharacter > "ö")) And _
              ((sLastCharacter < "ù") Or (sLastCharacter > "ÿ")) Then
              
              psText = Left$(psText, lngPos - 1) & UCase$(Mid$(psText, lngPos, 1)) & Right$(psText, Len(psText) - lngPos)
            ElseIf lngPos > 2 Then
              ' Catch the McName.
              If (Mid(psText, lngPos - 2, 1) = "M") And (sLastCharacter = "c") Then
                psText = Left$(psText, lngPos - 1) & UCase$(Mid$(psText, lngPos, 1)) & Right$(psText, Len(psText) - lngPos)
              End If
            End If
          Next lngPos
      End Select
    End If
  End If
  
  CaseConversion = psText
    
End Function


Public Property Get ViewID() As Long
  'MH20031002 Fault 7082
  If mobjTableView Is Nothing Then
    ViewID = 0
  Else
    ' Return the view ID.
    ViewID = mobjTableView.ViewID
  End If
End Property

Public Property Get ParentTableID() As Long
  ParentTableID = mlngParentTableID

End Property

Public Property Get TableView() As CTablePrivilege
  Set TableView = mobjTableView

End Property


Public Property Get ColumnSelectPrivileges() As CColumnPrivileges
  ' Get's the select column privileges.
  Set ColumnSelectPrivileges = mcolColumnPrivileges
  
End Property



Public Property Get Cancelled() As Boolean
  ' Return the cancelled flag.
  Cancelled = mfCancelled

End Property


Public Sub CancelCourse()
  ' TRAINING BOOKING MODULE SPECIFICS.
  ' Cancel the current course record.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fTransferred As Boolean
  Dim fInTransaction As Boolean
  Dim fTBRecordsExists As Boolean
  Dim iUserChoice As Integer
  Dim sSQL As String
  Dim sCourseTitle As String
  Dim sErrorMsg As String
  Dim frmCourseSelection As frmTransferCourseBookings
  Dim objColumns As CColumnPrivileges
  Dim objWLTable As CTablePrivilege
  Dim objTBColumn As CColumnPrivilege
  Dim objWLColumn As CColumnPrivilege
  Dim objWLColumnPrivileges As CColumnPrivileges
  Dim objTBTable As CTablePrivilege
  Dim objTBColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim lngCurrentRecordID As Long
  ' NPG20080414 Fault 13023
  Dim alngRelatedColumns() As Long
  Dim iLoop As Integer
  Dim asAddedColumns() As String
  Dim sColumnList As String
  Dim sValueList As String
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  
  
    
  fOK = True
  fInTransaction = False
  fTBRecordsExists = False
  lngCurrentRecordID = mlngRecordID
  
  'NHRD15012007 Fault 3905, 07022007 Fault 11943
  If COAMsgBox("Are you sure you want to cancel this Course ?", vbQuestion + vbYesNo, app.ProductName) = vbNo Then
    Exit Sub
  End If
  
  ' Save changes if required.
  If SaveChanges Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If
    
    If fOK Then
      ' Check that the Training Booking table can be read.
      Set objTBTable = gcoTablePrivileges.FindTableID(glngTrainBookTableID)
      fOK = Not objTBTable Is Nothing
      If fOK Then
        ' Check that the current user can read the table.
        fOK = objTBTable.AllowSelect
      Else
        COAMsgBox "Unable to find the '" & gsTrainBookTableName & "' table.", vbOKOnly, app.ProductName
      End If
    End If
    
    If fOK Then
      Set objTBColumnPrivileges = GetColumnPrivileges(objTBTable.TableName)
      
      fOK = objTBColumnPrivileges.Item(gsTrainBookStatusColumnName).AllowUpdate
      If Not fOK Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column.", vbOKOnly, app.ProductName
      End If
  
      ' If the training booking cancellation date is defined, check that the current user can update it.
      If Len(gsTrainBookCancelDateColumnName) > 0 Then
        fOK = objTBColumnPrivileges.Item(gsTrainBookCancelDateColumnName).AllowUpdate
        If Not fOK Then
          COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookCancelDateColumnName & "' column.", vbOKOnly, app.ProductName
        End If
      End If
    
      ' NPG20080414 Fault 13023
      ' Set objTBColumnPrivileges = Nothing
    End If
    
    If fOK Then
      ' Get the number of training booking records for the current course.
      sSQL = "SELECT COUNT(id) AS recCount" & _
        " FROM " & objTBTable.RealSource & _
        " WHERE id_" & Trim(Str(glngCourseTableID)) & " = " & Trim(Str(mlngRecordID))
          
      If gfCourseTransferProvisionals Then
        sSQL = sSQL & " AND (LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'" & _
          " OR LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'P')"
      Else
        sSQL = sSQL & " AND LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'"
      End If
      
      Set rsInfo = datGeneral.GetRecords(sSQL)
      fTBRecordsExists = (rsInfo!recCount > 0)
      rsInfo.Close
      Set rsInfo = Nothing
      
      If fTBRecordsExists Then
        ' Only ask the user if they want to transfer booking to another course if the current course has some bookings.
        'NHRD15012007 Fault 3905
        iUserChoice = COAMsgBox("Transfer bookings to another course ?", vbYesNo + vbQuestion, app.ProductName)
        'iUserChoice = COAMsgBox("Transfer bookings to another course ?", vbYesNoCancel + vbQuestion, App.ProductName)
      Else
        iUserChoice = vbNo
      End If
    
      fOK = (iUserChoice <> vbCancel)
    End If
    
    If fOK Then
      ' Check that the Cancellation Date column can be updated in the current view of the Course table.
      If mobjTableView.IsTable Then
        Set objColumns = GetColumnPrivileges(mobjTableView.TableName)
      Else
        Set objColumns = GetColumnPrivileges(mobjTableView.ViewName)
      End If
  
      fOK = objColumns.IsValid(gsCourseTitleColumnName)
      If Not fOK Then
        COAMsgBox "The '" & gsCourseTitleColumnName & "' column is not in your current view.", vbOKOnly, app.ProductName
      End If
  
      If fOK Then
        fOK = objColumns.Item(gsCourseTitleColumnName).AllowSelect
        If Not fOK Then
          COAMsgBox "You do not have 'read' permission on the '" & gsCourseTitleColumnName & "' column.", vbOKOnly, app.ProductName
        End If
      End If
  
      If fOK Then
        fOK = objColumns.IsValid(gsCourseCancelDateColumnName)
        If Not fOK Then
          COAMsgBox "The '" & gsCourseCancelDateColumnName & "' column is not in your current view.", vbOKOnly, app.ProductName
        End If
      End If
  
      If fOK Then
        fOK = objColumns.Item(gsCourseCancelDateColumnName).AllowUpdate
        If Not fOK Then
          COAMsgBox "You do not have 'edit' permission on the '" & gsCourseCancelDateColumnName & "' column.", vbOKOnly, app.ProductName
        End If
      End If
  
      If fOK Then
        If Len(gsCourseCancelledByColumnName) > 0 Then
          fOK = objColumns.IsValid(gsCourseCancelledByColumnName)
          If Not fOK Then
            COAMsgBox "The '" & gsCourseCancelledByColumnName & "' column is not in your current view.", vbOKOnly, app.ProductName
          End If
  
          If fOK Then
            fOK = objColumns.Item(gsCourseCancelledByColumnName).AllowUpdate
            If Not fOK Then
              COAMsgBox "You do not have 'edit' permission on the '" & gsCourseCancelledByColumnName & "' column.", vbOKOnly, app.ProductName
            End If
          End If
        End If
      End If
  
      Set objColumns = Nothing
  
      If fOK Then
        ' The current user has permission to cancel the courses.
        gADOCon.BeginTrans
        fInTransaction = True
        ' JPD20010828 Fault 2577 - Moved the record update code below as it was
        ' causing errors when the record was being moved out of the current view.
        ' Set the cancellation date and cancelled by fields of the course record.
'        sSQL = "UPDATE " & mobjTableView.RealSource & _
'          " SET " & gsCourseCancelDateColumnName & " = '" & Format(Date, "mm/dd/yyyy") & "'"
'
'        If Len(gsCourseCancelledByColumnName) > 0 Then
'          sSQL = sSQL & ", " & gsCourseCancelledByColumnName & " = '" & gsUserName & "'"
'        End If
'
'        sSQL = sSQL & " WHERE id = " & Trim(Str(mlngRecordID))
'
'        Screen.MousePointer = vbHourglass
'        fOK = datGeneral.UpdateTableRecord(sSQL, mobjTableView.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
'
'        If Not fOK Then
'          gADOCon.RollbackTrans
'          fInTransaction = False
'        End If
      End If
        
      If fOK And (iUserChoice = vbYes) Then
        ' Transfer bookings to another course if required.
        fTransferred = False
          
        ' Get the current course title.
        sSQL = "SELECT " & gsCourseTitleColumnName & _
          " FROM " & mobjTableView.RealSource & _
          " WHERE id = " & Trim(Str(mlngRecordID))
        Set rsInfo = datGeneral.GetRecords(sSQL)
        fOK = Not (rsInfo.EOF And rsInfo.BOF)
        If fOK Then
          sCourseTitle = IIf(IsNull(rsInfo.Fields(gsCourseTitleColumnName)), "", rsInfo.Fields(gsCourseTitleColumnName))
        End If
        rsInfo.Close
        Set rsInfo = Nothing
          
        If fOK Then
          Set frmCourseSelection = New frmTransferCourseBookings
          With frmCourseSelection
            If .Initialise(mlngRecordID, sCourseTitle) Then
              Screen.MousePointer = vbDefault
              .Show vbModal
              Screen.MousePointer = vbHourglass
  
              fOK = Not .ErrorTransferring
              fTransferred = Not .Cancelled
            End If
          End With
          Unload frmCourseSelection
          Set frmCourseSelection = Nothing
        End If
      
        If Not fOK Then
          gADOCon.RollbackTrans
          fInTransaction = False
        End If
      End If
        
      'JPD20010828 Fault 2577 - Code moved down from above to avoid errors when
      ' the course record is updated and henceforth out of the current view.
      If fOK Then
        ' Set the cancellation date and cancelled by fields of the course record.
        sSQL = "UPDATE " & mobjTableView.RealSource & _
          " SET " & gsCourseCancelDateColumnName & " = '" & Replace(Format(Date, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
        
        If Len(gsCourseCancelledByColumnName) > 0 Then
          'MH20060303 Fault 10871
          'sSQL = sSQL & ", " & gsCourseCancelledByColumnName & " = '" & gsUserName & "'"
          sSQL = sSQL & ", " & gsCourseCancelledByColumnName & " = '" & datGeneral.UserNameForSQL & "'"
        End If
        
        sSQL = sSQL & " WHERE id = " & Trim(Str(mlngRecordID))
        
        Screen.MousePointer = vbHourglass
        'MH20031002 Fault 7082 Reference Property instead of object to trap errors
        'fOK = datGeneral.UpdateTableRecord(sSQL, mobjTableView.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
        fOK = datGeneral.UpdateTableRecord(sSQL, Me.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
        
        If Not fOK Then
          gADOCon.RollbackTrans
          fInTransaction = False
        End If
      End If
      
      If fOK Then
        ' Change the Cancellation Date of the existing bookings.
        If Len(gsTrainBookCancelDateColumnName) > 0 Then
          sSQL = "UPDATE " & objTBTable.RealSource & _
            " SET " & gsTrainBookCancelDateColumnName & " = '" & Replace(Format(Date, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & _
            " WHERE id_" & Trim(Str(glngCourseTableID)) & " = " & Trim(Str(mlngRecordID))

          If gfCourseTransferProvisionals Then
            sSQL = sSQL & " AND (LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'" & _
              " OR LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'P')"
          Else
            sSQL = sSQL & " AND LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'"
          End If

          sErrorMsg = ""
          fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
          If Not fOK Then
            Screen.MousePointer = vbDefault
            COAMsgBox "Unable to update the Training Booking records." & vbNewLine & vbNewLine & sErrorMsg, vbOKOnly, app.ProductName
            Screen.MousePointer = vbHourglass
          End If

          If Not fOK Then
            gADOCon.RollbackTrans
            fInTransaction = False
          End If
        End If
      End If
              
      ' If the bookings have not been transferred then prompt the user if they
      ' want waiting list entries created.
      If fOK And _
        (fTBRecordsExists) And _
        (((iUserChoice = vbYes) And (Not fTransferred)) Or (iUserChoice = vbNo)) Then
        ' Check that the user has permission to add records to the Waiting List table.
        Set objWLTable = gcoTablePrivileges.FindTableID(glngWaitListTableID)
        fOK = Not objWLTable Is Nothing
        If fOK Then
          fOK = objWLTable.AllowInsert
        End If
          
        If fOK Then
          Set objWLColumnPrivileges = GetColumnPrivileges(objWLTable.TableName)
          fOK = objWLColumnPrivileges.Item(gsWaitListCourseTitleColumnName).AllowUpdate
          ' Set objWLColumnPrivileges = Nothing
        End If
          
'''        If fOK Then
'''          Set objTBColumnprivileges = GetColumnPrivileges(objTBTable.TableName)
'''          fOK = objTBColumnprivileges.Item(gsTrainBookCourseTitleName).AllowSelect
'''          Set objTBColumnprivileges = Nothing
'''        End If
        
        If fOK Then
          Screen.MousePointer = vbDefault
          iUserChoice = COAMsgBox("Create waiting list entries for the cancelled bookings ?", vbYesNo + vbQuestion, app.ProductName)
          Screen.MousePointer = vbHourglass
          
          If iUserChoice = vbYes Then
          
          
          
          
          'NPG20080414 Fault 13023
          ' Initialise the string for transfering info from the Training Booking
          ' table back to the Waiting List table.
          ReDim asAddedColumns(1)
          asAddedColumns(1) = UCase(Trim(gsWaitListCourseTitleColumnName))
          sColumnList = gsWaitListCourseTitleColumnName & _
            ", id_" & Trim(Str(glngEmployeeTableID))
          sValueList = "'" & Replace(mrsRecords.Fields(gsCourseTitleColumnName), "'", "''") & "'" & _
            ", id_" & Trim(Str(glngEmployeeTableID))
          
          alngRelatedColumns = RelatedColumns
          
          
          ' NPG20080905 Fault 13023
          ' For iLoop = 1 To UBound(alngRelatedColumns)
          For iLoop = 1 To UBound(alngRelatedColumns, 2)
            Set objTBColumn = objTBColumnPrivileges.FindColumnID(alngRelatedColumns(1, iLoop))
            Set objWLColumn = objWLColumnPrivileges.FindColumnID(alngRelatedColumns(2, iLoop))

            fOK = Not objTBColumn Is Nothing
            If Not fOK Then
              COAMsgBox "Unable to find all related columns in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
              Exit For
            Else
              fOK = objTBColumn.AllowSelect
              If Not fOK Then
                COAMsgBox "You do not have 'read' permission on the '" & objTBColumn.ColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
                Exit For
              End If
            End If
            
            If fOK Then
              fOK = Not objWLColumn Is Nothing
              If Not fOK Then
                COAMsgBox "Unable to find all related columns in the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
                Exit For
              Else
                fOK = objWLColumn.AllowUpdate
                If Not fOK Then
                  COAMsgBox "You do not have 'edit' permission on the '" & objWLColumn.ColumnName & "' column in the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
                  Exit For
                End If
              End If
            End If
            
            If fOK Then
              ' Check that the Training Booking column has not already been added to the 'insert' string.
              fFound = False
              For iNextIndex = 1 To UBound(asAddedColumns)
                If UCase(Trim(objWLColumn.ColumnName)) = asAddedColumns(iNextIndex) Then
                  fFound = True
                  Exit For
                End If
              Next iNextIndex
            
              If Not fFound Then
                ' The current WL column is not in the 'insert' string so add it now,
                ' and add it to the array of added columns.
                sColumnList = sColumnList & _
                  ", " & objWLColumn.ColumnName
              
                iNextIndex = UBound(asAddedColumns) + 1
                ReDim Preserve asAddedColumns(iNextIndex)
                asAddedColumns(iNextIndex) = UCase(Trim(objWLColumn.ColumnName))
                
                sValueList = sValueList & _
                  ", " & objTBColumn.ColumnName
              End If
            End If
            
            Set objTBColumn = Nothing
            Set objWLColumn = Nothing
          Next iLoop

          If fOK Then
            ' Validate the required Waiting List table parameters.
            ' Check that the user has permission to insert records from the Waiting List table.
            fOK = objWLTable.AllowInsert
            If Not fOK Then
              COAMsgBox "You do not have 'new' permission on the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
            End If
          End If
          
          If fOK Then
            ' Check that the user has permission to see the Waiting List Course Title column.
            fOK = objWLColumnPrivileges.Item(gsWaitListCourseTitleColumnName).AllowUpdate
            If Not fOK Then
              COAMsgBox "You do not have 'edit' permission on the '" & gsWaitListCourseTitleColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
            End If
          End If
            
            'NPG20080414 Fault 13027
            'NPG20080422 Fault 13122 - Added 'AND id_1 > 0"
            sSQL = "INSERT INTO " & objWLTable.RealSource & _
              " (" & sColumnList & ")" & _
              " (SELECT " & sValueList & _
              " FROM " & objTBTable.RealSource & _
              " WHERE id_" & glngCourseTableID & " = " & Trim(Str(mlngRecordID)) & _
              " AND id_" & Trim(Str(glngEmployeeTableID)) & " > 0" & _
              " AND '" & Replace(mrsRecords.Fields(gsCourseTitleColumnName), "'", "''") & "' NOT IN (SELECT " & objWLTable.RealSource & "." & gsWaitListCourseTitleColumnName & _
              " FROM " & objWLTable.RealSource & " WHERE " & objWLTable.RealSource & ".id_" & Trim(Str(glngEmployeeTableID)) & _
              " = " & objTBTable.RealSource & ".id_" & Trim(Str(glngEmployeeTableID)) & ")"

'            ' Transfer the booking to the employee's waiting list.
'            sSQL = "INSERT INTO " & objWLTable.RealSource & _
'              " (" & gsWaitListCourseTitleColumnName & "," & _
'              " id_" & Trim(Str(glngEmployeeTableID)) & ")" & _
'              " (SELECT '" & mrsRecords.Fields(gsCourseTitleColumnName) & "', " & _
'              "id_" & Trim(Str(glngEmployeeTableID)) & _
'              " FROM " & objTBTable.RealSource & _
'              " WHERE id_" & glngCourseTableID & " = " & Trim(Str(mlngRecordID)) & _
'              " AND '" & mrsRecords.Fields(gsCourseTitleColumnName) & "' NOT IN (SELECT " & objWLTable.RealSource & "." & gsWaitListCourseTitleColumnName & _
'              " FROM " & objWLTable.RealSource & " WHERE " & objWLTable.RealSource & ".id_" & Trim(Str(glngEmployeeTableID)) & _
'              " = " & objTBTable.RealSource & ".id_" & Trim(Str(glngEmployeeTableID)) & ")"
          
            If gfCourseTransferProvisionals Then
              ' JPD20021126 Fault 4814
              sSQL = sSQL & " AND (LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'" & _
                " OR LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'P'))"
            Else
              sSQL = sSQL & " AND LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B')"
            End If
            
            sErrorMsg = ""
            fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
            
            If Not fOK Then
              Screen.MousePointer = vbDefault
              COAMsgBox "Unable to create waiting list records." & vbNewLine & vbNewLine & sErrorMsg, vbOKOnly, app.ProductName
              Screen.MousePointer = vbHourglass
            End If
          End If
        End If
      
        If Not fOK Then
          gADOCon.RollbackTrans
          fInTransaction = False
        End If
      End If
          
      If fOK Then
        ' Change the status of the existing bookings to be 'CC'.
        sSQL = "UPDATE " & objTBTable.RealSource & _
          " SET " & gsTrainBookStatusColumnName & IIf(gfTrainBookStatus_CC, " = 'CC'", " = 'C'")

''        If Len(gsTrainBookCancelDateColumnName) > 0 Then
''          sSQL = sSQL & _
''            ", " & gsTrainBookCancelDateColumnName & " = '" & Replace(Format(Date, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
''        End If

        sSQL = sSQL & _
          " WHERE id_" & Trim(Str(glngCourseTableID)) & " = " & Trim(Str(mlngRecordID))

        If gfCourseTransferProvisionals Then
          sSQL = sSQL & " AND (LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'" & _
            " OR LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'P')"
        Else
          sSQL = sSQL & " AND LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'"
        End If

        sErrorMsg = ""
        fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
        If Not fOK Then
          Screen.MousePointer = vbDefault
          COAMsgBox "Unable to update the Training Booking records." & vbNewLine & vbNewLine & sErrorMsg, vbOKOnly, app.ProductName
          Screen.MousePointer = vbHourglass
        End If

        If Not fOK Then
          gADOCon.RollbackTrans
          fInTransaction = False
        End If
      End If
        
      ' Refresh the record editing screen and children.
      mfDataChanged = False
      ' JPD20021007 Fault 4498
      ReDim malngChangedOLEPhotos(0)
      
      Requery False

      'JPD 20050915 Fault 10351 - Filter/view checks now made in the Requery method (called in the preceding line)
      ''JPD20010828 Fault 2577 - Check if the course record is still in the recordset.
      'LocateRecord lngCurrentRecordID
      'If mrsRecords!ID <> lngCurrentRecordID Then
      '  If Filtered Then
      '    COAMsgBox "The record saved does not satisfy the current filter.", vbExclamation, App.ProductName
      '  'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '  'ElseIf mobjTableView.ViewID > 0 Then
      '  ElseIf Me.ViewID > 0 Then
      '    COAMsgBox "The record saved is no longer in the current view.", vbExclamation, App.ProductName
      '  End If
      'End If
      
      frmMain.RefreshMainForm Me
    End If
  End If
  
TidyUpAndExit:
  If fInTransaction Then
    If fOK Then
      gADOCon.CommitTrans
      objEmail.SendImmediateEmails
    Else
      gADOCon.RollbackTrans
    End If
    fInTransaction = False
  End If

  Screen.MousePointer = vbDefault
  Exit Sub

ErrorTrap:
  fOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit
    
End Sub


Public Function ReleaseFindWindow()
  ' Release the find window.
  Set mfrmFind = Nothing
  
End Function
Public Property Get TableID() As Long
  ' Return the ID of the screen's associated table.
  If mobjTableView Is Nothing Then
    TableID = 0
  Else
    TableID = mobjTableView.TableID
  End If
End Property

Private Function OLEType(iColumnID As Integer) As Boolean

  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT OleOnServer FROM ASRSysColumns WHERE ColumnID = " & iColumnID

  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.RecordCount > 0 Then
    If Not IsNull(rsTemp!OleOnServer) Then
      OLEType = rsTemp!OleOnServer
    Else
      OLEType = True
    End If
  Else
    OLEType = True
  End If
  
  rsTemp.Close
  
End Function



Public Function Find() As Boolean
  ' Find's a record using the frmFind2 form.
  On Error GoTo Err_Trap

  Find = False
  
  If Not SaveChanges Then
    Screen.MousePointer = vbDefault
    Exit Function
  End If
  
  'MH20040223 Fault 8126
  If mrsRecords.State = adStateClosed Then
    Exit Function
  End If


  If Not (mrsRecords.BOF And mrsRecords.EOF) Then
    If (mrsRecords.EditMode = adEditAdd) And _
      (Not Database.Validation) Then
      Exit Function
    End If
    
    If (mrsRecords.BOF And mrsRecords.EOF) Then
      AddNew
      Exit Function
    End If
  End If
  
  ' JPD20020924
  Screen.MousePointer = vbHourglass
  
  ' See if this window already has a find window
  ' if it doesn't then create a new one else
  ' just show the existing one
  If mfrmFind Is Nothing Then
    If mfrmSummary Is Nothing Then
      Screen.MousePointer = vbHourglass
      
      Set mfrmFind = New frmFind2
      
      With mfrmFind
        
        .Visible = False
        .CurrentRecordID = mlngRecordID
        
        If Not .FindStartFromPrimary(mobjTableView, mlngOrderID, Me, True) Then
          Unload mfrmFind
        End If

      End With
    Else
      mfrmSummary.Visible = True
      'TM20020719 Fault 4167
      mfrmSummary.UpdateSummaryWindow
      
      ' JPD20020920 Fault 4422
      mfrmSummary.CurrentRecordID = mlngRecordID
      mfrmSummary.SetCurrentRecord
      
      mfrmSummary.SetFocus
    End If
  Else
    With mfrmFind
      
      'MH20010523
      '.Show
      '.SetFocus
      .CurrentRecordID = mlngRecordID
      If Not .FindStartFromPrimary(mobjTableView, mlngOrderID, Me, True) Then
        Unload Me
      Else
        .Show
        .SetFocus
      End If
    
    End With
  End If
  
  ' Set the loading value to false to indicate that this form has loaded correctly.
  ' N.B. for other kinds of forms this gets set in the form activate event
  mfLoading = False
  
  ' Initialse the flag indicating if the find window should be updated.
  Screen.MousePointer = vbDefault
  
  Exit Function
  
Err_Trap:
  COAMsgBox Err.Description & " - Find", vbCritical
  
  ' JPD20020924
  Screen.MousePointer = vbDefault
  
End Function


Public Sub MoveNext()
  ' Moves to the NEXT record in the recordset.
  
  ' Save changes if required.
  If SaveChanges Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If
    
    ' Move to the next record.
    With mrsRecords
      
      If (Not .EOF) Then .MoveNext
      If .EOF Then
        If Not RefreshRecordset Then
          Exit Sub
        End If
          
        ' There are records in the refreshed recordset. Move to the last record.
        If .EditMode <> adEditAdd Then
          .MoveLast
        End If
      End If
    End With
  
    ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
    ' UpdateChildren call as the summary fields that are updated in
    ' UpdateChildren may be dependent on the parent recordset being
    ' refreshed first (in UpdateParentWindow).
    UpdateParentWindow
    UpdateControls
    UpdateChildren

    ' JPD20021126 Fault 4676
    If mrsRecords.State = adStateClosed Then
      Exit Sub
    End If

    UpdateFindWindow
'    UpdateParentWindow

    ' Highlight the currently selected control
    FocusCurrentControl

    DoEvents
    If Me.Visible Then Me.SetFocus
    frmMain.RefreshMainForm Me
  End If
    
End Sub


Public Sub UpdateChildren(Optional pvCallingFormID As Variant)
  ' Update any screens that hang off the current screen
  ' ie. find windows and histories of the current screen.
  On Error GoTo Err_Trap

  Dim fGoodChildForm As Boolean
  Dim frmTemp As Form

  ' Update the summary window if there is one.
  If Not mfrmSummary Is Nothing Then
    mfrmSummary.UpdateSummaryWindow
  End If

  ' JPD20021126 Fault 4676
  If mrsRecords.State = adStateClosed Then
    Exit Sub
  End If

  ' See if we have any screens that hang off this screen.
  For Each frmTemp In Forms
    With frmTemp
      If .Name = "frmRecEdit4" Then
        If (.ParentFormID = mlngFormID) Then
          fGoodChildForm = IsMissing(pvCallingFormID)
          If Not fGoodChildForm Then
            fGoodChildForm = (.FormID <> CLng(pvCallingFormID))
          End If
          
          If fGoodChildForm Then
            ' See if the parent is disabled or in edit mode.
            ' If it is then disable the child.
            If mrsRecords.State = adStateClosed Then
              .Enabled = False
            Else
              ' JPD20021209 Fault 4863
              If (Me.Enabled = False) Then 'Or (Not mrsRecords.EditMode = adEditNone) Then
                .Enabled = False
              Else
                .Enabled = True
              End If
              
              .UpdateChildRecords
              .ClearEmbeddedStreams
            
              ' JPD20021126 Fault 4676
              If mrsRecords.State = adStateClosed Then
                Exit Sub
              End If
            
              ' JPD20021104 Fault 4692
              EnableActiveBar .ActiveBar1, False
            End If
          End If
        End If

      'MH20010824 Fault 2449
      ElseIf .Name = "frmFind2" Then
        If (.ParentFormID = mlngFormID) Then
          .UpdateSummaryWindow
        
          ' JPD20021126 Fault 4676
          If mrsRecords.State = adStateClosed Then
            Exit Sub
          End If
        
          ' JPD20021104 Fault 4692
          EnableActiveBar .ActiveBar1, False
        End If

      End If
    End With
  Next frmTemp
  Set frmTemp = Nothing

  Exit Sub
  
Err_Trap:
  COAMsgBox Err.Description & " - UpdateChildren", vbCritical

End Sub


Public Sub UpdateChildRecords()
  
  On Error GoTo Err_Trap
  
  ' Refresh the resultset for a history
  
  ' Close the resultset if there is one
  If Not mrsRecords Is Nothing Then
    Set mrsRecords = Nothing
  End If
  
  GetRecords

  If frmMain.ActiveForm Is Me Then
    frmMain.RefreshMainForm Me
  End If

  'TM20020107 Fault 1379
'  ' If no records match the filter, then clear it.
'  If (mrsRecords.BOF And mrsRecords.EOF) And _
'    Filtered Then
'    COAMsgBox "No records match the current filter." & vbNewLine & _
'      "No filter is applied.", vbInformation + vbOKOnly, App.ProductName
'    ReDim mavFilterCriteria(3, 0)
'    mrsRecords.Close
'    Set mrsRecords = Nothing
'    GetRecords
'  End If

  If mrsRecords.State <> adStateClosed Then
    ' Check if the refreshed recordset is still empty.
    If mrsRecords.BOF And mrsRecords.EOF Then
      ' Create a new record if permitted.
      If mobjTableView.AllowInsert Then
        mrsRecords.AddNew
      Else
        ' JPD20030305 Fault 5079 (again)
        ' JPD20030225 Fault 5079
        If Me.Visible Then
          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'COAMsgBox "This " & IIf(mobjTableView.ViewID > 0, "view", "table") & " is empty" & _
            " and you do not have 'new' permission on it.", vbExclamation, "Security"
          COAMsgBox "This " & IIf(Me.ViewID > 0, "view", "table") & " is empty" & _
            " and you do not have 'new' permission on it.", vbExclamation, "Security"
        End If
        RefreshFormCaption
        If Not (mfrmSummary Is Nothing) Then
          mfrmSummary.Visible = True
        End If
        ShowHistorySummary
        mfrmSummary.UpdateSummaryWindow
        Me.Visible = False
        Unload Me
        
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
    Else
      ' Select the originally selected record if it is still in
      ' the recordset.
      LocateRecord mlngRecordID
    End If
  
    UpdateControls
    UpdateChildren
  End If
  
  Exit Sub
  
Err_Trap:
  COAMsgBox Err.Description & " - UpdateChildRecords", vbCritical

End Sub

Private Sub UpdateFindWindow()
  
  If Not mfrmFind Is Nothing Then
    mfrmFind.UpdateFindWindow
  
    ' JPD20021104 Fault 4692
    EnableActiveBar mfrmFind.ActiveBar1, False
  End If
  
End Sub
Public Sub MoveLast()
  ' Moves to the LAST record in the recordset.
  
  ' Save changes if required.
  If SaveChanges Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If
    
    ' Move to the last record.
    With mrsRecords
      If (Not .EOF) Then .MoveLast
      If .EOF Then
        If Not RefreshRecordset Then
          Exit Sub
        End If
          
        ' There are records in the refreshed recordset. Move to the last record.
        If .EditMode <> adEditAdd Then
          .MoveLast
        End If
      End If
    End With
  
    ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
    ' UpdateChildren call as the summary fields that are updated in
    ' UpdateChildren may be dependent on the parent recordset being
    ' refreshed first (in UpdateParentWindow).
    UpdateParentWindow
    UpdateControls
    UpdateChildren
    
    ' JPD20021126 Fault 4676
    If mrsRecords.State = adStateClosed Then
      Exit Sub
    End If
        
    UpdateFindWindow
'    UpdateParentWindow

    ' Highlight the currently selected control
    FocusCurrentControl
    
    frmMain.RefreshMainForm Me
  End If

End Sub

Public Sub MoveFirst()
  ' Moves to the FIRST record in the recordset.
  
  ' Save changes if required.
  If SaveChanges Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If
    
    ' Move to the first record.
    With mrsRecords
      If (Not .BOF) Then .MoveFirst
      If .BOF Then
        If Not RefreshRecordset Then
          Exit Sub
        End If
        
        ' There are records in the refreshed recordset. Move to the first record.
        If .EditMode <> adEditAdd Then
          .MoveFirst
        End If
      End If
    End With
  
    ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
    ' UpdateChildren call as the summary fields that are updated in
    ' UpdateChildren may be dependent on the parent recordset being
    ' refreshed first (in UpdateParentWindow).
    UpdateParentWindow
    UpdateControls
    UpdateChildren
    
    ' JPD20021126 Fault 4676
    If mrsRecords.State = adStateClosed Then
      Exit Sub
    End If
    
    UpdateFindWindow
'    UpdateParentWindow

    ' Highlight the currently selected control
    FocusCurrentControl
  
    frmMain.RefreshMainForm Me
  End If
        
End Sub


Public Sub MovePrevious()
  ' Moves to the PREVIOUS record in the recordset.
  Dim fAdding As Boolean
  
  fAdding = (mrsRecords.EditMode = adEditAdd)
  
  ' Save changes if required.
  If SaveChanges(False) Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If
    
    
    ' Move to the previous record.
    With mrsRecords
      If Not fAdding Then
        If (Not .BOF) Then .MovePrevious
      End If
      
      If .BOF Then
        If Not RefreshRecordset Then
          Exit Sub
        End If
          
        ' There are records in the refreshed recordset. Move to the first record.
        If .EditMode <> adEditAdd Then
          .MoveFirst
        End If
      End If
    End With
  
    
    ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
    ' UpdateChildren call as the summary fields that are updated in
    ' UpdateChildren may be dependent on the parent recordset being
    ' refreshed first (in UpdateParentWindow).
    UpdateParentWindow
    UpdateControls
    UpdateChildren
    
    ' JPD20021126 Fault 4676
    If mrsRecords.State = adStateClosed Then
      Exit Sub
    End If
    
    UpdateFindWindow
'    UpdateParentWindow

    ' Highlight the currently selected control
    FocusCurrentControl
  
    frmMain.RefreshMainForm Me
  End If
  

End Sub

Public Property Get AllowInsert() As Boolean
  ' Return whether or not the user can insert records.
  If mobjTableView Is Nothing Then
    AllowInsert = False
  Else
    AllowInsert = mobjTableView.AllowInsert
  End If
End Property


Public Sub AddNew()

  Dim iCount As Integer

  ' Add a new record.
  
  ' Save changes if required.
  If SaveChanges Then
    'If Not Database.Validation Then
    If Not Database.Validation Or mrsRecords Is Nothing Then
      Exit Sub
    End If

    ' Check the user has permission to add new records to the table/view.
    If mobjTableView.AllowInsert Then
      mrsRecords.CancelUpdate
      mrsRecords.AddNew
    
      mbDisableAURefresh = True
      ' JPD20030311 Fault 5142
      mfAddingNewInProgress = True
      
      ' JDM - Fault 9262 - Clear streams for new records
      ClearEmbeddedStreams
      
      UpdateControls
      UpdateChildren
      mfAddingNewInProgress = False
    Else
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      'COAMsgBox "You do not have 'new' permission on this " & IIf(mobjTableView.ViewID > 0, "view.", "table."), vbExclamation, "Security"
      COAMsgBox "You do not have 'new' permission on this " & IIf(Me.ViewID > 0, "view.", "table."), vbExclamation, "Security"
      Exit Sub
    End If
  End If

  ' Highlight the first object
  FocusFirstControl

  ' Refresh the main menu.
  frmMain.RefreshMainForm Me

End Sub
Public Function SaveChanges(Optional pfUpdateControls As Variant, _
  Optional pfCancelUpdate As Variant, _
  Optional pfDeactivating As Variant) As Boolean
  ' Prompt the user if they wish to save the changes if they have made any.
  Dim iResult As Integer
  Dim strSaveCaption As String
  ' JPD20021206 Fault 4854
  If mfSavingInProgress Then
    SaveChanges = True
    Exit Function
  End If

  ' JPD20030311 Fault 5142
  If mfAddingNewInProgress Then
    SaveChanges = True
    Exit Function
  End If

  mfCancelled = False
  iResult = vbYes
  
  If IsMissing(pfUpdateControls) Then pfUpdateControls = True
  If IsMissing(pfCancelUpdate) Then pfCancelUpdate = True
  If IsMissing(pfDeactivating) Then pfDeactivating = False
  
  ' Check if the current record has been modified.
  If RecordChanged(pfUpdateControls, pfCancelUpdate) Or mbResendingToAccord Then
    ' prompt the user to save changes if the record has been changed.
    If mfDataChanged Then
      iResult = COAMsgBox("Record changed, do you wish to save changes?", _
        vbYesNoCancel + vbQuestion, Me.Caption)
    Else
      iResult = vbYes
    End If

    Select Case iResult
      Case vbYes
           
        If mbResendingToAccord Then
          strSaveCaption = "Transfer to Payroll..."
        Else
          strSaveCaption = "Save changes..."
        End If
        
        ' Save the changes to the server.
        If Not UpdateWithAVI(pfDeactivating, strSaveCaption) Then
          iResult = vbCancel
          mfCancelled = True
          Database.Validation = False
        End If

      Case vbNo
        ' Cancel the changes and do not save them.
        mfDataChanged = False
        ' JPD20021007 Fault 4498
        ReDim malngChangedOLEPhotos(0)
        
        If mrsRecords.EditMode <> adEditNone Then
          mrsRecords.CancelUpdate
        End If
        
        ' JPD 14/3/01 Fault 1998 - AddNew record if BOF or EOF.
        If (mrsRecords.BOF And mrsRecords.EOF) Then
          If mobjTableView.AllowInsert Then
            mrsRecords.AddNew
          Else
            ' The refreshed recordset is empty and a new record cannot be created.
            ' Kill the record edit form.
            'MH20031002 Fault 7082 Reference Property instead of object to trap errors
            'COAMsgBox "This " & IIf(mobjTableView.ViewID > 0, "view", "table") & " is empty" & _
              " and you do not have 'new' permission on it.", vbExclamation, "Security"
            COAMsgBox "This " & IIf(Me.ViewID > 0, "view", "table") & " is empty" & _
              " and you do not have 'new' permission on it.", vbExclamation, "Security"
            Screen.MousePointer = vbDefault
            Unload Me
            SaveChanges = False
            Exit Function
           End If
        End If

        ' Update the screen controls.
        UpdateControls True
        ' JPD20021209 Fault 4863
        UpdateChildren
        
        If mfTableEntry Then
          mfLeaveLookup = True
        End If
        Database.Validation = True

      Case Else
        ' Do not save changes, and cancel the operation that called this function.
        mfCancelled = True
        Database.Validation = False
    End Select
       
    'If Not mfCancelled Then
    If Not mfCancelled And Not mfUnloading Then
      frmMain.RefreshMainForm Me
    End If
  Else
    Database.Validation = True
  End If

  SaveChanges = (iResult <> vbCancel)
  
End Function
'Private Function RecordChanged(Optional pfUpdateControls As Variant, _
  Optional pfCancelUpdate As Variant) As Boolean
Public Function RecordChanged(Optional pfUpdateControls As Variant, _
  Optional pfCancelUpdate As Variant) As Boolean
  ' Check to see if the record has changed by comparing it to the current record in the database.
  On Error GoTo Err_Trap
    
  ' Do nothing if the form is not visible.
  If Not Me.Visible Then
    Exit Function
  End If


  If IsMissing(pfUpdateControls) Then pfUpdateControls = True
  If IsMissing(pfCancelUpdate) Then pfCancelUpdate = True
  
  'JPD 20031009 Fault 7080
  If mrsRecords.State = adStateClosed Then
    RecordChanged = False
  Else
    If (mrsRecords.BOF And mrsRecords.EOF) Then
      ' Record has not changed if there is no record.
      RecordChanged = False
    Else
      If mrsRecords.EditMode <> adEditAdd Then
        ' Check the change flag if we are just in edit mode.
        RecordChanged = mfDataChanged
      Else
        ' Check the change flag if we are in add mode.
        RecordChanged = mfDataChanged
        
        If Not mfDataChanged Then
          If (mrsRecords.EditMode <> adEditNone) And pfCancelUpdate Then
            mrsRecords.CancelUpdate
          
            ' Check if the refreshed recordset is empty and filtered.
            'TM20020829 Fault 4168 & 4232 - only clear the filter if it is the parent, (Me.ParentTableID = 0).
            If (mrsRecords.BOF And mrsRecords.EOF) And Filtered And (Me.ParentTableID = 0) Then
              ' Clear the filter.
              COAMsgBox "No records match the current filter." & vbNewLine & _
                "The filter has been cleared.", vbInformation + vbOKOnly, app.ProductName
              ReDim mavFilterCriteria(3, 0)
              mrsRecords.Close
              Set mrsRecords = Nothing
              GetRecords
            End If
          End If
          
          ' Check if the refreshed recordset is still empty.
          If (mrsRecords.BOF And mrsRecords.EOF) Then
            ' The refreshed recordset is empty. Create a new record if permitted.
            If mobjTableView.AllowInsert Then
              mrsRecords.AddNew
            Else
              ' The refreshed recordset is empty and a new record cannot be created.
              ' Kill the record edit form.
              'MH20031002 Fault 7082 Reference Property instead of object to trap errors
              'COAMsgBox "This " & IIf(mobjTableView.ViewID > 0, "view", "table") & " is empty" & _
                " and you do not have 'new' permission on it.", vbExclamation, "Security"
              COAMsgBox "This " & IIf(Me.ViewID > 0, "view", "table") & " is empty" & _
                " and you do not have 'new' permission on it.", vbExclamation, "Security"
              Screen.MousePointer = vbDefault
              Unload Me
              RecordChanged = False
              Exit Function
             End If
          End If
          
          ' Update the screen controls.
          If pfUpdateControls Then
            UpdateControls
          End If
  
          If mfTableEntry Then
            mfLeaveLookup = True
          End If
        End If
      End If
    End If
  End If
  
  Exit Function
   
Err_Trap:
  COAMsgBox Err.Description & " - RecordChanged", vbCritical
    
End Function

'MH20001109 Fault 981 Pass in parent table and record now !
'Private Sub SetControlDefaults(psTag As String, pobjControl As Control)
Private Sub SetControlDefaults(psTag As String, pobjControl As Control, lngParentTableID As Long, lngParentRecordID As Long)
  ' Set control defaults.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fProgBarVisible As Boolean
  Dim sSQL As String
  Dim sDefaultValue As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim rsInfo As Recordset
  Dim iLoop As Integer
  Dim sFormat As String
  
  fOK = True
  
  'RH 16/08/00 - BUG 632. Only populate the control with the
  '              default value if we have read permision on
  '              the controls column.
  If Not mcolColumnPrivileges.FindColumnID(mobjScreenControls.Item(psTag).ColumnID) Is Nothing Then
    If Not mcolColumnPrivileges.FindColumnID(mobjScreenControls.Item(psTag).ColumnID).AllowSelect Then Exit Sub
  End If


  'MH20060725 Fault 11358
  'Don't calculate default values for columns on the parent table
  'If mobjScreenControls.Item(psTag).DfltValueExprID > 0 Then
  If mobjScreenControls.Item(psTag).DfltValueExprID > 0 And _
     mobjScreenControls.Item(psTag).TableID = Me.TableID Then


    'Check the default expression stored procedure exists.
    sSQL = "SELECT COUNT(*) AS recCount" & _
      " FROM sysobjects" & _
      " WHERE id = object_id(N'sp_ASRDfltExpr_" & Trim(Str(mobjScreenControls.Item(psTag).DfltValueExprID)) & "')" & _
      " AND OBJECTPROPERTY(id, N'IsProcedure') = 1"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    fOK = (rsInfo!recCount > 0)
    rsInfo.Close
    Set rsInfo = Nothing
      
    If fOK Then
      Set cmADO = New ADODB.Command
      With cmADO
        .CommandText = "sp_ASRDfltExpr_" & Trim(Str(mobjScreenControls.Item(psTag).DfltValueExprID))
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = gADOCon
  
        ' Append the result parameter.
        Select Case mobjScreenControls.Item(psTag).DataType
          Case sqlOle ' OLE columns do not have defaults.
            fOK = False
            
          Case sqlBoolean  ' Logic columns
            Set pmADO = .CreateParameter("result", adBoolean, adParamOutput)
          
          Case sqlNumeric   ' Numeric columns
            Set pmADO = .CreateParameter("result", adDouble, adParamOutput, 8)
          
          Case sqlInteger  ' Integer columns
            Set pmADO = .CreateParameter("result", adDouble, adParamOutput, 8)
          
          Case sqlDate
            Set pmADO = .CreateParameter("result", adDate, adParamOutput, 8)
                      
          Case sqlVarChar ' Character columns
            Set pmADO = .CreateParameter("result", adVarChar, adParamOutput, VARCHAR_MAX_Size)
                      
          Case sqlVarBinary ' Photo columns do not have defaults.
            fOK = False
            
          Case sqlLongVarChar ' Working Pattern columns
            Set pmADO = .CreateParameter("result", adVarChar, adParamOutput, 14)
            
          Case Else
            fOK = False
        End Select
        
        If fOK Then
          .Parameters.Append pmADO

          ' Append the parent table ID parameters.
          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'sSQL = "SELECT parentID" & _
            " FROM ASRSysRelations" & _
            " WHERE childID = " & Trim(Str(mobjTableView.TableID)) & _
            " ORDER BY parentID"
          sSQL = "SELECT parentID" & _
            " FROM ASRSysRelations" & _
            " WHERE childID = " & CStr(Me.TableID) & _
            " ORDER BY parentID"
          Set rsInfo = datGeneral.GetRecords(sSQL)
          Do While Not rsInfo.EOF
            Set pmADO = .CreateParameter("ID_" & Trim(Str(rsInfo!ParentID)), adInteger, adParamInput)
            .Parameters.Append pmADO
            
            'MH20001109 Fault 981 Use parameters which were passed in
            'and not module level variables
            'If rsInfo!ParentID = mlngParentTableID Then
            '  pmADO.Value = mlngParentRecordID
            If rsInfo!ParentID = lngParentTableID Then
              pmADO.Value = lngParentRecordID
            Else
              pmADO.Value = 0
            End If
            
            rsInfo.MoveNext
          Loop

          rsInfo.Close
          Set rsInfo = Nothing

          cmADO.Execute

          Select Case mobjScreenControls.Item(psTag).DataType
            Case sqlBoolean  ' Logic column
              sDefaultValue = IIf(IsNull(.Parameters("result").Value), "FALSE", IIf(.Parameters("result").Value, "TRUE", "FALSE"))
            
            Case sqlNumeric   ' Numeric column
              sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", Trim(Str(.Parameters("result").Value)))
            
            Case sqlInteger  ' Integer column
              sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", Trim(Str(.Parameters("result").Value)))
            
            Case sqlDate ' Date column
              
              sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", Replace(Format(.Parameters("result").Value, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/"))
              
            Case sqlVarChar ' Character column
              sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", .Parameters("result").Value)
                        
            Case sqlLongVarChar ' Working Pattern column
              sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", .Parameters("result").Value)
          End Select
        End If
      End With
      
      Set pmADO = Nothing
      Set cmADO = Nothing
    End If
  End If
  
  If Not fOK Then
    ' Error calculating the default, so reset.
    sDefaultValue = ""
  ElseIf mobjScreenControls.Item(psTag).DfltValueExprID = 0 Then
    ' Straight value as the default.
    sDefaultValue = mobjScreenControls.Item(psTag).DefaultValue
  End If

  With pobjControl
    If TypeOf pobjControl Is TDBText6Ctl.TDBText Then
      ' NB. If the default value is too long for the defined column length,
      ' then it is automatically truncated by the TDBText control.
      .Text = sDefaultValue
    
    ElseIf TypeOf pobjControl Is COA_Image Then
      ' No defaults, just reset the control.
      .Picture = Nothing
      .ASRDataField = ""
    
    ElseIf TypeOf pobjControl Is TDBMask6Ctl.TDBMask Then
      ' If the default value does not match the defined mask then the
      ' error is trapped below, and the control is left with its previous value.
      .Text = sDefaultValue
    
    ElseIf TypeOf pobjControl Is TextBox Then
      ' NB. If the default value is too long for the defined column length,
      ' then it is automatically truncated by the textbox control.
      .Text = sDefaultValue
    
    'JPD 20050302 Fault 9847
    ElseIf (TypeOf pobjControl Is TDBNumberCtrl.TDBNumber) Or _
      (TypeOf pobjControl Is TDBNumber6Ctl.TDBNumber) Then
      If Len(sDefaultValue) > 0 Then
        If mobjScreenControls.Item(psTag).DataType = sqlInteger Then
          
          'MH20010202 Fault 1785
          'Clng causes error for big numbers
          '.Text = Trim(Str(CLng(sDefaultValue)))
          .Text = Trim(sDefaultValue)

        Else
          ' Check that the default value fits into the column control.
          ' First, round any decimals to the columns number of decimals.
          sDefaultValue = Trim(Str(Round(Val(sDefaultValue), mobjScreenControls.Item(psTag).Decimals)))
          
          'MH20010108
          sDefaultValue = datGeneral.ConvertNumberForDisplay(sDefaultValue)

          ' Second, ensure that the rounded value does not exceed the columns size.
          
          
          'MH20010202 Fault 1785
          'Clng causes error for big numbers but good old fashioned VAL does the trick!
          'If Len(Trim(Str(IIf(CLng(sDefaultValue) < 0, (CLng(sDefaultValue) * -1), CLng(sDefaultValue))))) <= mobjScreenControls.Item(psTag).Size Then
          'If Len(Trim(Str(IIf(Val(sDefaultValue) < 0, (Val(sDefaultValue) * -1), Val(sDefaultValue))))) <= mobjScreenControls.Item(psTag).Size Then
          
          'JPD 20050309 - changed number control, so now need to use the value property.
          '.Text = Trim(sDefaultValue)
          .Value = Val(sDefaultValue)
          'End If
        End If
      Else
        .Text = sDefaultValue
      End If
      
    ElseIf TypeOf pobjControl Is XtremeSuiteControls.CheckBox Then
      .Value = IIf(CBool(IIf(sDefaultValue = "", 0, sDefaultValue)) = True, 1, 0)
    
    ElseIf TypeOf pobjControl Is XtremeSuiteControls.ComboBox Then
      ' JPD 10/4/01 Allow a blank default to be selected.
      If Len(sDefaultValue) > 0 Then
        .ListIndex = UI.cboSelect(pobjControl, sDefaultValue)
      Else
        For iLoop = 0 To (.ListCount - 1)
          If Len(.List(iLoop)) = 0 Then
            .ListIndex = iLoop
            Exit For
          End If
        Next iLoop
      End If
      
      ' If the default value is not in the combo,
      ' just select the first combo item.
      If .ListIndex < 0 Then
        If .ListCount > 0 Then
          .ListIndex = 0
        End If
      End If
      
    ElseIf TypeOf pobjControl Is COA_Lookup Then
      If Len(sDefaultValue) > 0 Then
        ' Ensure that the default value does not exceed the lookup column's size.
        If mobjScreenControls.Item(psTag).DataType = sqlVarChar Then
          If mobjScreenControls.Item(psTag).Size < Len(sDefaultValue) Then
            sDefaultValue = Left$(sDefaultValue, mobjScreenControls.Item(psTag).Size)
          End If
          .Text = sDefaultValue
        ElseIf mobjScreenControls.Item(psTag).DataType = sqlNumeric Then
          ' Check that the default value fits into the column control.
          ' First, round any decimals to the columns number of decimals.
          sDefaultValue = Trim(Str(Round(Val(sDefaultValue), mobjScreenControls.Item(psTag).Decimals)))
          
          ' Second, ensure that the rounded value does not exceed the columns size.
          If Len(Trim(Str(IIf(CLng(sDefaultValue) < 0, (CLng(sDefaultValue) * -1), CLng(sDefaultValue))))) <= mobjScreenControls.Item(psTag).Size Then
            'JPD 20050810 Fault 10165
            sFormat = "0"
            If mobjScreenControls.Item(psTag).Use1000Separator Then
              sFormat = "#,0"
            End If
            If mobjScreenControls.Item(psTag).Decimals > 0 Then
              sFormat = sFormat & "." & String(mobjScreenControls.Item(psTag).Decimals, "0")
            End If
                        
            .Text = Format(sDefaultValue, sFormat)
          End If
        ElseIf mobjScreenControls.Item(psTag).DataType = sqlInteger Then
          .Text = Trim(Str(CLng(sDefaultValue)))
        ElseIf mobjScreenControls.Item(psTag).DataType = sqlDate Then
          .Text = IIf(Len(sDefaultValue) > 0, ConvertSQLDateToLocale(sDefaultValue), "")
        ElseIf mobjScreenControls.Item(psTag).DataType = sqlLongVarChar Then
          .Text = Left$(sDefaultValue, 14)
        End If
      Else
        .Text = sDefaultValue
      End If
      
    ElseIf TypeOf pobjControl Is COA_OptionGroup Then
      
      'MH20070228 Fault 11876
      'Setting mvOldValue forces the option group value to change
      'even if a child record is part way through being saved
      '(This overrides a fix in the optiongroup_click event)
      mvOldValue = sDefaultValue
      
      .Text = sDefaultValue
      
      ' If the default value is not in the option group,
      ' just select the first option group item.
      If .Text <> sDefaultValue Then
        If UBound(mobjScreenControls.Item(psTag).ControlValues, 2) >= 0 Then
          .Value = 0
        End If
      End If
      
    ElseIf TypeOf pobjControl Is COA_OLE Then
      ' No defaults, just reset the control.
      '.Delete
      .FileName = vbNullString
      .OLEType = IIf(pobjControl.OLEType = OLE_EMBEDDED, OLE_UNC, pobjControl.OLEType)
      
    ElseIf TypeOf pobjControl Is COA_Spinner Then
      ' Round the default value to the nearest integer.
      ' If the rounded value exceeds the controls min or max
      ' then the control automatically set the value to be the nearest
      ' permissable value.
      If Len(sDefaultValue) > 0 Then
        .Text = Trim(Str(CLng(sDefaultValue)))
      Else
        .Text = sDefaultValue
      End If
      
    ElseIf TypeOf pobjControl Is GTMaskDate.GTMaskDate Then
      .Text = IIf((Len(sDefaultValue) > 0) And IsDate(sDefaultValue), ConvertSQLDateToLocale(sDefaultValue), "")
    
    ElseIf TypeOf pobjControl Is COA_WorkingPattern Then
      ' NB. If the default value is too long for the defined column length,
      ' then it is automatically truncated by the Working Pattern control.
      .Value = sDefaultValue
    
    ElseIf TypeOf pobjControl Is COA_Navigation Then
      .NavigateTo = sDefaultValue
    
    ElseIf TypeOf pobjControl Is COA_ColourSelector Then
      .BackColor = sDefaultValue
    
    End If
  End With

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  If Err.Number = 380 Then
    ' Mask control was being populated with a value that does not fit the mask.
    ' Just ignore it.
    
    'JPD 20051005 Fault 10403
    ' No, don't ignore it. Clear the control.
    If Not pobjControl Is Nothing Then
      If TypeOf pobjControl Is TDBMask6Ctl.TDBMask Then
        pobjControl.Text = ""
      End If
    End If

    Resume Next
  Else
    fProgBarVisible = gobjProgress.Visible
    gobjProgress.Visible = False
    COAMsgBox Err.Description & " - SetControlDefaults", vbExclamation + vbOKOnly, app.ProductName
    gobjProgress.Visible = fProgBarVisible
  End If
  
  Resume TidyUpAndExit
  
End Sub







Private Function GetColumnDefault(plngColumnID As Long) As String
  ' Set control defaults.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sSQL As String
  Dim lngDfltExprID As Long
  Dim sDefaultValue As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim rsInfo As Recordset
  Dim iDataType As SQLDataType

  fOK = True
  sDefaultValue = ""
  lngDfltExprID = 0
  
  If Not mcolColumnPrivileges.FindColumnID(plngColumnID) Is Nothing Then
    If Not mcolColumnPrivileges.FindColumnID(plngColumnID).AllowSelect Then
      GetColumnDefault = sDefaultValue
      Exit Function
    End If
  End If

  sSQL = "SELECT dfltValueExprID, defaultValue, dataType" & _
    " FROM ASRSysColumns" & _
    " WHERE columnID = " & CStr(plngColumnID)
  Set rsInfo = datGeneral.GetRecords(sSQL)
  If (rsInfo.BOF And rsInfo.EOF) Then
    fOK = False
  Else
    lngDfltExprID = IIf(IsNull(rsInfo!DfltValueExprID), 0, rsInfo!DfltValueExprID)
    sDefaultValue = IIf(IsNull(rsInfo!DefaultValue), "", rsInfo!DefaultValue)
    iDataType = IIf(IsNull(rsInfo!DataType), 0, rsInfo!DataType)
  End If
  rsInfo.Close
  Set rsInfo = Nothing
  
  If fOK Then
    If lngDfltExprID > 0 Then
      ' Calculated value as the default.
  
      'Check the default expression stored procedure exists.
      sSQL = "SELECT COUNT(*) AS recCount" & _
        " FROM sysobjects" & _
        " WHERE id = object_id(N'sp_ASRDfltExpr_" & Trim(Str(lngDfltExprID)) & "')" & _
        " AND OBJECTPROPERTY(id, N'IsProcedure') = 1"
      Set rsInfo = datGeneral.GetRecords(sSQL)
      fOK = (rsInfo!recCount > 0)
      rsInfo.Close
      Set rsInfo = Nothing
  
      If fOK Then
        Set cmADO = New ADODB.Command
        With cmADO
          .CommandText = "sp_ASRDfltExpr_" & Trim(Str(lngDfltExprID))
          .CommandType = adCmdStoredProc
          .CommandTimeout = 0
          Set .ActiveConnection = gADOCon
  
          ' Append the result parameter.
          Select Case iDataType
            Case sqlOle ' OLE columns do not have defaults.
              fOK = False
  
            Case sqlBoolean  ' Logic columns
              Set pmADO = .CreateParameter("result", adBoolean, adParamOutput)
  
            Case sqlNumeric   ' Numeric columns
              Set pmADO = .CreateParameter("result", adDouble, adParamOutput, 8)
  
            Case sqlInteger  ' Integer columns
              Set pmADO = .CreateParameter("result", adDouble, adParamOutput, 8)
  
            Case sqlDate
              Set pmADO = .CreateParameter("result", adDate, adParamOutput, 8)
  
            Case sqlVarChar ' Character columns
              Set pmADO = .CreateParameter("result", adLongVarChar, adParamOutput, -1)
  
            Case sqlVarBinary ' Photo columns do not have defaults.
              fOK = False
  
            Case sqlLongVarChar ' Working Pattern columns
              Set pmADO = .CreateParameter("result", adVarChar, adParamOutput, 14)
  
            Case Else
              fOK = False
          End Select
  
          If fOK Then
            .Parameters.Append pmADO
  
            ' Append the parent table ID parameters.
            'MH20031002 Fault 7082 Reference Property instead of object to trap errors
            'sSQL = "SELECT parentID" & _
              " FROM ASRSysRelations" & _
              " WHERE childID = " & Trim(Str(mobjTableView.TableID)) & _
              " ORDER BY parentID"
            sSQL = "SELECT parentID" & _
              " FROM ASRSysRelations" & _
              " WHERE childID = " & CStr(Me.TableID) & _
              " ORDER BY parentID"
            Set rsInfo = datGeneral.GetRecords(sSQL)
            Do While Not rsInfo.EOF
              Set pmADO = .CreateParameter("ID_" & Trim(Str(rsInfo!ParentID)), adInteger, adParamInput)
              .Parameters.Append pmADO
  
              If rsInfo!ParentID = mlngParentTableID Then
                pmADO.Value = mlngParentRecordID
              Else
                pmADO.Value = 0
              End If
  
              rsInfo.MoveNext
            Loop
  
            rsInfo.Close
            Set rsInfo = Nothing
  
            cmADO.Execute
  
            Select Case iDataType
              Case sqlBoolean  ' Logic column
                sDefaultValue = IIf(IsNull(.Parameters("result").Value), "0", IIf(.Parameters("result").Value, "1", "0"))
  
              Case sqlNumeric   ' Numeric column
                sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", Trim(Str(.Parameters("result").Value)))
  
              Case sqlInteger  ' Integer column
                sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", Trim(Str(.Parameters("result").Value)))
  
              Case sqlDate ' Date column
                sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", Replace(Format(.Parameters("result").Value, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/"))
                sDefaultValue = IIf(Len(sDefaultValue) > 0, ConvertSQLDateToLocale(sDefaultValue), "")
  
              Case sqlVarChar ' Character column
                sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", .Parameters("result").Value)
  
              Case sqlLongVarChar ' Working Pattern column
                sDefaultValue = IIf(IsNull(.Parameters("result").Value), "", .Parameters("result").Value)
            End Select
          End If
        End With
  
        Set pmADO = Nothing
        Set cmADO = Nothing
      End If
    Else
      Select Case iDataType
        Case sqlBoolean  ' Logic column
          sDefaultValue = IIf(CBool(IIf(sDefaultValue = "", 0, sDefaultValue)), "1", "0")
  
        Case sqlDate ' Date column
          sDefaultValue = IIf(Len(sDefaultValue) > 0, ConvertSQLDateToLocale(sDefaultValue), "")
      End Select
    End If
  End If

TidyUpAndExit:
  If Not fOK Then
    ' Error calculating the default, so reset.
    sDefaultValue = ""
  End If
  
  GetColumnDefault = sDefaultValue
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Sub RefreshFormCaption()
  ' Refresh the form caption.
  Dim fOK As Boolean
  Dim lngNextIndex As Long
  Dim sCaption As String
  Dim sRecordDescription As String
  Dim frmForm As Form
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim alngScreenIndexes() As Long
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngMinIndex As Long
  Dim lngMaxIndex As Long
  Dim fFound As Boolean

  ' Evaluate the Record Description Expression.
  sRecordDescription = ""
  If (Not (mrsRecords.BOF And mrsRecords.EOF)) Then
    If (mrsRecords.EditMode <> adEditAdd) And _
      (mlngRecDescID > 0) Then
    
      sRecordDescription = EvaluateRecordDescription(mlngRecordID, mlngRecDescID)
      
      If Len(sRecordDescription) > 0 Then
        sRecordDescription = " - " & sRecordDescription
      End If
    End If
  End If

  If miScreenType = screenHistoryTable Then
    ' JDM - Fault 1528 - Modified the captions to make them a little shorter
    Me.Caption = msScreenName & " (" & GetParentFormParameter(mlngParentFormID, "CAPTION") & ")"
    msFindCaption = Me.Caption
    msStatusCaption = msScreenName & " (" & GetParentFormParameter(mlngParentFormID, "STATUSCAPTION") & ")"
    msFindStatusCaption = msStatusCaption
    msFindPrintHeader = msScreenName & " (" & GetParentFormParameter(mlngParentFormID, "PRINTHEADER") & ")"
  Else
    If miScreenType = screenParentView Then
      sCaption = msScreenName & " - View"
      msFindCaption = "Find - " & msScreenName & " (" & Replace(ViewName, "_", " ") & " view)"
      msStatusCaption = msScreenName & " (" & Replace(ViewName, "_", " ") & " view)"
      msFindStatusCaption = msFindCaption
      msFindPrintHeader = msScreenName
    Else
      sCaption = msScreenName
      msFindCaption = "Find - " & sCaption
      msStatusCaption = msScreenName
      msFindStatusCaption = msFindCaption
      msFindPrintHeader = msScreenName
    End If

    'Loop through forms collection to see if there are any more of the same
    ReDim alngScreenIndexes(0)
    lngMinIndex = 1
    lngMaxIndex = 1
    For Each frmForm In Forms
      If TypeOf frmForm Is frmRecEdit4 Then
        If (Not frmForm Is Me) Then
          If frmForm.Recordset.State = adStateClosed Then
            Unload frmForm
          Else
            'MH20031002 Fault 7082 Reference Property instead of object to trap errors
            'If (frmForm.ScreenID = mlngScreenID) And _
              (frmForm.ViewID = mobjTableView.ViewID) Then
            If (frmForm.ScreenID = mlngScreenID) And _
              (frmForm.ViewID = Me.ViewID) Then
              ReDim Preserve alngScreenIndexes(UBound(alngScreenIndexes) + 1)
              alngScreenIndexes(UBound(alngScreenIndexes)) = frmForm.ScreenIndex
              If frmForm.ScreenIndex < lngMinIndex Then lngMinIndex = frmForm.ScreenIndex
              If frmForm.ScreenIndex > lngMaxIndex Then lngMaxIndex = frmForm.ScreenIndex
            End If
          End If
        End If
      End If
    Next

    If UBound(alngScreenIndexes) > 0 Then
      lngNextIndex = 0
      For lngLoop = 1 To lngMaxIndex
        fFound = False
        
        For lngLoop2 = 1 To UBound(alngScreenIndexes)
          If alngScreenIndexes(lngLoop2) = lngLoop Then
            fFound = True
            Exit For
          End If
        Next lngLoop2
        
        If Not fFound Then
          lngNextIndex = lngLoop
          Exit For
        End If
      Next lngLoop
      
      If lngNextIndex = 0 Then
        lngNextIndex = lngMaxIndex + 1
      End If
      
      If lngNextIndex > 1 Then
        Me.Caption = sCaption & " (" & lngNextIndex & ")" & sRecordDescription
        msFindCaption = msFindCaption & " (" & lngNextIndex & ")"
        msStatusCaption = msStatusCaption & " (" & lngNextIndex & ")" & sRecordDescription
        msFindStatusCaption = msFindStatusCaption & " (" & lngNextIndex & ")"
      Else
        Me.Caption = sCaption & sRecordDescription
        msStatusCaption = msStatusCaption & sRecordDescription
        msFindPrintHeader = msFindPrintHeader & sRecordDescription
      End If
      
      mlngScreenIndex = lngNextIndex
    Else
      mlngScreenIndex = 1
      Me.Caption = sCaption & sRecordDescription
      msStatusCaption = msStatusCaption & sRecordDescription
    End If
  
    msFindPrintHeader = msFindPrintHeader & sRecordDescription
  End If

  ' Get rid of the icon
  RemoveIcon Me

End Sub


Private Function LoadControls(pobjScreen As clsScreen) As Boolean
  ' Load controls onto the screen when the form is loaded.
  On Error GoTo ErrorTrap_LoadControls
  
  On Error GoTo 0
  
  Dim fFound As Boolean
  Dim fSelectOK As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim iTabIndex As Integer
  Dim iCurrentTabIndex As Integer
  Dim lngCount As Long
  Dim sText As String
  Dim sFormat As String
  Dim sPictureFile As String
  Dim iControlType As ControlTypes
  Dim iControlDataType As SQLDataType
  Dim objControlArray As Object
  Dim objNewControl As Object
  Dim objSuperControl As Object
  Dim asRadio As Variant
  Dim colTabIndices As Collection
  Dim objScreenControl As clsScreenControl
  Dim iDigitCount As Integer
  
  Dim lngControlLeft As Long
  Dim lngControlTop As Long
  Dim lngControlHeight As Long
  Dim lngCONTROLWIDTH As Long
  
  
  ' Instantiate a new collection to hold the control indices.
  Set colTabIndices = New Collection

  ' Clear the array of linked records.
  ' Column 1 = table ID.
  ' Column 2 = record ID.
  ReDim malngLinkRecordIDs(2, 0)
  
  '  Get the screen's control details.
  Set mobjScreenControls = mdatRecEdit.GetControls(mlngScreenID)
  
  ' Instantiate each control on the screen.
  For Each objScreenControl In mobjScreenControls.Collection
    
    ' Get the control type.
    iControlType = objScreenControl.ControlType
    ' Get the column's data type.
    iControlDataType = objScreenControl.DataType

    If Not mblnScreenHasAutoUpdate Then
      If objScreenControl.AutoUpdateLookupValues Then
        mblnScreenHasAutoUpdate = True
      End If
    End If
    
    ' Get the control array eg. text1 array
    Set objControlArray = GetControlArray(objScreenControl)

    ' If we have a control array then load it into memory
    If Not objControlArray Is Nothing Then
      
      ' Create an instance of the required control.
      Load objControlArray(objControlArray.Count)

      ' Set new control to be this control
      Set objNewControl = objControlArray(objControlArray.Count - 1)

      ' If the control is an OLE control then we need to put it into a frame
      ' so that the OLE can appear in front of other frames.
      If iControlType = ctlOle Then
      
        ' Stop
        Load OLEFrame(OLEFrame.Count)
        
        
        Set objSuperControl = OLEFrame(OLEFrame.Count - 1)
        objSuperControl.Visible = True
      Else
        Set objSuperControl = Nothing
      End If

      With objNewControl
        .Tag = objScreenControl.Key

        ' Set the container of the new control.
        If objSuperControl Is Nothing Then
          Set .Container = IIf(mfUseTab, fraTabPage(objScreenControl.PageNo), Me)
        Else
          Set .Container = objSuperControl
          Set objSuperControl.Container = IIf(mfUseTab, fraTabPage(objScreenControl.PageNo), Me)
        End If

'        JDM - 11/06/2007 - Never set to anything else
'        If iControlType = ctlLabel Then
'          .BorderStyle = 0
'        End If

        If iControlType = ctlLine Then
          .Alignment = objScreenControl.Alignment
        End If

''        If TypeOf objNewControl Is TDBText6Ctl.TDBText _
''          Or TypeOf objNewControl Is TDBNumber6Ctl.TDBNumber Then
''              If mobjBorders Is Nothing Then
''                Set mobjBorders = New clsBorders
''              End If
''              mobjBorders.SetBorder objNewControl.hwnd, ctTextBox, RGB(169, 177, 184)
''
''        ElseIf TypeOf objNewControl Is GTMaskDate.GTMaskDate Then
''              If mobjBorders Is Nothing Then
''                Set mobjBorders = New clsBorders
''              End If
''              mobjBorders.SetBorder objNewControl.hwnd, ctTextBox, RGB(169, 177, 184)
''
''        End If

'        ' Position the new control on the screen
'        If objSuperControl Is Nothing Then
'          lngControlTop = objScreenControl.TopCoord
'          lngControlLeft = objScreenControl.LeftCoord
'        Else
'          objSuperControl.Top = objScreenControl.TopCoord
'          objSuperControl.Left = objScreenControl.LeftCoord
'          lngControlTop = 0
'          lngControlLeft = 0
'        End If
'
'        ' Position the control
'        If TypeOf objNewControl Is ComboBox Then
'          .Move lngControlLeft, lngControlTop, objScreenControl.Width
'        Else
'          .Move lngControlLeft, lngControlTop, objScreenControl.Width, objScreenControl.Height
'        End If

        ' Position the new control on the screen
        If objSuperControl Is Nothing Then
          .Top = objScreenControl.TopCoord
          .Left = objScreenControl.LeftCoord
        Else
          objSuperControl.Top = objScreenControl.TopCoord
          objSuperControl.Left = objScreenControl.LeftCoord

          .Top = 0
          .Left = 0
        End If

        ' Check that the control is not a combo box
        ' and if it isn't then set it's height
        If Not TypeOf objNewControl Is XtremeSuiteControls.ComboBox Then
          .Height = objScreenControl.Height
        End If

        ' Set the controls width
        .Width = objScreenControl.Width


        ' Resize the OLE control's container frame if required.
        If Not objSuperControl Is Nothing Then
          objSuperControl.Height = .Height
          objSuperControl.Width = .Width
        End If

        ' Check that the caption isn't null and if it isn't use the
        ' SetCaption function to set it which tries to set both the
        ' caption and the text, ignoring all errors
        If iControlType <> ctlText And iControlType <> ctlColourPicker Then
          SetCaption objNewControl, objScreenControl.Caption
        End If

        If TypeOf objNewControl Is CommandButton Then
          ' See if the linked table is already in our array of links.
          fFound = False
          For iNextIndex = 1 To UBound(malngLinkRecordIDs, 2)
            If malngLinkRecordIDs(1, iNextIndex) = objScreenControl.LinkTableID Then
              fFound = True
              Exit For
            End If
          Next iNextIndex

          If Not fFound Then
            iNextIndex = UBound(malngLinkRecordIDs, 2) + 1
            ReDim Preserve malngLinkRecordIDs(2, iNextIndex)
            malngLinkRecordIDs(1, iNextIndex) = objScreenControl.LinkTableID
            malngLinkRecordIDs(2, iNextIndex) = 0
          End If
        End If

        ' Check that the control is not an image and that it has
        ' a tab index > 0 and set it
        If Not TypeOf objNewControl Is Image Then
          If objScreenControl.TabIndex > 0 Then
            ' Find the appropriate tabIndex to set for the current control.
            iTabIndex = 0
            iCurrentTabIndex = objScreenControl.TabIndex
            For iLoop = 1 To colTabIndices.Count
              If iCurrentTabIndex < colTabIndices.Item(iLoop) Then
                Exit For
              End If

              iTabIndex = iLoop
            Next iLoop

            If (iTabIndex + 1) > colTabIndices.Count Then
              colTabIndices.Add iCurrentTabIndex
            Else
              colTabIndices.Add iCurrentTabIndex, , iTabIndex + 1
            End If

            .TabIndex = iTabIndex
          End If
        End If

        ' Set the BackColor property for all controls
        ' except Images and tabs....and OLE's
        If (iControlType And (ctlImage Or ctlTab Or ctlOle Or ctlPhoto Or ctlCommand Or ctlLine Or ctlNavigation Or ctlColourPicker)) = 0 Then
          .BackColor = objScreenControl.BackColor
        End If

        ' If it is an OLE then set its Display type and help ID property
        If iControlType = ctlOle Or iControlType = ctlPhoto Then
          .ColumnID = objScreenControl.ColumnID
          .OLEType = IIf(objScreenControl.OLEType = OLE_EMBEDDED, OLE_UNC, objScreenControl.OLEType)
        End If

        ' Set the Font properties for all controls except
        ' images and OLE's
        If (iControlType And (ctlImage Or ctlOle Or ctlPhoto Or ctlLine Or ctlColourPicker)) = 0 Then
          If iControlType = ctlLabel Then
            .Font = objScreenControl.FontName
            .FontSize = objScreenControl.FontSize
            .FontBold = objScreenControl.FontBold
            .FontItalic = objScreenControl.FontItalic
            .Font.Strikethrough = objScreenControl.FontStrikethru
            .FontUnderline = objScreenControl.FontUnderline
          Else
            ' JPD20030211 Fault 5041
            If (iControlType = ctlWorkingPattern) Or _
              (iControlType = ctlRadio) Then
              Dim objFont As StdFont
              Set objFont = New StdFont
              objFont.Bold = objScreenControl.FontBold
              objFont.Italic = objScreenControl.FontItalic
              objFont.Name = objScreenControl.FontName
              objFont.Size = objScreenControl.FontSize
              objFont.Strikethrough = objScreenControl.FontStrikethru
              objFont.Underline = objScreenControl.FontUnderline
              Set .Font = objFont
            Else
              .Font.Name = objScreenControl.FontName
              .Font.Size = objScreenControl.FontSize
              .Font.Bold = objScreenControl.FontBold
              .Font.Italic = objScreenControl.FontItalic
              .Font.Strikethrough = objScreenControl.FontStrikethru
              .Font.Underline = objScreenControl.FontUnderline
            End If
          End If
        End If

        ' JDM - 16/05/2005 - Was on my machine... :-)
        'JPD 20030903 Fault 4290 - Background colour wasn't busted here !
        ' JDM - 08/07/03 - Fault 4290 - Background colour is busted
        'If iControlType = ctlworkingpattern Then
        '  .BackColor = Me.BackColor
        'End If

        ' Set the ForeColor property for all controls except
        ' images, ole's and tabs
        If (iControlType And (ctlImage Or ctlOle Or ctlTab Or ctlPhoto Or ctlCommand Or ctlLine Or ctlColourPicker)) = 0 Then
          .ForeColor = objScreenControl.ForeColor
        End If

        ' Photo type
        If iControlType = ctlPhoto Or iControlType = ctlImage Then
          .BorderStyle = objScreenControl.BorderStyle
        End If

        ' If the control is an image then load the Picture
        If iControlType = ctlImage Then

          sPictureFile = LoadScreenControlPicture(objScreenControl.PictureID)
          If LenB(sPictureFile) <> 0 Then
            .Picture = LoadPicture(sPictureFile)
            Kill sPictureFile
          End If

          ' JPD20021016 Fault 4605
          objNewControl.Enabled = False
        End If

        ' Navigation control
        If iControlType = ctlNavigation Then
          .ColumnID = objScreenControl.ColumnID
          .Caption = objScreenControl.Caption
          .DisplayType = objScreenControl.DisplayType
          .NavigateTo = objScreenControl.NavigateTo
          .NavigateIn = objScreenControl.NavigateIn
          .NavigateOnSave = objScreenControl.NavigateOnSave
          .TabStop = (.DisplayType = NavigationDisplayType.Button)
          .Enabled = True
        End If

        ' If the control is a radio then set it's Options and borderstyle properties
        If iControlType = ctlRadio Then
          asRadio = objScreenControl.ControlValues
          .SetOptions asRadio
          .BorderStyle = objScreenControl.BorderStyle
          .Alignment = objScreenControl.Alignment
        End If

        ' If the control is a spinner then set the min, max and increment properties
        If iControlType = ctlSpin Then
          .MinimumValue = objScreenControl.SpinnerMinimum
          .MaximumValue = objScreenControl.SpinnerMaximum
          .Increment = objScreenControl.SpinnerIncrement
          .Alignment = objScreenControl.ColumnAlignment
          .SpinnerPosition = objScreenControl.Alignment
        End If

        'Set the alignment property
        If iControlType = ctlCheck Then
          .Alignment = objScreenControl.Alignment
        ElseIf iControlType = ctlText Then
          
          If TypeOf objNewControl Is TDBText6Ctl.TDBText _
            Or TypeOf objNewControl Is TDBMask6Ctl.TDBMask _
            Or TypeOf objNewControl Is TDBNumber6Ctl.TDBNumber Then
            
            .AlignHorizontal = objScreenControl.ColumnAlignment
            
          'JPD 20050302 Fault 9847
          ElseIf (Not TypeOf objNewControl Is GTMaskDate.GTMaskDate) And _
            (Not TypeOf objNewControl Is TDBNumberCtrl.TDBNumber) And _
            (Not TypeOf objNewControl Is TDBNumber6Ctl.TDBNumber) Then

            .Alignment = objScreenControl.ColumnAlignment

          End If
        End If

        'MH20010108
        If iControlDataType = sqlDate Then
          datGeneral.FormatTDBNumberControl objNewControl
        End If

        'Set any masks and other settings for the mask control
        If iControlType = ctlText Then
          'Check if it is a text box (Only if multi-line though)
          'JPD 20050302 Fault 9847
          If (objScreenControl.Multiline Or _
            (LenB(objScreenControl.Mask) = 0)) And _
            (Not TypeOf objNewControl Is GTMaskDate.GTMaskDate) And _
            (Not TypeOf objNewControl Is TDBNumberCtrl.TDBNumber) And _
            (Not TypeOf objNewControl Is TDBNumber6Ctl.TDBNumber) Then

            If Not objScreenControl.Multiline Then
              .MaxLength = objScreenControl.Size
            End If

          Else
            If iControlDataType <> sqlDate Then
              If LenB(objScreenControl.Mask) <> 0 Then
                Select Case iControlDataType
                  Case sqlInteger, sqlNumeric     'DTPNumber control
                    sFormat = ConvertMaskToNumeric(objScreenControl.Mask)
                    .Format = sFormat

                    If objScreenControl.BlankIfZero Then
                      .DisplayFormat = ""
                    Else
                      .DisplayFormat = GetDisplayFormat(sFormat)
                      'JPD 20050309 - changed number control, so now need to specify the 'null' display format.
                      .DisplayFormat = .DisplayFormat & ";;" & Replace(.DisplayFormat, "#", "")
                    End If

                  Case Else
                    .Format = Replace(Replace(objScreenControl.Mask, "S", "A"), "s", "a")
                    
                End Select
              Else
                Select Case iControlDataType
                  Case sqlInteger

                    ' Loop and create the format mask
                    sFormat = ""
                    For lngCount = 1 To objScreenControl.Size
                      If objScreenControl.Use1000Separator Then
                        sFormat = IIf(lngCount Mod 3 = 0 And (lngCount <> (objScreenControl.Size)), ",#", "#") & sFormat
                      Else
                        sFormat = "#" & sFormat
                      End If
                    Next lngCount

                    If Not objScreenControl.BlankIfZero Then
                      If LenB(sFormat) <> 0 Then
                        sFormat = Left$(sFormat, Len(sFormat) - 1) & "0"
                      End If
                    End If

                    .Format = sFormat
                    .DisplayFormat = sFormat

                    'JPD 20050309 - changed number control, so now need to specify the 'null' display format.
                    .DisplayFormat = .DisplayFormat & ";;" & Replace(.DisplayFormat, "#", "")

                    .MaxValue = 2147483647#
                    .MinValue = -2147483648#

                  Case sqlNumeric
                    sFormat = ""
                    iDigitCount = 1

                    ' Loop and create the format mask
                    For lngCount = 1 To (objScreenControl.Size - objScreenControl.Decimals)
                      If objScreenControl.Use1000Separator Then
                        sFormat = IIf(lngCount Mod 3 = 0 And (lngCount <> (objScreenControl.Size - objScreenControl.Decimals)), ",#", "#") & sFormat
                      Else
                        sFormat = "#" & sFormat
                      End If
                    Next lngCount

                    'NPG20080418 Fault 10186
                    'If Not objScreenControl.BlankIfZero Then
                      If LenB(sFormat) <> 0 Then
                        sFormat = Left$(sFormat, Len(sFormat) - 1) & "0"
                      End If
                    'End If

                    If objScreenControl.Decimals > 0 Then
                      sFormat = sFormat & "."
                      For lngCount = 1 To objScreenControl.Decimals
                        'NPG20080418 Fault 10186
                        'If objScreenControl.BlankIfZero Then
                        '  sFormat = sFormat & "#"
                        'Else
                          sFormat = sFormat & "0"
                        'End If
                      Next lngCount
                    End If

                    .DecimalPoint = UI.GetSystemDecimalSeparator
                    .Separator = UI.GetSystemThousandSeparator
                    .Format = sFormat
                    .DisplayFormat = sFormat

                    'NPG20080418 Fault 10186
                    'JPD 20050309 - changed number control, so now need to specify the 'null' display format.
                    ' If InStr(.DisplayFormat, "0") <> 0 Then
                    '     .DisplayFormat = .DisplayFormat & ";;" & Replace(.DisplayFormat, "#", "")
                    ' Else
                    '   .DisplayFormat = .DisplayFormat & ";;"
                    ' End If
                    
                    If InStr(.DisplayFormat, "0") <> 0 Then
                      If objScreenControl.BlankIfZero Then
                        .DisplayFormat = .DisplayFormat & ";;{};{}"
                      Else
                        .DisplayFormat = .DisplayFormat & ";;" & Replace(.DisplayFormat, "#", "")
                      End If
                    Else
                      .DisplayFormat = .DisplayFormat & ";;"
                    End If
                    
                    'MH20010202 Fault 1785
                    .MaxValue = Val(String(objScreenControl.Size - objScreenControl.Decimals, "9") & "." & String(objScreenControl.Decimals, "9"))
                    .MinValue = (.MaxValue * -1)

                  Case Else
                    If LenB(objScreenControl.Mask) <> 0 Then
                      sFormat = ""
                      For lngCount = 1 To objScreenControl.Size
                        sFormat = sFormat & "&"
                      Next lngCount

                      .Format = sFormat
                      .ShowLiterals = 1
                      .AllowSpace = True
                    End If
                End Select
              End If
            End If
          End If
        End If

        ' See if this control belongs to this table
        ' ie. is it for database entry.
        'MH20031002 Fault 7082 Reference Property instead of object to trap errors
        'If objScreenControl.TableID = mobjTableView.TableID Then
        If objScreenControl.TableID = Me.TableID Then
          ' Check that it is mapped to a column
          If objScreenControl.ColumnID > 0 Then
            fSelectOK = False

            If mcolColumnPrivileges.IsValid(objScreenControl.ColumnName) Then
              fSelectOK = mcolColumnPrivileges.Item(objScreenControl.ColumnName).AllowSelect
            End If

            If fSelectOK Then
              ' If the control is a option group...
              If TypeOf objNewControl Is COA_OptionGroup Then
                .MaxLength = objScreenControl.Size
              End If

              If iControlType = ctlCombo Then
                If objScreenControl.ColumnType <> colLookup Then
                  PopulateComboFromArray objNewControl, objScreenControl.ControlValues, objScreenControl.Mandatory
                End If
              End If

              ' Disable control if no permission is granted.
              If TypeOf objNewControl Is COA_Navigation Then
                .Enabled = True
                .NavigateOnSave = objScreenControl.NavigateOnSave
              Else
                .Enabled = Not (objScreenControl.ReadOnly Or objScreenControl.ScreenReadOnly)
              End If
              
              If .Enabled Then
                'Fix for date and lookup control
                If (TypeOf objNewControl Is GTMaskDate.GTMaskDate) Or _
                  (TypeOf objNewControl Is COA_Lookup) Then

                  sText = .Text
                  .Enabled = mcolColumnPrivileges.Item(objScreenControl.ColumnName).AllowUpdate
                  .Text = sText
                ' JPD20030311 Fault 5139
                ElseIf TypeOf objNewControl Is CommandButton Then
                  .Enabled = (objScreenControl.LinkTableID > 0) And _
                    (objScreenControl.LinkTableID <> mlngParentTableID) And _
                    mcolColumnPrivileges.Item(objScreenControl.ColumnName).AllowUpdate

                ElseIf (iControlType And ctlOle) = ctlOle Then
                  .Enabled = True

                Else
                  If (iControlType = ctlText) _
                    And (objScreenControl.Multiline = True) _
                    And ((objScreenControl.DataType = sqlVarChar) _
                        Or (objScreenControl.DataType = sqlLongVarChar)) Then

                    .Enabled = mcolColumnPrivileges.Item(objScreenControl.ColumnName).AllowUpdate

                    If Not .Enabled Then
                      .Enabled = True
                      .ReadOnly = True
                      .BackColor = COL_GREY
                      .ForeColor = vbGrayText
                    End If

                  Else
                    .Enabled = mcolColumnPrivileges.Item(objScreenControl.ColumnName).AllowUpdate
                  End If
                End If

                If Not .Enabled Then
                  If (iControlType And (ctlImage Or ctlOle Or ctlTab Or ctlPhoto Or ctlWorkingPattern)) = 0 Then
                    .BackColor = COL_GREY
                  End If
                End If
              Else
                If (iControlType And (ctlImage Or ctlOle Or ctlTab Or ctlPhoto)) = 0 Then
                  .BackColor = COL_GREY
                  'NHRD23032007 Fault 11675 Making readonly multiline fields scrollable
                  If objScreenControl.Multiline Then
                    .Enabled = True
                    .ReadOnly = True
                    .ForeColor = vbGrayText
                  End If
                End If
              End If
            Else
              ' No Select privileges, grey it out and disable the control
              .Enabled = False
              If (iControlType And (ctlImage Or ctlOle Or ctlTab Or ctlPhoto)) = 0 Then
                .BackColor = COL_GREY
              End If

              'JPD 20050302 Fault 9847
              If (TypeOf objNewControl Is TDBNumberCtrl.TDBNumber) Or _
                (TypeOf objNewControl Is TDBNumber6Ctl.TDBNumber) Then
                .DisplayFormat = "#"
              End If
            End If
          End If
        Else
          ' Parent table control.
          If (iControlType And (ctlLabel Or ctlFrame Or ctlImage Or ctlLine Or ctlNavigation)) = 0 Then
            .Enabled = False

            If (iControlType <> ctlImage) And _
              (iControlType <> ctlPhoto) And _
              (iControlType <> ctlWorkingPattern) And _
              (iControlType <> ctlNavigation) And _
              (iControlType <> ctlOle) Then 'NPG20080519 Fault 13003
              .BackColor = COL_GREY
              'NHRD23032007 Fault 11675 Making readonly multiline fields scrollable
              If objScreenControl.Multiline Then
                .Enabled = True
                .ReadOnly = True
                .ForeColor = vbGrayText
              End If
            End If
          End If
        End If
        
'        If iControlType = ctlLabel Then
'          .ZOrder 0
'        End If
'
'        ' Show control if on first tab
'        .Visible = (objScreenControl.PageNo = 1) Or (objScreenControl.PageNo = 0)
        .Visible = True
      
      End With

    End If
  Next objScreenControl
  
  ' Set the correct zorder for the controls
  SetControlLevel
  
  Set objScreenControl = Nothing
  Set objSuperControl = Nothing
  
  Set colTabIndices = Nothing

  LoadControls = True
  Exit Function

ErrorTrap_LoadControls:
  gobjProgress.Visible = False
  COAMsgBox Err.Description & " - LoadControls", vbExclamation + vbOKOnly, app.ProductName
  LoadControls = False

End Function


Private Function PopulateComboFromArray(pobjCombo As XtremeSuiteControls.ComboBox, pasComboItems As Variant, ByVal pfMandatory As Boolean) As Boolean
  ' Populates a combo box with items from an array.
  Dim iCount As Integer

  ' Clear the combo box and loop through the array adding the items
  With pobjCombo
    .Clear
  
    For iCount = 0 To UBound(pasComboItems, 2)
      .AddItem pasComboItems(0, iCount)
    Next iCount
  
    If Not pfMandatory Then
      .AddItem ""
    End If
  End With
  
End Function


Public Property Get StatusCaption() As String
  StatusCaption = msStatusCaption

End Property

Public Property Get FindStatusCaption() As String
  FindStatusCaption = msFindStatusCaption

End Property


Public Property Get FindCaption() As String
  FindCaption = msFindCaption

End Property

Public Property Get FindPrintHeader() As String
  FindPrintHeader = msFindPrintHeader

End Property



Public Property Get ScreenIndex() As Long
  ScreenIndex = mlngScreenIndex

End Property




Public Property Get ScreenName() As String
  ' Get the screen name of the form.
  ScreenName = msScreenName

End Property



Private Function LoadScreenControlPicture(plngPictureID As Long) As String
  ' Read the given picture from the database.
  On Error GoTo ErrorTrap
  
  Dim iFragments As Integer
  Dim iTempFile  As Integer
  Dim lngOffset As Long
  Dim lngPictureSize As Long
  Dim sTempName As String
  Dim sPictureFile As String
  Dim bytChunks() As Byte
  Dim rsPictures As ADODB.Recordset
  
  Const conChunkSize = 2 ^ 14

  sPictureFile = ""
  
  Set rsPictures = datGeneral.GetPicture(plngPictureID)
  With rsPictures
    If Not (.BOF And .EOF) Then
      .MoveFirst
        
      sTempName = GetTmpFName
      iTempFile = 1
      Open sTempName For Binary Access Write As iTempFile
        
      lngPictureSize = !Picture.ActualSize
      iFragments = lngPictureSize Mod conChunkSize
        
      ReDim bytChunks(iFragments)
      
      Do While lngOffset < lngPictureSize
        bytChunks() = !Picture.GetChunk(conChunkSize)
        lngOffset = lngOffset + conChunkSize
        Put iTempFile, , bytChunks()
      Loop

      Close iTempFile
        
      sPictureFile = sTempName
    End If
    
    .Close
  End With
  Set rsPictures = Nothing
  
TidyUpAndExit:
  LoadScreenControlPicture = sPictureFile
  Exit Function
  
ErrorTrap:
  sPictureFile = ""
  Resume TidyUpAndExit
  
End Function



Private Function GetControlArray(ByRef pobjScreenControl As clsScreenControl) As Object
  ' Return a control array of the appropriate type.
  Dim objControlArray As Object
 
  Select Case pobjScreenControl.ControlType
    
    Case ctlText
      If pobjScreenControl.Multiline Then
        Set objControlArray = Me.TDBText1
      Else
        Select Case pobjScreenControl.DataType
          Case sqlDate
            'Set objControlArray = Me.TDBDate1
            Set objControlArray = Me.GTMaskDate1
            
          Case sqlNumeric, sqlInteger
            'JPD 20050302 Fault 9847
            'Set objControlArray = Me.TDBNumber1
            Set objControlArray = Me.TDBNumber2
            
          Case Else
            If LenB(pobjScreenControl.Mask) <> 0 Then
              Set objControlArray = Me.TDBMask1
            Else
              Set objControlArray = Me.Text1
            End If
        End Select
      End If
      
    Case ctlLabel
      Set objControlArray = Me.Label1
      
    Case ctlPhoto
      Set objControlArray = Me.ASRUserImage1
      
    Case ctlCheck
      Set objControlArray = Me.Check1
      
    Case ctlCombo
      If (pobjScreenControl.ColumnType = colLookup) Then
        Set objControlArray = Me.ctlNewLookup1
      Else
        Set objControlArray = Me.Combo1
      End If
      
    Case ctlImage
      Set objControlArray = Me.ASRUserImage1
      
    Case ctlOle
      Set objControlArray = Me.OLE1
      
    Case ctlRadio
      Set objControlArray = Me.OptionGroup1
      
    Case ctlSpin
      Set objControlArray = Me.Spinner1
      
    Case ctlTab
    
    Case ctlLine
      Set objControlArray = Me.ASRLine1
      
    Case ctlFrame
      Set objControlArray = Me.Frame1
      
    Case ctlCommand
      Set objControlArray = Me.Command1
      
    Case ctlWorkingPattern
      Set objControlArray = Me.ASRWorkingPattern1
      
    Case ctlNavigation
      Set objControlArray = Me.COA_Navigation1
      
    Case ctlColourPicker
      Set objControlArray = ColourSelector1
    
    Case Else
      Set objControlArray = Nothing
  End Select
      
  Set GetControlArray = objControlArray
  
End Function



Private Function GetIcon(ByVal plngPictureID As Long) As Boolean
  ' Gets an icon from the database based on the given id.
  On Error GoTo ErrorTrap_GetIcon
  
  Dim rsPictures As Recordset
  
  If plngPictureID > 0 Then
    Set rsPictures = datGeneral.GetPicture(plngPictureID)
    With rsPictures
      If Not (.BOF And .EOF) Then
        ReadPicture Me, rsPictures!Picture, rsPictures!Picture.ActualSize
      End If
      .Close
    End With
    Set rsPictures = Nothing
  End If
  
  GetIcon = True
  Exit Function
  
ErrorTrap_GetIcon:
  GetIcon = False
  Err = False
  
End Function

Public Property Get ScreenID() As Long
  ' Get the screen id of the form.
  ScreenID = mlngScreenID

End Property

Public Property Let ScreenID(plngNewValue As Long)
  ' Set the Screen ID of the form.
  mlngScreenID = plngNewValue

End Property

Private Function GetRecords() As Boolean
  ' Get a set of records from the database and set up the relevant security information.
  On Error GoTo ErrorTrap_GetRecords
    
  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim fColumnOK As Boolean
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnSelectSQL As String
  Dim sOrderColumnSQL As String
  Dim sOrderSelectSQL As String
  Dim sOrderBySQL As String
  Dim sJoinSQL As String
  Dim sWhereSQL As String
  Dim sFilterSQL As String
  Dim sGetRecordsSQL As String
  Dim sGetRecordCountSQL As String
  Dim rsInfo As Recordset
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim alngTableViews() As Long
  Dim asViews() As String
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)
    
  ' Initialise the SELECT string with the columns in the base table/view that
  ' are readable by the user.
  sColumnSelectSQL = GetSelectString()
  fOK = (LenB(sColumnSelectSQL) <> 0)
  
  If fOK Then
    ' Initialise the ORDER and JOIN SQL strings.
    sOrderSelectSQL = ""
    sOrderBySQL = ""
    sJoinSQL = ""
    sWhereSQL = ""
    sFilterSQL = ""
    
    'lngFirstOrderColumn = 0
    mlngFirstOrderColumnID = 0
    
    ' Get the parent table/view and record IDs if required.
    GetParentDetails
    
    ' Construct the ORDER string (including JOINing any parent tables/views used in the order).
    Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)
      
    Do While Not rsInfo.EOF
      If rsInfo!Type = "O" Then
        ' Get the column privileges collection for the given table.
        'MH20031002 Fault 7082 Reference Property instead of object to trap errors
        'If rsInfo!TableID = mobjTableView.TableID Then
        If rsInfo!TableID = Me.TableID Then
          sSource = msTableViewName
        Else
          sSource = rsInfo!TableName
        End If
        Set objColumnPrivileges = GetColumnPrivileges(sSource)
        sRealSource = gcoTablePrivileges.Item(sSource).RealSource

        fColumnOK = objColumnPrivileges.IsValid(rsInfo!ColumnName)

        If fColumnOK Then
          fColumnOK = objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect
        End If
        Set objColumnPrivileges = Nothing

        If fColumnOK Then
          ' The column can be read from the base table/view, or directly from a parent table.
          ' Add the column to the order string.
          sOrderBySQL = sOrderBySQL & _
            IIf(LenB(sOrderBySQL) <> 0, ", ", "") & _
            sRealSource & "." & Trim(rsInfo!ColumnName) & _
            IIf(rsInfo!Ascending, "", " DESC")

          If mlngFirstOrderColumnID = 0 Then
            mlngFirstOrderColumnID = rsInfo!ColumnID
            mfFirstOrderColumnAscending = rsInfo!Ascending
          End If
          
          ' If the column comes from a parent table, then add the table to the Join code.
          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'If rsInfo!TableID <> mobjTableView.TableID Then
          If rsInfo!TableID <> Me.TableID Then
            ' Check if the table has already been added to the join code.
            fFound = False
            For iNextIndex = 1 To UBound(alngTableViews, 2)
              If alngTableViews(1, iNextIndex) = 0 And _
                alngTableViews(2, iNextIndex) = rsInfo!TableID Then
                fFound = True
                Exit For
              End If
            Next iNextIndex
            
            If Not fFound Then
              ' The table has not yet been added to the join code, so add it to the array and the join code.
              iNextIndex = UBound(alngTableViews, 2) + 1
              ReDim Preserve alngTableViews(2, iNextIndex)
              alngTableViews(1, iNextIndex) = 0
              alngTableViews(2, iNextIndex) = rsInfo!TableID
              
              sJoinSQL = sJoinSQL & _
                " LEFT OUTER JOIN " & sRealSource & _
                " ON " & mobjTableView.RealSource & ".ID_" & Trim(Str(rsInfo!TableID)) & _
                " = " & sRealSource & ".ID"
            End If
          End If
        Else
          ' The column cannot be read from the base table/view, or directly from a parent table.
          ' If it is a column from a parent table, then try to read it from the views on the parent table.
          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'If rsInfo!TableID <> mobjTableView.TableID Then
          If rsInfo!TableID <> Me.TableID Then
            ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
            ReDim asViews(0)
            For Each objTableView In gcoTablePrivileges.Collection
              If (Not objTableView.IsTable) And _
                (objTableView.TableID = rsInfo!TableID) And _
                (objTableView.AllowSelect) Then

                sSource = objTableView.ViewName
                sRealSource = gcoTablePrivileges.Item(sSource).RealSource

                ' Get the column permission for the view.
                Set objColumnPrivileges = GetColumnPrivileges(sSource)

                If objColumnPrivileges.IsValid(rsInfo!ColumnName) Then
                  If objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect Then
                    ' Add the view info to an array to be put into the column list or order code below.
                    iNextIndex = UBound(asViews) + 1
                    ReDim Preserve asViews(iNextIndex)
                    asViews(iNextIndex) = objTableView.ViewName
                    
                    ' Add the view to the Join code.
                    ' Check if the view has already been added to the join code.
                    fFound = False
                    For iNextIndex = 1 To UBound(alngTableViews, 2)
                      If alngTableViews(1, iNextIndex) = 1 And _
                        alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                        fFound = True
                        Exit For
                      End If
                    Next iNextIndex

                    If Not fFound Then
                      ' The view has not yet been added to the join code, so add it to the array and the join code.
                      iNextIndex = UBound(alngTableViews, 2) + 1
                      ReDim Preserve alngTableViews(2, iNextIndex)
                      alngTableViews(1, iNextIndex) = 1
                      alngTableViews(2, iNextIndex) = objTableView.ViewID

                      sJoinSQL = sJoinSQL & _
                        " LEFT OUTER JOIN " & sRealSource & _
                        " ON " & mobjTableView.RealSource & ".ID_" & Trim(Str(objTableView.TableID)) & _
                        " = " & sRealSource & ".ID"
                    End If
                  End If
                End If
                Set objColumnPrivileges = Nothing

              End If
            Next objTableView
            Set objTableView = Nothing

            ' The current user does have permission to 'read' the column through a/some view(s) on the
            ' table.
            If UBound(asViews) > 0 Then
              ' Add the column to the column list.
              sOrderColumnSQL = ""
              For iNextIndex = 1 To UBound(asViews)
                If iNextIndex = 1 Then
                  sOrderColumnSQL = "CASE "
                End If

                sOrderColumnSQL = sOrderColumnSQL & _
                  " WHEN NOT " & asViews(iNextIndex) & "." & rsInfo!ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & rsInfo!ColumnName
              Next iNextIndex

              If LenB(sOrderColumnSQL) <> 0 Then
                sOrderColumnSQL = sOrderColumnSQL & _
                  " ELSE NULL" & _
                  " END AS '?" & rsInfo!ColumnName & "'"

                sOrderSelectSQL = sOrderSelectSQL & _
                  IIf(LenB(sOrderSelectSQL) <> 0, ", ", "") & _
                  sOrderColumnSQL

                ' Add the column to the order string.
                sOrderBySQL = sOrderBySQL & _
                  IIf(LenB(sOrderBySQL) <> 0, ", ", "") & _
                  "'?" & Trim(rsInfo!ColumnName) & "'" & _
                  IIf(rsInfo!Ascending, "", " DESC")

                If mlngFirstOrderColumnID = 0 Then
                  mlngFirstOrderColumnID = rsInfo!ColumnID
                  mfFirstOrderColumnAscending = rsInfo!Ascending
                End If
              End If
            End If
          End If
        End If
      End If
      
      rsInfo.MoveNext
    Loop

    rsInfo.Close
    Set rsInfo = Nothing
  
    ' Construct the WHERE string (including the filter).
    If (mlngParentTableID > 0) Then
      'TM14062004 Fault 8322 - If the Parent Record ID is 0 then we don't want to return any records.
      '...can't look for 0 or NULL as records that have empty links will be returned, therefore have
      'compared to -1 (no records should have negative ID)
      If (mlngParentRecordID > 0) Then
        sWhereSQL = " WHERE " & mobjTableView.RealSource & ".ID_" & mlngParentTableID & " = " & Trim(Str(mlngParentRecordID))
      Else
        sWhereSQL = " WHERE " & mobjTableView.RealSource & ".ID_" & mlngParentTableID & " = -1"
      End If
    End If

    sFilterSQL = GetFilter
    If LenB(sFilterSQL) <> 0 Then
      sWhereSQL = sWhereSQL & _
        IIf(LenB(sWhereSQL) <> 0, " AND ", " WHERE ") & sFilterSQL
    End If
    
    sGetRecordsSQL = "SELECT " & sColumnSelectSQL & _
      IIf(LenB(sOrderSelectSQL) <> 0, ", ", "") & sOrderSelectSQL & _
      " FROM " & mobjTableView.RealSource & _
      " " & sJoinSQL & _
      sWhereSQL & _
      IIf(LenB(sOrderBySQL) <> 0, " ORDER BY " & sOrderBySQL, "")
  
    sGetRecordCountSQL = "SELECT COUNT(" & mobjTableView.RealSource & ".id)" & _
      " FROM " & mobjTableView.RealSource & _
      sWhereSQL

    ' Set the record count variable instead of using the recordset's
    ' recordCount property as that takes a long time.
    msCountSQL = sGetRecordCountSQL
    mlngRecordCount = GetRecordCount
    
    If Not mrsRecords Is Nothing Then
      If mrsRecords.State <> adStateClosed Then
        mrsRecords.Close
        Set mrsRecords = Nothing
      End If
    End If
    
    ' Get the recordset.
    If mfRequiresLocalCursor Then gADOCon.CursorLocation = adUseClient
    Set mrsRecords = datGeneral.GetPersistentMainRecordset(sGetRecordsSQL)
    If mfRequiresLocalCursor Then gADOCon.CursorLocation = adUseServer
  End If
  
  GetRecords = fOK
    
  Exit Function
  
ErrorTrap_GetRecords:
  GetRecords = False
  COAMsgBox Err.Description & " - GetRecords", vbExclamation + vbOKOnly, app.ProductName
  Err = False
    
End Function

Private Function GetParentDetails() As Boolean
  ' This function will return true if the form has a parent.
  ' Variable defining the parent table/view IDs and the parent records ID
  ' will be set/reset accordingly.
  Dim fParentFound As Boolean
  Dim lngParentTableID As Long
  Dim lngParentViewID As Long
  Dim lngParentRecordID As Long
  Dim frmForm As Form

  fParentFound = False
  lngParentTableID = 0
  lngParentViewID = 0
  lngParentRecordID = 0
  
  ' Decide if we are a history table
  If (miScreenType = screenHistoryTable) Or _
    (miScreenType = screenHistoryView) Then

    ' We are a history table.
    ' Get the parent table and view ID
    For Each frmForm In Forms
      With frmForm
        If .Name = "frmRecEdit4" Then
          If .FormID = mlngParentFormID Then
            fParentFound = True
            lngParentTableID = .TableID
            lngParentViewID = .ViewID
            lngParentRecordID = .RecordID
            Exit For
          End If
        End If
      End With
    Next frmForm
    Set frmForm = Nothing
  End If
  
  mlngParentTableID = lngParentTableID
  mlngParentViewID = lngParentViewID
  mlngParentRecordID = lngParentRecordID
  GetParentDetails = fParentFound
  
End Function


Private Function GetFilter() As String
  ' Return the filter code for the recordset's defined filter.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim dtDateValue As Date
  Dim sCurrChar As String
  Dim sFilter As String
  Dim sFilterColName As String
  Dim sFilterValue As String
  Dim sModifiedFilterValue As String
  Dim sSubFilter As String
  Dim objColumn As CColumnPrivilege
  
  ' Construct the filter string.
  sFilter = ""
  For iLoop = 1 To UBound(mavFilterCriteria, 2)
    
    ' Find the column definition.
    For Each objColumn In mcolColumnPrivileges
      If objColumn.ColumnID = CLng(mavFilterCriteria(1, iLoop)) Then
      
        sFilterColName = mobjTableView.RealSource & "." & objColumn.ColumnName
        sFilterValue = mavFilterCriteria(3, iLoop)
      
        Select Case objColumn.DataType
          Case sqlOle  ' Not required as OLEs are not permitted in the Filter Column selection.
        
          Case sqlBoolean ' Logic columns.
            sSubFilter = sFilterColName & " = " & IIf(UCase$(sFilterValue) = "TRUE", "1", "0")
                  
          Case sqlNumeric, sqlInteger ' Numeric and Integer columns.
            If (LenB(sFilterValue) = 0) Then sFilterValue = "0"
            
            'MH20010108 Fault 1604
            sFilterValue = datGeneral.ConvertNumberForSQL(sFilterValue)

            Select Case mavFilterCriteria(2, iLoop)
              Case giFILTEROP_EQUALS
                sSubFilter = sFilterColName & " = " & sFilterValue
                If Val(sFilterValue) = 0 Then
                  sSubFilter = sSubFilter & " OR " & sFilterColName & " IS NULL"
                End If
          
              Case giFILTEROP_NOTEQUALTO
                sSubFilter = sFilterColName & " <> " & sFilterValue
                If Val(sFilterValue) = 0 Then
                  sSubFilter = sSubFilter & " AND " & sFilterColName & " IS NOT NULL"
                End If
          
              Case giFILTEROP_ISATMOST
                sSubFilter = sFilterColName & " <= " & sFilterValue
                If Val(sFilterValue) >= 0 Then
                  sSubFilter = sSubFilter & " OR " & sFilterColName & " IS NULL"
                End If
          
              Case giFILTEROP_ISATLEAST
                sSubFilter = sFilterColName & " >= " & sFilterValue
                If Val(sFilterValue) <= 0 Then
                  sSubFilter = sSubFilter & " OR " & sFilterColName & " IS NULL"
                End If
          
              Case giFILTEROP_ISMORETHAN
                sSubFilter = sFilterColName & " > " & sFilterValue
                If Val(sFilterValue) < 0 Then
                  sSubFilter = sSubFilter & " OR " & sFilterColName & " IS NULL"
                End If
          
              Case giFILTEROP_ISLESSTHAN
                sSubFilter = sFilterColName & " < " & sFilterValue
                If Val(sFilterValue) > 0 Then
                  sSubFilter = sSubFilter & " OR " & sFilterColName & " IS NULL"
                End If
            End Select
      
          Case sqlDate  ' Date columns.
            If LenB(sFilterValue) <> 0 Then
              dtDateValue = CDate(sFilterValue)
              sFilterValue = Replace(Format(dtDateValue, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
            End If
            
            Select Case mavFilterCriteria(2, iLoop)
              Case giFILTEROP_ON
                If LenB(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL"
                Else
                  sSubFilter = sFilterColName & " = '" & sFilterValue & "'"
                End If
            
              Case giFILTEROP_NOTON
                If LenB(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sFilterColName & " <> '" & sFilterValue & "'"
                End If
              
              Case giFILTEROP_ONORBEFORE
                If LenB(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL"
                Else
                  sSubFilter = sFilterColName & " <= '" & sFilterValue & "' OR " & sFilterColName & " IS NULL"
                End If
            
              Case giFILTEROP_ONORAFTER
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL OR " & sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sFilterColName & " >= '" & sFilterValue & "'"
                End If
            
              Case giFILTEROP_AFTER
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sFilterColName & " > '" & sFilterValue & "'"
                End If
            
              Case giFILTEROP_BEFORE
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL AND " & sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sFilterColName & " < '" & sFilterValue & "' OR " & sFilterColName & " IS NULL"
                End If
            End Select
          
          Case sqlVarChar, sqlVarBinary, sqlLongVarChar  ' Character and Photo columns (photo columns are really character columns).
            Select Case mavFilterCriteria(2, iLoop)
              Case giFILTEROP_IS
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " = '' OR " & sFilterColName & " IS NULL"
                Else
                  ' Replace the standard * and ? characters with the SQL % and _ characters.
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
                      Case "*"
                        sModifiedFilterValue = sModifiedFilterValue & "%"
                      Case "?"
                        sModifiedFilterValue = sModifiedFilterValue & "_"
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " LIKE '" & sModifiedFilterValue & "'"
                End If
            
              Case giFILTEROP_ISNOT
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " <> '' and " & sFilterColName & " IS NOT NULL"
                Else
                  ' Replace the standard * and ? characters with the SQL % and _ characters.
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
                      Case "*"
                        sModifiedFilterValue = sModifiedFilterValue & "%"
                      Case "?"
                        sModifiedFilterValue = sModifiedFilterValue & "_"
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " NOT LIKE '" & sModifiedFilterValue & "'"
                End If
            
              Case giFILTEROP_CONTAINS
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL OR " & sFilterColName & " IS NOT NULL"
                Else
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " LIKE '%" & sModifiedFilterValue & "%'"
                End If
            
              Case giFILTEROP_DOESNOTCONTAIN
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL AND " & sFilterColName & " IS NOT NULL"
                Else
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " NOT LIKE '%" & sModifiedFilterValue & "%'"
                End If
            End Select
        End Select
  
        Exit For
      End If
    Next objColumn
    Set objColumn = Nothing
    
    ' Add this filter criterion definition to the full global definition string.
    sFilter = sFilter & IIf(Len(sFilter) > 0, " AND (", "(") & sSubFilter & ")"
  Next iLoop
  
  ' Return the filter string.
  GetFilter = sFilter
  
End Function
Private Function GetSelectString() As String
  ' Return the Select string for getting the current recordset.
  Dim sSelect As String
  Dim objColumn As CColumnPrivilege
  
  sSelect = ""
  
  If mcolColumnPrivileges Is Nothing Then Exit Function
  
  For Each objColumn In mcolColumnPrivileges
    If objColumn.AllowSelect Then
       
'      ' JDM - Don't read in embedded objects as the size could seriously affect performance
'      If (objColumn.DataType = sqlOle And objColumn.OLEType = OLE_EMBEDDED) Or _
'        (objColumn.DataType = sqlVarBinary = objColumn.OLEType = OLE_EMBEDDED) Then
'        sSelect = sSelect & IIf(Len(sSelect) > 0, ", ", "") & "'" & objColumn.ColumnID & "' AS " & objColumn.ColumnName
'      Else
'        sSelect = sSelect & IIf(Len(sSelect) > 0, ", ", "") & mobjTableView.RealSource & "." & objColumn.ColumnName
'      End If

      ' JDM - Don't read in embedded objects as the size could seriously affect performance
      If objColumn.DataType = sqlOle And objColumn.DataType = sqlVarBinary Then
        sSelect = sSelect & IIf(LenB(sSelect) > 0, ", ", "") & "'" & objColumn.ColumnID & "' AS " & objColumn.ColumnName
      Else
        sSelect = sSelect & IIf(LenB(sSelect) > 0, ", ", "") & mobjTableView.RealSource & "." & objColumn.ColumnName
      End If

    End If
  Next objColumn
  Set objColumn = Nothing

  If LenB(sSelect) > 0 Then
    sSelect = sSelect & _
      ", CONVERT(integer," & mobjTableView.RealSource & ".TimeStamp) AS TimeStamp"
  End If
  
  GetSelectString = sSelect
    
End Function

Private Sub SetupColumnPrivileges()
  ' Call the function to return the column privileges collection for the given table.
  Set mcolColumnPrivileges = GetColumnPrivileges(msTableViewName)
  
End Sub

Public Property Get RecordCount() As Long
  ' Return the number of records in the recordset.
  RecordCount = mlngRecordCount
  
End Property



Public Property Get ParentViewID() As Long
  ParentViewID = mlngParentViewID

End Property


Public Property Get GetRecordCount() As Long
  ' Return the number of records in the recordset.
  Dim rsTemp As ADODB.Recordset

  Set rsTemp = New ADODB.Recordset
  rsTemp.Open msCountSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  GetRecordCount = rsTemp(0).Value
  rsTemp.Close
  Set rsTemp = Nothing

End Property


Public Property Get Recordset() As ADODB.Recordset
  ' Return the recordset for the form.
  Set Recordset = mrsRecords

End Property

Private Sub mnuOLEDelete_Click()
  If TypeOf Me.ActiveControl Is OLE Then
    Me.ActiveControl.Delete
    Me.ActiveControl.Class = vbNullString
  End If

End Sub


Private Sub mnuOLEEdit_Click()
  If TypeOf Me.ActiveControl Is OLE Then
    Me.ActiveControl.DoVerb vbOLEOpen
  End If

End Sub


Private Sub mnuOLEInsert_Click()
  If TypeOf Me.ActiveControl Is OLE Then
    Me.ActiveControl.InsertObjDlg
  End If

End Sub


Private Sub mnuOLEPaste_Click()
  If TypeOf Me.ActiveControl Is OLE Then
    If Me.ActiveControl.PasteOK Then
      Me.ActiveControl.PasteSpecialDlg
    End If
  End If

End Sub


Private Sub OLE1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As COA_OLE
  Dim fDataChanged As Boolean
  Dim fOleOnServer As Boolean
  Dim bEmbedded As Boolean
  Dim sFile As String
  Dim fOLEOK As Boolean
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim frmTemp As Form
    
  ' JPD20030115 Fault 4862
  ' Do not action the click event if there are other
  ' recEdit forms that have not had their changes saved.
  ' Set focus to this control to ensure the menu is updated
  ' correctly when the control is right-clicked on, and to fire
  ' the changed recEdit form's 'deactivate' event.
  ' NB. This code is put in the 'click' event rather than 'onFocus'
  ' as right-click fires the 'click' event, but doesn't fire the
  ' 'onFocus' event.
  If Not mfTableEntry Then
    OLE1(Index).SetFocus
    For Each frmTemp In Forms
      If (TypeOf frmTemp Is frmRecEdit4) Then
        If Not (frmTemp Is Me) Then
          If frmTemp.Changed Then
            Exit Sub
          End If
        End If
      End If
    Next frmTemp
    Set frmTemp = Nothing
  End If
  
  fOLEOK = True
  
  If Not mfLoading Then

    ' Get the control's tag.
    sTag = OLE1(Index).Tag
    
    ' If embedded document we'll need to read from the database
    If OLE1(Index).OLEType = OLE_EMBEDDED Or OLE1(Index).OLEType = OLE_UNC Then
    
      ' Read from the database if no stream is open
      If Not OLE1(Index).EmbeddedStream.State = adStateOpen Then
        fOLEOK = ReadStream(OLE1(Index), mobjScreenControls.Item(sTag), False)
      Else
        ' If only the header has been read, re-read full stream from database
        If OLE1(Index).OLEType = OLE_EMBEDDED And OLE1(Index).EmbeddedStream.Size <= 400 Then
          fOLEOK = ReadStream(OLE1(Index), mobjScreenControls.Item(sTag), False)
        Else
          fOLEOK = True
        End If
      End If
    
    Else
     
      If Len(sTag) > 0 Then
    
        ' Is this OLE column one that is located on the server ?
        fOleOnServer = (mobjScreenControls.Item(sTag).OLEType = OLE_SERVER)
      
        If (fOleOnServer = True) Then
          'JPD 20030828 Fault 4285
          If (Len(gsOLEPath) = 0) Or (Dir(gsOLEPath & IIf(Right(gsOLEPath, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString) Then
            fOLEOK = False
            COAMsgBox "Unable to edit OLE fields." & vbNewLine & _
                 "The OLE path has not been defined, or is invalid." & vbNewLine & _
                 "Please set the Server OLE Path in PC Configuration." _
                 , vbExclamation + vbOKOnly, app.ProductName
          End If
        Else
          'JPD 20030828 Fault 4285
          If (Len(gsLocalOLEPath) = 0) Or (Dir(gsLocalOLEPath & IIf(Right(gsLocalOLEPath, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString) Then
            fOLEOK = False
            COAMsgBox "Unable to edit this OLE field." & vbNewLine & _
                 "The Local OLE path has not been defined, or is invalid." & vbNewLine & _
                 "Please set the Local OLE Path in PC Configuration." _
                 , vbExclamation + vbOKOnly, app.ProductName
          End If
        End If
     
      End If
    End If

    If fOLEOK Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Check if it's a OLE control, if so load the select OLE form.
      'If mrsRecords.Fields(sThisControlsColumnName).Type = adVarChar Then
      If OLE1(Index).OLEType = OLE_EMBEDDED Or OLE1(Index).OLEType = OLE_UNC Then
        With frmSelectEmbedded
        
          '.Initialise mrsRecords.Fields(sThisControlsColumnName)
          .OLEType = OLE1(Index).OLEType
          .EmbeddedEnabled = mobjScreenControls.Item(sTag).EmbeddedEnabled
          .MaxOLESize = mobjScreenControls.Item(sTag).MaxOLESize
          .IsPhoto = False
          .IsReadOnly = mobjScreenControls.Item(sTag).ReadOnly Or Not mcolColumnPrivileges.Item(mobjScreenControls.Item(sTag).ColumnName).AllowUpdate
          .Initialise OLE1(Index).EmbeddedStream
          .Show vbModal
        
          Select Case .Selection
            Case optSelect
              OLE1(Index).OLEType = .OLEType
              OLE1(Index).EmbeddedStream = .EmbeddedFile
              OLE1(Index).FileName = LoadFileNameFromStream(OLE1(Index).EmbeddedStream, (.OLEType = OLE_UNC), False)
              OLE1(Index).ToolTipText = OLE1(Index).FileName
              fDataChanged = True
              
              fFound = False
              For iLoop = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              If Not fFound Then
                ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
              End If
              
            Case optCancel
              fDataChanged = False
            
            Case optNone
            
              fFound = False
              For iLoop = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              If Not fFound Then
                ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
              End If
            
              OLE1(Index).EmbeddedStream = .EmbeddedFile
              fDataChanged = True

          End Select
        
          Unload frmSelectEmbedded
        
        End With


      
      Else
      
        mfControlProcessing = True
      
        With frmSelectOLE
          
          ' RH 03/08/00 - FAULT xxx by ME. Oles optionally held on server
          ' If the column definition for this column states that the OLE
          ' object is held on the local machine, and the field contains
          ' an objects filename, and the sourcelink of the OLE field is
          ' nullstring, then it means the file does not exist on the
          ' current users machine, so let them know.
'            If fOleOnServer = False And _
'               Not IsNull(mrsRecords.Fields(sThisControlsColumnName)) And _
'               OLE1(Index).SourceDoc = vbNullString Then
          If fOleOnServer = False And _
             Not IsNull(mrsRecords.Fields(sThisControlsColumnName)) And _
             OLE1(Index).FileName = vbNullString Then
            
            'JPD 20040209 Fault ????
            If Len(mrsRecords.Fields(sThisControlsColumnName)) > 0 Then
              COAMsgBox "This field contains the OLE object below, however, the column" & vbNewLine & _
                     "definition states OLE objects for this column are held locally" & vbNewLine & _
                     "and the object does not exist on your machine." & vbNewLine & vbNewLine & _
                     Replace(gsLocalOLEPath & "\" & mrsRecords.Fields(sThisControlsColumnName), "\\", "\") _
                     , vbInformation + vbOKOnly, Application.Name
            End If
          End If
          
          .OleOnServer = fOleOnServer
'            .Initialise Mid(OLE1(Index).SourceDoc, InStrRev(OLE1(Index).SourceDoc, "\") + 1)
          .IsReadOnly = mobjScreenControls.Item(sTag).ReadOnly Or Not mcolColumnPrivileges.Item(mobjScreenControls.Item(sTag).ColumnName).AllowUpdate
          .Initialise Mid(OLE1(Index).FileName, InStrRev(OLE1(Index).FileName, "\") + 1)
          .Show vbModal
          
          Select Case .optSelection
            Case optSelect
              OLE1(Index).ToolTipText = .OLEFileName
              OLE1(Index).FileName = Mid(.OLEFileName, InStrRev(.OLEFileName, "\") + 1)
              OLE1(Index).OleOnServer = fOleOnServer

'                OLE1(Index).CreateLink .OLEFileName
              mfDataChanged = True
              fDataChanged = True
              
              ' JPD20021007 Fault 4498
              fFound = False
              For iLoop = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              If Not fFound Then
                ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
              End If

            Case optNone
              OLE1(Index).ToolTipText = ""
              OLE1(Index).FileName = vbNullString
              OLE1(Index).OleOnServer = fOleOnServer
              
'              mfDataChanged = True
              fDataChanged = True
              
              ' JPD20021007 Fault 4498
              fFound = False
              For iLoop = 1 To UBound(malngChangedOLEPhotos)
                If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
                  fFound = True
                  Exit For
                End If
              Next iLoop
              If Not fFound Then
                ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
                malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
              End If

            Case Else
              fDataChanged = False
          End Select

          Unload frmSelectOLE
        End With
        
      End If
    
    
      If fDataChanged Then
        ' Update all other screen control's that represent the same column.
        For Each objControl In OLE1
          If objControl.Index <> Index Then
            sTag = objControl.Tag

            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID

              If lngOtherControlsColumnID = lngThisControlsColumnID Then

                objControl.FileName = OLE1(Index).FileName
                objControl.OLEType = OLE1(Index).OLEType
                
                If objControl.OLEType = OLE_EMBEDDED Or objControl.OLEType = OLE_UNC Then
                  
                  If objControl.EmbeddedStream.State = adStateClosed Then
                    objControl.EmbeddedStream.Open
                    objControl.EmbeddedStream.Type = adTypeBinary
                  End If
                  
                  objControl.EmbeddedStream = New ADODB.Stream
                  objControl.EmbeddedStream.Type = adTypeBinary
                  objControl.EmbeddedStream.Open
                  
                  OLE1(Index).EmbeddedStream.Position = 0
                  OLE1(Index).EmbeddedStream.CopyTo objControl.EmbeddedStream
                  
                End If

              End If
            End If
          End If
        Next objControl
        Set objControl = Nothing
      End If
    
    
    
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = fDataChanged
    End If
    
    frmMain.RefreshMainForm Me
  End If

  Exit Sub
  
Err_Trap:
  Select Case Err.Number
    Case 440
      mfDataChanged = True

    ' JPD20020828 Fault 4176
    Case 52, 53, 75
      fOLEOK = False
      Resume Next
      
    Case Else
  End Select
  
  mfDataChanged = True

  ' JPD20021007 Fault 4498
  fFound = False
  For iLoop = 1 To UBound(malngChangedOLEPhotos)
    If malngChangedOLEPhotos(iLoop) = lngThisControlsColumnID Then
      fFound = True
      Exit For
    End If
  Next iLoop
  If Not fFound Then
    ReDim Preserve malngChangedOLEPhotos(UBound(malngChangedOLEPhotos) + 1)
    malngChangedOLEPhotos(UBound(malngChangedOLEPhotos)) = lngThisControlsColumnID
  End If

End Sub

Private Sub OLE1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck OLE1(Index)
  
End Sub


Private Sub OLE1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then OLE1_Click (Index)
End Sub

Private Sub OLE1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'  Dim fEmpty As Boolean
'
'  fEmpty = (Len(Trim(OLE1(Index).Class)) = 0)
'
'  Select Case KeyCode
'    Case vbKeyInsert
'      If fEmpty Then
'        OLE1(Index).InsertObjDlg
'      End If
'
'    Case vbKeyReturn
'      If Not fEmpty Then
'        OLE1(Index).DoVerb vbOLEOpen
'      End If
'
'    Case vbKeyDelete
'      If Not fEmpty Then
'        OLE1(Index).Delete
'        OLE1(Index).Class = vbNullString
'      End If
'  End Select

End Sub


Private Sub OLE1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(OLE1(Index))

End Sub

Private Sub OptionGroup1_Click(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
  Dim frmTemp As Form
    
  ' JPD20030115 Fault 4862
  ' Do not action the click event if there are other
  ' recEdit forms that have not had their changes saved.
  ' Set focus to this control to ensure the menu is updated
  ' correctly when the control is right-clicked on, and to fire
  ' the changed recEdit form's 'deactivate' event.
  ' NB. This code is put in the 'click' event rather than 'onFocus'
  ' as right-click fires the 'click' event, but doesn't fire the
  ' 'onFocus' event.
  If Not mfTableEntry Then
    OptionGroup1(Index).SetFocus
    For Each frmTemp In Forms
      If (TypeOf frmTemp Is frmRecEdit4) Then
        If Not (frmTemp Is Me) Then
          If frmTemp.Changed Then
            OptionGroup1(Index).Value = mvOldValue
            Exit Sub
          End If
        End If
      End If
    Next frmTemp
    Set frmTemp = Nothing
  End If
  
  If Not mfLoading Then
    ' Get the control's tag.
    sTag = OptionGroup1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In OptionGroup1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .Value = OptionGroup1(Index).Value
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (OptionGroup1(Index).Value <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))

      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True

    frmMain.RefreshMainForm Me
  End If

End Sub

Private Sub OptionGroup1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck OptionGroup1(Index)

End Sub


Private Sub OptionGroup1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(OptionGroup1(Index))

End Sub


Private Sub Spinner1_Change(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
  Static fChanging As Boolean 'Indicates if we are already changing a spinner
                              'to avoid triggering the other spinners change
                              'event and so stopping circular looping
  Dim frmTemp As Form
    
  If (Not mfLoading) And (Not fChanging) Then
    ' JPD20030115 Fault 4862
    ' Do not action the click event if there are other
    ' recEdit forms that have not had their changes saved.
    ' Set focus to this control to ensure the menu is updated
    ' correctly when the control is right-clicked on, and to fire
    ' the changed recEdit form's 'deactivate' event.
    ' NB. This code is put in the 'click' event rather than 'onFocus'
    ' as right-click fires the 'click' event, but doesn't fire the
    ' 'onFocus' event.
    If Not mfTableEntry Then
      Spinner1(Index).SetFocus
      For Each frmTemp In Forms
        If (TypeOf frmTemp Is frmRecEdit4) Then
          If Not (frmTemp Is Me) Then
            If frmTemp.Changed Then
              Spinner1(Index).Text = mvOldValue
              Exit Sub
            End If
          End If
        End If
      Next frmTemp
      Set frmTemp = Nothing
    End If
    
    fChanging = True
    
    ' Get the control's tag.
    sTag = Spinner1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In Spinner1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .Text = Spinner1(Index).Text
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    fChanging = False
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (Spinner1(Index).Text <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))

      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True

    frmMain.RefreshMainForm Me
  End If
  
End Sub

Private Sub Spinner1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck Spinner1(Index)

End Sub


Private Sub Spinner1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(Spinner1(Index))

End Sub





Private Sub TabStrip1_BeforeClick(Cancel As Integer)
  Dim frmTemp As Form
    
  ' JPD20030115 Fault 4862
  ' Do not action the click event if there are other
  ' recEdit forms that have not had their changes saved.
  ' Set focus to this control to fire
  ' the changed recEdit form's 'deactivate' event.
  If Not mfTableEntry Then
    If TabStrip1.Visible Then
      TabStrip1.SetFocus
      For Each frmTemp In Forms
        If (TypeOf frmTemp Is frmRecEdit4) Then
          If Not (frmTemp Is Me) Then
            If frmTemp.Changed Then
              Cancel = True
              Exit Sub
            End If
          End If
        End If
      Next frmTemp
      Set frmTemp = Nothing
    End If
  End If

End Sub

Private Sub TabStrip1_Click()
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim objControl As Control
  
  If Not mfLoading Then
  
    ' Lock the window refreshing.
    UI.LockWindow Me.hWnd
  
    ' Get the index of the selected tabpage.
    iIndex = TabStrip1.SelectedItem.Index
    
    If iIndex > 0 Then
      For iLoop = 1 To fraTabPage.UBound
        If iLoop = iIndex Then
          fraTabPage(iLoop).Visible = True
          
          For Each objControl In Me.Controls
          
            If Not TypeOf objControl Is ActiveBar And Not TypeOf objControl Is Menu And Not TypeOf objControl Is COA_ColourPicker Then
          
              If objControl.Container Is fraTabPage(iLoop) Then
                If Not objControl.Visible And LenB(objControl.Tag) > 0 Then
                  objControl.Visible = True
                End If
              End If
          
            End If
          Next
          
          fraTabPage(iLoop).ZOrder 0
        Else
          fraTabPage(iLoop).Visible = False
        End If
      Next
    End If
  
    StatusBar1.SimpleText = TabStrip1.SelectedItem.Caption
      
    ' Refresh any navigation controls because Version One has some teething troubles
    For Each objControl In Me.Controls
      If TypeOf objControl Is COA_Navigation Then
        objControl.RefreshControls
      End If
    Next
    DoEvents
      
    ' Unlock the window refreshing.
    UI.UnlockWindow
  End If

End Sub


'MH20001016 Fault 652
'Fix status bar caption.
Private Sub TabStrip1_GotFocus()
  'JPD 20031008 Fault 7080
  If TabStrip1.SelectedItem Is Nothing Then
    StatusBar1.SimpleText = ""
  Else
    StatusBar1.SimpleText = TabStrip1.SelectedItem.Caption
  End If

End Sub

Private Sub TDBMask1_Change(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If (Not mfLoading) Then
    ' Get the control's tag.
    sTag = TDBMask1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In TDBMask1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                mfLoading = True
                .Text = TDBMask1(Index).Text
                mfLoading = False
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (TDBMask1(Index).Value <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))
      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True

    frmMain.RefreshMainForm Me
  End If

End Sub

Private Sub TDBMask1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck TDBMask1(Index)

End Sub


Private Sub GotFocusCheck(pctlControl As VB.Control)
  On Error GoTo ErrorTrap
  
  ' Only try to evaluate the control's column properties if we can find the
  ' associated column (using the control's tag as the key into the control collection).
  If Len(pctlControl.Tag) > 0 Then
    With pctlControl
      ' Update the status bar display.
      StatusBar1.SimpleText = Trim(mobjScreenControls.Item(.Tag).StatusBarMessage)
        
      ' Remember the original value in case we need to restore it.
      If TypeOf pctlControl Is TDBMask6Ctl.TDBMask Then
        mvOldValue = .Text
      ElseIf TypeOf pctlControl Is TextBox Then
        mvOldValue = .Text
      ElseIf TypeOf pctlControl Is COA_Spinner Then
        mvOldValue = .Text
      ElseIf TypeOf pctlControl Is XtremeSuiteControls.CheckBox Then
        mvOldValue = .Value
      ElseIf TypeOf pctlControl Is COA_Lookup Then
        mvOldValue = .Text
      'JPD 20050302 Fault 9847
      ElseIf (TypeOf pctlControl Is TDBNumberCtrl.TDBNumber) Or _
        (TypeOf pctlControl Is TDBNumber6Ctl.TDBNumber) Then
        mvOldValue = .Value
      ElseIf TypeOf pctlControl Is GTMaskDate.GTMaskDate Then
        mvOldValue = .Text
      ElseIf TypeOf pctlControl Is COA_Image Then
        mvOldValue = .ASRDataField
      ElseIf TypeOf pctlControl Is COA_OptionGroup Then
        mvOldValue = .Value
      ElseIf TypeOf pctlControl Is XtremeSuiteControls.ComboBox Then
        mvOldValue = .ListIndex
      ElseIf TypeOf pctlControl Is OLE Then
        mvOldValue = BackupOLEData(pctlControl)
      ElseIf TypeOf pctlControl Is TDBText6Ctl.TDBText Then
        mvOldValue = .Text
      ElseIf TypeOf pctlControl Is COA_WorkingPattern Then
        mvOldValue = .Value
      ElseIf TypeOf pctlControl Is COA_ColourSelector Then
        mvOldValue = CStr(.BackColor)
      Else
        mvOldValue = vbNullString
      End If
    End With
  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  COAMsgBox Err.Description & " - GotFocusCheck", vbCritical
  Resume TidyUpAndExit
        
End Sub

Public Sub Requery(pfReset As Boolean, Optional pvCallingFormID As Variant)
  ' Requery (refresh the current recordset) the database.
  Dim lngOldRecordID As Long

  Screen.MousePointer = vbHourglass
  
  lngOldRecordID = mlngRecordID

  If pfReset Then
    ' Refresh the recordset and goto the first record.
    If Not RefreshRecordset Then
      Exit Sub
    End If
    
    If mrsRecords.EditMode <> adEditAdd Then
      mrsRecords.MoveFirst
    End If
  Else
    ' Save changes if required and valid.
    If SaveChanges Then
      If Database.Validation Then
        ' Refresh the recordset and locate the current record.
        If Not RefreshRecordset Then
          Exit Sub
        End If

        If mrsRecords.EditMode <> adEditAdd Then
          LocateRecord mlngRecordID
          
          'JPD 20041109 Fault 9008
          'NPG20080509 Fault 12938 - check if user cancelled entering a new record too.
          'If mrsRecords!ID <> lngOldRecordID Then
          If mrsRecords!ID <> lngOldRecordID And lngOldRecordID <> 0 Then
            If Filtered Then
              COAMsgBox "The '" & Replace(mobjTableView.TableName, "_", " ") & "' record does not satisfy the current filter.", vbExclamation, app.ProductName
            ElseIf Me.ViewID > 0 Then
              COAMsgBox "The '" & Replace(mobjTableView.TableName, "_", " ") & "' record is no longer in the current view.", vbExclamation, app.ProductName
            End If
          End If
        Else
          'JPD 20041109 Fault 9008
          If Filtered Then
            COAMsgBox "The '" & Replace(mobjTableView.TableName, "_", " ") & "' record does not satisfy the current filter.", vbExclamation, app.ProductName
          
          'MH20060619 Fault 11236
          'Supress message if no records as the "has been deleted" message should appear instead.
          'ElseIf Me.ViewID > 0 Then
          ElseIf Me.ViewID > 0 And mrsRecords.RecordCount > 0 Then
            COAMsgBox "The '" & Replace(mobjTableView.TableName, "_", " ") & "' record is no longer in the current view.", vbExclamation, app.ProductName
          End If
        End If
      End If
    End If
  End If

  
  'MH20001018 Fault 1124
  'Don't bother doing all this refresh stuff if cancelled save changes
  If Not mfCancelled Then
    ' Update all controls and associated screens.
    ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
    ' UpdateChildren call as the summary fields that are updated in
    ' UpdateChildren may be dependent on the parent recordset being
    ' refreshed first (in UpdateParentWindow).
    UpdateParentWindow pvCallingFormID
    UpdateControls
    
    'JPD 20030410 Fault 5315
    If lngOldRecordID = mlngRecordID Then
      UpdateChildren pvCallingFormID
    Else
      UpdateChildren
    End If
    
    'UpdateChildren does all of the UpdateFindWindow stuff so don't bother calling it again...
    'UpdateFindWindow
'    UpdateParentWindow pvCallingFormID

    ' Refresh the menu.
    frmMain.RefreshMainForm Me
  End If
  
  Screen.MousePointer = vbDefault
  
End Sub

Public Sub SelectOrder()
  ' Gives the user the choice of which sort order
  ' to display the records in by using the frmDefSel form.
  Dim lngOldID As Long
  Dim sSQL As String

  If SaveChanges Then
    If Database.Validation Then
    
      ' Reference Property instead of object to trap errors
      If Me.ViewID > 0 Then

        sSQL = "SELECT DISTINCT ASRSysOrders.name, ASRSysOrders.orderID" & _
          " FROM ASRSysOrders" & _
          " INNER JOIN ASRSysOrderItems ON ASRSysOrders.orderID = ASRSysOrderItems.orderID" & _
          " INNER JOIN ASRSysViewColumns ON ASRSysOrderItems.columnID = ASRSysViewColumns.columnID" & _
          " WHERE ASRSysOrders.tableID = " & Me.TableID & _
          " AND ASRSysViewColumns.inView = 1" & _
          " AND ASRSysOrderItems.type = 'O'" & _
          " AND ASRSysViewColumns.viewID = " & Me.ViewID & _
          " AND ASRSysOrders.type = 1"
      Else
        sSQL = "SELECT name, orderID" & _
          " FROM ASRSysOrders" & _
          " WHERE tableID=" & Me.TableID & _
          " AND type = 1"
      End If
      
      With frmDefSel
        .SelectedUtilityType = utlOrder
        .Caption = "Select Order"
        .Options = edtSelect
        .EnableRun = False
        .TableComboEnabled = False
        .TableComboVisible = True
        .HideDescription = True
        .TableID = Me.TableID
        .SelectedID = mlngOrderID
               
        If Not .ShowList(utlOrder) Then
          Exit Sub
        End If
      
        .Show vbModal
      
        If (.Action = edtSelect) And (mlngOrderID <> frmDefSel.SelectedID) Then
          Screen.MousePointer = vbHourglass
          mlngOrderID = frmDefSel.SelectedID
            
          If mrsRecords.EditMode = adEditAdd Then
            mrsRecords.CancelUpdate
          End If
          
          mrsRecords.Close
          Set mrsRecords = Nothing
          GetRecords
          
            
          ' Check if the refreshed recordset is still empty.
          If mrsRecords.BOF And mrsRecords.EOF Then
            ' Create a new record if permitted.
            If mobjTableView.AllowInsert Then
              mrsRecords.AddNew
            Else
              ' JPD20030311 Fault 5138
              If Me.Visible Then
                COAMsgBox "This " & IIf(Me.ViewID > 0, "view", "table") & " is empty" & _
                  " and you do not have 'new' permission on it.", vbExclamation, "Security"
              End If
              Unload Me
              Screen.MousePointer = vbDefault
              Exit Sub
            End If
          Else
            ' Select the originally selected record if it is still in the recordset.
            LocateRecord mlngRecordID
          End If
            
          ' Refresh the screen controls in this screen.
          ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
          ' UpdateChildren call as the summary fields that are updated in
          ' UpdateChildren may be dependent on the parent recordset being
          ' refreshed first (in UpdateParentWindow).
          UpdateParentWindow
          UpdateControls
          UpdateChildren
          UpdateFindWindow
            
          Screen.MousePointer = vbDefault
        End If
      End With
      Set frmDefSel = Nothing
  
      ' Refresh the main menu.
      frmMain.RefreshMainForm Me
    End If
  End If

End Sub

Public Sub ClearFilter()
  ' RH 14/07/00
  ' Remove all filters to the current recedit screen
  If SaveChanges Then
    If Database.Validation Then
      If Me.Filtered Then
        ' Clear the filter.
        ReDim mavFilterCriteria(3, 0)
        
        'TM20020107 Fault 1379
        'mrsRecords.Close
        Set mrsRecords = Nothing
        GetRecords

        If mrsRecords.BOF And mrsRecords.EOF Then
          ' The refreshed recordset is empty. Create a new record if permitted.
          If mobjTableView.AllowInsert Then
            mrsRecords.AddNew
          Else
            ' JPD20030311 Fault 5138
            If Me.Visible Then
              'MH20031002 Fault 7082 Reference Property instead of object to trap errors
              'COAMsgBox "This " & IIf(mobjTableView.ViewID > 0, "view", "table") & " is now empty" & _
                " and you do not have 'new' permission on it.", vbExclamation, "Security"
              COAMsgBox "This " & IIf(Me.ViewID > 0, "view", "table") & " is now empty" & _
                " and you do not have 'new' permission on it.", vbExclamation, "Security"
            End If
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
          End If
        Else
          ' Locate the record if it is still in the recordset.
          LocateRecord mlngRecordID
        End If
        
        ' Refresh the screen controls in this screen.
        ' JPD 21/02/2001 Moved the UpdateParentWindow call before the
        ' UpdateChildren call as the summary fields that are updated in
        ' UpdateChildren may be dependent on the parent recordset being
        ' refreshed first (in UpdateParentWindow).
        UpdateParentWindow
        UpdateControls
        UpdateChildren
        UpdateFindWindow
'        UpdateParentWindow
      End If
    End If
  End If
  
  ' Refresh the main menu.
  frmMain.RefreshMainForm Me
  
End Sub

Public Sub SelectFilter()
  ' Display the record filter definition form, and apply the defined filter to the current recordset.
  ' Save changes first if required.
  Dim lngOldID As Long
  Dim fFilterScreen As Boolean
  Dim fCancelled As Boolean
  
  fCancelled = False
  
  ' Save changes if necessary.
  If SaveChanges Then
    If Database.Validation Then
    
      ' Display the filter selection form.
      With frmRecordFilter
        
        fFilterScreen = True
        Do While fFilterScreen
        
          .Initialise mrsRecords, mobjTableView, mavFilterCriteria
          .Show vbModal

          If .Cancelled Then
            fCancelled = True
            Exit Do
          End If

          ' Apply the filter.
          mavFilterCriteria = frmRecordFilter.FilterArray
          
          ' Remember the currrent record.
          lngOldID = mlngRecordID
          
          ' Get the filtered recordset.
          ' JPD20030311 Fault 5138
          If Not (mrsRecords.BOF And mrsRecords.EOF) Then
            If mrsRecords.EditMode = adEditAdd Then
              mrsRecords.CancelUpdate
            End If
          End If
          
          mrsRecords.Close
          Set mrsRecords = Nothing
          GetRecords
          
          'TM20020107 Fault 1379
          'TM20020204 Fault 3416 - Clear the filter if no records AND is a parent table.
          If mrsRecords.BOF And mrsRecords.EOF And (Me.ParentTableID = 0) Then
            COAMsgBox "No records match the current filter." & vbNewLine & _
                   "No filter is applied.", vbInformation + vbOKOnly, app.ProductName

            ReDim mavFilterCriteria(3, 0)

            mrsRecords.Close
            Set mrsRecords = Nothing
            GetRecords

          Else
            fFilterScreen = False
          
          End If
          
          ' IMPORTANT : FIX PRODUCED WIERD ERRORS WHEN SELECTING NO TO THE COAMsgBox
          ' THEREFORE REMOVED FIX AND ORIGINAL CODE IS NOW IN PLACE, AS ABOVE
          
          ' RH 13/09/00 - SUG 635 - could be a bit dodgy, but worked ok in testing!
          ' If no records match the filter, then clear it.
          
'          If mrsRecords.BOF And mrsRecords.EOF Then
'            If COAMsgBox("No records match the current filter." & vbNewLine & _
'              "Would you like to define another filter?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'              ReDim mavFilterCriteria(3, 0)
'              mrsRecords.Close
'              Set mrsRecords = Nothing
'              GetRecords
'            Else
'              SelectFilter
'              Exit Sub
'            End If
'          End If
         
          ' Check if the refreshed recordset is still empty.
          If mrsRecords.BOF And mrsRecords.EOF Then
            ' Create a new record if permitted.
            If mobjTableView.AllowInsert Then
              mrsRecords.AddNew
            Else
              ' JPD20030311 Fault 5138
              If Me.Visible Then
                'MH20031002 Fault 7082 Reference Property instead of object to trap errors
                'COAMsgBox "This " & IIf(mobjTableView.ViewID > 0, "view", "table") & " is empty" & _
                  " and you do not have 'new' permission on it.", vbExclamation, "Security"
                COAMsgBox "This " & IIf(Me.ViewID > 0, "view", "table") & " is empty" & _
                  " and you do not have 'new' permission on it.", vbExclamation, "Security"
              End If
              Unload Me
              Exit Sub
            End If
          Else
            ' Select the originally selected record if it is still in
            ' the recordset.
            LocateRecord lngOldID
          End If
          
        Loop
      
        If Not fCancelled Then
          UpdateAll
        End If
      End With
        
      Unload frmRecordFilter
      
      ' Refresh the main menu.
      frmMain.RefreshMainForm Me
    End If
  End If
  
End Sub


Public Property Get AllowDelete() As Boolean
  ' Return whether or not the user can delete records.
  If mobjTableView Is Nothing Then
    AllowDelete = False
  Else
    AllowDelete = mobjTableView.AllowDelete
  End If
  
End Property



Public Property Get AllowUpdate() As Boolean
  ' Return whether or not the user can update records.
  If mobjTableView Is Nothing Then
    AllowUpdate = False
  Else
    AllowUpdate = mobjTableView.AllowUpdate
  End If
  
End Property

Public Property Get AutoUpdateLookup() As Boolean
  AutoUpdateLookup = mblnScreenHasAutoUpdate
End Property

Public Sub DeleteRecord(Optional pfPrompt As Variant, Optional pfFromFind As Boolean)
  ' Deletes the current record.
  On Error GoTo ErrorTrap
  
  Dim fGoLast As Boolean
  Dim fDeleteOK As Boolean
  Dim iChangeReason As Integer
  Dim lngRecordID As Long
  Dim sSQL As String
  Dim sErrorMsg As String
  Dim sRecordDescription As String
  Dim sTemp As String
  
  fDeleteOK = True
   
  If IsMissing(pfPrompt) Then pfPrompt = True
  
  ' Check that the user can delete from the current table/view.
  If (Not mobjTableView.AllowDelete) Or _
    (mrsRecords.BOF And mrsRecords.EOF) Then
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
      
  'First check that the record hasn't already been deleted or amended !
  'MH20031002 Fault 7082 Reference Property instead of object to trap errors
  'iChangeReason = RecordAmended2(mobjTableView.RealSource, mobjTableView.TableID, mlngRecordID, mlngTimeStamp)
  iChangeReason = datGeneral.RecordAmended(Me.TableID, mobjTableView.RealSource, mlngRecordID, mlngTimeStamp)
  
  If iChangeReason = 2 Then
    ' The current record has been amended AND is no longer in the given table/view.
    AmendedRecord2 True, iChangeReason
    Screen.MousePointer = vbDefault
    Exit Sub
  
  ElseIf iChangeReason = 3 Then
    ' The current record has already been deleted.
    mrsRecords.MoveNext
    If mrsRecords.EOF Then
      ' We thought there were records after the current one but there aren't.
      ' Another user must have deleted them, so refresh the recordset.
      If Not RefreshRecordset Then
        Exit Sub
      End If
      
      ' Move to the last record.
      If mrsRecords.EditMode <> adEditAdd Then
        mrsRecords.MoveLast
      End If
    End If
  Else
    ' The record still exists, and is still in the current realSource.
    ' Prompt the user to confirm the deletion.
    sRecordDescription = EvaluateRecordDescription(mlngRecordID, mlngRecDescID)
    
    If pfPrompt Then
      'JPD 20030423 Fault 3286
      If miScreenType = screenLookup Then
        sTemp = "Are you sure you want to delete this entry from the lookup table ?"
      Else
        sTemp = "Are you sure you want to delete this record ?"
      End If
      
      If COAMsgBox(sTemp, vbQuestion + vbYesNo + vbDefaultButton2, "Delete Record " & IIf(Len(sRecordDescription) = 0, "", " - " & sRecordDescription)) <> vbYes Then
        fDeleteOK = False
      End If
    End If
    
    If fDeleteOK Then
      ' Delete the record from the recordset.
      ' Check that the required stored procedure exists.
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      'fDeleteOK = datGeneral.DeleteTableRecord(mobjTableView.TableID, mobjTableView.RealSource, mlngRecordID)
      fDeleteOK = datGeneral.DeleteTableRecord(Me.TableID, mobjTableView.RealSource, mlngRecordID)
      
      If fDeleteOK Then
        mrsRecords.MoveNext
        fGoLast = mrsRecords.EOF
        If Not fGoLast Then
          lngRecordID = mrsRecords!ID
        End If

        If Not RefreshRecordset Then
          Exit Sub
        End If

        ' Locate the record if it is still in the recorset.
        If mrsRecords.EditMode <> adEditAdd Then
          If fGoLast Then
            mrsRecords.MoveLast
          Else
            LocateRecord lngRecordID
          End If
        End If
      End If
    End If
  End If
        
  If fDeleteOK Then
    mfDataChanged = False
    ' JPD20021007 Fault 4498
    ReDim malngChangedOLEPhotos(0)
    
    'JPD 20030905 Fault 5184
    If (Not pfFromFind) Then
      UpdateAll
      frmMain.RefreshMainForm Me
    End If
    objEmail.SendImmediateEmails  'MH20100325 HRPRO-828
  End If
  
ExitDelete:
  Screen.MousePointer = vbDefault
  Exit Sub

ErrorTrap:
  fDeleteOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
  Resume ExitDelete

End Sub



Public Property Get Filtered() As Boolean
  ' Return TRUE if a filter is applied.
  Filtered = (UBound(mavFilterCriteria, 2) > 0)

End Property

Public Sub LocateRecord(plngID As Long)
  ' Locate the given record in the recordset.
  Dim fFound As Boolean
  Dim fRecordFound As Boolean
  Dim iColumnDataType As Integer
  Dim sColumnName As String
  Dim vOrderValue As Variant
  Dim iComparisonResult As Integer
  Dim lngUpper As Long
  Dim lngLower As Long
  Dim lngJump As Long
  Dim varFoundBookmark As Variant
  Dim objColumn As CColumnPrivilege
  Dim objColumns As CColumnPrivileges

  'MH20060615 Fault 11052 (Added Error Handling)
  On Local Error GoTo LocalErr

  With mrsRecords
    ' Check if we can determine the order column.
    If mlngFirstOrderColumnID = 0 Then
      ' Move to the first record.
      If (Not .BOF) Then .MoveFirst
  
      If .BOF Then
        If Not RefreshRecordset Then
          Exit Sub
        End If
      
        ' There are records in the refreshed recordset. Move to the first record.
        If .EditMode <> adEditAdd Then
          .MoveFirst
        End If
      Else
        'JPD 20030916 Fault 6979
        Do While Not !ID = plngID
          If (Not .EOF) Then .MoveNext
        
          If .EOF Then
            If (Not .BOF) Then .MoveFirst
            Exit Do
          End If
        Loop
      End If
    Else
      ' Check if the first order column is in the current table/view.
      fFound = False
      For Each objColumn In mcolColumnPrivileges
        If objColumn.ColumnID = mlngFirstOrderColumnID Then
          fFound = True
          sColumnName = objColumn.ColumnName
          iColumnDataType = objColumn.DataType
        End If
      Next objColumn
      Set objColumn = Nothing
    
      If (Not fFound) Or _
        ((iColumnDataType <> sqlVarChar) And _
        (iColumnDataType <> sqlVarBinary) And _
        (iColumnDataType <> sqlNumeric) And _
        (iColumnDataType <> sqlInteger)) Then
  
        .MoveFirst
        .Find "ID = " & plngID
        If .EOF Then
          If (Not .BOF) Then .MoveFirst
          If .BOF Then
            If Not RefreshRecordset Then
              Exit Sub
            End If
        
            ' There are records in the refreshed recordset. Move to the first record.
            If .EditMode <> adEditAdd Then
              .MoveFirst
            End If
          End If
        End If
      Else
        ' Binary search the recordset for the required record.
        vOrderValue = datGeneral.GetOrderValue(plngID, sColumnName, mobjTableView.RealSource)

        If IsEmpty(vOrderValue) Or IsNull(vOrderValue) Then
          .MoveFirst
          .Find "ID = " & plngID

          If .EOF Then
            ' Dodgy bit of recordset handling. I encountered errors when trying to moveFirst
            ' even though BOF was false. Doing the movePrevious and then the moveFirst sorted it out.
            If (Not .BOF) Then .MovePrevious
            If (Not .BOF) Then .MoveFirst

            If .BOF Then
              If Not RefreshRecordset Then
                Exit Sub
              End If

              ' There are records in the refreshed recordset. Move to the first record.
              If .EditMode <> adEditAdd Then
                .MoveFirst
              End If
            End If
          End If
        Else
          fFound = False
          lngLower = 1
          lngUpper = RecordCount

          Do
            Select Case iColumnDataType
              Case sqlVarChar, sqlVarBinary
                ' JPD String comparison changed from using VB's strComp function to
                ' using our own DictionaryCompareStrings function. VB's strComp
                ' function does not use the same order as that used when SQL orders
                ' by a character column. The DictionaryCompareStrings does.
                'iComparisonResult = StrComp(UCase(Left(IIf(IsNull(.Fields(sColumnName).Value), "", .Fields(sColumnName).Value), _
                  Len(vOrderValue))), UCase(vOrderValue), vbBinaryCompare)
                iComparisonResult = datGeneral.DictionaryCompareStrings(.Fields(sColumnName).Value, vOrderValue)

              Case sqlNumeric, sqlInteger
                If IsNull(.Fields(sColumnName).Value) Then
                  iComparisonResult = -1
                Else
                  If Val(.Fields(sColumnName).Value) = Val(vOrderValue) Then
                    iComparisonResult = 0
                  ElseIf Val(.Fields(sColumnName).Value) < Val(vOrderValue) Then
                    iComparisonResult = -1
                  Else
                    iComparisonResult = 1
                  End If
                End If
            End Select

            If Not mfFirstOrderColumnAscending Then
              iComparisonResult = iComparisonResult * -1
            End If

            Select Case iComparisonResult
              Case 0    ' String found.
                fFound = True
                varFoundBookmark = .Bookmark
                lngUpper = .Bookmark - 1
                lngJump = -((.Bookmark - lngLower) \ 2) - 1
                If lngLower > lngUpper Then Exit Do

              Case -1   ' Current record is before the required record.
                lngLower = .Bookmark + 1
                lngJump = ((lngUpper - .Bookmark) \ 2)
                If lngLower > lngUpper Then Exit Do

              Case 1    ' Current record is after the required record.
                lngUpper = .Bookmark - 1
                lngJump = -((.Bookmark - lngLower) \ 2) - 1
                If lngLower > lngUpper Then Exit Do
            End Select

            If lngLower = lngUpper Then
              lngJump = lngUpper - .Bookmark
            End If

            ' Move to the middle record of the remaining records to search.
            ' Only move forward if we're not on the EOF marker already.
            ' Only move back if we're not on the BOF marker already.
            If ((lngJump > 0) And (Not .EOF)) Or _
              ((lngJump < 0) And (Not .BOF)) Then
              .Move lngJump

              ' Check if we're now BOF or EOF.
              If .BOF Or .EOF Then
                Exit Do
              End If
            Else
              Exit Do
            End If
          Loop

          If fFound Then
            ' Find the record that has the same ID as the required one.
            .Bookmark = varFoundBookmark
            Do While Not !ID = plngID
              If (Not .EOF) Then .MoveNext

              If .EOF Then
                .Bookmark = varFoundBookmark
                Exit Do
              End If
            Loop
          Else
            ' Move to the first record.
            ' Dodgy bit of recordset handling. I encountered errors when trying to moveFirst
            ' even though BOF was false. Doing the movePrevious and then the moveFirst sorted it out.
            If (Not .BOF) Then .MovePrevious
            If (Not .BOF) Then .MoveFirst

            If .BOF Then
              If Not RefreshRecordset Then
                Exit Sub
              End If

              ' There are records in the refreshed recordset. Move to the first record.
              If .EditMode <> adEditAdd Then
                .MoveFirst
              End If
            End If
          End If
        End If
      End If
    End If
  End With
  
  mlngTimeStamp = mrsRecords!Timestamp
  mlngRecordID = mrsRecords!ID

Exit Sub

'MH20060615 Fault 11052 (Added Error Handling)
LocalErr:
  mlngRecordID = 0

End Sub


Private Function LostFocusCheck(pctlControl As Control) As Boolean
  On Error GoTo ErrorTrap
  
  Dim fValid As Boolean
  Dim fChanged As Boolean
  Dim iDatePart1 As Integer
  Dim iDatePart2 As Integer
  Dim iDatePart3 As Integer

  fValid = True

  If Len(pctlControl.Tag) > 0 Then
    'Check if the Afd module is enabled, and then if the control is Afd
    If gfAFDEnabled And _
      mobjScreenControls.Item(pctlControl.Tag).AFDEnabled Then
    
      'Check if the value in the control has actually changed
      If TypeOf pctlControl Is TextBox Then
        fChanged = (UCase(mvOldValue) <> UCase(pctlControl.Text))
      Else
        fChanged = False
      End If
      
      If fChanged Then
        ' Initialise the Afd form.
        modAfdShowMappedFields mobjScreenControls.Item(pctlControl.Tag).TableID, _
          mobjScreenControls.Item(pctlControl.Tag).ColumnName, pctlControl.Text, Me
      End If
    End If
    
    ' Check if Quick address module is enabled and the control is Quick Addressed
    If giQAddressEnabled <> QADDRESS_DISABLED And mobjScreenControls.Item(pctlControl.Tag).QAddressEnabled Then
      
      'Check if the value in the control has actually changed
      If TypeOf pctlControl Is TextBox Then
        fChanged = (UCase(mvOldValue) <> UCase(pctlControl.Text))
      Else
        fChanged = False
      End If
      
      If fChanged Then
        ' Initialise the Afd form.
        modQAShowMappedFields mobjScreenControls.Item(pctlControl.Tag).TableID, _
          mobjScreenControls.Item(pctlControl.Tag).ColumnName, pctlControl.Text, Me
      End If

    End If
    
    
  If TypeOf pctlControl Is GTMaskDate.GTMaskDate Then
    With pctlControl
      If Len(Trim(Replace(.Text, UI.GetSystemDateSeparator, ""))) <> 0 Then
        'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or .DateValue > #12/31/9999# Then
        If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Then
  
          .ForeColor = vbRed
          COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, app.title
          .ForeColor = vbWindowText
          .DateValue = Null
          If .Visible And .Enabled Then
            .SetFocus
          End If
          fValid = False
        End If
      End If
    End With
  End If
    
  ' Auto trimming (if any)
  If TypeOf pctlControl Is TextBox Then
    With pctlControl
      Select Case mobjScreenControls.Item(pctlControl.Tag).TrimmingType
        Case giTRIMMING_NONE
          ' Do nothing
        Case giTRIMMING_LEFTRIGHT
          .Text = Trim(.Text)
        Case giTRIMMING_LEFTONLY
          .Text = LTrim(.Text)
        Case giTRIMMING_RIGHTONLY
          .Text = RTrim(.Text)
      End Select
    End With
  End If
    
    ' Poxy new date control lets you lose focus for some invalid dates.
'    If TypeOf pctlControl Is GTMaskDate.GTMaskDate Then
'      If Trim(pctlControl.Text) = "/  /" Then
'        LostFocusCheck = True
'        Exit Function
'      End If

'      iDatePart1 = Val(pctlControl.Text) Mod 100
'      iDatePart2 = Val(Mid(pctlControl.Text, InStr(1, pctlControl.Text, "/") + 1)) Mod 100
'      iDatePart3 = Val(Mid(pctlControl.Text, InStr(InStr(1, pctlControl.Text, "/") + 1, pctlControl.Text, "/") + 1)) Mod 100

'      pctlControl.DateValue = pctlControl.DateValue
      
      
      'If (iDatePart1 <> Val(pctlControl.Text) Mod 100) Or _
      '  (iDatePart2 <> Val(Mid(pctlControl.Text, InStr(1, pctlControl.Text, "/") + 1)) Mod 100) Or _
      '  (iDatePart3 <> Val(Mid(pctlControl.Text, InStr(InStr(1, pctlControl.Text, "/") + 1, pctlControl.Text, "/") + 1)) Mod 100) Then
      '  pctlControl.DateValue = Null
      '  COAMsgBox "You have entered an invalid date.", vbExclamation + vbOKOnly, App.Title
      '  LostFocusCheck = False
      '  pctlControl.SetFocus
      '  Exit Function
      'ElseIf Not IsDate(pctlControl.Text) Then
      'If Not IsDate(pctlControl.Text) Then
      '  pctlControl.DateValue = Null
      '  COAMsgBox "You have entered an invalid date.", vbExclamation + vbOKOnly, App.Title
      '  LostFocusCheck = False
      '  pctlControl.SetFocus
      '  Exit Function
      'ElseIf CDate(pctlControl.Text) < CDate("01/01/1800") Then
      '  pctlControl.DateValue = Null
      '  COAMsgBox "You have entered an invalid date." & vbNewLine & "Date must be after 01/01/1800.", vbExclamation + vbOKOnly, App.Title
      '  LostFocusCheck = False
      '  pctlControl.SetFocus
      '  Exit Function
      'End If
    End If
  
    
TidyUpAndExit:
  ' Disassociate object variables.
  LostFocusCheck = fValid
  Exit Function
  
ErrorTrap:
  COAMsgBox Err.Description & " - LostFocusCheck", vbCritical
  Resume TidyUpAndExit
  
End Function

Private Sub TDBMask1_LostFocus(Index As Integer)
  Dim sTag As String
  
  sTag = TDBMask1(Index).Tag
  If Len(sTag) > 0 Then
    TDBMask1(Index).Text = CaseConversion(TDBMask1(Index).Text, mobjScreenControls.Item(sTag).ConvertCase)
  End If
  
End Sub

Private Sub TDBMask1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(TDBMask1(Index))

End Sub


Private Sub TDBNumber2_Change(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If (Not mfLoading) Then
    ' Get the control's tag.
    sTag = TDBNumber2(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In TDBNumber2
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .Text = TDBNumber2(Index).Text
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (TDBNumber2(Index).Value <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))

      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True

    frmMain.RefreshMainForm Me
  End If

End Sub

Private Sub TDBNumber2_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck TDBNumber2(Index)

End Sub


Private Sub TDBNumber2_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(TDBNumber2(Index))

End Sub

Private Sub TDBText1_Change(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
DebugOutput "TDBText1_Change", "Start"
    
  'JPD 20030815 Fault 6737
  If mlngUpdatedMultiLineControl = 0 Then
    mlngUpdatedMultiLineControl = Index
  End If
  
  If (Not mfLoading) Then
    ' Get the control's tag.
    sTag = TDBText1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In TDBText1
        With objControl
          If (objControl.Index <> Index) And _
            (objControl.Index <> mlngUpdatedMultiLineControl) Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .Text = TDBText1(Index).Text
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (TDBText1(Index).Text <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))

      frmMain.RefreshMainForm Me
    End If
  End If

  If mlngUpdatedMultiLineControl = Index Then
    mlngUpdatedMultiLineControl = 0
  End If
  
DebugOutput "TDBText1_Change", "End"
  
  Exit Sub
  
Err_Trap:
DebugOutput "TDBText1_Change", "Error"
  If Err.Number = 440 Then
    mfDataChanged = True

    frmMain.RefreshMainForm Me
  End If

  If mlngUpdatedMultiLineControl = Index Then
    mlngUpdatedMultiLineControl = 0
  End If
End Sub

Private Sub TDBText1_GotFocus(Index As Integer)
  
DebugOutput "TDBText1_GotFocus", "Start"
  
  ' Run the column's 'GotFocus' Expression.
  GotFocusCheck TDBText1(Index)

DebugOutput "TDBText1_GotFocus", "End"

End Sub


Private Sub TDBText1_Validate(Index As Integer, Cancel As Boolean)
  
DebugOutput "TDBText1_Validate", "Start"
  
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(TDBText1(Index))

DebugOutput "TDBText1_Validate", "End"

End Sub



Private Sub UpdateSiblingWindows(plngRecordID As Long, pfDeleted As Boolean)
  ' Update any other record editing screens for the same table as this one,
  ' that have the given record displayed.
  Dim frmTemp As Form
  
  For Each frmTemp In Forms
    If TypeOf frmTemp Is frmRecEdit4 Then
      If frmTemp.Recordset.State <> adStateClosed Then
        If Not (frmTemp.Recordset.BOF And frmTemp.Recordset.EOF) Then
          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'If (frmTemp.TableID = mobjTableView.TableID) And _
            (frmTemp.Recordset!ID = plngRecordID) And _
            (Not frmTemp Is Me) Then
          If (frmTemp.TableID = Me.TableID) And _
            (frmTemp.Recordset!ID = plngRecordID) And _
            (Not frmTemp Is Me) Then
          
            frmTemp.Requery pfDeleted
          End If
        End If
      End If
    End If
  Next frmTemp
  Set frmTemp = Nothing
  
End Sub

Private Sub Text1_Change(Index As Integer)
  On Error GoTo Err_Trap
  
  Dim lngThisControlsColumnID As Long
  Dim lngOtherControlsColumnID As Long
  Dim sTag As String
  Dim sThisControlsColumnName As String
  Dim objControl As Control
    
  If (Not mfLoading) Then
    ' Get the control's tag.
    sTag = Text1(Index).Tag
    
    ' Get the control's associated column ID).
    If Len(sTag) > 0 Then
      lngThisControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
      sThisControlsColumnName = mobjScreenControls.Item(sTag).ColumnName
    
      ' Update all other screen control's that represent the same column.
      For Each objControl In Text1
        With objControl
          If objControl.Index <> Index Then
            sTag = .Tag
            
            If Len(sTag) > 0 Then
              lngOtherControlsColumnID = mobjScreenControls.Item(sTag).ColumnID
            
              If lngOtherControlsColumnID = lngThisControlsColumnID Then
                .Text = Text1(Index).Text
              End If
            End If
          End If
        End With
      Next objControl
      Set objControl = Nothing
    End If
    
    ' Set the 'changed' flag if required.
    If Not mfDataChanged Then
      mfDataChanged = (Text1(Index).Text <> IIf(IsNull(mrsRecords(sThisControlsColumnName)), "", mrsRecords(sThisControlsColumnName) & vbNullString))

      frmMain.RefreshMainForm Me
    End If
  End If

  Exit Sub
  
Err_Trap:
  If Err.Number = 440 Then
    mfDataChanged = True

    frmMain.RefreshMainForm Me
  End If

End Sub


Private Sub Text1_GotFocus(Index As Integer)
  ' Run the column's 'GotFocus' Expression.
  Text1(Index).SelStart = 0
  Text1(Index).SelLength = Len(Text1(Index).Text)
  
  GotFocusCheck Text1(Index)

End Sub


Private Sub Text1_LostFocus(Index As Integer)
  ' Perform case conversion.
  Dim sTag As String
  Dim fDataChanged As Boolean
  
  sTag = Text1(Index).Tag
  If Len(sTag) > 0 Then
    fDataChanged = mfDataChanged
    mfDataChanged = True
    Text1(Index).Text = CaseConversion(Text1(Index).Text, mobjScreenControls.Item(sTag).ConvertCase)
    mfDataChanged = fDataChanged
  End If
  
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
  ' Run the column's 'LostFocus' Expression.
  Cancel = Not LostFocusCheck(Text1(Index))

End Sub



Private Function ValidateTrainingBookingRecord(plngCurrentRecordID As Long, pasColumns As Variant) As Boolean
  ' TRAINING BOOKING MODULE SPECIFICS.
  ' Check that if we are saving a Training Booking record that we are not over-booking a course
  ' or over-lapping another booking.
  Dim fValid As Boolean
  Dim iNextIndex As Integer
  Dim lngCourseRecordID As Long
  Dim lngEmployeeRecordID As Long
  Dim lngBookingID As Long
  Dim sBookingStatus As String
  
  fValid = True
    
  'MH20031002 Fault 7082 Reference Property instead of object to trap errors
  'If gfTrainingBookingEnabled And _
    (mobjTableView.TableID = glngTrainBookTableID) Then
  If gfTrainingBookingEnabled And _
    (Me.TableID = glngTrainBookTableID) Then
    
    ' We are saving a training booking record so check that we are not overbooking the course.
    ' Get the associated employee and course record IDs.
    lngCourseRecordID = 0
    lngEmployeeRecordID = 0
    lngBookingID = mlngRecordID
    sBookingStatus = ""
    
    For iNextIndex = 1 To UBound(pasColumns, 2)
      If UCase(Trim(pasColumns(1, iNextIndex))) = "ID_" & Trim(Str(glngEmployeeTableID)) Then
        lngEmployeeRecordID = Val(pasColumns(2, iNextIndex))
      End If
    
      If UCase(Trim(pasColumns(1, iNextIndex))) = "ID_" & Trim(Str(glngCourseTableID)) Then
        lngCourseRecordID = Val(pasColumns(2, iNextIndex))
      End If
    
      If UCase(Trim(pasColumns(1, iNextIndex))) = UCase(Trim(gsTrainBookStatusColumnName)) Then
        sBookingStatus = UCase(Mid(pasColumns(2, iNextIndex), 2, 1))
      End If
    
      If (lngEmployeeRecordID > 0) And _
        (lngCourseRecordID > 0) And _
        (Len(sBookingStatus) > 0) Then
        Exit For
      End If
    Next iNextIndex

    If (lngCourseRecordID > 0) And _
      ((sBookingStatus = "B") Or (sBookingStatus = "P")) Then

      'JPD20010815 Fault 2239 Only validate the booking if the status, course or employee has changed.
      'JPD20010815 Fault 2239 Only validate the booking is now deemed booked (or provisoinal)
      ' and wasn't so before.
      If (gfCourseIncludeProvisionals Or (sBookingStatus = "B")) And _
        ((mlngTBOriginalCourseID <> lngCourseRecordID) Or _
          (gfCourseIncludeProvisionals And _
            (msTBOriginalStatus <> "P") And _
            (msTBOriginalStatus <> "B")) Or _
          ((Not gfCourseIncludeProvisionals) And _
            (sBookingStatus = "B") And _
            (msTBOriginalStatus <> "B"))) Then

        fValid = TrainingBooking_CheckOverbooking(lngCourseRecordID, lngBookingID)
      End If
      
      If fValid And (lngEmployeeRecordID > 0) Then
        If (mlngTBOriginalEmpID <> lngEmployeeRecordID) Or _
          (mlngTBOriginalCourseID <> lngCourseRecordID) Or _
          ((msTBOriginalStatus <> "P") And (msTBOriginalStatus <> "B")) Then
      
          ' Check that the employee has satisfied the pre-requisite criteria for the selected course.
          fValid = TrainingBooking_CheckPreRequisites(lngCourseRecordID, lngEmployeeRecordID)
          
          If fValid Then
            ' Check that the employee is available for the selected course.
            fValid = TrainingBooking_CheckAvailability(lngCourseRecordID, lngEmployeeRecordID)
          End If
        
          If fValid Then
            ' Check that the employee is available for the selected course.
            fValid = TrainingBooking_CheckOverlappedBooking(lngCourseRecordID, lngEmployeeRecordID, lngBookingID)
          End If
        Else
          'All the above things could be equal if we are copy a record and then try to add him to the same course
          'So do the ovelapping check to sort this one out NHRD26012007 Fault 10658
          fValid = TrainingBooking_CheckOverlappedBooking(lngCourseRecordID, lngEmployeeRecordID, lngBookingID)
        End If
      End If
    End If
  End If
  
  ValidateTrainingBookingRecord = fValid
  
End Function



Public Property Get Changed() As Boolean
  ' Return the 'changed' flag.
  Changed = mfDataChanged
  
End Property


Private Sub UpdateParentWindow(Optional pvCallingFormID As Variant)
  ' Tell the parent window to refresh itself.
  Dim fGoodParentForm As Boolean
  Dim frmForm As Form
  
  ' Decide if we are a history table
  If (miScreenType = screenHistoryTable) Or _
    (miScreenType = screenHistoryView) Then

    ' We are a history table.
    ' Get the parent table and view ID
    For Each frmForm In Forms
      With frmForm
        If .Name = "frmRecEdit4" Then
          If (.FormID = mlngParentFormID) Then
            fGoodParentForm = IsMissing(pvCallingFormID)
            If Not fGoodParentForm Then
              fGoodParentForm = (.FormID <> CLng(pvCallingFormID))
            End If
          
            If fGoodParentForm Then
              .Requery False, Me.FormID
              
              ' JPD20021104 Fault 4692
              EnableActiveBar .ActiveBar1, False

              Exit For
            End If
          End If
        End If
      End With
    Next frmForm
    Set frmForm = Nothing
  End If

End Sub

Private Sub Update_AutoUpdateScreens(Optional pvCallingFormID As Variant)
  
  Dim frmForm As Form
  
  For Each frmForm In Forms
    With frmForm
      If (.Name = "frmRecEdit4") Then
        If (.FormID <> pvCallingFormID) And (.AutoUpdateLookup = True) And (.RecordID <> 0) Then
          .Requery False, Me.FormID
          EnableActiveBar .ActiveBar1, False
        End If
      End If
    End With
  Next frmForm
  Set frmForm = Nothing

End Sub

Public Sub MailMergeClick()

  Dim objExecution As clsMailMergeRun
  Dim frmSelection As frmDefSel
  Dim lForms As Long
  Dim blnExit As Boolean
  Dim blnOK As Boolean

  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to do a mail merge for", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If
  
  Set frmSelection = New frmDefSel
  blnExit = False
  
  With frmSelection
    Do While Not blnExit
      
      .TableComboEnabled = False
      .TableComboVisible = True
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '.TableID = mobjTableView.TableID
      .TableID = Me.TableID

      .Options = edtSelect
      .EnableRun = True

      If .ShowList(utlMailMerge) Then
        
        .CustomShow vbModal
        
        Select Case .Action
        Case edtSelect
          Set objExecution = New clsMailMergeRun
          objExecution.ExecuteMailMerge .SelectedID, CStr(mlngRecordID)
          Set objExecution = Nothing
          blnExit = gbCloseDefSelAfterRun

        Case edtCancel
          blnExit = True  'cancel

        End Select
      
      End If
    Loop
  End With

  frmMain.RefreshRecordEditScreens

  frmMain.RefreshMainForm Me


End Sub

Public Sub AbsenceBreakdownClick()

Dim frmDef As frmConfigurationReports
Dim fOK As Boolean

  'JDM - 13/08/01 - Fault 2677 - Force save before running reports
  If Not SaveChanges Then
    Exit Sub
  End If

  If mlngRecordID = 0 Then
    COAMsgBox "The Absence Breakdown report can only be produced for existing records." & vbNewLine & "If you are currently adding a new record, please save the record first.", vbInformation + vbOKOnly, app.title
    Exit Sub
  Else

    'JDM - 24/07/01 - Fault 2478 - Not checking correct value
    fOK = ValidateAbsenceParameters_BreakdownReport
    If fOK Then
      'frmAbsenceBreakdown.SetRecordID mlngRecordID, Me.Caption
      'frmAbsenceBreakdown.Show vbModal
      'If frmCrossTabRun.AbsenceBreakdownExecuteReport(False, mlngRecordID) Then
      '  If frmCrossTabRun.PreviewOnScreen Then
      '    frmCrossTabRun.Show vbModal
      '  End If
      'End If
      Set frmDef = New frmConfigurationReports
      frmDef.SingleRecord = mlngRecordID
      frmDef.Run = True
      frmDef.ShowControls "Absence Breakdown"
      frmDef.Show vbModal
  
      'JPD 20060124 Fault 10674
      Unload frmDef
      Set frmDef = Nothing

      Unload frmCrossTabRun
      Set frmCrossTabRun = Nothing
    End If
  End If
  
End Sub


Public Sub AbsenceCalendarClick()

  'JDM - 13/08/01 - Fault 2677 - Force save before running reports
  If Not SaveChanges Then
    Exit Sub
  End If
  
  If mlngRecordID = 0 Then
    COAMsgBox "The Absence Calendar can only be produced for existing records." & vbNewLine & "If you are currently adding a new record, please save the record first.", vbInformation + vbOKOnly, app.title
    Exit Sub
  Else
    ' JDM - 13/08/01 - Fault 2629 - Validate parameters before running calendar
'    If ValidateAbsenceParameters Then
      frmAbsenceCalendar.Initialise
      Unload frmAbsenceCalendar
      Set frmAbsenceCalendar = Nothing
'    End If
  End If

End Sub

Public Sub BradfordFactorClick()

  Dim frmDef As frmConfigurationReports
  'Dim pobjBradfordIndex As clsCustomReportsRUN
  Dim fOK As Boolean
  
  
  'JDM - 13/08/01 - Fault 2677 - Force save before running reports
  If Not SaveChanges Then
    Exit Sub
  End If

  If mlngRecordID = 0 Then
    COAMsgBox "The Bradford Factor can only be produced for existing records." & vbNewLine & "If you are currently adding a new record, please save the record first.", vbInformation + vbOKOnly, app.title
    Exit Sub
  Else
    'JDM - 24/07/01 - Fault 2478 - Not checking correct value
    fOK = ValidateAbsenceParameters_BreakdownReport
    If fOK Then
      'frmBradfordIndex.SetRecordID mlngRecordID
      'frmBradfordIndex.Show vbModal
      
      'Set pobjBradfordIndex = New clsCustomReportsRUN
      'pobjBradfordIndex.RunBradfordReport False, mlngRecordID
      'Set pobjBradfordIndex = Nothing

      Set frmDef = New frmConfigurationReports
      frmDef.SingleRecord = mlngRecordID
      frmDef.Run = True
      frmDef.ShowControls "Bradford Factor"
      frmDef.Show vbModal
      
      'JPD 20060124 Fault 10674
      Unload frmDef
      Set frmDef = Nothing

    End If
  End If

End Sub

Public Sub DataTransferClick()

  Dim objExecution As clsDataTransferRun
  Dim frmSelection As frmDefSel
  Dim lForms As Long
  Dim blnExit As Boolean
  Dim blnOK As Boolean

  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to do a data transfer for", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If

  Set frmSelection = New frmDefSel
  blnExit = False
  
  With frmSelection
    Do While Not blnExit
      
      .TableComboEnabled = False
      .TableComboVisible = True
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '.TableID = mobjTableView.TableID
      .TableID = Me.TableID

      .Options = edtSelect
      .EnableRun = True

      If .ShowList(utlDataTransfer) Then

        .CustomShow vbModal
        
        Select Case .Action
          Case edtSelect
            Set objExecution = New clsDataTransferRun
            objExecution.ExecuteDataTransfer .SelectedID, CStr(mlngRecordID)
            Set objExecution = Nothing
            blnExit = gbCloseDefSelAfterRun
  
          Case edtCancel
            blnExit = True  'cancel
        End Select
      
      End If
    
    Loop
  End With

  frmMain.RefreshRecordEditScreens

  frmMain.RefreshMainForm Me

End Sub


Public Property Get RecordID() As Long
  ' Return the current record ID.
  RecordID = mlngRecordID
  
End Property

Public Function RefreshRecordset() As Boolean
  ' Refresh the recordset. Return TRUE if the recordset was refreshed OK.
  ' NB. this method does not refresh any controls or screens, and does not position
  ' the cursor in the recordset. This is a general purpose method used by other
  ' methods that will do the positioning and refreshment.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sErrorMsg As String
  
  fOK = True
  
  ' If it is empty, add a new record if permitted.
  With mrsRecords
    ' JPD 21/02/2001 - Only try to cancel the update if we're not BOF or EOF.
    ' ADO2.6 error occurs if we do.
    If Not (.BOF Or .EOF) Then
      .CancelUpdate
    End If
    
    .Requery

    mlngRecordCount = GetRecordCount
  
    ' Check if the refreshed recordset is empty and filtered.
    If (.BOF And .EOF) And Filtered Then
      ' Clear the filter.
      COAMsgBox "No records match the current filter." & vbNewLine & _
        "The filter has been cleared.", vbInformation + vbOKOnly, app.ProductName
      ReDim mavFilterCriteria(3, 0)
      .Close
      Set mrsRecords = Nothing
      GetRecords
    End If
  
  
  
  End With
  'MH20001019 Fault 1086
  ' There is one WITH loop above and one WITH loop below this point which appear
  ' to be referencing the same recordset.  However, a problem occured when "NO
  ' RECORDS MATCH THE CURRENT FILTER" (above).  The recordset was closed and
  ' refreshed.  As soon as it was checked for BOF and EOF below got the message
  ' that it was closed (I think that the with loop was still referencing the old
  ' closed recordset ????  Anyway, I changed the single with to be two with loops
  ' and that seemed to fix the problem (All I did was close the with loop and
  ' open a new one !
  With mrsRecords
    
    
    
    ' Check if the refreshed recordset is still empty.
    If (.BOF And .EOF) Then
      ' The refreshed recordset is empty. Create a new record if permitted.
      If mobjTableView.AllowInsert Then
        .AddNew
      Else
        ' The refreshed recordset is empty and a new record cannot be created.
        ' Kill the record edit form.
        fOK = False
        'MH20031002 Fault 7082 Reference Property instead of object to trap errors
        'sErrorMsg = "This " & IIf(mobjTableView.ViewID > 0, "view", "table") & " is empty" & _
          " and you do not have 'new' permission on it."
        sErrorMsg = "This " & IIf(Me.ViewID > 0, "view", "table") & " is empty" & _
          " and you do not have 'new' permission on it."
       End If
    End If
  End With
  
TidyUpAndExit:
  If Not fOK Then
    COAMsgBox sErrorMsg, vbExclamation, "Security"
    ' Reset the mousepointer as it may have been changed in the calling method.
    Screen.MousePointer = vbDefault
    Unload Me
  End If
  
  RefreshRecordset = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  sErrorMsg = Err.Description
  Resume TidyUpAndExit
  
End Function

Public Sub EmailClick()

  Dim frmSelection As frmEmailSel
  Dim strRecordDetails As String
  
  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to do a mail merge for", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If
  
  Set frmSelection = New frmEmailSel
      
  strRecordDetails = mobjTableView.TableName & " : " & _
                     EvaluateRecordDescription(mlngRecordID, mlngRecDescID)

  'MH20031002 Fault 7082 Reference Property instead of object to trap errors
  'frmSelection.Initialise mobjTableView.TableID, mlngRecordID, strRecordDetails
  frmSelection.Initialise Me.TableID, mlngRecordID, strRecordDetails
  frmSelection.Show vbModal

End Sub


Private Sub ControlSetFocus(ctl As Control)

  'MH20001120 Stops any SetFocus errors !!!
  'Now, I tried checking visible and enabled properties
  'then got problems with labels.  Just did this instead
  'cos its the easiest (and probably most reliable) way!
  On Local Error Resume Next
  ctl.SetFocus

End Sub

Public Sub LabelsClick()

  Dim objExecution As clsMailMergeRun
  Dim frmSelection As frmDefSel
  Dim lForms As Long
  Dim blnExit As Boolean
  Dim blnOK As Boolean

  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to do a labels & envelopes for", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If
  
  Set frmSelection = New frmDefSel
  blnExit = False
  
  With frmSelection
    Do While Not blnExit
      
      .TableComboEnabled = False
      .TableComboVisible = True
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '.TableID = mobjTableView.TableID
      .TableID = Me.TableID

      .Options = edtSelect
      .EnableRun = True

      If .ShowList(utlLabel) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtSelect
          Set objExecution = New clsMailMergeRun
          objExecution.ExecuteMailMerge .SelectedID, CStr(mlngRecordID)
          Set objExecution = Nothing
          blnExit = gbCloseDefSelAfterRun

        Case edtCancel
          blnExit = True  'cancel

        End Select
      
      End If
    Loop
  End With

  frmMain.RefreshRecordEditScreens

  frmMain.RefreshMainForm Me

End Sub


Public Sub MatchReportClick(mrtMatchReportType As MatchReportType)
  
  Dim lForms As Long
  Dim fExit As Boolean
  Dim frmSelection As frmDefSel
  Dim frmEdit As frmMatchDef
  Dim frmRun As frmMatchRun
  Dim lngTYPE As UtilityType
  
  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to run a report for.", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If

  If mrtMatchReportType <> mrtNormal Then
    If Not ValidatePostParameters Then
      Exit Sub
    End If
  End If

  Screen.MousePointer = vbHourglass

  fExit = False
  Set frmSelection = New frmDefSel

  With frmSelection
    
    Do While Not fExit
      .Options = edtSelect
      .EnableRun = True
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '.TableID = mobjTableView.TableID
      .TableID = Me.TableID
      .TableComboVisible = True
      .TableComboEnabled = False

      Select Case mrtMatchReportType
      Case mrtNormal: lngTYPE = utlMatchReport
      Case mrtSucession: lngTYPE = utlSuccession
      Case mrtCareer: lngTYPE = utlCareer
      End Select

      If .ShowList(lngTYPE, "MatchReportType = " & CStr(mrtMatchReportType)) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtAdd
          Set frmEdit = New frmMatchDef
          frmEdit.MatchReportType = mrtMatchReportType
          frmEdit.Initialise True, .FromCopy
          frmEdit.Show vbModal
          .SelectedID = frmEdit.SelectedID
          Unload frmEdit
          Set frmEdit = Nothing

        Case edtEdit
          Set frmEdit = New frmMatchDef
          frmEdit.MatchReportType = mrtMatchReportType
          frmEdit.Initialise False, .FromCopy, .SelectedID
          If Not frmEdit.Cancelled Then
            frmEdit.Show vbModal
            If .FromCopy And frmEdit.SelectedID > 0 Then
              .SelectedID = frmEdit.SelectedID
            End If
          End If
          Unload frmEdit
          Set frmEdit = Nothing

        Case edtSelect
          Set frmRun = New frmMatchRun
          frmRun.MatchReportType = mrtMatchReportType
          frmRun.MatchReportID = .SelectedID
          'MH20031002 Fault 7082 Reference Property instead of object to trap errors
          'frmRun.RunMatchReport mobjTableView.TableID, mlngRecordID
          frmRun.RunMatchReport Me.TableID, mlngRecordID
          If frmRun.PreviewOnScreen Then
            frmRun.Show vbModal
          End If
          Unload frmRun
          Set frmRun = Nothing
          fExit = gbCloseDefSelAfterRun

        Case edtPrint
          Set frmEdit = New frmMatchDef
          frmEdit.MatchReportType = mrtMatchReportType
          frmEdit.Initialise False, False, .SelectedID, True
          If Not frmEdit.Cancelled Then
            frmEdit.PrintDef .SelectedID
          End If
          Unload frmEdit
          Set frmEdit = Nothing

        Case edtCancel
          fExit = True

          End Select
        End If
    Loop
  End With
  
  Unload frmSelection
  Set frmSelection = Nothing

  frmMain.RefreshRecordEditScreens
  frmMain.RefreshMainForm Me

End Sub

Public Sub RecordProfileClick()
  Dim objRecordProfile As clsRecordProfileRUN
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean

  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to run a record profile for.", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If

  Set frmSelection = New frmDefSel
  blnExit = False

  With frmSelection
    Do While Not blnExit

      .TableComboEnabled = False
      .TableComboVisible = True
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '.TableID = mobjTableView.TableID
      .TableID = Me.TableID

      .Options = edtSelect
      .EnableRun = True

      If .ShowList(utlRecordProfile) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtSelect
          Set objRecordProfile = New clsRecordProfileRUN
          objRecordProfile.RecordProfileID = .SelectedID
          objRecordProfile.RunRecordProfile mlngRecordID
          Set objRecordProfile = Nothing
          blnExit = gbCloseDefSelAfterRun
        
        Case edtCancel
          blnExit = True  'cancel

        End Select
      End If
    Loop
  End With

  frmMain.RefreshRecordEditScreens
  frmMain.RefreshMainForm Me

End Sub

Public Sub CalendarReportClick()
  
  Dim objCalendarReport As clsCalendarReportsRUN
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean

  If mlngRecordID = 0 Then
    COAMsgBox "Please select which record you would like to run a calendar report for.", vbInformation
    Exit Sub
  End If

  If Not SaveChanges Then
    Exit Sub
  End If

  Set frmSelection = New frmDefSel
  blnExit = False

  With frmSelection
    Do While Not blnExit

      .TableComboEnabled = False
      .TableComboVisible = True
      'MH20031002 Fault 7082 Reference Property instead of object to trap errors
      '.TableID = mobjTableView.TableID
      .TableID = Me.TableID

      .Options = edtSelect
      .EnableRun = True

      If .ShowList(utlCalendarReport) Then

        .CustomShow vbModal
        Select Case .Action
        Case edtSelect
          Set objCalendarReport = New clsCalendarReportsRUN
          objCalendarReport.CalendarReportID = .SelectedID
          objCalendarReport.RunCalendarReport CStr(mlngRecordID)
          Set objCalendarReport = Nothing
          blnExit = gbCloseDefSelAfterRun
        
        Case edtCancel
          blnExit = True  'cancel

        End Select
      End If
    Loop
  End With

  frmMain.RefreshRecordEditScreens
  frmMain.RefreshMainForm Me

End Sub


'Public Sub CareerSuccessionClick(lngMatchReportType As MatchReportType)
'
'  Dim frmRun As frmMatchRun
'
'  Set frmRun = New frmMatchRun
'
'  If lngMatchReportType = mrtSucession Then
'    frmRun.SetPostReport mrtSucession, glngSuccessionDef, gblnSuccessionAllowEqual, gblnSuccessionRestrict, gblnSuccessionLevels
'  Else
'    frmRun.SetPostReport mrtCareer, glngCareerDef, gblnCareerAllowEqual, gblnCareerRestrict, gblnCareerLevels
'  End If
'
'  frmRun.RunMatchReport mobjTableView.TableID, mlngRecordID
'  If frmRun.PreviewOnScreen Then
'    frmRun.Show vbModal
'  End If
'  Set frmRun = Nothing
'
'End Sub
'

' Reads the field from the database into the passed in COA_OLE object
Private Function ReadStream(ByRef pobjOLE As Object, _
  ByRef pobjColumn As DataMgr.clsScreenControl, ByVal pbHeaderOnly As Boolean) As Boolean

  On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim sSQL As String
  Dim rsDocument As ADODB.Recordset
  Dim objDocument As ADODB.Stream

  Set rsDocument = New ADODB.Recordset
  Set objDocument = New ADODB.Stream

  bOK = True

  ' New record - thus no stream will exist
  If mlngRecordID = 0 Then
    ReadStream = bOK
    Exit Function
  End If

  If pbHeaderOnly Then
    sSQL = "SUBSTRING(" & pobjColumn.ColumnName & ",1,400) AS " & pobjColumn.ColumnName
  Else
    sSQL = pobjColumn.ColumnName
  End If

  sSQL = "SELECT " & sSQL _
    & " FROM " & mobjTableView.RealSource _
    & " WHERE ID=" & mlngRecordID

  ' Set as blank stream
  pobjOLE.EmbeddedStream = New ADODB.Stream

  rsDocument.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  With rsDocument
    objDocument.Open
    objDocument.Type = adTypeBinary
    
    If Not IsNull(rsDocument.Fields(pobjColumn.ColumnName).Value) Then
      objDocument.Write rsDocument.Fields(pobjColumn.ColumnName).Value
    End If
    
    If objDocument.Size > 0 Then
      If pobjOLE.EmbeddedStream.State = adStateClosed Then
        pobjOLE.EmbeddedStream.Open
      End If
      
      objDocument.Position = 0
      pobjOLE.EmbeddedStream.Type = adTypeBinary
      objDocument.CopyTo pobjOLE.EmbeddedStream
    End If
    
    objDocument.Close
  End With

  rsDocument.Close
  
TidyUpAndExit:
  Set rsDocument = Nothing
  ReadStream = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit
  
End Function

Private Function SaveStream(ByRef pobjOLE As Object, _
  ByRef pobjColumn As DataMgr.clsScreenControl, plngRecordID As Long) As Boolean

  Dim bOK As Boolean
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo ErrorTrap
  bOK = True

  If mcolColumnPrivileges.Item(pobjColumn.ColumnName).AllowUpdate Then
  
    If mfRequiresLocalCursor Then gADOCon.CursorLocation = adUseClient
    
    Set cmADO = New ADODB.Command
    With cmADO
    
      .CommandText = "spASRUpdateOLEField_" & pobjColumn.ColumnID
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
            
      Set pmADO = .CreateParameter("currentID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngRecordID
                      
      Set pmADO = .CreateParameter("UploadFile", adLongVarBinary, adParamInput, -1)
      .Parameters.Append pmADO
    
      If pobjOLE.EmbeddedStream.State = adStateClosed Then
        pobjOLE.EmbeddedStream.Open
      End If
      
      If pobjOLE.EmbeddedStream.Size > 0 Then
        pobjOLE.EmbeddedStream.Position = 0
        pmADO.Value = pobjOLE.EmbeddedStream.Read
      Else
        pmADO.Value = Null
      End If
    
    End With
            
    cmADO.Execute

    If mfRequiresLocalCursor Then gADOCon.CursorLocation = adUseServer
    
  End If

TidyUpAndExit:
  Set pmADO = Nothing
  Set cmADO = Nothing
  SaveStream = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

' Clears the embedded streams for OLEs and Photos
Public Sub ClearEmbeddedStreams()

  Dim iCount As Integer
      
  ' OLEs
  For iCount = OLE1.LBound To OLE1.UBound
    'JPD 20041222 Fault 9262
    If Not OLE1(iCount).EmbeddedStream Is Nothing Then
      If OLE1(iCount).EmbeddedStream.State = adStateOpen Then
        OLE1(iCount).EmbeddedStream.Close
      End If
    End If
  Next iCount

  ' Photos
  For iCount = ASRUserImage1.LBound To ASRUserImage1.UBound
    If ASRUserImage1(iCount).OLEType > 0 Then
      'JPD 20041222 Fault 9262
      If Not ASRUserImage1(iCount).EmbeddedStream Is Nothing Then
        If ASRUserImage1(iCount).EmbeddedStream.State = adStateOpen Then
          ASRUserImage1(iCount).EmbeddedStream.Close
        End If
      End If
    End If
  Next iCount

End Sub

Public Sub AccordClick()

  Dim sSQL As String
  Dim rsTemp As New ADODB.Recordset
  Dim bFound As Boolean
  Dim lngTransferTypeID As Long

  If Not SaveChanges Then
    Exit Sub
  End If

  If mlngRecordID = 0 Then
    COAMsgBox "Payroll transactions can only be produced for existing records." & vbNewLine & "If you are currently adding a new record, please save the record first.", vbInformation + vbOKOnly, app.title
    Exit Sub
  Else
  
    sSQL = "SELECT [ASRBaseTableID], [TransferTypeID] FROM ASRSysAccordTransferTypes WHERE ASRBaseTableID = " & datGeneral.GetTableIDFromTableViewName(msTableViewName)
    rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    If Not (rsTemp.BOF And rsTemp.EOF) Then
      bFound = True
      lngTransferTypeID = rsTemp.Fields("TransferTypeID").Value
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bFound Then
      frmAccordViewTransfers.ConnectionType = ACCORD_LOCAL
      frmAccordViewTransfers.ViewMode = iCURRENT_RECORD
      frmAccordViewTransfers.CurrentRecordID = mlngRecordID
      frmAccordViewTransfers.TransferType = lngTransferTypeID
      frmAccordViewTransfers.Initialise
      frmAccordViewTransfers.Show vbModal
      Set frmAccordViewTransfers = Nothing
    Else
      COAMsgBox "Payroll transactions can only be viewed for defined transfer types.", vbInformation + vbOKOnly, app.title
      Exit Sub
    End If
  
  End If
  
End Sub

Private Function SetControlLevel() As Boolean
  ' Set the correct z-order for controls.
  ' Bring labels to the front
  On Error GoTo ErrorTrap
  
  Dim ctlControl As VB.Control
  Dim fOK As Boolean
  
  fOK = True
  
  For Each ctlControl In Me.Controls
    If TypeOf ctlControl Is COA_Label Then
      ctlControl.ZOrder 0
    End If
  Next ctlControl
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SetControlLevel = fOK
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error setting control level." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, app.ProductName
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Sub SendToAccord(ByVal pbSendAsNew As Boolean)

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim iMappedAccordTransfers() As Integer
  Dim iCount As Integer

  ' Void all previous transactions for this record
  iMappedAccordTransfers = MappedAccordTransfers(mobjTableView.TableID)
  
  For iCount = LBound(iMappedAccordTransfers) To UBound(iMappedAccordTransfers)
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "spASRAccordVoidPreviousTransactions"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
  
      Set pmADO = .CreateParameter("TransferType", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = iMappedAccordTransfers(iCount)
      
      Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = mlngRecordID
              
      .Execute
  
    End With
  Next iCount

  
  ' Resave this record
  mbResendingToAccord = True
  SaveChanges
  mbResendingToAccord = False


  ' Set correct new/update transaction type
  For iCount = LBound(iMappedAccordTransfers) To UBound(iMappedAccordTransfers)
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "spASRAccordSetLatestToType"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
  
      Set pmADO = .CreateParameter("TransferType", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = iMappedAccordTransfers(iCount)
      
      Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = mlngRecordID
              
      Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = IIf(pbSendAsNew, 0, 1)
              
      .Execute
  
    End With
  Next iCount


End Sub

'NPG20080516 Fault 12973
Public Sub EnableNavigation(blnEnabled As Boolean)

    ActiveBar1.Tools("FirstRecord").Enabled = blnEnabled
    ActiveBar1.Tools("PreviousRecord").Enabled = blnEnabled
    ActiveBar1.Tools("NextRecord").Enabled = blnEnabled
    ActiveBar1.Tools("LastRecord").Enabled = blnEnabled

End Sub

' Executes any code thats in the hidden navigation control
Private Sub ExecutePostSaveCode()

  On Error GoTo ErrorTrap
  
  Dim objControl As Control

  For Each objControl In Me.Controls
    If TypeOf objControl Is COA_Navigation Then
      objControl.ExecutePostSave
    End If
  Next objControl

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  GoTo TidyUpAndExit

End Sub

Private Function CopyWhenParentRecordIsCopied(ByVal iParentTableID As Integer, ByVal iNewRecordID As Integer, ByVal iOriginalRecordID As Integer) As Boolean

  On Error GoTo ErrorTrap
  
  Dim bOK As Boolean
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  bOK = True

  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.spASRCopyChildRecords"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("iParentTableID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = iParentTableID

    Set pmADO = .CreateParameter("iParentTableID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = iNewRecordID

    Set pmADO = .CreateParameter("iOriginalRecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = iOriginalRecordID

    cmADO.Execute

  End With
  Set cmADO = Nothing

TidyUpAndExit:
  CopyWhenParentRecordIsCopied = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function
