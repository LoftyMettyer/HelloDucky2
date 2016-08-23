VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "COA_ColourPicker.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12660
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHRProUtilityLink 
      Caption         =   "Report / Utility :"
      Height          =   1485
      Left            =   2880
      TabIndex        =   31
      Top             =   6180
      Width           =   6300
      Begin VB.ComboBox cboHRProUtility 
         Height          =   315
         Left            =   1400
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   700
         Width           =   4700
      End
      Begin VB.ComboBox cboHRProUtilityType 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":000C
         Left            =   1400
         List            =   "frmSSIntranetLink.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   300
         Width           =   4700
      End
      Begin VB.Label lblHRProUtilityMessage 
         AutoSize        =   -1  'True
         Caption         =   "<message>"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1395
         TabIndex        =   36
         Top             =   1160
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHRProUtility 
         Caption         =   "Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   765
         Width           =   780
      End
      Begin VB.Label lblHRProUtilityType 
         Caption         =   "Type :"
         Height          =   195
         Left            =   195
         TabIndex        =   32
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame fraURLLink 
      Caption         =   "URL :"
      Height          =   1125
      Left            =   2880
      TabIndex        =   37
      Top             =   4470
      Width           =   6300
      Begin VB.TextBox txtURL 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   39
         Top             =   300
         Width           =   4515
      End
      Begin VB.CheckBox chkNewWindow 
         Caption         =   "D&isplay in new window"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1575
         TabIndex        =   40
         Top             =   690
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.Label lblURL 
         Caption         =   "URL :"
         Height          =   195
         Left            =   195
         TabIndex        =   38
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame fraChartLink 
      Caption         =   "Chart :"
      Height          =   6060
      Left            =   2880
      TabIndex        =   66
      Top             =   5760
      Width           =   6300
      Begin VB.CheckBox chkDataGridUseFormatting 
         Caption         =   "Use &formatting in Data Grids"
         Height          =   240
         Left            =   210
         TabIndex        =   80
         Top             =   4395
         Width           =   3495
      End
      Begin VB.CheckBox chkDataGrid1000Separator 
         Caption         =   "Use &1000 separator"
         Height          =   240
         Left            =   4065
         TabIndex        =   82
         Top             =   4755
         Width           =   2010
      End
      Begin VB.CheckBox chkShowPercentages 
         Caption         =   "Show Values as Perce&nt"
         Height          =   195
         Left            =   210
         TabIndex        =   74
         Top             =   2730
         Width           =   2355
      End
      Begin VB.CheckBox chkPrimaryDisplay 
         Caption         =   "D&isplay Data Grid First"
         Height          =   195
         Left            =   210
         TabIndex        =   73
         Top             =   2385
         Width           =   2235
      End
      Begin VB.CommandButton cmdChartData 
         Caption         =   "..."
         Height          =   315
         Left            =   5715
         TabIndex        =   76
         ToolTipText     =   "Select Chart Data"
         Top             =   3315
         Width           =   315
      End
      Begin VB.TextBox txtChartDescription 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1935
         TabIndex        =   75
         Top             =   3315
         Width           =   3780
      End
      Begin VB.CommandButton cmdChartReportClear 
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
         Left            =   5715
         MaskColor       =   &H000000FF&
         TabIndex        =   79
         ToolTipText     =   "Clear Drill Down Utility"
         Top             =   3705
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdChartReport 
         Caption         =   "..."
         Height          =   315
         Left            =   5385
         TabIndex        =   78
         ToolTipText     =   "Select Drill Down Utility"
         Top             =   3705
         Width           =   315
      End
      Begin VB.TextBox txtChartUtility 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1935
         TabIndex        =   77
         Top             =   3705
         Width           =   3450
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2505
         Left            =   2730
         OleObjectBlob   =   "frmSSIntranetLink.frx":0010
         TabIndex        =   83
         Top             =   555
         Width           =   3330
      End
      Begin VB.CheckBox chkShowValues 
         Caption         =   "Show &Values"
         Height          =   210
         Left            =   210
         TabIndex        =   71
         Top             =   1695
         Width           =   1665
      End
      Begin VB.CheckBox chkStackSeries 
         Caption         =   "Stac&k Series"
         Height          =   210
         Left            =   210
         TabIndex        =   72
         Top             =   2040
         Width           =   1665
      End
      Begin VB.CheckBox chkDottedGridlines 
         Caption         =   "Dotted &Gridlines"
         Height          =   195
         Left            =   210
         TabIndex        =   70
         Top             =   1350
         Width           =   1980
      End
      Begin VB.CheckBox chkShowLegend 
         Caption         =   "Show &Legend"
         Height          =   240
         Left            =   210
         TabIndex        =   69
         Top             =   990
         Width           =   1710
      End
      Begin VB.ComboBox cboChartType 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":2500
         Left            =   210
         List            =   "frmSSIntranetLink.frx":2502
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   555
         Width           =   2205
      End
      Begin COASpinner.COA_Spinner spnDataGridDecimals 
         Height          =   300
         Left            =   2535
         TabIndex        =   81
         Top             =   4710
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label lblDataGridDecimals 
         AutoSize        =   -1  'True
         Caption         =   "Fix decimal places :"
         Height          =   195
         Left            =   210
         TabIndex        =   135
         Top             =   4755
         Width           =   1695
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000015&
         X1              =   210
         X2              =   6015
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label lblChartData 
         AutoSize        =   -1  'True
         Caption         =   "Chart Data :"
         Height          =   195
         Left            =   210
         TabIndex        =   128
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label lblUtility 
         AutoSize        =   -1  'True
         Caption         =   "Drill Down Utility :"
         Height          =   195
         Left            =   210
         TabIndex        =   127
         Top             =   3750
         Width           =   1560
      End
      Begin VB.Label lblChartyType 
         AutoSize        =   -1  'True
         Caption         =   "Chart Type :"
         Height          =   195
         Left            =   210
         TabIndex        =   67
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraLinkSeparator 
      Caption         =   "Separator :"
      Height          =   4170
      Left            =   2880
      TabIndex        =   57
      Top             =   5250
      Width           =   6300
      Begin VB.CheckBox chkSeparatorUseFormatting 
         Caption         =   "Use &formatting"
         Height          =   195
         Left            =   210
         TabIndex        =   63
         Top             =   1260
         Width           =   1800
      End
      Begin VB.TextBox txtSeparatorColour 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   64
         Top             =   1545
         Width           =   1170
      End
      Begin VB.CommandButton cmdSeparatorColPick 
         Caption         =   "..."
         Height          =   315
         Left            =   3765
         TabIndex        =   65
         ToolTipText     =   "Select Border Colour"
         Top             =   1530
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "..."
         Height          =   315
         Left            =   4830
         TabIndex        =   60
         ToolTipText     =   "Select Icon"
         Top             =   315
         Width           =   315
      End
      Begin VB.TextBox txtIcon 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1050
         TabIndex        =   59
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
         TabIndex        =   61
         ToolTipText     =   "Clear Icon"
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chkNewColumn 
         Caption         =   "Column &break"
         Height          =   255
         Left            =   1050
         TabIndex        =   62
         Top             =   690
         Width           =   2040
      End
      Begin VB.Label lblDiaryWarning 
         AutoSize        =   -1  'True
         Caption         =   "lblDiaryWarningText"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   300
         TabIndex        =   136
         Top             =   2655
         Visible         =   0   'False
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSeparatorColour 
         AutoSize        =   -1  'True
         Caption         =   "Separator border colour :"
         Height          =   195
         Left            =   210
         TabIndex        =   126
         Top             =   1575
         Width           =   2205
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000015&
         X1              =   210
         X2              =   6075
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblNoOptions 
         AutoSize        =   -1  'True
         Caption         =   "There are no configurable options for this link type."
         Height          =   195
         Left            =   285
         TabIndex        =   84
         Top             =   2190
         Visible         =   0   'False
         Width           =   4410
      End
      Begin VB.Label lblIcon 
         Caption         =   "Icon :"
         Height          =   195
         Left            =   210
         TabIndex        =   58
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
   Begin VB.Frame fraLink 
      Caption         =   "Link :"
      Height          =   1710
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   9000
      Begin VB.ComboBox cboTableView 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":2504
         Left            =   1485
         List            =   "frmSSIntranetLink.frx":2506
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
         stylesets(0).Picture=   "frmSSIntranetLink.frx":2508
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
         stylesets(1).Picture=   "frmSSIntranetLink.frx":2524
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
   Begin VB.Frame fraApplicationLink 
      Caption         =   "Application :"
      Height          =   2190
      Left            =   2880
      TabIndex        =   41
      Top             =   3585
      Width           =   6300
      Begin VB.CommandButton cmdAppFilePathSel 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         TabIndex        =   44
         ToolTipText     =   "Select File Path"
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtAppFilePath 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   43
         Top             =   300
         Width           =   4185
      End
      Begin VB.TextBox txtAppParameters 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   46
         Top             =   700
         Width           =   4515
      End
      Begin VB.Label lblApplicationLinksUnavailable 
         AutoSize        =   -1  'True
         Caption         =   "Application links are no longer supported. The parameters shown here are for information only."
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   195
         TabIndex        =   137
         Top             =   1155
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAppFilePath 
         AutoSize        =   -1  'True
         Caption         =   "File Path :"
         Height          =   195
         Left            =   195
         TabIndex        =   42
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblAppParameters 
         AutoSize        =   -1  'True
         Caption         =   "Parameters :"
         Height          =   195
         Left            =   195
         TabIndex        =   45
         Top             =   765
         Width           =   930
      End
   End
   Begin VB.Frame fraDBValue 
      Caption         =   "Database Value :"
      Height          =   6060
      Left            =   2880
      TabIndex        =   85
      Top             =   6000
      Width           =   6300
      Begin VB.ComboBox cboDBValCFOperator 
         Height          =   315
         Index           =   2
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   4635
         Width           =   2730
      End
      Begin VB.TextBox txtDBValCFValue 
         Height          =   315
         Index           =   2
         Left            =   2970
         TabIndex        =   117
         Top             =   4635
         Width           =   645
      End
      Begin VB.ComboBox cboDBValCFStyle 
         Height          =   315
         Index           =   2
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   4635
         Width           =   1470
      End
      Begin VB.TextBox txtDBValCFColour 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5160
         TabIndex        =   119
         Top             =   4635
         Width           =   570
      End
      Begin VB.CommandButton cmdDBValueColPick 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   5745
         TabIndex        =   120
         ToolTipText     =   "Select Formatting Colour"
         Top             =   4635
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.ComboBox cboDBValCFOperator 
         Height          =   315
         Index           =   1
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   4275
         Width           =   2730
      End
      Begin VB.TextBox txtDBValCFValue 
         Height          =   315
         Index           =   1
         Left            =   2970
         TabIndex        =   112
         Top             =   4275
         Width           =   645
      End
      Begin VB.ComboBox cboDBValCFStyle 
         Height          =   315
         Index           =   1
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   4275
         Width           =   1470
      End
      Begin VB.TextBox txtDBValCFColour 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5160
         TabIndex        =   114
         Top             =   4275
         Width           =   570
      End
      Begin VB.CommandButton cmdDBValueColPick 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   5745
         TabIndex        =   115
         ToolTipText     =   "Select Formatting Colour"
         Top             =   4275
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtDBValueSample 
         Height          =   315
         Left            =   1215
         TabIndex        =   121
         Text            =   "12345.210"
         Top             =   5280
         Width           =   1290
      End
      Begin COASpinner.COA_Spinner spnDBValueDecimals 
         Height          =   300
         Left            =   2310
         TabIndex        =   99
         Top             =   2385
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.CommandButton cmdDBValueColPick 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   5745
         TabIndex        =   110
         ToolTipText     =   "Select Formatting Colour"
         Top             =   3915
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtDBValCFColour 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   5160
         TabIndex        =   109
         Top             =   3915
         Width           =   570
      End
      Begin VB.ComboBox cboDBValCFStyle 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSSIntranetLink.frx":2540
         Left            =   3660
         List            =   "frmSSIntranetLink.frx":2542
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   3915
         Width           =   1470
      End
      Begin VB.TextBox txtDBValCFValue 
         Height          =   315
         Index           =   0
         Left            =   2970
         TabIndex        =   107
         Top             =   3915
         Width           =   645
      End
      Begin VB.ComboBox cboDBValCFOperator 
         Height          =   315
         Index           =   0
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   3915
         Width           =   2730
      End
      Begin VB.CheckBox chkConditionalFormatting 
         Caption         =   "Use cond&itional formatting"
         Height          =   195
         Left            =   210
         TabIndex        =   105
         Top             =   3315
         Width           =   2835
      End
      Begin VB.TextBox txtDBValueSuffix 
         Height          =   300
         Left            =   4815
         TabIndex        =   104
         Top             =   2745
         Width           =   1125
      End
      Begin VB.TextBox txtDBValuePrefix 
         Height          =   300
         Left            =   1800
         TabIndex        =   102
         Top             =   2745
         Width           =   1125
      End
      Begin VB.CheckBox chkDBVaUseThousandSeparator 
         Caption         =   "Use &1000 separator"
         Height          =   240
         Left            =   3975
         TabIndex        =   100
         Top             =   2415
         Width           =   2040
      End
      Begin VB.CheckBox chkFormatting 
         Caption         =   "Use &formatting"
         Height          =   195
         Left            =   210
         TabIndex        =   97
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   705
         Width           =   4515
      End
      Begin VB.ComboBox cboParents 
         Height          =   315
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   315
         Width           =   4515
      End
      Begin VB.OptionButton optAggregateType 
         Caption         =   "Count"
         Height          =   285
         Index           =   0
         Left            =   2265
         TabIndex        =   95
         Top             =   1545
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAggregateType 
         Caption         =   "Total"
         Height          =   285
         Index           =   1
         Left            =   3390
         TabIndex        =   96
         Top             =   1545
         Width           =   765
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   5430
         TabIndex        =   92
         ToolTipText     =   "Select Filter"
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtFilter 
         Height          =   330
         Left            =   1575
         TabIndex        =   91
         Top             =   1080
         Width           =   3870
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
         Left            =   5745
         MaskColor       =   &H000000FF&
         TabIndex        =   93
         ToolTipText     =   "Clear Filter"
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label lblCFHeader4 
         AutoSize        =   -1  'True
         Caption         =   "Colour"
         Height          =   195
         Left            =   5160
         TabIndex        =   132
         Top             =   3660
         Width           =   570
      End
      Begin VB.Label lblCFHeader3 
         AutoSize        =   -1  'True
         Caption         =   "Format"
         Height          =   195
         Left            =   3690
         TabIndex        =   131
         Top             =   3660
         Width           =   600
      End
      Begin VB.Label lblCFHeader2 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   2985
         TabIndex        =   130
         Top             =   3660
         Width           =   480
      End
      Begin VB.Label lblCFHeader1 
         AutoSize        =   -1  'True
         Caption         =   "Operator"
         Height          =   195
         Left            =   240
         TabIndex        =   129
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label lblDBValueSample 
         Caption         =   "12345.21"
         Height          =   510
         Left            =   3660
         TabIndex        =   122
         Top             =   5325
         Width           =   2430
      End
      Begin VB.Label lblDBValuePreview 
         AutoSize        =   -1  'True
         Caption         =   "Preview :"
         Height          =   195
         Left            =   2655
         TabIndex        =   125
         Top             =   5310
         Width           =   810
      End
      Begin VB.Label lblConditionWarning 
         AutoSize        =   -1  'True
         Caption         =   "(in priority order)"
         Height          =   195
         Left            =   4530
         TabIndex        =   124
         Top             =   3315
         Width           =   1500
      End
      Begin VB.Label lblDBVSampleHeader 
         Caption         =   "Sample :"
         Height          =   195
         Left            =   210
         TabIndex        =   123
         Top             =   5310
         Width           =   870
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   195
         X2              =   6075
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   195
         X2              =   6075
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Label lblDBValueSuffix 
         AutoSize        =   -1  'True
         Caption         =   "Suffix :"
         Height          =   195
         Left            =   3975
         TabIndex        =   103
         Top             =   2790
         Width           =   630
      End
      Begin VB.Label lblDBValuePrefix 
         AutoSize        =   -1  'True
         Caption         =   "Prefix :"
         Height          =   195
         Left            =   210
         TabIndex        =   101
         Top             =   2790
         Width           =   630
      End
      Begin VB.Label lblDBValueDecimals 
         AutoSize        =   -1  'True
         Caption         =   "Decimal places :"
         Height          =   195
         Left            =   210
         TabIndex        =   98
         Top             =   2430
         Width           =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   195
         X2              =   6075
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblParents 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
         Height          =   195
         Left            =   210
         TabIndex        =   86
         Top             =   330
         Width           =   600
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         Caption         =   "Column :"
         Height          =   195
         Left            =   210
         TabIndex        =   88
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter :"
         Height          =   195
         Left            =   210
         TabIndex        =   90
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label lblAggregateType 
         AutoSize        =   -1  'True
         Caption         =   "Aggregate Function :"
         Height          =   195
         Left            =   210
         TabIndex        =   94
         Top             =   1575
         Width           =   1785
      End
   End
   Begin VB.Frame fraLinkType 
      Caption         =   "Link Type :"
      Height          =   6060
      Left            =   150
      TabIndex        =   9
      Top             =   1920
      Width           =   2500
      Begin VB.OptionButton optLink 
         Caption         =   "&On-screen Document Display"
         Height          =   450
         Index           =   11
         Left            =   195
         TabIndex        =   21
         Top             =   4065
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Today's Events"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   19
         Top             =   3480
         Width           =   2235
      End
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
         Caption         =   "Cha&rt"
         Height          =   315
         Index           =   6
         Left            =   195
         TabIndex        =   16
         Top             =   2385
         Width           =   1305
      End
      Begin VB.OptionButton optLink 
         Caption         =   "Se&parator"
         Height          =   315
         Index           =   5
         Left            =   195
         TabIndex        =   15
         Top             =   2040
         Width           =   1305
      End
      Begin VB.OptionButton optLink 
         Caption         =   "Or&ganisation Chart"
         Height          =   255
         Index           =   10
         Left            =   200
         TabIndex        =   20
         Top             =   3810
         Width           =   2175
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Application"
         Enabled         =   0   'False
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
         Caption         =   "&Screen"
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
         Caption         =   "Report / Utilit&y"
         Height          =   315
         Index           =   2
         Left            =   200
         TabIndex        =   11
         Top             =   650
         Width           =   2265
      End
   End
   Begin VB.Frame fraHRProScreenLink 
      Caption         =   "Screen :"
      Height          =   6060
      Left            =   2880
      TabIndex        =   22
      Top             =   1920
      Width           =   6300
      Begin VB.TextBox txtPageTitle 
         Height          =   315
         Left            =   1575
         MaxLength       =   100
         TabIndex        =   28
         Top             =   1100
         Width           =   4515
      End
      Begin VB.ComboBox cboHRProScreen 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   700
         Width           =   4515
      End
      Begin VB.ComboBox cboHRProTable 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   300
         Width           =   4515
      End
      Begin VB.ComboBox cboStartMode 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":2544
         Left            =   1575
         List            =   "frmSSIntranetLink.frx":2546
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1500
         Width           =   4515
      End
      Begin VB.Label lblPageTitle 
         AutoSize        =   -1  'True
         Caption         =   "Page Title :"
         Height          =   195
         Left            =   200
         TabIndex        =   27
         Top             =   1160
         Width           =   810
      End
      Begin VB.Label lblHRProScreen 
         Caption         =   "Screen :"
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblHRProTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblStartMode 
         AutoSize        =   -1  'True
         Caption         =   "Start Mode :"
         Height          =   195
         Left            =   200
         TabIndex        =   29
         Top             =   1560
         Width           =   900
      End
   End
   Begin VB.Frame fraEmailLink 
      Caption         =   "Email Link :"
      Height          =   1245
      Left            =   2880
      TabIndex        =   47
      Top             =   5310
      Width           =   6300
      Begin VB.TextBox txtEmailSubject 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   51
         Top             =   700
         Width           =   4515
      End
      Begin VB.TextBox txtEmailAddress 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   49
         Top             =   300
         Width           =   4515
      End
      Begin VB.Label lblEmailSubject 
         AutoSize        =   -1  'True
         Caption         =   "Email Subject :"
         Height          =   195
         Left            =   195
         TabIndex        =   50
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label lblEMailAddress 
         AutoSize        =   -1  'True
         Caption         =   "Email Address :"
         Height          =   195
         Left            =   195
         TabIndex        =   48
         Top             =   360
         Width           =   1320
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   810
      Top             =   11535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   6600
      TabIndex        =   56
      Top             =   12150
      Width           =   2600
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   135
         TabIndex        =   133
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   134
         Top             =   0
         Width           =   1200
      End
   End
   Begin COAColourPicker.COA_ColourPicker ColorPicker 
      Left            =   240
      Top             =   11535
      _ExtentX        =   820
      _ExtentY        =   820
      ShowSysColorButton=   0   'False
   End
   Begin VB.Frame fraDocument 
      Caption         =   "Document :"
      Height          =   1125
      Left            =   150
      TabIndex        =   52
      Top             =   5895
      Width           =   9045
      Begin VB.TextBox txtDocumentFilePath 
         Height          =   315
         Left            =   1400
         MaxLength       =   500
         TabIndex        =   54
         Top             =   300
         Width           =   7365
      End
      Begin VB.CheckBox chkDisplayDocumentHyperlink 
         Caption         =   "Displa&y hyperlink to document"
         Height          =   330
         Left            =   1395
         TabIndex        =   55
         Top             =   690
         Width           =   3720
      End
      Begin VB.Label lblDocumentFilePath 
         AutoSize        =   -1  'True
         Caption         =   "URL :"
         Height          =   195
         Left            =   195
         TabIndex        =   53
         Top             =   360
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmSSIntranetLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private mblnReadOnly As Boolean
Private miInitialDisplayMode As Integer
Private mlngChart_TableID_2 As Long
Private mlngChart_ColumnID_2 As Long
Private mlngChart_TableID_3 As Long
Private mlngChart_ColumnID_3 As Long
Private mlngChart_SortOrderID As Long
Private miChart_SortDirection As Integer
Private mlngChart_ColourID As Long

Private mcolSSITableViews As clsSSITableViews
Private mcolGroups As Collection

Private gForeColour As ColorConstants
Private jnCount As Integer



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
 
  ' Screen links only required for Button or Dropdown List Links
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
    
'    Enabled for v4.2 as multi-dimensional charts are now configurable
    .AddItem "3D Line"
    .ItemData(.NewIndex) = 2
        
'    Enabled for v4.2 as multi-dimensional charts are now configurable
    .AddItem "2D Line"
    .ItemData(.NewIndex) = 3
        
    .AddItem "3D Area"
    .ItemData(.NewIndex) = 4
        
'    Enabled for v4.2 as multi-dimensional charts are now configurable
    .AddItem "2D Area"
    .ItemData(.NewIndex) = 5
        
    .AddItem "3D Step"
    .ItemData(.NewIndex) = 6
        
    .AddItem "2D Step"
    .ItemData(.NewIndex) = 7
        
    .AddItem "Pie"
    .ItemData(.NewIndex) = 14
        
'   Leave 2d xy disabled, as it requires data to be in a 2-column format which isn't
'   catered for at present.
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
    
    If ASRDEVELOPMENT Or Application.NineBoxGridModule Then
      .AddItem "9-Box Grid Report"
      .ItemData(.NewIndex) = utlNineBoxGrid
    End If
      
    .AddItem "Calendar Report"
    .ItemData(.NewIndex) = utlCalendarreport
        
    .AddItem "Custom Report"
    .ItemData(.NewIndex) = utlCustomReport
    
    .AddItem "Mail Merge"
    .ItemData(.NewIndex) = utlMailMerge
    
    .AddItem "Organisation Report"
    .ItemData(.NewIndex) = utlOrganisation
    
    .AddItem "Talent Report"
    .ItemData(.NewIndex) = utlTalent
             
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
  Dim rsTables As DAO.Recordset
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
  Dim rsLocalUtilities As DAO.Recordset
  Dim sTableName As String
  Dim sIDColumnName As String
  Dim fLocalTable As Boolean
  
  fLocalTable = False
  
  cboHRProUtility.Clear

  Select Case pUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sIDColumnName = "ID"
      
    Case utlCalendarreport
      sTableName = "ASRSysCalendarReports"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sIDColumnName = "CrossTabID"
      sWhereSQL = "ASRSysCrossTab.CrossTabType = " & CStr(cttNormal)
    
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
        
    Case utlNineBoxGrid
       sTableName = "ASRSysCrossTab"
      sIDColumnName = "CrossTabID"
      sWhereSQL = "ASRSysCrossTab.CrossTabType = " & CStr(ctt9GridBox)
      
    Case utlTalent
      sTableName = "ASRSysTalentReports"
      sIDColumnName = "ID"
     
    Case utlOrganisation
      sTableName = "ASRSysOrganisationReport"
      sIDColumnName = "ID"
     
      
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
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsLocalUtilities!ID

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
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsUtilities!ID
  
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
  Dim rsScreens As DAO.Recordset

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
                      psPrompt As String, psText As String, psHRProScreenID As String, psPageTitle As String, _
                      psURL As String, plngTableID As Long, psStartMode As String, plngViewID As Long, _
                      psUtilityType As String, psUtilityID As String, pfCopy As Boolean, psHiddenGroups As String, _
                      psTableViewName As String, pfNewWindow As Boolean, psEMailAddress As String, psEMailSubject As String, _
                      psAppFilePath As String, psAppParameters As String, _
                      psDocumentFilePath As String, pfDisplayDocumentHyperlink As Boolean, _
                      piElement_Type As Integer, piSeparatorOrientation As Integer, plngPictureID As Long, _
                      pfChartShowLegend As Boolean, piChartType As Integer, pfChartShowGrid As Boolean, _
                      pfChartStackSeries As Boolean, plngChartViewID As Long, miChartTableID As Long, _
                      plngChartColumnID As Long, plngChartFilterID As Long, piChartAggregateType As Integer, _
                      pfChartShowValues As Boolean, pcolGroups As Collection, _
                      pfUseFormatting As Boolean, piFormatting_DecimalPlaces As Integer, pfFormatting_Use1000Separator As Boolean, _
                      psFormatting_Prefix As String, psFormatting_Suffix As String, pfUseConditionalFormatting As Boolean, _
                      psConditionalFormatting_Operator_1 As String, psConditionalFormatting_Value_1 As String, psConditionalFormatting_Style_1 As String, psConditionalFormatting_Colour_1 As String, _
                      psConditionalFormatting_Operator_2 As String, psConditionalFormatting_Value_2 As String, psConditionalFormatting_Style_2 As String, psConditionalFormatting_Colour_2 As String, _
                      psConditionalFormatting_Operator_3 As String, psConditionalFormatting_Value_3 As String, psConditionalFormatting_Style_3 As String, psConditionalFormatting_Colour_3 As String, _
                      psSeparatorColour As String, pfShowPercentages As Boolean, _
                      ByRef pcolSSITableViews As clsSSITableViews)
   
  Set mcolSSITableViews = pcolSSITableViews
  Set mcolGroups = pcolGroups
  
  mfLoading = True
  
  miLinkType = piType
  'mlngPersonnelTableID = plngPersonnelTableID
  mlngTableID = plngTableID
  mlngViewID = plngViewID
  msTableViewName = psTableViewName
    
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
    optLink(SSINTLINKTODAYS_EVENTS).Enabled = False
    ' fault HRPRO-907 - disable separators for all but dashboard and hypertext links
    If miLinkType <> SSINTLINK_HYPERTEXT Then optLink(SSINTLINKSEPARATOR).Enabled = False
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
  ElseIf piElement_Type = 5 Then
    optLink(SSINTLINKTODAYS_EVENTS).value = True
  ElseIf piElement_Type = 6 Then
    optLink(SSINTLINKORGCHART).value = True
  End If
 

  
  GetHRProTables
  
  If optLink(SSINTLINKSCREEN_UTILITY).value Then
    GetHRProUtilityTypes
  End If
  
  'If psUtilityType = "" Then
  '  cboHRProUtilityType.ListIndex = -1
  '  cboHRProUtility.ListIndex = -1
  'Else
    UtilityType = psUtilityType
    UtilityID = psUtilityID
  'End If
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
  chkPrimaryDisplay.value = IIf(InitialDisplayMode = 1, 1, 0)
  ChartShowPercentages = pfShowPercentages
  If chkShowValues.value = 0 Then ChartShowPercentages = 0
  chkShowPercentages.Enabled = (chkShowValues.value = 1)
  
  ' Set up 'Database Value' combos...
  PopulateParentsCombo (miChartTableID) ' populate and set default value
  PopulateColumnsCombo (cboParents.ItemData(cboParents.ListIndex))
  optAggregateType(0).value = IIf(ChartAggregateType = 0, True, False)
  optAggregateType(1).value = IIf(ChartAggregateType = 1, True, False)
  txtFilter.Tag = miChartFilterID
  txtFilter.Text = GetExpressionName(txtFilter.Tag)
  
  txtFilter.Enabled = False
  txtFilter.BackColor = vbButtonFace
  
  txtChartDescription = Chart_Description
  txtChartUtility = Chart_Utility_Description
  cmdChartReportClear.Enabled = (Len(txtChartUtility.Text) > 0)
  
  UseFormatting = pfUseFormatting
  Formatting_DecimalPlaces = piFormatting_DecimalPlaces
  Formatting_Use1000Separator = pfFormatting_Use1000Separator
  Formatting_Prefix = psFormatting_Prefix
  Formatting_Suffix = psFormatting_Suffix
  UseConditionalFormatting = pfUseConditionalFormatting
  
  PopulateDBValOperatorCombos
  PopulateDBValStyleCombos
  
  ConditionalFormatting_Operator_1 = psConditionalFormatting_Operator_1
  ConditionalFormatting_Value_1 = psConditionalFormatting_Value_1
  ConditionalFormatting_Style_1 = psConditionalFormatting_Style_1
  ConditionalFormatting_Colour_1 = psConditionalFormatting_Colour_1
  ConditionalFormatting_Operator_2 = psConditionalFormatting_Operator_2
  ConditionalFormatting_Value_2 = psConditionalFormatting_Value_2
  ConditionalFormatting_Style_2 = psConditionalFormatting_Style_2
  ConditionalFormatting_Colour_2 = psConditionalFormatting_Colour_2
  ConditionalFormatting_Operator_3 = psConditionalFormatting_Operator_3
  ConditionalFormatting_Value_3 = psConditionalFormatting_Value_3
  ConditionalFormatting_Style_3 = psConditionalFormatting_Style_3
  ConditionalFormatting_Colour_3 = psConditionalFormatting_Colour_3
      
  SeparatorBorderColour = psSeparatorColour
        
  ' Hide Workflow items if not licensed
  If Not IsModuleEnabled(modWorkflow) Then optLink(SSINTLINKPWFSTEPS).Visible = False
  
  PopulateAccessGrid psHiddenGroups

  mfChanged = False
  
  
  
  If pfCopy Then mfChanged = True
  
  RefreshControls
  
  mfLoading = False
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
  fraLinkSeparator.Visible = (optLink(SSINTLINKSEPARATOR).value Or optLink(SSINTLINKPWFSTEPS).value Or optLink(SSINTLINKTODAYS_EVENTS).value Or optLink(SSINTLINKORGCHART).value)
  fraChartLink.Visible = optLink(SSINTLINKCHART).value
  fraDBValue.Visible = optLink(SSINTLINKDB_VALUE).value
      
  ' Disable the screen controls as required.
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
  If Not optLink(SSINTLINKSCREEN_UTILITY).value And Not optLink(SSINTLINKCHART).value Then
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
  txtAppFilePath.Enabled = False  'optLink(SSINTLINKSCREEN_APPLICATION).value
  txtAppFilePath.BackColor = IIf(txtAppFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblAppFilePath.Enabled = False  'txtAppFilePath.Enabled
  txtAppParameters.Enabled = False  'optLink(SSINTLINKSCREEN_APPLICATION).value
  txtAppParameters.BackColor = IIf(txtAppFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblAppParameters.Enabled = False  'txtAppFilePath.Enabled
'  If Not txtAppFilePath.Enabled Then
'    txtAppFilePath.Text = ""
'    txtAppParameters.Text = ""
'  End If

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
      Or optLink(SSINTLINKPWFSTEPS).value Or optLink(SSINTLINKDB_VALUE).value _
      Or optLink(SSINTLINKTODAYS_EVENTS).value Or optLink(SSINTLINKORGCHART).value Then
        
    txtPrompt.Enabled = False
    txtPrompt.BackColor = vbButtonFace
    
    If optLink(SSINTLINKPWFSTEPS).value Or optLink(SSINTLINKTODAYS_EVENTS).value Then
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
    If optLink(SSINTLINKTODAYS_EVENTS).value Then txtPrompt.Text = "<TODAYS EVENTS>"
    If optLink(SSINTLINKORGCHART).value Then txtPrompt.Text = "<ORGCHART>"
    
    If optLink(SSINTLINKCHART).value Then
      MSChart1.RowCount = 1
    End If
    
    ' Chart DataGrid formatting
    lblDataGridDecimals.Enabled = chkFormatting
    If chkDataGridUseFormatting Then
      spnDataGridDecimals.Enabled = True
      lblDataGridDecimals.Enabled = True
    Else
      spnDataGridDecimals.Enabled = False
      lblDataGridDecimals.Enabled = False
      spnDataGridDecimals.value = 0
    End If
    chkDataGrid1000Separator.Enabled = chkDataGridUseFormatting
    
    ' disable the DB Value items
    lblDBValueDecimals.Enabled = chkFormatting
    If chkFormatting And Not optAggregateType(0).value Then
      spnDBValueDecimals.Enabled = True
    Else
      spnDBValueDecimals.Enabled = False
      spnDBValueDecimals.value = 0
    End If
    chkDBVaUseThousandSeparator.Enabled = chkFormatting
    lblDBValuePrefix.Enabled = chkFormatting
    txtDBValuePrefix.Enabled = chkFormatting
    txtDBValuePrefix.BackColor = IIf(chkFormatting, vbWindowBackground, vbButtonFace)
    lblDBValueSuffix.Enabled = chkFormatting
    txtDBValueSuffix.Enabled = chkFormatting
    txtDBValueSuffix.BackColor = IIf(chkFormatting, vbWindowBackground, vbButtonFace)
    
    For jnCount = 0 To 2
      cboDBValCFOperator(jnCount).Enabled = chkConditionalFormatting
      txtDBValCFValue(jnCount).Enabled = chkConditionalFormatting
      txtDBValCFValue(jnCount).BackColor = IIf(chkConditionalFormatting, vbWindowBackground, vbButtonFace)
      cboDBValCFStyle(jnCount).Enabled = chkConditionalFormatting
'      txtDBValCFColour(jnCount).Enabled = chkConditionalFormatting
      ' txtDBValCFColour(jnCount).BackColor = IIf(chkConditionalFormatting, vbWindowBackground, vbButtonFace)
      cmdDBValueColPick(jnCount).Enabled = chkConditionalFormatting And (cboDBValCFStyle(jnCount).Text <> "Hidden")
      If cboDBValCFStyle(jnCount).Text = "Hidden" Then
        txtDBValCFColour(jnCount).BackColor = &HFFFFFF
      End If
    Next
    
    lblSeparatorColour.Enabled = chkSeparatorUseFormatting
    txtSeparatorColour.Enabled = chkSeparatorUseFormatting
    cmdSeparatorColPick.Enabled = chkSeparatorUseFormatting
    If chkSeparatorUseFormatting = False Then
      ' reset the colour to white
      txtSeparatorColour.BackColor = &HFFFFFF
    End If
    
    txtDBValueSample.Enabled = chkFormatting Or chkConditionalFormatting
    lblDBValuePreview.Enabled = chkFormatting Or chkConditionalFormatting
    lblDBValueSample.Enabled = chkFormatting Or chkConditionalFormatting
    lblDBVSampleHeader.Enabled = chkFormatting Or chkConditionalFormatting
    lblCFHeader1.Enabled = chkConditionalFormatting
    lblCFHeader2.Enabled = chkConditionalFormatting
    lblCFHeader3.Enabled = chkConditionalFormatting
    lblCFHeader4.Enabled = chkConditionalFormatting
        
  cmdChartReportClear.Enabled = (Len(txtChartUtility.Text) > 0)
        
        
  Else
    ' NPG20100427 Fault HRPRO-888
    If Not txtPrompt.Enabled Then txtPrompt.Text = ""
    
    txtPrompt.Enabled = True
    txtPrompt.BackColor = vbWindowBackground
    ' NPG20100427 Fault HRPRO-910
    txtText.Enabled = True
    txtText.BackColor = vbWindowBackground
  End If
  
  If optLink(SSINTLINKPWFSTEPS).value Or _
          miLinkType = SSINTLINK_DROPDOWNLIST Or _
          optLink(SSINTLINKTODAYS_EVENTS).value Or _
          optLink(SSINTLINKORGCHART).value Then
    ' Disable the icon and new column options for dropdown lists & pwfs...
    chkNewColumn.Visible = False
    lblIcon.Visible = False
    txtIcon.Visible = False
    cmdIcon.Visible = False
    cmdIconClear.Visible = False
    cmdIconClear.Enabled = False
    imgIcon.Visible = False
    lblNoOptions.Visible = True
    lblNoOptions.Top = 345
    lblSeparatorColour.Visible = False
    txtSeparatorColour.Visible = False
    cmdSeparatorColPick.Visible = False
    chkSeparatorUseFormatting.Visible = False
    Line4.Visible = False
    lblDiaryWarning.Visible = (optLink(SSINTLINKTODAYS_EVENTS))
    lblDiaryWarning.Caption = "The following items will appear for all users in the selected visibility groups :" & vbCrLf & vbCrLf & _
    "   Manual Diary Events" & vbCrLf & _
    "   System Diary Events" & vbCrLf & _
    "   Outlook Calendar Links" & vbCrLf & _
    "   Attendance Status (Workflow Out of Office)"
    lblDiaryWarning.Top = 745
    
  ElseIf optLink(SSINTLINKSEPARATOR).value Then
    ' Enable the icon and new column options for dashboard link separators...
    lblIcon.Visible = True
    txtIcon.Visible = True
    cmdIcon.Visible = True
    cmdIcon.Enabled = False
    cmdIconClear.Visible = True
    cmdIconClear.Enabled = False  'txtIcon.Text <> ""
    imgIcon.Visible = True
    chkNewColumn.Visible = (miLinkType = SSINTLINK_BUTTON)
    lblNoOptions.Visible = False
    lblNoOptions.Top = 345
    lblDiaryWarning.Visible = False
    lblSeparatorColour.Visible = (miLinkType = SSINTLINK_BUTTON)
    txtSeparatorColour.Visible = (miLinkType = SSINTLINK_BUTTON)
    cmdSeparatorColPick.Visible = (miLinkType = SSINTLINK_BUTTON)
    chkSeparatorUseFormatting.Visible = (miLinkType = SSINTLINK_BUTTON)
    Line4.Visible = (miLinkType = SSINTLINK_BUTTON)
  Else
    lblNoOptions.Visible = False
    lblNoOptions.Top = 345
    lblDiaryWarning.Visible = False
    cmdIconClear.Enabled = txtIcon.Text <> ""
    cmdFilterClear.Enabled = txtFilter.Text <> ""
  End If
  
  
  
  If optLink(SSINTLINKPWFSTEPS).value Then
    fraLinkSeparator.Caption = "Pending Workflow Steps :"
  ElseIf optLink(SSINTLINKTODAYS_EVENTS).value Then
    fraLinkSeparator.Caption = "Today's Events :"
  ElseIf optLink(SSINTLINKORGCHART).value Then
    fraLinkSeparator.Caption = "Organisation Chart :"
  Else
    fraLinkSeparator.Caption = "Separator :"
  End If
  
  If optLink(SSINTLINKCHART).value And cboChartType.ListIndex >= 0 Then
    
    Select Case cboChartType.ItemData(cboChartType.ListIndex)
      Case 0  '3D Bar
        
      Case 1  '2D Bar
            
      Case 4  '3D Area
        
      Case 6  '3D Step
    
      Case 7  '2D Step
        
      Case 14 '2D Pie


    End Select
  
  End If
  
  If optLink(SSINTLINKDB_VALUE) Then RefreshDBValueSample
  
  mblnRefreshing = True
  GetStartModes
  mblnRefreshing = False
  
  lblHRProUtilityMessage.Caption = sUtilityMessage
  
  ' Disable the OK button as required.
  cmdOk.Enabled = mfChanged
  

End Sub

Private Function RefreshDBValueSample()
  ' Display the sample value with all the formatting applied
  Dim strFormatString As String
  Dim pfBold As Boolean
  Dim pfItalic As Boolean
  
  If mfLoading Then Exit Function
  
  strFormatString = IIf(chkDBVaUseThousandSeparator, "###,###,###", "")
  If spnDBValueDecimals.value > 0 Then
    strFormatString = strFormatString & "."
    For jnCount = 1 To spnDBValueDecimals.value
      strFormatString = strFormatString + "0"
    Next
  End If
  
  lblDBValueSample.ForeColor = vbButtonText
  lblDBValueSample.FontBold = False
  lblDBValueSample.FontItalic = False
  lblDBValueSample.Visible = True
  
  If chkConditionalFormatting Then
    ' Conditional formatting
    ' Conditions are in priority order!
    For jnCount = 0 To 2
    If IsNumeric(txtDBValueSample) And IsNumeric(txtDBValCFValue(jnCount)) Then
      pfBold = cboDBValCFStyle(jnCount) = "Bold" Or cboDBValCFStyle(jnCount) = "Bold & Italic"
      pfItalic = cboDBValCFStyle(jnCount) = "Italic" Or cboDBValCFStyle(jnCount) = "Bold & Italic"
    
      Select Case cboDBValCFOperator(jnCount)
        Case "is equal to"
          If CDec(txtDBValueSample) = CDec(txtDBValCFValue(jnCount)) Then
            lblDBValueSample.ForeColor = txtDBValCFColour(jnCount).BackColor
            lblDBValueSample.FontBold = pfBold
            lblDBValueSample.FontItalic = pfItalic
            lblDBValueSample.Visible = IIf(cboDBValCFStyle(jnCount) = "Hidden", False, True)
            Exit For
          End If
      
        Case "is not equal to"
          If CDec(txtDBValueSample) <> CDec(txtDBValCFValue(jnCount)) Then
            lblDBValueSample.ForeColor = txtDBValCFColour(jnCount).BackColor
            lblDBValueSample.FontBold = pfBold
            lblDBValueSample.FontItalic = pfItalic
            lblDBValueSample.Visible = IIf(cboDBValCFStyle(jnCount) = "Hidden", False, True)
            Exit For
          End If
      
        Case "is less than or equal to"
          If CDec(txtDBValueSample) <= CDec(txtDBValCFValue(jnCount)) Then
            lblDBValueSample.ForeColor = txtDBValCFColour(jnCount).BackColor
            lblDBValueSample.FontBold = pfBold
            lblDBValueSample.FontItalic = pfItalic
            lblDBValueSample.Visible = IIf(cboDBValCFStyle(jnCount) = "Hidden", False, True)
            Exit For
          End If
      
        Case "is greater than or equal to"
          If CDec(txtDBValueSample) >= CDec(txtDBValCFValue(jnCount)) Then
            lblDBValueSample.ForeColor = txtDBValCFColour(jnCount).BackColor
            lblDBValueSample.FontBold = pfBold
            lblDBValueSample.FontItalic = pfItalic
            lblDBValueSample.Visible = IIf(cboDBValCFStyle(jnCount) = "Hidden", False, True)
            Exit For
          End If
      
        Case "is less than"
          If CDec(txtDBValueSample) < CDec(txtDBValCFValue(jnCount)) Then
            lblDBValueSample.ForeColor = txtDBValCFColour(jnCount).BackColor
            lblDBValueSample.FontBold = pfBold
            lblDBValueSample.FontItalic = pfItalic
            lblDBValueSample.Visible = IIf(cboDBValCFStyle(jnCount) = "Hidden", False, True)
            Exit For
          End If
      
        Case "is greater than"
          If CDec(txtDBValueSample) > CDec(txtDBValCFValue(jnCount)) Then
            lblDBValueSample.ForeColor = txtDBValCFColour(jnCount).BackColor
            lblDBValueSample.FontBold = pfBold
            lblDBValueSample.FontItalic = pfItalic
            lblDBValueSample.Visible = IIf(cboDBValCFStyle(jnCount) = "Hidden", False, True)
            Exit For
          End If
              
      End Select
    End If
    Next
  End If
  
  If chkFormatting Then
    lblDBValueSample = txtDBValuePrefix + Format(txtDBValueSample, strFormatString) + txtDBValueSuffix
  Else
    lblDBValueSample = txtDBValueSample
  End If
End Function

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
  

  ' disable dotted gridlines option for pie charts
  chkDottedGridlines.value = IIf(MSChart1.ChartType = VtChChartType2dPie, 0, chkDottedGridlines.value)
  chkDottedGridlines.Enabled = (MSChart1.ChartType <> VtChChartType2dPie)
    
  ' disable and clear stacking option
  chkStackSeries.value = IIf(MSChart1.ChartType = VtChChartType2dPie, 0, chkStackSeries.value)
  chkStackSeries.Enabled = (MSChart1.ChartType <> VtChChartType2dPie)

  
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
        If chkShowPercentages Then
          .Plot.SeriesCollection.Item(iCount).DataPoints.Item(-1).DataPointLabel.Component = VtChLabelComponentPercent
          .Plot.SeriesCollection.Item(iCount).DataPoints.Item(-1).DataPointLabel.PercentFormat = "0%"
        Else
          .Plot.SeriesCollection.Item(iCount).DataPoints.Item(-1).DataPointLabel.Component = VtChLabelComponentValue
        End If
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
  
  ' Check that text has been entered
  If fValid Then
    If (Len(txtText.Text) = 0) And Not optLink(SSINTLINKPWFSTEPS).value And Not optLink(SSINTLINKSEPARATOR).value And Not optLink(SSINTLINKTODAYS_EVENTS).value Then
      fValid = False
      MsgBox "No text has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtText.SetFocus
    End If
  End If
  
  ' Check that text Text cannot contain apostrophes for workflows
  If fValid Then
    If cboHRProUtilityType.ListIndex >= 0 Then
      If cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex) = utlWorkflow Then
        If InStr(txtText.Text, "'") > 0 Then
          fValid = False
          MsgBox "This Link Text cannot contain apostrophes.", vbOKOnly + vbExclamation, Application.Name
          txtText.SetFocus
        End If
      End If
    End If
  End If
  
  ' Check that the screen has been selected (if required)
  If fValid Then
    If (miLinkType <> SSINTLINK_HYPERTEXT) And _
      (optLink(SSINTLINKSCREEN_HRPRO).value) Then
      If cboHRProScreen.ListIndex < 0 Then
        fValid = False
        MsgBox "No screen has been selected.", vbOKOnly + vbExclamation, Application.Name
        cboHRProTable.SetFocus
      End If
    End If
  End If
  
  ' Check that the page title been entered
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
      MsgBox "No chart data has been defined.", vbOKOnly + vbExclamation, Application.Name
      cmdChartData.SetFocus
    End If
  End If

  ' Only one Pending Workflow Steps per security group...
  If fValid Then
    If optLink(SSINTLINKPWFSTEPS).value Then
      ' loop through the chosen security groups and check they're in the combined string
      
      psDuplicateGroups = ""
      
      With grdAccess
        For iLoop = 1 To (.Rows - 1)  ' exclude item 0 as it's the '(All Groups)' item.
          .Bookmark = .AddItemBookmark(iLoop)
          If .Columns("Access").value Then
            If mcolGroups(.Columns("GroupName").Text & "3").Allow = False Then
              fValid = False
              psDuplicateGroups = psDuplicateGroups & vbCrLf & .Columns("GroupName").Text
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

  ' Only one Today's Events per security group...
  If fValid Then
    If optLink(SSINTLINKTODAYS_EVENTS).value Then
      ' loop through the chosen security groups and check they're in the combined string
      
      psDuplicateGroups = ""
      
      With grdAccess
        For iLoop = 1 To (.Rows - 1)  ' exclude item 0 as it's the '(All Groups)' item.
          .Bookmark = .AddItemBookmark(iLoop)
          If .Columns("Access").value Then
            If mcolGroups(.Columns("GroupName").Text & "5").Allow = False Then
              fValid = False
              psDuplicateGroups = psDuplicateGroups & vbCrLf & .Columns("GroupName").Text
            End If
          End If
        Next iLoop
      .MoveFirst
      End With
      
      If Not fValid Then
        MsgBox "'Today's Events' can only be defined once per user group." & vbCrLf & _
                "It has already been defined for the following groups:" & vbCrLf & _
                psDuplicateGroups, vbOKOnly + vbExclamation, Application.Name
        grdAccess.SetFocus
      End If
    End If
  End If
  
  ' DBValues - fault HRPRO-1256
  If fValid Then
    If optLink(SSINTLINKDB_VALUE).value Then
      If ChartColumnID = 0 Then
        fValid = False
        MsgBox "No Database Value data has been defined.", vbOKOnly + vbExclamation, Application.Name
      End If
      
      For jnCount = 0 To 2
      If Trim(cboDBValCFOperator(jnCount).Text) <> vbNullString And cboDBValCFStyle(jnCount).Text = vbNullString Then
        fValid = False
        MsgBox "Conditional formatting style has not been defined.", vbOKOnly + vbExclamation, Application.Name
        Exit For
      End If
      Next
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
        cmdIconClear.Enabled = True
      End If
    End With
  Else
    Set imgIcon.Picture = LoadPicture(vbNullString)
    txtIcon.Text = vbNullString
    cmdIconClear.Enabled = False
  End If
End Function

Private Sub PopulateDBValOperatorCombos()
  For jnCount = 0 To 2
    cboDBValCFOperator(jnCount).Clear
    cboDBValCFOperator(jnCount).AddItem ""
    cboDBValCFOperator(jnCount).AddItem "is equal to"
    cboDBValCFOperator(jnCount).AddItem "is NOT equal to"
    cboDBValCFOperator(jnCount).AddItem "is less than or equal to"
    cboDBValCFOperator(jnCount).AddItem "is greater than or equal to"
    cboDBValCFOperator(jnCount).AddItem "is greater than"
    cboDBValCFOperator(jnCount).AddItem "is less than"
  Next
End Sub

Private Sub PopulateDBValStyleCombos()
  For jnCount = 0 To 2
    cboDBValCFStyle(jnCount).Clear
    If cboDBValCFOperator(jnCount).Text = vbNullString Then
      cboDBValCFStyle(jnCount).AddItem ""
    End If
    cboDBValCFStyle(jnCount).AddItem "Normal"
    cboDBValCFStyle(jnCount).AddItem "Bold"
    cboDBValCFStyle(jnCount).AddItem "Italic"
    cboDBValCFStyle(jnCount).AddItem "Bold & Italic"
    cboDBValCFStyle(jnCount).AddItem "Hidden"
  Next
End Sub

Private Sub PopulateParentsCombo(plngDefaultID As Long)
  
  Dim i As Integer
  ' Clear the contents of the combo.
  cboParents.Clear
  
  cboParents.AddItem ""
  
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
        If recColEdit!DataType <> dtLONGVARCHAR And _
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
 
  ' Display new chart details
  RefreshChart
  If Not mfLoading Then
    mfChanged = True
    RefreshControls
  End If
End Sub

Private Sub cboColumns_Click()
  Dim piColumnDataType As Integer
  Dim lngColumnID As Long
  Dim jniCount As Integer
  
  mfChanged = True

  miChartColumnID = cboColumns.ItemData(cboColumns.ListIndex)
  
  lngColumnID = cboColumns.ItemData(cboColumns.ListIndex)
  
  piColumnDataType = GetColumnDataType(lngColumnID)
  
  ' Disable 'total' option if not numeric or integer
  If piColumnDataType <> dtINTEGER And piColumnDataType <> dtNUMERIC Then
    optAggregateType(0).value = True
    optAggregateType(1).Enabled = False
    optAggregateType(1).ForeColor = vbButtonFace
  Else
    optAggregateType(1).Enabled = True
    optAggregateType(1).ForeColor = vbWindowBackground
  End If
  
  If optLink(SSINTLINKDB_VALUE).value Then
  ' Clear out the formatting options - fault HRPRO-1145
  chkConditionalFormatting.value = 0
  chkFormatting.value = 0
  spnDBValueDecimals.value = 0
  chkDBVaUseThousandSeparator.value = 0
  txtDBValuePrefix.Text = vbNullString
  txtDBValueSuffix.Text = vbNullString
  chkConditionalFormatting.value = 0
  For jniCount = 0 To 2
    If cboDBValCFOperator(jniCount).ListCount > 0 Then
      cboDBValCFOperator(jniCount).ListIndex = 0
    End If
  Next
  End If
  RefreshControls
End Sub

Private Sub cboDBValCFOperator_Click(Index As Integer)
  
  If cboDBValCFOperator(Index).Text = vbNullString Then
    ' set all other rows to empty
    txtDBValCFValue(Index) = vbNullString
    cboDBValCFStyle(Index).ListIndex = -1
    txtDBValCFColour(Index).BackColor = &HFFFFFF
  End If
  
  ' remove the 'blank' option fro the format combo if this operator is not set to nullstring
  'PopulateDBValStyleCombos

  mfChanged = True
  RefreshControls
End Sub

Private Sub cboDBValCFStyle_Click(Index As Integer)
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboParents_Click()

  

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
  
  
  If Not mfLoading Then
    mfChanged = True
    RefreshControls
  End If
End Sub

Private Sub chkConditionalFormatting_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkDataGrid1000Separator_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkDataGridUseFormatting_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkDBVaUseThousandSeparator_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkFormatting_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkPrimaryDisplay_Click()
  mfChanged = True

  InitialDisplayMode = IIf(chkPrimaryDisplay.value, 1, 0)
  
  RefreshControls
End Sub

Private Sub chkSeparatorUseFormatting_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub chkShowPercentages_Click()
If mfLoading Then Exit Sub
  mfChanged = True
  ' refresh the chart
  RefreshChart
  
  RefreshControls
End Sub

Private Sub cmdChartReport_Click()
  Dim frmSSIUtility As New frmSSIntranetUtility

  With frmSSIUtility
    ' .UtilityID = UtilityID
    ' .UtilityType = UtilityType
    .Initialize UtilityType, UtilityID
    
    .Show vbModal
    
    If Not .Cancelled Then
      
      ' Set the utilitytype and utility id for charts...
      GetHRProUtilityTypes
      
      UtilityType = .UtilityType
      UtilityID = .UtilityID
      
      txtChartUtility = Chart_Utility_Description
      cmdChartReportClear.Enabled = (Len(txtChartUtility.Text) > 0)

      mfChanged = True
      
      RefreshControls
      
    End If
  
    UnLoad frmSSIUtility
    Set frmSSIUtility = Nothing
    
  End With

End Sub

Private Sub cmdChartReportClear_Click()

    ' Reset the utilitytype and id to -1
    cboHRProUtilityType.ListIndex = -1
    cboHRProUtility.ListIndex = -1


  txtChartUtility = Chart_Utility_Description
  mfChanged = True
  RefreshControls
End Sub

Private Sub cmdDBValueColPick_Click(Index As Integer)
  On Error GoTo ErrorTrap
  
  gForeColour = txtDBValCFColour(Index).BackColor
  
  With ColorPicker
    ' Set the colour properties of the dialogue box.
    .Color = gForeColour
    ' Display the dialogue box.
    .ShowPalette
    ' Read the colour properties of the dialogue box.
    gForeColour = .Color
  End With

  txtDBValCFColour(Index).BackColor = gForeColour
  txtDBValCFColour(Index).ForeColor = UI.GetInverseColor(gForeColour)
    
  mfChanged = True
  
  RefreshControls
ErrorTrap:
  ' User pressed cancel.
  
End Sub

Private Sub cmdFilterClear_Click()
  txtFilter.Text = vbNullString
  txtFilter.Tag = 0
  miChartFilterID = 0
  
  cmdFilterClear.Enabled = txtFilter.Text <> ""
  mfChanged = True
    
  RefreshControls
End Sub

Private Sub cmdSeparatorColPick_Click()
  On Error GoTo ErrorTrap
  gForeColour = txtSeparatorColour.BackColor
  With ColorPicker
    ' Set the colour properties of the dialogue box.
    .Color = gForeColour
    ' Display the dialogue box.
    .ShowPalette
    ' Read the colour properties of the dialogue box.
    gForeColour = .Color
  End With

  txtSeparatorColour.BackColor = gForeColour
  txtSeparatorColour.ForeColor = UI.GetInverseColor(gForeColour)
    
  mfChanged = True
  
  RefreshControls
ErrorTrap:
  ' User pressed cancel.
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select

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
      cmdFilterClear.Enabled = txtFilter.Text <> ""
      miChartFilterID = .ExpressionID
      mfChanged = True
      
      RefreshControls
      
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
      End If
      
      cmdFilterClear.Enabled = txtFilter.Text <> ""
      
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
  
  If chkShowValues.value = 0 Then ChartShowPercentages = 0
  
  chkShowPercentages.Enabled = (chkShowValues.value = 1)
  
  ' refresh the chart
  RefreshChart
  RefreshControls
End Sub


Private Sub cmdChartData_Click()
  
  
  Dim frmSSIChart As New frmSSIntranetChart

  With frmSSIChart
    .Initialize 1, ChartTableID, ChartColumnID, ChartFilterID, ChartAggregateType, Chart_TableID_2, Chart_ColumnID_2, Chart_TableID_3, Chart_ColumnID_3, Chart_SortOrderID, Chart_SortDirection, Chart_ColourID
    
    .Show vbModal
    
    If Not .Cancelled Then
      ChartTableID = .ChartTableID
      ChartColumnID = .ChartColumnID
      ChartAggregateType = .ChartAggregateType
      ChartFilterID = .txtFilter.Tag
      Chart_TableID_2 = .Chart_TableID_2
      Chart_ColumnID_2 = .Chart_ColumnID_2
      Chart_TableID_3 = .Chart_TableID_3
      Chart_ColumnID_3 = .Chart_ColumnID_3
      Chart_SortOrderID = .Chart_SortOrderID
      Chart_SortDirection = .Chart_SortDirection
      Chart_ColourID = .Chart_ColourID
      
      txtChartDescription = Chart_Description
      
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
  cmdIconClear.Enabled = False
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
  If Not mfLoading Then
    mfChanged = True
    RefreshControls
  End If
End Sub

Private Sub cboHRProUtilityType_Click()
  If cboHRProUtilityType.ListIndex >= 0 Then
    GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
    If Not mfLoading Then
      mfChanged = True
      RefreshControls
    End If
  End If
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
  
  If Not mfLoading Then
    mfChanged = True
    RefreshControls
  End If

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

  Dim intAnswer As Integer
  
  ' Check if any changes have been made.
  If mfChanged Then
    intAnswer = MsgBox("The link type definition has changed.  Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
    If intAnswer = vbYes Then
      Call cmdOK_Click
      Exit Sub
    ElseIf intAnswer = vbCancel Then
      Exit Sub
    End If
  End If


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
  UtilityType = CStr(utlNineBoxGrid)
  
  'dashboard
  If optLink(SSINTLINKSEPARATOR).value Then
    ElementType = 1
  ElseIf optLink(SSINTLINKCHART).value Then
    ElementType = 2
  ElseIf optLink(SSINTLINKPWFSTEPS).value Then
    ElementType = 3
  ElseIf optLink(SSINTLINKDB_VALUE).value Then
    ElementType = 4
  ElseIf optLink(SSINTLINKTODAYS_EVENTS).value Then
    ElementType = 5
  ElseIf optLink(SSINTLINKORGCHART).value Then
    ElementType = 6
  Else
    ElementType = 0
  End If
  
  If Not optLink(SSINTLINKSCREEN_UTILITY).value Then
    ' Reset the utilitytype and id to -1
    cboHRProUtilityType.ListIndex = -1
    cboHRProUtility.ListIndex = -1
  Else
    ' Reset the utilitytype and id to 0
    SetComboItem cboHRProUtilityType, 0   ' cboHRProUtilityType.ListIndex = -1
    SetComboItem cboHRProUtility, 0  ' cboHRProUtility.ListIndex = -1
  End If
  
  If Not mfLoading Then
    mfChanged = True
  ' mfChanged = False
     RefreshControls
  End If
End Sub

Private Sub txtDBValCFValue1_LostFocus()
'  If Not IsNumeric(txtDBValCFValue1) Then
'    MsgBox "Value must be numeric.", vbOKOnly + vbExclamation, Application.Name
'    txtDBValCFValue1.SetFocus
'  End If
End Sub

Private Sub txtDBValCFValue2_LostFocus()
'  If Not IsNumeric(txtDBValCFValue2) Then
'    MsgBox "Value must be numeric.", vbOKOnly + vbExclamation, Application.Name
'    txtDBValCFValue2.SetFocus
'  End If
End Sub

Private Sub txtDBValCFValue3_LostFocus()
'  If Not IsNumeric(txtDBValCFValue3) Then
'    MsgBox "Value must be numeric.", vbOKOnly + vbExclamation, Application.Name
'    txtDBValCFValue3.SetFocus
'  End If
End Sub

Private Sub spnDataGridDecimals_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub spnDBValueDecimals_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtDBValCFColour_Change(Index As Integer)
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtDBValCFValue_Change(Index As Integer)
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtDBValCFValue_KeyPress(Index As Integer, KeyAscii As Integer)
If Len(txtDBValCFValue(Index)) > 9 And KeyAscii <> 8 Then KeyAscii = 0
  Select Case KeyAscii
    Case 8
      'backspace is OK
    Case 45
      ' minus key is OK
    Case 46
      ' decimal point
    Case 48 To 57
      ' Numbers are fine...
    Case Else
      KeyAscii = 0
  End Select
  
  
  ' If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
  ' 8 = bsp
  ' 46 = .
  ' 45 = -
  
End Sub

Private Sub txtDBValuePrefix_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtDBValueSample_Change()
  RefreshControls
End Sub

Private Sub txtDBValueSample_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      'backspace is OK
    Case 45
      ' minus key is OK
    Case 46
      ' decimal point
    Case 48 To 57
      ' Numbers are fine...
    Case Else
      KeyAscii = 0
  End Select
  
  
  ' If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
  ' 8 = bsp
  ' 46 = .
  ' 45 = -
  
End Sub

Private Sub txtDBValueSuffix_Change()
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
  
  If Not mfLoading Then
    mfChanged = True
    RefreshControls
  End If
End Sub

Private Sub txtPrompt_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtText_Change()
  If Not mfLoading Then
  
    ' Check that text Text cannot contain apostrophes for workflows only
    If cboHRProUtilityType.ListIndex >= 0 Then
      If cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex) = utlWorkflow Then
        If InStr(txtText.Text, "'") > 0 Then
          MsgBox "This Link Text cannot contain apostrophes.", vbOKOnly + vbExclamation, Application.Name
          txtText.SetFocus
        End If
      End If
    End If

    mfChanged = True
    RefreshControls
 End If
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
    ((Not optLink(SSINTLINKSCREEN_UTILITY).value) And (Not optLink(SSINTLINKCHART).value)) Then
    UtilityType = ""
  Else
    UtilityType = CStr(cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex))
  End If

End Property

Public Property Get UtilityID() As String

  If (cboHRProUtility.ListIndex < 0) Or _
    ((Not optLink(SSINTLINKSCREEN_UTILITY).value) And (Not optLink(SSINTLINKCHART).value)) Then
    
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
  Dim rsScreens As DAO.Recordset
  
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
  
  If ((optLink(SSINTLINKSCREEN_UTILITY).value) Or (optLink(SSINTLINKCHART).value)) And _
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

  If ((optLink(SSINTLINKSCREEN_UTILITY).value) Or (optLink(SSINTLINKCHART).value)) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtilityType.ListCount - 1
      If cboHRProUtilityType.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtilityType.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
    
    
    If cboHRProUtilityType.ListIndex >= 0 Then GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
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

Public Property Get ChartShowPercentages() As Boolean
  ChartShowPercentages = chkShowPercentages.value
End Property

Public Property Let ChartShowPercentages(ByVal pfNewValue As Boolean)
  chkShowPercentages.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property


Public Property Get UseFormatting() As Boolean
  If optLink(SSINTLINKDB_VALUE) Then
    UseFormatting = chkFormatting.value
  ElseIf optLink(SSINTLINKSEPARATOR) Then
    UseFormatting = chkSeparatorUseFormatting
  ElseIf optLink(SSINTLINKCHART) Then
    UseFormatting = chkDataGridUseFormatting
  End If
End Property

Public Property Let UseFormatting(ByVal pfNewValue As Boolean)
  If optLink(SSINTLINKDB_VALUE) Then
    chkFormatting.value = IIf(pfNewValue, vbChecked, vbUnchecked)
  ElseIf optLink(SSINTLINKSEPARATOR) Then
    chkSeparatorUseFormatting.value = IIf(pfNewValue, vbChecked, vbUnchecked)
  ElseIf optLink(SSINTLINKCHART) Then
    chkDataGridUseFormatting.value = IIf(pfNewValue, vbChecked, vbUnchecked)
  End If
End Property

Public Property Get Formatting_DecimalPlaces() As Integer
  If optLink(SSINTLINKDB_VALUE) Then
    Formatting_DecimalPlaces = spnDBValueDecimals.value
  ElseIf optLink(SSINTLINKCHART) Then
    Formatting_DecimalPlaces = spnDataGridDecimals.value
  End If
End Property

Public Property Let Formatting_DecimalPlaces(ByVal piNewValue As Integer)
  If optLink(SSINTLINKDB_VALUE) Then
    spnDBValueDecimals.value = piNewValue
  ElseIf optLink(SSINTLINKCHART) Then
    spnDataGridDecimals.value = piNewValue
  End If
End Property

Public Property Get Formatting_Use1000Separator() As Boolean
  If optLink(SSINTLINKDB_VALUE) Then
    Formatting_Use1000Separator = chkDBVaUseThousandSeparator.value
  ElseIf optLink(SSINTLINKCHART) Then
    Formatting_Use1000Separator = chkDataGrid1000Separator.value
  End If
End Property

Public Property Let Formatting_Use1000Separator(ByVal pfNewValue As Boolean)
  If optLink(SSINTLINKDB_VALUE) Then
    chkDBVaUseThousandSeparator.value = IIf(pfNewValue, vbChecked, vbUnchecked)
  ElseIf optLink(SSINTLINKCHART) Then
    chkDataGrid1000Separator.value = IIf(pfNewValue, vbChecked, vbUnchecked)
  End If
End Property

Public Property Get Formatting_Prefix() As String
  Formatting_Prefix = txtDBValuePrefix.Text
End Property

Public Property Let Formatting_Prefix(ByVal psNewValue As String)
  txtDBValuePrefix.Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get Formatting_Suffix() As String
  Formatting_Suffix = txtDBValueSuffix.Text
End Property

Public Property Let Formatting_Suffix(ByVal psNewValue As String)
  txtDBValueSuffix.Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get UseConditionalFormatting() As Boolean
  UseConditionalFormatting = chkConditionalFormatting.value
End Property

Public Property Let UseConditionalFormatting(ByVal pfNewValue As Boolean)
  chkConditionalFormatting.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get ConditionalFormatting_Operator_1() As String
  ConditionalFormatting_Operator_1 = cboDBValCFOperator(0).Text
End Property

Public Property Let ConditionalFormatting_Operator_1(ByVal psNewValue As String)
  Dim iLoop As Integer

  If (optLink(SSINTLINKDB_VALUE).value And psNewValue <> vbNullString And chkConditionalFormatting) Then
    For iLoop = 0 To cboDBValCFOperator(0).ListCount - 1
      If cboDBValCFOperator(0).List(iLoop) = psNewValue Then
        cboDBValCFOperator(0).ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

 'cboDBValCFOperator(0).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Value_1() As String
  ConditionalFormatting_Value_1 = txtDBValCFValue(0).Text
End Property

Public Property Let ConditionalFormatting_Value_1(ByVal psNewValue As String)
  If chkConditionalFormatting Then txtDBValCFValue(0).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Style_1() As String
  ConditionalFormatting_Style_1 = cboDBValCFStyle(0).Text
End Property

Public Property Let ConditionalFormatting_Style_1(ByVal psNewValue As String)
  Dim iLoop As Integer

  If (optLink(SSINTLINKDB_VALUE).value And psNewValue <> vbNullString And chkConditionalFormatting) Then
    For iLoop = 0 To cboDBValCFStyle(0).ListCount - 1
      If cboDBValCFStyle(0).List(iLoop) = psNewValue Then
        cboDBValCFStyle(0).ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

  'cboDBValCFStyle(0).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Colour_1() As String
  ConditionalFormatting_Colour_1 = UI.SysColorToHex(txtDBValCFColour(0).BackColor)
End Property

Public Property Let ConditionalFormatting_Colour_1(ByVal psNewValue As String)
  If chkConditionalFormatting Then
    txtDBValCFColour(0).BackColor = IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground)
    txtDBValCFColour(0).ForeColor = UI.GetInverseColor(IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground))
  End If
End Property

Public Property Get ConditionalFormatting_Operator_2() As String
  ConditionalFormatting_Operator_2 = cboDBValCFOperator(1).Text
End Property

Public Property Let ConditionalFormatting_Operator_2(ByVal psNewValue As String)
  Dim iLoop As Integer

  If (optLink(SSINTLINKDB_VALUE).value And psNewValue <> vbNullString And chkConditionalFormatting) Then
    For iLoop = 0 To cboDBValCFOperator(1).ListCount - 1
      If cboDBValCFOperator(1).List(iLoop) = psNewValue Then
        cboDBValCFOperator(1).ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

'  cboDBValCFOperator(1).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Value_2() As String
  ConditionalFormatting_Value_2 = txtDBValCFValue(1).Text
End Property

Public Property Let ConditionalFormatting_Value_2(ByVal psNewValue As String)
  If chkConditionalFormatting Then txtDBValCFValue(1).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Style_2() As String
  ConditionalFormatting_Style_2 = cboDBValCFStyle(1).Text
End Property

Public Property Let ConditionalFormatting_Style_2(ByVal psNewValue As String)
  Dim iLoop As Integer

  If (optLink(SSINTLINKDB_VALUE).value And psNewValue <> vbNullString And chkConditionalFormatting) Then
    For iLoop = 0 To cboDBValCFStyle(1).ListCount - 1
      If cboDBValCFStyle(1).List(iLoop) = psNewValue Then
        cboDBValCFStyle(1).ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

'  cboDBValCFStyle(1).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Colour_2() As String
  ConditionalFormatting_Colour_2 = UI.SysColorToHex(txtDBValCFColour(1).BackColor)
End Property

Public Property Let ConditionalFormatting_Colour_2(ByVal psNewValue As String)
If chkConditionalFormatting Then
  txtDBValCFColour(1).BackColor = IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground)
  txtDBValCFColour(1).ForeColor = UI.GetInverseColor(IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground))
End If
End Property

Public Property Get ConditionalFormatting_Operator_3() As String
  ConditionalFormatting_Operator_3 = cboDBValCFOperator(2).Text
End Property

Public Property Let ConditionalFormatting_Operator_3(ByVal psNewValue As String)
  Dim iLoop As Integer

  If (optLink(SSINTLINKDB_VALUE).value And psNewValue <> vbNullString And chkConditionalFormatting) Then
    For iLoop = 0 To cboDBValCFOperator(2).ListCount - 1
      If cboDBValCFOperator(2).List(iLoop) = psNewValue Then
        cboDBValCFOperator(2).ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

'  cboDBValCFOperator(2).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Value_3() As String
  ConditionalFormatting_Value_3 = txtDBValCFValue(2).Text
End Property

Public Property Let ConditionalFormatting_Value_3(ByVal psNewValue As String)
  If chkConditionalFormatting Then txtDBValCFValue(2).Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get ConditionalFormatting_Style_3() As String
  ConditionalFormatting_Style_3 = cboDBValCFStyle(2).Text
End Property

Public Property Let ConditionalFormatting_Style_3(ByVal psNewValue As String)
  Dim iLoop As Integer

  If (optLink(SSINTLINKDB_VALUE).value And psNewValue <> vbNullString And chkConditionalFormatting) Then
    For iLoop = 0 To cboDBValCFStyle(2).ListCount - 1
      If cboDBValCFStyle(2).List(iLoop) = psNewValue Then
        cboDBValCFStyle(2).ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
End Property

Public Property Get ConditionalFormatting_Colour_3() As String
  ConditionalFormatting_Colour_3 = UI.SysColorToHex(txtDBValCFColour(2).BackColor)
End Property

Public Property Let ConditionalFormatting_Colour_3(ByVal psNewValue As String)
If chkConditionalFormatting Then
  txtDBValCFColour(2).BackColor = IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground)
  txtDBValCFColour(2).ForeColor = UI.GetInverseColor(IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground))
End If
End Property

Public Property Get SeparatorBorderColour() As String
  SeparatorBorderColour = UI.SysColorToHex(txtSeparatorColour.BackColor)
End Property

Public Property Let SeparatorBorderColour(ByVal psNewValue As String)
  txtSeparatorColour.BackColor = IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground)
  txtSeparatorColour.ForeColor = UI.GetInverseColor(IIf(psNewValue <> vbNullString, UI.HexToSysColor(psNewValue), vbWindowBackground))
End Property

Public Property Get Chart_TableID_2() As Long
  Chart_TableID_2 = mlngChart_TableID_2
End Property

Public Property Let Chart_TableID_2(ByVal plngNewValue As Long)
  mlngChart_TableID_2 = plngNewValue
End Property

Public Property Get Chart_ColumnID_2() As Long
  Chart_ColumnID_2 = mlngChart_ColumnID_2
End Property

Public Property Let Chart_ColumnID_2(ByVal plngNewValue As Long)
  mlngChart_ColumnID_2 = plngNewValue
End Property

Public Property Get Chart_TableID_3() As Long
  Chart_TableID_3 = mlngChart_TableID_3
End Property

Public Property Let Chart_TableID_3(ByVal plngNewValue As Long)
  mlngChart_TableID_3 = plngNewValue
End Property

Public Property Get Chart_ColumnID_3() As Long
  Chart_ColumnID_3 = mlngChart_ColumnID_3
End Property

Public Property Let Chart_ColumnID_3(ByVal plngNewValue As Long)
  mlngChart_ColumnID_3 = plngNewValue
End Property

Public Property Get Chart_SortOrderID() As Long
  Chart_SortOrderID = mlngChart_SortOrderID
End Property

Public Property Let Chart_SortOrderID(ByVal plngNewValue As Long)
  mlngChart_SortOrderID = plngNewValue
End Property
      
Public Property Get Chart_SortDirection() As Integer
  Chart_SortDirection = miChart_SortDirection
End Property

Public Property Let Chart_SortDirection(ByVal piNewValue As Integer)
  miChart_SortDirection = piNewValue
End Property
      
Public Property Get InitialDisplayMode() As Integer
  InitialDisplayMode = miInitialDisplayMode
End Property

Public Property Let InitialDisplayMode(ByVal piNewValue As Integer)
  miInitialDisplayMode = piNewValue
End Property

Public Property Get Chart_ColourID() As Long
  Chart_ColourID = mlngChart_ColourID
End Property

Public Property Let Chart_ColourID(ByVal plngNewValue As Long)
  mlngChart_ColourID = plngNewValue
End Property

Public Property Get Chart_Description() As String

  If ChartTableID = 0 Then
    Chart_Description = ""
  ElseIf Chart_TableID_2 = 0 And Chart_TableID_3 = 0 Then
    Chart_Description = "1-T"
  ElseIf Chart_TableID_2 > 0 Then  ' table 2 is actually the Z-Axis
    Chart_Description = "3-T"
  Else
    Chart_Description = "2-T"
  End If
    
  Chart_Description = Chart_Description & IIf(ChartTableID > 0, ", " & Replace(GetTableName(ChartTableID), "_", " "), "")
  Chart_Description = Chart_Description & IIf(Chart_TableID_2 > 0, ", " & Replace(GetTableName(Chart_TableID_2), "_", " "), "")
  Chart_Description = Chart_Description & IIf(Chart_TableID_3 > 0, ", " & Replace(GetTableName(Chart_TableID_3), "_", " "), "")
  
End Property

Public Property Get Chart_Utility_Description() As String
    
  If Len(UtilityType) > 0 And ChartColumnID <> 0 Then
    Chart_Utility_Description = cboHRProUtilityType.Text & " - " & cboHRProUtility.Text
  Else
    Chart_Utility_Description = ""
  End If
    
End Property



