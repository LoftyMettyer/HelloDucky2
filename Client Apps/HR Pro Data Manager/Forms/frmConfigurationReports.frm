VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmConfigurationReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Configuration"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigurationReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   8010
      TabIndex        =   63
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   9285
      TabIndex        =   64
      Top             =   5760
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Definition"
      TabPicture(0)   =   "frmConfigurationReports.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDateRangeDef"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDateRangeRun"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDisplayOptions"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraSortOrder"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraAbsence"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraColumns"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraRecordSelection"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "O&utput"
      TabPicture(1)   =   "frmConfigurationReports.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOutputFormat"
      Tab(1).Control(1)=   "fraOutputDestination"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   5020
         Left            =   -72240
         TabIndex        =   49
         Top             =   360
         Width           =   7470
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3900
            TabIndex        =   62
            Tag             =   "0"
            Top             =   3460
            Width           =   3240
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1760
            Width           =   2940
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1240
            Width           =   3240
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   2160
            Width           =   3240
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3900
            TabIndex        =   61
            Top             =   3060
            Width           =   3240
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   2660
            Width           =   2940
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   56
            Top             =   1760
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   60
            Top             =   2660
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   850
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   1300
            Width           =   1740
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   54
            Top             =   1820
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   58
            Top             =   2720
            Width           =   1515
         End
         Begin VB.CheckBox chkPreview 
            Caption         =   "Preview on &screen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   400
            Width           =   3495
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   2430
            TabIndex        =   78
            Top             =   3525
            Width           =   1065
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2430
            TabIndex        =   77
            Top             =   1815
            Width           =   1095
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   2430
            TabIndex        =   76
            Top             =   3120
            Width           =   1305
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   2430
            TabIndex        =   75
            Top             =   2715
            Width           =   1200
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2430
            TabIndex        =   74
            Top             =   2220
            Width           =   1350
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2430
            TabIndex        =   73
            Top             =   1305
            Width           =   1410
         End
      End
      Begin VB.Frame fraOutputFormat 
         Caption         =   "Output Format :"
         Height          =   5020
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   2505
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel P&ivot Table"
            Height          =   195
            Index           =   6
            Left            =   200
            TabIndex        =   48
            Top             =   2800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Char&t"
            Height          =   195
            Index           =   5
            Left            =   200
            TabIndex        =   47
            Top             =   2400
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   46
            Top             =   2000
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   45
            Top             =   1600
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   44
            Top             =   1200
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CS&V File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   43
            Top             =   800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "D&ata Only"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   42
            Top             =   400
            Value           =   -1  'True
            Width           =   1900
         End
      End
      Begin VB.Frame fraRecordSelection 
         Caption         =   "Record Selection :"
         Height          =   3130
         Left            =   5235
         TabIndex        =   28
         Top             =   2280
         Width           =   4980
         Begin VB.CheckBox chkMinimumBradfordFactor 
            Caption         =   "&Minimum Bradford Factor"
            Height          =   210
            Left            =   210
            TabIndex        =   39
            Top             =   2490
            Width           =   2520
         End
         Begin VB.CheckBox chkOmitAbsenceBeforeStart 
            Caption         =   "Omit absences starting &before the report start date"
            Height          =   300
            Left            =   210
            TabIndex        =   37
            Top             =   1800
            Width           =   4680
         End
         Begin VB.CheckBox chkOmitAbsenceAfterEnd 
            Caption         =   "Omit absences ending af&ter the report end date"
            Height          =   315
            Left            =   210
            TabIndex        =   38
            Top             =   2115
            Width           =   4515
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4380
            TabIndex        =   35
            Top             =   1040
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4380
            TabIndex        =   32
            Top             =   640
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   200
            TabIndex        =   29
            Top             =   300
            Value           =   -1  'True
            Width           =   3660
         End
         Begin VB.OptionButton optPicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   200
            TabIndex        =   30
            Top             =   700
            Width           =   1065
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   200
            TabIndex        =   33
            Top             =   1100
            Width           =   840
         End
         Begin VB.TextBox txtPicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   31
            Tag             =   "0"
            Top             =   640
            Width           =   2340
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   34
            Tag             =   "0"
            Top             =   1040
            Width           =   2340
         End
         Begin VB.CheckBox chkPrintFilterHeader 
            Caption         =   "Display filter or pic&klist title in the report header"
            Enabled         =   0   'False
            Height          =   240
            Left            =   210
            TabIndex        =   36
            Tag             =   "PrintFilterHeader"
            Top             =   1500
            Width           =   4440
         End
         Begin COASpinner.COA_Spinner spnMinimumBradfordFactor 
            Height          =   300
            Left            =   2745
            TabIndex        =   40
            Top             =   2430
            Width           =   705
            _ExtentX        =   1244
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
            MaximumValue    =   99999
            Text            =   "0"
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns :"
         Height          =   1890
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   5000
         Begin VB.CheckBox chkNewStarters 
            Caption         =   "&Include new starters within report period"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   4575
         End
         Begin VB.ComboBox cboVerCol 
            Height          =   315
            Left            =   2025
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   765
            Width           =   2820
         End
         Begin VB.ComboBox cboHorCol 
            Height          =   315
            Left            =   2025
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   2820
         End
         Begin VB.Label lblVerCol 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vertical Column :"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   825
            Width           =   1560
         End
         Begin VB.Label lblHorCol 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horizontal Column :"
            Height          =   195
            Left            =   240
            TabIndex        =   65
            Top             =   420
            Width           =   1710
         End
      End
      Begin VB.Frame fraAbsence 
         Caption         =   "Absence Types :"
         Height          =   1890
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4995
         Begin VB.ListBox lstTypes 
            Height          =   1410
            Left            =   195
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   300
            Width           =   4605
         End
      End
      Begin VB.Frame fraSortOrder 
         Caption         =   "Sort Order :"
         Height          =   1200
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   4995
         Begin VB.CheckBox chkSort1Asc 
            Caption         =   "Asc&ending"
            Height          =   210
            Left            =   3720
            TabIndex        =   9
            Top             =   360
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.ComboBox cboGroupBy 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   700
            Width           =   2445
         End
         Begin VB.ComboBox cboOrderBy 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   300
            Width           =   2445
         End
         Begin VB.CheckBox chkSort2Asc 
            Caption         =   "Ascendin&g"
            Height          =   210
            Left            =   3720
            TabIndex        =   11
            Top             =   760
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.Label lblGroupBy 
            Caption         =   "Then By :"
            Height          =   255
            Left            =   195
            TabIndex        =   68
            Top             =   765
            Width           =   885
         End
         Begin VB.Label lblOrderBy 
            Caption         =   "Order By :"
            Height          =   345
            Left            =   195
            TabIndex        =   67
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame fraDisplayOptions 
         Caption         =   "Display :"
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   4995
         Begin VB.CheckBox chkDisplayDetail 
            Caption         =   "Sho&w Absence Details"
            Height          =   255
            Left            =   195
            TabIndex        =   17
            Top             =   1485
            Width           =   2865
         End
         Begin VB.CheckBox chkSRV 
            Caption         =   "Suppress Repeated Perso&nnel Details"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   300
            Width           =   3780
         End
         Begin VB.CheckBox chkShowWorkings 
            Caption         =   "Show Bradford Factor Formu&la"
            Height          =   195
            Left            =   195
            TabIndex        =   16
            Top             =   1200
            Width           =   3570
         End
         Begin VB.CheckBox chkShowTotals 
            Caption         =   "S&how Duration Totals"
            Height          =   195
            Left            =   195
            TabIndex        =   14
            Top             =   600
            Value           =   1  'Checked
            Width           =   3570
         End
         Begin VB.CheckBox chkShowCount 
            Caption         =   "Show &Instances Count"
            Height          =   195
            Left            =   195
            TabIndex        =   15
            Top             =   900
            Width           =   3015
         End
      End
      Begin VB.Frame fraDateRangeRun 
         Caption         =   "Date Range :"
         Height          =   1890
         Left            =   5235
         TabIndex        =   18
         Top             =   360
         Width           =   4980
         Begin GTMaskDate.GTMaskDate dtDate 
            Height          =   315
            Index           =   0
            Left            =   1635
            TabIndex        =   19
            Top             =   300
            Width           =   1575
            _Version        =   65537
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            NullText        =   "__/__/____"
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            MaskCentury     =   2
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin GTMaskDate.GTMaskDate dtDate 
            Height          =   315
            Index           =   1
            Left            =   1635
            TabIndex        =   20
            Top             =   705
            Width           =   1575
            _Version        =   65537
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            NullText        =   "__/__/____"
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            MaskCentury     =   2
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date :"
            Height          =   195
            Left            =   195
            TabIndex        =   72
            Top             =   765
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date :"
            Height          =   195
            Left            =   195
            TabIndex        =   71
            Top             =   360
            Width           =   1350
         End
      End
      Begin VB.Frame fraDateRangeDef 
         Caption         =   "Date Range :"
         Height          =   1890
         Left            =   5235
         TabIndex        =   21
         Top             =   360
         Width           =   4980
         Begin VB.CommandButton cmdExprDate 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   4380
            TabIndex        =   27
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdExprDate 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   4380
            TabIndex        =   25
            Top             =   1040
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.OptionButton optDate 
            Caption         =   "Default (&12 months to end of last month)"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   22
            Top             =   300
            Value           =   -1  'True
            Width           =   3885
         End
         Begin VB.OptionButton optDate 
            Caption         =   "Cu&stom"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   23
            Top             =   700
            Width           =   1275
         End
         Begin VB.TextBox txtDateExpr 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1995
            Locked          =   -1  'True
            TabIndex        =   24
            Tag             =   "0"
            Top             =   1040
            Width           =   2385
         End
         Begin VB.TextBox txtDateExpr 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1995
            Locked          =   -1  'True
            TabIndex        =   26
            Tag             =   "0"
            Top             =   1440
            Width           =   2385
         End
         Begin VB.Label lblStartDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   510
            TabIndex        =   69
            Top             =   1095
            Width           =   1215
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   510
            TabIndex        =   70
            Top             =   1500
            Width           =   1125
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConfigurationReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjOutputDef As clsOutputDef
Private mrsColumns As Recordset
Private mstrReportType As String
Private mblnRun As Boolean
Private mlngSingleRecord As Long

Private mblnPrintFilterHeader As Boolean
Private mlngVerColID As Long
Private mlngHorColID As Long

Private mlngFormat As Long
Private mblnPreview As Boolean
Private mblnScreen As Boolean
Private mblnPrinter As Boolean
Private mstrPrinterName As String
Private mblnSave As Boolean
Private mlngSaveExisting As Long
Private mstrFileName As String

Private lngAction As ReportOptions
Private mblnForceInitialChanged As Boolean


Private Property Let Changed(pblnChanged As Boolean)
  cmdOK.Enabled = pblnChanged
End Property

Private Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Run(ByVal blnNewValue As Boolean)
  mblnRun = blnNewValue
  fraDateRangeDef.Visible = Not mblnRun
  fraDateRangeRun.Visible = mblnRun
  cmdOK.Caption = IIf(mblnRun, "&Run", "&OK")
End Property

Public Property Let SingleRecord(ByVal lngNewValue As Long)
  mlngSingleRecord = lngNewValue
End Property


Public Sub ShowControls(strReportType As String)

  Dim blnAbsBreakdown As Boolean
  Dim blnBradford As Boolean
  
  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos True, True, True

  Me.Caption = IIf(strReportType = "Stability", "Stability Index", strReportType) & _
               " Report Configuration"
  mstrReportType = Replace(strReportType, " ", "")
   
  blnAbsBreakdown = (strReportType = "Absence Breakdown")
  blnBradford = (strReportType = "Bradford Factor")
  
  fraColumns.Visible = Not (blnAbsBreakdown Or blnBradford)
  fraColumns.Enabled = fraColumns.Visible
  chkNewStarters.Visible = (strReportType = "Turnover")
  
  fraAbsence.Visible = blnAbsBreakdown Or blnBradford
  fraAbsence.Enabled = fraAbsence.Visible
  fraSortOrder.Visible = blnBradford
  fraSortOrder.Enabled = fraSortOrder.Visible
  fraDisplayOptions.Visible = blnBradford
  fraDisplayOptions.Enabled = fraDisplayOptions.Visible
  
  chkOmitAbsenceAfterEnd.Visible = blnBradford
  chkOmitAbsenceAfterEnd.Enabled = blnBradford
  chkOmitAbsenceBeforeStart.Visible = blnBradford
  chkOmitAbsenceBeforeStart.Enabled = blnBradford
  chkMinimumBradfordFactor.Visible = blnBradford
  chkMinimumBradfordFactor.Enabled = blnBradford
  'optOutputFormat(fmtExcelPivotTable).Visible = Not blnBradford
  'optOutputFormat(fmtExcelPivotTable).Enabled = Not blnBradford
  
  'MH20040705 - Remove CSV and Pivot for Bradford...
  mobjOutputDef.ShowFormats True, Not blnBradford, True, True, True, True, Not blnBradford

  If Not blnBradford Then
    Me.Height = Me.Height - 840
    SSTab1.Height = SSTab1.Height - 840
    fraRecordSelection.Height = fraRecordSelection.Height - 840
    fraOutputFormat.Height = fraOutputFormat.Height - 840
    fraOutputDestination.Height = fraOutputDestination.Height - 840
    cmdOK.Top = cmdOK.Top - 840
    cmdCancel.Top = cmdCancel.Top - 840
  End If
  
  
  Select Case strReportType
  'Absence Breakdown
  Case "Absence Breakdown"
    fraAbsence.Height = 4230
    lstTypes.Height = 3635

    LoadAbsenceTypes
    Me.HelpContextID = 1004

  'Bradford Factor
  Case "Bradford Factor"
    fraAbsence.Height = 1890
    lstTypes.Height = 1410

    LoadAbsenceTypes
    LoadBradford
    Me.HelpContextID = 1011
  
  
  'Stability Index
  Case "Stability"
    
    fraColumns.Height = 4230
    
    If GetColumns Then
      PopulateCombo cboHorCol, True
      PopulateCombo cboVerCol, False
      SetComboItem cboHorCol, GetSystemSetting(strReportType, "HorColID", 0)
      SetComboItem cboVerCol, GetSystemSetting(strReportType, "VerColID", 0)
    End If
    Me.HelpContextID = 1018
  
  
  'Turnover Report
  Case "Turnover"
    
    fraColumns.Height = 4230
    
    If GetColumns Then
      PopulateCombo cboHorCol, True
      PopulateCombo cboVerCol, False
      SetComboItem cboHorCol, GetSystemSetting(strReportType, "HorColID", 0)
      SetComboItem cboVerCol, GetSystemSetting(strReportType, "VerColID", 0)
      chkNewStarters.Value = IIf(GetSystemSetting(mstrReportType, "IncludeNewStarters", 0) = 1, vbChecked, vbUnchecked)
    End If
    Me.HelpContextID = 1019

  End Select


  LoadDates
  LoadPicklistFilter
  LoadOutputOptions

  chkPrintFilterHeader.Value = IIf(GetSystemSetting(mstrReportType, "PrintFilterHeader", 0) = 1, vbChecked, vbUnchecked)
    
  EnableControls
    
  If Not mblnRun Then
    Changed = mblnForceInitialChanged
    mblnForceInitialChanged = False
  End If
  
End Sub

Private Sub cboGroupBy_Click()
  
  If Not (mlngSingleRecord > 0) Then
    If cboGroupBy.ItemData(cboGroupBy.ListIndex) > 0 Then
      chkSort2Asc.Enabled = True
    Else
      chkSort2Asc.Value = vbUnchecked
      chkSort2Asc.Enabled = False
    End If
  End If
  
  Changed = True
End Sub

Private Sub cboHorCol_Click()
  
  Dim lngHorColID As Long
  Dim lngVerColID As Long

  With cboHorCol
    
    lngHorColID = .ItemData(.ListIndex)
    
    If lngHorColID <> Val(.Tag) Then
    
      lngVerColID = cboVerCol.ItemData(cboVerCol.ListIndex)

      .Tag = lngHorColID

      If GetColumns(lngHorColID) Then
        PopulateCombo cboVerCol, False

        'Restore Vertical column, if still selectable...
        SetComboItem cboVerCol, lngVerColID
      End If
    
    End If
  
  End With

  Changed = True
  
End Sub

Private Sub cboOrderBy_Click()
  If Not (mlngSingleRecord > 0) Then
    If cboOrderBy.ItemData(cboOrderBy.ListIndex) > 0 Then
      chkSort1Asc.Enabled = True
    Else
      chkSort1Asc.Value = vbUnchecked
      chkSort1Asc.Enabled = False
    End If
  End If
  Changed = True
End Sub

Private Sub cboPrinterName_Click()
  Changed = True
End Sub



Private Sub cboSaveExisting_Click()
  Changed = True
End Sub

Private Sub cboVerCol_Click()
Changed = True
End Sub

Private Sub chkDisplayDetail_Click()
  If chkDisplayDetail.Value = vbUnchecked Then
    chkSRV.Value = vbUnchecked
    chkSRV.Enabled = False
  Else
    chkSRV.Enabled = True
  End If
  Changed = True
End Sub

Private Sub chkMinimumBradfordFactor_Click()

  spnMinimumBradfordFactor.Enabled = IIf(chkMinimumBradfordFactor.Value = vbChecked, True, False)
  'NHRD04062003 Fault 5760
  spnMinimumBradfordFactor.BackColor = IIf(spnMinimumBradfordFactor.Enabled, vbWindowBackground, vbButtonFace)
  'NHRD17072003 Fault 6251
  If spnMinimumBradfordFactor.Enabled = False Then spnMinimumBradfordFactor.Value = 0
  
  Changed = True

End Sub

Private Sub chkNewStarters_Click()
Changed = True
End Sub

Private Sub chkOmitAbsenceAfterEnd_Click()
Changed = True
End Sub

Private Sub chkOmitAbsenceBeforeStart_Click()
Changed = True
End Sub

Private Sub chkPreview_Click()
Changed = True
End Sub

Private Sub chkPrintFilterHeader_Click()
Changed = True
End Sub

Private Sub chkShowCount_Click()
Changed = True
End Sub

Private Sub chkShowTotals_Click()
Changed = True
End Sub

Private Sub chkShowWorkings_Click()
Changed = True
End Sub

Private Sub chkSort1Asc_Click()
Changed = True
End Sub

Private Sub chkSort2Asc_Click()
Changed = True
End Sub

Private Sub chkSRV_Click()
Changed = True
End Sub

Private Sub cmdCancel_Click()
  
  lngAction = rptCancel
  Unload Me

End Sub

Private Sub cmdExprDate_Click(Index As Integer)
  
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  fOK = True
  
  Set objExpression = New clsExprExpression
  With objExpression
    
    fOK = .Initialise(0, Val(txtDateExpr(Index).Tag), giEXPR_RECORDINDEPENDANTCALC, giEXPRVALUE_DATE)
    If fOK Then

      Do
        .SelectExpression True
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        .ValidateExpression True, True
        
        fOK = (.ReturnType = giEXPRVALUE_DATE)
        
        If Not fOK Then
          COAMsgBox "This calculation does not return a date value.", vbExclamation, Me.Caption
        End If
      Loop While Not fOK

      If fOK Then
        'NHRD14022007 Fault 10744
        If (txtDateExpr(Index).Text = IIf((.Name = vbNullString), "<None>", .Name) And txtDateExpr(Index).Tag = .ExpressionID) = False Then
          txtDateExpr(Index).Text = IIf((.Name = vbNullString), "<None>", .Name)
          txtDateExpr(Index).Tag = .ExpressionID
          Changed = True
        End If
      End If
    End If
  End With
  
  Set objExpression = Nothing

End Sub

Private Sub cmdOK_Click()

  lngAction = IIf(mblnRun, rptRun, rptOK)

  If ValidDefinition Then
    If mblnRun Then
      Me.Hide
      RunDefinition
    Else
      SaveDefinition
    
      Unload Me
    End If
  End If

End Sub

Private Sub dtDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    dtDate(Index).DateValue = Date
  End If
End Sub

Private Sub dtDate_LostFocus(Index As Integer)
  ValidateGTMaskDate dtDate(Index)
End Sub

Private Sub Form_Activate()
  'NHRD04062003 Fault 5760
  spnMinimumBradfordFactor.Enabled = IIf(chkMinimumBradfordFactor.Value = vbChecked, True, False)
  spnMinimumBradfordFactor.BackColor = IIf(spnMinimumBradfordFactor.Enabled, vbWindowBackground, vbButtonFace)
  lngAction = rptCancel
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
  
  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos True, True, True

  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl dtDate(0)
  UI.FormatGTDateControl dtDate(1)
  
  SSTab1.Tab = 0
  SSTab1_Click 0
  Changed = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim pintAnswer As Integer
  
  If (Changed) Then
    
    If (lngAction = rptCancel) And (Not mblnRun) Then
      pintAnswer = COAMsgBox("You have changed the report configuration. Save changes ?", vbQuestion + vbYesNoCancel, "Report Configuration")
        
      If pintAnswer = vbYes Then
        If Not ValidDefinition Then
          Cancel = 1
        Else
          cmdOK_Click
          Cancel = 0
        End If
        Exit Sub
      ElseIf pintAnswer = vbCancel Then
        Cancel = 1
        Exit Sub
      End If
    End If
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mobjOutputDef = Nothing
End Sub

Private Sub lstTypes_ItemCheck(Item As Integer)
  Changed = True
End Sub

Private Sub optDate_Click(Index As Integer)

  Dim blnCustomDate As Boolean
  Dim lngCount As Long

  blnCustomDate = (Index = 1)

  lblStartDate.Enabled = blnCustomDate
  lblEndDate.Enabled = blnCustomDate

  For lngCount = 0 To 1
    cmdExprDate(lngCount).Enabled = blnCustomDate
    'txtDateExpr(lngCount).Enabled = blnCustomDate
    If blnCustomDate = False Then
      txtDateExpr(lngCount).Text = vbNullString
      txtDateExpr(lngCount).Tag = 0
    ElseIf optDate(1).Value Then
      txtDateExpr(lngCount).Text = "<None>"
      txtDateExpr(lngCount).Tag = 0
    End If
  Next

  Changed = True

End Sub

Private Sub optOutputFormat_Click(Index As Integer)
  mobjOutputDef.FormatClick Index
  Changed = True
End Sub

Private Sub chkDestination_Click(Index As Integer)
  mobjOutputDef.DestinationClick Index
  Changed = True
End Sub


Private Sub LoadPicklistFilter()

  Dim strRecSelStatus As String
  Dim pstrType As String
  Dim plngID As Long

  If mlngSingleRecord > 0 Then
    optAllRecords.Caption = "Current Record"
    optAllRecords.Value = True
    optAllRecords.Enabled = False
    Me.optFilter.Enabled = False
    Me.optPicklist.Enabled = False
    Me.chkPrintFilterHeader.Enabled = False
    chkPrintFilterHeader.Value = vbUnchecked
  
  Else
    pstrType = GetSystemSetting(mstrReportType, "Type", "A")
    plngID = GetSystemSetting(mstrReportType, "ID", 0)
    
    Select Case pstrType
    Case "A"
      optAllRecords.Value = True
    Case "F"
      optFilter.Value = True
      strRecSelStatus = IsFilterValid(plngID)
      If strRecSelStatus <> vbNullString Then
        COAMsgBox strRecSelStatus & vbCrLf & "It has been removed from the definition.", vbExclamation, Me.Caption
        txtFilter.Text = "<None>"
        txtFilter.Tag = 0
        mblnForceInitialChanged = True
      Else
        txtFilter.Text = datGeneral.GetFilterName(plngID)
        txtFilter.Tag = plngID
      End If
    Case "P"
      optPicklist.Value = True
      strRecSelStatus = IsPicklistValid(plngID)
      If strRecSelStatus <> vbNullString Then
        COAMsgBox strRecSelStatus & vbCrLf & "It has been removed from the definition.", vbExclamation, Me.Caption
        txtPicklist.Text = "<None>"
        txtPicklist.Tag = 0
        mblnForceInitialChanged = True
      Else
        txtPicklist.Text = datGeneral.GetPicklistName(plngID)
        txtPicklist.Tag = plngID
      End If
    End Select

  End If

End Sub


Private Sub optAllRecords_Click()
  Call RecordSelectionClick(False, False)
End Sub

Private Sub optPicklist_Click()
  Call RecordSelectionClick(True, False)
End Sub

Private Sub optFilter_Click()
  Call RecordSelectionClick(False, True)
End Sub

Public Sub RecordSelectionClick(blnPicklist As Boolean, blnFilter As Boolean)

  cmdPicklist.Enabled = blnPicklist
  If blnPicklist = False Then
    txtPicklist.Text = vbNullString
    txtPicklist.Tag = vbNullString
  ElseIf txtPicklist.Text = vbNullString Then
    txtPicklist.Text = "<None>"
  End If

  cmdFilter.Enabled = blnFilter
  If blnFilter = False Then
    txtFilter.Text = vbNullString
    txtFilter.Tag = vbNullString
  ElseIf txtFilter.Text = vbNullString Then
    txtFilter.Text = "<None>"
  End If

  chkPrintFilterHeader.Enabled = (blnPicklist Or blnFilter)
  If Not (blnPicklist Or blnFilter) Then
    chkPrintFilterHeader.Value = vbUnchecked
  End If

  Changed = True
  
End Sub

Private Sub cmdPicklist_Click()

  Dim rsTemp As Recordset
  Dim sSQL As String
  Dim lParent As Long
  Dim fExit As Boolean
  Dim frmPick As frmPicklists
  Dim blnEnabled As Boolean
  Dim blnHiddenPicklist As Boolean

  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass

  fExit = False

  With frmDefSel
      
    .TableID = glngPersonnelTableID
    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(txtPicklist.Tag) > 0 Then
      .SelectedID = Val(txtPicklist.Tag)
    End If

    'loop until a picklist has been selected or cancelled
    Do While Not fExit

      If .ShowList(utlPicklist) Then
        .Show vbModal

        Select Case frmDefSel.Action
        Case edtAdd
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(True, False, glngPersonnelTableID) Then
              .Show vbModal
            End If
            frmDefSel.SelectedID = .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          End With

        Case edtEdit
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(False, frmDefSel.FromCopy, glngPersonnelTableID, frmDefSel.SelectedID) Then
              .Show vbModal
            End If
            If frmDefSel.FromCopy And .SelectedID > 0 Then
              frmDefSel.SelectedID = .SelectedID
            End If
            Unload frmPick
            Set frmPick = Nothing
          End With

        'MH20050728 Fault 10232
        Case edtPrint
          Set frmPick = New frmPicklists
          frmPick.PrintDef .TableID, .SelectedID
          Unload frmPick
          Set frmPick = Nothing

        Case edtSelect

          txtPicklist.Text = frmDefSel.SelectedText
          txtPicklist.Tag = frmDefSel.SelectedID
          txtFilter.Text = ""
          txtFilter.Tag = ""
          fExit = True
          Changed = True
          
        Case 0
          If IsPicklistValid(txtPicklist.Tag) <> vbNullString Then
            txtPicklist.Text = "<None>"
            txtPicklist.Tag = 0
          Else
            txtPicklist.Text = datGeneral.GetPicklistName(Val(txtPicklist.Tag))
          End If
          fExit = True
      
        End Select
      End If

    Loop

  End With

  Set frmDefSel = Nothing


Exit Sub

LocalErr:
  COAMsgBox "Error selecting picklist", vbCritical, Me.Caption

End Sub

Private Sub cmdFilter_Click()
  
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression

  Dim rsTemp As Recordset
  Dim sSQL As String
  Dim blnHiddenPicklist As Boolean
  
  On Error GoTo LocalErr
  
  Set objExpression = New clsExprExpression
  With objExpression
    fOK = .Initialise(glngPersonnelTableID, Val(txtFilter.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)

    If fOK Then
      If .SelectExpression(True) Then
        txtFilter.Text = .Name
        txtFilter.Tag = .ExpressionID
        txtPicklist.Text = ""
        txtPicklist.Tag = ""
        Changed = True
      Else
        txtFilter.Text = datGeneral.GetFilterName(Val(txtFilter.Tag))
      End If

    End If
  End With

  Set objExpression = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Error selecting filter", vbCritical, Me.Caption

End Sub


Private Sub LoadAbsenceTypes()

  'Dim objAbsBreakdown As clsAbsenceBreakdown
  Dim rsType As Recordset
  Dim strSQL As String
  Dim strSettingKey As String

  'Set objAbsBreakdown = New clsAbsenceBreakdown
  'objAbsBreakdown.ReportType = mstrReportType

  strSQL = "SELECT * " & _
           "FROM " & gsAbsenceTypeTableName & " " & _
           "ORDER BY " & gsAbsenceTypeTypeColumnName
  Set rsType = datGeneral.GetReadOnlyRecords(strSQL)

  Do Until rsType.EOF

    lstTypes.AddItem rsType.Fields(gsAbsenceTypeTypeColumnName).Value
    'lstTypes.Selected(lstTypes.NewIndex) = objAbsBreakdown.CheckIfAbsenceTypeSelected(rsType.Fields(gsAbsenceTypeTypeColumnName).Value)
    strSettingKey = "Absence Type " & rsType.Fields(gsAbsenceTypeTypeColumnName).Value
    lstTypes.Selected(lstTypes.NewIndex) = _
        (GetSystemSetting(mstrReportType, strSettingKey, vbNullString) = "1")
    
    rsType.MoveNext

  Loop

  lstTypes.ListIndex = 0

  'Set objAbsBreakdown = Nothing
  Set rsType = Nothing

End Sub


Public Sub LoadBradford()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmBradfordIndex.LoadSortCombos()"

  ' Loads the Base combo with all columns for personnel records
  
  Dim sSQL As String
  Dim rsColumns As New Recordset
  Dim datData As New clsDataAccess
  Dim lngOrderBy As Long
  Dim lngGroupBy As Long

  sSQL = "SELECT ColumnID, ColumnName " & _
         "FROM ASRSysColumns " & _
         "WHERE TableID = " & CStr(glngPersonnelTableID) & _
         " AND DataType <> " & Trim(Str(sqlOle)) & _
         " AND DataType <> " & Trim(Str(sqlVarBinary)) & _
         " ORDER BY ColumnName"
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  cboOrderBy.Clear
  cboGroupBy.Clear
  
  'Add blanks
  cboOrderBy.AddItem "<None>"
  cboOrderBy.ItemData(cboOrderBy.NewIndex) = 0
  
  cboGroupBy.AddItem "<None>"
  cboGroupBy.ItemData(cboOrderBy.NewIndex) = 0
  
  ' Populate with columns from personnel records
  rsColumns.MoveFirst
  Do While Not rsColumns.EOF
  
    If Not rsColumns!ColumnName = "ID" Then
      cboOrderBy.AddItem rsColumns!ColumnName
      cboOrderBy.ItemData(cboOrderBy.NewIndex) = rsColumns!ColumnID
      
      cboGroupBy.AddItem rsColumns!ColumnName
      cboGroupBy.ItemData(cboGroupBy.NewIndex) = rsColumns!ColumnID
    End If
    
    rsColumns.MoveNext
  Loop

  rsColumns.Close
  Set rsColumns = Nothing


  ' Set current groups
'  SetComboItem cboOrderBy, Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
  lngOrderBy = GetSystemSetting(mstrReportType, "Order By", 0)
  If lngOrderBy > 0 Then
    SetComboItem cboOrderBy, lngOrderBy
    If cboOrderBy.ListIndex < 0 Then
      cboOrderBy.ListIndex = 0
    End If
  Else
    cboOrderBy.ListIndex = 0
  End If
  chkSort1Asc.Value = IIf(GetSystemSetting(mstrReportType, "Order By Asc", "1") = "1", vbChecked, vbUnchecked)

'  SetComboItem cboGroupBy, Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
  lngGroupBy = GetSystemSetting(mstrReportType, "Group By", 0)
  If lngGroupBy > 0 Then
    SetComboItem cboGroupBy, lngGroupBy
    If cboGroupBy.ListIndex < 0 Then
      cboGroupBy.ListIndex = 0
    End If
  Else
    cboGroupBy.ListIndex = 0
  End If
  chkSort2Asc.Value = IIf(GetSystemSetting(mstrReportType, "Group By Asc", "1") = "1", vbChecked, vbUnchecked)

  chkSRV.Value = IIf(GetSystemSetting(mstrReportType, "SRV", "0") = "1", vbChecked, vbUnchecked)
  chkShowCount.Value = IIf(GetSystemSetting(mstrReportType, "Show Count", "0") = "1", vbChecked, vbUnchecked)
  chkShowTotals.Value = IIf(GetSystemSetting(mstrReportType, "Show Totals", "1") = "1", vbChecked, vbUnchecked)
  chkShowWorkings.Value = IIf(GetSystemSetting(mstrReportType, "Show Workings", "0") = "1", vbChecked, vbUnchecked)
  chkOmitAbsenceBeforeStart.Value = IIf(GetSystemSetting(mstrReportType, "Omit Before", "0") = "1", vbChecked, vbUnchecked)
  chkOmitAbsenceAfterEnd.Value = IIf(GetSystemSetting(mstrReportType, "Omit After", "0") = "1", vbChecked, vbUnchecked)
  
  chkMinimumBradfordFactor.Value = IIf(GetSystemSetting(mstrReportType, "Minimum Bradford Factor", "0") = "1", vbChecked, vbUnchecked)
  spnMinimumBradfordFactor.Value = GetSystemSetting(mstrReportType, "Minimum Bradford Factor Amount", "0")
  chkDisplayDetail.Value = IIf(GetSystemSetting(mstrReportType, "Display Absence Details", "1") = "1", vbChecked, vbUnchecked)

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub

ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub LoadDates()

  Dim blnCustom As Boolean
  Dim strRecSelStatus As String
  Dim lngID As Long
  Dim lngCount As Long
  Dim lngStartDateExprID As Long
  Dim lngEndDateExprID As Long
  
  blnCustom = (GetSystemSetting(mstrReportType, "Custom Dates", "0") = "1")

  If Not mblnRun Then
    If blnCustom Then
      optDate(1).Value = True
  
      For lngCount = 0 To 1
        
        lngID = GetSystemSetting(mstrReportType, IIf(lngCount = 0, "Start Date", "End Date"), "0")
        
        strRecSelStatus = IsCalcValid(lngID)
        If strRecSelStatus <> vbNullString Then
          COAMsgBox strRecSelStatus & vbCrLf & "It has been removed from the definition.", vbExclamation, Me.Caption
          txtDateExpr(lngCount).Text = "<None>"
          txtDateExpr(lngCount).Tag = 0
          mblnForceInitialChanged = True
        Else
          txtDateExpr(lngCount).Text = datGeneral.GetFilterName(lngID)
          txtDateExpr(lngCount).Tag = lngID
        End If
  
      Next
    
    End If

  Else
    If blnCustom Then
      lngStartDateExprID = GetSystemSetting(mstrReportType, "Start Date", 0)
      strRecSelStatus = IsCalcValid(lngStartDateExprID)
      If strRecSelStatus <> vbNullString Then
        COAMsgBox strRecSelStatus & vbCrLf & "It has been removed from the definition.", vbExclamation, Me.Caption
        mblnForceInitialChanged = True
      Else
        dtDate(0).DateValue = datGeneral.GetValueForRecordIndependantCalc(lngStartDateExprID)
      End If

      lngEndDateExprID = GetSystemSetting(mstrReportType, "End Date", 0)
      strRecSelStatus = IsCalcValid(lngEndDateExprID)
      If strRecSelStatus <> vbNullString Then
        COAMsgBox strRecSelStatus & vbCrLf & "It has been removed from the definition.", vbExclamation, Me.Caption
        mblnForceInitialChanged = True
      Else
        dtDate(1).DateValue = datGeneral.GetValueForRecordIndependantCalc(lngEndDateExprID)
      End If
      
    Else
      dtDate(1).DateValue = DateAdd("d", Day(Date) * -1, Date)
      dtDate(0).DateValue = DateAdd("d", 1, DateAdd("yyyy", -1, dtDate(1).DateValue))
    End If
  
  End If

End Sub


Private Sub LoadOutputOptions()

  Dim lngFormat As Long
  Dim blnPreview As Boolean
  Dim blnScreen As Boolean
  Dim blnPrinter As Boolean
  Dim strPrinterName As String
  Dim blnSave As Boolean
  Dim lngSaveExisting As Long
  Dim blnEmail As Boolean
  Dim lngEmailAddr As Long
  Dim strEmailSubject As String
  Dim strEmailAttachAs As String
  Dim strFileName As String
  'Dim lngFileFormat As Long
  'Dim lngEmailFileFormat As Long

  lngFormat = GetSystemSetting(mstrReportType, "Format", 0)
  blnPreview = GetSystemSetting(mstrReportType, "Preview", 0)
  blnScreen = GetSystemSetting(mstrReportType, "Screen", 1)
  blnPrinter = GetSystemSetting(mstrReportType, "Printer", 0)
  strPrinterName = GetSystemSetting(mstrReportType, "PrinterName", vbNullString)
  blnSave = GetSystemSetting(mstrReportType, "Save", 0)
  lngSaveExisting = GetSystemSetting(mstrReportType, "SaveExisting", -1)
  strFileName = GetSystemSetting(mstrReportType, "FileName", vbNullString)
  'lngFileFormat = GetSystemSetting(mstrReportType, "SaveFormat", vbNullString)
  
  blnEmail = GetSystemSetting(mstrReportType, "Email", "0")
  If blnEmail Then
    lngEmailAddr = GetSystemSetting(mstrReportType, "EmailAddr", 0)
    strEmailSubject = GetSystemSetting(mstrReportType, "EmailSubject", vbNullString)
    strEmailAttachAs = GetSystemSetting(mstrReportType, "EmailAttachAs", vbNullString)
    'lngEmailFileFormat = GetSystemSetting(mstrReportType, "EmailFileFormat", vbNullString)
  Else
    lngEmailAddr = 0
    strEmailSubject = vbNullString
    strEmailAttachAs = vbNullString
    'lngEmailFileFormat = 0
  End If


  optOutputFormat(lngFormat).Value = True
  chkPreview.Value = IIf(blnPreview, vbChecked, vbUnchecked)
  chkDestination(desScreen).Value = IIf(blnScreen, vbChecked, vbUnchecked)

  chkDestination(desPrinter).Value = IIf(blnPrinter, vbChecked, vbUnchecked)
  If strPrinterName <> vbNullString Then
    SetComboText cboPrinterName, strPrinterName
    If cboPrinterName.Text <> strPrinterName Then
      cboPrinterName.AddItem strPrinterName
      cboPrinterName.ListIndex = cboPrinterName.NewIndex
      COAMsgBox "This definition is set to output to printer " & strPrinterName & _
             " which is not set up on your PC.", vbInformation, Me.Caption
    End If
  End If

  chkDestination(desSave).Value = IIf(blnSave, vbChecked, vbUnchecked)
  SetComboItem cboSaveExisting, lngSaveExisting

  chkDestination(desEmail).Value = IIf(blnEmail, vbChecked, vbUnchecked)
  'SetComboItem cboEmailAddr, lngEmailAddr
  If blnEmail Then
    txtEmailGroup.Text = datGeneral.GetEmailGroupName(lngEmailAddr)
    txtEmailGroup.Tag = lngEmailAddr
    txtEmailSubject.Text = strEmailSubject
    txtEmailAttachAs.Text = strEmailAttachAs
    'txtEmailAttachAs.Tag = lngEmailFileFormat
  End If
  txtFilename.Text = strFileName
  'txtFilename.Tag = lngFileFormat

End Sub


Private Function GetColumns(Optional lngExcludeColumnID As Long) As Boolean

  Dim strSQL As String


  strSQL = "SELECT columnName, columnID FROM ASRSysColumns" & _
           " WHERE tableID = " & CStr(glngPersonnelTableID) & _
           " AND columnType <> " & Trim(Str(colSystem)) & _
           " AND columnType <> " & Trim(Str(colLink)) & _
           " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
           " AND dataType <> " & Trim(Str(sqlOle))
  
  If lngExcludeColumnID > 0 Then
    strSQL = strSQL & _
           " AND columnID <> " & CStr(lngExcludeColumnID)
  End If
  
  Set mrsColumns = datGeneral.GetReadOnlyRecords(strSQL)
  GetColumns = (Not mrsColumns.BOF Or Not mrsColumns.EOF)

End Function


Private Sub PopulateCombo(cboTemp As ComboBox, blnAllowNone)

  With cboTemp
    .Clear
    
    If blnAllowNone Then
      .AddItem "<None>"   'Horizontal Column is Optional
      .ItemData(.NewIndex) = 0
    End If

    mrsColumns.MoveFirst
    Do While Not mrsColumns.EOF
      .AddItem Replace(mrsColumns.Fields("ColumnName").Value, "_", " ")
      .ItemData(.NewIndex) = mrsColumns.Fields("ColumnID").Value
      mrsColumns.MoveNext
    Loop

    If .ListCount > 0 Then
      .ListIndex = 0
    End If
  
  End With

End Sub


Private Function ValidDefinition() As Boolean

  ValidDefinition = False

  If mblnRun Then
    If IsNull(dtDate(0).DateValue) Then
      COAMsgBox "You must enter a start date.", vbExclamation, Me.Caption
      Exit Function
    ElseIf Not ValidateGTMaskDate(dtDate(0)) Then
      Exit Function
    End If
  
    If IsNull(dtDate(1).DateValue) Then
      'AE20071005 Fault #9959
      COAMsgBox "You must enter an end date.", vbExclamation, Me.Caption
      Exit Function
    ElseIf Not ValidateGTMaskDate(dtDate(1)) Then
      Exit Function
    End If
  
    If DateDiff("d", dtDate(0).DateValue, dtDate(1).DateValue) < 0 Then
      COAMsgBox "The report end date is before the report start date.", vbExclamation, Me.Caption
      Exit Function
    End If
  
  Else
    If optDate(1).Value = True Then
      If Val(txtDateExpr(0).Tag) = 0 Then
        COAMsgBox "You must select a Start Date calculation.", vbExclamation, Me.Caption
        Exit Function
      ElseIf Val(txtDateExpr(1).Tag) = 0 Then
        COAMsgBox "You must select an End Date calculation.", vbExclamation, Me.Caption
        Exit Function
      End If
    End If
  
  End If
  
  ' Check for valid personnel record selection criteria
  If optPicklist And Val(txtPicklist.Tag) = 0 Then
    COAMsgBox "You must select a picklist.", vbExclamation, Me.Caption
    Exit Function
  ElseIf optFilter And Val(txtFilter.Tag) = 0 Then
    COAMsgBox "You must select a filter.", vbExclamation, Me.Caption
    Exit Function
  End If
  
  
  If mstrReportType = "AbsenceBreakdown" Or mstrReportType = "BradfordFactor" Then
    ' Check at least 1 absence type has been selected
    If lstTypes.Visible = True Then
      If lstTypes.SelCount = 0 Then
        COAMsgBox "You must have at least 1 absence type selected.", vbExclamation + vbOKOnly, Me.Caption
        Exit Function
      End If
    End If
  End If


  If Not mobjOutputDef.ValidDestination Then
    SSTab1.Tab = 1
    Exit Function
  End If

  ValidDefinition = True

End Function


Private Function SaveDefinition() As Boolean

  Dim lngCount As Long
  
  
  If mstrReportType = "AbsenceBreakdown" Or mstrReportType = "BradfordFactor" Then
    For lngCount = 0 To lstTypes.ListCount - 1
      SaveSystemSetting mstrReportType, "Absence Type " & lstTypes.List(lngCount), IIf(lstTypes.Selected(lngCount), "1", "0")
    Next
  End If

  SaveSystemSetting mstrReportType, "Custom Dates", IIf(optDate(1).Value = True, "1", "0")
  SaveSystemSetting mstrReportType, "Start Date", txtDateExpr(0).Tag
  SaveSystemSetting mstrReportType, "End Date", txtDateExpr(1).Tag

  If optAllRecords.Value Then
    SaveSystemSetting mstrReportType, "Type", "A"
    SaveSystemSetting mstrReportType, "ID", 0
  ElseIf optPicklist.Value Then
    SaveSystemSetting mstrReportType, "Type", "P"
    SaveSystemSetting mstrReportType, "ID", txtPicklist.Tag
  Else
    SaveSystemSetting mstrReportType, "Type", "F"
    SaveSystemSetting mstrReportType, "ID", txtFilter.Tag
  End If
  SaveSystemSetting mstrReportType, "PrintFilterHeader", IIf(chkPrintFilterHeader.Value = vbChecked, "1", "0")


  Select Case mstrReportType
  Case "BradfordFactor"
    SaveSystemSetting mstrReportType, "Order By", cboOrderBy.ItemData(cboOrderBy.ListIndex)
    SaveSystemSetting mstrReportType, "Order By Asc", IIf(chkSort1Asc.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Group By", cboGroupBy.ItemData(cboGroupBy.ListIndex)
    SaveSystemSetting mstrReportType, "Group By Asc", IIf(chkSort2Asc.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "SRV", IIf(chkSRV.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Show Count", IIf(chkShowCount.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Show Totals", IIf(chkShowTotals.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Show Workings", IIf(chkShowWorkings.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Omit Before", IIf(chkOmitAbsenceBeforeStart.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Omit After", IIf(chkOmitAbsenceAfterEnd.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Minimum Bradford Factor", IIf(chkMinimumBradfordFactor.Value = vbChecked, "1", "0")
    SaveSystemSetting mstrReportType, "Minimum Bradford Factor Amount", spnMinimumBradfordFactor.Value
    SaveSystemSetting mstrReportType, "Display Absence Details", IIf(chkDisplayDetail.Value = vbChecked, "1", "0")

  Case "Stability", "Turnover"
    SaveSystemSetting mstrReportType, "HorColID", cboHorCol.ItemData(cboHorCol.ListIndex)
    SaveSystemSetting mstrReportType, "VerColID", cboVerCol.ItemData(cboVerCol.ListIndex)

    If mstrReportType = "Turnover" Then
      SaveSystemSetting mstrReportType, "IncludeNewStarters", IIf(chkNewStarters.Value = vbChecked, "1", "0")
    End If

  End Select


  SaveSystemSetting mstrReportType, "Format", mobjOutputDef.GetSelectedFormatIndex
  SaveSystemSetting mstrReportType, "Preview", IIf(chkPreview.Value = vbChecked, "1", "0")
  SaveSystemSetting mstrReportType, "Screen", IIf(chkDestination(desScreen).Value = vbChecked, "1", "0")
  SaveSystemSetting mstrReportType, "Printer", IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0")
  SaveSystemSetting mstrReportType, "PrinterName", cboPrinterName.Text
  SaveSystemSetting mstrReportType, "Save", IIf(chkDestination(desSave).Value = vbChecked, "1", "0")
  SaveSystemSetting mstrReportType, "SaveExisting", cboSaveExisting.ListIndex
  SaveSystemSetting mstrReportType, "FileName", txtFilename.Text
  'SaveSystemSetting mstrReportType, "SaveFormat", Val(txtFilename.Tag)

  If chkDestination(desEmail).Value = vbChecked Then
    SaveSystemSetting mstrReportType, "Email", 1
    SaveSystemSetting mstrReportType, "EmailAddr", txtEmailGroup.Tag
    SaveSystemSetting mstrReportType, "EmailSubject", txtEmailSubject.Text
    SaveSystemSetting mstrReportType, "EmailAttachAs", txtEmailAttachAs.Text
    'SaveSystemSetting mstrReportType, "EmailFileFormat", txtEmailAttachAs.Tag
  Else
    SaveSystemSetting mstrReportType, "Email", 0
    SaveSystemSetting mstrReportType, "EmailAddr", 0
    SaveSystemSetting mstrReportType, "EmailSubject", vbNullString
    SaveSystemSetting mstrReportType, "EmailAttachAs", vbNullString
    'SaveSystemSetting mstrReportType, "EmailFileFormat", 0
  End If

End Function


Private Sub RunDefinition()

  Dim frmRun As frmCrossTabRun
  Dim pobjBradfordIndex As clsCustomReportsRUN
  Dim fOK As Boolean
  Dim lngCount As Long
  Dim strAbsTypes As String
  Dim strPicklistFilterType As String
  Dim lngPicklistFilterID As Long


  strAbsTypes = vbNullString
  If mstrReportType = "AbsenceBreakdown" Or mstrReportType = "BradfordFactor" Then
    For lngCount = 0 To lstTypes.ListCount - 1
      If lstTypes.Selected(lngCount) Then
        strAbsTypes = strAbsTypes & _
          IIf(strAbsTypes <> vbNullString, ", ", "") & _
          "'" & Replace(lstTypes.List(lngCount), "'", "''") & "'"
      End If
    Next
  End If

  If optAllRecords.Value Then
    strPicklistFilterType = "A"
    lngPicklistFilterID = 0
  ElseIf optPicklist.Value Then
    strPicklistFilterType = "P"
    lngPicklistFilterID = Val(txtPicklist.Tag)
  Else
    strPicklistFilterType = "F"
    lngPicklistFilterID = Val(txtFilter.Tag)
  End If

  
  Select Case mstrReportType
  Case "BradfordFactor"
    
    Me.HelpContextID = 1133
    Set pobjBradfordIndex = New clsCustomReportsRUN
    
    pobjBradfordIndex.SetOutputParameters _
          mobjOutputDef.GetSelectedFormatIndex, _
          (chkDestination(desScreen).Value = vbChecked), _
          (chkDestination(desPrinter).Value = vbChecked), _
          cboPrinterName.Text, _
          (chkDestination(desSave).Value = vbChecked), _
          cboSaveExisting.ListIndex, _
          (chkDestination(desEmail).Value), _
          Val(txtEmailGroup.Tag), _
          txtEmailSubject.Text, _
          txtEmailAttachAs.Text, _
          txtFilename.Text, _
          (chkPreview.Value = vbChecked), _
          (chkPrintFilterHeader.Value = vbChecked)

    If (mlngSingleRecord > 0) Then
      pobjBradfordIndex.SetBradfordParameters _
            dtDate(0).DateValue, dtDate(1).DateValue, _
            lngPicklistFilterID, strPicklistFilterType, _
            strAbsTypes, _
            cboOrderBy.List(cboOrderBy.ListIndex), _
            IIf(chkSort1Asc.Value = vbChecked, "1", "0"), _
            cboGroupBy.List(cboGroupBy.ListIndex), _
            IIf(chkSort2Asc.Value = vbChecked, "1", "0"), _
            IIf(chkSRV.Value = vbChecked, "1", "0"), _
            IIf(chkShowCount.Value = vbChecked, "1", "0"), _
            IIf(chkShowTotals.Value = vbChecked, "1", "0"), _
            IIf(chkShowWorkings.Value = vbChecked, "1", "0"), _
            IIf(chkOmitAbsenceBeforeStart.Value = vbChecked, "1", "0"), _
            IIf(chkOmitAbsenceAfterEnd.Value = vbChecked, "1", "0"), _
            IIf(chkMinimumBradfordFactor.Value = vbChecked, "1", "0"), _
            spnMinimumBradfordFactor.Value, _
            IIf(chkDisplayDetail.Value = vbChecked, "1", "0"), _
            0, _
            0
      
    
    Else
      pobjBradfordIndex.SetBradfordParameters _
            dtDate(0).DateValue, dtDate(1).DateValue, _
            lngPicklistFilterID, strPicklistFilterType, _
            strAbsTypes, _
            cboOrderBy.List(cboOrderBy.ListIndex), _
            IIf(chkSort1Asc.Value = vbChecked, "1", "0"), _
            cboGroupBy.List(cboGroupBy.ListIndex), _
            IIf(chkSort2Asc.Value = vbChecked, "1", "0"), _
            IIf(chkSRV.Value = vbChecked, "1", "0"), _
            IIf(chkShowCount.Value = vbChecked, "1", "0"), _
            IIf(chkShowTotals.Value = vbChecked, "1", "0"), _
            IIf(chkShowWorkings.Value = vbChecked, "1", "0"), _
            IIf(chkOmitAbsenceBeforeStart.Value = vbChecked, "1", "0"), _
            IIf(chkOmitAbsenceAfterEnd.Value = vbChecked, "1", "0"), _
            IIf(chkMinimumBradfordFactor.Value = vbChecked, "1", "0"), _
            spnMinimumBradfordFactor.Value, _
            IIf(chkDisplayDetail.Value = vbChecked, "1", "0"), _
            cboOrderBy.ItemData(cboOrderBy.ListIndex), _
            cboGroupBy.ItemData(cboGroupBy.ListIndex)
      
    End If
    
    fOK = pobjBradfordIndex.RunBradfordReport(mlngSingleRecord)
    Set pobjBradfordIndex = Nothing
  
  Case Else
    Set frmRun = New frmCrossTabRun

    frmRun.SetOutputParameters _
          mobjOutputDef.GetSelectedFormatIndex, _
          (chkDestination(desScreen).Value = vbChecked), _
          (chkDestination(desPrinter).Value = vbChecked), _
          cboPrinterName.Text, _
          (chkDestination(desSave).Value = vbChecked), _
          cboSaveExisting.ListIndex, _
          (chkDestination(desEmail).Value), _
          Val(txtEmailGroup.Tag), _
          txtEmailSubject.Text, _
          txtEmailAttachAs.Text, _
          txtFilename.Text, _
          (chkPreview.Value = vbChecked), _
          (chkPrintFilterHeader.Value = vbChecked)

    Select Case mstrReportType
    Case "AbsenceBreakdown"
      Me.HelpContextID = 1134
      frmRun.SetAbsenceBreakdownParameters _
            dtDate(0).DateValue, dtDate(1).DateValue, _
            lngPicklistFilterID, strPicklistFilterType, _
            strAbsTypes
      fOK = frmRun.AbsenceBreakdownExecuteReport(mlngSingleRecord)

    Case "Stability"
      Me.HelpContextID = 1135
      frmRun.SetTurnoverParameters _
            dtDate(0).DateValue, dtDate(1).DateValue, _
            lngPicklistFilterID, strPicklistFilterType, _
            cboVerCol.ItemData(cboVerCol.ListIndex), _
            cboHorCol.ItemData(cboHorCol.ListIndex), _
            False
      fOK = frmRun.TurnoverStabilityReport

    Case "Turnover"
      Me.HelpContextID = 1136
      frmRun.SetTurnoverParameters _
            dtDate(0).DateValue, dtDate(1).DateValue, _
            lngPicklistFilterID, strPicklistFilterType, _
            cboVerCol.ItemData(cboVerCol.ListIndex), _
            cboHorCol.ItemData(cboHorCol.ListIndex), _
            (chkNewStarters.Value = vbChecked)
      fOK = frmRun.TurnoverExecuteReport

    End Select
  
    If fOK Then
      If frmRun.PreviewOnScreen Then
        'Needed for the correct Context Sensitive Help page
        frmRun.HelpContextID = Me.HelpContextID
        frmRun.Show vbModal
      End If
    End If

    'JPD 20060105 Fault 10674
    Unload frmRun
  End Select

  Set frmRun = Nothing

End Sub




Private Sub spnMinimumBradfordFactor_Change()
Changed = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  EnableControls
End Sub


Private Sub EnableControls()

  fraAbsence.Enabled = (SSTab1.Tab = 0)
  fraColumns.Enabled = ((SSTab1.Tab = 0) And ((mstrReportType = "Stability") Or (mstrReportType = "Turnover")))
  fraDateRangeDef.Enabled = (SSTab1.Tab = 0)
  fraDateRangeRun.Enabled = (SSTab1.Tab = 0)
  fraDisplayOptions.Enabled = (SSTab1.Tab = 0)
  fraRecordSelection.Enabled = (SSTab1.Tab = 0)
  
  fraSortOrder.Enabled = ((SSTab1.Tab = 0) And (mlngSingleRecord < 1))
  ControlsDisableAll fraSortOrder, fraSortOrder.Enabled
  If (mlngSingleRecord > 0) Then
    cboOrderBy.ListIndex = -1
    cboGroupBy.ListIndex = -1
    chkSort1Asc.Value = vbUnchecked
    chkSort2Asc.Value = vbUnchecked
    chkPrintFilterHeader.Value = vbUnchecked
    chkMinimumBradfordFactor.Enabled = False
    chkMinimumBradfordFactor.Value = vbUnchecked
    spnMinimumBradfordFactor.Enabled = False
  End If
  
  If cboOrderBy.ListIndex > -1 Then
    If Not (cboOrderBy.ItemData(cboOrderBy.ListIndex) > 0) Then
      chkSort1Asc.Value = vbUnchecked
      chkSort1Asc.Enabled = False
    End If
  End If
  
  If cboGroupBy.ListIndex > -1 Then
   If Not (cboGroupBy.ItemData(cboGroupBy.ListIndex) > 0) Then
     chkSort2Asc.Value = vbUnchecked
     chkSort2Asc.Enabled = False
   End If
  End If
  
  If chkDisplayDetail.Value = vbUnchecked Then
    chkSRV.Value = vbUnchecked
    chkSRV.Enabled = False
  Else
    chkSRV.Enabled = True
  End If

  
  fraOutputDestination.Enabled = (SSTab1.Tab = 1)
  'fraOutputFilename.Enabled = (SSTab1.Tab = 1)
  fraOutputFormat.Enabled = (SSTab1.Tab = 1)

End Sub

Public Property Get Action() As ReportOptions
  Action = lngAction
End Property

Private Sub txtDateExpr_Change(Index As Integer)
  Changed = True
End Sub

Private Sub txtEmailAttachAs_Change()
Changed = True
End Sub

Private Sub txtEmailGroup_Change()
Changed = True
End Sub

Private Sub txtEmailSubject_Change()
Changed = True
End Sub

Private Sub txtFilename_Change()
Changed = True
End Sub

Private Sub txtFilter_Change()
Changed = True
End Sub

Private Sub txtPicklist_Change()
Changed = True
End Sub

