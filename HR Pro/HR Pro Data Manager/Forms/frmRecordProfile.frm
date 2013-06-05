VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmRecordProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Profile Definition"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1063
   Icon            =   "frmRecordProfile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDocument 
      Height          =   510
      Left            =   4110
      Picture         =   "frmRecordProfile.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8700
      TabIndex        =   92
      Top             =   5900
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7395
      TabIndex        =   91
      Top             =   5900
      Width           =   1200
   End
   Begin VB.PictureBox picNoDrop 
      Height          =   495
      Left            =   1380
      Picture         =   "frmRecordProfile.frx":0596
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   5865
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2595
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   50
      TabIndex        =   93
      Top             =   50
      Width           =   9850
      _ExtentX        =   17383
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmRecordProfile.frx":0E60
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBase"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraInformation"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Related Ta&bles"
      TabPicture(1)   =   "frmRecordProfile.frx":0E7C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRelatedTables"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Colu&mns"
      TabPicture(2)   =   "frmRecordProfile.frx":0E98
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTable"
      Tab(2).Control(1)=   "fraFieldsAvailable"
      Tab(2).Control(2)=   "fraFieldsSelected"
      Tab(2).Control(3)=   "cmdAddSeparator"
      Tab(2).Control(4)=   "cmdAddHeading"
      Tab(2).Control(5)=   "cmdAdd"
      Tab(2).Control(6)=   "cmdRemove"
      Tab(2).Control(7)=   "cmdMoveUp"
      Tab(2).Control(8)=   "cmdMoveDown"
      Tab(2).Control(9)=   "cmdAddAll"
      Tab(2).Control(10)=   "cmdRemoveAll"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Outpu&t"
      TabPicture(3)   =   "frmRecordProfile.frx":0EB4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraOutputFormat"
      Tab(3).Control(1)=   "fraReportOptions"
      Tab(3).Control(2)=   "fraOutputDestination"
      Tab(3).ControlCount=   3
      Begin VB.Frame fraInformation 
         Height          =   1950
         Left            =   150
         TabIndex        =   0
         Top             =   400
         Width           =   9550
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5950
            MaxLength       =   30
            TabIndex        =   6
            Top             =   300
            Width           =   3465
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   2
            Top             =   300
            Width           =   3150
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1350
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   700
            Width           =   3150
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1080
            Left            =   5985
            TabIndex        =   98
            Top             =   720
            Width           =   3405
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
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
            stylesets(0).Picture=   "frmRecordProfile.frx":0ED0
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
            stylesets(1).Picture=   "frmRecordProfile.frx":0EEC
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
            Columns.Count   =   3
            Columns(0).Width=   2963
            Columns(0).Caption=   "User Group"
            Columns(0).Name =   "GroupName"
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   2566
            Columns(1).Caption=   "Access"
            Columns(1).Name =   "Access"
            Columns(1).AllowSizing=   0   'False
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(1).Style=   3
            Columns(1).Row.Count=   3
            Columns(1).Col.Count=   2
            Columns(1).Row(0).Col(0)=   "Read / Write"
            Columns(1).Row(1).Col(0)=   "Read Only"
            Columns(1).Row(2).Col(0)=   "Hidden"
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "SysSecMgr"
            Columns(2).Name =   "SysSecMgr"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   6006
            _ExtentY        =   1905
            _StockProps     =   79
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
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Left            =   5115
            TabIndex        =   5
            Top             =   360
            Width           =   675
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   1
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   195
            TabIndex        =   3
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   5115
            TabIndex        =   7
            Top             =   750
            Width           =   690
         End
      End
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   3975
         Left            =   -72500
         TabIndex        =   96
         Top             =   1680
         Width           =   7200
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3315
            TabIndex        =   90
            Tag             =   "0"
            Top             =   3460
            Width           =   3705
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3315
            Locked          =   -1  'True
            TabIndex        =   79
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1760
            Width           =   3360
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3315
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1240
            Width           =   3705
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3315
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   2160
            Width           =   3705
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3315
            TabIndex        =   88
            Top             =   3060
            Width           =   3705
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3315
            Locked          =   -1  'True
            TabIndex        =   85
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   2660
            Width           =   3360
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            TabIndex        =   80
            Top             =   1760
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            TabIndex        =   86
            Top             =   2660
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   73
            Top             =   850
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   74
            Top             =   1300
            Width           =   1650
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   77
            Top             =   1820
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   83
            Top             =   2720
            Width           =   1560
         End
         Begin VB.CheckBox chkPreview 
            Caption         =   "P&review on screen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   195
            TabIndex        =   72
            Top             =   400
            Width           =   3495
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   1845
            TabIndex        =   89
            Top             =   3525
            Width           =   1155
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1845
            TabIndex        =   78
            Top             =   1815
            Width           =   1140
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1845
            TabIndex        =   87
            Top             =   3120
            Width           =   1395
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   1845
            TabIndex        =   84
            Top             =   2715
            Width           =   1290
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1845
            TabIndex        =   81
            Top             =   2220
            Width           =   1440
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1845
            TabIndex        =   75
            Top             =   1305
            Width           =   1545
         End
      End
      Begin VB.Frame fraReportOptions 
         Caption         =   "Display Options :"
         Enabled         =   0   'False
         Height          =   1250
         Left            =   -74850
         TabIndex        =   60
         Top             =   400
         Width           =   9550
         Begin VB.CheckBox chkShowTableRelationshipTitle 
            Caption         =   "&Show Table Relationship Titles"
            Height          =   240
            Left            =   200
            TabIndex        =   63
            Top             =   880
            Width           =   4000
         End
         Begin VB.CheckBox chkIndent 
            Caption         =   "I&ndent Related Tables"
            Height          =   240
            Left            =   195
            TabIndex        =   61
            Top             =   280
            Width           =   2310
         End
         Begin VB.CheckBox chkSuppressEmptyRelatedTableTitles 
            Caption         =   "Suppress Empty Related Table Tit&les"
            Height          =   240
            Left            =   200
            TabIndex        =   62
            Top             =   580
            Width           =   4000
         End
      End
      Begin VB.Frame fraOutputFormat 
         Caption         =   "Output Format :"
         Height          =   3975
         Left            =   -74850
         TabIndex        =   64
         Top             =   1680
         Width           =   2200
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "D&ata Only"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   65
            Top             =   400
            Value           =   -1  'True
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CS&V File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   71
            Top             =   3200
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   66
            Top             =   800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   67
            Top             =   1200
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   68
            Top             =   1600
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Char&t"
            Enabled         =   0   'False
            Height          =   195
            Index           =   5
            Left            =   200
            TabIndex        =   69
            Top             =   2400
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel P&ivot Table"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   200
            TabIndex        =   70
            Top             =   2800
            Visible         =   0   'False
            Width           =   1900
         End
      End
      Begin VB.Frame fraTable 
         Caption         =   "Table :"
         Height          =   700
         Left            =   -74850
         TabIndex        =   39
         Top             =   400
         Width           =   9550
         Begin VB.ComboBox cboTblAvailable 
            Height          =   315
            Left            =   200
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   240
            Width           =   3400
         End
         Begin VB.Label lblTable 
            Caption         =   "Table :"
            Height          =   255
            Left            =   200
            TabIndex        =   40
            Top             =   300
            Visible         =   0   'False
            Width           =   500
         End
      End
      Begin VB.Frame fraFieldsAvailable 
         Caption         =   "Columns Available :"
         Height          =   4425
         Left            =   -74850
         TabIndex        =   42
         Top             =   1200
         Width           =   3800
         Begin ComctlLib.ListView ListView1 
            Height          =   3900
            Left            =   200
            TabIndex        =   43
            Top             =   300
            Width           =   3400
            _ExtentX        =   6006
            _ExtentY        =   6879
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   5644
            EndProperty
         End
      End
      Begin VB.Frame fraFieldsSelected 
         Caption         =   "Columns Selected :"
         Height          =   4425
         Left            =   -69100
         TabIndex        =   52
         Top             =   1200
         Width           =   3800
         Begin VB.TextBox txtProp_ColumnHeading 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   55
            Top             =   3150
            Width           =   2400
         End
         Begin COASpinner.COA_Spinner spnSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   57
            Top             =   3555
            Width           =   1005
            _ExtentX        =   1773
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
            Enabled         =   0   'False
            MaximumValue    =   2147483647
            Text            =   "0"
         End
         Begin ComctlLib.ListView ListView2 
            Height          =   2745
            Left            =   195
            TabIndex        =   53
            Top             =   300
            Width           =   3400
            _ExtentX        =   6006
            _ExtentY        =   4842
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Column"
               Object.Tag             =   "Column"
               Text            =   "Column"
               Object.Width           =   5644
            EndProperty
         End
         Begin COASpinner.COA_Spinner spnDec 
            Height          =   315
            Left            =   1200
            TabIndex        =   59
            Top             =   3945
            Width           =   1005
            _ExtentX        =   1773
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
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
         Begin VB.Label lblProp_ColumnHeading 
            BackStyle       =   0  'Transparent
            Caption         =   "Heading :"
            Height          =   195
            Left            =   195
            TabIndex        =   54
            Top             =   3210
            Width           =   1260
         End
         Begin VB.Label lblProp_Size 
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            Height          =   195
            Left            =   195
            TabIndex        =   56
            Top             =   3615
            Width           =   615
         End
         Begin VB.Label lblProp_Decimals 
            BackStyle       =   0  'Transparent
            Caption         =   "Decimals :"
            Height          =   195
            Left            =   195
            TabIndex        =   58
            Top             =   4010
            Width           =   1140
         End
      End
      Begin VB.Frame fraBase 
         Caption         =   "Data :"
         Height          =   3175
         Left            =   150
         TabIndex        =   8
         Top             =   2450
         Width           =   9555
         Begin VB.CheckBox chkPrintFilterHeader 
            Caption         =   "Display &title in the report header"
            Height          =   195
            Left            =   5150
            TabIndex        =   22
            Top             =   1560
            Width           =   3960
         End
         Begin VB.CheckBox chkBasePageBreak 
            Caption         =   "Pag&e Break"
            Height          =   195
            Left            =   5150
            TabIndex        =   23
            Top             =   1960
            Width           =   1380
         End
         Begin VB.Frame fraBaseOrientation 
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   6900
            TabIndex        =   25
            Top             =   2600
            Width           =   2600
            Begin VB.OptionButton optBaseOrientation 
               Caption         =   "&Vertical"
               Height          =   195
               Index           =   1
               Left            =   1450
               TabIndex        =   27
               Top             =   0
               Value           =   -1  'True
               Width           =   1100
            End
            Begin VB.OptionButton optBaseOrientation 
               Caption         =   "Hori&zontal"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.CommandButton cmdBaseOrder 
            Caption         =   "..."
            Height          =   315
            Left            =   4050
            Picture         =   "frmRecordProfile.frx":0F08
            TabIndex        =   13
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtBaseOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   12
            Tag             =   "0"
            Top             =   700
            Width           =   2700
         End
         Begin VB.TextBox txtBaseFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   20
            Tag             =   "0"
            Top             =   1100
            Width           =   2100
         End
         Begin VB.TextBox txtBasePicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   17
            Tag             =   "0"
            Top             =   700
            Width           =   2100
         End
         Begin VB.OptionButton optBaseFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   5950
            TabIndex        =   19
            Top             =   1160
            Width           =   840
         End
         Begin VB.OptionButton optBasePicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   5950
            TabIndex        =   16
            Top             =   760
            Width           =   930
         End
         Begin VB.OptionButton optBaseAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   5950
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   630
         End
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   300
            Width           =   3030
         End
         Begin VB.CommandButton cmdBasePicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   9000
            TabIndex        =   18
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdBaseFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   9000
            TabIndex        =   21
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.Label lblPageBreak 
            Caption         =   " (applies only for output to Word    and printing Data Only output)"
            Height          =   405
            Left            =   6500
            TabIndex        =   95
            Top             =   1965
            Width           =   2850
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblBaseOrientation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Orientation :"
            Height          =   195
            Left            =   5145
            TabIndex        =   24
            Top             =   2595
            Width           =   1545
         End
         Begin VB.Label lblBaseOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order :"
            Height          =   195
            Left            =   200
            TabIndex        =   11
            Top             =   760
            Width           =   525
         End
         Begin VB.Label lblBaseTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Left            =   200
            TabIndex        =   9
            Top             =   360
            Width           =   885
         End
         Begin VB.Label lblBaseRecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   5115
            TabIndex        =   14
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame fraRelatedTables 
         Height          =   5210
         Left            =   -74850
         TabIndex        =   28
         Top             =   400
         Width           =   9550
         Begin VB.CommandButton cmdAddAllTableColumns 
            Caption         =   "Add All Column&s"
            Height          =   400
            Left            =   7750
            TabIndex        =   37
            Top             =   3900
            Width           =   1600
         End
         Begin VB.CommandButton cmdAddAllRelatedTables 
            Caption         =   "Add A&ll..."
            Height          =   400
            Left            =   7750
            TabIndex        =   31
            Top             =   750
            Width           =   1600
         End
         Begin VB.CommandButton cmdAutoArrangeRelatedTables 
            Caption         =   "Au&to Arrange"
            Height          =   400
            Left            =   7750
            TabIndex        =   38
            Top             =   4500
            Width           =   1600
         End
         Begin VB.CommandButton cmdRemoveRelatedTable 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   7750
            TabIndex        =   33
            Top             =   1800
            Width           =   1600
         End
         Begin VB.CommandButton cmdAddRelatedTable 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7750
            TabIndex        =   30
            Top             =   300
            Width           =   1600
         End
         Begin VB.CommandButton cmdRemoveAllRelatedTables 
            Caption         =   "Remo&ve All "
            Height          =   400
            Left            =   7750
            TabIndex        =   34
            Top             =   2250
            Width           =   1600
         End
         Begin VB.CommandButton cmdEditRelatedTable 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   7750
            TabIndex        =   32
            Top             =   1200
            Width           =   1600
         End
         Begin SSDataWidgets_B.SSDBGrid grdRelatedTables 
            Height          =   4600
            Left            =   195
            TabIndex        =   29
            Top             =   300
            Width           =   7410
            ScrollBars      =   3
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   10
            AllowUpdate     =   0   'False
            MultiLine       =   0   'False
            AllowRowSizing  =   0   'False
            AllowGroupSizing=   0   'False
            AllowGroupMoving=   0   'False
            AllowColumnMoving=   0
            AllowGroupSwapping=   0   'False
            AllowColumnSwapping=   0
            AllowGroupShrinking=   0   'False
            AllowColumnShrinking=   0   'False
            AllowDragDrop   =   0   'False
            SelectTypeCol   =   0
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   10
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "TableID"
            Columns(0).Name =   "TableID"
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2619
            Columns(1).Caption=   "Table"
            Columns(1).Name =   "Table"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "FilterID"
            Columns(2).Name =   "FilterID"
            Columns(2).AllowSizing=   0   'False
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   2593
            Columns(3).Caption=   "Filter"
            Columns(3).Name =   "Filter"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   2064
            Columns(4).Caption=   "Order"
            Columns(4).Name =   "Order"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   1931
            Columns(5).Caption=   "Records"
            Columns(5).Name =   "Records"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "OrderID"
            Columns(6).Name =   "OrderID"
            Columns(6).AllowSizing=   0   'False
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   1244
            Columns(7).Caption=   "Break"
            Columns(7).Name =   "PageBreak"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(7).Style=   2
            Columns(8).Width=   2143
            Columns(8).Caption=   "Orientation"
            Columns(8).Name =   "Orientation"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "OrientationCode"
            Columns(9).Name =   "OrientationCode"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13070
            _ExtentY        =   8114
            _StockProps     =   79
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
         Begin VB.CommandButton cmdMoveRelatedTableUp 
            Caption         =   "&Up"
            Height          =   400
            Left            =   7750
            TabIndex        =   35
            Top             =   2865
            Width           =   1600
         End
         Begin VB.CommandButton cmdMoveRelatedTableDown 
            Caption         =   "Do&wn"
            Height          =   400
            Left            =   7750
            TabIndex        =   36
            Top             =   3300
            Width           =   1600
         End
      End
      Begin VB.CommandButton cmdAddSeparator 
         Caption         =   "Add &Separator"
         Height          =   400
         Left            =   -70800
         TabIndex        =   47
         Top             =   2800
         Width           =   1450
      End
      Begin VB.CommandButton cmdAddHeading 
         Caption         =   "Add &Heading"
         Height          =   400
         Left            =   -70800
         TabIndex        =   46
         Top             =   2300
         Width           =   1450
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   400
         Left            =   -70800
         TabIndex        =   44
         Top             =   1300
         Width           =   1450
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   400
         Left            =   -70800
         TabIndex        =   48
         Top             =   3500
         Width           =   1450
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "&Up"
         Enabled         =   0   'False
         Height          =   400
         Left            =   -70800
         TabIndex        =   50
         Top             =   4675
         Width           =   1450
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Do&wn"
         Enabled         =   0   'False
         Height          =   400
         Left            =   -70800
         TabIndex        =   51
         Top             =   5175
         Width           =   1450
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "Add A&ll"
         Height          =   400
         Left            =   -70800
         TabIndex        =   45
         Top             =   1800
         Width           =   1450
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remo&ve All"
         Enabled         =   0   'False
         Height          =   400
         Left            =   -70800
         TabIndex        =   49
         Top             =   4000
         Width           =   1450
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   240
      Top             =   5700
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
      Bands           =   "frmRecordProfile.frx":0F80
   End
End
Attribute VB_Name = "frmRecordProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DataAccess Class
Private datData As DataMgr.clsDataAccess

' Collection Class (Holds column details such as heading, size etc)
Public mcolRecordProfileColumnDetails As clsRecordProfileColDtls

' Long to hold current ID
Private mlngRecordProfileID As Long

' Variables to hold current (or previously) selected table details
Private mstrBaseTable As String
Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mblnFormPrint As Boolean
Private mblnColumnDrag As Boolean
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mlngTimeStamp As Long
Private mblnForceHidden As Boolean
Private mblnRecordSelectionInvalid As Boolean
Private mblnDefinitionCreator As Boolean
Private mblnDeleted As Boolean
Private mbNeedsSave As Boolean

Private Const lng_GRIDROWHEIGHT = 239
Private Const sALL_RECORDS = "All Records"

Private Const sDFLTTEXT_HEADING = "<Heading>"
Private Const sDFLTTEXT_SEPARATOR = "<Separator>"
Private Const sTYPECODE_HEADING = "H"
Private Const sTYPECODE_SEPARATOR = "S"
Private Const sTYPECODE_COLUMN = "C"

Private mavTables() As Variant

Private miDataType As SQLDataType

Private mobjOutputDef As clsOutputDef

Private miRelatedTableCount As Integer
Private Sub ForceAccess(Optional pvAccess As Variant)
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    .MoveFirst

    For iLoop = 0 To (.Rows - 1)
      varBookmark = .Bookmark
      
      If iLoop = 0 Then
        .Columns("Access").Text = ""
      Else
        If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
          If mblnForceHidden Then
            .Columns("Access").Text = AccessDescription(ACCESS_HIDDEN)
          Else
            If Not IsMissing(pvAccess) Then
              .Columns("Access").Text = AccessDescription(CStr(pvAccess))
            End If
          End If
        End If
      End If
      
      .MoveNext
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow

End Sub

Private Function AllHiddenAccess() As Boolean
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        If .Columns("Access").CellText(varBookmark) <> AccessDescription(ACCESS_HIDDEN) Then
          AllHiddenAccess = False
          Exit Function
        End If
      End If
    Next iLoop
  End With

  AllHiddenAccess = True
  
End Function


Private Function HiddenGroups() As String
  'Return a TAB delimited string of the user groups to which this definition is hidden.
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim sHiddenGroups As String
  
  sHiddenGroups = ""
  
  With grdAccess
    .Update
    For iLoop = 1 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        If .Columns("Access").CellText(varBookmark) = AccessDescription(ACCESS_HIDDEN) Then
          sHiddenGroups = sHiddenGroups & .Columns("GroupName").CellText(varBookmark) & vbTab
        End If
      End If
    Next iLoop
  End With

  If Len(sHiddenGroups) > 0 Then
    sHiddenGroups = vbTab & sHiddenGroups
  End If
  
  HiddenGroups = sHiddenGroups
  
End Function



Public Function Initialise(bNew As Boolean, _
  bCopy As Boolean, _
  Optional plngRecordProfileID As Long, _
  Optional bPrint As Boolean) As Boolean
  ' This function is called from frmMain and prepares the form depending
  ' on whether the user is creating a new definition or editing an existing
  ' one.
  
  Screen.MousePointer = vbHourglass
  
  ' Set references to class modules
  Set datData = New DataMgr.clsDataAccess
  
  mblnLoading = True

  If bNew Then
    mblnDefinitionCreator = True
    
    'Set ID to 0 to indicate new record
    mlngRecordProfileID = 0

    'Set controls to defaults
    ClearForNew
    
    'Load All Possible Base Tables into combo
    LoadBaseCombo

    UpdateDependantFields
    
    PopulateTableAvailable True

    ' Set command button status
    UpdateButtonStatus (SSTab1.Tab)
    
    PopulateAccessGrid
    
    Changed = False
  Else
    ' Make the record profile ID visible to the rest of the module
    mlngRecordProfileID = plngRecordProfileID
    
    ' Is is a copy of an existing one ?
    FromCopy = bCopy
    
    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint
    
    PopulateAccessGrid
    
    If Not RetrieveRecordProfileDetails(plngRecordProfileID) Then
      If mblnDeleted Or Cancelled Then
        Initialise = False
        Exit Function
      Else
        If COAMsgBox("OpenHR could not load all of the definition successfully. The recommendation is that" & vbCrLf & _
               "you delete the definition and create a new one, however, you may edit the existing" & vbCrLf & _
               "definition if you wish. Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, "Record Profile") = vbNo Then
          Cancelled = True
          Initialise = False
          Exit Function
        End If
      End If
    End If

    If bCopy = True Then
      mlngRecordProfileID = 0
      Changed = True
    Else
      Changed = mblnRecordSelectionInvalid And (Not mblnReadOnly) ' False
    End If
  End If
    
  EnableDisableTabControls
  Cancelled = False
  Screen.MousePointer = vbDefault
  mblnLoading = False
  
End Function




Private Function InsertRelatedTableDetails() As String
  Dim pvarbookmark  As Variant
  Dim i As Integer
  Dim sSQL As String

  With grdRelatedTables
    .MoveFirst
    For i = 0 To .Rows - 1 Step 1
      pvarbookmark = .GetBookmark(i)

      sSQL = "INSERT INTO ASRSysRecordProfileTables" & _
        " (recordProfileID, tableID, filterID, orderID, maxRecords, orientation, pageBreak, sequence)" & _
        " VALUES (" & mlngRecordProfileID & "," & _
        .Columns("TableID").CellValue(pvarbookmark) & "," & _
        IIf(.Columns("FilterID").CellValue(pvarbookmark) = vbNullString, 0, .Columns("FilterID").CellValue(pvarbookmark)) & "," & _
        IIf(.Columns("OrderID").CellValue(pvarbookmark) = vbNullString, 0, .Columns("OrderID").CellValue(pvarbookmark)) & "," & _
        IIf(.Columns("Records").CellValue(pvarbookmark) = sALL_RECORDS, 0, Val(.Columns("Records").CellValue(pvarbookmark))) & "," & _
        .Columns("OrientationCode").CellValue(pvarbookmark) & "," & _
        IIf(.Columns("PageBreak").CellValue(pvarbookmark), 1, 0) & "," & _
        Trim(Str(i + 1)) & ")"

      datData.ExecuteSql (sSQL)
    Next i
  End With

End Function


Private Sub RefreshTableAvailableOrder()
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim varBookmark As Variant
  Dim lngSelectedTableID As Long
    
  'JPD 20031009 Fault 7098
  lngSelectedTableID = 0
  If cboTblAvailable.ListCount > 0 Then
    lngSelectedTableID = cboTblAvailable.ItemData(cboTblAvailable.ListIndex)
  End If

  With grdRelatedTables
    If .Rows > 0 Then
      For iLoop = 0 To .Rows - 1 Step 1
        varBookmark = .AddItemBookmark(iLoop)
        
        If CInt(.Columns("TableID").CellValue(varBookmark)) <> cboTblAvailable.ItemData(iLoop + 1) Then
          cboTblAvailable.AddItem .Columns("Table").CellValue(varBookmark), iLoop + 1
          cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = CInt(.Columns("TableID").CellValue(varBookmark))
          
          For iLoop2 = (cboTblAvailable.ListCount - 1) To (cboTblAvailable.NewIndex + 1) Step -1
            If cboTblAvailable.ItemData(iLoop2) = CInt(.Columns("TableID").CellValue(varBookmark)) Then
              cboTblAvailable.RemoveItem iLoop2
            End If
          Next iLoop2
        End If
      Next iLoop
    End If
  End With

  If (cboTblAvailable.ListCount > 0) And (lngSelectedTableID > 0) Then
    For iLoop = 0 To (cboTblAvailable.ListCount - 1)
      If lngSelectedTableID = cboTblAvailable.ItemData(iLoop) Then
        cboTblAvailable.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

End Sub

Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysRecordProfileAccess WHERE ID = " & mlngRecordProfileID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysRecordProfileAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngRecordProfileID & ", sysusers.name," & _
    " CASE" & _
    "   WHEN (SELECT count(*)" & _
    "     FROM ASRSysGroupPermissions" & _
    "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    "     WHERE sysusers.Name = ASRSysGroupPermissions.groupname" & _
    "       AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
    "   ELSE '" & ACCESS_HIDDEN & "'" & _
    " END" & _
    " FROM sysusers" & _
    " WHERE sysusers.uid = sysusers.gid" & _
    " AND sysusers.name <> 'ASRSysGroup'" & _
    " AND sysusers.uid <> 0)"
  datData.ExecuteSql (sSQL)

  ' Update the new access records with the real access values.
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      sSQL = "IF EXISTS (SELECT * FROM ASRSysRecordProfileAccess" & _
        " WHERE ID = " & CStr(mlngRecordProfileID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysRecordProfileAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngRecordProfileID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub



Private Function InsertRecordProfile(pstrSQL As String) As Long
  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertRecordProfile_ERROR

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fSavedOK As Boolean
  
  fSavedOK = True

  Set cmADO = New ADODB.Command

  With cmADO
    .CommandText = "sp_ASRInsertNewUtility"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0

    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("newID", adInteger, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("insertString", adLongVarChar, adParamInput, -1)
    .Parameters.Append pmADO
    pmADO.Value = pstrSQL

    Set pmADO = .CreateParameter("tablename", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = "ASRSysRecordProfileName"

    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "recordProfileID"

    Set pmADO = Nothing

    cmADO.Execute

    If Not fSavedOK Then
      COAMsgBox "The new record profile could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertRecordProfile = 0
        Set cmADO = Nothing
        Exit Function
    End If

    InsertRecordProfile = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)

  End With

  Set cmADO = Nothing

TidyUpAndExit:
  Exit Function

InsertRecordProfile_ERROR:

  fSavedOK = False
  Resume TidyUpAndExit

End Function


Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property


Private Sub RefreshCollectionSequence()
  Dim iLoop As Integer
  
  For iLoop = 1 To ListView2.ListItems.Count
    mcolRecordProfileColumnDetails.Item(ListView2.ListItems(iLoop).Key).Sequence = iLoop
  Next iLoop

End Sub

Private Function RetrieveRecordProfileDetails(plngRecordProfileID As Long) As Boolean
  Dim rsTemp As Recordset
  Dim sText As String
  Dim fAlreadyNotified As Boolean
  Dim sMessage As String
  Dim rsTables As ADODB.Recordset
  Dim sSQL As String
  Dim sKey As String
  Dim rsOrder As ADODB.Recordset
  Dim sOrderName As String
  Dim lngFilterID As Long
  Dim sFilterName As String
  Dim lngOrderID As Long
  Dim fIsAscendant As Boolean
  Dim iLoop As Integer
  Dim iDataType As SQLDataType
  Dim iResult As RecordSelectionValidityCodes
  
  On Error GoTo Load_ERROR

  'Load the basic guff first
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysRecordProfileName.*, " & _
    "CONVERT(integer, ASRSysRecordProfileName.TimeStamp) AS intTimeStamp " & _
    "FROM ASRSysRecordProfileName WHERE recordProfileID = " & plngRecordProfileID)

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This Record Profile definition has been deleted by another user.", vbExclamation + vbOKOnly, "Record Profile"
    Set rsTemp = Nothing
    RetrieveRecordProfileDetails = False
    mblnDeleted = True
    Exit Function
  End If

  ' Set Definition Name
  txtName.Text = rsTemp!Name

  ' Set Definition Description
  txtDesc.Text = IIf(IsNull(rsTemp!Description), "", rsTemp!Description)

  If FromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If

  ' Set Base Table
  mblnLoading = True
  LoadBaseCombo

  SetComboText cboBaseTable, datGeneral.GetTableName(rsTemp!BaseTable)
  mstrBaseTable = cboBaseTable.Text
  UpdateDependantFields

  ' Set Base Table Record Select Options
  If rsTemp!AllRecords Then optBaseAllRecords.Value = True
  
  If rsTemp!PicklistID > 0 Then
    optBasePicklist.Value = True
    txtBasePicklist.Tag = rsTemp!PicklistID
    txtBasePicklist.Text = datGeneral.GetPicklistName(rsTemp!PicklistID)
  End If

  If rsTemp!FilterID > 0 Then
    optBaseFilter.Value = True
    txtBaseFilter.Tag = rsTemp!FilterID
    txtBaseFilter.Text = datGeneral.GetFilterName(rsTemp!FilterID)
  End If

  If rsTemp!OrderID > 0 Then
    sSQL = "SELECT name " & _
      "FROM ASRSysOrders " & _
      "WHERE orderID=" & rsTemp!OrderID
        
    Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
    With rsOrder
      If Not (.BOF And .EOF) Then
        sOrderName = Trim(!Name)
      Else
        sOrderName = vbNullString
      End If
      .Close
    End With
    
    Set rsOrder = Nothing
  
    txtBaseOrder.Tag = rsTemp!OrderID
    txtBaseOrder.Text = sOrderName
  End If

  If rsTemp!Orientation = giHORIZONTAL Then optBaseOrientation(0).Value = True Else optBaseOrientation(1).Value = True
  chkBasePageBreak.Value = IIf(rsTemp!PageBreak, vbChecked, vbUnchecked)
  If (rsTemp!PrintFilterHeader) And _
    ((rsTemp!FilterID > 0) Or (rsTemp!PicklistID > 0)) Then chkPrintFilterHeader.Value = vbChecked

  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  GetRelatedTables rsTemp!BaseTable, "PARENT"
  GetRelatedTables rsTemp!BaseTable, "CHILD"

  'Get the all the related table information.
  sSQL = "SELECT A.tableID, B.tableName, A.filterID, X.Name, A.maxRecords, A.orderID, A.orientation, A.pageBreak" & _
    " FROM ASRSysRecordProfileTables A " & _
    " INNER JOIN ASRSysTables B ON A.tableID = B.tableID " & _
    " LEFT OUTER JOIN ASRSysExpressions X ON A.filterID = X.exprID " & _
    " WHERE A.recordProfileID = " & plngRecordProfileID & " " & _
    " ORDER BY A.sequence"

  Set rsTables = datGeneral.GetRecords(sSQL)

  With rsTables
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        fIsAscendant = False
        For iLoop = 1 To UBound(mavTables, 2)
          If mavTables(1, iLoop) = rsTables!TableID Then
            fIsAscendant = mavTables(3, iLoop)
            Exit For
          End If
        Next iLoop
        
        If (Not fIsAscendant) Then
          lngFilterID = !FilterID
          sFilterName = IIf(IsNull(!Name), vbNullString, !Name)
        Else
          lngFilterID = 0
          sFilterName = vbNullString
        End If
        
        If (Not fIsAscendant) And (rsTables!OrderID > 0) Then
          sSQL = "SELECT name " & _
            "FROM ASRSysOrders " & _
            "WHERE orderID=" & !OrderID
              
          Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
          With rsOrder
            If Not (.BOF And .EOF) Then
              lngOrderID = rsTables!OrderID
              sOrderName = Trim(!Name)
            Else
              lngOrderID = 0
              sOrderName = vbNullString
            End If
            .Close
          End With
          
          Set rsOrder = Nothing
        Else
          lngOrderID = 0
          sOrderName = vbNullString
        End If
        
        grdRelatedTables.AddItem !TableID & _
          vbTab & !TableName & _
          vbTab & lngFilterID & _
          vbTab & sFilterName & _
          vbTab & sOrderName & _
          vbTab & IIf(fIsAscendant, "", IIf(!MaxRecords = 0, sALL_RECORDS, !MaxRecords)) & _
          vbTab & lngOrderID & _
          vbTab & !PageBreak & _
          vbTab & IIf(!Orientation = giHORIZONTAL, "Horizontal", "Vertical") & _
          vbTab & !Orientation

        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsTables = Nothing

  mlngTimeStamp = rsTemp!intTimestamp

  ' =========================

  mblnReadOnly = Not datGeneral.SystemPermission("RECORDPROFILE", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlRecordProfile, mlngRecordProfileID) = ACCESS_READONLY)
  End If
  
  'Only the creator of a definition can change the access regardless of definition access
  'Output Options.
  optOutputFormat(IIf(IsNull(rsTemp!OutputFormat), 0, rsTemp!OutputFormat)).Value = True
  mobjOutputDef.FormatClick IIf(IsNull(rsTemp!OutputFormat), 0, rsTemp!OutputFormat)

  chkPreview.Value = IIf(IIf(IsNull(rsTemp!OutputPreview), True, rsTemp!OutputPreview), vbChecked, vbUnchecked)
  chkDestination(desScreen).Value = IIf(IIf(IsNull(rsTemp!OutputScreen), True, rsTemp!OutputScreen), vbChecked, vbUnchecked)

  chkDestination(desPrinter).Value = IIf(IIf(IsNull(rsTemp!OutputPrinter), False, rsTemp!OutputPrinter), vbChecked, vbUnchecked)
  SetComboText cboPrinterName, IIf(IsNull(rsTemp!OutputPrinterName), "", rsTemp!OutputPrinterName)
  If rsTemp!OutputPrinterName <> vbNullString Then
    If cboPrinterName.Text <> rsTemp!OutputPrinterName Then
      cboPrinterName.AddItem rsTemp!OutputPrinterName
      cboPrinterName.ListIndex = cboPrinterName.NewIndex
      COAMsgBox "This definition is set to output to printer " & rsTemp!OutputPrinterName & _
             " which is not set up on your PC.", vbInformation, Me.Caption
    End If
  End If

  chkDestination(desSave).Value = IIf(IIf(IsNull(rsTemp!OutputSave), False, rsTemp!OutputSave), vbChecked, vbUnchecked)
  If chkDestination(desSave).Value Then
    SetComboItem cboSaveExisting, IIf(IsNull(rsTemp!OutputSaveExisting), 0, rsTemp!OutputSaveExisting)
  End If

  chkDestination(desEmail).Value = IIf(IIf(IsNull(rsTemp!OutputEmail), False, rsTemp!OutputEmail), vbChecked, vbUnchecked)

  If IIf(IsNull(rsTemp!OutputEmail), False, rsTemp!OutputEmail) Then
    txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
    txtEmailGroup.Tag = rsTemp!OutputEmailAddr
    txtEmailSubject.Text = rsTemp!OutputEmailSubject
    txtEmailAttachAs.Text = IIf(IsNull(rsTemp!OutputEmailAttachAs), vbNullString, rsTemp!OutputEmailAttachAs)
  End If
  txtFilename.Text = IIf(IsNull(rsTemp!OutputFilename), "", rsTemp!OutputFilename)

  chkIndent.Value = IIf(IIf(IsNull(rsTemp!IndentRelatedTables), False, rsTemp!IndentRelatedTables), vbChecked, vbUnchecked)
  chkSuppressEmptyRelatedTableTitles.Value = IIf(IIf(IsNull(rsTemp!SuppressEmptyRelatedTableTitles), False, rsTemp!SuppressEmptyRelatedTableTitles), vbChecked, vbUnchecked)
  chkShowTableRelationshipTitle.Value = IIf(IIf(IsNull(rsTemp!SuppressTableRelationshipTitles), False, Not rsTemp!SuppressTableRelationshipTitles), vbChecked, vbUnchecked)

  If mblnReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True

  ' =========================

  sMessage = vbNullString

  ' Now load the columns guff
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysRecordProfileDetails WHERE recordProfileID = " & plngRecordProfileID & " ORDER BY [Sequence]")

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "Cannot load the column definition for this Record Profile", vbExclamation + vbOKOnly, "Record Profile"
    RetrieveRecordProfileDetails = False
    Set rsTemp = Nothing
    Exit Function
  End If

  Do Until rsTemp.EOF

    Select Case rsTemp!Type
      Case "H":
        ' HEADING
        sText = "<Heading> " & rsTemp!Heading
        sKey = rsTemp!Type & CStr(rsTemp!ColumnID)
      Case "S":
        ' SEPARATOR
        sText = "<Separator>"
        sKey = rsTemp!Type & CStr(rsTemp!ColumnID)
      Case Else:
        ' COLUMN
        sText = datGeneral.GetColumnName(rsTemp!ColumnID)
        sKey = rsTemp!Type & CStr(rsTemp!ColumnID)
    End Select

    ' Add to collection
    If sText <> vbNullString And sMessage = vbNullString Then
      If rsTemp!Type = "C" Then
        iDataType = datGeneral.GetDataType(rsTemp!TableID, rsTemp!ColumnID)
      Else
        iDataType = 0
      End If
      
      mcolRecordProfileColumnDetails.Add sKey, rsTemp!Type, rsTemp!ColumnID, rsTemp!Heading, rsTemp!Size, rsTemp!dp, iDataType, rsTemp!TableID, sText, rsTemp!Sequence
    End If
    sMessage = vbNullString
    rsTemp.MoveNext
  Loop

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Cancelled = True
    RetrieveRecordProfileDetails = False
    Exit Function
  End If
  mblnLoading = False

  PopulateTableAvailable True

  UpdateButtonStatus (SSTab1.Tab)

  RetrieveRecordProfileDetails = True
  Exit Function

Load_ERROR:

  COAMsgBox "Warning : Error whilst retrieving the record profile definition." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Record Profile"
  RetrieveRecordProfileDetails = False
  Set rsTemp = Nothing

End Function

Private Function ForceDefinitionToBeHiddenIfNeeded(Optional pvOnlyFatalMessages As Variant) As Boolean
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim lngFilterID As Long
  Dim sRow As String
  Dim iResult As RecordSelectionValidityCodes
  Dim sBigMessage As String
  Dim asDeletedParameters() As String
  Dim asHiddenBySelfParameters() As String
  Dim asHiddenByOtherParameters() As String
  Dim asInvalidParameters() As String
  Dim fChangesRequired As Boolean
  Dim fDefnAlreadyHidden As Boolean
  Dim fNeedToForceHidden As Boolean
  Dim fRemove As Boolean
  Dim fOnlyFatalMessages As Boolean
  
  If IsMissing(pvOnlyFatalMessages) Then
    fOnlyFatalMessages = mblnLoading
  Else
    fOnlyFatalMessages = CBool(pvOnlyFatalMessages)
  End If
  
  ' Return false if some of the filters/picklists need to be removed from the definition,
  ' or if the definition needs to be made hidden.
  fChangesRequired = False
  fDefnAlreadyHidden = AllHiddenAccess
  fNeedToForceHidden = False

  ' Dimension arrays to hold details of the filters/picklists that
  ' have been deleted, made hidden or are now invalid.
  ' Column 1 - parameter type
  ' Column 2 - table name
  ReDim asDeletedParameters(2, 0)
  ReDim asHiddenBySelfParameters(2, 0)
  ReDim asHiddenByOtherParameters(2, 0)
  ReDim asInvalidParameters(2, 0)
  
  ' Check Base Table Picklist
  If (Len(txtBasePicklist.Tag) > 0) And (txtBasePicklist.Tag <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtBasePicklist.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        If fRemove Then
          sBigMessage = "The selected '" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, "Record Profile"
        Else
          fNeedToForceHidden = True
          
          ReDim Preserve asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2) + 1)
          asHiddenBySelfParameters(1, UBound(asHiddenBySelfParameters, 2)) = "picklist"
          asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(2, UBound(asDeletedParameters, 2) + 1)
        asDeletedParameters(1, UBound(asDeletedParameters, 2)) = "picklist"
        asDeletedParameters(2, UBound(asDeletedParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2) + 1)
          asHiddenByOtherParameters(1, UBound(asHiddenByOtherParameters, 2)) = "picklist"
          asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(2, UBound(asInvalidParameters, 2) + 1)
        asInvalidParameters(1, UBound(asInvalidParameters, 2)) = "picklist"
        asInvalidParameters(2, UBound(asInvalidParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)
        
        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtBasePicklist.Tag = 0
      txtBasePicklist.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Base Table Filter
  If Len(txtBaseFilter.Tag) > 0 And txtBaseFilter.Tag <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtBaseFilter.Tag))
    
    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        
        If fRemove Then
          sBigMessage = "The selected '" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, "Record Profile"
        Else
          fNeedToForceHidden = True
          
          ReDim Preserve asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2) + 1)
          asHiddenBySelfParameters(1, UBound(asHiddenBySelfParameters, 2)) = "filter"
          asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(2, UBound(asDeletedParameters, 2) + 1)
        asDeletedParameters(1, UBound(asDeletedParameters, 2)) = "filter"
        asDeletedParameters(2, UBound(asDeletedParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2) + 1)
          asHiddenByOtherParameters(1, UBound(asHiddenByOtherParameters, 2)) = "filter"
          asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(2, UBound(asInvalidParameters, 2) + 1)
        asInvalidParameters(1, UBound(asInvalidParameters, 2)) = "filter"
        asInvalidParameters(2, UBound(asInvalidParameters, 2)) = cboBaseTable.List(cboBaseTable.ListIndex)
    
        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select
    
    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtBaseFilter.Tag = 0
      txtBaseFilter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If
  
  ' Related Table Filters
  With grdRelatedTables
    If .Rows > 0 Then
      For iLoop = 0 To .Rows - 1 Step 1
        varBookmark = .AddItemBookmark(iLoop)
        lngFilterID = .Columns("FilterID").CellValue(varBookmark)

        If lngFilterID > 0 Then
          fRemove = False
          iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
          
          Select Case iResult
            Case REC_SEL_VALID_HIDDENBYUSER
              ' Filter hidden by the current user.
              ' Only a problem if the current definition is NOT owned by the current user,
              ' or if the current definition is not already hidden.
              fRemove = (Not mblnDefinitionCreator) And _
                (Not mblnReadOnly) And _
                (Not FormPrint)
              
              If fRemove Then
                sBigMessage = "The selected '" & .Columns("Table").CellValue(varBookmark) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
                COAMsgBox sBigMessage, vbExclamation + vbOKOnly, "Record Profile"
              Else
                fNeedToForceHidden = True
  
                ReDim Preserve asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2) + 1)
                asHiddenBySelfParameters(1, UBound(asHiddenBySelfParameters, 2)) = "filter"
                asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2)) = .Columns("Table").CellValue(varBookmark)
              End If
              
            Case REC_SEL_VALID_DELETED
              ' Filter deleted by another user.
              ReDim Preserve asDeletedParameters(2, UBound(asDeletedParameters, 2) + 1)
              asDeletedParameters(1, UBound(asDeletedParameters, 2)) = "filter"
              asDeletedParameters(2, UBound(asDeletedParameters, 2)) = .Columns("Table").CellValue(varBookmark)

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)

            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                ' Filter hidden by another user.
                ReDim Preserve asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2) + 1)
                asHiddenByOtherParameters(1, UBound(asHiddenByOtherParameters, 2)) = "filter"
                asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2)) = .Columns("Table").CellValue(varBookmark)
  
                fRemove = (Not mblnReadOnly) And _
                  (Not FormPrint)
              End If
            Case REC_SEL_VALID_INVALID
              ' Filter invalid.
              ReDim Preserve asInvalidParameters(2, UBound(asInvalidParameters, 2) + 1)
              asInvalidParameters(1, UBound(asInvalidParameters, 2)) = "filter"
              asInvalidParameters(2, UBound(asInvalidParameters, 2)) = .Columns("Table").CellValue(varBookmark)
            
              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)
          End Select

          If fRemove Then
            ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
            sRow = .Columns("TableID").CellValue(varBookmark) _
              & vbTab & .Columns("Table").CellValue(varBookmark) _
              & vbTab & 0 _
              & vbTab & vbNullString _
              & vbTab & .Columns("Order").CellValue(varBookmark) _
              & vbTab & .Columns("Records").CellValue(varBookmark) _
              & vbTab & .Columns("OrderID").CellValue(varBookmark) _
              & vbTab & .Columns("PageBreak").CellValue(varBookmark) _
              & vbTab & .Columns("Orientation").CellValue(varBookmark) _
              & vbTab & .Columns("OrientationCode").CellValue(varBookmark)

            .RemoveItem iLoop
            .AddItem sRow, iLoop

            .Bookmark = .AddItemBookmark(iLoop)
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(iLoop)

            If Not FormPrint Then
              SSTab1.Tab = 1
              .SetFocus
            End If
            
            mblnRecordSelectionInvalid = True
          End If
        End If
      Next iLoop
    End If
  End With

  ' Construct one big message with all of the required error messages.
  sBigMessage = ""
  
  If UBound(asHiddenBySelfParameters, 2) = 1 Then
    If FormPrint Or mblnReadOnly Then
      'JPD 20040219 Fault 7897
      If Not fDefnAlreadyHidden Then
        sBigMessage = "This definition needs to be made hidden as the selected '" & asHiddenBySelfParameters(2, 1) & "' table " & asHiddenBySelfParameters(1, 1) & " is hidden."
      End If
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If (Not mblnForceHidden) And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the selected '" & asHiddenBySelfParameters(2, 1) & "' table " & asHiddenBySelfParameters(1, 1) & " is hidden."
        End If
      Else
        If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden as the selected '" & asHiddenBySelfParameters(2, 1) & "' table " & asHiddenBySelfParameters(1, 1) & " is hidden."
        End If
      End If
    Else
      sBigMessage = "The selected '" & asHiddenBySelfParameters(2, 1) & "' table " & asHiddenBySelfParameters(1, 1) & " will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
    End If
  ElseIf UBound(asHiddenBySelfParameters, 2) > 1 Then
    If FormPrint Or mblnReadOnly Then
      'JPD 20040308 Fault 7897
      If Not fDefnAlreadyHidden Then
        sBigMessage = "This definition needs to be made hidden as the following parameters are hidden :" & vbCrLf
      End If
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If (Not mblnForceHidden) And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the following parameters are hidden :" & vbCrLf
        End If
      Else
        If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden as the following parameters are hidden :" & vbCrLf
        End If
      End If
    Else
      sBigMessage = "The following parameters will be removed from this definition as they are hidden and you do not have permission to make this definition hidden :" & vbCrLf
    End If
    
    If Len(sBigMessage) > 0 Then
      For iLoop = 1 To UBound(asHiddenBySelfParameters, 2)
        sBigMessage = sBigMessage & vbCrLf & vbTab & "'" & asHiddenBySelfParameters(2, iLoop) & "' table " & asHiddenBySelfParameters(1, iLoop)
      Next iLoop
    End If
  End If
      
  If UBound(asDeletedParameters, 2) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected '" & asDeletedParameters(2, 1) & "' table " & asDeletedParameters(1, 1) & " has been deleted."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected '" & asDeletedParameters(2, 1) & "' table " & asDeletedParameters(1, 1) & " will be removed from this definition as it has been deleted."
    End If
  ElseIf UBound(asDeletedParameters, 2) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been deleted :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been deleted :" & vbCrLf
    End If
    
    For iLoop = 1 To UBound(asDeletedParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & "'" & asDeletedParameters(2, iLoop) & "' table " & asDeletedParameters(1, iLoop)
    Next iLoop
  End If
  
  If UBound(asHiddenByOtherParameters, 2) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected '" & asHiddenByOtherParameters(2, 1) & "' table " & asHiddenByOtherParameters(1, 1) & " has been made hidden by another user."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected '" & asHiddenByOtherParameters(2, 1) & "' table " & asHiddenByOtherParameters(1, 1) & " will be removed from this definition as it has been made hidden by another user."
    End If
  ElseIf UBound(asHiddenByOtherParameters, 2) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been made hidden by another user :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been made hidden by another user :" & vbCrLf
    End If
    
    For iLoop = 1 To UBound(asHiddenByOtherParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & "'" & asHiddenByOtherParameters(2, iLoop) & "' table " & asHiddenByOtherParameters(1, iLoop)
    Next iLoop
  End If
  
  If UBound(asInvalidParameters, 2) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected '" & asInvalidParameters(2, 1) & "' table " & asInvalidParameters(1, 1) & " is invalid."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected '" & asInvalidParameters(2, 1) & "' table " & asInvalidParameters(1, 1) & " will be removed from this definition as it is invalid."
    End If
  ElseIf UBound(asInvalidParameters, 2) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters are invalid :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they are invalid :" & vbCrLf
    End If
    
    For iLoop = 1 To UBound(asInvalidParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & "'" & asInvalidParameters(2, iLoop) & "' table " & asInvalidParameters(1, iLoop)
    Next iLoop
  End If
  
  If Not FormPrint Then
    If mblnForceHidden And (Not fNeedToForceHidden) And (Not fOnlyFatalMessages) Then
      sBigMessage = "This definition no longer has to be hidden." & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        sBigMessage
    End If
    
    mblnForceHidden = fNeedToForceHidden
    ForceAccess
  End If
  
  If Len(sBigMessage) > 0 Then
    If FormPrint Then
      sBigMessage = "Record Profile print failed. The definition is currently invalid : " & vbCrLf & vbCrLf & sBigMessage
    End If
  
    COAMsgBox sBigMessage, vbExclamation + vbOKOnly, "Record Profile"
  End If
      
  ForceDefinitionToBeHiddenIfNeeded = (Len(sBigMessage) = 0)
      
End Function


Private Sub SelectAllColumns(plngTableID As Long)
  ' Add all columns from the selected table into the collection of selected columns.
  ' Copy items to the 'Selected' listview
  Dim rsColumns As New Recordset
  Dim sSQL As String
  Dim sColType As String
  Dim lID As Long
  Dim sHeading As String
  Dim lSize As Long
  Dim iDecPlaces As Integer
  Dim objItem As clsRecordProfileColDtl
  Dim iTemp As Integer
  
  Screen.MousePointer = vbHourglass

  iTemp = 1
  For Each objItem In mcolRecordProfileColumnDetails
    If objItem.Sequence >= iTemp Then
      iTemp = objItem.Sequence + 1
    End If
  Next objItem
  Set objItem = Nothing
  
  ' Get the columns of the given table
  sSQL = "SELECT columnID, tableID, columnName, dataType, defaultDisplayWidth, decimals" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & Trim(Str(plngTableID)) & _
    " AND columnType <> " & Trim(Str(colSystem)) & _
    " AND columnType <> " & Trim(Str(colLink)) & _
    " ORDER BY columnName"

  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  ' Check if the column has already been selected. If so, dont add it
  ' to the available listview
  Do While Not rsColumns.EOF
    If Not ColumnAlreadyUsed(rsColumns!ColumnID) Then
      sColType = sTYPECODE_COLUMN
      lID = rsColumns!ColumnID
      sHeading = rsColumns!ColumnName
      lSize = rsColumns!DefaultDisplayWidth
      iDecPlaces = rsColumns!Decimals

      mcolRecordProfileColumnDetails.Add sColType & Trim(Str(lID)), sColType, lID, sHeading, lSize, iDecPlaces, rsColumns!DataType, plngTableID, sHeading, iTemp
      iTemp = iTemp + 1
    End If
    rsColumns.MoveNext
  Loop
  ' Clear recordset reference
  Set rsColumns = Nothing
  
  Screen.MousePointer = vbDefault
  
End Sub

Public Property Get SelectedID() As Long
  SelectedID = mlngRecordProfileID
  
End Property


Public Sub PrintDef(plngRecordProfileID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsOrder As Recordset
  Dim sSQL As String
  Dim rsColumns As Recordset
  Dim rsRelatedTables As ADODB.Recordset
  Dim sTemp As String
  Dim sOrderName As String
  Dim lngLastTableID As Long
  Dim fIsAscendant As Boolean
  Dim iLoop As Integer
  Dim fFirstLoop As Boolean
  Dim varBookmark As Variant
  
  mlngRecordProfileID = plngRecordProfileID

  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysRecordProfileName.*, " & _
    "CONVERT(integer, ASRSysRecordProfileName.TimeStamp) AS intTimeStamp " & _
    "FROM ASRSysRecordProfileName WHERE recordProfileID = " & mlngRecordProfileID)

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Print Definition"
    Set rsTemp = Nothing
    Exit Sub
  End If

  GetRelatedTables rsTemp!BaseTable, "PARENT"
  GetRelatedTables rsTemp!BaseTable, "CHILD"
  
  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
    With objPrintDef
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader "Record Profile : " & rsTemp!Name
        
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal "Owner : " & rsTemp!UserName

        ' Access section --------------------------------------------------------
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop

        ' Data section --------------------------------------------------------
        .PrintTitle "Data"
        
        .PrintNormal "Base Table : " & datGeneral.GetTableName(rsTemp!BaseTable)
  
        If rsTemp!OrderID > 0 Then
          sSQL = "SELECT name " & _
            "FROM ASRSysOrders " & _
            "WHERE orderID=" & rsTemp!OrderID
              
          Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
          With rsOrder
            If Not (.BOF And .EOF) Then
              sOrderName = Trim(!Name)
            Else
              sOrderName = "<None>"
            End If
            .Close
          End With
          
          Set rsOrder = Nothing
        
          .PrintNormal "Order : " & sOrderName
        Else
          .PrintNormal "Order : <None>"
        End If
  
        If rsTemp!AllRecords Then
          .PrintNormal "Records : All Records"
        ElseIf rsTemp!PicklistID Then
          .PrintNormal "Records : '" & datGeneral.GetPicklistName(rsTemp!PicklistID) & "' picklist"
        ElseIf rsTemp!FilterID Then
          .PrintNormal "Records : '" & datGeneral.GetFilterName(rsTemp!FilterID) & "' filter"
        End If
        .PrintNormal
        .PrintNormal "Display filter or picklist title in the report header : " & IIf(rsTemp!PrintFilterHeader = True, "Yes", "No")
        .PrintNormal "Page Break : " & IIf(rsTemp!PageBreak, "Yes", "No")
        .PrintNormal "Data Orientation : " & IIf(rsTemp!Orientation = giHORIZONTAL, "Horizontal", "Vertical")
    
        ' Related Tables section --------------------------------------------------------
        .PrintTitle "Related Tables"
        
        Set rsRelatedTables = datGeneral.GetRecords("SELECT * FROM ASRSysRecordProfileTables WHERE recordProfileID = " & mlngRecordProfileID & " ORDER BY sequence")
        If rsRelatedTables.EOF And rsRelatedTables.BOF Then
          .PrintNormal "Related Tables : <None>"
          .PrintNormal
        Else
          fFirstLoop = True
          
          Do While Not rsRelatedTables.EOF
            fIsAscendant = False
            For iLoop = 1 To UBound(mavTables, 2)
              If mavTables(1, iLoop) = rsRelatedTables!TableID Then
                fIsAscendant = mavTables(3, iLoop)
                Exit For
              End If
            Next iLoop
              
            If fFirstLoop Then
              fFirstLoop = False
            Else
              .PrintNormal
            End If
            
            .PrintNormal "Related Table : " & datGeneral.GetTableName(rsRelatedTables!TableID)
            
            If Not fIsAscendant Then
              .PrintNormal "Filter : " & IIf(rsRelatedTables!FilterID > 0, datGeneral.GetFilterName(rsRelatedTables!FilterID), "<None>")
            
              If rsRelatedTables!OrderID > 0 Then
                sSQL = "SELECT name " & _
                  "FROM ASRSysOrders " & _
                  "WHERE orderID=" & rsRelatedTables!OrderID
                    
                Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockOptimistic)
                With rsOrder
                  If Not (.BOF And .EOF) Then
                    sOrderName = Trim(!Name)
                  Else
                    sOrderName = "<None>"
                  End If
                  .Close
                End With
                
                Set rsOrder = Nothing
              
                .PrintNormal "Order : " & sOrderName
              Else
                .PrintNormal "Order : <None>"
              End If
                
              .PrintNormal "Max Records : " & IIf(rsRelatedTables!MaxRecords = 0, "All Records", rsRelatedTables!MaxRecords)
            End If
            
            .PrintNormal "Page Break : " & IIf(rsRelatedTables!PageBreak, "Yes", "No")
            .PrintNormal "Data Orientation : " & IIf(rsRelatedTables!Orientation = giHORIZONTAL, "Horizontal", "Vertical")
            
            rsRelatedTables.MoveNext
          Loop
          rsRelatedTables.Close
          Set rsRelatedTables = Nothing
        End If
        
        ' Columns section --------------------------------------------------------
        .PrintTitle "Columns"
  
        .PrintBold "'" & datGeneral.GetTableName(rsTemp!BaseTable) & "' Table"
        .PrintNormal

        sSQL = "SELECT ASRSysRecordProfileDetails.type," & _
          " ASRSysRecordProfileDetails.heading," & _
          " ASRSysRecordProfileDetails.size," & _
          " ASRSysRecordProfileDetails.dp," & _
          " ASRSysRecordProfileDetails.tableID," & _
          " ASRSysColumns.columnName," & _
          " ASRSysColumns.dataType" & _
          " FROM ASRSysRecordProfileDetails" & _
          " INNER JOIN ASRSysColumns ON ASRSysRecordProfileDetails.columnID = ASRSysColumns.columnID" & _
          " WHERE ASRSysRecordProfileDetails.RecordProfileID = " & mlngRecordProfileID & _
          "   AND ASRSysRecordProfileDetails.tableID = " & rsTemp!BaseTable & _
          " ORDER BY ASRSysRecordProfileDetails.sequence"

        Set rsColumns = datGeneral.GetRecords(sSQL)
        Do While Not rsColumns.EOF
          Select Case rsColumns!Type
            Case sTYPECODE_HEADING:
              .PrintNormal "     " & "Type : Heading"
              .PrintNormal "     " & "Heading : " & rsColumns!Heading
            Case sTYPECODE_SEPARATOR:
              .PrintNormal "     " & "Type : Separator"
            Case Else
              .PrintNormal "     " & "Type : Column"
              .PrintNormal "     " & "Name : " & rsColumns!ColumnName
              .PrintNormal "     " & "Heading : " & rsColumns!Heading
              .PrintNormal "     " & "Size : " & rsColumns!Size

              If rsColumns!DataType = sqlNumeric Then
                .PrintNormal "     " & "Decimal Places : " & rsColumns!dp
              End If
          End Select

          .PrintNormal " "

          rsColumns.MoveNext

        Loop
        rsColumns.Close
        Set rsColumns = Nothing
        
        sSQL = "SELECT ASRSysRecordProfileDetails.type," & _
          " ASRSysRecordProfileDetails.heading," & _
          " ASRSysRecordProfileDetails.size," & _
          " ASRSysRecordProfileDetails.dp," & _
          " ASRSysRecordProfileDetails.tableID," & _
          " ASRSysColumns.columnName," & _
          " ASRSysColumns.dataType," & _
          " ASRSysTables.tableName" & _
          " FROM ASRSysRecordProfileTables" & _
          " INNER JOIN ASRSysTables ON ASRSysRecordProfileTables.tableID = ASRSysTables.tableID" & _
          " INNER JOIN ASRSysRecordProfileDetails ON (ASRSysRecordProfileTables.recordProfileID = ASRSysRecordProfileDetails.recordProfileID" & _
          "   AND ASRSysRecordProfileTables.tableID = ASRSysRecordProfileDetails.tableID)" & _
          " INNER JOIN ASRSysColumns ON ASRSysRecordProfileDetails.columnID = ASRSysColumns.columnID" & _
          " WHERE ASRSysRecordProfileTables.RecordProfileID = " & mlngRecordProfileID & _
          " ORDER BY ASRSysRecordProfileTables.sequence, ASRSysRecordProfileDetails.sequence"
        'sSQL = "SELECT ASRSysRecordProfileDetails.*, ASRSysRecordProfileTables.tableName, ASRSysColumns.dataType" & _
          " FROM ASRSysRecordProfileDetails" & _
          " INNER JOIN ASRSysRecordProfileTables ON ASRSysRecordProfileDetails.tableID = ASRSysRecordProfileTables.tableID" & _
          " WHERE ASRSysRecordProfileDetails.recordProfileID = " & mlngRecordProfileID & _
          " ORDER BY ASRSysRecordProfileTables.sequence, ASRSysRecordProfileDetails.sequence"
        
        Set rsColumns = datGeneral.GetRecords(sSQL)
        
        lngLastTableID = 0
        fFirstLoop = True
        Do While Not rsColumns.EOF
          If fFirstLoop Then
            fFirstLoop = False
          Else
            .PrintNormal
          End If
          
          If lngLastTableID <> rsColumns!TableID Then
            .PrintBold "'" & rsColumns!TableName & "' Table"
            .PrintNormal
            lngLastTableID = rsColumns!TableID
          End If
          
          Select Case rsColumns!Type
            Case sTYPECODE_HEADING:
              .PrintNormal "     " & "Type : Heading"
              .PrintNormal "     " & "Heading : " & rsColumns!Heading
            Case sTYPECODE_SEPARATOR:
              .PrintNormal "     " & "Type : Separator"
            Case Else
              .PrintNormal "     " & "Type : Column"
              .PrintNormal "     " & "Name : " & rsColumns!ColumnName
              .PrintNormal "     " & "Heading : " & rsColumns!Heading
              .PrintNormal "     " & "Size : " & rsColumns!Size
          
              If (rsColumns!DataType = sqlNumeric) Or (rsColumns!DataType = sqlInteger) Then
                .PrintNormal "     " & "Decimal Places : " & rsColumns!dp
              End If
          End Select
          
          rsColumns.MoveNext
        Loop
        rsColumns.Close
        Set rsColumns = Nothing
        
        ' Options section --------------------------------------------------------
        .PrintTitle "Output"
  
        .PrintNormal "Indent Related Tables : " & IIf(chkIndent.Value = vbChecked, "Yes", "No")
        .PrintNormal "Suppress Empty Related Table Titles : " & IIf(chkSuppressEmptyRelatedTableTitles.Value = vbChecked, "Yes", "No")
        .PrintNormal "Show Table Relationship Titles : " & IIf(chkShowTableRelationshipTitle.Value = vbChecked, "Yes", "No")
        
        .PrintNormal " "
        
        If optOutputFormat(0).Value Then .PrintNormal "Output Format : Data Only"
        If optOutputFormat(1).Value Then .PrintNormal "Output Format : CSV File"
        If optOutputFormat(2).Value Then .PrintNormal "Output Format : HTML Document"
        If optOutputFormat(3).Value Then .PrintNormal "Output Format : Word Document"
        If optOutputFormat(4).Value Then .PrintNormal "Output Format : Excel Worksheet"
        If optOutputFormat(5).Value Then .PrintNormal "Output Format : Excel Chart"
        If optOutputFormat(6).Value Then .PrintNormal "Output Format : Excel Pivot Table"
        
        If chkPreview.Value = vbChecked Then
          .PrintNormal "Output Destination : Preview on screen prior to output"
        End If

        If chkDestination(0).Value = vbChecked Then
          .PrintNormal "Output Destination : Display on screen"
        End If
        
        If chkDestination(1).Value = vbChecked Then
          .PrintNormal "Output Destination : Send to printer"
          .PrintNormal "Printer Location : " & cboPrinterName.List(cboPrinterName.ListIndex)
        End If
        
        If chkDestination(2).Value = vbChecked Then
          .PrintNormal "Output Destination : Save to file"
          .PrintNormal "File Name : " & txtFilename.Text
          .PrintNormal "File Options : " & cboSaveExisting.List(cboSaveExisting.ListIndex)
        End If
        
        If chkDestination(3).Value = vbChecked Then
          .PrintNormal "Output Destination : Send to email"
          .PrintNormal "Email Group : " & txtEmailGroup.Text
          .PrintNormal "Email Subject : " & txtEmailSubject.Text
          .PrintNormal "Email Attach As : " & txtEmailAttachAs.Text
        End If
        
        .PrintEnd
        .PrintConfirm "Record Profile : " & rsTemp!Name, "Record Profile Definition"
      End If
    End With
  End If

  Set rsTemp = Nothing
  Set rsColumns = Nothing

  Exit Sub

LocalErr:
  COAMsgBox "Printing Record Profile Definition Failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"

End Sub

Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property


Public Property Get FromCopy() As Boolean
  FromCopy = mblnFromCopy
End Property

Public Property Let FromCopy(ByVal bCopy As Boolean)
  mblnFromCopy = bCopy
End Property


Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOK.Enabled = pblnChanged
End Property


Private Function UniqueKey(psType As String) As String
  Dim objItem As clsRecordProfileColDtl
  Dim iKey As Integer
  Dim iNewKey As Integer
  
  iNewKey = 1
  
  For Each objItem In mcolRecordProfileColumnDetails
    If objItem.ColType = psType Then
      iKey = objItem.ID
    
      If iKey >= iNewKey Then
        iNewKey = iKey + 1
      End If
    End If
  Next objItem
  Set objItem = Nothing
  
  UniqueKey = psType & Trim(Str(iNewKey))
  
End Function

Private Function UpdateButtonStatus(iTab As Integer)
  On Error Resume Next

  Dim tempItem As ListItem
  Dim iCount As Integer

  Select Case iTab
    Case 1:
      cmdAddRelatedTable.Enabled = fraRelatedTables.Enabled And (miRelatedTableCount > 0)
      cmdAddAllRelatedTables.Enabled = cmdAddRelatedTable.Enabled
      
      If grdRelatedTables.Rows = 0 Then
        cmdEditRelatedTable.Enabled = False
        cmdRemoveRelatedTable.Enabled = False
        cmdRemoveAllRelatedTables.Enabled = False
        cmdAutoArrangeRelatedTables.Enabled = False
        cmdAddAllTableColumns.Enabled = False
      Else
        If grdRelatedTables.SelBookmarks.Count > 0 Then
          cmdEditRelatedTable.Enabled = Not mblnReadOnly
          cmdRemoveRelatedTable.Enabled = Not mblnReadOnly
        Else
          cmdEditRelatedTable.Enabled = False
          cmdRemoveRelatedTable.Enabled = False
        End If
        
        'AE20071025 Fault #7097
        'cmdAutoArrangeRelatedTables.Enabled = Not mblnReadOnly
        cmdAutoArrangeRelatedTables.Enabled = (grdRelatedTables.Rows > 1) And (Not mblnReadOnly)
        cmdRemoveAllRelatedTables.Enabled = Not mblnReadOnly
        cmdAddAllTableColumns.Enabled = Not mblnReadOnly
      End If
    
      With grdRelatedTables
        If grdRelatedTables.SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveRelatedTableUp.Enabled = False
            cmdMoveRelatedTableDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveRelatedTableUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveRelatedTableDown.Enabled = False
          Else
            cmdMoveRelatedTableUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveRelatedTableDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          End If
        Else
          cmdMoveRelatedTableUp.Enabled = False
          cmdMoveRelatedTableDown.Enabled = False
        End If
      End With
    
    Case 2:
      ' If there are no items to be selected, disable Add buttons
      If ListView1.ListItems.Count = 0 Then
        cmdAdd.Enabled = False
        cmdAddAll.Enabled = False
      Else
        cmdAdd.Enabled = Not mblnReadOnly
        cmdAddAll.Enabled = Not mblnReadOnly
      End If
  
      ' If there are no items in the 'Selected' Listview then disable move buttons and exit
      If ListView2.ListItems.Count = 0 Then
        cmdMoveUp.Enabled = False
        cmdMoveDown.Enabled = False
        cmdRemove.Enabled = False
        cmdRemoveAll.Enabled = False
        EnableColProperties False
      Else
        cmdRemove.Enabled = Not mblnReadOnly
        cmdRemoveAll.Enabled = Not mblnReadOnly
  
        ' If there are more than 1 items selected then disable the move buttons and exit
        For Each tempItem In ListView2.ListItems
          If tempItem.Selected Then iCount = iCount + 1
        Next tempItem
  
        If iCount <> 1 Then
          cmdMoveUp.Enabled = False
          cmdMoveDown.Enabled = False
          EnableColProperties False
        Else
          If ListView2.SelectedItem.Index <> 1 Then cmdMoveUp.Enabled = Not mblnReadOnly Else cmdMoveUp.Enabled = False
          If ListView2.SelectedItem.Index <> ListView2.ListItems.Count Then cmdMoveDown.Enabled = Not mblnReadOnly Else cmdMoveDown.Enabled = False
          EnableColProperties Not mblnReadOnly
        End If
      End If
    
  End Select
  
  Call CheckListViewColWidth(ListView1)
  Call CheckListViewColWidth(ListView2)

End Function


Private Sub CheckListViewColWidth(lstvw As ListView)
  Dim objItem As ListItem
  Dim lngMax As Long
  Dim lngLen As Long
  Dim lngSelectedItem As Long
  
  lngMax = 0
  lngSelectedItem = 0

  If lstvw.ListItems.Count > 0 Then
    
    For Each objItem In lstvw.ListItems
      If lngSelectedItem = 0 And objItem.Selected Then
        objItem.Selected = True
        lngSelectedItem = objItem.Index
      End If
    
      lngLen = Me.TextWidth(objItem.Text)
      If lngMax < lngLen Then
        lngMax = lngLen
      End If
    Next

    If lngSelectedItem = 0 Then
      lstvw.ListItems(1).Selected = True
    End If
  
  End If

  lngMax = lngMax + 60
  lstvw.ColumnHeaders(1).Width = lngMax
  lstvw.Refresh

End Sub



Private Function CheckUniqueName(sName As String, lngCurrentID As Long) As Boolean
  ' Is there already a definition with the same name (that isnt the
  ' definition we are editing ?)
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSYSRecordProfileName " & _
         "WHERE UPPER(Name) = '" & UCase(Replace(sName, "'", "''")) & "' " & _
         "AND recordProfileID <> " & lngCurrentID
  
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckUniqueName = True
  Else
    CheckUniqueName = False
  End If
  
  Set rsTemp = Nothing

End Function



Private Sub ClearRelatedTables(plngRecordProfileID As Long)

  ' Delete all column information from the Details table.
  
  Dim sSQL As String
  
  sSQL = "DELETE FROM ASRSysRecordProfileTables Where recordProfileID = " & plngRecordProfileID
  datData.ExecuteSql sSQL

End Sub


Private Sub ClearDetailTables(plngRecordProfileID As Long)
  ' Delete all column information from the Details table.
  Dim sSQL As String
  
  sSQL = "DELETE FROM ASRSysRecordProfileDetails Where recordProfileID = " & plngRecordProfileID
  datData.ExecuteSql sSQL

End Sub


Private Function EnableColProperties(bStatus As Boolean)
  Dim sType As String
  
  mblnLoading = True
  
  If Not ListView2.SelectedItem Is Nothing Then
    If ListView2.ListItems.Count > 0 Then GetCurrentDetails ListView2.SelectedItem.Key
  End If
  
  If mblnReadOnly Then
    Exit Function
  End If
  
  lblProp_ColumnHeading.Enabled = bStatus
  txtProp_ColumnHeading.Enabled = bStatus
  txtProp_ColumnHeading.BackColor = IIf(bStatus, &H80000005, &H8000000F)
  If Not bStatus Then txtProp_ColumnHeading.Text = ""
  
  lblProp_Size.Enabled = bStatus
  spnSize.Enabled = bStatus
  spnSize.BackColor = IIf(bStatus, &H80000005, &H8000000F)
  If Not bStatus Then spnSize.Text = ""
  
  lblProp_Decimals.Enabled = bStatus
  spnDec.Enabled = bStatus
  spnDec.BackColor = IIf(bStatus, &H80000005, &H8000000F)
  If Not bStatus Then spnDec.Text = ""
  
  If bStatus Then
    If (miDataType <> sqlNumeric) And (miDataType <> sqlInteger) Then
      spnDec.Enabled = False
      lblProp_Decimals.Enabled = spnDec.Enabled
      spnDec.BackColor = &H8000000F
    End If
    
    ' JPD20021204 - Disable 'Column Heading' controls for Separators
    sType = Left(ListView2.SelectedItem.Key, 1)
    If sType = sTYPECODE_SEPARATOR Then
      lblProp_ColumnHeading.Enabled = False
      txtProp_ColumnHeading.Enabled = False
      txtProp_ColumnHeading.BackColor = &H8000000F
    End If
    
    'JPD 20030610 Fault 5721 - Disable 'Size' controls for Photo columns
    ' JPD20021204 - Disable 'Size' controls for Separators & Headings
    If (sType = sTYPECODE_SEPARATOR) Or (sType = sTYPECODE_HEADING) Or (miDataType = sqlVarBinary) Then
      spnSize.Text = ""
      lblProp_Size.Enabled = False
      spnSize.Enabled = False
      spnSize.BackColor = &H8000000F
    End If
  End If
  
  mblnLoading = False

End Function


Private Function GetCurrentDetails(sKey As String) As Boolean
  ' This function returns the details held in the collection
  ' for the currently highlighted item in the 'selected'
  ' listview
  
  Dim objTemp As clsRecordProfileColDtl
  
  Set objTemp = mcolRecordProfileColumnDetails.Item(sKey)
  
  If objTemp Is Nothing Then
    txtProp_ColumnHeading.Text = ""
    spnSize.Text = 0
    spnDec.Text = 0
    EnableColProperties False
  Else
    txtProp_ColumnHeading.Text = objTemp.Heading
    spnSize.Text = objTemp.Size
    spnDec.Text = objTemp.DecPlaces
    
    miDataType = objTemp.DataType
  End If
    
  Set objTemp = Nothing
    
End Function

Public Sub PopulateTableAvailable(Optional pbSetToBase As Boolean)
  ' Populate the TableAvailable combo
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim pvarbookmark As Variant
  Dim fAdded As Boolean
  
  If pbSetToBase Then
    cboTblAvailable.Clear
    ' Clear the listview
    ListView1.ListItems.Clear
  End If
  
  ' Add the base table to the top of the combo
  If cboBaseTable.ListIndex >= 0 Then
    If Not TableAlreadyAvailable(cboBaseTable.ItemData(cboBaseTable.ListIndex)) Or pbSetToBase Then
      cboTblAvailable.AddItem cboBaseTable.Text
      cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = cboBaseTable.ItemData(cboBaseTable.ListIndex)
    End If
  End If

  ' Add the related tables to the combo if selected
  With grdRelatedTables
    If .Rows > 0 Then
      For iLoop = 0 To .Rows - 1 Step 1
        pvarbookmark = .AddItemBookmark(iLoop)
        If Not TableAlreadyAvailable(CInt(.Columns("TableID").CellValue(pvarbookmark))) Or pbSetToBase Then
          cboTblAvailable.AddItem .Columns("Table").CellValue(pvarbookmark)
          cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = CInt(.Columns("TableID").CellValue(pvarbookmark))
        End If
      Next iLoop
    End If
  End With

  If Not IsMissing(pbSetToBase) Then
    If pbSetToBase Then
      ' Select the base table in the combo by default
      If cboTblAvailable.ListCount > 0 Then
        cboTblAvailable.ListIndex = 0
      End If
    End If
  End If

  'JPD 20030609 Fault 5540
  ' Check the order of the tables in the combo matches the order in the related tables grid.
  RefreshTableAvailableOrder

  ' If theres only 1 table, then disable the combo, otherwise enable it
  If cboTblAvailable.ListCount = 1 Then
    cboTblAvailable.Enabled = False
    cboTblAvailable.BackColor = &H8000000F
  Else
    cboTblAvailable.Enabled = True
    cboTblAvailable.BackColor = &H80000005
  End If

End Sub


Private Function TableAlreadyAvailable(plngTableID As Long) As Boolean
  Dim i As Integer
  
  With cboTblAvailable
    For i = 0 To .ListCount - 1 Step 1
      If .ItemData(i) = plngTableID Then
        TableAlreadyAvailable = True
        Exit Function
      End If
    Next i
  End With
  
End Function



Private Sub UpdateDependantFields()
  ' This sub populates the parent/child combos depending
  ' on the base table selected

  Dim rsTables As New Recordset
  Dim sSQL As String
  Dim lngTableID As Long
  
  lngTableID = 0
  If cboBaseTable.ListIndex <> -1 Then
    lngTableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)
  End If

  ' Check if the selected base table has any related tables
  sSQL = "SELECT COUNT(*) AS result" & _
         " FROM ASRSysRelations " & _
         " WHERE ASRSysRelations.parentID = " & CStr(lngTableID) & _
         " OR ASRSysRelations.childID = " & CStr(lngTableID)

  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  miRelatedTableCount = rsTables!Result

  If rsTables!Result = 0 Then
    fraRelatedTables.Enabled = False
    cmdAddRelatedTable.Enabled = False
    cmdAddAllRelatedTables.Enabled = False
    cmdEditRelatedTable.Enabled = False
    cmdRemoveRelatedTable.Enabled = False
    cmdRemoveAllRelatedTables.Enabled = False
    cmdMoveRelatedTableUp.Enabled = False
    cmdMoveRelatedTableDown.Enabled = False
    cmdAutoArrangeRelatedTables.Enabled = False
    cmdAddAllTableColumns.Enabled = False
    grdRelatedTables.Enabled = False
  Else
    fraRelatedTables.Enabled = True
    grdRelatedTables.Enabled = True
  End If
  
  rsTables.Close
  Set rsTables = Nothing

  EnableDisableTabControls
  
End Sub



Public Sub LoadBaseCombo()
  ' Loads the Base combo with all tables (even lookups)
  
  Dim sSQL As String
  Dim rsTables As New Recordset

  sSQL = "SELECT tableName, tableID" & _
    " FROM ASRSysTables" & _
    " ORDER BY tableName"
  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With cboBaseTable
    .Clear
    Do While Not rsTables.EOF
      .AddItem rsTables!TableName
      .ItemData(.NewIndex) = rsTables!TableID
      rsTables.MoveNext
    Loop
    
    If .ListCount > 0 Then
      If gsPersonnelTableName <> "" Then
        SetComboText cboBaseTable, gsPersonnelTableName
      Else
        .ListIndex = 0
      End If

      'TM20020424 Fault 3802
      mstrBaseTable = .List(.ListIndex)
    End If
  End With

  rsTables.Close
  Set rsTables = Nothing
  
End Sub


Private Function IsRecordSelectionValid() As Boolean

  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim objCol As clsRecordProfileColDtl
  Dim iLoop As Integer
  Dim sKey As String
  Dim fNotified As Boolean
  Dim fNotifiedHdn As Boolean
  Dim i As Integer
  Dim pvarbookmark As Variant
  
  IsRecordSelectionValid = True

  ' Check that the filter on the Base Table still exists, and that it has
  ' not been made hidden by another user.
  
  If optBaseFilter.Value And txtBaseFilter.Tag > 0 Then
    
    sSQL = "SELECT * FROM AsrSysExpressions WHERE ExprID = " & txtBaseFilter.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    
    If rsTemp.BOF And rsTemp.EOF Then
      ' Filter has been deleted by another user
      COAMsgBox "The '" & txtBaseFilter.Text & "' filter has been deleted by another user.", vbExclamation + vbOKOnly, "Record Profile"
      txtBaseFilter.Tag = 0
      txtBaseFilter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(rsTemp!UserName) <> LCase(gsUserName) Then
      ' Filter has been made hidden by its owner
      COAMsgBox "The '" & txtBaseFilter.Text & "' filter has been made hidden by another user.", vbExclamation + vbOKOnly, "Record Profile"
      txtBaseFilter.Tag = 0
      txtBaseFilter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  
  ElseIf optBasePicklist.Value And txtBasePicklist.Tag > 0 Then
    sSQL = "SELECT * FROM AsrSysPicklistName WHERE PickListID = " & txtBasePicklist.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    
    If rsTemp.BOF And rsTemp.EOF Then
      ' Picklist has been deleted by another user
      COAMsgBox "The '" & txtBasePicklist.Text & "' picklist has been deleted by another user.", vbExclamation + vbOKOnly, "Record Profile"
      txtBasePicklist.Tag = 0
      txtBasePicklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(rsTemp!UserName) <> LCase(gsUserName) Then
      ' Picklist has been made hidden by its owner
      COAMsgBox "The '" & txtBasePicklist.Text & "' picklist has been made hidden by another user.", vbExclamation + vbOKOnly, "Record Profile"
      txtBasePicklist.Tag = 0
      txtBasePicklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  End If
  
  ' Related Tables
  With grdRelatedTables
    If .Rows > 0 Then
      .MoveFirst
      For i = 0 To .Rows - 1 Step 1
        pvarbookmark = .GetBookmark(i)
        If .Columns("FilterID").CellValue(pvarbookmark) > 0 Then
          sSQL = "SELECT * FROM AsrSysExpressions WHERE ExprID = " & .Columns("FilterID").CellValue(pvarbookmark)
          Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
          If rsTemp.BOF And rsTemp.EOF Then
            ' filter no longer exists !
            COAMsgBox "The '" & .Columns("Filter").CellValue(pvarbookmark) & "' filter has been deleted by another user.", vbExclamation + vbOKOnly, "Record Profile"
            .Columns("FilterID").CellValue(pvarbookmark) = 0
            .Columns("Filter").CellValue(pvarbookmark) = vbNullString
            Set rsTemp = Nothing
            IsRecordSelectionValid = False
            Exit Function
          ElseIf rsTemp!Access = "HD" And LCase(rsTemp!UserName) <> LCase(gsUserName) Then
            ' Filter has been made hidden by its owner
            COAMsgBox "The '" & .Columns("Filter").CellValue(pvarbookmark) & "' filter has been made hidden by another user.", vbExclamation + vbOKOnly, "Record Profile"
            .Columns("FilterID").CellValue(pvarbookmark) = 0
            .Columns("Filter").CellValue(pvarbookmark) = vbNullString
            Set rsTemp = Nothing
            IsRecordSelectionValid = False
            Exit Function
          End If
        End If
      Next i
    End If
  End With

  If IsRecordSelectionValid Then
    PopulateAvailable
  Else
    SelectLast ListView2
    GetCurrentDetails ListView2.SelectedItem.Key
  End If
  
  Set objCol = Nothing
  Set rsTemp = Nothing
  
End Function



Private Sub ClearForNew()
  
  'Clear out all fields required to be blank for a new record profile definition
  
  optBaseAllRecords.Value = True
  txtBasePicklist.Text = ""
  txtBasePicklist.Tag = 0
  txtBaseFilter.Text = ""
  txtBaseFilter.Tag = 0
  txtBaseOrder.Text = ""
  txtBaseOrder.Tag = 0
  optBaseOrientation(1).Value = True  ' Vertical for the Base Table
  chkBasePageBreak.Value = vbUnchecked
  chkPrintFilterHeader.Value = vbUnchecked
  
  If mblnDefinitionCreator Then txtUserName.Text = gsUserName
  
  ' Related tables tab
  grdRelatedTables.RemoveAll
  
  ' Columns Tab
  txtProp_ColumnHeading = ""
  'txtProp_Size = 0
  spnSize.Text = 0
  
  'txtProp_DecPlaces = 0
  spnDec.Text = "0"
  ListView2.ListItems.Clear

'  optOutputFormat(0).Value = True
'  mobjOutputDef.FormatClick 0

End Sub





Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  ' Process the tool click.
  Select Case Tool.Name
        
    Case "ID_Add"
      cmdAdd_Click
      
    Case "ID_AddAll"
      cmdAddAll_Click
      
    Case "ID_Remove"
      cmdRemove_Click
      
    Case "ID_RemoveAll"
      cmdRemoveAll_Click
      
    Case "ID_MoveUp"
      cmdMoveUp_Click

    Case "ID_MoveDown"
      cmdMoveDown_Click
  
    Case "ID_AddHeading"
      cmdAddHeading_Click
      
    Case "ID_AddSeparator"
      cmdAddSeparator_Click
      
  End Select

End Sub

Private Sub cboBaseTable_Click()
  ' When the user changes the Base Table, check to see if the user
  ' has defined any columns in the record profile. If they have, check that
  ' they have selected a different table in the combo to the one that
  ' was there before. If so, then prompt user, otherwise, go ahead and
  ' clear the definition
  If mblnLoading = True Then Exit Sub
  If mstrBaseTable = cboBaseTable.Text And (mblnLoading = False) Then Exit Sub

  If mcolRecordProfileColumnDetails.Count > 0 Or grdRelatedTables.Rows > 0 Then
    If COAMsgBox("Warning: Changing the base table will result in all table/column " & _
          "specific aspects of this record profile definition being cleared." & vbCrLf & _
          "Are you sure you wish to continue?", _
          vbQuestion + vbYesNo + vbDefaultButton2, "Record Profile") = vbYes Then

      mblnLoading = True
      ClearForNew
      mblnLoading = False
      Changed = True
    Else
      ' User opted to abort the base table change
      SetComboText cboBaseTable, mstrBaseTable
      Exit Sub
    End If
  Else
    Changed = True
  End If

  mstrBaseTable = cboBaseTable.Text

  '01/08/2001 MH Fault 2615
  optBaseAllRecords.Value = True

  mcolRecordProfileColumnDetails.RemoveAll ' this leaves 1 when u check the count prop!!!
  Set mcolRecordProfileColumnDetails = New clsRecordProfileColDtls

  UpdateDependantFields
  PopulateTableAvailable True
  ForceDefinitionToBeHiddenIfNeeded

End Sub





Private Sub cboPrinterName_Click()
  If Not mblnLoading Then
    Changed = True
  End If

End Sub


Private Sub cboSaveExisting_Click()
  If Not mblnLoading Then
    Changed = True
  End If

End Sub


Private Sub cboTblAvailable_Click()
  PopulateAvailable
  PopulateSelected
  UpdateButtonStatus (SSTab1.Tab)

End Sub


Public Sub PopulateAvailable()
  ' This function is called whenever a new table is selected in the
  ' table combo, or when cols are removed from the record profile
  ' definition. It checks through each item in the 'Selected'
  ' listview and if it doesnt find them, it adds them to the
  ' 'Available' listview.

  Dim rsColumns As New Recordset
  Dim sSQL As String

  If cboBaseTable.ListIndex = -1 Then Exit Sub

  ' Clear the contents of the Available Listview
  ListView1.ListItems.Clear

  ' Add the Columns of the selected table to the listview
  sSQL = "SELECT columnID, tableID, columnName" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & cboTblAvailable.ItemData(cboTblAvailable.ListIndex) & _
    " AND columnType <> " & Trim(Str(colSystem)) & _
    " AND columnType <> " & Trim(Str(colLink)) & _
    " ORDER BY columnName"
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  ' Check if the column has already been selected. If so, dont add it
  ' to the available listview
  Do While Not rsColumns.EOF
    If Not ColumnAlreadyUsed(rsColumns!ColumnID) Then
      ListView1.ListItems.Add , sTYPECODE_COLUMN & rsColumns!ColumnID, rsColumns!ColumnName
      'Debug.Print rsColumns!ColumnName & "-" & rsColumns!ColumnID
    Else
      'Debug.Print rsColumns!ColumnName & "-" & rsColumns!ColumnID
    End If
    rsColumns.MoveNext
  Loop
  ' Clear recordset reference
  Set rsColumns = Nothing

End Sub

Public Sub PopulateSelected()
  ' This function is called whenever the table is selected in the
  ' table combo.
  Dim objItem As clsRecordProfileColDtl
  Dim sText As String
  Dim iLoop As Integer
  Dim fAdded As Boolean
  
  If cboBaseTable.ListIndex = -1 Then Exit Sub

  ' Clear the contents of the Selected Listview
  ListView2.ListItems.Clear

  For Each objItem In mcolRecordProfileColumnDetails
    If objItem.TableID = cboTblAvailable.ItemData(cboTblAvailable.ListIndex) Then
      ' Add the selected columns of the selected table to the listview
      Select Case objItem.ColType
        Case sTYPECODE_SEPARATOR:
          sText = sDFLTTEXT_SEPARATOR
        Case sTYPECODE_HEADING:
          sText = sDFLTTEXT_HEADING
        Case Else
          sText = objItem.ColumnName
      End Select
      
      ' Add the items to the listview in the correct order.
      fAdded = False
      
      If ListView2.ListItems.Count > 0 Then
        For iLoop = 1 To ListView2.ListItems.Count
          If objItem.Sequence < mcolRecordProfileColumnDetails.Item(ListView2.ListItems(iLoop).Key).Sequence Then
            ListView2.ListItems.Add iLoop, objItem.ColType & objItem.ID, sText
            fAdded = True
            Exit For
          End If
        Next iLoop
      End If
      
      If Not fAdded Then
        ListView2.ListItems.Add , objItem.ColType & objItem.ID, sText
      End If
    End If
  Next objItem
  Set objItem = Nothing
  
  'JPD 20030610 Fault 5796 Select the top item if there is one.
  If ListView2.ListItems.Count > 0 Then
    ListView2.ListItems(1).Selected = True
  End If
  
End Sub



Private Function RemoveFromCollection(sKey As String) As Boolean

  mcolRecordProfileColumnDetails.Remove sKey

End Function



Private Function ColumnAlreadyUsed(plngColumnID As Long) As Boolean

  Dim objItem As clsRecordProfileColDtl
  Dim fUsed As Boolean
  
  fUsed = False
  
  For Each objItem In mcolRecordProfileColumnDetails
    If objItem.ID = plngColumnID And objItem.ColType = sTYPECODE_COLUMN Then
      fUsed = True
      Exit For
    End If
  Next objItem
  
  Set objItem = Nothing

  ColumnAlreadyUsed = fUsed
  
End Function


Private Sub chkBasePageBreak_Click()
  Changed = True

End Sub




Private Sub chkDestination_Click(Index As Integer)
  mobjOutputDef.DestinationClick Index
  Changed = True

End Sub


Private Sub chkIndent_Click()
  Changed = True

End Sub

Private Sub chkPreview_Click()
  Changed = True

End Sub


Private Sub chkPrintFilterHeader_Click()
  Changed = True
  
End Sub

Private Sub chkShowTableRelationshipTitle_Click()
  Changed = True

End Sub

Private Sub chkSuppressEmptyRelatedTableTitles_Click()
  Changed = True

End Sub



Private Sub cmdAdd_Click()
  ' Add the selected items to the 'Selected' Listview
  CopyToSelected False

End Sub


Private Function CopyToSelected(pfAll As Boolean, Optional piBeforeIndex As Integer)
  ' Copy items to the 'Selected' listview
  Dim fOK As Boolean
  Dim objTempItem As ListItem
  Dim objWorkingItem As ListItem
  Dim iSelectedCount As Integer
  Dim iItemToSelect As Integer
  Dim iItemsToDelete() As Variant
  ReDim iItemsToDelete(0)
  Dim intTemp As Integer

  Screen.MousePointer = vbHourglass

  'If user has clicked ADD ALL then do this...
  If pfAll Then
    For Each objTempItem In ListView1.ListItems
      ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text
      fOK = True
      If fOK Then AddToCollection objTempItem
    Next objTempItem

    ListView1.ListItems.Clear
RefreshCollectionSequence
    SelectFirst ListView2
    UpdateButtonStatus (SSTab1.Tab)
    Screen.MousePointer = vbDefault
    Changed = True
    Exit Function
  End If

  'Get count of how many items we are moving
  For Each objTempItem In ListView1.ListItems
    If objTempItem.Selected = True Then
      iSelectedCount = iSelectedCount + 1
      If iSelectedCount = 1 Then
        Set objWorkingItem = objTempItem
        iItemToSelect = objWorkingItem.Index
      End If
    End If
  Next objTempItem

  'If its just one item do this...
  If iSelectedCount = 1 Then
    Set objTempItem = objWorkingItem
    'If we are not inserting it before existing columns...
    If piBeforeIndex = 0 Then
      ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text
      fOK = True

      If fOK Then
        AddToCollection objTempItem
        ListView1.ListItems.Remove iItemToSelect
      End If
    Else
      ' Before index
      ListView2.ListItems.Add piBeforeIndex, objTempItem.Key, objTempItem.Text
      fOK = True

      If fOK Then
        AddToCollection objTempItem
        ListView1.ListItems.Remove iItemToSelect
      End If
    End If

    If ListView1.ListItems.Count > 0 Then
      If iItemToSelect > ListView1.ListItems.Count Then
        iItemToSelect = ListView1.ListItems.Count
      End If
      ListView1.ListItems(iItemToSelect).Selected = True
    End If

    If piBeforeIndex = 0 Then
      SelectLast ListView2
    Else
      For Each objTempItem In ListView2.ListItems
        objTempItem.Selected = (objTempItem.Index = piBeforeIndex)
      Next objTempItem
      Set ListView2.DropHighlight = Nothing
    End If

RefreshCollectionSequence
    UpdateButtonStatus (SSTab1.Tab)
    Screen.MousePointer = vbDefault
    Changed = True
    Exit Function
  End If

  'There are more than one item selected
  For Each objTempItem In ListView1.ListItems
    If objTempItem.Selected Then
      If piBeforeIndex = 0 Then
        ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text
        fOK = True
      Else
        ' Before an existing item
        ListView2.ListItems.Add piBeforeIndex, objTempItem.Key, objTempItem.Text
        fOK = True
        piBeforeIndex = piBeforeIndex + 1
      End If

      If fOK = True Then
        AddToCollection objTempItem
        ReDim Preserve iItemsToDelete(UBound(iItemsToDelete) + 1)
        iItemsToDelete(UBound(iItemsToDelete) - 1) = objTempItem.Index
      End If
    End If
  Next objTempItem

  ' Remove the selected items from the available listview
  For intTemp = UBound(iItemsToDelete) - 1 To 0 Step -1
    ListView1.ListItems.Remove iItemsToDelete(intTemp)
  Next intTemp

  ' Select the top available item in the listview
  If ListView1.ListItems.Count > 0 Then ListView1.ListItems(1).Selected = True

  If piBeforeIndex = 0 Then
    SelectLast ListView2
  Else
    For Each objTempItem In ListView2.ListItems
      objTempItem.Selected = (objTempItem.Index = piBeforeIndex)
    Next objTempItem
    Set ListView2.DropHighlight = Nothing
  End If

  RefreshCollectionSequence
  UpdateButtonStatus (SSTab1.Tab)
  Screen.MousePointer = vbDefault
  Changed = True

End Function


Public Property Get DefinitionOwner() As Boolean
  DefinitionOwner = mblnDefinitionCreator
End Property


Private Function SelectLast(lvwCtl As ListView)
  Dim objItem As ListItem
  
  For Each objItem In lvwCtl.ListItems
    objItem.Selected = IIf(objItem.Index = lvwCtl.ListItems.Count, True, False)
  Next objItem

End Function


Private Function SelectFirst(lvwCtl As ListView)
  Dim objItem As ListItem
  
  For Each objItem In lvwCtl.ListItems
    objItem.Selected = IIf(objItem.Index = 1, True, False)
  Next objItem

End Function


Private Function AddToCollection(objTempItem As ListItem) As Boolean
  Dim sColType As String
  Dim lID As Long
  Dim sHeading As String
  Dim lSize As Long
  Dim iDecPlaces As Integer
  
  mblnLoading = True
  GetDefaultDetails objTempItem.Key
  mblnLoading = False

  sColType = Left(objTempItem.Key, 1)
  lID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
  sHeading = txtProp_ColumnHeading.Text
  lSize = spnSize.Text
  iDecPlaces = IIf(spnDec.Text = "", "0", spnDec.Text)

  mcolRecordProfileColumnDetails.Add objTempItem.Key, sColType, lID, sHeading, lSize, iDecPlaces, miDataType, cboTblAvailable.ItemData(cboTblAvailable.ListIndex), objTempItem.Text, mcolRecordProfileColumnDetails.Count

End Function


Private Function GetDefaultDetails(sKey As String) As Boolean
  ' This function returns the default Column Name, Size and
  ' Decimal Places. These can then be edited by the user if desired.
  
  Dim rsTemp As Recordset

  Set rsTemp = datGeneral.GetColumnDefinition(Right(sKey, Len(sKey) - 1))
  
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    txtProp_ColumnHeading.Text = Replace(rsTemp!ColumnName, "_", " ")

    spnSize.Text = rsTemp!DefaultDisplayWidth
    
    miDataType = rsTemp!DataType
    If (rsTemp!DataType = sqlNumeric) Then ' its numeric
      spnDec.Text = rsTemp!Decimals
    ElseIf (rsTemp!DataType = sqlInteger) Then
      spnSize.Text = rsTemp!DefaultDisplayWidth ' 10 '5
      spnDec.Text = rsTemp!Decimals
    ElseIf rsTemp!DataType = sqlDate Then ' its a date
      spnSize.Text = rsTemp!DefaultDisplayWidth '10
      spnDec.Text = 0
    ElseIf rsTemp!DataType = sqlBoolean Then ' its a logic
      spnSize.Text = 1
    ElseIf rsTemp!DataType = sqlLongVarChar Then      ' working pattern field
      spnSize.Text = 14
    Else                                               ' its not
      spnDec.Text = 0
    End If
  End If

  If spnSize.Text = 0 Then
    If Len(rsTemp!SpinnerMaximum) > 0 Then spnSize.Text = Len(rsTemp!SpinnerMaximum)
  End If

  rsTemp.Close
  Set rsTemp = Nothing
  
End Function



Private Function GetTableIDFromColumn(plngColumnID As Long) As String
  Dim rsInfo As Recordset
  Dim sSQL As String
  
  sSQL = "SELECT ASRSysTables.TableID" & _
           " FROM ASRSysColumns JOIN ASRSysTables ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           " WHERE ColumnID = " & CStr(plngColumnID)

  Set rsInfo = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableIDFromColumn = rsInfo!TableID
  
  Set rsInfo = Nothing

End Function

Private Function GetTableNameFromColumn(plngColumnID As Long) As String

  Dim rsInfo As Recordset
  Dim sSQL As String
  
  sSQL = "SELECT ASRSysTables.TableName " & _
    " FROM ASRSysColumns JOIN ASRSysTables ON (ASRSysTables.TableID = ASRSysColumns.TableID)" & _
    " WHERE ColumnID = " & CStr(plngColumnID)

  Set rsInfo = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableNameFromColumn = rsInfo!TableName
  
  Set rsInfo = Nothing

End Function


Private Sub cmdAdd_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdAdd.Picture = cmdAdd.Picture
  
End Sub

Private Sub cmdAddAll_Click()
  ' Add All items from to the 'Selected' Listview
  CopyToSelected True

End Sub

Private Sub cmdAddAll_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdAddAll.Picture = cmdAddAll.Picture

End Sub


Private Sub cmdAddAllRelatedTables_Click()
  Dim sRow As String
  Dim pfrmTable As New frmRecordProfileTable
  Dim iTablesToAdd As Integer
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim fAlreadyAdded As Boolean
  Dim fCancelled As Boolean
  
  fCancelled = False
  
  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "PARENT"
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "CHILD"

  iTablesToAdd = UBound(mavTables, 2) - cboTblAvailable.ListCount + 1
  
  With pfrmTable
    .Initialize True, Me, 0, , 0, , 0, , 0, giHORIZONTAL, False, iTablesToAdd

    If Not .Cancelled Then .Show vbModal

    fCancelled = .Cancelled
    If Not fCancelled Then
      For iLoop = 1 To UBound(mavTables, 2)
        fAlreadyAdded = False
      
        For iLoop2 = 0 To cboTblAvailable.ListCount - 1
          If mavTables(1, iLoop) = cboTblAvailable.ItemData(iLoop2) Then
            fAlreadyAdded = True
            Exit For
          End If
        Next iLoop2
        
        If Not fAlreadyAdded Then
          sRow = mavTables(1, iLoop) _
            & vbTab & mavTables(2, iLoop) _
            & vbTab & 0 _
            & vbTab _
            & vbTab _
            & vbTab & IIf(mavTables(3, iLoop), "", sALL_RECORDS) _
            & vbTab & 0 _
            & vbTab & .PageBreak _
            & vbTab & IIf(.Orientation = giHORIZONTAL, "Horizontal", "Vertical") _
            & vbTab & .Orientation

          With grdRelatedTables
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .Bookmark
          End With
      
          Changed = True
        End If
      Next iLoop
    End If
  End With

  Unload pfrmTable
  Set pfrmTable = Nothing

  If Not fCancelled Then
    PopulateTableAvailable False

    EnableDisableTabControls
  
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
  UpdateButtonStatus (SSTab1.Tab)
  
  If Not fCancelled Then
    Changed = True
  End If

End Sub

Private Sub cmdAddAllTableColumns_Click()
  Dim iLoop As Integer
  
  With grdRelatedTables
    If .SelBookmarks.Count > 0 Then
      For iLoop = 0 To .SelBookmarks.Count - 1
        SelectAllColumns CLng(.Columns("TableID").CellText(.SelBookmarks(iLoop)))
      Next iLoop
    
      COAMsgBox "All columns added for the selected table" & IIf(.SelBookmarks.Count = 1, "", "s") & ".", vbInformation + vbOKOnly, "Record Profile"
    
      PopulateAvailable
      PopulateSelected
      UpdateButtonStatus (SSTab1.Tab)
    Else
      COAMsgBox "No tables selected.", vbExclamation + vbOKOnly, "Record Profile"
    End If
  End With

End Sub

Private Sub cmdAddHeading_Click()
  Dim sKey As String
  Dim lngID As Long
  
  sKey = UniqueKey(sTYPECODE_HEADING)
  
  ListView2.ListItems.Add , sKey, sDFLTTEXT_HEADING
  
  lngID = Right(sKey, Len(sKey) - 1)
  mcolRecordProfileColumnDetails.Add sKey, sTYPECODE_HEADING, lngID, sDFLTTEXT_HEADING, 0, 0, 0, cboTblAvailable.ItemData(cboTblAvailable.ListIndex), sDFLTTEXT_HEADING
  
  SelectLast ListView2

  RefreshCollectionSequence
  
  UpdateButtonStatus (SSTab1.Tab)
  Screen.MousePointer = vbDefault
  Changed = True

End Sub

Private Sub GetRelatedTables(plngTableID As Long, psRelationship As String)
  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
  Dim iLoop As Integer
  Dim fFound As Boolean
  
  If psRelationship = "CHILD" Then
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID" & _
      " WHERE ASRSysRelations.parentID = " & Trim(Str(plngTableID)) & _
      " ORDER BY ASRSysTables.tableName"
  Else
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.parentID" & _
      " WHERE ASRSysRelations.childID = " & Trim(Str(plngTableID)) & _
      " ORDER BY ASRSysTables.tableName"
  End If
  
  Set rsTables = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  Do While Not rsTables.EOF
    fFound = False
    
    For iLoop = 1 To UBound(mavTables, 2)
      If mavTables(1, iLoop) = rsTables!TableID Then
        fFound = True
        Exit For
      End If
    Next iLoop
    
    If fFound = False Then
      ReDim Preserve mavTables(3, UBound(mavTables, 2) + 1)
      mavTables(1, UBound(mavTables, 2)) = rsTables!TableID
      mavTables(2, UBound(mavTables, 2)) = rsTables!TableName
      mavTables(3, UBound(mavTables, 2)) = (psRelationship = "PARENT")
      
      GetRelatedTables rsTables!TableID, psRelationship
    End If
    
    rsTables.MoveNext
  Loop
  rsTables.Close
  Set rsTables = Nothing

End Sub



Private Sub cmdAddHeading_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdAddHeading.Picture = cmdAddHeading.Picture

End Sub

Private Sub cmdAddRelatedTable_Click()
  Dim sRow As String
  Dim pfrmTable As New frmRecordProfileTable
  Dim iLoop As Integer
  Dim fIsAscendant As Boolean
  
  With pfrmTable
    .Initialize True, Me, 0, , 0, , 0, , 0, giHORIZONTAL, False, 1

    If Not .Cancelled Then .Show vbModal

    If Not .Cancelled Then
      ' Get the tables related to the selected base table
      ' Put the table info into an array
      '   Column 1 = table ID
      '   Column 2 = table name
      '   Column 3 = true if this table is an ASCENDENT of the base table
      '            = false if this table is an DESCENDENT of the base table
      ReDim mavTables(3, 0)
      
      GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "PARENT"
      GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "CHILD"
      
      fIsAscendant = False
      For iLoop = 1 To UBound(mavTables, 2)
        If mavTables(1, iLoop) = .RelatedTableID Then
          fIsAscendant = mavTables(3, iLoop)
          Exit For
        End If
      Next iLoop
      
      sRow = .RelatedTableID _
        & vbTab & .RelatedTable _
        & vbTab & IIf(fIsAscendant, 0, .FilterID) _
        & vbTab & IIf(fIsAscendant, "", .Filter) _
        & vbTab & IIf(fIsAscendant, "", .Order) _
        & vbTab & IIf(fIsAscendant, "", IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)) _
        & vbTab & IIf(fIsAscendant, 0, .OrderID) _
        & vbTab & .PageBreak _
        & vbTab & IIf(.Orientation = giHORIZONTAL, "Horizontal", "Vertical") _
        & vbTab & .Orientation

      With grdRelatedTables
        .AddItem sRow
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With
      
      Changed = True
    End If
  End With

  Unload pfrmTable
  Set pfrmTable = Nothing

  PopulateTableAvailable False

  EnableDisableTabControls

  ForceDefinitionToBeHiddenIfNeeded

  UpdateButtonStatus (SSTab1.Tab)

  'AE20071005 Fault #8135
  'Changed = True

End Sub


Private Sub cmdAddSeparator_Click()
  Dim sKey As String
  Dim lngID As Long
  
  sKey = UniqueKey(sTYPECODE_SEPARATOR)
  
  ListView2.ListItems.Add , sKey, sDFLTTEXT_SEPARATOR
  
  lngID = Right(sKey, Len(sKey) - 1)
  mcolRecordProfileColumnDetails.Add sKey, sTYPECODE_SEPARATOR, lngID, "", 0, 0, 0, cboTblAvailable.ItemData(cboTblAvailable.ListIndex), sDFLTTEXT_SEPARATOR
  
  SelectLast ListView2

  RefreshCollectionSequence
  
  UpdateButtonStatus (SSTab1.Tab)
  Screen.MousePointer = vbDefault
  Changed = True

End Sub

Private Sub cmdAddSeparator_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdAddSeparator.Picture = cmdAddSeparator.Picture

End Sub


Private Sub cmdAutoArrangeRelatedTables_Click()
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iCurrentIndex As Integer
  
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "PARENT"
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "CHILD"
  
  grdRelatedTables.Redraw = False
  
  iCurrentIndex = 0
  
  For iLoop = 1 To UBound(mavTables, 2)
    For iLoop2 = 0 To grdRelatedTables.Rows - 1
      grdRelatedTables.Bookmark = grdRelatedTables.AddItemBookmark(iLoop2)
      If CLng(grdRelatedTables.Columns("TableID").Value) = mavTables(1, iLoop) Then
        If iLoop2 <> iCurrentIndex Then
          strSourceRow = grdRelatedTables.Columns(0).Text & vbTab & _
            grdRelatedTables.Columns(1).Text & vbTab & _
            grdRelatedTables.Columns(2).Text & vbTab & _
            grdRelatedTables.Columns(3).Text & vbTab & _
            grdRelatedTables.Columns(4).Text & vbTab & _
            grdRelatedTables.Columns(5).Text & vbTab & _
            grdRelatedTables.Columns(6).Text & vbTab & _
            grdRelatedTables.Columns(7).Text & vbTab & _
            grdRelatedTables.Columns(8).Text & vbTab & _
            grdRelatedTables.Columns(9).Text
  
          grdRelatedTables.Bookmark = grdRelatedTables.AddItemBookmark(iCurrentIndex)
          strDestinationRow = grdRelatedTables.Columns(0).Text & vbTab & _
            grdRelatedTables.Columns(1).Text & vbTab & _
            grdRelatedTables.Columns(2).Text & vbTab & _
            grdRelatedTables.Columns(3).Text & vbTab & _
            grdRelatedTables.Columns(4).Text & vbTab & _
            grdRelatedTables.Columns(5).Text & vbTab & _
            grdRelatedTables.Columns(6).Text & vbTab & _
            grdRelatedTables.Columns(7).Text & vbTab & _
            grdRelatedTables.Columns(8).Text & vbTab & _
            grdRelatedTables.Columns(9).Text
  
          grdRelatedTables.AddItem strSourceRow, iCurrentIndex
          grdRelatedTables.RemoveItem iLoop2 + 1
  
          Changed = True
        End If
        
        iCurrentIndex = iCurrentIndex + 1
        Exit For
      End If
    Next iLoop2
  Next iLoop
  
  ' Select the top row in the grid
  If grdRelatedTables.Rows > 0 Then
    grdRelatedTables.SelBookmarks.RemoveAll
    grdRelatedTables.MoveFirst
    grdRelatedTables.SelBookmarks.Add grdRelatedTables.Bookmark
  End If

  grdRelatedTables.Redraw = True
  RefreshTableAvailableOrder
  UpdateButtonStatus (SSTab1.Tab)

End Sub

Private Sub cmdBaseFilter_Click()
  GetFilter cboBaseTable, txtBaseFilter

End Sub


Private Sub GetFilter(ctlSource As Control, ctlTarget As Control)
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    If TypeOf ctlSource Is TextBox Then
      fOK = .Initialise(ctlSource.Tag, Val(ctlTarget.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    ElseIf TypeOf ctlSource Is ComboBox Then
      fOK = .Initialise(ctlSource.ItemData(ctlSource.ListIndex), Val(ctlTarget.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    End If
      
    If fOK Then
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) = True Then
        ' Read the selected expression info.
        ctlTarget.Text = IIf(Len(.Name) = 0, "<None>", .Name)
        ctlTarget.Tag = .ExpressionID
        
        Changed = True
      End If
    End If
  End With
  
  Set objExpression = Nothing

  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub GetOrder(ctlSource As Control, ctlTarget As Control)
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  On Error GoTo ErrorTrap

  Dim fExit As Boolean
  Dim lngTableID As Long
  Dim objOrder As clsOrder
  Dim sSQL As String
  Dim rsOrders As Recordset
  Dim fOK As Boolean

  fOK = True
  
  Screen.MousePointer = vbHourglass

  fExit = False

  If TypeOf ctlSource Is TextBox Then
    lngTableID = ctlSource.Tag
  Else
    lngTableID = ctlSource.ItemData(ctlSource.ListIndex)
  End If

  ' Instantiate an order object.
  Set objOrder = New clsOrder

  With objOrder
    ' Initialize the order object.
    .OrderID = Val(ctlTarget.Tag)
    .TableID = lngTableID
    .OrderType = giORDERTYPE_DYNAMIC
    
    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      ctlTarget.Text = .OrderName
      ctlTarget.Tag = .OrderID
      
      Changed = True
      fExit = True
    Else
      ' Check in case the original order has been deleted.
      sSQL = "SELECT *" & _
        " FROM ASRSysOrders" & _
        " WHERE orderID = " & Trim(Str(Val(ctlTarget.Tag)))
      Set rsOrders = datGeneral.GetRecords(sSQL)
      With rsOrders
        If (.EOF And .BOF) Then
          ctlTarget.Text = ""
          ctlTarget.Tag = 0
        End If

        .Close
      End With
      Set rsOrders = Nothing
    End If
  End With

TidyUpAndExit:
  Set objOrder = Nothing
  If Not fOK Then
    COAMsgBox "Error changing order ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub



Private Sub cmdBaseOrder_Click()
  GetOrder cboBaseTable, txtBaseOrder

End Sub

Private Sub cmdBasePicklist_Click()

  GetPicklist cboBaseTable, txtBasePicklist

End Sub

Private Sub GetPicklist(ctlSource As Control, ctlTarget As Control)

  Dim sSQL As String
  Dim lParent As Long
  Dim fExit As Boolean
  Dim frmPick As frmPicklists
  Dim lngTableID As Long
  Dim rsTemp As Recordset
  Dim blnHiddenPicklist As Boolean
  
  Screen.MousePointer = vbHourglass

  fExit = False

  If TypeOf ctlSource Is TextBox Then
    lngTableID = ctlSource.Tag
  Else
    lngTableID = ctlSource.ItemData(ctlSource.ListIndex)
  End If
  
  With frmDefSel
    .TableID = lngTableID
    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(ctlTarget.Tag) > 0 Then
      .SelectedID = Val(ctlTarget.Tag)
    End If
  End With

  'loop until a picklist has been selected or cancelled
  Do While Not fExit

    If frmDefSel.ShowList(utlPicklist) Then
      frmDefSel.Show vbModal
  
      Select Case frmDefSel.Action
        Case edtAdd
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(True, False, lngTableID) Then
              .Show vbModal
            End If
            frmDefSel.SelectedID = .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          End With
  
        Case edtEdit
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(False, frmDefSel.FromCopy, lngTableID, frmDefSel.SelectedID) Then
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
          frmPick.PrintDef frmDefSel.TableID, frmDefSel.SelectedID
          Unload frmPick
          Set frmPick = Nothing

        Case edtSelect
          ctlTarget.Text = IIf(Len(frmDefSel.SelectedText) = 0, "<None>", frmDefSel.SelectedText)
          ctlTarget.Tag = frmDefSel.SelectedID

          Changed = True
          fExit = True
          
        Case 0
          fExit = True
          
      End Select
    End If
  Loop

  Set frmDefSel = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub cmdCancel_Click()
  Unload Me

End Sub

Private Sub cmdEditRelatedTable_Click()
  Dim sRow As String
  Dim lngRow As Long
  Dim frmTable As New frmRecordProfileTable
  Dim lngInitTableID As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iCount2 As Integer
  Dim fNeedRefreshAvail As Boolean
  Dim aiSelectedRows() As Integer
  Dim fIsAscendant As Boolean
  Dim fCancelled As Boolean
  
  fCancelled = False
  
  ReDim aiSelectedRows(0)
  With grdRelatedTables
    For iLoop = 0 To .SelBookmarks.Count - 1
      ReDim aiSelectedRows(UBound(aiSelectedRows) + 1)
      aiSelectedRows(UBound(aiSelectedRows)) = .AddItemRowIndex(.SelBookmarks(iLoop))
    Next iLoop
    
    If .SelBookmarks.Count > 1 Then
      frmTable.Initialize False, Me, _
        0, , _
        0, , _
        0, , _
        0, _
        giHORIZONTAL, _
        False, _
        .SelBookmarks.Count
    Else
      .Bookmark = .SelBookmarks(0)
      lngRow = .AddItemRowIndex(.Bookmark)
      lngInitTableID = .Columns("TableID").Value
      frmTable.Initialize False, Me, _
        .Columns("TableID").Value, .Columns("Table").Value, _
        .Columns("FilterID").Value, .Columns("Filter").Value, _
        .Columns("OrderID").Value, .Columns("Order").Value, _
        IIf(.Columns("Records").Value = sALL_RECORDS, 0, Val(.Columns("Records").Value)), _
        .Columns("OrientationCode").Value, _
        .Columns("PageBreak").Value, 1
    End If
  End With

  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "PARENT"
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "CHILD"
      
  With frmTable
    .Show vbModal

    fCancelled = .Cancelled
    
    If Not fCancelled Then
      If grdRelatedTables.SelBookmarks.Count > 1 Then
        For iLoop = 0 To grdRelatedTables.SelBookmarks.Count - 1
          grdRelatedTables.Bookmark = grdRelatedTables.SelBookmarks(iLoop)
          
          grdRelatedTables.Columns("Orientation").Value = IIf(.Orientation = giHORIZONTAL, "Horizontal", "Vertical")
          grdRelatedTables.Columns("OrientationCode").Value = .Orientation
          grdRelatedTables.Columns("PageBreak").Value = .PageBreak
        Next iLoop
      Else
        fIsAscendant = False
        For iLoop = 1 To UBound(mavTables, 2)
          If mavTables(1, iLoop) = .RelatedTableID Then
            fIsAscendant = mavTables(3, iLoop)
            Exit For
          End If
        Next iLoop
        
        If .RelatedTableID <> lngInitTableID Then
          ' Check if any columns in the record profile definition are from the table that was
          ' previously selected in the related table combo box. If so, prompt user for action.
          Select Case AnyRelatedTableColumnsUsed(lngInitTableID)
            Case 2: ' related table cols used and user wants to continue with the change
              sRow = .RelatedTableID _
                & vbTab & .RelatedTable _
                & vbTab & IIf(fIsAscendant, 0, .FilterID) _
                & vbTab & IIf(fIsAscendant, "", .Filter) _
                & vbTab & IIf(fIsAscendant, "", .Order) _
                & vbTab & IIf(fIsAscendant, "", IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)) _
                & vbTab & IIf(fIsAscendant, 0, .OrderID) _
                & vbTab & .PageBreak _
                & vbTab & IIf(.Orientation = giHORIZONTAL, "Horizontal", "Vertical") _
                & vbTab & .Orientation

            Case 1: ' related tables cols used and user has aborted the change
              With grdRelatedTables
                .Bookmark = .AddItemBookmark(lngRow)
                .SelBookmarks.RemoveAll
                .SelBookmarks.Add .AddItemBookmark(lngRow)
              End With
    
              Exit Sub
            
            Case 0: ' no related table cols used
              sRow = .RelatedTableID _
                & vbTab & .RelatedTable _
                & vbTab & IIf(fIsAscendant, 0, .FilterID) _
                & vbTab & IIf(fIsAscendant, "", .Filter) _
                & vbTab & IIf(fIsAscendant, "", .Order) _
                & vbTab & IIf(fIsAscendant, "", IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)) _
                & vbTab & IIf(fIsAscendant, 0, .OrderID) _
                & vbTab & .PageBreak _
                & vbTab & IIf(.Orientation = giHORIZONTAL, "Horizontal", "Vertical") _
                & vbTab & .Orientation
          End Select
        Else
          sRow = .RelatedTableID _
            & vbTab & .RelatedTable _
            & vbTab & IIf(fIsAscendant, 0, .FilterID) _
            & vbTab & IIf(fIsAscendant, "", .Filter) _
            & vbTab & IIf(fIsAscendant, "", .Order) _
            & vbTab & IIf(fIsAscendant, "", IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)) _
            & vbTab & IIf(fIsAscendant, 0, .OrderID) _
            & vbTab & .PageBreak _
            & vbTab & IIf(.Orientation = giHORIZONTAL, "Horizontal", "Vertical") _
            & vbTab & .Orientation
        End If

        With grdRelatedTables
          'Find and remove from Table Available
          For iLoop = 0 To cboTblAvailable.ListCount - 1 Step 1
            If cboTblAvailable.ItemData(iLoop) = lngInitTableID Then
  
              If cboTblAvailable.ListIndex = iLoop Then
                fNeedRefreshAvail = True
              End If
  
              cboTblAvailable.RemoveItem iLoop
              Exit For
            End If
          Next iLoop
          .RemoveItem lngRow
          .AddItem sRow, lngRow
          .Bookmark = .AddItemBookmark(lngRow)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
        End With
      End If
    End If
  End With

  Unload frmTable
  Set frmTable = Nothing

  If Not fCancelled Then
    PopulateTableAvailable fNeedRefreshAvail
  
    EnableDisableTabControls
  
    ForceDefinitionToBeHiddenIfNeeded
  
    With grdRelatedTables
      If .SelBookmarks.Count <= 1 Then
        .SelBookmarks.RemoveAll
        For iLoop = 1 To UBound(aiSelectedRows)
          .SelBookmarks.Add .AddItemBookmark(aiSelectedRows(iLoop))
        Next iLoop
      End If
    End With
  
    UpdateButtonStatus (SSTab1.Tab)
  
    Changed = True
  End If
  
End Sub





Private Sub cmdMoveDown_Click()

  ChangeSelectedOrder ListView2.SelectedItem.Index + 2, True

End Sub

Private Function ChangeSelectedOrder(Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)
  ' This function changes the order of listitems in the selected listview.
  ' At the moment, different arrays are used depending on what information you
  ' need to store...change the array to a type if it would suit the purpose
  ' better
  
  ' Dimension arrays
  Dim iLoop As Integer
  Dim Key() As String
  Dim Text() As String
  Dim Icon() As Variant
  Dim SmallIcon() As Variant
  
  ReDim Key(0)
  ReDim Text(0)
  ReDim Icon(0)
  ReDim SmallIcon(0)
  
  ' Clear the highlight
  Set ListView2.DropHighlight = Nothing
  
  ' If drop point is below all other items, then fix the intbeforeindex var
  If intBeforeIndex = 0 Then intBeforeIndex = ListView2.ListItems.Count + 1
  
  ' First get all the items that are above the drop point that arent selected
  For iLoop = 1 To (intBeforeIndex - 1)
    If ListView2.ListItems(iLoop).Selected = False Then
      ReDim Preserve Key(UBound(Key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
      Key(UBound(Key) - 1) = ListView2.ListItems(iLoop).Key
      Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
    End If
  Next iLoop
  
  ' Now get all the items that are selected
  For iLoop = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(iLoop).Selected = True Then
      ReDim Preserve Key(UBound(Key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
      Key(UBound(Key) - 1) = ListView2.ListItems(iLoop).Key
      Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
    End If
  Next iLoop
  
  ' Now get all the items below the drop point that arent selected
  If intBeforeIndex <> 0 Then
    For iLoop = (intBeforeIndex) To ListView2.ListItems.Count
      If ListView2.ListItems(iLoop).Selected = False Then
        ReDim Preserve Key(UBound(Key) + 1)
        ReDim Preserve Text(UBound(Text) + 1)
        ReDim Preserve Icon(UBound(Icon) + 1)
        ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
        Key(UBound(Key) - 1) = ListView2.ListItems(iLoop).Key
        Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
        Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
        SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
      End If
    Next iLoop
  End If
  
  ' Clear all items from the listview
  ListView2.ListItems.Clear
  
  ' Add items in the right order from the array
  For iLoop = LBound(Key) To (UBound(Key) - 1)
    ListView2.ListItems.Add , Key(iLoop), Text(iLoop), Icon(iLoop), SmallIcon(iLoop)
  Next iLoop
  
  If mfFromButtons = True Then
    ListView2.ListItems(intBeforeIndex - 1).Selected = True
  Else
    If intBeforeIndex < ListView2.ListItems.Count Then ListView2.ListItems(intBeforeIndex).Selected = True Else ListView2.ListItems(ListView2.ListItems.Count).Selected = True
  End If
  
  mfFromButtons = False
  
  Changed = True
  
  UpdateButtonStatus (SSTab1.Tab)

  ' Remember the order the columns are now in.
  RefreshCollectionSequence
  
End Function



Private Sub cmdMoveDown_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdMoveDown.Picture = cmdMoveDown.Picture

End Sub

Private Sub cmdMoveRelatedTableDown_Click()
  ChangeRelatedTableOrder "DOWN"

End Sub

Private Sub cmdMoveRelatedTableDown_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdMoveRelatedTableDown.Picture = cmdMoveRelatedTableDown.Picture

End Sub


Private Sub cmdMoveRelatedTableUp_Click()
  ChangeRelatedTableOrder "UP"

End Sub


Private Sub ChangeRelatedTableOrder(psDirection As String)
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdRelatedTables.AddItemRowIndex(grdRelatedTables.Bookmark)
  strSourceRow = grdRelatedTables.Columns(0).Text & vbTab & _
    grdRelatedTables.Columns(1).Text & vbTab & _
    grdRelatedTables.Columns(2).Text & vbTab & _
    grdRelatedTables.Columns(3).Text & vbTab & _
    grdRelatedTables.Columns(4).Text & vbTab & _
    grdRelatedTables.Columns(5).Text & vbTab & _
    grdRelatedTables.Columns(6).Text & vbTab & _
    grdRelatedTables.Columns(7).Text & vbTab & _
    grdRelatedTables.Columns(8).Text & vbTab & _
    grdRelatedTables.Columns(9).Text
  
  If psDirection = "UP" Then
    intDestinationRow = intSourceRow - 1
    grdRelatedTables.MovePrevious
  Else
    intDestinationRow = intSourceRow + 1
    grdRelatedTables.MoveNext
  End If
  
  strDestinationRow = grdRelatedTables.Columns(0).Text & vbTab & _
    grdRelatedTables.Columns(1).Text & vbTab & _
    grdRelatedTables.Columns(2).Text & vbTab & _
    grdRelatedTables.Columns(3).Text & vbTab & _
    grdRelatedTables.Columns(4).Text & vbTab & _
    grdRelatedTables.Columns(5).Text & vbTab & _
    grdRelatedTables.Columns(6).Text & vbTab & _
    grdRelatedTables.Columns(7).Text & vbTab & _
    grdRelatedTables.Columns(8).Text & vbTab & _
    grdRelatedTables.Columns(9).Text
  
  If psDirection = "UP" Then
    grdRelatedTables.AddItem strSourceRow, intDestinationRow
    
    grdRelatedTables.RemoveItem intSourceRow + 1
    
    grdRelatedTables.SelBookmarks.RemoveAll
    grdRelatedTables.MovePrevious
  Else
    grdRelatedTables.RemoveItem intDestinationRow
    grdRelatedTables.RemoveItem intSourceRow
    
    grdRelatedTables.AddItem strDestinationRow, intSourceRow
    grdRelatedTables.AddItem strSourceRow, intDestinationRow
    
    grdRelatedTables.SelBookmarks.RemoveAll
    grdRelatedTables.MoveNext
  End If
  
  grdRelatedTables.Bookmark = grdRelatedTables.AddItemBookmark(intDestinationRow)
  grdRelatedTables.SelBookmarks.Add grdRelatedTables.AddItemBookmark(intDestinationRow)
  
  RefreshTableAvailableOrder
  
  UpdateButtonStatus (SSTab1.Tab)
  Changed = True

End Sub

Private Sub cmdMoveRelatedTableUp_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdMoveRelatedTableUp.Picture = cmdMoveRelatedTableUp.Picture

End Sub

Private Sub cmdMoveUp_Click()

  ChangeSelectedOrder ListView2.SelectedItem.Index - 1

End Sub

Private Sub cmdMoveUp_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdMoveUp.Picture = cmdMoveUp.Picture

End Sub


Private Sub cmdOK_Click()
  If Changed = True Then
    'NHRD24042002 Fault 3728 Switch on and off the hourglass when saving definitions
    Screen.MousePointer = vbHourglass
    
    'TM20020508 Fault 3839 - need to set the mouse pointer to default before exit sub is called.
    If Not ValidateDefinition(mlngRecordProfileID) Then
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    
    If Not SaveDefinition Then
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
  End If
  
  Me.Hide

End Sub

Private Function ValidateDefinition(lngCurrentID As Long) As Boolean

  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim strRecSelStatus As String
  Dim i As Integer
  Dim pvarbookmark As Variant
  Dim lngFilterID As Long
  Dim sRow As String
  Dim sTableName As String
  Dim objItem As clsRecordProfileColDtl
  Dim fFound As Boolean
  Dim asTableErrors() As String
  Dim sText As String
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  
  Dim iCount_Owner As Integer
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobScheduledUserGroups As String
  Dim sHiddenGroups As String
  
  ValidateDefinition = False
  fBatchJobsOK = True
  
  ' Check a name has been entered
  If Trim(txtName.Text) = "" Then
    COAMsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SetFocus
    Exit Function
  End If

  'Check if this definition has been changed by another user
  Call UtilityAmended(utlRecordProfile, mlngRecordProfileID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngRecordProfileID = 0
    mblnReadOnly = False
    ForceAccess
  End If

  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngRecordProfileID) Then
    COAMsgBox "A Record Profile definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    Exit Function
  End If

  ' BASE TABLE - If using a picklist, check one has been selected
  If optBasePicklist.Value Then
    If txtBasePicklist.Text = "" Or txtBasePicklist.Tag = "0" Or txtBasePicklist.Tag = "" Then
      COAMsgBox "You must select a picklist, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Record Profile"
      SSTab1.Tab = 0
      cmdBasePicklist.SetFocus
      Exit Function
    End If
  End If

  ' BASE TABLE - If using a filter, check one has been selected
  If optBaseFilter.Value Then
    If txtBaseFilter.Text = "" Or txtBaseFilter.Tag = "0" Or txtBaseFilter.Tag = "" Then
      COAMsgBox "You must select a filter, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Record Profile"
      SSTab1.Tab = 0
      cmdBaseFilter.SetFocus
      Exit Function
    End If
  End If

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
  
  ' Check that no grandchild or grandparent tables are included in the
  ' record profile without the related child or parent table also being included.
  ' Start at index 1 of the combo as index 0 is the base table, and so does not require this check.
  ReDim asTableErrors(0)
      
  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "PARENT"
  GetRelatedTables cboBaseTable.ItemData(cboBaseTable.ListIndex), "CHILD"
  
  For iLoop = 1 To cboTblAvailable.ListCount - 1
    ' See if the table is a descendant or ascendant of the base table.
    For iLoop2 = 1 To UBound(mavTables, 2)
      If mavTables(1, iLoop2) = cboTblAvailable.ItemData(iLoop) Then
        If Not LineageExists(cboTblAvailable.ItemData(iLoop), CBool(mavTables(3, iLoop2))) Then
          ReDim Preserve asTableErrors(UBound(asTableErrors) + 1)
          asTableErrors(UBound(asTableErrors)) = cboTblAvailable.List(iLoop)
        End If
        
        Exit For
      End If
    Next iLoop2
  Next iLoop
  
  If UBound(asTableErrors) = 1 Then
    COAMsgBox "You must include the tables that relate the '" & asTableErrors(1) & "' table to the base table in the record profile.", vbExclamation + vbOKOnly, "Record Profile"
    SSTab1.Tab = 1
    Exit Function
  ElseIf UBound(asTableErrors) > 0 Then
    sText = "You must include the tables that relate the following tables to the base table in the record profile:" & vbCrLf & vbCrLf
    
    For iLoop = 1 To UBound(asTableErrors)
      sText = sText & _
        vbTab & asTableErrors(iLoop) & vbCrLf
    Next iLoop
    
    COAMsgBox sText, vbExclamation + vbOKOnly, "Record Profile"
    SSTab1.Tab = 1
    Exit Function
  End If
  
  ' Check that there are columns defined in the record profile definition
  ReDim asTableErrors(0)
  
  For iLoop = 0 To cboTblAvailable.ListCount - 1
    fFound = False
    
    For Each objItem In mcolRecordProfileColumnDetails
      If objItem.TableID = cboTblAvailable.ItemData(iLoop) Then
        fFound = True
        Exit For
      End If
    Next objItem
    Set objItem = Nothing
    
    If Not fFound Then
      ReDim Preserve asTableErrors(UBound(asTableErrors) + 1)
      asTableErrors(UBound(asTableErrors)) = cboTblAvailable.List(iLoop)
    End If
  Next iLoop
  
  If UBound(asTableErrors) = 1 Then
    COAMsgBox "You must select at least 1 column for your record profile from the '" & asTableErrors(1) & "' table.", vbExclamation + vbOKOnly, "Record Profile"
    SSTab1.Tab = 2
    Exit Function
  ElseIf UBound(asTableErrors) > 0 Then
    sText = "You must select at least 1 column for your record profile from the following tables:" & vbCrLf & vbCrLf
    
    For iLoop = 1 To UBound(asTableErrors)
      sText = sText & _
        vbTab & asTableErrors(iLoop) & vbCrLf
    Next iLoop
    
    COAMsgBox sText, vbExclamation + vbOKOnly, "Record Profile"
    SSTab1.Tab = 2
    Exit Function
  End If

  If Not mobjOutputDef.ValidDestination Then
    SSTab1.Tab = 3
    Exit Function
  End If

If mlngRecordProfileID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlRecordProfile, _
      CStr(mlngRecordProfileID), _
      txtUserName.Text, _
      iCount_Owner, _
      sBatchJobDetails_Owner, _
      sBatchJobIDs, _
      sBatchJobDetails_NotOwner, _
      fBatchJobsOK, _
      sBatchJobDetails_ScheduledForOtherUsers, _
      sBatchJobScheduledUserGroups, _
      sHiddenGroups

    If (Not fBatchJobsOK) Then
      If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
        COAMsgBox "This definition cannot be made hidden from the following user groups :" & vbCrLf & vbCrLf & sBatchJobScheduledUserGroups & vbCrLf & _
               "as it is used in the following batch jobs which are scheduled to be run by these user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, _
               vbExclamation + vbOKOnly, "Record Profile"
      Else
        COAMsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Record Profile"
      End If

      Screen.MousePointer = vbDefault
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If COAMsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Record Profile") = vbNo Then
        Screen.MousePointer = vbDefault
        SSTab1.Tab = 0
        Exit Function
      Else
        ' Ok, we are continuing, so lets update all those utils to hidden !
        If Len(Trim(sBatchJobIDs)) > 0 Then
          HideUtilities utlBatchJob, sBatchJobIDs, sHiddenGroups
          Call UtilUpdateLastSavedMultiple(utlBatchJob, sBatchJobIDs)
        End If
      End If
    End If
  End If
End If

  ValidateDefinition = True

End Function


Private Function LineageExists(plngTableID As Long, _
  pfIsAscendant As Boolean) As Boolean
  ' Check if the given table's lineage to the record profile base table
  ' is via tables that are included in the record profile definition.
  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
  Dim iLoop As Integer
  Dim fExists As Boolean
  
  fExists = False
  
  If pfIsAscendant Then
    sSQL = "SELECT ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID" & _
      " WHERE ASRSysRelations.parentID = " & Trim(Str(plngTableID))
  Else
    sSQL = "SELECT ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.parentID" & _
      " WHERE ASRSysRelations.childID = " & Trim(Str(plngTableID))
  End If
    
  Set rsTables = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  Do While (Not rsTables.EOF) And (Not fExists)
    If rsTables!TableID = cboBaseTable.ItemData(cboBaseTable.ListIndex) Then
      ' Related table IS the base table. Yippee, we've struck gold !
      fExists = True
      Exit Do
    Else
      ' Related table is not the base table, so do the lineage check on this one (if its included in the record profile definition).
      For iLoop = 1 To cboTblAvailable.ListCount - 1
        If rsTables!TableID = cboTblAvailable.ItemData(iLoop) Then
          If LineageExists(rsTables!TableID, pfIsAscendant) Then
            fExists = True
            Exit For
          End If
        End If
      Next iLoop
    End If
    
    rsTables.MoveNext
  Loop
  rsTables.Close
  Set rsTables = Nothing

  LineageExists = fExists
  
End Function




Private Function SaveDefinition() As Boolean
  On Error GoTo Save_ERROR

  Dim sSQL As String
  Dim objItem As clsRecordProfileColDtl
  Dim iDefExportTo As Integer
  Dim iDefSave As Integer
  Dim sDefSaveAs As String
  Dim iDefCloseApp As Integer
  
  '########################### 1 Of 3 - SAVE THE BASIC DETAILS
  If mlngRecordProfileID > 0 Then
    ' Construct the SQL Update string (Editing an existing definition)

    sSQL = "UPDATE ASRSYSRecordProfileName SET" & _
      " Name = '" & Trim(Replace(txtName.Text, "'", "''")) & "'," & _
      " Description = '" & Replace(txtDesc.Text, "'", "''") & "'," & _
      " BaseTable = " & cboBaseTable.ItemData(cboBaseTable.ListIndex) & "," & _
      " AllRecords = " & IIf(optBaseAllRecords.Value, 1, 0) & "," & _
      " PicklistID = " & IIf(optBasePicklist.Value, txtBasePicklist.Tag, 0) & "," & _
      " FilterID = " & IIf(optBaseFilter.Value, txtBaseFilter.Tag, 0) & "," & _
      " OrderID = " & txtBaseOrder.Tag & "," & _
      " Orientation = " & IIf(optBaseOrientation(0).Value, giHORIZONTAL, giVERTICAL) & "," & _
      " PageBreak = " & IIf(chkBasePageBreak.Value = vbChecked, 1, 0) & "," & _
      " PrintFilterHeader = " & IIf(chkPrintFilterHeader.Value = vbChecked, 1, 0) & ","

    sSQL = sSQL & _
        " OutputPreview = " & IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        " OutputFormat = " & CStr(mobjOutputDef.GetSelectedFormatIndex) & ", " & _
        " OutputScreen = " & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", " & _
        " OutputPrinter = " & IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0") & ", " & _
        " OutputPrinterName = '" & Replace(cboPrinterName.Text, " '", "''") & "', "
        
    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & _
        "OutputSave = 1, " & _
        "OutputSaveExisting = " & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
    Else
      sSQL = sSQL & _
        "OutputSave = 0, " & _
        "OutputSaveExisting = 0, "
    End If
        
    If chkDestination(desEmail).Value = vbChecked Then
      sSQL = sSQL & _
          "OutputEmail = 1, " & _
          "OutputEmailAddr = " & txtEmailGroup.Tag & ", " & _
          "OutputEmailSubject = '" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "OutputEmailAttachAs = '" & Replace(txtEmailAttachAs.Text, "'", "''") & "', "
    Else
      sSQL = sSQL & _
          "OutputEmail = 0, " & _
          "OutputEmailAddr = 0, " & _
          "OutputEmailSubject = '', " & _
          "OutputEmailAttachAs = '', "
    End If
    
    sSQL = sSQL & _
        "OutputFilename = '" & Replace(txtFilename.Text, "'", "''") & "',"

    sSQL = sSQL & _
      " IndentRelatedTables = " & IIf(chkIndent.Value = vbChecked, 1, 0) & "," & _
      " SuppressEmptyRelatedTableTitles = " & IIf(chkSuppressEmptyRelatedTableTitles.Value = vbChecked, 1, 0) & "," & _
      " SuppressTableRelationshipTitles = " & IIf(chkShowTableRelationshipTitle.Value = vbChecked, 0, 1)
      
    sSQL = sSQL & " WHERE recordProfileID = " & mlngRecordProfileID
    
'''''    If IsRecordSelectionValid = False Then
'''''      SaveDefinition = False
'''''      Exit Function
'''''    End If

    datData.ExecuteSql (sSQL)

    Call UtilUpdateLastSaved(utlRecordProfile, mlngRecordProfileID)
  Else
    ' Construct the SQL Insert string (Adding a new definition)
    sSQL = "INSERT INTO ASRSYSRecordProfileName (" & _
      "Name, Description, BaseTable, " & _
      "AllRecords, PicklistID, FilterID, OrderID, " & _
      "UserName," & _
      "Orientation, PageBreak, PrintFilterHeader," & _
      "OutputPreview, OutputFormat, OutputScreen, OutputPrinter, " & _
      "OutputPrinterName, OutputSave, OutputSaveExisting, OutputEmail, " & _
      "OutputEmailAddr, OutputEmailSubject, OutputEmailAttachAs, OutputFileName, " & _
      "IndentRelatedTables, SuppressEmptyRelatedTableTitles, SuppressTableRelationshipTitles)"
'      "Access, UserName," & _

    sSQL = sSQL & _
      " VALUES ('" & _
      Trim(Replace(txtName.Text, "'", "''")) & "','" & _
      Replace(txtDesc.Text, "'", "''") & "'," & _
      cboBaseTable.ItemData(cboBaseTable.ListIndex)

    If optBaseAllRecords Then
      sSQL = sSQL & ", 1, 0, 0"
    ElseIf optBasePicklist Then
      sSQL = sSQL & ", 0, " & Val(txtBasePicklist.Tag) & ", 0"
    Else
      sSQL = sSQL & ", 0, 0, " & Val(txtBaseFilter.Tag)
    End If

    sSQL = sSQL & ", " & Val(txtBaseOrder.Tag)

    sSQL = sSQL & ", '" & datGeneral.UserNameForSQL & "',"
    sSQL = sSQL & CStr(IIf(optBaseOrientation(0).Value, giHORIZONTAL, giVERTICAL)) & ","
    sSQL = sSQL & CStr(IIf(chkBasePageBreak.Value = vbChecked, 1, 0)) & ","
    sSQL = sSQL & CStr(IIf(chkPrintFilterHeader.Value = vbChecked, 1, 0)) & ","
    
    'Output Options
    sSQL = sSQL & CStr(IIf(chkPreview.Value = vbChecked, "1", "0")) & ","  'OutputPreview
    sSQL = sSQL & CStr(mobjOutputDef.GetSelectedFormatIndex) & ","           'OutputFormat
    sSQL = sSQL & CStr(IIf(chkDestination(desScreen).Value = vbChecked, "1", "0")) & ","  'OutputScreen
    sSQL = sSQL & CStr(IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0")) & "," 'OutputPrinter
    sSQL = sSQL & "'" & Replace(cboPrinterName.Text, "'", "''") & "',"              'OutputPrinterName

    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & "1, " & _
        cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "    'OutputSave, OutputSaveExisting
    Else
      sSQL = sSQL & "0, 0, "    'OutputSave, OutputSaveExisting
    End If

    If chkDestination(desEmail).Value = vbChecked Then
      sSQL = sSQL & "1, " & _
          txtEmailGroup.Tag & ", " & _
          "'" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "'" & Replace(txtEmailAttachAs.Text, "'", "''") & "', "      'OutputEmail, OutputEmailAddr, OutputEmailSubject
    Else
      sSQL = sSQL & "0, 0, '', '', "   'OutputEmail, OutputEmailAddr, OutputEmailSubject
    End If

    sSQL = sSQL & _
        "'" & Replace(txtFilename.Text, "'", "''") & "',"  'OutputFilename
        
    sSQL = sSQL & _
      IIf(chkIndent.Value = vbChecked, "1", "0") & "," & _
      IIf(chkSuppressEmptyRelatedTableTitles.Value = vbChecked, "1", "0") & "," & _
      IIf(chkShowTableRelationshipTitle.Value = vbChecked, "0", "1")
        
    sSQL = sSQL & ")"

    If IsRecordSelectionValid = False Then
      SaveDefinition = False
      Exit Function
    End If

    mlngRecordProfileID = InsertRecordProfile(sSQL)

    Call UtilCreated(utlRecordProfile, mlngRecordProfileID)
  End If

'########################### 2 Of 3 - SAVE THE RELATED TABLE DETAILS
' First, remove any records from the relate table detail tables.
SaveAccess

  '########################### 2 Of 3 - SAVE THE RELATED TABLE DETAILS
  ' First, remove any records from the relate table detail tables.
  ClearRelatedTables mlngRecordProfileID
  InsertRelatedTableDetails

  '########################### 3 Of 3 - SAVE THE COLUMN DETAILS
  ' First, remove any records from the 2 detail tables.
  ClearDetailTables mlngRecordProfileID

  For Each objItem In mcolRecordProfileColumnDetails
    sSQL = "INSERT INTO ASRSysRecordProfileDetails (" & _
      "recordProfileID, Sequence, Type, " & _
      "columnID, Heading, Size, DP, " & _
      "IsNumeric, tableID) "

    sSQL = sSQL & _
      "VALUES(" & _
       mlngRecordProfileID & "," & _
       objItem.Sequence & "," & _
       "'" & objItem.ColType & "'," & _
       objItem.ID & "," & _
       "'" & Replace(objItem.Heading, "'", "''") & "'," & _
       objItem.Size & "," & _
       objItem.DecPlaces & "," & _
       "0," & _
       objItem.TableID & ")"

    datData.ExecuteSql (sSQL)
  Next objItem

  Set objItem = Nothing

  SaveDefinition = True
  Changed = False

  Exit Function

Save_ERROR:

  SaveDefinition = False
  COAMsgBox "Warning : An error has occurred whilst saving..." & vbCrLf & Err.Description & vbCrLf & "Please cancel and try again. If this error continues, delete the definition.", vbCritical + vbOKOnly, "Record Profile"

End Function



Private Sub cmdRemove_Click()

  CopyToAvailable False
  If ListView2.ListItems.Count = 0 Then EnableColProperties False

End Sub

Private Function CopyToAvailable(bAll As Boolean, Optional intBeforeIndex As Integer)
  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  
  ' Dont add the to the first listview...just remove em and
  ' repopulate the available listview...much quicker
  
  Screen.MousePointer = vbHourglass
  
  For iLoop = ListView2.ListItems.Count To 1 Step -1
    If Not bAll Then
      If ListView2.ListItems(iLoop).Selected Then
        iTempItemIndex = iLoop
        RemoveFromCollection ListView2.ListItems(iLoop).Key
        ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
      End If
    Else
      RemoveFromCollection ListView2.ListItems(iLoop).Key
      ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
    End If
  Next iLoop
  
  If ListView2.ListItems.Count > 0 Then
    If iTempItemIndex > ListView2.ListItems.Count Then iTempItemIndex = ListView2.ListItems.Count
    If iTempItemIndex > 0 Then ListView2.ListItems(iTempItemIndex).Selected = True
  End If
  
  PopulateAvailable
  
  UpdateButtonStatus (SSTab1.Tab)

  Changed = True

  Screen.MousePointer = vbDefault

End Function



Private Sub cmdRemove_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdRemove.Picture = cmdRemove.Picture

End Sub

Private Sub cmdRemoveAll_Click()
  ' Remove All items from the 'Selected' Listview
  If COAMsgBox("Are you sure you wish to remove all of the selected table's columns from this definition ?", vbYesNo + vbQuestion, "Record Profile") = vbYes Then
    CopyToAvailable True
    EnableColProperties False
  End If

End Sub

Private Sub cmdRemoveAll_LostFocus()
  'JPD 20031013 Fault 5498 & fault 5500
  cmdRemoveAll.Picture = cmdRemoveAll.Picture

End Sub


Private Sub cmdRemoveAllRelatedTables_Click()
  Dim i As Integer
  Dim i2 As Integer
  Dim pvarbookmark As Variant
  Dim bContinueRemoval As Boolean
  Dim lngSelectedTableID As Long
  Dim lngRowCount As Long
  Dim bRemovedFromAvailable As Boolean
  Dim lRow As Long
  Dim bNeedRefreshAvail As Boolean
  
  bContinueRemoval = (COAMsgBox("Removing all the related tables will remove all related table columns " & _
                            "included in the record profile definition. " & vbCrLf & _
                            "Do you wish to continue ?" _
                            , vbYesNo + vbQuestion, "Record Profile") = vbYes)

  If Not bContinueRemoval Then Exit Sub

  With grdRelatedTables
    lngRowCount = .Rows
    For i = 0 To lngRowCount - 1 Step 1
      .MoveFirst
      pvarbookmark = .GetBookmark(0)
      lRow = .AddItemRowIndex(pvarbookmark)
      
      lngSelectedTableID = .Columns("TableID").CellValue(pvarbookmark)
      bRemovedFromAvailable = False
      'Find and remove from Table Available
      For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
        If Not bRemovedFromAvailable Then
          If cboTblAvailable.ItemData(i2) = lngSelectedTableID Then
          
            If cboTblAvailable.ListIndex = i2 Then
              bNeedRefreshAvail = True
            End If

            cboTblAvailable.RemoveItem i2
            bRemovedFromAvailable = True
          End If
        End If
      Next i2

      AnyRelatedTableColumnsUsed lngSelectedTableID, bContinueRemoval
      .RemoveItem lRow
    Next i
    .RemoveAll
  End With
  
  If bNeedRefreshAvail Then
    PopulateTableAvailable bNeedRefreshAvail
  Else
    'TM20020424 Fault 3715
    'Using PopulateTableAvailable slows down the addition of the child table,
    'so do the stuff we need to do here.
    DoEvents
    ' If theres only 1 table, then disable the combo, otherwise enable it
    If cboTblAvailable.ListCount = 1 Then
      cboTblAvailable.Enabled = False
      cboTblAvailable.BackColor = &H8000000F
    Else
      cboTblAvailable.Enabled = True
      cboTblAvailable.BackColor = &H80000005
    End If
  End If
  
  EnableDisableTabControls

  ForceDefinitionToBeHiddenIfNeeded

  UpdateButtonStatus (SSTab1.Tab)
  
  Changed = True

End Sub


Private Function AnyRelatedTableColumnsUsed(lngTableID As Long, Optional bAutoYes As Boolean, Optional pfJustChecking As Boolean) As Integer
  ' This sub checks if any columns from the Child table which has just
  ' been deselected have been used in the current record profile definition.
  ' If so, user is prompted if they wish to continue. Continuing will
  ' delete the columns in the record profile from the old Child table.

  ' Return value  = 0 if no columns from this table are used in the record profile
  '               = 1 if columns from this table ARE used in the record profile, and the
  '                 user wishes to leave the table in the record profile
  '               = 2 if columns from this table ARE used in the record profile, and the
  '                 user still wishes to remove the table from the record profile

  Dim objRecordProfileColumn As clsRecordProfileColDtl
  Dim fUsed As Boolean
  
  fUsed = False
  
  For Each objRecordProfileColumn In mcolRecordProfileColumnDetails
    If objRecordProfileColumn.TableID = lngTableID Then
      fUsed = True
    End If
  Next objRecordProfileColumn
  Set objRecordProfileColumn = Nothing
  
  If Not fUsed Then
    AnyRelatedTableColumnsUsed = 0
    Exit Function
  End If
  
  If (Not bAutoYes) And (Not pfJustChecking) Then
    If COAMsgBox("One or more columns from the '" & datGeneral.GetTableName(lngTableID) & "' table have been included in the current record profile definition." & vbCrLf & _
              "Changing the table will remove these columns from the record profile definition." & vbCrLf & _
              "Do you wish to continue ?" _
              , vbYesNo + vbQuestion, "Record Profile") = vbNo Then
      AnyRelatedTableColumnsUsed = 1
      Exit Function
    End If
  End If
    
  If Not pfJustChecking Then
    ' Remove the table's columns from the definition.
    For Each objRecordProfileColumn In mcolRecordProfileColumnDetails
      If objRecordProfileColumn.TableID = lngTableID Then
        mcolRecordProfileColumnDetails.Remove objRecordProfileColumn.ColType & objRecordProfileColumn.ID
      End If
    Next objRecordProfileColumn
    Set objRecordProfileColumn = Nothing
  End If
  
  AnyRelatedTableColumnsUsed = 2
  
End Function


Private Sub cmdRemoveRelatedTable_Click()
  Dim i2 As Integer
  Dim lRow As Long
  Dim lngSelectedRelatedTableID As Long
  Dim bNeedRefreshAvail As Boolean
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim alngTables() As Long
  Dim varBookmark As Variant
  Dim sTablesWithSelectedColumns As String
  
  ' Column 1  = Table ID
  ' Column 2  = 0 if no columns from this table are used in the record profile
  '           = 1 if columns from this table ARE used in the record profile, and the
  '             user wishes to leave the table in the record profile
  '           = 2 if columns from this table ARE used in the record profile, and the
  '             user still wishes to remove the table from the record profile
  ReDim alngTables(2, 0)
  
  sTablesWithSelectedColumns = ""
  
  With grdRelatedTables
    For iCount = 0 To .SelBookmarks.Count - 1
      .Bookmark = .SelBookmarks(iCount)
      lngSelectedRelatedTableID = .Columns("TableID").Value
     
      ReDim Preserve alngTables(2, UBound(alngTables, 2) + 1)
      alngTables(1, UBound(alngTables, 2)) = lngSelectedRelatedTableID
      alngTables(2, UBound(alngTables, 2)) = AnyRelatedTableColumnsUsed(lngSelectedRelatedTableID, , True)
      
      If alngTables(2, UBound(alngTables, 2)) = 2 Then
        sTablesWithSelectedColumns = sTablesWithSelectedColumns & _
          vbTab & datGeneral.GetTableName(lngSelectedRelatedTableID) & vbCrLf
      End If
    Next iCount
      
    If Len(sTablesWithSelectedColumns) > 0 Then
      If COAMsgBox("The following tables have one or more columns included in the current record profile definition." & vbCrLf & vbCrLf & _
        sTablesWithSelectedColumns & vbCrLf & _
        "Removing these tables will remove these columns from the record profile definition." & vbCrLf & _
        "Do you wish to continue ?" _
        , vbYesNo + vbQuestion, "Record Profile") = vbNo Then
        
        Exit Sub
      End If
    End If
    
    For iCount = 1 To UBound(alngTables, 2)
      If alngTables(2, iCount) = 2 Then
        AnyRelatedTableColumnsUsed alngTables(1, iCount), True
      End If
      
      'Find and remove from Table Available
      For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
        If cboTblAvailable.ItemData(i2) = alngTables(1, iCount) Then
        
          If cboTblAvailable.ListIndex = i2 Then
            bNeedRefreshAvail = True
          End If

          cboTblAvailable.RemoveItem i2
          Exit For
        End If
      Next i2
        
      If .Rows = 1 Then
        .RemoveAll
      Else
        .MoveFirst
        For iCount2 = 0 To .Rows - 1
          .Bookmark = .AddItemBookmark(iCount2)
          If .Columns("TableID").Value = alngTables(1, iCount) Then
            .RemoveItem iCount2
            Exit For
          End If
        Next iCount2
      End If
    Next iCount
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
    
  If bNeedRefreshAvail Then
    PopulateTableAvailable bNeedRefreshAvail
  Else
    'TM20020424 Fault 3715
    'Using PopulateTableAvailable slows down the addition of the child table,
    'so do the stuff we need to do here.
    DoEvents
    ' If theres only 1 table, then disable the combo, otherwise enable it
    If cboTblAvailable.ListCount = 1 Then
      cboTblAvailable.Enabled = False
      cboTblAvailable.BackColor = &H8000000F
    Else
      cboTblAvailable.Enabled = True
      cboTblAvailable.BackColor = &H80000005
    End If
  End If
  
  EnableDisableTabControls

  ForceDefinitionToBeHiddenIfNeeded

  UpdateButtonStatus (SSTab1.Tab)
  
  Changed = True

End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture

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
  ' Instantiate collection class
  Set mcolRecordProfileColumnDetails = New clsRecordProfileColDtls
  SSTab1.Tab = 0
  
  grdRelatedTables.RowHeight = 239
  grdAccess.RowHeight = 239
  
  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos True, True, True
  mobjOutputDef.ShowFormats True, False, True, True, True, False, False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim pintAnswer As Integer
    
  If Changed = True And Not FormPrint Then
    pintAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, "Record Profile")
      
    If pintAnswer = vbYes Then
      cmdOK_Click
      Cancel = True
      Exit Sub
    ElseIf pintAnswer = vbCancel Then
      Cancel = True
      Exit Sub
    End If
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Frafieldsavailable_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture

End Sub


Private Sub Frafieldsselected_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture

End Sub


Private Sub grdAccess_ComboCloseUp()
  Changed = True
  
  'JPD 20030728 Fault 6486
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) And _
    (Len(grdAccess.Columns("Access").Text) > 0) Then
    ' The 'All Groups' access has changed. Apply the selection to all other groups.
    ForceAccess AccessCode(grdAccess.Columns("Access").Text)
    
    grdAccess.MoveFirst
    grdAccess.Col = 1
  End If

End Sub

Private Sub grdAccess_GotFocus()
  grdAccess.Col = 1

End Sub


Private Sub grdAccess_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  Dim varBkmk As Variant

  'JPD 20030728 Fault 6486
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) Then
    grdAccess.Columns("Access").Text = ""
  End If

  With grdAccess
    varBkmk = .SelBookmarks(0)

    If ((Not mblnDefinitionCreator) Or mblnReadOnly Or mblnForceHidden) Or _
      (.Columns("SysSecMgr").CellText(varBkmk) = "1") Then
      .Columns("Access").Style = ssStyleEdit
    Else
      .Columns("Access").Style = ssStyleComboBox
      .Columns("Access").RemoveAll
      .Columns("Access").AddItem AccessDescription(ACCESS_READWRITE)
      .Columns("Access").AddItem AccessDescription(ACCESS_READONLY)
      .Columns("Access").AddItem AccessDescription(ACCESS_HIDDEN)
    End If
  End With

  If Me.ActiveControl Is grdAccess Then
    grdAccess.Col = 1
  End If

End Sub


Private Sub grdAccess_RowLoaded(ByVal Bookmark As Variant)
  With grdAccess
    If (Not mblnDefinitionCreator) Or mblnReadOnly Or mblnForceHidden Then
      .Columns("GroupName").CellStyleSet "ReadOnly"
      .Columns("Access").CellStyleSet "ReadOnly"
      .ForeColor = vbGrayText
    ElseIf (.Columns("SysSecMgr").CellText(Bookmark) = "1") Then
      .Columns("GroupName").CellStyleSet "SysSecMgr"
      .Columns("Access").CellStyleSet "SysSecMgr"
      .ForeColor = vbWindowText
    Else
      .ForeColor = vbWindowText
    End If
  End With

End Sub


Private Sub grdRelatedTables_DblClick()
  If Not mblnReadOnly Then
    If grdRelatedTables.Rows > 0 Then
      cmdEditRelatedTable_Click
    Else
      cmdAddRelatedTable_Click
    End If
  End If

End Sub

Private Sub grdRelatedTables_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  UpdateButtonStatus (SSTab1.Tab)

End Sub

Private Sub grdRelatedTables_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  If Not cmdEditRelatedTable.Enabled Then
    UpdateButtonStatus (SSTab1.Tab)
  End If

End Sub

Private Sub ListView1_DblClick()
  If mblnReadOnly Then
    Exit Sub
  End If

  ' Copy the item doubleclicked on to the 'Selected' Listview
  CopyToSelected False

End Sub


Private Sub ListView1_DragDrop(Source As Control, X As Single, Y As Single)
  ' Perform the drop operation
  If Source Is ListView2 Then
    CopyToAvailable False
    ListView2.Drag vbCancel
  Else
    ListView2.Drag vbCancel
  End If

End Sub


Private Sub ListView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument.Picture
  End If

End Sub


Private Sub ListView1_GotFocus()
  'JPD 20030912 Fault 5781
  If cmdAdd.Enabled Then
    cmdAdd.Default = True
  End If
  
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Capture the Ctl-A event
  
  Dim objItem As ListItem
  
  If Shift = vbCtrlMask And KeyCode = 65 Then
    For Each objItem In ListView1.ListItems
      objItem.Selected = True
    Next objItem
    Set objItem = Nothing
  ElseIf KeyCode = vbKeyReturn Then
    If cmdAdd.Enabled Then
      cmdAdd_Click
      ListView1.SetFocus
    End If
  End If

End Sub


Private Sub ListView1_LostFocus()
  'JPD 20030912 Fault 5781
  cmdOK.Default = True

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mblnReadOnly Then
    Exit Sub
  End If
  
  ' SUB COMPLETED 28/01/00
  ' Start the drag operation
  Dim objItem As ComctlLib.ListItem

  If Button = vbLeftButton Then
    If ListView1.ListItems.Count > 0 Then
      mblnColumnDrag = True
      ListView1.Drag vbBeginDrag
    End If
  End If

End Sub


Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mblnReadOnly Then
    Exit Sub
  End If
  
    ' Popup menu on right button.
  If Button = vbRightButton Then
    
    With ActiveBar1.Bands("ColSelPopup")
      ' Enable/disable the required tools.
      .Tools("ID_Add").Enabled = ((Not mblnReadOnly) And (cmdAdd.Enabled))
      .Tools("ID_AddAll").Enabled = ((Not mblnReadOnly) And (cmdAddAll.Enabled))
      .Tools("ID_AddHeading").Visible = (cmdAddHeading.Visible)
      .Tools("ID_AddHeading").Enabled = ((Not mblnReadOnly) And (cmdAddHeading.Enabled))
      .Tools("ID_AddSeparator").Visible = (cmdAddSeparator.Visible)
      .Tools("ID_AddSeparator").Enabled = ((Not mblnReadOnly) And (cmdAddSeparator.Enabled))
      .Tools("ID_Remove").Enabled = False
      .Tools("ID_RemoveAll").Enabled = False
      .Tools("ID_MoveUp").Visible = (cmdMoveUp.Visible)
      .Tools("ID_MoveUp").Enabled = False
      .Tools("ID_MoveDown").Visible = (cmdMoveDown.Visible)
      .Tools("ID_MoveDown").Enabled = False

      .TrackPopup -1, -1
    End With
  
  Else
    ' If we are dragging, from the Available listview then end drag operation
    If mblnColumnDrag Then
      ListView1.Drag vbCancel
      mblnColumnDrag = False
    End If
    
  End If
  
End Sub


Private Sub ListView2_DblClick()
  If mblnReadOnly Then
    Exit Sub
  End If
  
  ' Remove the item doubleclicked on from the 'Selected' Listview
  CopyToAvailable False

End Sub


Private Sub ListView2_DragDrop(Source As Control, X As Single, Y As Single)
  ' Perform the drop operation - action depends on source and destination
  
  If Source Is ListView1 Then
    If ListView2.HitTest(X, Y) Is Nothing Then
      CopyToSelected False
    Else
      CopyToSelected False, ListView2.HitTest(X, Y).Index
    End If
    ListView1.Drag vbCancel
  Else
    If ListView2.HitTest(X, Y) Is Nothing Then
      ChangeSelectedOrder
    Else
      ChangeSelectedOrder ListView2.HitTest(X, Y).Index
    End If
    ListView2.Drag vbCancel
  End If

End Sub


Private Sub ListView2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument.Picture
  End If

  ' Set DropHighlight to the mouse's coordinates.
  Set ListView2.DropHighlight = ListView2.HitTest(X, Y)

End Sub


Private Sub ListView2_GotFocus()
  'JPD 20030912 Fault 5781
  If cmdRemove.Enabled Then
    cmdRemove.Default = True
  End If

End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
  UpdateButtonStatus (SSTab1.Tab)

End Sub


Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim objItem As ListItem
  
  If Shift = vbCtrlMask And KeyCode = 65 Then
    For Each objItem In ListView2.ListItems
      objItem.Selected = True
    Next objItem
    Set objItem = Nothing
    UpdateButtonStatus (SSTab1.Tab)
  ElseIf KeyCode = vbKeyReturn Then
    If cmdRemove.Enabled Then
      cmdRemove_Click
      ListView2.SetFocus
    End If
  End If

End Sub


Private Sub ListView2_LostFocus()
  'JPD 20030912 Fault 5781
  cmdOK.Default = True

End Sub

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Start the drag operation
  Dim objItem As ComctlLib.ListItem

  If mblnReadOnly Then
    Exit Sub
  End If
  
  If Button = vbLeftButton Then
    If ListView2.ListItems.Count > 0 Then
      mblnColumnDrag = True
      ListView2.Drag vbBeginDrag
    End If
  End If

End Sub


Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mblnReadOnly Then
    Exit Sub
  End If
  
    ' Popup menu on right button.
  If Button = vbRightButton Then
    
    With ActiveBar1.Bands("ColSelPopup")
      ' Enable/disable the required tools.
      .Tools("ID_Add").Enabled = False
      .Tools("ID_AddAll").Enabled = False
      .Tools("ID_AddHeading").Visible = (cmdAddHeading.Visible)
      .Tools("ID_AddHeading").Enabled = False
      .Tools("ID_AddSeparator").Visible = (cmdAddSeparator.Visible)
      .Tools("ID_AddSeparator").Enabled = False
      .Tools("ID_Remove").Enabled = ((Not mblnReadOnly) And (cmdRemove.Enabled))
      .Tools("ID_RemoveAll").Enabled = ((Not mblnReadOnly) And (cmdRemoveAll.Enabled))
      .Tools("ID_MoveUp").Visible = (cmdMoveUp.Visible)
      .Tools("ID_MoveUp").Enabled = ((Not mblnReadOnly) And (cmdMoveUp.Enabled))
      .Tools("ID_MoveDown").Visible = (cmdMoveDown.Visible)
      .Tools("ID_MoveDown").Enabled = ((Not mblnReadOnly) And (cmdMoveDown.Enabled))
      
      .TrackPopup -1, -1
    End With
  
  Else
    ' If we are dragging, from the Available listview then end drag operation
    If mblnColumnDrag Then
      ListView2.Drag vbCancel
      mblnColumnDrag = False
    End If
    
  End If
  
End Sub


Private Sub optBaseAllRecords_Click()
  
  Changed = True

  With txtBasePicklist
    .Text = ""
    .Tag = 0
  End With
  
  cmdBasePicklist.Enabled = False

  With txtBaseFilter
    .Text = ""
    .Tag = 0
  End With

  chkPrintFilterHeader.Value = vbUnchecked
  cmdBaseFilter.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded
  EnableDisableTabControls
  
End Sub

Private Sub optBaseFilter_Click()

  cmdBaseFilter.Enabled = True

  With txtBasePicklist
    .Text = ""
    .Tag = 0
  End With

  txtBaseFilter.Text = "<None>"
  cmdBasePicklist.Enabled = False
  
  If Not mblnLoading Then
    Changed = True
    ForceDefinitionToBeHiddenIfNeeded
    EnableDisableTabControls
  End If
  
End Sub


Private Sub optBaseOrientation_Click(Index As Integer)
  Changed = True

End Sub

Private Sub optBasePicklist_Click()

  cmdBasePicklist.Enabled = True

  With txtBaseFilter
    .Text = ""
    .Tag = 0
  End With

  txtBasePicklist.Text = "<None>"
  cmdBaseFilter.Enabled = False
  
  If Not mblnLoading Then
    Changed = True
    ForceDefinitionToBeHiddenIfNeeded
    EnableDisableTabControls
  End If
  
End Sub


Private Sub optOutputFormat_Click(Index As Integer)
  mobjOutputDef.FormatClick Index
  Changed = True
  
End Sub






Private Sub spnDec_Change()
  Dim objItem As clsRecordProfileColDtl
  
  If spnDec.Text = "" Then spnDec.Text = "0"
  
  If Not mblnLoading Then
    Set objItem = mcolRecordProfileColumnDetails.Item(ListView2.SelectedItem.Key)
    objItem.DecPlaces = spnDec.Text
    Set objItem = Nothing
    Changed = True
  End If

End Sub


Private Sub spnSize_Change()
  Dim objItem As clsRecordProfileColDtl
  
  If spnSize.Text = "" Then spnSize.Text = "0"
  
  If Not mblnLoading Then
    Set objItem = mcolRecordProfileColumnDetails.Item(ListView2.SelectedItem.Key)
    objItem.Size = spnSize.Text
    Set objItem = Nothing
    Changed = True
  End If

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
  EnableDisableTabControls

End Sub

Public Sub EnableDisableTabControls()
  
  Dim mblnWasNotChanged As Boolean
  Dim objItem As ListItem

  mblnWasNotChanged = Changed

  If mblnReadOnly Then
    Exit Sub
  End If

  ' TAB 1 CONTROLS
  fraInformation.Enabled = (SSTab1.Tab = 0)
  fraBase.Enabled = (SSTab1.Tab = 0)
  chkPrintFilterHeader.Enabled = (optBaseFilter.Value) Or (optBasePicklist.Value)

  ' TAB 2 CONTROLS
  fraRelatedTables.Enabled = (SSTab1.Tab = 1)
  If (SSTab1.Tab = 1) Then
    UpdateButtonStatus (SSTab1.Tab)
  End If

  ' TAB 3 CONTROLS
  fraTable.Enabled = (SSTab1.Tab = 2)
  fraFieldsAvailable.Enabled = (SSTab1.Tab = 2)
  fraFieldsSelected.Enabled = (SSTab1.Tab = 2)
  cmdAdd.Enabled = (SSTab1.Tab = 2)
  cmdAddAll.Enabled = (SSTab1.Tab = 2)
  cmdAddHeading.Enabled = (SSTab1.Tab = 2)
  cmdAddSeparator.Enabled = (SSTab1.Tab = 2)
  cmdRemove.Enabled = (SSTab1.Tab = 2)
  cmdRemoveAll.Enabled = (SSTab1.Tab = 2)
  cmdMoveUp.Enabled = (SSTab1.Tab = 2)
  cmdMoveDown.Enabled = (SSTab1.Tab = 2)

  ' TAB 4 CONTROLS
  'JPD 20030728 Fault 6408
  fraReportOptions.Enabled = (SSTab1.Tab = 3) And (grdRelatedTables.Rows > 0)
  chkIndent.Enabled = fraReportOptions.Enabled
  chkSuppressEmptyRelatedTableTitles.Enabled = fraReportOptions.Enabled
  chkShowTableRelationshipTitle.Enabled = fraReportOptions.Enabled
  
  If (grdRelatedTables.Rows = 0) Then
    chkIndent.Value = vbUnchecked
    chkSuppressEmptyRelatedTableTitles.Value = vbUnchecked
    chkShowTableRelationshipTitle.Value = vbUnchecked
  End If
  
  fraOutputFormat.Enabled = (SSTab1.Tab = 3)
  fraOutputDestination.Enabled = (SSTab1.Tab = 3)
  'fraOutputFilename.Enabled = (SSTab1.Tab = 3)

  If SSTab1.Tab = 2 Then
    ' column tab
    If ListView1.ListItems.Count > 0 Then
      ListView1.ListItems(1).Selected = True
      ListView1.Refresh
      cmdAdd.SetFocus
    End If

    If ListView2.ListItems.Count > 0 Then
      For Each objItem In ListView2.ListItems
        objItem.Selected = IIf(objItem.Index = 1, True, False)
      Next objItem
    End If

    UpdateButtonStatus (SSTab1.Tab)
  End If

  If mblnWasNotChanged = False Then Changed = False
  
End Sub




Private Sub txtDesc_Change()
  Changed = True

End Sub


Private Sub txtDesc_GotFocus()
  With txtDesc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
  cmdOK.Default = False

End Sub


Private Sub txtDesc_LostFocus()
  cmdOK.Default = True

End Sub


Private Sub txtEmailGroup_Change()
  If Not mblnLoading Then
    Changed = True
  End If

End Sub

Private Sub txtEmailSubject_Change()
  If Not mblnLoading Then
    Changed = True
  End If

End Sub


Private Sub txtEmailAttachAs_Change()
  If Not mblnLoading Then
    Changed = True
  End If

End Sub


Private Sub txtFilename_Change()
  If Not mblnLoading Then
    Changed = True
  End If

End Sub

Private Sub txtName_Change()
  Changed = True

End Sub


Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub


Private Sub txtProp_ColumnHeading_Change()
  Dim objItem As clsRecordProfileColDtl
  
  If Not mblnLoading Then
    Set objItem = mcolRecordProfileColumnDetails.Item(ListView2.SelectedItem.Key)
    objItem.Heading = txtProp_ColumnHeading.Text
    Set objItem = Nothing
    Changed = True
  End If

End Sub


Private Sub txtProp_ColumnHeading_GotFocus()
  UI.txtSelText

End Sub


Private Sub txtProp_ColumnHeading_LostFocus()
  txtProp_ColumnHeading.Text = Trim(txtProp_ColumnHeading.Text)

End Sub



Private Sub PopulateAccessGrid()
  ' Populate the access grid.
  Dim rsAccess As ADODB.Recordset
  
  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With
  
  ' Get the recordset of user groups and their access on this definition.
  Set rsAccess = GetUtilityAccessRecords(utlRecordProfile, mlngRecordProfileID, mblnFromCopy)
  If Not rsAccess Is Nothing Then
    ' Add the user groups and their access on this definition to the access grid.
    With rsAccess
      Do While Not .EOF
        grdAccess.AddItem !Name & vbTab & AccessDescription(!Access) & vbTab & !sysSecMgr
        
        .MoveNext
      Loop
    
      .Close
    End With
  End If
  Set rsAccess = Nothing

End Sub

