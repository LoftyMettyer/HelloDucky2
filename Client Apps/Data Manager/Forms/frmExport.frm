VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Definition"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1037
   Icon            =   "frmExport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8325
      TabIndex        =   119
      Top             =   5550
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   7050
      TabIndex        =   118
      Top             =   5550
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5385
      Left            =   45
      TabIndex        =   120
      Top             =   45
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   9499
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
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
      TabPicture(0)   =   "frmExport.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBase"
      Tab(0).Control(1)=   "fraInformation"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Related &Tables"
      TabPicture(1)   =   "frmExport.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraParent1"
      Tab(1).Control(1)=   "fraParent2"
      Tab(1).Control(2)=   "fraChild"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Colu&mns"
      TabPicture(2)   =   "frmExport.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraColumns"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Sort Order"
      TabPicture(3)   =   "frmExport.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraExportOrder"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "O&ptions"
      TabPicture(4)   =   "frmExport.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraHeaderOptions"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fraDateOptions"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "O&utput"
      TabPicture(5)   =   "frmExport.frx":0098
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "fraDelimFile"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "fraCMGFile"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "fraXML"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "fraOutputDestination"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "fraOutputType"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      Begin VB.Frame fraInformation 
         Height          =   2355
         Left            =   -74850
         TabIndex        =   121
         Top             =   400
         Width           =   9180
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   3090
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1395
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1110
            Width           =   3090
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   1
            Top             =   300
            Width           =   3090
         End
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5625
            MaxLength       =   30
            TabIndex        =   4
            Top             =   300
            Width           =   3405
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1485
            Left            =   5625
            TabIndex        =   5
            Top             =   705
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
            stylesets(0).Picture=   "frmExport.frx":00B4
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
            stylesets(1).Picture=   "frmExport.frx":00D0
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
            _ExtentY        =   2619
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
         Begin VB.Label lblCategory 
            Caption         =   "Category :"
            Height          =   240
            Left            =   195
            TabIndex        =   126
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   4770
            TabIndex        =   125
            Top             =   765
            Width           =   825
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   195
            TabIndex        =   124
            Top             =   1155
            Width           =   1080
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   123
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Left            =   4770
            TabIndex        =   122
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame fraOutputType 
         Caption         =   "Output Format :"
         Height          =   2835
         Left            =   150
         TabIndex        =   82
         Top             =   405
         Width           =   2400
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "X&ML"
            Height          =   195
            Index           =   9
            Left            =   200
            TabIndex        =   127
            Top             =   1600
            Width           =   1200
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   85
            Top             =   1200
            Width           =   1920
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "F&ixed Length File"
            Height          =   195
            Index           =   7
            Left            =   200
            TabIndex        =   84
            Top             =   800
            Width           =   1965
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "De&limited File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   83
            Top             =   400
            Width           =   1470
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CM&G File"
            Height          =   195
            Index           =   8
            Left            =   200
            TabIndex        =   86
            Top             =   2000
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "SQL &Table"
            Height          =   195
            Index           =   99
            Left            =   200
            TabIndex        =   87
            Top             =   2400
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   2835
         Left            =   2655
         TabIndex        =   88
         Top             =   405
         Width           =   6675
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3360
            TabIndex        =   102
            Tag             =   "0"
            Top             =   2145
            Width           =   3135
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   315
            Width           =   2835
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            TabIndex        =   100
            Top             =   1740
            Width           =   3135
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1335
            Width           =   2835
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6195
            TabIndex        =   92
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6195
            TabIndex        =   98
            Top             =   1335
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   89
            Top             =   375
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   95
            Top             =   1395
            Width           =   1515
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   1935
            TabIndex        =   101
            Top             =   2205
            Width           =   1155
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1935
            TabIndex        =   90
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1935
            TabIndex        =   99
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   1935
            TabIndex        =   96
            Top             =   1395
            Width           =   1245
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1935
            TabIndex        =   93
            Top             =   780
            Width           =   1350
         End
      End
      Begin VB.Frame fraHeaderOptions 
         Caption         =   "Header && Footer :"
         Height          =   2085
         Left            =   -74850
         TabIndex        =   64
         Top             =   405
         Width           =   9180
         Begin VB.CommandButton cmdFooterText 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4800
            TabIndex        =   144
            Top             =   1485
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdHeaderText 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4800
            TabIndex        =   143
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkForceHeader 
            Caption         =   "&Force header if no records"
            Height          =   195
            Left            =   5445
            TabIndex        =   73
            Top             =   360
            Width           =   3345
         End
         Begin VB.ComboBox cboFooterOptions 
            Height          =   315
            ItemData        =   "frmExport.frx":00EC
            Left            =   1845
            List            =   "frmExport.frx":00F9
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1100
            Width           =   3300
         End
         Begin VB.TextBox txtCustomFooter 
            Height          =   315
            Left            =   1845
            MaxLength       =   255
            TabIndex        =   72
            Top             =   1500
            Width           =   2955
         End
         Begin VB.TextBox txtCustomHeader 
            Height          =   315
            Left            =   1845
            MaxLength       =   255
            TabIndex        =   68
            Top             =   700
            Width           =   2955
         End
         Begin VB.ComboBox cboHeaderOptions 
            Height          =   315
            ItemData        =   "frmExport.frx":0125
            Left            =   1845
            List            =   "frmExport.frx":0135
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   300
            Width           =   3315
         End
         Begin VB.CheckBox chkOmitHeader 
            Caption         =   "Omit header when &appending to file"
            Height          =   195
            Left            =   5445
            TabIndex        =   74
            Top             =   660
            Width           =   3465
         End
         Begin VB.Label lblFooterLine 
            AutoSize        =   -1  'True
            Caption         =   "Footer Line :"
            Height          =   195
            Left            =   195
            TabIndex        =   69
            Top             =   1155
            Width           =   1170
         End
         Begin VB.Label lblHeaderLine 
            AutoSize        =   -1  'True
            Caption         =   "Header Line :"
            Height          =   195
            Left            =   195
            TabIndex        =   65
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label lblCustomHeader 
            AutoSize        =   -1  'True
            Caption         =   "Custom Header :"
            Height          =   195
            Left            =   195
            TabIndex        =   67
            Top             =   765
            Width           =   1485
         End
         Begin VB.Label lblCustomFooter 
            AutoSize        =   -1  'True
            Caption         =   "Custom Footer :"
            Height          =   195
            Left            =   195
            TabIndex        =   71
            Top             =   1560
            Width           =   1425
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns :"
         Height          =   4560
         Left            =   -74850
         TabIndex        =   47
         Top             =   405
         Width           =   9180
         Begin VB.CommandButton cmdAddAllColumns 
            Caption         =   "Add A&ll..."
            Height          =   400
            Left            =   7800
            TabIndex        =   50
            Top             =   800
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "Move U&p"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   54
            Top             =   3460
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Move Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   55
            Top             =   3960
            Width           =   1200
         End
         Begin VB.CommandButton cmdClearColumn 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   53
            Top             =   2300
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteColumn 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   52
            Top             =   1800
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditColumn 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   51
            Top             =   1300
            Width           =   1200
         End
         Begin VB.CommandButton cmdNewColumn 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7800
            TabIndex        =   49
            Top             =   315
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdColumns 
            Height          =   4065
            Left            =   195
            TabIndex        =   48
            Top             =   300
            Width           =   7545
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   11
            stylesets.count =   5
            stylesets(0).Name=   "ssetHeaderDisabled"
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
            stylesets(0).Picture=   "frmExport.frx":0181
            stylesets(1).Name=   "ssetSelected"
            stylesets(1).ForeColor=   -2147483634
            stylesets(1).BackColor=   -2147483635
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
            stylesets(1).Picture=   "frmExport.frx":019D
            stylesets(2).Name=   "ssetEnabled"
            stylesets(2).ForeColor=   -2147483640
            stylesets(2).BackColor=   -2147483643
            stylesets(2).HasFont=   -1  'True
            BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(2).Picture=   "frmExport.frx":01B9
            stylesets(3).Name=   "ssetHeaderEnabled"
            stylesets(3).ForeColor=   -2147483630
            stylesets(3).BackColor=   -2147483633
            stylesets(3).HasFont=   -1  'True
            BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(3).Picture=   "frmExport.frx":01D5
            stylesets(4).Name=   "ssetDisabled"
            stylesets(4).ForeColor=   -2147483631
            stylesets(4).BackColor=   -2147483633
            stylesets(4).HasFont=   -1  'True
            BeginProperty stylesets(4).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(4).Picture=   "frmExport.frx":01F1
            CheckBox3D      =   0   'False
            AllowUpdate     =   0   'False
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
            SelectTypeRow   =   1
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ActiveRowStyleSet=   "ssetSelected"
            Columns.Count   =   11
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "Type"
            Columns(0).Name =   "Type"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "TableID"
            Columns(1).Name =   "TableID"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "ColExprID"
            Columns(2).Name =   "ColExprID"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   9287
            Columns(3).Caption=   "Data"
            Columns(3).Name =   "Data"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   1852
            Columns(4).Caption=   "Size"
            Columns(4).Name =   "Length"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "CMG Code"
            Columns(5).Name =   "CMG Code"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "Audit"
            Columns(6).Name =   "Audit"
            Columns(6).Alignment=   2
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   11
            Columns(6).FieldLen=   50
            Columns(6).Style=   2
            Columns(7).Width=   1852
            Columns(7).Caption=   "Decimals"
            Columns(7).Name =   "Decimals"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "Heading"
            Columns(8).Name =   "Heading"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "ConvertCase"
            Columns(9).Name =   "ConvertCase"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "SuppressNulls"
            Columns(10).Name=   "SuppressNulls"
            Columns(10).Alignment=   2
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   11
            Columns(10).FieldLen=   50
            Columns(10).Style=   2
            TabNavigation   =   1
            _ExtentX        =   13309
            _ExtentY        =   7170
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
      End
      Begin VB.Frame fraChild 
         Caption         =   "Child :"
         Height          =   1240
         Left            =   -74850
         TabIndex        =   38
         Top             =   3720
         Width           =   9180
         Begin VB.CommandButton cmdChildFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            TabIndex        =   43
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboChild 
            Height          =   315
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   300
            Width           =   3000
         End
         Begin VB.TextBox txtChildFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5800
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   300
            Width           =   2900
         End
         Begin COASpinner.COA_Spinner spnMaxRecords 
            Height          =   315
            Left            =   5800
            TabIndex        =   45
            Top             =   700
            Width           =   1000
            _ExtentX        =   1746
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
         Begin VB.Label lblMaxRecordsAll 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(All Records)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   7005
            TabIndex        =   46
            Top             =   765
            Width           =   1275
         End
         Begin VB.Label lblMaxRecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4860
            TabIndex        =   44
            Top             =   765
            Width           =   870
         End
         Begin VB.Label lblChildTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   180
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblChildFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4860
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraParent2 
         Caption         =   "Parent 2 :"
         Height          =   1600
         Left            =   -74850
         TabIndex        =   27
         Top             =   2060
         Width           =   9180
         Begin VB.OptionButton optParent2Filter 
            Caption         =   "Filter"
            Height          =   195
            Left            =   5715
            TabIndex        =   35
            Top             =   1160
            Width           =   840
         End
         Begin VB.OptionButton optParent2Picklist 
            Caption         =   "Picklist"
            Height          =   195
            Left            =   5715
            TabIndex        =   32
            Top             =   760
            Width           =   930
         End
         Begin VB.OptionButton optParent2AllRecords 
            Caption         =   "All"
            Height          =   195
            Left            =   5715
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   630
         End
         Begin VB.TextBox txtParent2Picklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   33
            Tag             =   "0"
            Top             =   700
            Width           =   2000
         End
         Begin VB.CommandButton cmdParent2Picklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            TabIndex        =   34
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdParent2Filter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            TabIndex        =   37
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtParent2Filter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1100
            Width           =   2000
         End
         Begin VB.TextBox txtParent2 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1000
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   300
            Width           =   3000
         End
         Begin VB.Label lblParent2Records 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   4860
            TabIndex        =   30
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblParent2Table 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   200
            TabIndex        =   28
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraParent1 
         Caption         =   "Parent 1 :"
         Height          =   1600
         Left            =   -74850
         TabIndex        =   16
         Top             =   400
         Width           =   9180
         Begin VB.CommandButton cmdParent1Picklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            TabIndex        =   23
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtParent1Picklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   22
            Tag             =   "0"
            Top             =   700
            Width           =   2000
         End
         Begin VB.OptionButton optParent1AllRecords 
            Caption         =   "All"
            Height          =   195
            Left            =   5715
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton optParent1Picklist 
            Caption         =   "Picklist"
            Height          =   195
            Left            =   5715
            TabIndex        =   21
            Top             =   760
            Width           =   885
         End
         Begin VB.OptionButton optParent1Filter 
            Caption         =   "Filter"
            Height          =   195
            Left            =   5715
            TabIndex        =   24
            Top             =   1160
            Width           =   795
         End
         Begin VB.CommandButton cmdParent1Filter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            TabIndex        =   26
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtParent1Filter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1100
            Width           =   2000
         End
         Begin VB.TextBox txtParent1 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1000
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   300
            Width           =   3000
         End
         Begin VB.Label lblParent1Records 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4860
            TabIndex        =   19
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblParent1Table 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   200
            TabIndex        =   17
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraExportOrder 
         Caption         =   "Sort Order :"
         Height          =   4560
         Left            =   -74850
         TabIndex        =   56
         Top             =   405
         Width           =   9180
         Begin VB.CommandButton cmdSortMoveDown 
            Caption         =   "Move Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   63
            Top             =   3960
            Width           =   1200
         End
         Begin VB.CommandButton cmdSortMoveUp 
            Caption         =   "Move U&p"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   62
            Top             =   3460
            Width           =   1200
         End
         Begin VB.CommandButton cmdNewOrder 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7800
            TabIndex        =   58
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditOrder 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   59
            Top             =   800
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteOrder 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   60
            Top             =   1300
            Width           =   1200
         End
         Begin VB.CommandButton cmdClearOrder 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   61
            Top             =   1800
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdExportOrder 
            Height          =   4065
            Left            =   195
            TabIndex        =   57
            Top             =   300
            Width           =   7395
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            stylesets.count =   2
            stylesets(0).Name=   "ssetDormant"
            stylesets(0).Picture=   "frmExport.frx":020D
            stylesets(1).Name=   "ssetActive"
            stylesets(1).ForeColor=   -2147483634
            stylesets(1).BackColor=   -2147483635
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
            stylesets(1).Picture=   "frmExport.frx":0229
            AllowUpdate     =   0   'False
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
            SelectTypeRow   =   1
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   1
            StyleSet        =   "ssetDormant"
            ForeColorEven   =   0
            BackColorOdd    =   16777215
            RowHeight       =   423
            ActiveRowStyleSet=   "ssetActive"
            Columns.Count   =   3
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ColExprID"
            Columns(0).Name =   "ColExprID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   10769
            Columns(1).Caption=   "Column"
            Columns(1).Name =   "Column"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   2196
            Columns(2).Caption=   "Sort Order"
            Columns(2).Name =   "Sort Order"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13044
            _ExtentY        =   7170
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
      End
      Begin VB.Frame fraBase 
         Caption         =   "Data :"
         Height          =   2115
         Left            =   -74850
         TabIndex        =   0
         Top             =   2850
         Width           =   9180
         Begin VB.CommandButton cmdBaseFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   15
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdBasePicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   12
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtBasePicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   11
            Tag             =   "0"
            Top             =   700
            Width           =   2000
         End
         Begin VB.TextBox txtBaseFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   14
            Tag             =   "0"
            Top             =   1100
            Width           =   2000
         End
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   300
            Width           =   3000
         End
         Begin VB.OptionButton optBaseAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   5715
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton optBasePicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   5715
            TabIndex        =   10
            Top             =   760
            Width           =   885
         End
         Begin VB.OptionButton optBaseFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   5715
            TabIndex        =   13
            Top             =   1160
            Width           =   800
         End
         Begin VB.Label lblBaseRecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   4770
            TabIndex        =   8
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblBaseTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   6
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame fraDateOptions 
         Caption         =   "Date Format :"
         Height          =   2450
         Left            =   -74850
         TabIndex        =   75
         Top             =   2550
         Width           =   9180
         Begin VB.ComboBox cboDateFormat 
            Height          =   315
            ItemData        =   "frmExport.frx":0245
            Left            =   1845
            List            =   "frmExport.frx":0255
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   300
            Width           =   2000
         End
         Begin VB.ComboBox cboDateSeparator 
            Height          =   315
            ItemData        =   "frmExport.frx":026D
            Left            =   1845
            List            =   "frmExport.frx":027D
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   700
            Width           =   2000
         End
         Begin VB.ComboBox cboDateYearDigits 
            Height          =   315
            ItemData        =   "frmExport.frx":0292
            Left            =   1845
            List            =   "frmExport.frx":029C
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   1100
            Width           =   1000
         End
         Begin VB.Label lblDateFormat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Regional :"
            Height          =   195
            Left            =   200
            TabIndex        =   76
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblDateSeparator 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separator :"
            Height          =   195
            Left            =   200
            TabIndex        =   78
            Top             =   760
            Width           =   825
         End
         Begin VB.Label lblDateYearDigits 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year Digits :"
            Height          =   195
            Left            =   200
            TabIndex        =   80
            Top             =   1160
            Width           =   870
         End
      End
      Begin VB.Frame fraXML 
         Caption         =   "XML Options :"
         Height          =   1950
         Left            =   150
         TabIndex        =   128
         Top             =   3270
         Width           =   9180
         Begin VB.CheckBox chkSplitXMLNodesFile 
            Caption         =   "&Split nodes into individual files"
            Height          =   210
            Left            =   5925
            TabIndex        =   142
            Top             =   1665
            Width           =   3030
         End
         Begin VB.CheckBox chkPreserveTransformPath 
            Caption         =   "Preser&ve transformation path"
            Height          =   270
            Left            =   5925
            TabIndex        =   141
            Top             =   1215
            Width           =   2865
         End
         Begin VB.CheckBox chkPreserveXSDPath 
            Caption         =   "Preserve XSD pat&h"
            Height          =   270
            Left            =   5925
            TabIndex        =   136
            Top             =   780
            Width           =   2355
         End
         Begin VB.CommandButton cmdXSDFile 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   5115
            TabIndex        =   134
            Top             =   750
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdXSDFileClear 
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5460
            MaskColor       =   &H000000FF&
            TabIndex        =   135
            ToolTipText     =   "Clear Path"
            Top             =   750
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtXSDFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2115
            TabIndex        =   139
            Top             =   750
            Width           =   3000
         End
         Begin VB.CheckBox chkAuditChangesOnly 
            Caption         =   "Only &include audited changes"
            Height          =   210
            Left            =   5940
            TabIndex        =   133
            Top             =   360
            Width           =   3030
         End
         Begin VB.TextBox txtXMLDataNodeName 
            Height          =   315
            Left            =   2115
            MaxLength       =   50
            TabIndex        =   132
            Top             =   315
            Width           =   3000
         End
         Begin VB.CommandButton cmdTransformFileClear 
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5460
            MaskColor       =   &H000000FF&
            TabIndex        =   140
            ToolTipText     =   "Clear Path"
            Top             =   1200
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdTransformFile 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   5115
            TabIndex        =   138
            Top             =   1200
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtTransformFile 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   130
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1200
            Width           =   3000
         End
         Begin VB.Label lblXSDFile 
            Caption         =   "XSD File :"
            Height          =   210
            Left            =   240
            TabIndex        =   137
            Top             =   825
            Width           =   1410
         End
         Begin VB.Label lblXMLNodeName 
            Caption         =   "Custom Node Name : "
            Height          =   285
            Left            =   225
            TabIndex        =   131
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label lblTransformFile 
            Caption         =   "Transformation File :"
            Height          =   255
            Left            =   225
            TabIndex        =   129
            Top             =   1275
            Width           =   1815
         End
      End
      Begin VB.Frame fraCMGFile 
         Caption         =   "CMG Options :"
         Height          =   1665
         Left            =   150
         TabIndex        =   111
         Top             =   3285
         Width           =   9180
         Begin VB.CheckBox chkUpdateAuditPointer 
            Caption         =   "Commit &after run"
            Height          =   195
            Left            =   5500
            TabIndex        =   117
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtFileExportCode 
            Height          =   315
            Left            =   1995
            MaxLength       =   10
            TabIndex        =   116
            Top             =   700
            Width           =   1000
         End
         Begin VB.ComboBox cboParentFields 
            Height          =   315
            Left            =   2000
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   300
            Width           =   3000
         End
         Begin VB.Label lblFileExportCode 
            Caption         =   "File Export Code :"
            Height          =   195
            Left            =   195
            TabIndex        =   115
            Top             =   765
            Width           =   1605
         End
         Begin VB.Label lblCMGRecordIdentifier 
            Caption         =   "Record Identifier :"
            Height          =   195
            Left            =   195
            TabIndex        =   113
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.Frame fraDelimFile 
         Caption         =   "Delimited File Options :"
         Height          =   1665
         Left            =   150
         TabIndex        =   103
         Top             =   3285
         Width           =   9180
         Begin COASpinner.COA_Spinner spnSplitFileSize 
            Height          =   300
            Left            =   7920
            TabIndex        =   109
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
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
            MaximumValue    =   99999
            Text            =   "0"
         End
         Begin VB.CheckBox chkSplitFile 
            Caption         =   "Split file in &blocks"
            Height          =   300
            Left            =   4710
            TabIndex        =   108
            Top             =   315
            Width           =   1875
         End
         Begin VB.CheckBox chkStripDelimiter 
            Caption         =   "Strip delimiter from d&ata"
            Height          =   240
            Left            =   4710
            TabIndex        =   112
            Top             =   1125
            Width           =   2475
         End
         Begin VB.TextBox txtDelimiter 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   3420
            MaxLength       =   1
            TabIndex        =   107
            Top             =   300
            Width           =   345
         End
         Begin VB.CheckBox chkQuotes 
            Caption         =   "Enclose in &quotes"
            Height          =   195
            Left            =   4710
            TabIndex        =   110
            Top             =   750
            Width           =   1875
         End
         Begin VB.ComboBox cboDelimiter 
            Height          =   315
            ItemData        =   "frmExport.frx":02A6
            Left            =   1455
            List            =   "frmExport.frx":02B3
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblSplitFileSize 
            Caption         =   "Block Size :"
            Height          =   300
            Left            =   6810
            TabIndex        =   145
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label lblOtherDelimiter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other :"
            Height          =   210
            Left            =   2730
            TabIndex        =   106
            Top             =   360
            Width           =   585
         End
         Begin VB.Label lblDelimiter 
            BackStyle       =   0  'Transparent
            Caption         =   "Delimiter :"
            Height          =   195
            Left            =   315
            TabIndex        =   104
            Top             =   360
            Width           =   915
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DataAccess Class
Private mdatData As DataMgr.clsDataAccess
Private mobjOutputDef As clsOutputDef

' Long to hold current Export ID
Private mlngExportID As Long

' Variables to hold current (or previously) selected table details
Private mstrBaseTable As String
Private mstrChildTable As String
Private mlngChildTable As Long

Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mblnFormPrint As Boolean
'Private mblnChanged As Boolean
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mlngTimeStamp As Long

Private mblnColGrid As Boolean
Private mblnSortGrid As Boolean

Public mblnDefinitionCreator As Boolean
Private mblnForceHidden As Boolean
Private mblnRecordSelectionInvalid As Boolean
Private mblnBaseTableSpecificChanged As Boolean
Private mblnDeleted As Boolean
Private mblnNotified As Boolean
Private sMessage As String
Private mbNeedsSave As Boolean

' CMG variables
Private mbCMGExportFileCode As Boolean
Private mbCMGExportFieldCode As Boolean

Private Const lng_CMGCodeCOLUMNWIDTH = 1050
Private Const lng_AuditCOLUMNWIDTH = 1050
Private Const lng_DecimalCOLUMNWIDTH = 1050
Private Const lng_LengthCOLUMNWIDTH = 1050
Private lng_DataCOLUMNWIDTH As Double
Private Const lng_GRIDROWHEIGHT = 239
Private Const lng_SCROLLBARWIDTH = 240
Private Const lng_SortCOLUMNWIDTH = 5890
Private Const lng_SortOrderCOLUMNWIDTH = 1250

Private pstrType As String


Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property
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


Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOk.Enabled = pblnChanged
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngExportID
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get FromCopy() As Boolean
  FromCopy = mblnFromCopy
End Property

Public Property Let FromCopy(ByVal bCopy As Boolean)
  mblnFromCopy = bCopy
End Property

Public Function Initialise(pblnNew As Boolean, pblnCopy As Boolean, Optional plngExportID As Long, Optional bPrint As Boolean) As Boolean
  
  CheckIfCMGEnabled
  optOutputFormat(fmtXML).Enabled = gbXMLExportEnabled
  
  mblnLoading = True
 
  ' Set reference to data access class module
  Set mdatData = New DataMgr.clsDataAccess
  
  Screen.MousePointer = vbHourglass

  If pblnNew Then

    mlngExportID = 0

    'Set controls to defaults
    ClearForNew
        
    'Load All Possible Base Tables into combo
    LoadBaseCombo

    ' Set the categories combo
    GetObjectCategories cboCategory, utlExport, 0, cboBaseTable.ItemData(cboBaseTable.ListIndex)
    SetComboItem cboCategory, IIf(glngCurrentCategoryID = -1, 0, glngCurrentCategoryID)

    UpdateDependantFields
    
    ' Set command button status
    UpdateButtonStatus

    mblnDefinitionCreator = True

    ' Default filetype for new defs is delimited, for which size doesnt matter
    grdColumns.Columns(4).ForeColor = vbWindowText
    grdColumns.Columns(4).HeadForeColor = vbWindowText
    
    PopulateAccessGrid
    Changed = False

  Else

    ' Make the ExportID visible to the rest of the module
    mlngExportID = plngExportID
    
    ' Is is a copy of an existing one ?
    FromCopy = pblnCopy

    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint

    PopulateAccessGrid

    If Not RetrieveExportDetails(mlngExportID) Then
      If mblnDeleted Or Me.Cancelled Then
        Changed = False
        Initialise = False
        Exit Function
      Else
        If COAMsgBox("OpenHR could not load all of the definition successfully. The recommendation is that" & vbCrLf & _
               "you delete the definition and create a new one, however, you may edit the existing" & vbCrLf & _
               "definition if you wish. Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, "Export") = vbNo Then
          Initialise = False
          Exit Function
        End If
      End If
    End If
    
    UpdateButtonStatus

    'Reset pointer so copy will be saved as new
    If pblnCopy Then
      mlngExportID = 0
    End If
    
  End If
    

  If mblnFromCopy Then
    mlngExportID = 0
    Changed = True
  Else
    Changed = mblnRecordSelectionInvalid And (Not mblnReadOnly) ' False
  End If
    
  mblnLoading = False
  Cancelled = False
  If mblnForceHidden Then mblnForceHidden = True
  Screen.MousePointer = vbDefault
  Initialise = True
  
End Function

Public Sub LoadBaseCombo()

  'Purpose : Populate the base table combo with all tables
  'Input   : None
  'Output  : None
  
  On Error GoTo LoadBaseCombo_ERROR
  
  Dim pstrSQL As String
  Dim prstTables As New Recordset

  pstrSQL = "Select TableName, TableID From ASRSysTables "
  
  ' Uncomment this line to exclude lookup tables
  ' pstrSQL = pstrsql & "WHERE TableType = 1 OR TableType = 2 "
  
  pstrSQL = pstrSQL & "ORDER BY TableName"
  
  Set prstTables = mdatData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)

  With cboBaseTable
    .Clear
    If Not prstTables.EOF Then mstrBaseTable = gsPersonnelTableName ' prstTables!TableName
    Do While Not prstTables.EOF
      .AddItem prstTables!TableName
      .ItemData(.NewIndex) = prstTables!TableID
      prstTables.MoveNext
    Loop
'    .ListIndex = 0
    If .ListCount > 0 Then
      If gsPersonnelTableName <> "" Then
        SetComboText cboBaseTable, gsPersonnelTableName
      Else
        .ListIndex = 0
      End If
    End If
  End With

  pstrSQL = vbNullString
  Set prstTables = Nothing
  Exit Sub
  
LoadBaseCombo_ERROR:
  
  pstrSQL = vbNullString
  Set prstTables = Nothing
  COAMsgBox "Error populating the base table combo box." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Export"
  
End Sub

Private Sub cboBaseTable_Click()
  
  'Purpose : When the user changes the Base Table, prompt the user that the definition
  'Input   : None
  'Output  : None
  If mblnLoading = True Then Exit Sub
  
'  If (mstrBaseTable = cboBaseTable.Text) Or mstrBaseTable = "" Then Exit Sub
'  If (mstrBaseTable = cboBaseTable.Text) Then Exit Sub
  
  'MH20010828 Can't do this as we still need to show the parent tables and
  'populate the child tables combo etc..
  '''TM20010823 Fault 2707
  '''Only need to do validation if there are no columns already selected.
  '''If (grdColumns.Rows = 0) Or (mstrBaseTable = cboBaseTable.Text) Then Exit Sub
  
  'If changed = False Then
  '  mstrBaseTable = cboBaseTable.Text
  '  UpdateDependantFields
  '  changed = True
  '  Exit Sub
  'End If

  '01/08/2001 MH Fault 2615
  'If mblnBaseTableSpecificChanged = True Then
    
  If mstrBaseTable <> cboBaseTable.Text Then
    If grdColumns.Rows > 0 Then
      If COAMsgBox("Warning: Changing the base table will result in all table/column " & _
              "specific aspects of this export definition being cleared." & vbCrLf & _
              "Are you sure you wish to continue?", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Export") = vbYes Then
        mblnLoading = True
        ClearForNew True
        mstrBaseTable = cboBaseTable.Text
        mblnLoading = False
        mblnBaseTableSpecificChanged = False
      
      'MH20010831 Fault 2771
      Else
        SetComboText cboBaseTable, mstrBaseTable
        
        'TM20011212 Fault 2791
        SetComboText cboChild, mstrChildTable
      End If
    Else
            
      'JDM - 20/11/01 - Fault 3176 - Not clearing picklists
      ClearForNew True
    
    End If
    UpdateDependantFields
    Changed = True
  'MH20010831 Fault 2771
    mstrBaseTable = cboBaseTable.Text
  'Else
  '  SetComboText cboBaseTable, mstrBaseTable
  End If
  'Else
  '  UpdateDependantFields
  'End If

  UpdateButtonStatus
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub

Private Sub cboBaseTable_LostFocus()

'  If (mstrBaseTable = cboBaseTable.Text) Or mstrBaseTable = "" Then Exit Sub
'
'  If mblnBaseTableSpecificChanged = False Then
'    mstrBaseTable = cboBaseTable.Text
'    UpdateDependantFields
'    changed = True
'    Exit Sub
'  End If
'
'  If mblnBaseTableSpecificChanged = True Then
'    If COAMsgBox("Warning : Changing the base table will result in all table/column specific aspects of this" & vbCrLf & " export definition being cleared. Are you sure you wish to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Export") = vbYes Then
'      ClearForNew True
'      mstrBaseTable = cboBaseTable.Text
'      UpdateDependantFields
'      changed = True
'      mblnBaseTableSpecificChanged = False
'    Else
'      SetComboText cboBaseTable, mstrBaseTable
'    End If
'  End If
  
End Sub

Public Sub UpdateDependantFields()

  'Purpose : Populates the parent/child combos depending on the base table selected
  'Input   : None
  'Output  : None
  
  Dim pstrSQL As String
  Dim prstTables As New Recordset
  Dim prstFields As New Recordset
  Dim fOriginalLoading As Boolean
  
'  If mblnLoading Then Exit Sub
  
  ' Get the parent(s) of the selected base table
    
  pstrSQL = "SELECT ASRSysTables.tablename, ASRSysTables.tableid " & _
            "FROM ASRSysTables " & _
            "WHERE ASRSysTables.tableid in " & _
            "(select parentid from ASRSysRelations " & _
            "WHERE childid = " & cboBaseTable.ItemData(cboBaseTable.ListIndex) & ") " & _
            "ORDER BY tablename"
    
  Set prstTables = mdatData.OpenRecordset(pstrSQL, adOpenKeyset, adLockReadOnly)
    
  If Not prstTables.BOF And Not prstTables.EOF Then
    prstTables.MoveLast
    prstTables.MoveFirst
  End If
  
  Select Case prstTables.RecordCount
    Case 0
      txtParent1.Text = "" '"<None>"
      txtParent1.Tag = 0
      optParent1AllRecords.Value = True
      txtParent1Filter.Text = ""
      txtParent1Filter.Tag = 0
      txtParent1Picklist.Text = ""
      txtParent1Picklist.Tag = 0
      fraParent1.Enabled = False
'      cmdParent1Filter.Enabled = False
'      lblParent1Filter.Enabled = False
'      lblParent1.Enabled = False
      
      txtParent2.Text = "" '"<None>"
      txtParent2.Tag = 0
      optParent2AllRecords.Value = True
      txtParent2Filter.Text = ""
      txtParent2Filter.Tag = 0
      txtParent2Picklist.Text = ""
      txtParent2Picklist.Tag = 0
      fraParent2.Enabled = False
'      cmdParent2Filter.Enabled = False
'      lblParent2Filter.Enabled = False
    
    Case 1
      txtParent1.Text = prstTables!TableName
      txtParent1.Tag = prstTables!TableID
      optParent1AllRecords.Value = True
      txtParent1Filter.Text = "" '"<None>"
      txtParent1Filter.Tag = 0
      txtParent1Picklist.Text = ""
      txtParent1Picklist.Tag = 0
      fraParent1.Enabled = True
'      cmdParent1Filter.Enabled = True
'      lblParent1Filter.Enabled = True
'      lblParent1.Enabled = True
      
      txtParent2.Text = "" '"<None>"
      txtParent2.Tag = 0
      optParent2AllRecords.Value = True
      txtParent2Filter.Text = ""
      txtParent2Filter.Tag = 0
      txtParent2Picklist.Text = ""
      txtParent2Picklist.Tag = 0
      fraParent2.Enabled = False
'      cmdParent2Filter.Enabled = False
'      lblParent2Filter.Enabled = False
    
    Case 2
      txtParent1.Text = prstTables!TableName
      txtParent1.Tag = prstTables!TableID
      optParent1AllRecords.Value = True
      txtParent1Filter.Text = "" '"<None>"
      txtParent1Filter.Tag = 0
      txtParent1Picklist.Text = ""
      txtParent1Picklist.Tag = 0
      fraParent1.Enabled = True
'      cmdParent1Filter.Enabled = True
'      lblParent1Filter.Enabled = True
'      lblParent1.Enabled = True
      
      prstTables.MoveNext
      
      txtParent2.Text = prstTables!TableName
      txtParent2.Tag = prstTables!TableID
      optParent2AllRecords.Value = True
      txtParent2Filter.Text = "" '"<None>"
      txtParent2Filter.Tag = 0
      txtParent2Picklist.Text = ""
      txtParent2Picklist.Tag = 0
      fraParent2.Enabled = True
'      cmdParent2Filter.Enabled = True
'      lblParent2Filter.Enabled = True
  End Select
    
  ' Get the children of the selected base table and add <None> entry
  
  With cboChild
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    fOriginalLoading = mblnLoading
    mblnLoading = True
    .ListIndex = 0
    mblnLoading = fOriginalLoading
  End With
  
  pstrSQL = "SELECT ASRSysTables.tablename, ASRSysTables.tableid " & _
            "FROM ASRSysTables " & _
            "WHERE ASRSysTables.tableid in " & _
            "(select childid from ASRSysRelations " & _
            "WHERE parentid = " & cboBaseTable.ItemData(cboBaseTable.ListIndex) & ") " & _
            "ORDER BY tablename"
  
  Set prstTables = mdatData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)
  
  If Not prstTables.BOF And Not prstTables.EOF Then

    Do Until prstTables.EOF
      cboChild.AddItem prstTables!TableName
      cboChild.ItemData(cboChild.NewIndex) = prstTables!TableID
      prstTables.MoveNext
    Loop
  
  End If
  
  txtChildFilter.Text = ""
  txtChildFilter.Tag = 0
  spnMaxRecords.Value = 0
  
  'TM20011212 Fault 2791
  SetComboText cboChild, mstrChildTable
  
  ' Stuff info into the CMG Record identifier column
  With cboParentFields
    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'    mblnLoading = True
'    .ListIndex = 0
'    mblnLoading = False
  End With
  
  ' Load column names for this table
  Set prstFields = datGeneral.GetColumnNames(cboBaseTable.ItemData(cboBaseTable.ListIndex), True)
  If Not prstFields.BOF And Not prstFields.EOF Then
    Do Until prstFields.EOF
      cboParentFields.AddItem prstFields!ColumnName
      cboParentFields.ItemData(cboParentFields.NewIndex) = prstFields!ColumnID
      prstFields.MoveNext
    Loop
  End If

  If cboParentFields.ListCount > 0 Then
    cboParentFields.ListIndex = 0
  End If

  Set prstFields = Nothing
  Set prstTables = Nothing

End Sub


Private Sub cboDelimiter_Click()
  With txtDelimiter
      If cboDelimiter.Text = "<Other>" Then
        'If <Other> is selected as a delimiter choice...
        lblOtherDelimiter.Enabled = True
        txtDelimiter.Enabled = True
        txtDelimiter.BackColor = &H80000005
      Else
        lblOtherDelimiter.Enabled = False
        txtDelimiter.Enabled = False
        txtDelimiter.BackColor = &H8000000F
        txtDelimiter.Text = ""
      End If
    Changed = True
  End With
End Sub

Private Sub chkAuditChangesOnly_Click()
  Changed = True
End Sub

Private Sub chkPreserveTransformPath_Click()
  Changed = True
End Sub

Private Sub chkPreserveXSDPath_Click()
  Changed = True
End Sub

Private Sub chkSplitFile_Click()
  
  If chkSplitFile.Value = vbChecked Then
    lblSplitFileSize.Enabled = True
    spnSplitFileSize.Enabled = True
    spnSplitFileSize.BackColor = &H80000005
  Else
    lblSplitFileSize.Enabled = False
    spnSplitFileSize.Enabled = False
    spnSplitFileSize.BackColor = &H8000000F
    spnSplitFileSize.Value = 0
  End If
  
  Changed = True
   
End Sub

Private Sub chkSplitXMLNodesFile_Click()
  Changed = True
End Sub

Private Sub chkStripDelimiter_Click()
  Changed = True
End Sub

Private Sub cmdAddAllColumns_Click()

  Dim pstrRow As String
  Dim pfrmColumnEdit As frmExportColumns
  Dim bIsAudited As Boolean
  Dim objExpr As New clsExprExpression
  Dim lngTableID As Long
  Dim lngColumnID As Long
  Dim lngCount As Long
  Dim lngDecimals As Long
  
  Set pfrmColumnEdit = New frmExportColumns
  
  Screen.MousePointer = vbHourglass
  
  With pfrmColumnEdit
    
    .Initialise True, "", 0, 0, , , , Me, , , , , True
    
    'Initialise the edit column form for CMG options
    If optOutputFormat(fmtCMGFile).Value And mbCMGExportFieldCode Then
      .SetCMGOptions ("")
    End If
    
    .Show vbModal
    
    If Not .Cancelled Then
            
      Changed = True
      mblnBaseTableSpecificChanged = True
      bIsAudited = False

      lngTableID = .cboFromTable.ItemData(.cboFromTable.ListIndex)
      
      For lngCount = 1 To .cboFromColumn.ListCount - 1
      
        lngColumnID = .cboFromColumn.ItemData(lngCount)
      
        lngDecimals = 0
        If datGeneral.GetDataType(lngTableID, lngColumnID) = sqlNumeric Then
          lngDecimals = .GetColumnDecimals(lngColumnID)
        End If
      
        pstrRow = "C" & vbTab & _
                  CStr(lngTableID) & vbTab & _
                  CStr(lngColumnID) & vbTab & _
                  .cboFromTable.Text & "." & .cboFromColumn.List(lngCount) & vbTab & _
                  CStr(.GetColumnSize(lngColumnID) + IIf(lngDecimals > 0, 1, 0)) & vbTab & _
                  vbNullString & vbTab & _
                  IIf(datGeneral.IsColumnAudited(lngColumnID), "1", "0") & vbTab

                  '.GetColumnSize(lngColumnID) & vbTab & _

        If datGeneral.GetDataType(lngTableID, lngColumnID) = sqlNumeric Then
          pstrRow = pstrRow & CStr(lngDecimals)
        Else
          pstrRow = pstrRow & vbNullString
        End If

        'pstrRow = pstrRow & vbTab & _
                  .cboFromTable.Text & "." & .cboFromColumn.List(lngCount)
        pstrRow = pstrRow & vbTab & _
                  Left(.cboFromTable.Text & "." & .cboFromColumn.List(lngCount), 50)
        
        'NPG20080620 Fault 13241
        pstrRow = pstrRow & vbTab & 0


        grdColumns.AddItem pstrRow

      Next

      'cmdEditColumn.Enabled = True
      'cmdDeleteColumn.Enabled = True
      'cmdClearColumn.Enabled = True

    End If
  
    With grdColumns
      .SelBookmarks.RemoveAll
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End With

  End With

  Unload pfrmColumnEdit
  Set pfrmColumnEdit = Nothing
  UpdateButtonStatus
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub cmdFooterText_Click()

  Dim frmFooter As frmExportHeaderFooter

  Set frmFooter = New frmExportHeaderFooter

  With frmFooter
    .IsHeader = True
    .Text = txtCustomFooter.Text
    .Initialise
    .Show vbModal
    
    If Not .Cancelled Then
      txtCustomFooter.Text = .Text
      Changed = True
    End If
  
  End With
  
End Sub

Private Sub cmdHeaderText_Click()

  Dim frmHeader As frmExportHeaderFooter

  Set frmHeader = New frmExportHeaderFooter

  With frmHeader
    .IsHeader = True
    .Text = txtCustomHeader.Text
    .Initialise
    .Show vbModal
    
    If Not .Cancelled Then
      txtCustomHeader.Text = .Text
      Changed = True
    End If
  
  End With

End Sub

Private Sub cmdTransformFile_Click()

  Dim cd1 As CommonDialog
  
  On Local Error GoTo LocalErr
  
  Set cd1 = frmMain.CommonDialog1
    
  With cd1
    .Filter = "XSLT File (*.xslt)|*.xslt"
    .FileName = txtTransformFile.Text
    If txtTransformFile.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    End If

    .CancelError = True
    .DialogTitle = "Transformation file name"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
        
    .ShowSave

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        COAMsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
      Else
        txtTransformFile.Text = .FileName    'activates the change event
      End If
    End If
  End With

  Changed = True

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    COAMsgBox Err.Description, vbCritical
  End If


End Sub

Private Sub cmdTransformFileClear_Click()
  txtTransformFile.Text = vbNullString
  Changed = True
End Sub

Private Sub cmdXSDFile_Click()

  Dim cd1 As CommonDialog
  
  On Local Error GoTo LocalErr
  
  Set cd1 = frmMain.CommonDialog1
    
  With cd1
    .Filter = "XSD File (*.xsd)|*.xsd"
    .FileName = txtXSDFilename.Text
    If txtXSDFilename.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    End If

    .CancelError = True
    .DialogTitle = "XSD file name"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
        
    .ShowSave

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        COAMsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
      Else
        txtXSDFilename.Text = .FileName    'activates the change event
      End If
    End If
  End With

  Changed = True

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    COAMsgBox Err.Description, vbCritical
  End If

End Sub

Private Sub cmdXSDFileClear_Click()
  txtXSDFilename.Text = vbNullString
  Changed = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

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

Private Sub grdColumns_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  UpdateButtonStatus
End Sub

Private Sub grdExportOrder_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  
  If Not mblnReadOnly Then
  
    With grdExportOrder
      ' Set the styleSet of the rows to show which is selected.
      For iLoop = 0 To .Rows - 1
        If iLoop = .Row Then
          .Columns(1).CellStyleSet "ssetActive", iLoop
          .Columns(2).CellStyleSet "ssetActive", iLoop
        Else
          .Columns(1).CellStyleSet "ssetDormant", iLoop
          .Columns(2).CellStyleSet "ssetDormant", iLoop
        End If
      Next iLoop

      ' Activate the 'values' column.
      If .Col = 1 Then
        .Col = 0
      End If
      
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End With
    
  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub spnSplitFileSize_Change()
  Changed = True
End Sub

Private Sub txtDelimiter_Change()
  Changed = True
End Sub

Private Sub txtDelimiter_GotFocus()

  With txtDelimiter
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
End Sub

Private Sub txtEmailGroup_Change()
  Changed = True
End Sub

Private Sub RefreshHeaderOptions()
  With txtCustomHeader
    .Enabled = (cboHeaderOptions.ListIndex = 2 Or cboHeaderOptions.ListIndex = 3)   'Visible if custom header selected
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With
  cmdHeaderText.Enabled = txtCustomHeader.Enabled
  CheckIfOmitHeaderEnabled
End Sub

Private Sub cboHeaderOptions_Click()
  txtCustomHeader.Text = ""
  RefreshHeaderOptions
  Changed = True
End Sub

Private Sub CheckIfOmitHeaderEnabled()
  
  Dim blnEnabled As Boolean
  
  blnEnabled = False
  If cboSaveExisting.ListIndex <> -1 Then
    blnEnabled = (cboHeaderOptions.ListIndex > 0 And cboSaveExisting.ItemData(cboSaveExisting.ListIndex) = 3)
  End If

  chkOmitHeader.Enabled = blnEnabled
  If Not blnEnabled And Not mblnLoading Then
    chkOmitHeader.Value = vbUnchecked
  End If

End Sub

Private Sub RefreshFooterOptions()
  With txtCustomFooter
    .Enabled = (cboFooterOptions.ListIndex = 2) 'Visible if custom header selected
    .BackColor = IIf(cboFooterOptions.ListIndex = 2, vbWindowBackground, vbButtonFace)
  End With
  cmdFooterText.Enabled = txtCustomFooter.Enabled
End Sub

Private Sub cboFooterOptions_Click()
  txtCustomFooter.Text = vbNullString
  RefreshFooterOptions
  Changed = True
End Sub

Private Sub cboParentFields_Click()
  Changed = True
End Sub

Private Sub cboSaveExisting_Click()
  CheckIfOmitHeaderEnabled
  Changed = True
End Sub

Private Sub chkForceHeader_Click()
  Changed = True
End Sub

Private Sub chkOmitHeader_Click()
  Changed = True
End Sub

Private Sub chkQuotes_Click()
  Changed = True
End Sub

Private Sub chkUpdateAuditPointer_Click()
  Changed = True
End Sub

Private Sub cmdCancel_Click()

  Unload Me
  
'  Dim pintAnswer As Integer
'
'  If changed = True Then
'
'    pintAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, "Export")
'
'    If pintAnswer = vbYes Then
'      cmdOK_Click
'      Exit Sub
'    ElseIf pintAnswer = vbCancel Then
'      Exit Sub
'    End If
'
'  End If
'
'  Me.Hide
  
End Sub

Private Sub cmdOK_Click()

  If Changed = True Then
    If Not ValidateDefinition Then Exit Sub
    If Not SaveDefinition Then Exit Sub
  End If
  
  Me.Hide
  
End Sub

Private Sub cmdParent1Picklist_Click()
  GetPicklist txtParent1, txtParent1Picklist

End Sub

Private Sub cmdParent2Picklist_Click()
  GetPicklist txtParent2, txtParent2Picklist

End Sub

Private Sub cmdSortMoveDown_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdExportOrder.AddItemRowIndex(grdExportOrder.Bookmark)
  strSourceRow = grdExportOrder.Columns(0).Text & vbTab & grdExportOrder.Columns(1).Text & vbTab & grdExportOrder.Columns(2).Text
  
  intDestinationRow = intSourceRow + 1
  grdExportOrder.MoveNext
  strDestinationRow = grdExportOrder.Columns(0).Text & vbTab & grdExportOrder.Columns(1).Text & vbTab & grdExportOrder.Columns(2).Text
  
  grdExportOrder.RemoveItem intDestinationRow
  grdExportOrder.RemoveItem intSourceRow
  
  grdExportOrder.AddItem strDestinationRow, intSourceRow
  grdExportOrder.AddItem strSourceRow, intDestinationRow
  
  grdExportOrder.SelBookmarks.RemoveAll
  'grdExportOrder.MoveNext
  '
  ' JDM - 14/11/01 - Fault 3159 - Bookmark of wrong grid - someone was careless with cut'n'paste methinks...
  grdExportOrder.Bookmark = grdExportOrder.AddItemBookmark(intDestinationRow)
  grdExportOrder.SelBookmarks.Add grdExportOrder.AddItemBookmark(intDestinationRow)
  
  UpdateButtonStatus

  Changed = True
  mblnBaseTableSpecificChanged = True

End Sub

Private Sub cmdSortMoveUp_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdExportOrder.AddItemRowIndex(grdExportOrder.Bookmark)
  strSourceRow = grdExportOrder.Columns(0).Text & vbTab & grdExportOrder.Columns(1).Text & vbTab & grdExportOrder.Columns(2).Text
  
  intDestinationRow = intSourceRow - 1
  grdExportOrder.MovePrevious
  strDestinationRow = grdExportOrder.Columns(0).Text & vbTab & grdExportOrder.Columns(1).Text & vbTab & grdExportOrder.Columns(2).Text
  
  grdExportOrder.AddItem strSourceRow, intDestinationRow
  
  grdExportOrder.RemoveItem intSourceRow + 1
  
  grdExportOrder.SelBookmarks.RemoveAll
  grdExportOrder.SelBookmarks.Add grdExportOrder.AddItemBookmark(intDestinationRow)
  grdExportOrder.MovePrevious
  grdExportOrder.MovePrevious
  UpdateButtonStatus
  
  Changed = True
  mblnBaseTableSpecificChanged = True

End Sub
Private Sub Form_Load()
  SSTab1.Tab = 0
  'mblnEnableSQLTable = (GetSystemSetting("Output", "ExportToSQLTable", 0) = 1)

  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos False, True, True

  grdAccess.RowHeight = 239

  lng_DataCOLUMNWIDTH = 5045.236

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer
  
  If Changed = True Then
    
    pintAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, "Export")
      
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


Private Sub Form_Unload(Cancel As Integer)
  Set mobjOutputDef = Nothing
  frmMain.RefreshMainForm Me, True
End Sub

Private Sub grdColumns_DblClick()

  If grdColumns.BackColorEven <> vbButtonFace Then
    If grdColumns.Rows > 0 Then cmdEditColumn_Click Else cmdNewColumn_Click
  End If

End Sub

Private Sub grdColumns_GotFocus()

  mblnColGrid = True

End Sub

Private Sub grdColumns_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  UpdateButtonStatus
End Sub

Private Sub grdExportOrder_GotFocus()

mblnSortGrid = True

End Sub

Private Sub grdExportOrder_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

UpdateButtonStatus

End Sub

Private Sub optBasePicklist_Click()

  Changed = True

  cmdBasePicklist.Enabled = True

  With txtBaseFilter
    .Text = ""
    .Tag = 0
  End With

  txtBasePicklist.Text = "<None>"
  cmdBaseFilter.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub

Private Sub optBaseFilter_Click()

  Changed = True

  cmdBaseFilter.Enabled = True

  With txtBasePicklist
    .Text = ""
    .Tag = 0
  End With

  txtBaseFilter.Text = "<None>"
  cmdBasePicklist.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded

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

  cmdBaseFilter.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub cmdBasePicklist_Click()

  'changed = True
  GetPicklist cboBaseTable, txtBasePicklist

End Sub

Private Sub cmdBaseFilter_Click()
  
  'changed = True
  GetFilter cboBaseTable, txtBaseFilter

End Sub

Private Sub cmdParent1Filter_Click()
  
  'changed = True
  GetFilter txtParent1, txtParent1Filter

End Sub

Private Sub cmdParent2Filter_Click()
  
  'changed = True
  GetFilter txtParent2, txtParent2Filter

End Sub

Private Sub cmdChildFilter_Click()
  
  'changed = True
  GetFilter cboChild, txtChildFilter

End Sub

Private Sub optOutputFormat_Click(Index As Integer)
  
  Changed = True
  mobjOutputDef.FormatClick Index
  
  TextOptionsStatus Index
  ShowRelevantColumns
  
  cmdAddAllColumns.Enabled = (Index <> fmtXML)
  
  'Select Case Index
  'Case fmtCSV
  '  TextOptionsStatus "D"
  'Case fmtFixedLengthFile
  '  TextOptionsStatus "F"
  'Case fmtExcelWorksheet
  '  TextOptionsStatus "X"
  'Case fmtCMGFile
  '  TextOptionsStatus "C"
  'Case fmtSQLTable
  '  TextOptionsStatus "S"
  'End Select

End Sub

Private Sub chkDestination_Click(Index As Integer)
  mobjOutputDef.DestinationClick Index
  Changed = True
End Sub

Private Sub optParent1AllRecords_Click()

  Changed = True

  With txtParent1Picklist
    .Text = ""
    .Tag = 0
  End With
  cmdParent1Picklist.Enabled = False

  With txtParent1Filter
    .Text = ""
    .Tag = 0
  End With
  cmdParent1Filter.Enabled = False

  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub optParent1Filter_Click()
  Changed = True

  cmdParent1Filter.Enabled = True

  With txtParent1Picklist
    .Text = ""
    .Tag = 0
  End With

  txtParent1Filter.Text = "<None>"
  cmdParent1Picklist.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub optParent1Picklist_Click()

  Changed = True

  cmdParent1Picklist.Enabled = True

  With txtParent1Filter
    .Text = ""
    .Tag = 0
  End With

  txtParent1Picklist.Text = "<None>"
  cmdParent1Filter.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub optParent2AllRecords_Click()

  Changed = True

  With txtParent2Picklist
    .Text = ""
    .Tag = 0
  End With
  cmdParent2Picklist.Enabled = False

  With txtParent2Filter
    .Text = ""
    .Tag = 0
  End With
  cmdParent2Filter.Enabled = False

  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub optParent2Filter_Click()
  Changed = True

  cmdParent2Filter.Enabled = True

  With txtParent2Picklist
    .Text = ""
    .Tag = 0
  End With

  txtParent2Filter.Text = "<None>"
  cmdParent2Picklist.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub optParent2Picklist_Click()

  Changed = True

  cmdParent2Picklist.Enabled = True

  With txtParent2Filter
    .Text = ""
    .Tag = 0
  End With

  txtParent2Picklist.Text = "<None>"
  cmdParent2Filter.Enabled = False
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub spnMaxRecords_Change()

  Changed = True
  lblMaxRecordsAll.Visible = spnMaxRecords.Value = 0

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
  
  EnableDisableTabControls
  
End Sub


Private Sub txtCustomFooter_Change()
  Changed = True
End Sub

Private Sub txtCustomHeader_Change()
  Changed = True
End Sub

Private Sub txtEmailSubject_Change()
  Changed = True
End Sub

Private Sub txtEmailAttachAs_Change()
  Changed = True
End Sub


Private Sub txtFileExportCode_Change()

  Changed = True

End Sub

Private Sub txtFilename_Change()
  Changed = True
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

Private Sub txtDesc_GotFocus()
  With txtDesc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDesc_Change()

  Changed = True

End Sub

Private Sub GetPicklist(ctlSource As Control, ctlTarget As Control)

  'Purpose : Show the picklist selector form and populate the relevant control
  'Input   : None
  'Output  : None

  Dim plngParent As Long
  Dim pblnExit As Boolean
  Dim pfrmPicklist As frmPicklists
  Dim lngTableID As Long
  Dim rsTemp As Recordset
  Dim blnHiddenPicklist As Boolean

  Screen.MousePointer = vbHourglass

  pblnExit = False

  If TypeOf ctlSource Is TextBox Then
    lngTableID = ctlSource.Tag
  Else
    lngTableID = ctlSource.ItemData(ctlSource.ListIndex)
  End If
   
  With frmDefSel
    .SelectedUtilityType = utlPicklist
    .TableID = lngTableID
    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(ctlTarget.Tag) > 0 Then
      .SelectedID = Val(ctlTarget.Tag)
    End If
  End With
  
  Do While Not pblnExit

    If frmDefSel.ShowList(utlPicklist) Then
      frmDefSel.Show vbModal
  
      Select Case frmDefSel.Action
        Case edtAdd
          Set pfrmPicklist = New frmPicklists
          With pfrmPicklist
            If .InitialisePickList(True, False, lngTableID) Then
              .Show vbModal
            End If
            frmDefSel.SelectedID = .SelectedID
            Unload pfrmPicklist
            Set pfrmPicklist = Nothing
          End With
  
        Case edtEdit
          Set pfrmPicklist = New frmPicklists
          With pfrmPicklist
            If .InitialisePickList(False, frmDefSel.FromCopy, lngTableID, frmDefSel.SelectedID) Then
              .Show vbModal
            End If
            If frmDefSel.FromCopy And .SelectedID > 0 Then
              frmDefSel.SelectedID = .SelectedID
            End If
            Unload pfrmPicklist
            Set pfrmPicklist = Nothing
          End With
  
        'MH20050728 Fault 10232
        Case edtPrint
          Set pfrmPicklist = New frmPicklists
          pfrmPicklist.PrintDef frmDefSel.TableID, frmDefSel.SelectedID
          Unload pfrmPicklist
          Set pfrmPicklist = Nothing
        
        Case edtSelect
          Changed = True
  
          ctlTarget.Text = IIf(Len(frmDefSel.SelectedText) = 0, "<None>", frmDefSel.SelectedText)
          ctlTarget.Tag = frmDefSel.SelectedID
          pblnExit = True
      
        Case 0
          pblnExit = True
      
      End Select
    End If

  Loop
    
  Set frmDefSel = Nothing
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub GetFilter(ctlSource As Control, ctlTarget As Control)
  'Purpose : Show the expression.dll form and populate the relevant control
  'Input   : Source control and Target control - used to know which tags/text/listindex
  '          properties to set once an expression has been selected/cleared.
  'Output  : None

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
        
        If ctlTarget.Tag <> .ExpressionID Then
          Changed = True
          mblnBaseTableSpecificChanged = True
        End If
  
        ctlTarget.Tag = .ExpressionID
      End If
    End If
  End With
  
  Set objExpression = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub

Private Sub cboChild_Click()


  If mblnLoading = True Then Exit Sub
  
  ' If no change the exit sub, else set changed flag
  If cboChild.Text = mstrChildTable Then
    Exit Sub
  Else
    Changed = True
  End If
  
  
  ' Check if any columns in the export definition are from the table that was
  ' previously selected in the child combo box. If so, prompt user for action.
  
  Select Case AnyChildColumnsUsed(mlngChildTable)
    Case 1: ' Child cols used but user has aborted the change
      SetComboText cboChild, mstrChildTable
    Case Else: ' Child cols are used and user wants to continue with the change
      mstrChildTable = cboChild.Text
      mlngChildTable = cboChild.ItemData(cboChild.ListIndex)
      txtChildFilter.Text = "<None>" '""
      txtChildFilter.Tag = 0
  End Select
  
  ' If a table has been selected, enable filter button
  cmdChildFilter.Enabled = cboChild.ListIndex > 0
  
  UpdateButtonStatus
  ForceDefinitionToBeHiddenIfNeeded
  EnableDisableTabControls
  
End Sub

Public Function AnyChildColumnsUsed(lngTableID As Long) As Integer

  ' Purpose : Checks if any columns from the child table which has just been
  '           deselected have been used in the current export. If so, the user
  '           is prompted. Continuing will delete those columns from the export.
  ' Input   : The Table ID to search for
  ' Output  : 0 not used, 1 used so abort change, 2 used but continue with change

  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim pintCount As Integer
  Dim plngColExprIDToDelete As Long
  Dim pintRowToDelete As Integer
  
  If lngTableID = 0 Then
    AnyChildColumnsUsed = 0
    Exit Function
  End If
  
  With grdColumns
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("TableID").CellText(pvarbookmark) = lngTableID Then pintCount = pintCount + 1
      pintLoop = pintLoop + 1
    Loop
  End With

  If pintCount = 0 Then
    AnyChildColumnsUsed = 0
    Exit Function
  End If

  If COAMsgBox("One or more columns from the '" & datGeneral.GetTableName(lngTableID) & "' table have been included in the current export definition. Changing the child table will remove these columns from the export definition." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, "Export") = vbNo Then
    AnyChildColumnsUsed = 1
    Exit Function
  End If

  ' If we are here then we gonna delete the rows from the table

  pintCount = 0
  pintLoop = 0

  With grdColumns
    .MoveFirst
      Do Until pintLoop >= .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        plngColExprIDToDelete = .Columns("ColExprID").CellText(pvarbookmark)
        If .Columns("TableID").CellText(pvarbookmark) = lngTableID Then
          'delete the row
          pintRowToDelete = .AddItemRowIndex(pvarbookmark)
          
          If .Rows = 1 Then
            .RemoveAll
          Else
            .RemoveItem pintRowToDelete
          End If
          
          'delete corresponding row from sort order if found
          RemoveFromSortOrder plngColExprIDToDelete
          .MoveFirst
          pintLoop = 0
        Else
          'MH20000814 Fault 6005
          'Moved this increment from outside the IF to inside an ELSE
          pintLoop = pintLoop + 1
        End If
      Loop
  End With
  AnyChildColumnsUsed = 2
  
End Function


Private Sub cmdNewColumn_Click()

  'Purpose : Add a new row in the column grid.
  'Input   : None
  'Output  : None

  Dim pstrRow As String
  Dim pfrmColumnEdit As frmExportColumns
  Dim bIsAudited As Boolean
  Dim objExpr As New clsExprExpression
  
  Set pfrmColumnEdit = New frmExportColumns
  
  Screen.MousePointer = vbHourglass
  
  With pfrmColumnEdit
    
    .IsXML = optOutputFormat(fmtXML).Value
    
    .Initialise True, "", 0, 0, , , , Me, , , False, False, False
    
    'Initialise the edit column form for CMG options
    If optOutputFormat(fmtCMGFile).Value And mbCMGExportFieldCode Then
      .SetCMGOptions ("")
    End If
    
    If optOutputFormat(fmtXML).Value Then
      .SetXMLOptions
    End If
    
    .SetConvertCaseOptions (0)
    .Show vbModal
    
    If Not .Cancelled Then
            
      Changed = True
      mblnBaseTableSpecificChanged = True
      bIsAudited = False

      'If .OptFiller Then
      If .optOther Then
        Select Case .cboOther.ListIndex
        Case 0
          pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>"
        Case 1
          pstrRow = "R" & vbTab & "0" & vbTab & "0" & vbTab & "<Carriage Return>"
        Case 2
          pstrRow = "N" & vbTab & "0" & vbTab & "0" & vbTab & "<Record Number>"
        End Select
      ElseIf .optCalculation Then
        pstrRow = "X" & vbTab & "0" & vbTab & .txtCalculation.Tag & vbTab & .txtCalculation.Text
      ElseIf .optText Then
        pstrRow = "T" & vbTab & "0" & vbTab & "0" & vbTab & .txtOther
      Else
        pstrRow = "C" & vbTab & .cboFromTable.ItemData(.cboFromTable.ListIndex) & vbTab & .cboFromColumn.ItemData(.cboFromColumn.ListIndex) & vbTab & .cboFromTable.Text & "." & .cboFromColumn.Text
        bIsAudited = datGeneral.IsColumnAudited(.cboFromColumn.ItemData(.cboFromColumn.ListIndex))
      End If
      
      If .txtLength.Text <> "" Then
        pstrRow = pstrRow & vbTab & .txtLength.Text
      Else
        pstrRow = pstrRow & vbTab & 0
      End If
      
      ' Add the CMG Details
      pstrRow = pstrRow & vbTab & .txtCMGCode & vbTab & IIf(bIsAudited = True, "1", "0")
      
      'TM20011130 Fault 3182
      If .optCalculation Then
        If .txtCalculation.Tag > 0 Then
          objExpr.ExpressionID = .txtCalculation.Tag
          objExpr.ConstructExpression
          
          'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
          ' when all we need is the expression return type (time saving measure).
          objExpr.ValidateExpression True, True
          
          If (objExpr.ReturnType = giEXPRVALUE_NUMERIC Or _
                              objExpr.ReturnType = giEXPRVALUE_BYREF_NUMERIC) Then
            pstrRow = pstrRow & vbTab & .spnDec.Text
          Else
            pstrRow = pstrRow & vbTab & vbNullString
          End If
        End If
        Set objExpr = Nothing
      
      ElseIf .optTable Then
        If datGeneral.GetDataType(.cboFromTable.ItemData(.cboFromTable.ListIndex), .cboFromColumn.ItemData(.cboFromColumn.ListIndex)) = sqlNumeric Then
          pstrRow = pstrRow & vbTab & .spnDec.Text
        Else
          pstrRow = pstrRow & vbTab & vbNullString
        End If

      Else
        pstrRow = pstrRow & vbTab & vbNullString
      End If

      pstrRow = pstrRow & vbTab & .txtHeading.Text
      
      'NPG20071213 Fault 12867
      pstrRow = pstrRow & vbTab & .cboConvCase.ListIndex
      
      'NPG20080617 Suggestion S000816
      pstrRow = pstrRow & vbTab & IIf(.chkSuppressNulls.Value = 1, True, False)
      
      With grdColumns
        .AddItem pstrRow
        .MoveLast
        .SelBookmarks.Add .Bookmark
      End With
      
      cmdEditColumn.Enabled = True
      cmdDeleteColumn.Enabled = True
      cmdClearColumn.Enabled = True
    
    End If
  
  End With
  
  Unload pfrmColumnEdit
  Set pfrmColumnEdit = Nothing
  UpdateButtonStatus
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub


Private Sub cmdEditColumn_Click()
  
  'Purpose : Edit the current row in the column grid.
  'Input   : None
  'Output  : None
  
  Dim pstrRow As String
  Dim plngRow As Long
  Dim pfrmColumnEdit As frmExportColumns
  Dim lID As Long
  Dim iCount As Integer
  Dim fFoundInSortOrder As Boolean
  Dim bIsAudited As Boolean

  Dim objExpr As New clsExprExpression

  Screen.MousePointer = vbHourglass
  Set pfrmColumnEdit = New frmExportColumns
   
  With grdColumns
      
    plngRow = .AddItemRowIndex(.Bookmark)
    
    pfrmColumnEdit.IsXML = optOutputFormat(fmtXML).Value

    ' Pass in CMG Codes
    If optOutputFormat(fmtCMGFile).Value = True And mbCMGExportFieldCode Then
      pfrmColumnEdit.SetCMGOptions (Trim(.Columns("CMG Code").Value))
    End If
      
    If optOutputFormat(fmtXML).Value Then
      pfrmColumnEdit.SetXMLOptions
    End If
    
    Select Case .Columns("Type").Text
      
      Case "C":
        lID = .Columns("ColExprID").Value
        pfrmColumnEdit.Initialise False _
                                  , "C" _
                                  , .Columns("TableID").Value _
                                  , .Columns("ColExprID").Value _
                                  , Left(.Columns("Data").Text _
                                  , InStr(.Columns("Data").Text, ".") - 1) _
                                  , Mid(.Columns("Data").Text _
                                  , InStr(.Columns("Data").Text, ".") + 1 _
                                  , Len(.Columns("Data").Text)) _
                                  , .Columns("Length").Value _
                                  , Me _
                                  , CInt(IIf((IsNull(.Columns("Decimals").Value)) Or (.Columns("Decimals").Value = vbNullString), 0, .Columns("Decimals").Value)) _
                                  , .Columns("Heading").Value _
                                  , .Columns("ConvertCase").Value _
                                  , .Columns("SuppressNulls").Value 'NPG20080617 Suggestion S000816
                                  
      Case "X"
        pfrmColumnEdit.Initialise False _
                                  , "X" _
                                  , 0 _
                                  , .Columns("ColExprID").Value _
                                  , _
                                  , .Columns("Data").Text _
                                  , .Columns("Length").Value _
                                  , Me _
                                  , CInt(IIf((IsNull(.Columns("Decimals").Value)) Or (.Columns("Decimals").Value = vbNullString), 0, .Columns("Decimals").Value)) _
                                  , .Columns("Heading").Value _
                                  , .Columns("ConvertCase").Value _
                                  , .Columns("SuppressNulls").Value 'NPG20080617 Suggestion S000816
      Case "T"
        pfrmColumnEdit.Initialise False _
                                  , "T" _
                                  , 0 _
                                  , 0 _
                                  , .Columns("Data").Text _
                                  , _
                                  , .Columns("Length").Value _
                                  , Me _
                                  , CInt(IIf((IsNull(.Columns("Decimals").Value)) Or (.Columns("Decimals").Value = vbNullString), 0, .Columns("Decimals").Value)) _
                                  , .Columns("Heading").Value
      Case "F"
        pfrmColumnEdit.Initialise False _
                                  , "F" _
                                  , 0 _
                                  , 0 _
                                  , "" _
                                  , _
                                  , .Columns("Length").Value _
                                  , Me _
                                  , CInt(IIf((IsNull(.Columns("Decimals").Value)) Or (.Columns("Decimals").Value = vbNullString), 0, .Columns("Decimals").Value)) _
                                  , .Columns("Heading").Value
      Case "R"
        pfrmColumnEdit.Initialise False _
                                  , "R" _
                                  , 0 _
                                  , 0 _
                                  , "" _
                                  , _
                                  , .Columns("Length").Value _
                                  , Me _
                                  , CInt(IIf((IsNull(.Columns("Decimals").Value)) Or (.Columns("Decimals").Value = vbNullString), 0, .Columns("Decimals").Value)) _
                                  , .Columns("Heading").Value
      Case "N"
        pfrmColumnEdit.Initialise False _
                                  , "N" _
                                  , 0 _
                                  , 0 _
                                  , "" _
                                  , _
                                  , .Columns("Length").Value _
                                  , Me _
                                  , CInt(IIf((IsNull(.Columns("Decimals").Value)) Or (.Columns("Decimals").Value = vbNullString), 0, .Columns("Decimals").Value)) _
                                  , .Columns("Heading").Value
    
    End Select
    
  End With
  
  With pfrmColumnEdit
    
    .Show vbModal
    
    If Not .Cancelled Then
      
      'RH 29/11/00 - BUG 1463
      'Check if in sort order...if so, prompt user for action, remove from sort
      'order if necessary.
      If .optTable Then
        
'        If grdExportOrder.Rows > 0 Then
'          For iCount = 0 To (grdExportOrder.Rows)
'            grdExportOrder.Bookmark = grdExportOrder.AddItemBookmark(iCount)
'            If grdExportOrder.Columns("ColExprID").Text = lID Then
'              If .cboFromColumn.ItemData(.cboFromColumn.ListIndex) <> lID Then
'                fFoundInSortOrder = True
'                Exit For
'              End If
'            End If
'          Next iCount
'
'          If fFoundInSortOrder Then
        If IsUsedInSortOrder(grdColumns.Columns("ColExprID").Value) = True Then
            
            If COAMsgBox("You have changed a column that is used in the export sort order." & vbCrLf & _
                      "Continuing will remove the old column from the sort order." & vbCrLf & _
                      "Do you wish to continue ?", vbYesNo + vbQuestion, app.title) = vbNo Then
              Exit Sub
            End If
            RemoveFromSortOrder lID
'          End If
'        End If
        End If
      End If
      '###
      
      Changed = True
      mblnBaseTableSpecificChanged = True
      bIsAudited = False

      'If .OptFiller Then
      If .optOther Then
        Select Case .cboOther.ListIndex
        Case 0
          pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>"
        Case 1
          pstrRow = "R" & vbTab & "0" & vbTab & "0" & vbTab & "<Carriage Return>"
        Case 2
          pstrRow = "N" & vbTab & "0" & vbTab & "0" & vbTab & "<Record Number>"
        End Select
      ElseIf .optCalculation Then
        pstrRow = "X" & vbTab & "0" & vbTab & .txtCalculation.Tag & vbTab & .txtCalculation.Text
      ElseIf .optText Then
        pstrRow = "T" & vbTab & "0" & vbTab & "0" & vbTab & .txtOther
      Else
        pstrRow = "C" & vbTab & .cboFromTable.ItemData(.cboFromTable.ListIndex) & vbTab & .cboFromColumn.ItemData(.cboFromColumn.ListIndex) & vbTab & .cboFromTable.Text & "." & .cboFromColumn.Text
        bIsAudited = datGeneral.IsColumnAudited(.cboFromColumn.ItemData(.cboFromColumn.ListIndex))
      End If
      
      If .txtLength.Text <> "" Then
        pstrRow = pstrRow & vbTab & .txtLength.Text
      Else
        pstrRow = pstrRow & vbTab & 0
      End If

      ' Add the CMG Details
      pstrRow = pstrRow & vbTab & .txtCMGCode & vbTab & IIf(bIsAudited = True, "1", "0")

      'TM20011130 Fault 3182
      If .optCalculation Then
        If .txtCalculation.Tag > 0 Then
          objExpr.ExpressionID = .txtCalculation.Tag
          objExpr.ConstructExpression
          
          'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
          ' when all we need is the expression return type (time saving measure).
          objExpr.ValidateExpression True, True
          
          If (objExpr.ReturnType = giEXPRVALUE_NUMERIC Or _
                              objExpr.ReturnType = giEXPRVALUE_BYREF_NUMERIC) Then
            pstrRow = pstrRow & vbTab & .spnDec.Text
          Else
            pstrRow = pstrRow & vbTab & vbNullString
          End If
        End If
        Set objExpr = Nothing
      
      ElseIf .optTable Then
        If datGeneral.GetDataType(.cboFromTable.ItemData(.cboFromTable.ListIndex), .cboFromColumn.ItemData(.cboFromColumn.ListIndex)) = sqlNumeric Then
          pstrRow = pstrRow & vbTab & .spnDec.Text
        Else
          pstrRow = pstrRow & vbTab & vbNullString
        End If

      Else
        pstrRow = pstrRow & vbTab & vbNullString
      End If

      'MH20030120
      pstrRow = pstrRow & vbTab & .txtHeading.Text
      
      'NPG20071213 Fault 12867
      pstrRow = pstrRow & vbTab & .cboConvCase.ListIndex
      
      'NPG20080617 Suggestion S000816
      pstrRow = pstrRow & vbTab & IIf(.chkSuppressNulls.Value = 1, True, False)
      
      With grdColumns
        .RemoveItem plngRow
        .AddItem pstrRow, plngRow
      End With
    
    End If
  
  End With
  
    grdColumns.Bookmark = grdColumns.AddItemBookmark(plngRow)
    grdColumns.SelBookmarks.Add grdColumns.AddItemBookmark(plngRow)
  
  Unload pfrmColumnEdit
  Set pfrmColumnEdit = Nothing
  UpdateButtonStatus
  ForceDefinitionToBeHiddenIfNeeded

End Sub


Private Sub cmdDeleteColumn_Click()

  'Purpose : Remove the selected row from the order grid. First checks if the selected
  '          item is defined in the export sort order. If so, user is prompted for action.
  '          If they continue, the column is removed from both grids.
  'Input   : None
  'Output  : None
  
  Dim intRow As Integer
  Dim lRow As Long
  
  Changed = True
  mblnBaseTableSpecificChanged = True
 
  If IsUsedInSortOrder(grdColumns.Columns("ColExprID").Value) = True Then
    If COAMsgBox("The '" & grdColumns.Columns("Data").Value & "' column is defined in the export sort order." & vbCrLf & "If you delete the column from the export, it will be removed from the sort order." & vbCrLf & "Do you wish to continue ?", vbQuestion + vbYesNo, "Export") = vbYes Then
      RemoveFromSortOrder (grdColumns.Columns("ColExprID").Value)
      UpdateButtonStatus
    Else
      Exit Sub
    End If
  End If
  
  If grdColumns.Rows = 1 Then
    grdColumns.RemoveAll
  Else

    ' Store the row to be deleted
    lRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
    
    ' Delete the row
    grdColumns.RemoveItem grdColumns.AddItemRowIndex(grdColumns.Bookmark)

    If grdColumns.Rows > 0 Then
      If lRow < grdColumns.Rows Then
        grdColumns.SelBookmarks.Add grdColumns.GetBookmark(lRow)
      ElseIf lRow = grdColumns.Rows Then
        grdColumns.MoveLast
        grdColumns.SelBookmarks.Add grdColumns.Bookmark
      End If
    End If
  
  End If
  
  UpdateButtonStatus
    
  ForceDefinitionToBeHiddenIfNeeded
    

'#####################

''    intRow = Me.grdColumns.AddItemRowIndex(Me.grdColumns.Bookmark) 'Me.grdColumns.Row
''    Me.grdColumns.RemoveItem intRow
''    Me.grdColumns.MovePrevious
''    Me.grdColumns.SelBookmarks.Add grdColumns.Bookmark
'    With grdColumns
'
'      lRow = .AddItemRowIndex(.Bookmark)
'      .RemoveItem lRow
'
'      If .Rows = 0 Then
'        UpdateButtonStatus
'        Else
'        If lRow < .Rows Then
'
'          ' RH 29/09/00 - BUG 1037 - DONT THINK WE NEED TO DO THIS
'          '.Bookmark = lRow
'          '.MoveNext
'        Else
'          .Bookmark = (.Rows - 1)
'        End If
'        .SelBookmarks.Add .Bookmark
'      End If
'
'    End With
'
'  End If
'
'  UpdateButtonStatus
'
'  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Function IsUsedInSortOrder(plngColExprID As Long) As Boolean

  'Purpose : Check if the specified column is used in the sort order.
  'Input   : ColExprID
  'Output  : True/False
  
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim pintPrevRow As Integer
  Dim lngColumnCount As Long
  Dim pvarRestoreBookmark As Variant


  IsUsedInSortOrder = False
  
  
  'Check if this column exists more than once in the columns list
  pintLoop = 0
  lngColumnCount = 0
  With grdColumns
    pvarRestoreBookmark = .Bookmark
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("Type").CellText(pvarbookmark) = "C" And _
         .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
        lngColumnCount = lngColumnCount + 1
        If lngColumnCount > 1 Then
          .Bookmark = pvarRestoreBookmark
          Exit Function
        End If
      End If
      pintLoop = pintLoop + 1
    Loop
    .Bookmark = pvarRestoreBookmark
  End With
  
  
  
  With grdExportOrder
  
    pintPrevRow = .AddItemRowIndex(.Bookmark)
    pintLoop = 0
    .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
          IsUsedInSortOrder = True
          pvarbookmark = .GetBookmark(pintPrevRow)
          .Bookmark = pvarbookmark
          .SelBookmarks.Add .Bookmark
          Exit Function
        End If
        pintLoop = pintLoop + 1
      Loop
  End With

End Function

Private Sub RemoveFromSortOrder(plngColExprID As Long)

  'Purpose : Removes the specified column from the sort order.
  'Input   : ColExprID
  'Output  : None
  
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim pintRowToDelete As Integer
  
  With grdExportOrder
  
    .MoveFirst
    
    If .Rows = 1 Then
      .RemoveAll
      Exit Sub
    End If
    
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
          pintRowToDelete = .AddItemRowIndex(pvarbookmark)
          .RemoveItem pintRowToDelete
          If .Rows > 0 Then
            .MoveLast
            .Bookmark = pvarbookmark
            .SelBookmarks.Add .Bookmark
          End If

          Exit Sub
        End If
        pintLoop = pintLoop + 1
      Loop
  End With

End Sub

Private Sub cmdClearColumn_Click()
  
  'Purpose : Check the user really wishes to clear the column grid. This will
  '          automatically clear the export order grid too.
  'Input   : None
  'Output  : None
  If grdExportOrder.Rows > 0 Then
  
    If COAMsgBox("Clearing all export columns will automatically clear the sort order definition." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, "Export") = vbYes Then
      grdColumns.RemoveAll
      grdExportOrder.RemoveAll
      UpdateButtonStatus
      ForceDefinitionToBeHiddenIfNeeded
      Changed = True
    End If

  Else
    
    If COAMsgBox("Are you sure you wish to clear all columns / calculations from this definition ?", vbYesNo + vbQuestion, "Export") = vbYes Then
      grdColumns.RemoveAll
      grdExportOrder.RemoveAll
      UpdateButtonStatus
      ForceDefinitionToBeHiddenIfNeeded
      Changed = True
    End If
  
  End If
  
End Sub


Private Sub cmdNewOrder_Click()

  'Purpose : Add a new row in the order grid.
  'Input   : None
  'Output  : None
  
  Dim pfrmOrderEdit As New frmExportOrder
  Dim intOriginalRow As Integer
  
  ' Store the bookmark of the export order
  intOriginalRow = grdExportOrder.AddItemRowIndex(grdExportOrder.Bookmark)
  
  If pfrmOrderEdit.Initialise(True, Me, 0, "") = True Then
    pfrmOrderEdit.Show vbModal
    mblnBaseTableSpecificChanged = True
  Else
    ' Reset the bookmark of the export order
    grdExportOrder.SelBookmarks.RemoveAll
    grdExportOrder.Bookmark = grdExportOrder.AddItemBookmark(intOriginalRow)
    grdExportOrder.SelBookmarks.Add grdExportOrder.AddItemBookmark(intOriginalRow)
  End If
  
  Set pfrmOrderEdit = Nothing
  
  UpdateButtonStatus

End Sub

Private Sub cmdEditOrder_Click()

  'Purpose : Edit the current row in the order grid.
  'Input   : None
  'Output  : None
  
  Dim pfrmOrderEdit As New frmExportOrder
  Dim intOriginalRow As Integer
  
  ' Store the bookmark of the export order
  intOriginalRow = grdExportOrder.AddItemRowIndex(grdExportOrder.Bookmark)
  
  If pfrmOrderEdit.Initialise(False, Me, grdExportOrder.Columns("ColExprID").CellValue(grdExportOrder.Bookmark), grdExportOrder.Columns("Sort Order").CellText(grdExportOrder.Bookmark)) = True Then
    pfrmOrderEdit.Show vbModal
  End If

  ' Reset the bookmark of the export order
  grdExportOrder.SelBookmarks.RemoveAll
  grdExportOrder.Bookmark = grdExportOrder.AddItemBookmark(intOriginalRow)
  grdExportOrder.SelBookmarks.Add grdExportOrder.AddItemBookmark(intOriginalRow)
  
  mblnBaseTableSpecificChanged = True
  Set pfrmOrderEdit = Nothing
  UpdateButtonStatus
  
End Sub


Private Sub cmdDeleteOrder_Click()

  'Purpose : Remove the selected row from the order grid
  'Input   : None
  'Output  : None

  Dim lRow As Long
  
  If grdExportOrder.Rows = 1 Then
    grdExportOrder.RemoveAll
  Else
    With grdExportOrder
      lRow = .AddItemRowIndex(.Bookmark)
      .RemoveItem lRow
      If .Rows <> 0 Then
        If lRow < .Rows Then
          .Bookmark = lRow
        Else
          .Bookmark = (.Rows - 1)
        End If
        .SelBookmarks.Add .Bookmark
      End If
    End With
  
  End If
  
  UpdateButtonStatus
  Changed = True
  mblnBaseTableSpecificChanged = True
  
End Sub

Private Sub cmdClearOrder_Click()
  
  'Purpose : Check the user really wishes to clear the export order
  'Input   : None
  'Output  : None
  
  If COAMsgBox("Are you sure you wish to clear the sort order ?", vbYesNo + vbQuestion, "Export") = vbYes Then
    grdExportOrder.RemoveAll
    UpdateButtonStatus
    Changed = True
    mblnBaseTableSpecificChanged = True
  End If

End Sub

Private Sub grdExportOrder_DblClick()
  
  'Purpose : If now rows exist in the grid when it is double clicked, a new one
  '          is added, otherwise the current row is edited
  'Input   : None
  'Output  : None
  
  If grdColumns.BackColorEven <> vbButtonFace Then
    If grdExportOrder.Rows > 0 Then cmdEditOrder_Click Else cmdNewOrder_Click
  End If
  
End Sub

Private Sub UpdateButtonStatus()

  'Purpose : Updates all command button status depending on grid.rows etc
  'Input   : None
  'Output  : None
  Dim lngScrollbarSize As Long
  
  'If mblnReadOnly Then Exit Sub

  'Select the highlighted row
  grdColumns.SelBookmarks.RemoveAll
  grdColumns.SelBookmarks.Add grdColumns.Bookmark

  ' Ensure the columns are a snug fit
  If Me.Visible Then
    If grdColumns.VisibleRows < grdColumns.Rows Then
      'MH20050105 Fault 9540
      'grdColumns.ScrollBars = ssScrollBarsAutomatic
      grdColumns.ScrollBars = ssScrollBarsVertical
      lngScrollbarSize = 0
    Else
      grdColumns.ScrollBars = ssScrollBarsNone
      lngScrollbarSize = lng_SCROLLBARWIDTH
    End If
  End If
 
  grdColumns.Columns("Data").Width = lng_DataCOLUMNWIDTH + lngScrollbarSize
  grdColumns.Columns("Length").Width = lng_LengthCOLUMNWIDTH
  grdColumns.Columns("Decimals").Width = lng_DecimalCOLUMNWIDTH
 
'###########
  If Not mblnReadOnly Then
    If grdColumns.Rows = 0 Then
      cmdEditColumn.Enabled = False
      cmdDeleteColumn.Enabled = False
      cmdClearColumn.Enabled = False
      cmdMoveUp.Enabled = False
      cmdMoveDown.Enabled = False
      cmdEditOrder.Enabled = False
      cmdDeleteOrder.Enabled = False
      cmdClearOrder.Enabled = False
      cmdSortMoveUp.Enabled = False
      cmdSortMoveDown.Enabled = False
      Exit Sub
    Else
      If grdColumns.AddItemRowIndex(grdColumns.Bookmark) >= 0 Or mblnColGrid = True Then
        cmdEditColumn.Enabled = True
        cmdDeleteColumn.Enabled = True
        cmdClearColumn.Enabled = True
      End If
    End If
  
    If grdColumns.AddItemRowIndex(grdColumns.Bookmark) < 1 Then
      cmdMoveUp.Enabled = False
    Else
      cmdMoveUp.Enabled = True
    End If
    
    If grdColumns.AddItemRowIndex(grdColumns.Bookmark) = (grdColumns.Rows - 1) Then
      cmdMoveDown.Enabled = False
    Else
      cmdMoveDown.Enabled = True
    End If
  
    If grdColumns.SelBookmarks.Count = 0 Then
      cmdMoveUp.Enabled = False
      cmdMoveDown.Enabled = False
    End If
  End If

'###########

  ' Ensure sort order column tabs are a snug fit
  If Me.Visible Then
    DoEvents
    If grdExportOrder.VisibleRows < grdExportOrder.Rows Then
      'grdExportOrder.ScrollBars = ssScrollBarsAutomatic
      grdExportOrder.ScrollBars = ssScrollBarsVertical
      lngScrollbarSize = 0
    Else
      grdExportOrder.ScrollBars = ssScrollBarsNone
      lngScrollbarSize = lng_SCROLLBARWIDTH
    End If
  End If
  
  grdExportOrder.Columns("Column").Width = lng_SortCOLUMNWIDTH + lngScrollbarSize
  grdExportOrder.Columns("Sort Order").Width = lng_SortOrderCOLUMNWIDTH

  If Not mblnReadOnly Then
    'TM20020828 Fault 4351
    If grdExportOrder.Rows = 1 Then
      grdExportOrder.MoveFirst
    End If
    
    grdExportOrder.SelBookmarks.RemoveAll
    grdExportOrder.SelBookmarks.Add grdExportOrder.Bookmark

    If grdExportOrder.Rows = 0 Then
      cmdEditOrder.Enabled = False
      cmdDeleteOrder.Enabled = False
      cmdClearOrder.Enabled = False
      cmdSortMoveUp.Enabled = False
      cmdSortMoveDown.Enabled = False
      Exit Sub
    Else
      If grdExportOrder.SelBookmarks.Count > 0 Or mblnSortGrid = True Then
        cmdEditOrder.Enabled = True
        cmdDeleteOrder.Enabled = True
        cmdClearOrder.Enabled = True
      End If
    End If
    
    DoEvents
    
    If grdExportOrder.AddItemRowIndex(grdExportOrder.Bookmark) < 1 Then
      cmdSortMoveUp.Enabled = False
    Else
      cmdSortMoveUp.Enabled = grdExportOrder.Rows > 1
    End If
    
    If grdExportOrder.AddItemRowIndex(grdExportOrder.Bookmark) = (grdExportOrder.Rows - 1) Then
      cmdSortMoveDown.Enabled = False
    Else
      cmdSortMoveDown.Enabled = grdExportOrder.Rows > 1
    End If
  
    If (grdExportOrder.SelBookmarks.Count = 0) Or (grdExportOrder.Rows <= 1) Then
      cmdSortMoveUp.Enabled = False
      cmdSortMoveDown.Enabled = False
    End If
  End If

End Sub

Private Sub cboDateFormat_Click()
  Changed = True
End Sub

Private Sub cboDateSeparator_Click()
  Changed = True
End Sub

Private Sub cboDateYearDigits_Click()
  Changed = True
End Sub

Private Sub ClearForNew(Optional bPartialClear As Boolean)
  
  'Purpose : Clear out all fields required to be blank for a new export definition
  'Input   : Optional True/False for if its a partial or complete clear up
  'Output  : None
  
  With Me
    .optBaseAllRecords.Value = True
    .txtBasePicklist.Text = ""
    .txtBasePicklist.Tag = 0
    .txtBaseFilter.Text = ""
    .txtBaseFilter.Tag = 0
    
    .txtParent1.Text = ""
    .txtParent1.Tag = 0
    .txtParent1Picklist.Text = ""
    .txtParent1Picklist.Tag = 0
    .txtParent1Filter.Text = ""
    .txtParent1Filter.Tag = 0
    
    .txtParent2.Text = ""
    .txtParent2.Tag = 0
    .txtParent2Picklist.Text = ""
    .txtParent2Picklist.Tag = 0
    .txtParent2Filter.Text = ""
    .txtParent2Filter.Tag = 0
    
    .cboChild.Clear
    .cboChild.AddItem "<None>"
    .cboChild.ItemData(.cboChild.NewIndex) = 0
    .cboChild.ListIndex = 0
    .txtChildFilter.Text = ""
    .txtChildFilter.Tag = 0
    
    .grdColumns.RemoveAll
    .grdExportOrder.RemoveAll
    
    If bPartialClear Then Exit Sub
    
    .txtName = vbNullString
    .txtDesc = vbNullString
    .txtUserName = gsUserName
    '.optDELIMITED.Value = True
    .optOutputFormat(fmtCSV).Value = True
    .cboDelimiter.ListIndex = 0
    .txtFilename.Text = ""
    '.txtFilename(2).Text = ""
    '.txtSQLTableName.Text = ""
    .cboDateYearDigits.ListIndex = 1
    .cboDateFormat.ListIndex = 0
    'AE20071005 Fault #9614
    'cboDateSeparator.ListIndex = 0
    SetComboText cboDateSeparator, UI.GetSystemDateSeparator
    If cboDateSeparator.ListIndex = -1 Then cboDateSeparator.ListIndex = 0
    'UpdateDate
    .chkQuotes.Value = vbUnchecked
    '.chkAppendToFile = vbUnchecked
    .chkOmitHeader.Value = vbUnchecked
    cboHeaderOptions.ListIndex = 0
    cboFooterOptions.ListIndex = 0

    ControlsDisableAll fraCMGFile, False

  End With
  
End Sub

Private Sub TextOptionsStatus(intFormat As Integer)

  fraDelimFile.Visible = (intFormat <> fmtCMGFile)
  EnableFrame fraDelimFile, (intFormat = fmtCSV)

  fraCMGFile.Visible = (intFormat = fmtCMGFile)
  EnableFrame fraCMGFile, (intFormat = fmtCMGFile)

  fraXML.Visible = (intFormat = fmtXML)
  EnableFrame fraXML, (intFormat = fmtXML)
  EnableControl txtXSDFilename, False
  EnableControl txtTransformFile, False
      
  If intFormat = fmtXML Then
    EnableControl lblHeaderLine, False
    EnableControl cboHeaderOptions, False
    cboHeaderOptions.ListIndex = 2
    EnableControl txtCustomHeader, True
    EnableControl lblFooterLine, False
    EnableControl cboFooterOptions, False
    EnableControl lblCustomFooter, False
    EnableControl txtCustomFooter, False
    EnableControl chkOmitHeader, False
    EnableControl chkForceHeader, False
  Else
  
    EnableControl cboHeaderOptions, True
    EnableControl lblHeaderLine, True
    EnableControl lblCustomHeader, True
    EnableControl txtCustomHeader, Not cboHeaderOptions.ListIndex = 1
    
    EnableControl cboFooterOptions, True
    EnableControl lblFooterLine, True
    EnableControl lblCustomFooter, True
    EnableControl txtCustomFooter, Not cboFooterOptions.ListIndex = 1
    
    EnableControl chkOmitHeader, True
    EnableControl chkForceHeader, True
  End If
  
  EnableFrame fraDateOptions, (intFormat <> fmtXML)
  

End Sub

Private Sub EnableFrame(fraTemp As Control, blnEnabled As Boolean)

  Dim ctl As Control

  ControlsDisableAll fraTemp, blnEnabled
  For Each ctl In fraTemp.Parent
    If Not (TypeOf ctl Is CommonDialog) Then
      If ctl.Container.Name = fraTemp.Name Then
        If TypeOf ctl Is ComboBox Then
          If ctl.ListCount > 0 Then
            ctl.ListIndex = IIf(blnEnabled, 0, -1)
          End If
        ElseIf TypeOf ctl Is CheckBox Then
          If Not blnEnabled Then
            ctl.Value = vbUnchecked
          End If
        End If
      End If
    End If
  Next

End Sub


Public Sub EnableDisableTabControls()

  'Purpose : To enable/disable appropriate frames so tab order is correct
  'Input   : None
  'Output  : None

  If mblnReadOnly Then
    Exit Sub
  End If

  ' TAB 1 CONTROLS
  fraInformation.Enabled = (SSTab1.Tab = 0)
  fraBase.Enabled = (SSTab1.Tab = 0)

  ' TAB 2 CONTROLS
  fraParent1.Enabled = (SSTab1.Tab = 1) And (Len(txtParent1.Text) > 0)
  lblParent1Table.Enabled = fraParent1.Enabled
  lblParent1Records.Enabled = fraParent1.Enabled
  optParent1AllRecords.Enabled = fraParent1.Enabled
  optParent1Picklist.Enabled = fraParent1.Enabled
  optParent1Filter.Enabled = fraParent1.Enabled

  fraParent2.Enabled = (SSTab1.Tab = 1) And (Len(txtParent2.Text) > 0)
  lblParent2Table.Enabled = fraParent2.Enabled
  lblParent2Records.Enabled = fraParent2.Enabled
  optParent2AllRecords.Enabled = fraParent2.Enabled
  optParent2Picklist.Enabled = fraParent2.Enabled
  optParent2Filter.Enabled = fraParent2.Enabled

  fraChild.Enabled = (SSTab1.Tab = 1) And (cboChild.ListCount > 1)
  cboChild.Enabled = fraChild.Enabled
  cboChild.BackColor = IIf(cboChild.Enabled, &H80000005, &H8000000F)
  lblChildTable.Enabled = fraChild.Enabled
  lblChildFilter.Enabled = fraChild.Enabled And (cboChild.ListIndex > 0)
  cmdChildFilter.Enabled = lblChildFilter.Enabled
  If cboChild.ListIndex = 0 Then
    txtChildFilter.Text = ""
    txtChildFilter.Tag = 0
    spnMaxRecords.Value = 0
  End If

  lblMaxRecords.Enabled = lblChildFilter.Enabled
  spnMaxRecords.Enabled = lblChildFilter.Enabled
  spnMaxRecords.BackColor = IIf(spnMaxRecords.Enabled, &H80000005, &H8000000F)
  lblMaxRecordsAll.Enabled = lblChildFilter.Enabled


  ' TAB 3 CONTROLS
  fraColumns.Enabled = (SSTab1.Tab = 2)

  ' TAB 4 CONTROLS
  fraExportOrder.Enabled = (SSTab1.Tab = 3)

  fraDateOptions.Enabled = (SSTab1.Tab = 4)
  fraHeaderOptions.Enabled = (SSTab1.Tab = 4)
  
  ' TAB 5 CONTROLS
  'fraOutput.Enabled = (SSTab1.Tab = 5)
  fraOutputType.Enabled = (SSTab1.Tab = 5)
  fraOutputDestination.Enabled = (SSTab1.Tab = 5)
  'fraOutputFilename.Enabled = (SSTab1.Tab = 5)
  'fraSQLTable.Enabled = (SSTab1.Tab = 5)
  fraDelimFile.Enabled = (SSTab1.Tab = 5)
  fraCMGFile.Enabled = (SSTab1.Tab = 5)
  
  Select Case SSTab1.Tab
  Case 2
    DoEvents
    If (grdColumns.VisibleRows < grdColumns.Rows) And Me.Visible = True Then
      grdColumns.ScrollBars = ssScrollBarsVertical
    Else
      grdColumns.ScrollBars = ssScrollBarsNone
    End If
    grdColumns.Columns("Length").Width = lng_LengthCOLUMNWIDTH
    grdColumns.Columns("Decimals").Width = lng_DecimalCOLUMNWIDTH

    UpdateButtonStatus

  Case 3
    UpdateButtonStatus
  
  Case 4
    RefreshHeaderOptions
    RefreshFooterOptions
  
  Case 5
    With txtDelimiter
        If cboDelimiter.Text = "<Other>" Then
          'If <Other> is selected as a delimiter choice...
          lblOtherDelimiter.Enabled = True
          txtDelimiter.Enabled = True
          txtDelimiter.BackColor = &H80000005
        Else
          lblOtherDelimiter.Enabled = False
          txtDelimiter.Enabled = False
          txtDelimiter.BackColor = &H8000000F
          txtDelimiter.Text = ""
        End If
      'Changed = False
    End With
        
    If chkSplitFile.Value = vbChecked Then
      lblSplitFileSize.Enabled = True
      spnSplitFileSize.Enabled = True
      spnSplitFileSize.BackColor = &H80000005
    Else
      lblSplitFileSize.Enabled = False
      spnSplitFileSize.Enabled = False
      spnSplitFileSize.BackColor = &H8000000F
      spnSplitFileSize.Value = 0
    End If
    
  End Select

End Sub


Private Function ValidateDefinition() As Boolean

  'Purpose : Check all mandatory informaiton is entered and also check that
  '          there are no fields belonging to tables thats arent defined as
  '          either a base, parent or child. Also that no fields in the sort order
  '          are missing in the export. - shouldnt happen, but just in case !
  '          If there is a problem with validation, the program will display
  '          the tab containing the problem to the user.
  'Input   : None
  'Output  : True/False
  
  On Error GoTo ValidateDefinition_ERROR
  
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim rsTemp As Recordset
  Dim strRecSelStatus As String
  Dim strOutputName As String
  
  Dim iCount_Owner As Integer
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobScheduledUserGroups As String
  Dim sHiddenGroups As String
  
  Dim blnContinue As Boolean
  
  ValidateDefinition = False
  fBatchJobsOK = True
  blnContinue = False
  
  ' Check a name has been entered
  If Trim(txtName.Text) = "" Then
    COAMsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  'Check if this definition has been changed by another user
  Call UtilityAmended(utlExport, mlngExportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngExportID = 0
    mblnReadOnly = False
    ForceAccess
  End If
  
  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngExportID) Then
    COAMsgBox "An Export definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    ValidateDefinition = False
    Exit Function
  End If
  
  ' BASE TABLE - If using a picklist, check one has been selected
  If optBasePicklist.Value Then
    If txtBasePicklist.Text = "" Or txtBasePicklist.Tag = "0" Or txtBasePicklist.Tag = "" Then
      COAMsgBox "You must select a picklist, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 0
      cmdBasePicklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' BASE TABLE - If using a filter, check one has been selected
  If optBaseFilter.Value Then
    If txtBaseFilter.Text = "" Or txtBaseFilter.Tag = "0" Or txtBasePicklist.Tag = "" Then
      COAMsgBox "You must select a filter, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 0
      cmdBaseFilter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 1 TABLE - If using a picklist, check one has been selected
  If optParent1Picklist.Value Then
    If txtParent1Picklist.Text = "" Or txtParent1Picklist.Tag = "0" Or txtParent1Picklist.Tag = "" Then
      COAMsgBox "You must select a picklist, or change the record selection for your first parent table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 1
      cmdParent1Picklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 1 TABLE - If using a filter, check one has been selected
  If optParent1Filter.Value Then
    If txtParent1Filter.Text = "" Or txtParent1Filter.Tag = "0" Or txtParent1Filter.Tag = "" Then
      COAMsgBox "You must select a filter, or change the record selection for your first parent table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 1
      cmdParent1Filter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 2 TABLE - If using a picklist, check one has been selected
  If optParent2Picklist.Value Then
    If txtParent2Picklist.Text = "" Or txtParent2Picklist.Tag = "0" Or txtParent2Picklist.Tag = "" Then
      COAMsgBox "You must select a picklist, or change the record selection for your second parent table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 1
      cmdParent2Picklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 2 TABLE - If using a filter, check one has been selected
  If optParent2Filter.Value Then
    If txtParent2Filter.Text = "" Or txtParent2Filter.Tag = "0" Or txtParent2Filter.Tag = "" Then
      COAMsgBox "You must select a filter, or change the record selection for your second parent table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 1
      cmdParent2Filter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  If txtChildFilter.Tag <> 0 Then
    If txtChildFilter.Text = "" Or txtChildFilter.Tag = "0" Or txtChildFilter.Tag = "" Then
      COAMsgBox "You must select a filter, or change the record selection for the child table.", vbExclamation + vbOKOnly, "Export"
      SSTab1.Tab = 1
      cmdChildFilter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
  
  ' Check that there are columns defined in the export definition
  If grdColumns.Rows = 0 Then
    COAMsgBox "You must select at least 1 column for your export.", vbExclamation + vbOKOnly, "Export"
    SSTab1.Tab = 2
    ValidateDefinition = False
    Exit Function
  End If
  
  '  Check that a delimiter is specified if the file format is ASCII Delimited
  If optOutputFormat(fmtCSV) Then
    If cboDelimiter.Text = "<Other>" And Trim(txtDelimiter.Text) = "" Then
      COAMsgBox "You must specify a delimiter for delimited files.", vbExclamation + vbOKOnly, "Import"
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  
  ' XML specific validation
  If optOutputFormat(fmtXML) Then
    If ContainsInvalidXML(txtCustomHeader.Text, False) Then
      COAMsgBox "The XML custom header cannot contain spaces or any of the following characters: ~/\;?$&%@^=*+()|""'`{}[]<>"
      SSTab1.Tab = 4
      ValidateDefinition = False
      Exit Function
    End If
    
    If ContainsInvalidXML(txtXMLDataNodeName.Text, False) Then
      COAMsgBox "The XML custom node name cannot contain spaces or any of the following characters: ~/\;?$&%@^=*+()|""'`{}[]<>"
      SSTab1.Tab = 5
      ValidateDefinition = False
      Exit Function
    End If

  End If
  
  ' Check that at least 1 column has been defined as the export order
  If grdExportOrder.Rows = 0 Then
    COAMsgBox "You must select at least 1 column to order the export by.", vbExclamation + vbOKOnly, "Export"
    SSTab1.Tab = 3
    ValidateDefinition = False
    Exit Function
  End If

  
  If optOutputFormat(fmtCMGFile).Value = True Then
    If cboParentFields.Text = vbNullString Then
      COAMsgBox "No record identifier selected for the CMG file.", vbExclamation, "Export"
      ValidateDefinition = False
      SSTab1.Tab = 5
      Exit Function
    End If
  End If
  
  
  If Not mobjOutputDef.ValidDestination Then
    SSTab1.Tab = 5
    Exit Function
  End If
  
  
  
  ' Check that only tables selected in the combos are in the grid.
  ' THIS SHOULD NEVER HAPPEN, BUT JUST IN CASE...
  grdColumns.MoveFirst
  For pintLoop = 0 To grdColumns.Rows - 1
    pvarbookmark = grdColumns.GetBookmark(pintLoop)
    If grdColumns.Columns("TableID").CellValue(pvarbookmark) <> 0 Then
      If grdColumns.Columns("TableID").CellValue(pvarbookmark) <> Val(cboBaseTable.ItemData(cboBaseTable.ListIndex)) And _
         grdColumns.Columns("TableID").CellValue(pvarbookmark) <> Val(txtParent1.Tag) And _
         grdColumns.Columns("TableID").CellValue(pvarbookmark) <> Val(txtParent2.Tag) And _
         grdColumns.Columns("TableID").CellValue(pvarbookmark) <> Val(cboChild.ItemData(cboChild.ListIndex)) Then
        COAMsgBox "The '" & Left(grdColumns.Columns("Data").CellText(pvarbookmark), InStr(grdColumns.Columns("Data").CellText(pvarbookmark), ".") - 1) & "' table has not been selected as a Base, Parent or Child table, but exists in the export definition." & Chr(10) & "Please either include this table or remove the grid entries referring to it.", vbExclamation, "Export"
        ValidateDefinition = False
        Exit Function
      End If
    End If


    If optOutputFormat(fmtFixedLengthFile).Value = True Then
      If grdColumns.Columns("Length").CellValue(pvarbookmark) = 0 Then
        If grdColumns.Columns("Type").CellValue(pvarbookmark) <> "R" Then
          
          If blnContinue = False Then
            blnContinue = (COAMsgBox("You have selected zero for the length of one or more export columns." & vbCrLf & _
                    "Leaving a column with a length of zero will result in the values not" & vbCrLf & _
                    "appearing." & vbCrLf & vbCrLf & _
                    "Do you wish to continue ?", vbQuestion + vbYesNo, "Export") = vbYes)
            If blnContinue = False Then
              ValidateDefinition = False
              Exit Function
            End If
          End If
        
        End If
      End If
    End If


    'MH21012002
    If grdColumns.Columns("Type").CellText(pvarbookmark) = "N" Then  'Rec Num
      If optOutputFormat(fmtCMGFile).Value = True Then
        grdColumns.SelBookmarks.RemoveAll
        grdColumns.SelBookmarks.Add pvarbookmark
        grdColumns.Bookmark = pvarbookmark
        SSTab1.Tab = 2
        COAMsgBox "Record numbers are not available when exporting to CMG format", vbExclamation, "Export"
        ValidateDefinition = False
        Exit Function
'      ElseIf optOutputFormat(fmtSQLTable).Value = True Then
'        grdColumns.SelBookmarks.RemoveAll
'        grdColumns.SelBookmarks.Add pvarbookmark
'        grdColumns.Bookmark = pvarbookmark
'        SSTab1.Tab = 2
'        COAMsgBox "Record numbers are not available when exporting to a SQL table", vbExclamation, "Export"
'        ValidateDefinition = False
'        Exit Function
      End If
    End If
  
  Next pintLoop



  ' Check that all sort order cols exist in the columns grid.
  ' THIS SHOULD NEVER HAPPEN, BUT JUST IN CASE...
  grdExportOrder.MoveFirst
  For pintLoop = 0 To grdExportOrder.Rows - 1
    pvarbookmark = grdExportOrder.GetBookmark(pintLoop)
    If Not IsInExportDefinition(CLng(grdExportOrder.Columns("ColExprID").CellText(pvarbookmark))) Then
      COAMsgBox "The '" & Right(grdExportOrder.Columns("Column").CellText(pvarbookmark), Len(grdExportOrder.Columns("Column").CellText(pvarbookmark)) - InStr(grdExportOrder.Columns("Column").CellText(pvarbookmark), ".")) & "' column has not been selected to appear in the export." & Chr(10) & "Please either include this column or remove it form the sort order.", vbExclamation, "Export"
      ValidateDefinition = False
      Exit Function
    End If
  Next pintLoop

If mlngExportID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlExport, _
      CStr(mlngExportID), _
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
               vbExclamation + vbOKOnly, "Export"
      Else
        COAMsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Export"
      End If

      Screen.MousePointer = vbDefault
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If COAMsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Export") = vbNo Then
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

  Exit Function
  
ValidateDefinition_ERROR:
  
  COAMsgBox "Error whilst validating export definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Export"
  ValidateDefinition = False
  
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




Private Function IsInExportDefinition(plngColExprID As Long) As Boolean

  'Purpose : Check if the specified column is used in the export.
  'Input   : ColExprID
  'Output  : True/False
  
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim pintPrevRow As Integer

  With grdColumns
  
    .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
          IsInExportDefinition = True
          Exit Function
        End If
        pintLoop = pintLoop + 1
      Loop
  End With

  IsInExportDefinition = False
  
End Function

Private Function CheckUniqueName(sName As String, lngCurrentID As Long) As Boolean

  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSysExportName " & _
         "WHERE UPPER(Name) = '" & UCase(Replace(sName, "'", "''")) & "' " & _
         "AND ID <> " & lngCurrentID
  
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckUniqueName = True
  Else
    CheckUniqueName = False
  End If
  
  Set rsTemp = Nothing

End Function


Private Function SaveDefinition() As Boolean

  Dim strSQL As String
  Dim lCount As Long
  Dim lExportID As Long
  Dim rsExport As New Recordset
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim strColumnsDone As String
  Dim lngColID As Long
  
  On Error GoTo Err_Trap
    
  Screen.MousePointer = vbHourglass

  If mlngExportID > 0 Then

    'We are updating an existing export definition
      
    'First save the basic definition
  
    strSQL = "UPDATE ASRSysExportName SET " & _
             "Name = '" & Trim(Replace(txtName.Text, "'", "''")) & "'," & _
             "Description = '" & Replace(txtDesc.Text, "'", "''") & "'," & _
             "BaseTable = " & cboBaseTable.ItemData(cboBaseTable.ListIndex) & "," & _
             "AllRecords = " & IIf(optBaseAllRecords.Value, 1, 0) & "," & _
             "Picklist = " & IIf(optBasePicklist.Value, txtBasePicklist.Tag, 0) & "," & _
             "Filter = " & IIf(optBaseFilter.Value, txtBaseFilter.Tag, 0) & "," & _
             "Parent1Table = " & txtParent1.Tag & "," & _
             "Parent1Filter= " & txtParent1Filter.Tag & "," & _
             "Parent2Table = " & txtParent2.Tag & "," & _
             "Parent2Filter= " & txtParent2Filter.Tag & "," & _
             "ChildTable = " & cboChild.ItemData(cboChild.ListIndex) & "," & _
             "ChildFilter = " & txtChildFilter.Tag & "," & _
             "ChildMaxRecords = " & Me.spnMaxRecords.Value & ","
  
    If optOutputFormat(fmtCSV).Value Then
      strSQL = strSQL & "Delimiter = '" & cboDelimiter.Text & "',"
      strSQL = strSQL & "OtherDelimiter = '" & Replace(Me.txtDelimiter.Text, "'", "''") & "',"
    Else
      strSQL = strSQL & "Delimiter = NULL,"
      strSQL = strSQL & "OtherDelimiter = NULL,"
    End If
    
    strSQL = strSQL & "OutputFormat = " & CStr(mobjOutputDef.GetSelectedFormatIndex) & ","

    If chkDestination(desSave).Value = vbChecked Then
      strSQL = strSQL & _
        "OutputSave = 1, " & _
        "OutputSaveExisting = " & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
    Else
      strSQL = strSQL & _
        "OutputSave = 0, " & _
        "OutputSaveExisting = 0, "
    End If
        
    If chkDestination(desEmail).Value = vbChecked Then
      strSQL = strSQL & _
          "OutputEmail = 1, " & _
          "OutputEmailAddr = " & txtEmailGroup.Tag & ", " & _
          "OutputEmailSubject = '" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "OutputEmailAttachAs = '" & Replace(txtEmailAttachAs.Text, "'", "''") & "', "
    Else
      strSQL = strSQL & _
          "OutputEmail = 0, " & _
          "OutputEmailAddr = 0, " & _
          "OutputEmailSubject = '', " & _
          "OutputEmailAttachAs = '', "
    End If
    
'    If optOutputFormat(fmtSQLTable).Value Then
'      strSQL = strSQL & _
'          "OutputFilename = '" & Replace(txtSQLTableName.Text, "'", "''") & "',"
'    Else
      strSQL = strSQL & _
          "OutputFilename = '" & Replace(txtFilename.Text, "'", "''") & "',"
'    End If
    
    
    ' Save the CMG export options
    If optOutputFormat(fmtCMGFile).Value = True Then
      strSQL = strSQL & "CMGExportFileCode = '" & Replace(txtFileExportCode.Text, "'", "''") & "'" & _
        ",CMGExportUpdateAudit = " & IIf(Me.chkUpdateAuditPointer.Value, 1, 0) & _
        ",CMGExportRecordID = " & cboParentFields.ItemData(cboParentFields.ListIndex) & ","
    End If
    
    ' Save the XML export options
    If optOutputFormat(fmtXML).Value = True Then
      strSQL = strSQL & "TransformFile = '" & Replace(txtTransformFile.Text, "'", "''") & "'," _
              & "XMLDataNodeName = '" & Replace(txtXMLDataNodeName.Text, "'", "''") & "'," _
              & "XSDFilename = '" & Replace(txtXSDFilename.Text, "'", "''") & "'," _
              & "PreserveTransformPath = " & IIf(chkPreserveTransformPath.Value = vbChecked, "1", "0") & ", " _
              & "PreserveXSDPath = " & IIf(chkPreserveXSDPath.Value = vbChecked, "1", "0") & ", " _
              & "SplitXMLNodesFile = " & IIf(chkSplitXMLNodesFile.Value = vbChecked, "1", "0") & ", "
    End If
       
    strSQL = strSQL & "Quotes = " & IIf(Me.chkQuotes.Value, 1, 0) & "," & _
                  "StripDelimiterFromData = " & IIf(Me.chkStripDelimiter.Value, 1, 0) & "," & _
                  "SplitFile = " & IIf(Me.chkSplitFile.Value, 1, 0) & "," & _
                  "SplitFileSize = '" & Replace(spnSplitFileSize.Value, "'", "''") & "'," & _
                  "Header = " & cboHeaderOptions.ListIndex & "," & _
                  "HeaderText = '" & Replace(txtCustomHeader.Text, "'", "''") & "'," & _
                  "Footer = " & cboFooterOptions.ListIndex & "," & _
                  "FooterText = '" & Replace(txtCustomFooter.Text, "'", "''") & "'," & _
                  "DateFormat = '" & cboDateFormat.Text & "'," & _
                  "DateSeparator = '" & cboDateSeparator.Text & "'," & _
                  "DateYearDigits = '" & cboDateYearDigits.Text & "'," & _
                  "AuditChangesOnly = " & IIf(chkAuditChangesOnly.Value = vbChecked, "1", "0") & ", " & _
                  "OmitHeader = " & IIf(chkOmitHeader.Value, 1, 0) & "," & _
                  "ForceHeader = " & IIf(chkForceHeader.Value, 1, 0) & ","

    strSQL = strSQL & _
                  "Parent1AllRecords = " & IIf(optParent1AllRecords.Value, 1, 0) & "," & _
                  "Parent1Picklist = " & IIf(optParent1Picklist.Value, txtParent1Picklist.Tag, 0) & "," & _
                  "Parent2AllRecords = " & IIf(optParent2AllRecords.Value, 1, 0) & "," & _
                  "Parent2Picklist = " & IIf(optParent2Picklist.Value, txtParent2Picklist.Tag, 0) & " " & _
                  "WHERE ID = " & mlngExportID


    If IsRecordSelectionValid = False Then
      SaveDefinition = False
      Exit Function
    End If
              
     mdatData.ExecuteSql (strSQL)
  
    Call UtilUpdateLastSaved(utlExport, mlngExportID)
  
  Else

    ' Adding a new export definition

    strSQL = "INSERT ASRSysExportName (" & _
           "Name, Description, BaseTable, " & _
           "AllRecords, Picklist, Filter, " & _
           "Parent1Table, Parent1Filter, " & _
           "Parent2Table, Parent2Filter, " & _
           "ChildTable, ChildFilter, ChildMaxRecords, " & _
           "Delimiter, OtherDelimiter, Quotes, StripDelimiterFromData," & _
           "SplitFile, SplitFileSize," & _
           "Header, HeaderText, Footer, FooterText, " & _
           "DateFormat, DateSeparator, DateYearDigits, UserName," & _
           "CMGExportFileCode, CMGExportUpdateAudit, CMGExportRecordID," & _
           "Parent1AllRecords, Parent1Picklist, Parent2AllRecords, Parent2Picklist, " & _
           "AuditChangesOnly, OmitHeader, ForceHeader, OutputFormat, OutputSave, " & _
           "OutputSaveExisting, OutputEmail, OutputEmailAddr, OutputEmailSubject, OutputEmailAttachAs, OutputFilename," & _
           "TransformFile, XMLDataNodeName, XSDFilename, PreserveTransformPath, PreserveXSDPath, SplitXMLNodesFile) "
                     
    strSQL = strSQL & _
           "Values('" & _
           Trim(Replace(txtName.Text, "'", "''")) & "','" & _
           Replace(txtDesc.Text, "'", "''") & "'," & _
           cboBaseTable.ItemData(cboBaseTable.ListIndex)

    If optBaseAllRecords Then
      strSQL = strSQL & ", 1, 0, 0"
    ElseIf optBasePicklist Then
      strSQL = strSQL & ", 0, " & Val(txtBasePicklist.Tag) & ", 0"
    Else
      strSQL = strSQL & ", 0, 0, " & Val(txtBaseFilter.Tag)
    End If

    strSQL = strSQL & ", " & CStr(txtParent1.Tag)
    strSQL = strSQL & ", " & CStr(txtParent1Filter.Tag)
    strSQL = strSQL & ", " & CStr(txtParent2.Tag)
    strSQL = strSQL & ", " & CStr(txtParent2Filter.Tag)
    strSQL = strSQL & ", " & CStr(cboChild.ItemData(cboChild.ListIndex))
    strSQL = strSQL & ", " & CStr(txtChildFilter.Tag)
    strSQL = strSQL & ", " & CStr(spnMaxRecords.Value)

    If optOutputFormat(fmtCSV).Value Then
      strSQL = strSQL & ",'" & cboDelimiter.Text & "'"
      strSQL = strSQL & ",'" & txtDelimiter.Text & "'" '"OtherDelimiter = '" & Replace(Me.txtDelimiter.Text, "'", "''") & "',"
    Else
      strSQL = strSQL & ",NULL,NULL"
    End If


    If Me.chkQuotes.Value Then strSQL = strSQL & ", 1" Else strSQL = strSQL & ", 0"
    If Me.chkStripDelimiter.Value Then strSQL = strSQL & ", 1" Else strSQL = strSQL & ", 0"
    If Me.chkSplitFile.Value Then strSQL = strSQL & ", 1" Else strSQL = strSQL & ", 0"
                 
    strSQL = strSQL & ", '" & spnSplitFileSize.Value & "'" & _
                  ", " & CStr(cboHeaderOptions.ListIndex) & _
                  ", '" & Replace(txtCustomHeader.Text, "'", "''") & "'" & _
                  ", " & CStr(cboFooterOptions.ListIndex) & _
                  ", '" & Replace(txtCustomFooter.Text, "'", "''") & "'"

    strSQL = strSQL & ", '" & cboDateFormat.Text & "'"
    strSQL = strSQL & ", '" & cboDateSeparator.Text & "'"
    strSQL = strSQL & ", '" & cboDateYearDigits.Text & "'"
            
    strSQL = strSQL & ", '" & datGeneral.UserNameForSQL & "'"
    
    ' Save the CMG export options
    strSQL = strSQL & ", '" & Replace(txtFileExportCode.Text, "'", "''") & "'"
    strSQL = strSQL & ", " & IIf(Me.chkUpdateAuditPointer.Value, 1, 0)
    cboParentFields.ListIndex = IIf(cboParentFields.ListIndex = -1, 0, cboParentFields.ListIndex)
    strSQL = strSQL & ", " & cboParentFields.ItemData(cboParentFields.ListIndex) & ","
   
    strSQL = strSQL & IIf(optParent1AllRecords, "1", "0") & ","
    strSQL = strSQL & IIf(optParent1Picklist, CStr(txtParent1Picklist.Tag), "0") & ","
    strSQL = strSQL & IIf(optParent2AllRecords, "1", "0") & ","
    strSQL = strSQL & IIf(optParent2Picklist, CStr(txtParent2Picklist.Tag), "0") & ","
  
    ' Header and append options
    'strSQL = strSQL & IIf(chkAppendToFile.Value = vbChecked, "1", "0") & ","
    strSQL = strSQL & IIf(chkAuditChangesOnly.Value = vbChecked, "1", "0") & ", "
    strSQL = strSQL & IIf(chkOmitHeader.Value = vbChecked, "1", "0") & ","
    strSQL = strSQL & IIf(chkForceHeader.Value = vbChecked, "1", "0") & ","
    
    strSQL = strSQL & CStr(mobjOutputDef.GetSelectedFormatIndex) & ","
    
    If chkDestination(desSave).Value = vbChecked Then
      strSQL = strSQL & "1, " & _
        cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
    Else
      strSQL = strSQL & "0, 0, "
    End If

    If chkDestination(desEmail).Value = vbChecked Then
      strSQL = strSQL & "1, " & _
          txtEmailGroup.Tag & ", " & _
          "'" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "'" & Replace(txtEmailAttachAs.Text, "'", "''") & "', "
    Else
      strSQL = strSQL & "0, 0, '', '', "
    End If

    strSQL = strSQL & _
        "'" & Replace(txtFilename.Text, "'", "''") & "','" & Replace(txtTransformFile.Text, "'", "''") & "'," & _
        "'" & Replace(txtXMLDataNodeName.Text, "'", "''") & "'," & _
        "'" & Replace(txtXSDFilename.Text, "'", "''") & "'," & _
        IIf(chkPreserveTransformPath.Value = vbChecked, "1", "0") & ", " & _
        IIf(chkPreserveXSDPath.Value = vbChecked, "1", "0") & ", " & _
        IIf(chkSplitXMLNodesFile.Value = vbChecked, "1", "0") & ")"
          
    If IsRecordSelectionValid = False Then
      SaveDefinition = False
      Exit Function
    End If
    
    mlngExportID = InsertExport(strSQL)
  
    If mlngExportID = 0 Then
      SaveDefinition = False
      Exit Function
    End If
    
    Call UtilCreated(utlExport, mlngExportID)
  
  End If

  SaveAccess
  SaveObjectCategories cboCategory, utlExport, mlngExportID

  ' Now save the column details
  
  ' First, remove any records from the detail table with the specified ExportID
  ClearDetailTables mlngExportID
  
  ' Loop through the details grid, and also the sortorder grid
  With grdColumns
    
    '.MoveFirst
  
    strColumnsDone = "\"
    Do Until pintLoop = .Rows

      pvarbookmark = .GetBookmark(pintLoop)
      lngColID = .Columns("ColExprID").CellValue(pvarbookmark)
      
      strSQL = "INSERT ASRSysExportDetails (" & _
             "ExportID, " & _
             "Type , " & _
             "TableID, " & _
             "ColExprID, " & _
             "Data, " & _
             "FillerLength, " & _
             "Heading, " & _
             "Decimals, " & _
             "CMGColumnCode, " & _
             "ConvertCase, " & _
             "SuppressNulls, " & _
             "SortOrderSequence, " & _
             "SortOrder) "
  
      strSQL = strSQL & "VALUES(" & mlngExportID & ", "
      
      strSQL = strSQL & "'" & .Columns("Type").CellText(pvarbookmark) & "', "
      strSQL = strSQL & .Columns("TableID").CellText(pvarbookmark) & ", "
      strSQL = strSQL & CStr(lngColID) & ", "
      strSQL = strSQL & "'" & Replace(.Columns("Data").CellText(pvarbookmark), "'", "''") & "', "
      strSQL = strSQL & .Columns("Length").CellText(pvarbookmark) & ", "
      strSQL = strSQL & "'" & Replace(.Columns("Heading").CellText(pvarbookmark), "'", "''") & "', "
      strSQL = strSQL & IIf(IsNull(.Columns("Decimals").CellText(pvarbookmark)) Or (.Columns("Decimals").CellText(pvarbookmark) = vbNullString), 0, .Columns("Decimals").CellText(pvarbookmark)) & ", "
      strSQL = strSQL & "'" & Replace(.Columns("CMG Code").CellText(pvarbookmark), "'", "''") & "', "
      'NPG20071213 Fault 12867
      strSQL = strSQL & "'" & Replace(.Columns("ConvertCase").CellText(pvarbookmark), "'", "''") & "', "
      'NPG20080617 Suggestion S000816
      'strSQL = strSQL & IIf(IsNull(.Columns("SuppressNulls").CellText(pvarbookmark)) Or (.Columns("SuppressNulls").CellText(pvarbookmark) = vbNullString), 0, 1) & ", "
      
      If IsNull(.Columns("SuppressNulls").CellText(pvarbookmark)) Or (.Columns("SuppressNulls").CellText(pvarbookmark) = vbNullString) Then
        strSQL = strSQL & "0, "
      Else
        strSQL = strSQL & IIf(CBool(.Columns("SuppressNulls").CellText(pvarbookmark)), 1, 0) & ", "
      End If
                       
      If .Columns("Type").CellText(pvarbookmark) <> "X" And InStr(strColumnsDone, "\" & CStr(lngColID) & "\") = 0 Then
        strSQL = strSQL & GetSortOrderString(lngColID)
        strColumnsDone = strColumnsDone & .Columns("ColExprID").CellText(pvarbookmark) & "\"
      Else
        strSQL = strSQL & "0, NULL)"
      End If
      
      pintLoop = pintLoop + 1
      
      mdatData.ExecuteSql (strSQL)
  
    Loop
  
  End With
  
  Screen.MousePointer = vbDefault
  SaveDefinition = True
  Changed = False

  Exit Function

Err_Trap:
  Screen.MousePointer = vbDefault
  COAMsgBox "Error whilst saving Export definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Export"
  SaveDefinition = False
Resume
End Function

Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysExportAccess WHERE ID = " & mlngExportID
  mdatData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysExportAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngExportID & ", sysusers.name," & _
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
    " AND ISNULL(sysusers.uid, 0) <> 0)"
  mdatData.ExecuteSql (sSQL)

  ' Update the new access records with the real access values.
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      sSQL = "IF EXISTS (SELECT * FROM ASRSysExportAccess" & _
        " WHERE ID = " & CStr(mlngExportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysExportAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngExportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      mdatData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub




Private Function GetSortOrderString(plngColExprID As Long) As String

  'Purpose : Constructs a string to insert into the SQL statement when saving
  '          an export definition.
  'Input   : ColExprID
  'Output  : String for insertion into the SQL statement
  
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim pintPrevRow As Integer

  With grdExportOrder
  
    pintPrevRow = .AddItemRowIndex(.Bookmark)
    
    .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
        
          GetSortOrderString = (.AddItemRowIndex(pvarbookmark) + 1) & ", '"
          GetSortOrderString = GetSortOrderString & IIf(.Columns("Sort Order").CellText(pvarbookmark) = "Ascending", "Asc", "Desc") & "')"
          Exit Function
          
        End If
        pintLoop = pintLoop + 1
      Loop
  End With

  GetSortOrderString = "0, NULL)"

End Function


Private Function InsertExport(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertExport_ERROR
'
'  Dim rsExport As Recordset
'
'  mdatData.ExecuteSql pstrSQL
'
'  pstrSQL = "Select Max(ID) From ASRSysExportName"
'  Set rsExport = mdatData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)
'  InsertExport = rsExport(0)
'
'  rsExport.Close
'
'  Set rsExport = Nothing

'###################

  Dim sSQL As String
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
    pmADO.Value = "AsrSysExportName"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "ID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, app.ProductName
        InsertExport = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertExport = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertExport_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function

Private Sub ClearDetailTables(plngExportID As Long)

  ' Delete all column information from the Details table.
  
  Dim pstrSQL As String
  
  pstrSQL = "Delete From ASRSysExportDetails Where ExportID = " & plngExportID
  mdatData.ExecuteSql pstrSQL

End Sub


Private Function RetrieveExportDetails(plngExportID As Long) As Boolean

  Dim rsTemp As Recordset
  Dim pintLoop As Integer
  Dim pstrText As String
  Dim bIsAudited As Boolean
  Dim objExpr As New clsExprExpression
  Dim fAlreadyNotified As Boolean
  Dim sMessage As String
  
  Dim strHeading As String  'MH20030120
  Dim lngCalcCount As Long
  Dim lngFillerCount As Long
  Dim lngTextCount As Long
  Dim lngRecNumCount As Long
  Dim pintConvertCase As Integer  'NPG20071213 Fault 12867
  Dim bSuppNulls As Boolean   'NPG20080617 Suggestion S00816
  
  Dim lngLength As Long

  On Error GoTo Load_ERROR
  
  'Load the basic guff first
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysExportName.*, " & _
                                     "CONVERT(integer, ASRSysExportName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysExportName WHERE ID = " & plngExportID)

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This Export definition has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
    Set rsTemp = Nothing
    RetrieveExportDetails = False
    Exit Function
  End If
  
  
  If rsTemp!OutputFormat = fmtSQLTable Then
    'If Not mblnEnableSQLTable Then
      COAMsgBox "This Export definition is invalid as export to SQL Table is no longer supported.", vbExclamation + vbOKOnly, "Export"
      Set rsTemp = Nothing
      RetrieveExportDetails = False
      mblnDeleted = True
      Exit Function
    'End If
  End If


  ' Set name, username, access etc
  If FromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!userName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!userName) = LCase$(gsUserName))
  End If
   
  mblnReadOnly = Not datGeneral.SystemPermission("EXPORT", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlExport, plngExportID) = ACCESS_READONLY)
  End If
  
  ' Set Definition Description
  txtDesc.Text = IIf(IsNull(rsTemp!Description), "", rsTemp!Description)
    
  Changed = False
  
  mblnLoading = True
  
  LoadBaseCombo
  SetComboText cboBaseTable, datGeneral.GetTableName(rsTemp!BaseTable)
  mstrBaseTable = cboBaseTable.Text
  
  ' Set the categories combo
  GetObjectCategories cboCategory, utlExport, mlngExportID
  
  UpdateDependantFields
  
  ' Set Base Table Record Select Options
  If rsTemp!AllRecords Then optBaseAllRecords.Value = True
  If rsTemp!picklist Then
    optBasePicklist.Value = True
    txtBasePicklist.Tag = rsTemp!picklist
    txtBasePicklist.Text = datGeneral.GetPicklistName(rsTemp!picklist)
  End If
  
  If rsTemp!Filter Then
    optBaseFilter.Value = True
    txtBaseFilter.Tag = rsTemp!Filter
    txtBaseFilter.Text = datGeneral.GetFilterName(rsTemp!Filter)
  End If
  
  ' Set Parent 1 Table Record Select Options
  If (rsTemp!parent1AllRecords) Or (rsTemp!parent1table <= 0) Then optParent1AllRecords.Value = True
  
  If rsTemp!parent1picklist > 0 Then
    optParent1Picklist.Value = True
    txtParent1Picklist.Tag = rsTemp!parent1picklist
    txtParent1Picklist.Text = datGeneral.GetPicklistName(rsTemp!parent1picklist)
  End If
    
  If rsTemp!parent1filter > 0 Then
    optParent1Filter.Value = True
    txtParent1Filter.Tag = rsTemp!parent1filter
    txtParent1Filter.Text = datGeneral.GetFilterName(txtParent1Filter.Tag)
  End If
  
  ' Set Parent 2 Table Record Select Options
  If (rsTemp!parent2AllRecords) Or (rsTemp!parent2table <= 0) Then optParent2AllRecords.Value = True
  
  If rsTemp!parent2picklist > 0 Then
    optParent2Picklist.Value = True
    txtParent2Picklist.Tag = rsTemp!parent2picklist
    txtParent2Picklist.Text = datGeneral.GetPicklistName(rsTemp!parent2picklist)
  End If
      
  If rsTemp!parent2filter > 0 Then
    optParent2Filter.Value = True
    txtParent2Filter.Tag = rsTemp!parent2filter
    txtParent2Filter.Text = datGeneral.GetFilterName(txtParent2Filter.Tag)
  End If
  
  ' Set Child Table
  If rsTemp!ChildTable = 0 Then
    If cboChild.ListCount = 0 Then
      cboChild.AddItem "<None>"
      cboChild.ItemData(cboChild.NewIndex) = 0
    End If
    cboChild.ListIndex = 0
  Else
    SetComboText cboChild, datGeneral.GetTableName(rsTemp!ChildTable)
  End If
  
  ' Set Child Table Filter
  If rsTemp!childFilter Then
    txtChildFilter.Tag = rsTemp!childFilter
    txtChildFilter.Text = datGeneral.GetFilterName(txtChildFilter.Tag)
  End If
  
  ' Set Child Max Records
  spnMaxRecords.Value = rsTemp!ChildMaxRecords
      
  ' Header / append options
  'chkAppendToFile.Value = IIf(rsTemp!AppendToFile = True, vbChecked, vbUnchecked)
  chkForceHeader.Value = IIf(rsTemp!ForceHeader = True, vbChecked, vbUnchecked)
  chkOmitHeader.Value = IIf(rsTemp!OmitHeader = True, vbChecked, vbUnchecked)
  
  If Not IsNull(rsTemp!SplitFile) Then
    If rsTemp!SplitFile Then
      chkSplitFile.Value = vbChecked
      spnSplitFileSize.Value = IIf(IsNull(rsTemp!SplitFileSize), 0, rsTemp!SplitFileSize)
    End If
  End If
  
  'chkOmitHeader.Enabled = (rsTemp!OutputSaveExisting = 4) 'chkAppendToFile.Value
  
'  ' Set output info
'  Select Case rsTemp!outputtype
'
'    Case "D"
'      optOutputFormat(fmtCSV).Value = True
'      txtFilename.Text = rsTemp!OutputFilename
'      SetComboText cboDelimiter, rsTemp!delimiter
'      grdColumns.Columns(4).ForeColor = vbWindowText
'      grdColumns.Columns(4).HeadForeColor = vbWindowText
''      grdColumns.Columns(4).ForeColor = vbBlack
''      grdColumns.Columns(4).HeadForeColor = vbBlack
'      TextOptionsStatus "D"
'
'    Case "F"
'      optOutputFormat(fmtFixedLengthFile).Value = True
'      txtFilename.Text = rsTemp!OutputFilename
'      grdColumns.Columns(4).ForeColor = vbWindowText
'      grdColumns.Columns(4).HeadForeColor = vbWindowText
'      TextOptionsStatus "F"
'
'    Case "C"
'      optOutputFormat(fmtCMGFile).Value = True
'      txtFilename.Text = rsTemp!OutputFilename
'      grdColumns.Columns(4).ForeColor = vbWindowText
'      grdColumns.Columns(4).HeadForeColor = vbWindowText
'      TextOptionsStatus "C"
'
'    Case "S"
'      optOutputFormat(fmtSQLTable).Value = True
'      txtSQLTableName.Text = rsTemp!OutputFilename
'
'      'TM20010823 Fault 2389
'      'Could not see the text in the length column.
'      grdColumns.Columns(4).ForeColor = vbWindowText
'      grdColumns.Columns(4).HeadForeColor = vbWindowText
''      grdColumns.Columns(4).ForeColor = vbButtonFace
''      grdColumns.Columns(4).HeadForeColor = vbButtonShadow
'      TextOptionsStatus "S"
'
'    Case "X"
'      optOutputFormat(fmtExcelWorksheet).Value = True
'      txtSQLTableName.Text = rsTemp!OutputFilename
'      TextOptionsStatus "X"
'
'  End Select

  mblnLoading = True

  grdColumns.Columns(4).ForeColor = vbWindowText
  grdColumns.Columns(4).HeadForeColor = vbWindowText

  optOutputFormat(rsTemp!OutputFormat).Value = True
  optOutputFormat_Click rsTemp!OutputFormat
  'TextOptionsStatus rsTemp!OutputFormat

  chkDestination(desSave).Value = IIf(rsTemp!OutputSave, vbChecked, vbUnchecked)
  mobjOutputDef.DestinationClick desSave
  SetComboItem cboSaveExisting, rsTemp!OutputSaveExisting

  chkDestination(desEmail).Value = IIf(rsTemp!OutputEmail, vbChecked, vbUnchecked)
  If rsTemp!OutputEmail Then
    txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
    txtEmailGroup.Tag = rsTemp!OutputEmailAddr
    txtEmailSubject.Text = rsTemp!OutputEmailSubject
    txtEmailAttachAs.Text = IIf(IsNull(rsTemp!OutputEmailAttachAs), vbNullString, rsTemp!OutputEmailAttachAs)
  End If

'  If rsTemp!OutputFormat = fmtSQLTable Then
'    txtSQLTableName.Text = rsTemp!OutputFilename
'  Else
    txtFilename.Text = rsTemp!OutputFilename
'  End If


  'Text file options
  If rsTemp!Quotes Then chkQuotes.Value = vbChecked
  If Not IsNull(rsTemp!StripDelimiterFromData) Then
    If rsTemp!StripDelimiterFromData Then chkStripDelimiter.Value = vbChecked
  End If
  
  If IsNull(rsTemp!delimiter) Then
    cboDelimiter.ListIndex = -1
  Else
    'AE20071004 Fault 12489
'    SetComboText cboDelimiter, rsTemp!delimiter
    SetComboText cboDelimiter, rsTemp!delimiter, True
    txtDelimiter.Text = IIf(IsNull(rsTemp!otherdelimiter), vbNullString, rsTemp!otherdelimiter)
  End If

  'If rsTemp!header Then chkHeader.Value = True
  cboHeaderOptions.ListIndex = rsTemp!Header
  txtCustomHeader.Text = IIf(IsNull(rsTemp!HeaderText), vbNullString, rsTemp!HeaderText)

  cboFooterOptions.ListIndex = rsTemp!Footer
  txtCustomFooter.Text = IIf(IsNull(rsTemp!FooterText), vbNullString, rsTemp!FooterText)

  'mblnLoading = False
  CheckIfOmitHeaderEnabled

  'CMG Options
  If rsTemp!CMGExportUpdateAudit = True Then chkUpdateAuditPointer.Value = 1
  
  If rsTemp!OutputFormat = fmtCMGFile Then
    If Not IsNull(rsTemp!CMGExportFileCode) Then txtFileExportCode.Text = rsTemp!CMGExportFileCode
    If Not IsNull(rsTemp!CMGExportRecordID) And rsTemp!CMGExportRecordID > 0 Then SetComboText cboParentFields, datGeneral.GetColumnName(rsTemp!CMGExportRecordID)
  End If

  ' XML specifics
  If rsTemp!OutputFormat = fmtXML Then
    If Not IsNull(rsTemp!TransformFile) Then txtTransformFile.Text = rsTemp!TransformFile
    If Not IsNull(rsTemp!XMLDataNodeName) Then txtXMLDataNodeName.Text = rsTemp!XMLDataNodeName
    
    If Not IsNull(rsTemp!XSDFilename) Then txtXSDFilename.Text = rsTemp!XSDFilename
    chkPreserveTransformPath.Value = IIf(rsTemp!PreserveTransformPath = True, vbChecked, vbUnchecked)
    chkPreserveXSDPath.Value = IIf(rsTemp!PreserveXSDPath = True, vbChecked, vbUnchecked)
    chkSplitXMLNodesFile.Value = IIf(rsTemp!SplitXMLNodesFile = True, vbChecked, vbUnchecked)
  End If

  chkAuditChangesOnly.Value = IIf(rsTemp!AuditChangesOnly = True, vbChecked, vbUnchecked)

  ' Set Date Format
  SetComboText cboDateFormat, rsTemp!DateFormat
  SetComboText cboDateSeparator, rsTemp!dateseparator
  SetComboText cboDateYearDigits, rsTemp!Dateyeardigits
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    grdColumns.Enabled = True
    grdExportOrder.Enabled = True
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  mlngTimeStamp = rsTemp!intTimestamp
  
  ' =========================
  
  ' Now load the details
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysExportDetails WHERE ExportID = " & plngExportID & " ORDER BY ID")
  
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "No column information found for this Export.", vbExclamation + vbOKOnly, "Export"
    Set rsTemp = Nothing
    RetrieveExportDetails = False
    Exit Function
  End If
  
  Do Until rsTemp.EOF
    
    bIsAudited = datGeneral.IsColumnAudited(rsTemp!ColExprID)
    lngLength = IIf(rsTemp!fillerlength > 999999, 999999, rsTemp!fillerlength)
    
    If rsTemp!Type = "X" Then
      pstrText = rsTemp!Type & vbTab & rsTemp!TableID & vbTab & rsTemp!ColExprID & vbTab _
          & datGeneral.GetExpression(rsTemp!ColExprID) & vbTab & CStr(lngLength) & vbTab & rsTemp!CMGColumnCode & vbTab _
          & IIf(bIsAudited = True, "1", "0") & vbTab
         
      If rsTemp!ColExprID > 0 Then
        objExpr.ExpressionID = rsTemp!ColExprID
        objExpr.ConstructExpression
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        objExpr.ValidateExpression True, True
        
        If (objExpr.ReturnType = giEXPRVALUE_NUMERIC Or _
                            objExpr.ReturnType = giEXPRVALUE_BYREF_NUMERIC) Then
          pstrText = pstrText & IIf(IsNull(rsTemp!Decimals) Or (rsTemp!Decimals = vbNullString), 0, rsTemp!Decimals)
        Else
          pstrText = pstrText & vbNullString
        End If
        
        Set objExpr = Nothing
      End If
         
      'MH20030120
      lngCalcCount = lngCalcCount + 1
      strHeading = IIf(IsNull(rsTemp!Heading), "Calculation" & CStr(lngCalcCount), rsTemp!Heading)
      
      'NPG20080717 Fault 13275
      ' pstrText = pstrText & vbTab & strHeading
      pstrText = pstrText & vbTab & strHeading & vbTab
      
      'NPG20080717 Fault 13275
      pintConvertCase = rsTemp!ConvertCase
      pstrText = pstrText & pintConvertCase & vbTab
      
      'NPG20080717 Fault 13274
      pstrText = pstrText & rsTemp!SuppressNulls & vbTab
      
      grdColumns.AddItem pstrText
    ElseIf rsTemp!Type = "C" Then
      ' RH 19/10/00 - BUG - Get table.column name from db not from the definition...
      '               it might have been renamed!
      'pstrText = rsTemp!Type & vbTab & rsTemp!TableID & vbTab & rsTemp!ColExprID & vbTab & rsTemp!Data & vbTab & rsTemp!fillerlength
      strHeading = datGeneral.GetTableName(rsTemp!TableID) & "." & datGeneral.GetColumnName(rsTemp!ColExprID)
      pstrText = rsTemp!Type & vbTab & rsTemp!TableID & vbTab & rsTemp!ColExprID & vbTab _
        & strHeading _
        & vbTab & CStr(lngLength) & vbTab & rsTemp!CMGColumnCode & vbTab & IIf(bIsAudited = True, "1", "0") & vbTab
        
      If datGeneral.GetDataType(rsTemp!TableID, rsTemp!ColExprID) = sqlNumeric Then
        pstrText = pstrText & IIf(IsNull(rsTemp!Decimals) Or (rsTemp!Decimals = vbNullString), 0, rsTemp!Decimals)
      Else
        pstrText = pstrText & vbNullString
      End If
      
      'MH20030120
      pstrText = pstrText & vbTab & _
          IIf(IsNull(rsTemp!Heading), strHeading, rsTemp!Heading) & vbTab
    
      'NPG20071213 Fault 12867
      pintConvertCase = rsTemp!ConvertCase
      pstrText = pstrText & pintConvertCase & vbTab
      
      'NPG20080617 Suggestion S000816
      pstrText = pstrText & rsTemp!SuppressNulls & vbTab
      
      grdColumns.AddItem pstrText
    
    Else

      pstrText = rsTemp!Type & vbTab & rsTemp!TableID & vbTab & rsTemp!ColExprID & vbTab & rsTemp!Data _
        & vbTab & CStr(lngLength) & vbTab & rsTemp!CMGColumnCode & vbTab & IIf(bIsAudited = True, "1", "0") & vbTab & _
        vbNullString

      'Default to old heading
      '(still need to increase counter even if heading is overwritten!)
      Select Case rsTemp!Type
      Case "F", "R"   'Filler or Carriage Return
        lngFillerCount = lngFillerCount + 1
        strHeading = "Filler" & CStr(lngFillerCount)
      Case "T"        'Text
        lngTextCount = lngTextCount + 1
        strHeading = "Text" & CStr(lngTextCount)
      Case "N"        'Record Number
        lngRecNumCount = lngRecNumCount + 1
        strHeading = "Record Number" & CStr(lngRecNumCount)
      End Select

      'Overwrite if there is a new heading
      If Not IsNull(rsTemp!Heading) Then
        strHeading = rsTemp!Heading
      End If
      pstrText = pstrText & vbTab & strHeading
     
      grdColumns.AddItem pstrText

    End If
    
    rsTemp.MoveNext
  
  Loop

  grdColumns.RowHeight = lng_GRIDROWHEIGHT
  
  ' Now load the sort order details
  
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysExportDetails WHERE ExportID = " & plngExportID & " AND SortOrderSequence > 0 AND Type = 'C' ORDER BY [SortOrderSequence]")
  
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "No sort order information found for this Export.", vbExclamation + vbOKOnly, "Export"
    Set rsTemp = Nothing
    RetrieveExportDetails = False
    Exit Function
  End If
  
  ' Add to the sort order grid
  Do Until rsTemp.EOF
    
    grdExportOrder.AddItem rsTemp!ColExprID & vbTab & _
                           rsTemp!Data & vbTab & _
                           IIf(rsTemp!SortOrder = "Asc", "Ascending", "Descending")
    
    rsTemp.MoveNext
  
  Loop
  
  mblnBaseTableSpecificChanged = True
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Cancelled = True
    RetrieveExportDetails = False
    Exit Function
  End If
  
  ' Tidyup
  Set rsTemp = Nothing
  RetrieveExportDetails = True
  Exit Function

Load_ERROR:

  COAMsgBox "Error whilst retrieving the Export definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Export"
  RetrieveExportDetails = False
  Set rsTemp = Nothing

End Function


Private Sub PopulateAccessGrid()
  ' Populate the access grid.
  Dim rsAccess As ADODB.Recordset
  
  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With
  
  ' Get the recordset of user groups and their access on this definition.
  Set rsAccess = GetUtilityAccessRecords(utlExport, mlngExportID, mblnFromCopy)
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
  Dim strColumnType As String
  Dim lngColumnID As Long
  Dim sCalcName As String
  Dim fOnlyFatalMessages As Boolean
  
  If IsMissing(pvOnlyFatalMessages) Then
    fOnlyFatalMessages = mblnLoading
  Else
    fOnlyFatalMessages = CBool(pvOnlyFatalMessages)
  End If
  
  ' Return false if some of the filters/picklists/calcs need to be removed from the definition,
  ' or if the definition needs to be made hidden.
  fChangesRequired = False
  fDefnAlreadyHidden = AllHiddenAccess
  fNeedToForceHidden = False

  ' Dimension arrays to hold details of the filters/picklists that
  ' have been deleted, made hidden or are now invalid.
  ' Column 1 - parameter description
  ReDim asDeletedParameters(0)
  ReDim asHiddenBySelfParameters(0)
  ReDim asHiddenByOtherParameters(0)
  ReDim asInvalidParameters(0)

  ' Check Base Table Picklist
  If (Len(txtBasePicklist.Tag) > 0) And (Val(txtBasePicklist.Tag) <> 0) Then
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
          sBigMessage = "The '" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"

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
  If Len(txtBaseFilter.Tag) > 0 And Val(txtBaseFilter.Tag) <> 0 Then
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
          sBigMessage = "The '" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"

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

  ' Check Parent 1 Picklist
  If (Len(txtParent1Picklist.Tag) > 0) And (Val(txtParent1Picklist.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtParent1Picklist.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        If fRemove Then
          sBigMessage = "The '" & txtParent1.Text & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & txtParent1.Text & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & txtParent1.Text & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & txtParent1.Text & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & txtParent1.Text & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtParent1Picklist.Tag = 0
      txtParent1Picklist.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Parent 1 Filter
  If Len(txtParent1Filter.Tag) > 0 And Val(txtParent1Filter.Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtParent1Filter.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)

        If fRemove Then
          sBigMessage = "The '" & txtParent1.Text & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & txtParent1.Text & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & txtParent1.Text & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & txtParent1.Text & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & txtParent1.Text & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtParent1Filter.Tag = 0
      txtParent1Filter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Check Parent 2 Picklist
  If (Len(txtParent2Picklist.Tag) > 0) And (Val(txtParent2Picklist.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtParent2Picklist.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        If fRemove Then
          sBigMessage = "The '" & txtParent2.Text & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & txtParent2.Text & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & txtParent2.Text & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & txtParent2.Text & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & txtParent2.Text & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtParent2Picklist.Tag = 0
      txtParent2Picklist.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Parent 2 Filter
  If Len(txtParent2Filter.Tag) > 0 And Val(txtParent2Filter.Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtParent2Filter.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)

        If fRemove Then
          sBigMessage = "The '" & txtParent2.Text & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & txtParent2.Text & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & txtParent2.Text & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & txtParent2.Text & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & txtParent2.Text & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtParent2Filter.Tag = 0
      txtParent2Filter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Child Filter
  If Len(txtChildFilter.Tag) > 0 And Val(txtChildFilter.Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtChildFilter.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)

        If fRemove Then
          sBigMessage = "The '" & cboChild.List(cboChild.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboChild.List(cboChild.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboChild.List(cboChild.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboChild.List(cboChild.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboChild.List(cboChild.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtChildFilter.Tag = 0
      txtChildFilter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Calcs
  With grdColumns
    If .Rows > 0 Then
      For iLoop = .Rows - 1 To 0 Step -1
        varBookmark = .AddItemBookmark(iLoop)
        strColumnType = .Columns("Type").CellValue(varBookmark)
        lngColumnID = .Columns("ColExprID").CellValue(varBookmark)
        
        If strColumnType = "X" Then
          fRemove = False
          iResult = ValidateCalculation(lngColumnID)
  
          sCalcName = .Columns("Data").CellValue(varBookmark)

          Select Case iResult
            Case REC_SEL_VALID_HIDDENBYUSER
              ' Calculation hidden by the current user.
              ' Only a problem if the current definition is NOT owned by the current user,
              ' or if the current definition is not already hidden.
              fRemove = (Not mblnDefinitionCreator) And _
                (Not mblnReadOnly) And _
                (Not FormPrint)

              If fRemove Then
                sBigMessage = "The '" & sCalcName & "' calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
                COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
              Else
                fNeedToForceHidden = True
  
                ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
                asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & sCalcName & "' calculation"
              End If

            Case REC_SEL_VALID_DELETED
              ' Calc deleted by another user.
              ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
              asDeletedParameters(UBound(asDeletedParameters)) = "'" & sCalcName & "' calculation"

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)

            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                ' Calc hidden by another user.
                ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
                asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & sCalcName & "' calculation"
  
                fRemove = (Not mblnReadOnly) And _
                  (Not FormPrint)
              End If
            Case REC_SEL_VALID_INVALID
              ' Calc invalid.
              ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
              asInvalidParameters(UBound(asInvalidParameters)) = "'" & sCalcName & "' calculation"

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)
          End Select

          If fRemove Then
            ' Calc invalid, deleted or hidden by another user. Remove it from this definition.
            If .Rows > 1 Then
              .RemoveItem iLoop
            Else
              .RemoveAll
            End If

            If Not FormPrint Then
              SSTab1.Tab = 2
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

  If UBound(asHiddenBySelfParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      'JPD 20040219 Fault 7897
      If Not fDefnAlreadyHidden Then
        sBigMessage = "This definition needs to be made hidden as the " & asHiddenBySelfParameters(1) & " is hidden."
      End If
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If (Not mblnForceHidden) And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the " & asHiddenBySelfParameters(1) & " is hidden."
        End If
      Else
        If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden as the " & asHiddenBySelfParameters(1) & " is hidden."
        End If
      End If
    Else
      sBigMessage = "The " & asHiddenBySelfParameters(1) & " will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
    End If
  ElseIf UBound(asHiddenBySelfParameters) > 1 Then
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
      For iLoop = 1 To UBound(asHiddenBySelfParameters)
        sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenBySelfParameters(iLoop)
      Next iLoop
    End If
  End If

  If UBound(asDeletedParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asDeletedParameters(1) & " has been deleted."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asDeletedParameters(1) & " will be removed from this definition as it has been deleted."
    End If
  ElseIf UBound(asDeletedParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been deleted :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been deleted :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asDeletedParameters)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asDeletedParameters(iLoop)
    Next iLoop
  End If

  If UBound(asHiddenByOtherParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asHiddenByOtherParameters(1) & " has been made hidden by another user."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asHiddenByOtherParameters(1) & " will be removed from this definition as it has been made hidden by another user."
    End If
  ElseIf UBound(asHiddenByOtherParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been made hidden by another user :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been made hidden by another user :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asHiddenByOtherParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenByOtherParameters(iLoop)
    Next iLoop
  End If

  If UBound(asInvalidParameters) = 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asInvalidParameters(1) & " is invalid."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asInvalidParameters(1) & " will be removed from this definition as it is invalid."
    End If
  ElseIf UBound(asInvalidParameters) > 1 Then
    If FormPrint Or mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters are invalid :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they are invalid :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asInvalidParameters)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asInvalidParameters(iLoop)
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
      sBigMessage = Me.Caption & " print failed. The definition is currently invalid : " & vbCrLf & vbCrLf & sBigMessage
    End If

    COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
  End If

  ForceDefinitionToBeHiddenIfNeeded = (Len(sBigMessage) = 0)

End Function


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



Private Sub cmdMoveDown_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  Dim lngCount As Long
  
  intSourceRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
  'strSourceRow = grdColumns.Columns(0).Text & vbTab & grdColumns.Columns(1).Text & vbTab & grdColumns.Columns(2).Text & vbTab & grdColumns.Columns(3).Text & vbTab & grdColumns.Columns(4).Text & vbTab & grdColumns.Columns(5).Text & vbTab & grdColumns.Columns(6).Text & vbTab & grdColumns.Columns(7).Text
  strSourceRow = vbNullString
  For lngCount = 0 To grdColumns.Columns.Count - 1
    strSourceRow = strSourceRow & _
        IIf(lngCount > 0, vbTab, "") & _
        grdColumns.Columns(lngCount).Text
  Next
  
  intDestinationRow = intSourceRow + 1
  grdColumns.MoveNext
  'strDestinationRow = grdColumns.Columns(0).Text & vbTab & grdColumns.Columns(1).Text & vbTab & grdColumns.Columns(2).Text & vbTab & grdColumns.Columns(3).Text & vbTab & grdColumns.Columns(4).Text & vbTab & grdColumns.Columns(5).Text & vbTab & grdColumns.Columns(6).Text & vbTab & grdColumns.Columns(7).Text
  strDestinationRow = vbNullString
  For lngCount = 0 To grdColumns.Columns.Count - 1
    strDestinationRow = strDestinationRow & _
        IIf(lngCount > 0, vbTab, "") & _
        grdColumns.Columns(lngCount).Text
  Next
  
  grdColumns.RemoveItem intDestinationRow
  grdColumns.RemoveItem intSourceRow
  
  grdColumns.AddItem strDestinationRow, intSourceRow
  grdColumns.AddItem strSourceRow, intDestinationRow
  
  grdColumns.SelBookmarks.RemoveAll
  
'  grdColumns.MoveNext
  '
  grdColumns.Bookmark = grdColumns.AddItemBookmark(intDestinationRow)
  '
  grdColumns.SelBookmarks.Add grdColumns.AddItemBookmark(intDestinationRow)
  
  UpdateButtonStatus

  Changed = True
  mblnBaseTableSpecificChanged = True
  
End Sub

Private Sub cmdMoveUp_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  Dim lngCount As Long
  
  
  intSourceRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
  strSourceRow = vbNullString
  For lngCount = 0 To grdColumns.Columns.Count - 1
    strSourceRow = strSourceRow & _
        IIf(lngCount > 0, vbTab, "") & _
        grdColumns.Columns(lngCount).Text
  Next
  
  intDestinationRow = intSourceRow - 1
  grdColumns.MovePrevious
  'strDestinationRow = grdColumns.Columns(0).Text & vbTab & grdColumns.Columns(1).Text & vbTab & grdColumns.Columns(2).Text & vbTab & grdColumns.Columns(3).Text & vbTab & grdColumns.Columns(4).Text & vbTab & grdColumns.Columns(5).Text & vbTab & grdColumns.Columns(6).Text & vbTab & grdColumns.Columns(7).Text
  strDestinationRow = vbNullString
  For lngCount = 0 To grdColumns.Columns.Count - 1
    strDestinationRow = strDestinationRow & _
        IIf(lngCount > 0, vbTab, "") & _
        grdColumns.Columns(lngCount).Text
  Next
  
  grdColumns.AddItem strSourceRow, intDestinationRow
  
  grdColumns.RemoveItem intSourceRow + 1
  
  grdColumns.SelBookmarks.RemoveAll
  grdColumns.SelBookmarks.Add grdColumns.AddItemBookmark(intDestinationRow)
  grdColumns.MovePrevious
  grdColumns.MovePrevious
  UpdateButtonStatus
  
  Changed = True
  mblnBaseTableSpecificChanged = True
  
End Sub

Private Function IsAuditColumn(lColumnID As Long) As Boolean

  Dim sSQL As String
  Dim rsColumns As Recordset
  Dim datData As DataMgr.clsDataAccess

  Set datData = New DataMgr.clsDataAccess

'  sSQL = "Select Size, Datatype From ASRSysColumns Where ColumnID = " & lColumnID
  sSQL = "Select Audit From ASRSysColumns Where ColumnID = " & lColumnID
  
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  IsAuditColumn = IIf(rsColumns(0).Value = 1, True, False)
  
  Set rsColumns = Nothing
  Set datData = Nothing
  
End Function

'###################################
Public Sub PrintDef(lExportID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsColumns As Recordset
  Dim sSQL As String
  Dim sTemp As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim strPrintString As String
  
  Dim lngLength As Long
  
  mlngExportID = lExportID
  
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysExportName.*, " & _
                                     "CONVERT(integer, ASRSysExportName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysExportName WHERE ID = " & mlngExportID)
                                        
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Print Definition"
    Set rsTemp = Nothing
    Exit Sub
  End If

  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        
        .TabsOnPage = 8     'MH20030929 Fault 6155
        
        ' First section --------------------------------------------------------
        .PrintHeader "Export : " & rsTemp!Name
        
        .PrintNormal "Category : " & GetObjectCategory(utlExport, mlngExportID)
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal "Owner : " & rsTemp!userName
        
        ' Access section --------------------------------------------------------
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop
          
        ' Data section --------------------------------------------------------
        .PrintTitle "Data"
        
        .PrintNormal "Base Table : " & datGeneral.GetTableName(rsTemp!BaseTable)
        
        If rsTemp!AllRecords Then
          .PrintNormal "Records : All Records"
        ElseIf rsTemp!picklist Then
          .PrintNormal "Records : '" & datGeneral.GetPicklistName(rsTemp!picklist) & "' picklist"
        ElseIf rsTemp!Filter Then
          .PrintNormal "Records : '" & datGeneral.GetFilterName(rsTemp!Filter) & "' filter"
        End If
        
        .PrintNormal
        
        .PrintNormal "Parent 1 Table : " & IIf(rsTemp!parent1table > 0, datGeneral.GetTableName(rsTemp!parent1table), "<None>")
        If (rsTemp!parent1picklist > 0) Then
          .PrintNormal "Parent 1 Records : '" & datGeneral.GetPicklistName(rsTemp!parent1picklist) & "' picklist"
        ElseIf (rsTemp!parent1filter > 0) Then
          .PrintNormal "Parent 1 Records : '" & datGeneral.GetFilterName(rsTemp!parent1filter) & "' filter"
        Else
          .PrintNormal "Parent 1 Records : N/A"
        End If
        
        .PrintNormal
        
        .PrintNormal "Parent 2 Table : " & IIf(rsTemp!parent2table > 0, datGeneral.GetTableName(rsTemp!parent2table), "<None>")
        If (rsTemp!parent2picklist > 0) Then
          .PrintNormal "Parent 2 Records : '" & datGeneral.GetPicklistName(rsTemp!parent2picklist) & "' picklist"
        ElseIf (rsTemp!parent2filter > 0) Then
          .PrintNormal "Parent 2 Records : '" & datGeneral.GetFilterName(rsTemp!parent2filter) & "' filter"
        Else
          .PrintNormal "Parent 2 Records : N/A"
        End If
        
        .PrintNormal
        
        .PrintNormal "Child Table : " & IIf(rsTemp!ChildTable > 0, datGeneral.GetTableName(rsTemp!ChildTable), "<None>")
        .PrintNormal "Child Filter : " & IIf(rsTemp!childFilter > 0, datGeneral.GetFilterName(rsTemp!childFilter), "<None>")
        .PrintNormal "Child Records : " & IIf(rsTemp!ChildMaxRecords = 0, "All Records", rsTemp!ChildMaxRecords)
        
        ' Now do the Columns Section
      
        Set rsColumns = datGeneral.GetRecords("SELECT * FROM ASRSysExportDetails WHERE ExportID = " & mlngExportID & " ORDER BY ID")
      
        .PrintTitle "Columns"
        
        If rsTemp!OutputFormat = fmtCMGFile Then
          'NPG20071217 Fault 12867
          ' .PrintBold "Type" & vbTab & vbTab & "Data" & vbTab & vbTab & vbTab & "CMG Code"
          'NPG20080711 Suggestion S000816
          ' .PrintBold "Type" & vbTab & vbTab & "Data" & vbTab & vbTab & vbTab & "CMG Code" & vbTab & "Conversion"
          .PrintBold "Type" & vbTab & "Data" & vbTab & vbTab & vbTab & "CMG Code" & vbTab & "Conversion" & vbTab & "Exclude if Empty"
        Else
          'NHRD17072003 Fault 6155
          'NPG20071217 Fault 12867 Added 'Conversion' heading
'          .PrintBold "Type" & _
'          vbTab & vbTab & "Data" & _
'          vbTab & vbTab & vbTab & "Size" & _
'          vbTab & "Decimals"

          .PrintBold "Type" & _
          vbTab & vbTab & "Data" & _
          vbTab & vbTab & vbTab & "Size" & _
          vbTab & "Decimals" & vbTab & "Conversion"
        End If
        
        Do While Not rsColumns.EOF
        
          Select Case rsColumns!Type
            Case Is = "C"
                pstrType = "Field"
            Case Is = "X"
                pstrType = "Calculation"
            Case Is = "T"
                pstrType = "Value"
            Case Is = "F"
                pstrType = "Other"
          End Select
        
          If rsTemp!OutputFormat = fmtCMGFile Then
          'NPG20080711 Suggestion S000816
'            .PrintNonBold pstrType & _
'            vbTab & vbTab & rsColumns!Data & _
'            vbTab & vbTab & vbTab & rsColumns!CMGColumnCode & _
'            vbTab & GetConvertCaseText(rsColumns!ConvertCase)
            .PrintNonBold pstrType & _
            vbTab & rsColumns!Data & _
            vbTab & vbTab & vbTab & rsColumns!CMGColumnCode & _
            vbTab & GetConvertCaseText(rsColumns!ConvertCase) & _
            vbTab & IIf((rsColumns!SuppressNulls), "True", "")
          Else
            'NPG20071217 Fault 12867
'            strPrintString = pstrType & _
'                            vbTab & vbTab & rsColumns!Data & _
'                            vbTab & vbTab & vbTab & "  " & rsColumns!fillerlength & _
'                            vbTab & "    "
            lngLength = IIf(rsColumns!fillerlength > 999999, 999999, rsColumns!fillerlength)
            strPrintString = pstrType & _
                            vbTab & vbTab & rsColumns!Data & _
                            vbTab & vbTab & vbTab & "  " & CStr(lngLength) & _
                            vbTab & "    "
            
            If (rsColumns!Type = "C") Then
              If (datGeneral.GetColumnDataType(rsColumns!ColExprID) <> SQLDataType.sqlNumeric) Then
                strPrintString = strPrintString & IIf(rsColumns!Decimals = 0, "-", rsColumns!Decimals)
              Else
                strPrintString = strPrintString & IIf(rsColumns!Decimals = 0, "0", rsColumns!Decimals)
              End If
            ElseIf (rsColumns!Type = "X") Then
              If Not (datGeneral.NumericColumn(rsColumns!Type, Trim(rsColumns!TableID), rsColumns!ColExprID)) Then
                strPrintString = strPrintString & IIf(rsColumns!Decimals = 0, "-", rsColumns!Decimals)
              Else
                strPrintString = strPrintString & IIf(rsColumns!Decimals = 0, "0", rsColumns!Decimals)
              End If
            Else
              strPrintString = strPrintString & IIf(rsColumns!Decimals = 0, "-", rsColumns!Decimals)
            End If
            
            
            'NPG20071217 Fault 12867
            strPrintString = strPrintString & vbTab & GetConvertCaseText(rsColumns!ConvertCase)
            
            
            
            .PrintNonBold strPrintString
          End If
          
          rsColumns.MoveNext
        Loop
        
        ' Now do the Sort Order n Options Section
      
        Set rsColumns = datGeneral.GetRecords("SELECT * FROM ASRSysExportDetails WHERE ExportID = " & mlngExportID & " AND SortOrderSequence > 0 AND Type = 'C' ORDER BY [SortOrderSequence]")
      
        .PrintTitle "Sort Order & Output Options"
        .PrintBold "Column" & vbTab & vbTab & vbTab & vbTab & "Sort Order"
          
        Do While Not rsColumns.EOF
         .PrintNonBold rsColumns!Data & vbTab & vbTab & vbTab & vbTab & _
                       IIf(rsColumns!SortOrder = "Asc", "Ascending", "Descending")
          rsColumns.MoveNext
        Loop
          
        .PrintNormal
        .PrintNormal
        
        Select Case rsTemp!OutputFormat
        Case fmtCSV
          .PrintNormal "Output Format : Delimited File"
          .PrintNormal "Delimiter : " & IIf(rsTemp!delimiter = "<Other>", rsTemp!otherdelimiter, rsTemp!delimiter)
          .PrintNormal "Quotes : " & IIf(rsTemp!Quotes = True, "Yes", "No")
        
          .PrintNormal "Split File into blocks : " & IIf(rsTemp!SplitFile = True, "Yes", "No")
          .PrintNormal "File Block Size : " & rsTemp!SplitFileSize
                 
        Case fmtFixedLengthFile
          .PrintNormal "Output Format : Fixed Length File"
        '  .PrintNormal "Filename : " & txtFilename.Text
        '  .PrintNormal
        
        Case fmtExcelWorksheet
          .PrintNormal "Output Format : Excel Worksheet"
        
        'Case fmtSQLTable
        '  .PrintNormal "Output Format : SQL Table"
        '  .PrintNormal "Table name : " & rsTemp!OutputFilename
        '  .PrintNormal
        
        Case fmtCMGFile
          .PrintNormal "Output Type : CMG File"
          .PrintNormal "File Export Code : " & rsTemp!CMGExportFileCode
          .PrintNormal "Record Identifier : " & datGeneral.GetColumnName(rsTemp!CMGExportRecordID)
          .PrintNormal "Commit after run : " & IIf(rsTemp!CMGExportUpdateAudit, "True", "False")
          '.PrintNormal "Filename : " & txtFilename.Text
        
        Case fmtXML
          .PrintNormal "Output Type : XML File"
          .PrintNormal "XML Custom Node Name : " & IIf(rsTemp!XMLDataNodeName = "", "<None>", rsTemp!XMLDataNodeName)
          .PrintNormal "XSD File : " & IIf(rsTemp!txtXSDFilename = "", "<None>", rsTemp!txtXSDFilename)
          .PrintNormal "XSD File Preserve Path : " & IIf(rsTemp!chkPreserveXSDPath, "True", "False")
          .PrintNormal "XML Transformation File : " & IIf(rsTemp!TransformFile = "", "<None>", rsTemp!TransformFile)
          .PrintNormal "XML Transformation File Preserve Path : " & IIf(rsTemp!PreserveTransformPath, "True", "False")
          .PrintNormal "Split nodes into individual files : " & IIf(rsTemp!chkSplitXMLNodesFile, "True", "False")

        End Select
        
        .PrintNormal
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
        .PrintNormal


        .PrintTitle "Options"

        If rsTemp!OutputFormat <> fmtSQLTable And _
           rsTemp!OutputFormat <> fmtCMGFile Then
          Select Case rsTemp!Header
          Case 0: .PrintNormal "Header Line : No Header"
          Case 1: .PrintNormal "Header Line : Column Names"
          Case 2: .PrintNormal "Header Line : Custom Header '" & rsTemp!HeaderText & "'"
          Case 3: .PrintNormal "Header Line : Custom Header and Column Names '" & Replace(rsTemp!HeaderText, vbTab, "   ") & "'"
          End Select
          .PrintNormal
  
          Select Case rsTemp!Footer
          Case 0: .PrintNormal "Footer Line : No Footer"
          Case 1: .PrintNormal "Footer Line : Column Names"
          Case 2: .PrintNormal "Footer Line : Custom Footer '" & Replace(rsTemp!FooterText, vbTab, "   ") & "'"
          End Select
          
          .PrintNormal "Force header if no records : " & IIf(chkForceHeader.Value, "Yes", "No")
          .PrintNormal "Omit header when appending to file : " & IIf(chkOmitHeader.Value, "Yes", "No")
          .PrintNormal
        End If
        
        If rsTemp!OutputFormat <> fmtSQLTable Then
          .PrintNormal "Date Format : " & rsTemp!DateFormat
          .PrintNormal "Date Separator : " & rsTemp!dateseparator
          .PrintNormal "Date Year Digits : " & rsTemp!Dateyeardigits
        End If
        
        .PrintEnd
        .PrintConfirm "Export : " & rsTemp!Name, "Export Definition"
      End If
  
    End With
    
  End If
  
  Set rsTemp = Nothing
  Set rsColumns = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing export definition failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"

End Sub


Private Function IsRecordSelectionValid() As Boolean

  Dim pvarbookmark As Variant
  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim aRowsToDelete() As String
  Dim sMsgTxt As String
  
  Dim blnDeletedCalcs As Boolean
  Dim blnHiddenCalcs As Boolean
  
  On Error GoTo Valid_ERROR
  
  ReDim aRowsToDelete(2, 0)
  
  ' Base Table First
  
  If optBaseFilter.Value Then
    sSQL = "SELECT * FROM AsrSysExpressions WHERE ExprID = " & txtBaseFilter.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
      ' filter no longer exists !
      COAMsgBox "The '" & txtBaseFilter.Text & "' filter has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtBaseFilter.Tag = 0
      txtBaseFilter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) Then
      ' Filter has been made hidden by its owner
      COAMsgBox "The '" & txtBaseFilter.Text & "' filter has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtBaseFilter.Tag = 0
      txtBaseFilter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  ElseIf optBasePicklist.Value Then
    sSQL = "SELECT * FROM AsrSysPicklistName WHERE PickListID = " & txtBasePicklist.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
      ' picklist no longer exists !
      COAMsgBox "The '" & txtBasePicklist.Text & "' picklist has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtBasePicklist.Tag = 0
      txtBasePicklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) Then
      ' picklist has been made hidden by its owner
      COAMsgBox "The '" & txtBasePicklist.Text & "' picklist has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtBasePicklist.Tag = 0
      txtBasePicklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  End If
  
  ' Parent 1 Table
  If txtParent1Filter.Tag > 0 Then
    sSQL = "SELECT * FROM AsrSysExpressions WHERE ExprID = " & txtParent1Filter.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
      ' filter no longer exists !
      COAMsgBox "The '" & txtParent1Filter.Text & "' filter has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent1Filter.Tag = 0
      txtParent1Filter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) Then
      ' filter has been made hidden by its owner
      COAMsgBox "The '" & txtParent1Filter.Text & "' filter has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent1Filter.Tag = 0
      txtParent1Filter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  ElseIf optParent1Picklist.Value And txtParent1Picklist.Tag > 0 Then
    sSQL = "SELECT * FROM AsrSysPicklistName WHERE PickListID = " & txtParent1Picklist.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    
    If rsTemp.BOF And rsTemp.EOF Then
      ' Picklist has been deleted by another user
      COAMsgBox "The '" & txtParent1Picklist.Text & "' picklist has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent1Picklist.Tag = 0
      txtParent1Picklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(rsTemp!userName) <> LCase(gsUserName) Then
      ' Picklist has been made hidden by its owner
      COAMsgBox "The '" & txtParent1Picklist.Text & "' picklist has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent1Picklist.Tag = 0
      txtParent1Picklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  End If
  
  ' Parent 2 Table
  If txtParent2Filter.Tag > 0 Then
    sSQL = "SELECT * FROM AsrSysExpressions WHERE ExprID = " & txtParent2Filter.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
      ' filter no longer exists !
      COAMsgBox "The '" & txtParent2Filter.Text & "' filter has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent2Filter.Tag = 0
      txtParent2Filter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) Then
      ' filter has been made hidden by its owner
      COAMsgBox "The '" & txtParent2Filter.Text & "' filter has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent2Filter.Tag = 0
      txtParent2Filter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  ElseIf optParent2Picklist.Value And txtParent2Picklist.Tag > 0 Then
    sSQL = "SELECT * FROM AsrSysPicklistName WHERE PickListID = " & txtParent2Picklist.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    
    If rsTemp.BOF And rsTemp.EOF Then
      ' Picklist has been deleted by another user
      COAMsgBox "The '" & txtParent2Picklist.Text & "' picklist has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent2Picklist.Tag = 0
      txtParent2Picklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(rsTemp!userName) <> LCase(gsUserName) Then
      ' Picklist has been made hidden by its owner
      COAMsgBox "The '" & txtParent2Picklist.Text & "' picklist has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtParent2Picklist.Tag = 0
      txtParent2Picklist.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  End If
  
  ' Child Table
  If txtChildFilter.Tag > 0 Then
    sSQL = "SELECT * FROM AsrSysExpressions WHERE ExprID = " & txtChildFilter.Tag
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    If rsTemp.BOF And rsTemp.EOF Then
      ' filter no longer exists !
      COAMsgBox "The '" & txtChildFilter.Text & "' filter has been deleted by another user.", vbExclamation + vbOKOnly, "Export"
      txtChildFilter.Tag = 0
      txtChildFilter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    ElseIf rsTemp!Access = "HD" And LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) Then
      ' filter has been made hidden by its owner
      COAMsgBox "The '" & txtChildFilter.Text & "' filter has been made hidden by another user.", vbExclamation + vbOKOnly, "Export"
      txtChildFilter.Tag = 0
      txtChildFilter.Text = "<None>"
      Set rsTemp = Nothing
      IsRecordSelectionValid = False
      Exit Function
    End If
  End If
  
  ' Now check the columns for any calcs that have been deleted since they were added
  ' to this definition
  '
  ' aRowsToDelete(0, x) = Calculation Name
  ' aRowsToDelete(1, x) = Calculation ID
  ' aRowsToDelete(2, x) = Row in the grid
  '
  
  'lCount = 0
  'grdColumns.MoveFirst
    
  
  
  'Do Until lCount = grdColumns.Rows
  For lCount = 0 To grdColumns.Rows - 1
    pvarbookmark = grdColumns.GetBookmark(lCount)
    
    'MH20040602 Fault 8737
    'If grdColumns.Columns("Type").Text = "X" Then
    If grdColumns.Columns("Type").CellValue(pvarbookmark) = "X" Then
      Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & grdColumns.Columns("ColExprID").CellValue(pvarbookmark))
      If rsTemp.BOF And rsTemp.EOF Then
        ReDim Preserve aRowsToDelete(2, UBound(aRowsToDelete, 2) + 1)
        aRowsToDelete(0, UBound(aRowsToDelete, 2)) = grdColumns.Columns("Data").CellValue(pvarbookmark)
        aRowsToDelete(1, UBound(aRowsToDelete, 2)) = grdColumns.Columns("ColExprID").CellValue(pvarbookmark)
        aRowsToDelete(2, UBound(aRowsToDelete, 2)) = lCount
        blnDeletedCalcs = True    'MH20001102
      ElseIf rsTemp!Access = "HD" And LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) Then
        ReDim Preserve aRowsToDelete(2, UBound(aRowsToDelete, 2) + 1)
        aRowsToDelete(0, UBound(aRowsToDelete, 2)) = grdColumns.Columns("Data").CellValue(pvarbookmark)
        aRowsToDelete(1, UBound(aRowsToDelete, 2)) = grdColumns.Columns("ColExprID").CellValue(pvarbookmark)
        aRowsToDelete(2, UBound(aRowsToDelete, 2)) = lCount
        blnHiddenCalcs = True     'MH20001102
      End If
    End If
    'grdColumns.MoveNext
    'lCount = lCount + 1
  'Loop
  Next
  
  If UBound(aRowsToDelete, 2) > 0 Then
    'sMsgTxt = "The following calculations have been deleted by another user" & vbCrLf & _
             "and will be removed from this Export:" & vbCrLf & vbCrLf
    'For lCount = 1 To UBound(aRowsToDelete, 2)
    '  sMsgTxt = sMsgTxt & aRowsToDelete(0, lCount) & Space(35 - Len(aRowsToDelete(0, lCount))) & "(Row " & aRowsToDelete(2, lCount) + 1 & ")" & vbCrLf
    'Next lCount
    ' RH - 09/10/00 - QA prefer generic msg rather than specific calc names
    '  sMsgTxt = "One or more calculation(s) have been deleted and/or made hidden" & vbCrLf & _
    '            "by their owner and will be removed from this Export."
    '  COAMsgBox sMsgTxt, vbOKOnly + vbExclamation, "Export"
    
    
    'MH20001102 Fault 1248
    If blnHiddenCalcs Then
      sMsgTxt = "This definition contains one or more calculation(s) which have been made hidden by another user." & vbCrLf & _
                "They will be automatically removed from this definition."
      COAMsgBox sMsgTxt, vbOKOnly + vbExclamation, "Export"
    End If
    If blnDeletedCalcs Then
      sMsgTxt = "This definition contains one or more calculation(s) which have been deleted by another user." & vbCrLf & _
                "They will be automatically removed from this definition."
      COAMsgBox sMsgTxt, vbOKOnly + vbExclamation, "Export"
    End If
              
    ' Delete the rows. Work from bottom-up otherwise the row indexes will be wrong!
    For lCount = UBound(aRowsToDelete, 2) To 1 Step -1
      grdColumns.RemoveItem (aRowsToDelete(2, lCount))
    Next lCount
    
    ' Select the last row if there are any
    If grdColumns.Rows > 0 Then
      grdColumns.MoveLast
      grdColumns.SelBookmarks.Add grdColumns.Bookmark
    End If
    
    UpdateButtonStatus
    
    ' Tidy up
    Set rsTemp = Nothing
    IsRecordSelectionValid = False
    Exit Function
    
  End If
  
  Set rsTemp = Nothing
  IsRecordSelectionValid = True
  Exit Function
  
Valid_ERROR:

  COAMsgBox "Error whilst checking if picklist/filters/calculations selected in the" & vbCrLf & _
         "Export still exist:" & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly
  Set rsTemp = Nothing
  IsRecordSelectionValid = False
  
End Function

'Private Sub txtSQLTableName_Change()
'
'  If mblnLoading Then Exit Sub
'  Changed = True
'
'End Sub

Private Sub ShowRelevantColumns()

'TM20011012 Fault 2197
'Set the visible property of the Decimals column.

'Display the appropraite columns for the selected output type
  If optOutputFormat(fmtCMGFile).Value Then
    lng_DataCOLUMNWIDTH = 5045.236

    grdColumns.Columns(7).Visible = False
    grdColumns.Columns(4).Visible = False
    grdColumns.Columns(5).Visible = mbCMGExportFieldCode
    grdColumns.Columns(6).Visible = True
    grdColumns.Columns(8).Visible = False
    
    'TM20011012 Fault 2197
    'Set the width of all the visible columns.
    grdColumns.Columns(3).Width = lng_DataCOLUMNWIDTH
    grdColumns.Columns(5).Width = lng_LengthCOLUMNWIDTH
    grdColumns.Columns(6).Width = lng_AuditCOLUMNWIDTH
    
  ElseIf optOutputFormat(fmtXML).Value Then
    grdColumns.Columns(4).Visible = False
    grdColumns.Columns(5).Visible = False
    grdColumns.Columns(6).Visible = True
    grdColumns.Columns(7).Visible = False
    grdColumns.Columns(8).Visible = True
    
    lng_DataCOLUMNWIDTH = 3045.236
    
    grdColumns.Columns(3).Width = lng_DataCOLUMNWIDTH
    grdColumns.Columns(6).Width = lng_AuditCOLUMNWIDTH
    grdColumns.Columns(8).Width = (lng_LengthCOLUMNWIDTH + lng_DecimalCOLUMNWIDTH) + lng_AuditCOLUMNWIDTH
    
  Else
    lng_DataCOLUMNWIDTH = 5045.236
  
    grdColumns.Columns(7).Visible = True
    grdColumns.Columns(4).Visible = True
    grdColumns.Columns(5).Visible = False
    grdColumns.Columns(6).Visible = False
    grdColumns.Columns(8).Visible = False
    
    'TM20011012 Fault 2197
    'Set the width of all the visible columns.
    grdColumns.Columns(3).Width = lng_DataCOLUMNWIDTH
    grdColumns.Columns(4).Width = lng_LengthCOLUMNWIDTH
    grdColumns.Columns(7).Width = lng_DecimalCOLUMNWIDTH

  End If
  
  'Set the height of the rows.
  grdColumns.RowHeight = lng_GRIDROWHEIGHT
  
'Resize the data column so everything fits in the grid
'grdColumns.Columns(3).Width = IIf(psOutputCode = "C" And mbCMGExportFieldCode, lng_DataCOLUMNWIDTH, lng_DataCOLUMNWIDTH)
  UpdateButtonStatus

End Sub

Public Function ValidNameChar(ByVal piAsciiCode As Integer, ByVal piPosition As Integer) As Integer
  ' Validate the characters used to create table and column names.
  On Error GoTo ErrorTrap
  
  If piAsciiCode = Asc(" ") Then
    ' Substitute underscores for spaces.
    If piPosition <> 0 Then
      piAsciiCode = Asc("_")
    Else
      piAsciiCode = 0
    End If
  Else
    ' Allow only pure alpha-numerics and underscores.
    ' Do not allow numerics in the first chracter position.
  
  ' RH 15/08/2000 - BUG...we should be able to start filter/calcs with a number char
    If Not (piAsciiCode = 8 Or piAsciiCode = Asc("_") Or _
      (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9")) Or _
      (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or _
      (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z"))) Then
      piAsciiCode = 0
    End If
  End If
  
  ValidNameChar = piAsciiCode
  Exit Function
  
ErrorTrap:
  ValidNameChar = 0
  Err = False
  
End Function

'Private Sub txtSQLTableName_KeyPress(KeyAscii As Integer)
'
'  ' Validate the character entered.
'  KeyAscii = ValidNameChar(KeyAscii, txtSQLTableName.SelStart)
'
'End Sub


Private Sub CheckIfCMGEnabled()

  fraCMGFile.Visible = gbCMGEnabled
  optOutputFormat(fmtCMGFile).Visible = gbCMGEnabled

  If gbCMGEnabled Then
    mbCMGExportFileCode = GetSystemSetting("CMGExport", "FileCode", False)
    mbCMGExportFieldCode = GetSystemSetting("CMGExport", "FieldCode", False)
  Else
    mbCMGExportFileCode = False
    mbCMGExportFieldCode = False
  End If

End Sub

Private Function GetConvertCaseText(ByVal iConvertCase As Integer) As String
  Select Case iConvertCase
    Case 1
        GetConvertCaseText = "Uppercase"
    Case 2
        GetConvertCaseText = "Lowercase"
    Case Else
        GetConvertCaseText = ""
  End Select
End Function

Private Sub cboCategory_Click()
  Changed = True
End Sub

Private Sub txtXMLDataNodeName_Change()
  Changed = True
End Sub
