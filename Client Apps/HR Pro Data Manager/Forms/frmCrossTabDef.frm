VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCrossTabDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cross Tab Definition"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   1025
   Icon            =   "frmCrossTabDef.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8400
      TabIndex        =   49
      Top             =   4740
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7140
      TabIndex        =   48
      Top             =   4740
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4560
      Left            =   90
      TabIndex        =   50
      Top             =   75
      Width           =   9500
      _ExtentX        =   16748
      _ExtentY        =   8043
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   520
      OLEDropMode     =   1
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
      TabPicture(0)   =   "frmCrossTabDef.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinition(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDefinition(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Colu&mns"
      TabPicture(1)   =   "frmCrossTabDef.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraColumns(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraColumns(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "O&utput"
      TabPicture(2)   =   "frmCrossTabDef.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOutputFormat"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraOutputDestination"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   3975
         Left            =   -72255
         TabIndex        =   60
         Top             =   405
         Width           =   6600
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3315
            TabIndex        =   79
            Tag             =   "0"
            Top             =   3460
            Width           =   3180
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3315
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1760
            Width           =   2880
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3315
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1240
            Width           =   3180
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3315
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   2160
            Width           =   3180
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3315
            TabIndex        =   77
            Top             =   3060
            Width           =   3180
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3315
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   2660
            Width           =   2880
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            DisabledPicture =   "frmCrossTabDef.frx":0060
            Enabled         =   0   'False
            Height          =   315
            Left            =   6150
            Picture         =   "frmCrossTabDef.frx":03C1
            TabIndex        =   69
            Top             =   1760
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            DisabledPicture =   "frmCrossTabDef.frx":0722
            Enabled         =   0   'False
            Height          =   315
            Left            =   6150
            Picture         =   "frmCrossTabDef.frx":0A83
            TabIndex        =   75
            Top             =   2660
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   62
            Top             =   850
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   63
            Top             =   1300
            Width           =   1650
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   66
            Top             =   1820
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   72
            Top             =   2720
            Width           =   1560
         End
         Begin VB.CheckBox chkPreview 
            Caption         =   "P&review on screen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   195
            TabIndex        =   61
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
            TabIndex        =   78
            Top             =   3525
            Width           =   1155
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1845
            TabIndex        =   67
            Top             =   1815
            Width           =   1005
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1845
            TabIndex        =   76
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
            TabIndex        =   73
            Top             =   2715
            Width           =   1290
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1845
            TabIndex        =   70
            Top             =   2220
            Width           =   1350
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1845
            TabIndex        =   64
            Top             =   1305
            Width           =   1455
         End
      End
      Begin VB.Frame fraOutputFormat 
         Caption         =   "Output Format :"
         Height          =   3990
         Left            =   -74880
         TabIndex        =   52
         Top             =   405
         Width           =   2500
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel P&ivot Table"
            Height          =   195
            Index           =   6
            Left            =   200
            TabIndex        =   59
            Top             =   2800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Char&t"
            Height          =   195
            Index           =   5
            Left            =   200
            TabIndex        =   58
            Top             =   2400
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   57
            Top             =   2000
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   56
            Top             =   1600
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   55
            Top             =   1200
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CS&V File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   54
            Top             =   800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "D&ata Only"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   53
            Top             =   400
            Value           =   -1  'True
            Width           =   1900
         End
      End
      Begin VB.Frame fraDefinition 
         Caption         =   "Data :"
         Height          =   2015
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   2385
         Width           =   9200
         Begin VB.CheckBox chkPrintFilterHeader 
            Caption         =   "Display &title in the report header"
            Enabled         =   0   'False
            Height          =   240
            Left            =   4815
            TabIndex        =   18
            Tag             =   "PrintFilterHeader"
            Top             =   1560
            Width           =   3165
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1080
            Width           =   1950
         End
         Begin VB.TextBox txtPicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   705
            Width           =   1950
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   5715
            TabIndex        =   15
            Top             =   1120
            Width           =   840
         End
         Begin VB.OptionButton optPicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   5715
            TabIndex        =   12
            Top             =   750
            Width           =   885
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   5715
            TabIndex        =   11
            Top             =   365
            Value           =   -1  'True
            Width           =   540
         End
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   315
            Width           =   2910
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   14
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            TabIndex        =   17
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   8
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Index           =   5
            Left            =   4815
            TabIndex        =   10
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame fraDefinition 
         Height          =   1950
         Index           =   0
         Left            =   135
         TabIndex        =   51
         Top             =   405
         Width           =   9200
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5625
            MaxLength       =   30
            TabIndex        =   5
            Top             =   315
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   1
            Top             =   315
            Width           =   2910
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1620
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   705
            Width           =   2910
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1080
            Left            =   5625
            TabIndex        =   80
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
            stylesets(0).Picture=   "frmCrossTabDef.frx":0DE4
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
            stylesets(1).Picture=   "frmCrossTabDef.frx":0E00
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
            ForeColor       =   -2147483630
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Index           =   2
            Left            =   4815
            TabIndex        =   4
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   0
            Top             =   365
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   2
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Index           =   3
            Left            =   4815
            TabIndex        =   6
            Top             =   750
            Width           =   735
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Intersection :"
         Height          =   2015
         Index           =   1
         Left            =   -74865
         TabIndex        =   39
         Top             =   2385
         Width           =   9200
         Begin VB.CheckBox chkThousandSeparators 
            Caption         =   "Use 1000 &separators (,)"
            Height          =   330
            Left            =   5100
            TabIndex        =   47
            Top             =   1290
            Width           =   2520
         End
         Begin VB.CheckBox chkPercentageofPage 
            Caption         =   "Percentage of &Page"
            Height          =   195
            Left            =   5100
            TabIndex        =   45
            Top             =   690
            Width           =   2250
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "frmCrossTabDef.frx":0E1C
            Left            =   1620
            List            =   "frmCrossTabDef.frx":0E1E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   720
            Width           =   3000
         End
         Begin VB.CheckBox chkPercentage 
            Caption         =   "Percentage of &Type"
            Height          =   195
            Left            =   5100
            TabIndex        =   44
            Top             =   360
            Width           =   2250
         End
         Begin VB.ComboBox cboIntersectionCol 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   315
            Width           =   3000
         End
         Begin VB.CheckBox chkSuppressZeros 
            Caption         =   "Suppress &Zeros"
            Height          =   195
            Left            =   5100
            TabIndex        =   46
            Top             =   1020
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Type :"
            Height          =   195
            Index           =   14
            Left            =   225
            TabIndex        =   42
            Top             =   765
            Width           =   570
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Column :"
            Height          =   195
            Index           =   13
            Left            =   225
            TabIndex        =   40
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Headings && Breaks :"
         Height          =   1950
         Index           =   0
         Left            =   -74865
         TabIndex        =   19
         Top             =   405
         Width           =   9200
         Begin VB.ComboBox cboHorizontalCol 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   585
            Width           =   3000
         End
         Begin VB.ComboBox cboPageBreakCol 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1395
            Width           =   3000
         End
         Begin VB.ComboBox cboVerticalCol 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   990
            Width           =   3000
         End
         Begin TDBNumberCtrl.TDBNumber mskHorizontalRange 
            Height          =   315
            Index           =   0
            Left            =   5100
            TabIndex        =   26
            Top             =   585
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   99999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0E20
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskHorizontalRange 
            Height          =   315
            Index           =   1
            Left            =   6457
            TabIndex        =   27
            Top             =   585
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   9999999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0E3C
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskHorizontalRange 
            Height          =   315
            Index           =   2
            Left            =   7815
            TabIndex        =   28
            Top             =   585
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   99999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0E58
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskVerticalRange 
            Height          =   315
            Index           =   0
            Left            =   5100
            TabIndex        =   31
            Top             =   990
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   99999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0E74
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskVerticalRange 
            Height          =   315
            Index           =   1
            Left            =   6457
            TabIndex        =   32
            Top             =   990
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   9999999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0E90
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskVerticalRange 
            Height          =   315
            Index           =   2
            Left            =   7815
            TabIndex        =   33
            Top             =   990
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   99999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0EAC
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskPageBreakRange 
            Height          =   315
            Index           =   0
            Left            =   5100
            TabIndex        =   36
            Top             =   1395
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   99999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0EC8
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskPageBreakRange 
            Height          =   315
            Index           =   1
            Left            =   6457
            TabIndex        =   37
            Top             =   1395
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   9999999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0EE4
            MousePointer    =   0
         End
         Begin TDBNumberCtrl.TDBNumber mskPageBreakRange 
            Height          =   315
            Index           =   2
            Left            =   7815
            TabIndex        =   38
            Top             =   1395
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   -2147483640
            MaxValue        =   99999999
            MinValue        =   -9999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0"
            Format          =   "#########0"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483633
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   0
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCrossTabDef.frx":0F00
            MousePointer    =   0
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Column"
            Height          =   195
            Index           =   6
            Left            =   1620
            TabIndex        =   20
            Top             =   270
            Width           =   3000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Horizontal :"
            Height          =   195
            Index           =   10
            Left            =   225
            TabIndex        =   24
            Top             =   645
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Page Break :"
            Height          =   195
            Index           =   12
            Left            =   225
            TabIndex        =   34
            Top             =   1455
            Width           =   1200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Vertical :"
            Height          =   195
            Index           =   11
            Left            =   225
            TabIndex        =   29
            Top             =   1050
            Width           =   1110
         End
         Begin VB.Label lblRange 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   195
            Index           =   0
            Left            =   5100
            TabIndex        =   21
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label lblRange 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Stop"
            Height          =   195
            Index           =   1
            Left            =   6457
            TabIndex        =   22
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label lblRange 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Increment"
            Height          =   195
            Index           =   2
            Left            =   7815
            TabIndex        =   23
            Top             =   270
            Width           =   1200
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCrossTabDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjOutputDef As clsOutputDef

Private mblnReadOnly As Boolean
Private datData As HRProDataMgr.clsDataAccess          'DataAccess Class
Private fOK As Boolean
Private mrsColumns As New Recordset
Private mstrPrimaryTable As String
Private mlngCrossTabID As Long
'Private mblnChanged As Boolean
Private mlngTimeStamp As Long
Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mblnDefinitionCreator As Boolean
Private mblnForceHidden As Boolean
Private mblnRecordSelectionInvalid As Boolean

Private mstrBaseTablePicklistFilter As String


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
  If (Len(txtPicklist.Tag) > 0) And (Val(txtPicklist.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtPicklist.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly)
        If fRemove Then
          sBigMessage = "The '" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"
  
          fRemove = (Not mblnReadOnly)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtPicklist.Tag = 0
      txtPicklist.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Base Table Filter
  If Len(txtFilter.Tag) > 0 And Val(txtFilter.Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtFilter.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly)

        If fRemove Then
          sBigMessage = "The '" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtFilter.Tag = 0
      txtFilter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Construct one big message with all of the required error messages.
  sBigMessage = ""

  If UBound(asHiddenBySelfParameters) = 1 Then
    If mblnReadOnly Then
      'JPD 20040308 Fault 7897
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
    If mblnReadOnly Then
      'JPD 20040308 Fault 7897
      If Not fDefnAlreadyHidden Then
        sBigMessage = "This definition needs to be made hidden as the following parameters are hidden :" & vbCrLf
      End If
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If Not mblnForceHidden And (Not fOnlyFatalMessages) Then
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
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asDeletedParameters(1) & " has been deleted."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asDeletedParameters(1) & " will be removed from this definition as it has been deleted."
    End If
  ElseIf UBound(asDeletedParameters) > 1 Then
    If mblnReadOnly Then
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
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asHiddenByOtherParameters(1) & " has been made hidden by another user."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asHiddenByOtherParameters(1) & " will be removed from this definition as it has been made hidden by another user."
    End If
  ElseIf UBound(asHiddenByOtherParameters) > 1 Then
    If mblnReadOnly Then
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
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asInvalidParameters(1) & " is invalid."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The " & asInvalidParameters(1) & " will be removed from this definition as it is invalid."
    End If
  ElseIf UBound(asInvalidParameters) > 1 Then
    If mblnReadOnly Then
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

  If mblnForceHidden And (Not fNeedToForceHidden) And (Not fOnlyFatalMessages) Then
    sBigMessage = "This definition no longer has to be hidden." & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
      sBigMessage
  End If

  mblnForceHidden = fNeedToForceHidden
  ForceAccess

  If Len(sBigMessage) > 0 Then
    MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
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




Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOk.Enabled = blnChanged
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngCrossTabID
End Property



Private Function ErrorMsgbox(strMessage As String) As Boolean
  Screen.MousePointer = vbDefault
  MsgBox strMessage & vbCrLf & Err.Description, vbCritical, "Cross Tab"
End Function

Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional lCrossTabID As Long) As Boolean

  Dim lngCount As Long
  
  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass
  
  mblnLoading = True
  fOK = True

  Set datData = New HRProDataMgr.clsDataAccess
  
  For lngCount = 0 To 2
    datGeneral.FormatTDBNumberControl mskHorizontalRange(lngCount)
    datGeneral.FormatTDBNumberControl mskVerticalRange(lngCount)
    datGeneral.FormatTDBNumberControl mskPageBreakRange(lngCount)
  Next

  LoadPrimaryCombo
  'LoadPrimaryDependantCombos
  LoadOtherCombos

  If bNew Then
    'cboBaseTable.ListIndex = 0
    LoadPrimaryDependantCombos
    If cboHorizontalCol.ListCount > 0 Then
      cboHorizontalCol.ListIndex = 0
    End If
    mlngCrossTabID = 0          'Indicate new record
    optAllRecords.Value = True  'Default to all records
    'optOutput(0).Value = True
    txtUserName = gsUserName
    mblnDefinitionCreator = True
    PopulateAccessGrid
    Me.Changed = False

  Else
    mlngCrossTabID = lCrossTabID
    mblnFromCopy = bCopy
    
    PopulateAccessGrid
    
    Call RetreiveDefinition
    
    If fOK Then
      Me.Changed = False
    
      'Set ID to zero so that if saved a new record will
      'be created rather than updating this definition
      If mblnFromCopy Then
        mlngCrossTabID = 0        'Indicate new record
        Me.Changed = True
      Else
        Me.Changed = (mblnRecordSelectionInvalid And Not mblnReadOnly)
      End If
    End If

  End If

  mblnLoading = False
  Screen.MousePointer = vbDefault

  Initialise = fOK

Exit Function

LocalErr:
  ErrorMsgbox "Error with Cross Tab Definition"

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
  Set rsAccess = GetUtilityAccessRecords(utlCrossTab, mlngCrossTabID, mblnFromCopy)
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


Private Sub cboPrinterName_Click()
  Me.Changed = True
End Sub

Private Sub cboSaveExisting_Click()
  Me.Changed = True
End Sub

'Private Sub cboExportTo_Click()
'
'  ' If user changes exportto, and a filename has already been selected
'  ' then change the extension of the filename automatically
'
'  Dim sText As String
'
'  If txtFilename.Text <> "" Then
'    sText = txtFilename.Text
'    Select Case UCase(cboExportTo.Text)
'    Case "MICROSOFT EXCEL": Mid(sText, Len(txtFilename.Text) - 2, 3) = "xls"
'    Case "HTML":  Mid(sText, Len(txtFilename.Text) - 2, 3) = "htm"
'    Case "MICROSOFT WORD":  Mid(sText, Len(txtFilename.Text) - 2, 3) = "doc"
'    End Select
'    txtFilename.Text = sText
'  End If
'
'  Me.Changed = True
'
'End Sub

Private Sub cboType_Click()
  Me.Changed = True
End Sub

'Private Sub chkCloseApplication_Click()
'  Me.Changed = True
'End Sub

Private Sub chkPercentage_Click()
  
  chkPercentageofPage.Enabled = _
    (chkPercentage.Value = vbChecked And _
     Val(cboPageBreakCol.Tag) > 0)
  
  If chkPercentageofPage.Enabled = False Then
    chkPercentageofPage.Value = vbUnchecked
  End If

  Me.Changed = True

End Sub

Private Sub chkPercentageOfPage_Click()
  Me.Changed = True
End Sub

Private Sub chkPreview_Click()
  Changed = True
End Sub

Private Sub chkPrintFilterHeader_Click()
  Changed = True
End Sub

'Private Sub chkSave_Click()
'
'  Dim blnSaveChecked As Boolean
'
'  blnSaveChecked = (chkSave = vbChecked)
'
'  txtFilename.Enabled = blnSaveChecked
'  cmdFilename.Enabled = blnSaveChecked
'  chkCloseApplication.Enabled = blnSaveChecked
'
'  If blnSaveChecked = False Then
'    txtFilename.Text = vbNullString
'    chkCloseApplication.Value = vbUnchecked
'  End If
'
'  Me.Changed = True
'
'End Sub

Private Sub chkSuppressZeros_Click()
  Me.Changed = True
End Sub

Private Sub chkThousandSeparators_Click()
  Me.Changed = True
End Sub

Private Sub Form_Load()
  
  SSTab1.Tab = 0
  Call SSTab1_Click(0)
  
  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos True, True, True
  
  grdAccess.RowHeight = 239
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'If form is visible then assume that unload is the same as pressing
  'the cancel button.  Do not unload !!!!!
  If Me.Visible Then
    Cancel = True
    Call cmdCancel_Click
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mobjOutputDef = Nothing
  frmMain.RefreshMainForm Me, True
End Sub

'Private Sub txt_GotFocus(txtTemp As TextBox)
Private Sub txt_GotFocus(txtTemp As Variant)
  With txtTemp
    .SelStart = 0
    .SelLength = Len(txtTemp.Text)
  End With
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


Private Sub txtDesc_Change()
  Me.Changed = True
End Sub

Private Sub txtDesc_LostFocus()
  cmdOk.Default = True
End Sub

Private Sub txtEmailGroup_Change()
  Me.Changed = True
End Sub

Private Sub txtEmailSubject_Change()
  Me.Changed = True
End Sub

Private Sub txtEmailAttachAs_Change()
  Me.Changed = True
End Sub

Private Sub txtFilename_Change()
  Me.Changed = True
End Sub

Private Sub txtName_Change()
  Me.Changed = True
End Sub

Private Sub txtName_GotFocus()
  txt_GotFocus txtName
End Sub

Private Sub txtDesc_GotFocus()
  txt_GotFocus txtDesc
  cmdOk.Default = False
End Sub

Private Sub cboBaseTable_Click()

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  
  If mblnLoading Or (cboBaseTable.Text = mstrPrimaryTable) Then
    Exit Sub
  End If


  intMBResponse = vbYes

  'MH20021004 Only prompt if editting a definition.
  'If mstrPrimaryTable <> vbNullString Then
  If mlngCrossTabID > 0 Then
    strMBText = "Changing the Base Table will reset all of the selected columns. Continue ?"
    intMBButtons = vbQuestion + vbYesNo   'Cancel
    strMBTitle = Me.Caption
    intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)
  End If

  If intMBResponse = vbYes Then
    mstrPrimaryTable = cboBaseTable.Text
    Call LoadPrimaryDependantCombos
    optAllRecords.Value = True
    Me.Changed = True
  Else
    SetComboText cboBaseTable, mstrPrimaryTable
  End If

  ForceDefinitionToBeHiddenIfNeeded

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

  'MH20040726 Fault 8956
  'Enable Print Header checkbox even if loading!
  'If mblnLoading Then
  '  Exit Sub
  'End If

  cmdPicklist.Enabled = blnPicklist
  If blnPicklist = False Then
    txtPicklist.Text = vbNullString
    txtPicklist.Tag = 0
  ElseIf txtPicklist.Text = vbNullString Then
    txtPicklist.Text = "<None>"
  End If
  
  cmdFilter.Enabled = blnFilter
  If blnFilter = False Then
    txtFilter.Text = vbNullString
    txtFilter.Tag = 0
  ElseIf txtFilter.Text = vbNullString Then
    txtFilter.Text = "<None>"
  End If

  chkPrintFilterHeader.Enabled = (blnPicklist Or blnFilter)
  If Not mblnLoading Then
    If Not (blnPicklist Or blnFilter) Then
      chkPrintFilterHeader.Value = vbUnchecked
    End If

    ForceDefinitionToBeHiddenIfNeeded
    Me.Changed = True
  End If

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
  
  'If cboBaseTable.Text = "<None>" Then
  '  MsgBox "No primary table specified", vbExclamation, Me.Caption
  '  Exit Sub
  'End If

  Screen.MousePointer = vbHourglass

  fExit = False

  'set the sql to only include tables for the selected export base table
  'sSQL = "Select Name, PickListID From ASRSysPickListName"
  'sSQL = sSQL & " Where TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
  'sSQL = "TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)

  With frmDefSel
      
    .TableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)
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
            If .InitialisePickList(True, False, cboBaseTable.ItemData(cboBaseTable.ListIndex)) Then
              .Show vbModal
            End If
            frmDefSel.SelectedID = .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          End With

        Case edtEdit
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(False, frmDefSel.FromCopy, cboBaseTable.ItemData(cboBaseTable.ListIndex), frmDefSel.SelectedID) Then
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
          Me.Changed = True

          txtPicklist = frmDefSel.SelectedText
          txtPicklist.Tag = frmDefSel.SelectedID
          txtFilter.Text = ""
          txtFilter.Tag = 0
          fExit = True
      
        Case 0
          If IsPicklistValid(txtPicklist.Tag) <> vbNullString Then
            txtPicklist.Text = "<None>"
            txtPicklist.Tag = 0
          End If
          fExit = True
      
        End Select
      End If

    Loop

  End With

  Set frmDefSel = Nothing

  ForceDefinitionToBeHiddenIfNeeded

Exit Sub

LocalErr:
  ErrorMsgbox "Error selecting picklist"

End Sub

Private Sub cmdFilter_Click()
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression

  Dim rsTemp As Recordset
  Dim sSQL As String
  Dim blnHiddenPicklist As Boolean
  
  On Error GoTo LocalErr
  
  'If cboBaseTable.Text = "<None>" Then
  '  MsgBox "No primary table specified", vbExclamation, Me.Caption
  '  Exit Sub
  'End If

  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression

  With objExpression
    ' Initialise the expression object.
    fOK = .Initialise(cboBaseTable.ItemData(cboBaseTable.ListIndex), Val(txtFilter.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)

    If fOK Then
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) Then
  
        Me.Changed = True

        ' Read the selected expression info.
        txtFilter.Text = .Name
        txtFilter.Tag = .ExpressionID
        txtPicklist.Text = ""
        txtPicklist.Tag = 0

      End If

    End If
  End With

  Set objExpression = Nothing

  ForceDefinitionToBeHiddenIfNeeded

Exit Sub

LocalErr:
  ErrorMsgbox "Error selecting filter"

End Sub


Private Sub cmdOK_Click()
  
  If ValidateDefinition = False Then
    Exit Sub
  End If
  Call SaveDefinition

  Me.Hide
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdCancel_Click()
  
  Dim strSQL As String
  
  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  If Me.Changed And Not mblnReadOnly Then

    'strMBText = "Cross Tab definition has changed.  Save changes ?"
    strMBText = "You have changed the current definition. Save changes ?"
    intMBButtons = vbQuestion + vbYesNoCancel + vbDefaultButton1
    strMBTitle = "Cross Tab"
    intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)
    
    Select Case intMBResponse
    Case vbYes
      Call cmdOK_Click
      Exit Sub
    Case vbCancel
      Exit Sub
    End Select
  End If

  Me.Hide
  Screen.MousePointer = vbDefault

End Sub


Private Function GetItemName(bTable As Boolean, lItemID As Long) As String

    Dim sSQL As String
    Dim rsItem As Recordset

    If bTable Then
      sSQL = "Select TableName From ASRSysTables Where TableID = " & lItemID
    Else
      sSQL = "Select ColumnName From ASRSysColumns Where ColumnID = " & lItemID
    End If

    Set rsItem = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    GetItemName = rsItem(0)

    rsItem.Close
    Set rsItem = Nothing

End Function


'=================================================================================
'
' Possible re-usable code above this point
'
'=================================================================================

Private Sub SSTab1_Click(PreviousTab As Integer)

  Dim ctl As Control

  If Not mblnReadOnly Then
    For Each ctl In Me.Controls
      If TypeOf ctl Is VB.Frame Then
        ctl.Enabled = ctl.Left >= 0
      End If
    Next
  End If

  If SSTab1.Tab = 1 Then
    CheckIfEnableRangeLabels
  End If

End Sub

Private Sub RetreiveDefinition()

  Dim rsTemp As Recordset
  'Dim strSQL As String
  Dim blnHiddenPicklistOrFilter As Boolean
  Dim strRecSelStatus As String

  On Error GoTo LocalErr

  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Cross Tab"
    fOK = False
    Exit Sub
  End If
  
  SetComboText cboBaseTable, GetItemName(True, rsTemp!TableID)
  If cboBaseTable.ItemData(cboBaseTable.ListIndex) <> rsTemp!TableID Then
    MsgBox "This definition contains an invalid base table and could not be loaded", vbExclamation, "Cross Tab"
    fOK = False
    Exit Sub
  End If
  Let mstrPrimaryTable = cboBaseTable.Text

  Call LoadPrimaryDependantCombos
    
  'Horizontal column
  SetComboText cboHorizontalCol, GetItemName(False, rsTemp!HorizontalColID)
  Call RangeEnabled(1, rsTemp!HorizontalColID)
  If rsTemp!HorizontalStart <> 0 Or rsTemp!HorizontalStop <> 0 Then
    'mskHorizontalRange(0) = rsTemp!HorizontalStart
    'mskHorizontalRange(1) = rsTemp!HorizontalStop
    'mskHorizontalRange(2) = rsTemp!HorizontalStep
    mskHorizontalRange(0).Value = datGeneral.ConvertNumberForDisplay(rsTemp!HorizontalStart)
    mskHorizontalRange(1).Value = datGeneral.ConvertNumberForDisplay(rsTemp!HorizontalStop)
    mskHorizontalRange(2).Value = datGeneral.ConvertNumberForDisplay(rsTemp!HorizontalStep)
  End If

  SetComboText cboVerticalCol, GetItemName(False, rsTemp!VerticalColID)
  Call RangeEnabled(2, rsTemp!VerticalColID)
  If rsTemp!VerticalStart <> 0 Or rsTemp!VerticalStop <> 0 Then
    'mskVerticalRange(0) = rsTemp!VerticalStart
    'mskVerticalRange(1) = rsTemp!VerticalStop
    'mskVerticalRange(2) = rsTemp!VerticalStep
    mskVerticalRange(0).Value = datGeneral.ConvertNumberForDisplay(rsTemp!VerticalStart)
    mskVerticalRange(1).Value = datGeneral.ConvertNumberForDisplay(rsTemp!VerticalStop)
    mskVerticalRange(2).Value = datGeneral.ConvertNumberForDisplay(rsTemp!VerticalStep)
  End If

  If rsTemp!PageBreakColID > 0 Then
    SetComboText cboPageBreakCol, GetItemName(False, rsTemp!PageBreakColID)
    Call RangeEnabled(3, rsTemp!PageBreakColID)
    If rsTemp!PageBreakStart <> 0 Or rsTemp!PageBreakStop <> 0 Then
      'mskPageBreakRange(0) = rsTemp!PageBreakStart
      'mskPageBreakRange(1) = rsTemp!PageBreakStop
      'mskPageBreakRange(2) = rsTemp!PageBreakStep
      mskPageBreakRange(0).Value = datGeneral.ConvertNumberForDisplay(rsTemp!PageBreakStart)
      mskPageBreakRange(1).Value = datGeneral.ConvertNumberForDisplay(rsTemp!PageBreakStop)
      mskPageBreakRange(2).Value = datGeneral.ConvertNumberForDisplay(rsTemp!PageBreakStep)
    End If
  End If

  If rsTemp!IntersectionColID > 0 Then
    SetComboText cboIntersectionCol, GetItemName(False, rsTemp!IntersectionColID)
  End If
    
  SetComboItem cboType, rsTemp!IntersectionType

  chkPercentage = Abs(rsTemp!Percentage)
  chkPercentageofPage = Abs(rsTemp!PercentageofPage)
  chkSuppressZeros = Abs(rsTemp!SuppressZeros)
  
  If Not IsNull(rsTemp!ThousandSeparators) Then
    chkThousandSeparators.Value = Abs(rsTemp!ThousandSeparators)
  End If
'  Select Case CLng(rsTemp!DefaultOutput)
'  Case 0  'Printer
'    optOutput(0) = True
'  Case 1  'Export
'    optOutput(1) = True
'    SetComboItem cboExportTo, rsTemp!DefaultExportTo
'    chkSave = Abs(rsTemp!DefaultSave)
'    txtFilename = rsTemp!DefaultSaveAs
'    chkCloseApplication = Abs(rsTemp!DefaultCloseApp)
'    Call chkSave_Click
'  End Select
  
  mobjOutputDef.ReadDefFromRecset rsTemp
  
  
  ' === Standard access stuff ===
  
  txtDesc.Text = IIf(rsTemp!Description <> vbNullString, rsTemp!Description, vbNullString)
  
  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If
    
  blnHiddenPicklistOrFilter = False
  mblnRecordSelectionInvalid = False
  
  If rsTemp!PicklistID > 0 Then
    optPicklist = True
    txtPicklist.Tag = rsTemp!PicklistID
    txtPicklist.Text = rsTemp!PicklistName
            
    ''NHRD09042002 Fault 3322 - Code Added
    chkPrintFilterHeader.Value = Abs(rsTemp!PrintFilterHeader)
    
    cmdPicklist.Enabled = (Not mblnReadOnly)
    cmdFilter.Enabled = False
    
  ElseIf rsTemp!FilterID > 0 Then
    optFilter = True
    txtFilter.Tag = rsTemp!FilterID
    txtFilter.Text = rsTemp!FilterName
    
    'NHRD09042002 Fault 3322 - Code Added
    chkPrintFilterHeader.Value = Abs(rsTemp!PrintFilterHeader)
    
    cmdPicklist.Enabled = False
    cmdFilter.Enabled = (Not mblnReadOnly)
    
  Else
    optAllRecords = True
    'NHRD09042002 Fault 3322 - Code Added
    chkPrintFilterHeader.Value = Abs(rsTemp!PrintFilterHeader)
    
    cmdPicklist.Enabled = False
    cmdFilter.Enabled = False
  
  End If

  mblnReadOnly = Not datGeneral.SystemPermission("CROSSTABS", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlCrossTab, mlngCrossTabID) = ACCESS_READONLY)
  End If
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  ' =============================

  mlngTimeStamp = rsTemp!intTimestamp

  rsTemp.Close
  Set rsTemp = Nothing

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    fOK = False
    Exit Sub
  End If

Exit Sub

LocalErr:
  ErrorMsgbox "Error retrieving Cross Tab definition"

End Sub

Private Sub SaveDefinition()

  Dim rsTemp As Recordset
  Dim strSQL As String
  
  Dim strName As String
  Dim strDesc As String
  Dim strTableID As String
  Dim strSelection As Integer
  Dim strPicklist As String
  Dim strFilter As String
  Dim strUserName As String
  
  Dim strHorizontalColID As String
  Dim strHorizontalStart As String
  Dim strHorizontalStop As String
  Dim strHorizontalStep As String
  
  Dim strVerticalColID As String
  Dim strVerticalStart As String
  Dim strVerticalStop As String
  Dim strVerticalStep As String
  
  Dim strPageBreakColID As String
  Dim strPageBreakStart As String
  Dim strPageBreakStop As String
  Dim strPageBreakStep As String
  
  Dim strType As String
  Dim strIntersectionColID As String
  Dim strPercentage As String
  Dim strPercentageofPage As String
  Dim strSuppressZeros As String
  Dim strThousandSeparators As String
  
  'Dim strDefaultOutput As String
  'Dim strDefaultExportTo As String
  'Dim strDefaultSave As String
  'Dim strDefaultSaveAs As String
  'Dim strDefaultCloseApp As String
  
  Dim strPrintFilterHeader As String
    
  On Error GoTo LocalErr

  Screen.MousePointer = vbHourglass

  strName = "'" & Replace(Trim(txtName.Text), "'", "''") & "'"
  strDesc = "'" & Replace(txtDesc.Text, "'", "''") & "'"
  strTableID = CStr(cboBaseTable.ItemData(cboBaseTable.ListIndex))
  strUserName = "'" & datGeneral.UserNameForSQL & "'"
    
  strSelection = "0"
  strPicklist = "0"
  strFilter = "0"
  If optPicklist = True Then
    strSelection = "1"
    strPicklist = txtPicklist.Tag
  ElseIf optFilter = True Then
    strSelection = "2"
    strFilter = txtFilter.Tag
  End If

  'If optOutput(1).Value = True Then   'Export To
  '  strDefaultOutput = "1"
  '  strDefaultExportTo = CStr(cboExportTo.ItemData(cboExportTo.ListIndex))
  '  strDefaultSave = CStr(Abs(chkSave <> 0))
  '  strDefaultSaveAs = "'" & Replace(txtFilename.Text, "'", "''") & "'"
  '  strDefaultCloseApp = CStr(Abs(chkCloseApplication <> 0))
  'Else
  '  strDefaultOutput = "0"
  '  strDefaultExportTo = "0"
  '  strDefaultSave = "0"
  '  strDefaultSaveAs = "''"
  '  strDefaultCloseApp = "0"
  'End If

  strHorizontalColID = CStr(cboHorizontalCol.Tag)
  strHorizontalStart = "'" & datGeneral.ConvertNumberForSQL(mskHorizontalRange(0).Value) & "'"
  strHorizontalStop = "'" & datGeneral.ConvertNumberForSQL(mskHorizontalRange(1).Value) & "'"
  strHorizontalStep = "'" & datGeneral.ConvertNumberForSQL(mskHorizontalRange(2).Value) & "'"
  
  strVerticalColID = CStr(cboVerticalCol.Tag)
  strVerticalStart = "'" & datGeneral.ConvertNumberForSQL(mskVerticalRange(0).Value) & "'"
  strVerticalStop = "'" & datGeneral.ConvertNumberForSQL(mskVerticalRange(1).Value) & "'"
  strVerticalStep = "'" & datGeneral.ConvertNumberForSQL(mskVerticalRange(2).Value) & "'"
  
  strPageBreakColID = CStr(cboPageBreakCol.Tag)
  strPageBreakStart = "'" & datGeneral.ConvertNumberForSQL(mskPageBreakRange(0).Value) & "'"
  strPageBreakStop = "'" & datGeneral.ConvertNumberForSQL(mskPageBreakRange(1).Value) & "'"
  strPageBreakStep = "'" & datGeneral.ConvertNumberForSQL(mskPageBreakRange(2).Value) & "'"

  strType = CStr(cboType.ItemData(cboType.ListIndex))
  strIntersectionColID = CStr(cboIntersectionCol.ItemData(cboIntersectionCol.ListIndex))
  strPercentage = CStr(Abs(chkPercentage <> 0))
  strPercentageofPage = CStr(Abs(chkPercentageofPage <> 0))
  strSuppressZeros = CStr(Abs(chkSuppressZeros <> 0))
  strThousandSeparators = CStr(Abs(chkThousandSeparators.Value <> vbUnchecked))
    
  strPrintFilterHeader = CStr(Abs(chkPrintFilterHeader <> 0))
    
  If mlngCrossTabID > 0 Then
    strSQL = "UPDATE ASRSysCrossTab SET " & _
               "Name = " & strName & ", " & _
               "Description = " & strDesc & ", " & _
               "TableID = " & strTableID & ", " & _
               "Selection = " & strSelection & ", " & _
               "PicklistID = " & strPicklist & ", " & _
               "FilterID = " & strFilter & ", "

    strSQL = strSQL & _
               "HorizontalColID = " & strHorizontalColID & ", " & _
               "HorizontalStart = " & strHorizontalStart & ", " & _
               "HorizontalStop = " & strHorizontalStop & ", " & _
               "HorizontalStep = " & strHorizontalStep & ", " & _
               "VerticalColID = " & strVerticalColID & ", " & _
               "VerticalStart = " & strVerticalStart & ", " & _
               "VerticalStop = " & strVerticalStop & ", " & _
               "VerticalStep = " & strVerticalStep & ", " & _
               "PageBreakColID = " & strPageBreakColID & ", " & _
               "PageBreakStart = " & strPageBreakStart & ", " & _
               "PageBreakStop = " & strPageBreakStop & ", " & _
               "PageBreakStep = " & strPageBreakStep & ", "

    strSQL = strSQL & _
               "IntersectionType = " & strType & ", " & _
               "IntersectionColID = " & strIntersectionColID & ", " & _
               "Percentage = " & strPercentage & ", " & _
               "PercentageofPage = " & strPercentageofPage & ", " & _
               "SuppressZeros = " & strSuppressZeros & ", " & _
               "ThousandSeparators = " & strThousandSeparators & ", " & _
               "PrintFilterHeader = " & strPrintFilterHeader & ","
               '& _
               "DefaultOutput = " & strDefaultOutput & ", " & _
               "DefaultExportTo = " & strDefaultExportTo & ", " & _
               "DefaultSave = " & strDefaultSave & ", " & _
               "DefaultSaveAs = " & strDefaultSaveAs & ", " & _
               "DefaultCloseApp = " & strDefaultCloseApp & _

    strSQL = strSQL & _
        "OutputPreview = " & IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        "OutputFormat = " & CStr(mobjOutputDef.GetSelectedFormatIndex) & ", " & _
        "OutputScreen = " & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", " & _
        "OutputPrinter = " & IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0") & ", " & _
        "OutputPrinterName = '" & Replace(cboPrinterName.Text, " '", "''") & "', "
        
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
    
    strSQL = strSQL & _
        "OutputFilename = '" & Replace(txtFilename.Text, "'", "''") & "'"
               

    strSQL = strSQL & _
        " WHERE CrossTabID = " & CStr(mlngCrossTabID)


  'NHRD09042002 Fault 3322 - Code Addedtype
  'added strPrintFilterHeader variable in a few places here
    gADOCon.Execute strSQL, , adCmdText
  
    Call UtilUpdateLastSaved(utlCrossTab, mlngCrossTabID)
  
  Else
    strSQL = "INSERT ASRSysCrossTab (" & _
               "Name, Description, TableID, " & _
               "Selection, PicklistID, FilterID, " & _
               "HorizontalColID, HorizontalStart, HorizontalStop, HorizontalStep, " & _
               "VerticalColID, VerticalStart, VerticalStop, VerticalStep, " & _
               "PageBreakColID, PageBreakStart, PageBreakStop, PageBreakStep, " & _
               "IntersectionType, IntersectionColID, Percentage, PercentageofPage, SuppressZeros, ThousandSeparators, " & _
               "PrintFilterHeader, " & _
               "UserName, " & _
               "OutputPreview, OutputFormat, OutputScreen, OutputPrinter, OutputPrinterName, OutputSave, " & _
               "OutputSaveExisting, OutputEmail, OutputEmailAddr, OutputEmailSubject, OutputEmailAttachAs, OutputFilename) "
               '"DefaultOutput, DefaultExportTo, DefaultSave, DefaultSaveAs, DefaultCloseApp) "
               
    strSQL = strSQL & _
               "VALUES( " & _
               strName & ", " & strDesc & ", " & strTableID & ", " & _
               strSelection & ", " & strPicklist & ", " & strFilter & ", " & _
               strHorizontalColID & ", " & strHorizontalStart & ", " & strHorizontalStop & ", " & strHorizontalStep & ", " & _
               strVerticalColID & ", " & strVerticalStart & ", " & strVerticalStop & ", " & strVerticalStep & ", " & _
               strPageBreakColID & ", " & strPageBreakStart & ", " & strPageBreakStop & ", " & strPageBreakStep & ", " & _
               strType & ", " & strIntersectionColID & ", " & strPercentage & ", " & strPercentageofPage & ", " & strSuppressZeros & ", " & strThousandSeparators & ", " & _
               strPrintFilterHeader & ", " & _
               strUserName & ", " '& _
               strDefaultOutput & ", " & strDefaultExportTo & ", " & strDefaultSave & ", " & strDefaultSaveAs & ", " & strDefaultCloseApp & ")"
  
    strSQL = strSQL & _
        IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        CStr(mobjOutputDef.GetSelectedFormatIndex) & ", " & _
        IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", " & _
        IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0") & ", " & _
        "'" & Replace(cboPrinterName.Text, "'", "''") & "', "

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
        "'" & Replace(txtFilename.Text, "'", "''") & "')"

  
  
    
    ' RH 04/09/00 - Use the new util def stored procedure
    mlngCrossTabID = InsertCrossTab(strSQL)
    
    'gADOCon.Execute strSQL, , adCmdText
  
    'strSQL = "SELECT MAX(CrossTabID) FROM ASRSysCrossTab"
    'Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
    'mlngCrossTabID = Val(rsTemp(0))
  
    Call UtilCreated(utlCrossTab, mlngCrossTabID)
  
  End If
  
  SaveAccess

  Screen.MousePointer = vbDefault

Exit Sub

LocalErr:
  ErrorMsgbox "Error saving Cross Tab definition"

End Sub

Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysCrossTabAccess WHERE ID = " & mlngCrossTabID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysCrossTabAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngCrossTabID & ", sysusers.name," & _
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
      sSQL = "IF EXISTS (SELECT * FROM ASRSysCrossTabAccess" & _
        " WHERE ID = " & CStr(mlngCrossTabID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysCrossTabAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngCrossTabID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub





Private Function InsertCrossTab(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertCrossTab_ERROR

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
    pmADO.Value = "AsrSysCrossTab"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "CrossTabID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertCrossTab = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertCrossTab = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertCrossTab_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function



Private Function ValidateDefinition() As Boolean

  'Check that all required information has been completed before attempting to save
  
  Dim strRecSelStatus As String
  Dim blnContinueSave As Boolean
  Dim blnSaveAsNew As Boolean
  Dim strName As String
  
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
  
  On Error GoTo LocalErr
  
  ValidateDefinition = False
  strName = Trim(txtName.Text)

  If Len(strName) = 0 Then
    SSTab1.Tab = 0
    MsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If
  
  If optFilter Then
    If Val(txtFilter.Tag) = 0 Then
      SSTab1.Tab = 0
      MsgBox "No Filter entered for the base table.", vbExclamation
      cmdFilter.SetFocus
      Exit Function
    End If
  End If
    
  If optPicklist Then
    If Val(txtPicklist.Tag) = 0 Then
      SSTab1.Tab = 0
      MsgBox "No Picklist entered for the base table.", vbExclamation
      cmdPicklist.SetFocus
      Exit Function
    End If
  End If
    
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
    

  'If Val(mskHorizontalRange(0)) > 0 Then
  If Val(mskHorizontalRange(0).Value) <> 0 Or Val(mskHorizontalRange(1).Value) <> 0 Then
    If Val(mskHorizontalRange(1).Value) <= Val(mskHorizontalRange(0).Value) Then
      SSTab1.Tab = 1
      MsgBox "Horizontal stop value must be greater than Horizontal start value", vbExclamation
      mskHorizontalRange(1).SetFocus
      Exit Function
    End If
    If Val(mskHorizontalRange(2).Value) <= 0 Then
      SSTab1.Tab = 1
      MsgBox "Horizontal increment must be greater than zero", vbExclamation
      mskHorizontalRange(2).SetFocus
      Exit Function
    End If
  
    ' RH 05/10/00 - Ensure steps does not exceed maximum imposed by combo control
    If (Val(mskHorizontalRange(1).Value) - Val(mskHorizontalRange(0).Value)) / Val(mskHorizontalRange(2).Value) > 32768 Then
      MsgBox "Maximum number of steps between start, stop and increment value for the Horizontal Range " & _
             "has been exceeded. You must either increase the increment value or decrease the stop value.", vbExclamation
      SSTab1.Tab = 1
      Exit Function
    End If
  
  End If
    
  'If Val(mskVerticalRange(0)) > 0 Then
  If Val(mskVerticalRange(0).Value) <> 0 Or Val(mskVerticalRange(1).Value) <> 0 Then
    If Val(mskVerticalRange(1).Value) <= Val(mskVerticalRange(0).Value) Then
      SSTab1.Tab = 1
      MsgBox "Vertical stop value must be greater than Vertical start value", vbExclamation
      mskVerticalRange(1).SetFocus
      Exit Function
    End If
    If Val(mskVerticalRange(2).Value) <= 0 Then
      SSTab1.Tab = 1
      MsgBox "Vertical increment must be greater than zero", vbExclamation
      mskVerticalRange(2).SetFocus
      Exit Function
    End If
  
    ' RH 05/10/00 - Ensure steps does not exceed maximum imposed by combo control
    If (Val(mskVerticalRange(1).Value) - Val(mskVerticalRange(0).Value)) / Val(mskVerticalRange(2).Value) > 32768 Then
      MsgBox "Maximum number of steps between start, stop and increment value for the Vertical Range " & _
             "has been exceeded. You must either increase the increment value or decrease the stop value.", vbExclamation
      SSTab1.Tab = 1
      Exit Function
    End If
  
  End If
    
  'If Val(mskPageBreakRange(0).Value) > 0 Then
  If Val(mskPageBreakRange(0).Value) <> 0 Or Val(mskPageBreakRange(1).Value) <> 0 Then
    If Val(mskPageBreakRange(1).Value) <= Val(mskPageBreakRange(0).Value) Then
      SSTab1.Tab = 1
      MsgBox "Page Break stop value must be greater than Page Break start value", vbExclamation
      mskPageBreakRange(1).SetFocus
      Exit Function
    End If
    If Val(mskPageBreakRange(2).Value) <= 0 Then
      SSTab1.Tab = 1
      MsgBox "Page Break increment must be greater than zero", vbExclamation
      mskPageBreakRange(2).SetFocus
      Exit Function
    End If
        
    ' RH 05/10/00 - Ensure steps does not exceed maximum imposed by combo control
    If (Val(mskPageBreakRange(1).Value) - Val(mskPageBreakRange(0).Value)) / Val(mskPageBreakRange(2).Value) > 32768 Then
      MsgBox "Maximum number of steps between start, stop and increment value for the Page Break Range " & _
             "has been exceeded. You must either increase the increment value or decrease the stop value.", vbExclamation
      SSTab1.Tab = 1
      Exit Function
    End If
  
  End If
    
  
  'If optOutput(1).Value Then
  '  If chkSave.Value And txtFilename = "" Then
  '    SSTab1.Tab = 2
  '    MsgBox "You must select a filename if you opt to save the document !", vbExclamation
  '    cmdFilename.SetFocus
  '    Exit Function
  '  End If
  'End If
  If Not mobjOutputDef.ValidDestination Then
    SSTab1.Tab = 2
    Exit Function
  End If


  'Check if this definition has been changed by another user
  Call UtilityAmended(utlCrossTab, mlngCrossTabID, mlngTimeStamp, blnContinueSave, blnSaveAsNew)
  If blnContinueSave = False Then
    Exit Function
  ElseIf blnSaveAsNew Then
    txtUserName = gsUserName
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngCrossTabID = 0
    mblnReadOnly = False
    ForceAccess
  End If


  If ValidateDefinitionUniqueName(strName) = False Then
    SSTab1.Tab = 0
    MsgBox "A Cross Tab definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If
  
If mlngCrossTabID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlCrossTab, _
      CStr(mlngCrossTabID), _
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
        MsgBox "This definition cannot be made hidden from the following user groups :" & vbCrLf & vbCrLf & sBatchJobScheduledUserGroups & vbCrLf & _
               "as it is used in the following batch jobs which are scheduled to be run by these user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, _
               vbExclamation + vbOKOnly, "Cross Tabs"
      Else
        MsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Cross Tabs"
      End If

      Screen.MousePointer = vbNormal
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If MsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Cross Tabs") = vbNo Then
        Screen.MousePointer = vbNormal
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

LocalErr:
  ErrorMsgbox "Error validating Cross Tab definition"

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





Private Function ValidateDefinitionUniqueName(sName As String) As Boolean

  Dim rsName As Recordset
  Dim sSQL As String
    
  sSQL = "SELECT * FROM ASRSysCrossTab " & _
         "WHERE Name = '" & Replace(sName, "'", "''") & "' AND CrossTabID <> " & mlngCrossTabID
    
  Set rsName = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  ValidateDefinitionUniqueName = (rsName.BOF And rsName.EOF)
  rsName.Close
    
  Set rsName = Nothing

End Function


Private Sub LoadPrimaryCombo()

  Dim sSQL As String
  Dim rsTables As New Recordset

  'sSQL = "Select TableName, TableID From ASRSysTables " & _
         "JOIN ASRSYSColumns ON ASRSYSColumns.TableID = ASRSysColumns.TableID " & _
         "WHERE TableType='1' OR TableType='2' AND "
  
  'Only retreive tables which have more than one suitable column
  'sSQL = "SELECT ASRSysTables.TableID, ASRSysTables.TableName " & _
         " FROM ASRSysTables" & _
         " WHERE (tableType = " & Trim(Str(tabTopLevel)) & _
         " OR tableType = " & Trim(Str(tabChild)) & ")" & _
         " AND (SELECT COUNT(*) FROM ASRSysColumns" & _
         " WHERE ASRSysColumns.TableID = ASRSysTables.TableID" & _
         " AND columnType <> " & Trim(Str(colSystem)) & _
         " AND columnType <> " & Trim(Str(colLink)) & _
         " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
         " AND dataType <> " & Trim(Str(sqlOle)) & ") > 1"
  sSQL = "SELECT ASRSysTables.TableID, ASRSysTables.TableName " & _
         " FROM ASRSysTables" & _
         " WHERE (SELECT COUNT(*) FROM ASRSysColumns" & _
         " WHERE ASRSysColumns.TableID = ASRSysTables.TableID" & _
         " AND columnType <> " & Trim(Str(colSystem)) & _
         " AND columnType <> " & Trim(Str(colLink)) & _
         " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
         " AND dataType <> " & Trim(Str(sqlOle)) & ") > 1"
  
  LoadTableCombo cboBaseTable, sSQL
'  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'  With cboBaseTable
'    .Clear
'    Do While Not rsTables.EOF
'      .AddItem rsTables!TableName
'      .ItemData(.NewIndex) = rsTables!TableID
'      rsTables.MoveNext
'    Loop
'  End With
'
'  rsTables.Close
'  Set rsTables = Nothing
  
End Sub

Public Sub LoadPrimaryDependantCombos()

  Dim sSQL As String
  Dim lngTableID As Long
  Dim fOriginalLoading As Boolean
  
  lngTableID = 0
  If cboBaseTable.ListIndex <> -1 Then
    lngTableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)
  End If

  fOriginalLoading = mblnLoading
  mblnLoading = True

  sSQL = "SELECT columnName, columnID, columnType, dataType, size, decimals" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & CStr(lngTableID) & _
    " AND columnType <> " & Trim(Str(colSystem)) & _
    " AND columnType <> " & Trim(Str(colLink)) & _
    " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
    " AND dataType <> " & Trim(Str(sqlOle))

  'MH20000904 Fault 837
  'After going into a filter and clicking the test button, a 'zombie' error
  'would sometimes be produced.  Making the recordset persistent seems to fix it!
  
  'Set mrsColumns = datData.OpenRecordset(sSQL, adOpenDynamic, adLockReadOnly)
  Set mrsColumns = datData.OpenPersistentRecordset(sSQL, adOpenDynamic, adLockReadOnly)

  LoadColCombo cboHorizontalCol, False, False
  If cboHorizontalCol.ListCount > 0 Then
    cboHorizontalCol.ListIndex = 0
  End If

  'LoadColCombo cboVerticalCol, False, False
  'LoadColCombo cboPageBreakCol, True, False
  LoadColCombo cboIntersectionCol, True, True
  If cboIntersectionCol.ListCount > 0 Then
    cboIntersectionCol.ListIndex = 0
  End If

  mblnLoading = fOriginalLoading

End Sub

Private Sub LoadColCombo(cboOutput As ComboBox, blnAllowNone As Boolean, blnOnlyNumeric, _
                  Optional lngDisAllowHor As Long = 0, Optional lngDisAllowVer As Long = 0)

  With cboOutput
    .Visible = False
    .Clear
    
    If blnAllowNone Then
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End If
    
    If Not mrsColumns.BOF Or Not mrsColumns.EOF Then
      mrsColumns.MoveFirst
      Do While Not mrsColumns.EOF
  
        If mrsColumns!ColumnID <> lngDisAllowHor And mrsColumns!ColumnID <> lngDisAllowVer Then
        
          If Not blnOnlyNumeric Or _
            (mrsColumns!DataType = sqlNumeric Or mrsColumns!DataType = sqlInteger) Then
              .AddItem mrsColumns!ColumnName
              .ItemData(.NewIndex) = mrsColumns!ColumnID
          End If
        End If
  
        mrsColumns.MoveNext
      Loop
    End If

    '.Enabled = (.ListCount > 1)
    '.BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
    .Visible = True
  End With

End Sub

Private Sub LoadOtherCombos()

  Dim intCount As Integer
  
  With cboType
    .Clear
    .AddItem "Count": .ItemData(.NewIndex) = 0
    .AddItem "Average": .ItemData(.NewIndex) = 1
    .AddItem "Maximum": .ItemData(.NewIndex) = 2
    .AddItem "Minimum": .ItemData(.NewIndex) = 3
    .AddItem "Total": .ItemData(.NewIndex) = 4
    .ListIndex = 0
  End With



  'With cboExportTo
  '  .AddItem "Html": .ItemData(.NewIndex) = 0
  '  .AddItem "Microsoft Excel": .ItemData(.NewIndex) = 1
  '  .AddItem "Microsoft Word": .ItemData(.NewIndex) = 2
  '  .ListIndex = 0
  'End With

End Sub

Private Sub cboIntersectionCol_Click()

  Dim blnEnabled As Boolean
  'If mblnLoading Then
  '  Exit Sub
  'End If

  blnEnabled = (cboIntersectionCol.ListIndex <> 0)

  With cboType
    .Enabled = blnEnabled
    .BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    If blnEnabled = False Then SetComboItem cboType, 0
  End With

  Me.Changed = True

End Sub

Private Sub RangeEnabled(Index As Integer, ColID As Long)
  Dim txtTemp As TDBNumberCtrl.TDBNumber
  'Dim txtTemp As TDBNumber
  Dim intCount As Integer
  Dim blnNumericColumn As Boolean
  Dim strMask As String
  Dim dblMaxValue As Double
  Dim intSize As Integer
  Dim lngDigitsBeforeDecimal As Long
  
  strMask = "#"
  dblMaxValue = 0
  blnNumericColumn = False
  
  If ColID > 0 Then

    If Not mrsColumns.BOF Or Not mrsColumns.EOF Then
      With mrsColumns
        .MoveFirst
        .Find "ColumnID = " & CStr(ColID)
      End With
    End If

    If Not mrsColumns.BOF Or Not mrsColumns.EOF Then

      Select Case mrsColumns!DataType
      Case sqlNumeric
        lngDigitsBeforeDecimal = mrsColumns!Size - mrsColumns!Decimals
        strMask = String$(lngDigitsBeforeDecimal - 1, "#") & "0"
        dblMaxValue = Val(String$(lngDigitsBeforeDecimal, "9"))
        If mrsColumns!Decimals > 0 Then
          'strMask = strMask & UI.GetSystemDecimalSeparator & String$(mrsColumns!Decimals, "0")
          strMask = strMask & "." & String$(mrsColumns!Decimals, "0")
          dblMaxValue = dblMaxValue + Val("." & String$(mrsColumns!Decimals, "9"))
        End If
        blnNumericColumn = True
      
      Case sqlInteger
        strMask = String$(9, "#") & "0"
        dblMaxValue = Val(String$(10, "9"))
        blnNumericColumn = True
      
      End Select
  
    End If
  End If
  
  
  For intCount = 0 To 2

    Set txtTemp = Choose(Index, mskHorizontalRange(intCount), mskVerticalRange(intCount), mskPageBreakRange(intCount))

    With txtTemp
      .Visible = False
      .Enabled = blnNumericColumn
      .BackColor = IIf(blnNumericColumn, vbWindowBackground, vbButtonFace)
      .Text = vbNullString
      .Format = strMask
      .DisplayFormat = strMask
      .MaxValue = dblMaxValue
      .MinValue = .MaxValue * -1
      .Visible = True
    End With

    Set txtTemp = Nothing

  Next

  CheckIfEnableRangeLabels

End Sub

Private Sub cboHorizontalCol_Click()

  With cboHorizontalCol
  
    'Check if it has changed
    If .ListIndex <> -1 Then
      If Val(.Tag) <> .ItemData(.ListIndex) Then
        .Tag = .ItemData(.ListIndex)
        Call RangeEnabled(1, Val(.Tag))
      End If
      
      With cboVerticalCol

        Call LoadColCombo(cboVerticalCol, False, False, Val(cboHorizontalCol.Tag))
        SetComboItem cboVerticalCol, Val(.Tag)

        If .ListCount > 0 And .ListIndex = -1 Then
          .ListIndex = 0
        Else
          Call cboVerticalCol_Click
        End If

      End With
      
    End If

  End With

  Me.Changed = True

End Sub

Private Sub cboVerticalCol_Click()

  Dim lngHorColID As Long
  Dim lngVerColID As Long
  Dim lngPgbColID As Long
  
  With cboVerticalCol
  
    If .ListIndex <> -1 Then
      If Val(.Tag) <> .ItemData(.ListIndex) Then
        .Tag = .ItemData(.ListIndex)
        Call RangeEnabled(2, Val(.Tag))
      End If

      With cboPageBreakCol

        Call LoadColCombo(cboPageBreakCol, True, False, Val(cboHorizontalCol.Tag), Val(cboVerticalCol.Tag))
        SetComboItem cboPageBreakCol, Val(.Tag)

        If .ListCount > 0 And .ListIndex = -1 Then
          .ListIndex = 0
        Else
          Call cboPageBreakCol_Click
        End If

      End With

    End If

  End With

  Me.Changed = True

End Sub

Private Sub cboPageBreakCol_Click()
  'Enable start, stop, increment text boxes if numeric column selected
  
  With cboPageBreakCol
  
    'Check if it has changed
    If .ListIndex <> -1 Then
      If Val(.Tag) <> .ItemData(.ListIndex) Or .Tag = vbNullString Then
        .Tag = .ItemData(.ListIndex)
        Call RangeEnabled(3, Val(.Tag))
      End If
    End If
  
    Call chkPercentage_Click

  End With
  
  Me.Changed = True

End Sub

Private Sub mskHorizontalRange_GotFocus(Index As Integer)
  txt_GotFocus mskHorizontalRange(Index)
End Sub
Private Sub mskHorizontalRange_Change(Index As Integer)
  Me.Changed = True
End Sub
Private Sub mskVerticalRange_GotFocus(Index As Integer)
  txt_GotFocus mskVerticalRange(Index)
End Sub
Private Sub mskVerticalRange_Change(Index As Integer)
  Me.Changed = True
End Sub
Private Sub mskPageBreakRange_GotFocus(Index As Integer)
  txt_GotFocus mskPageBreakRange(Index)
End Sub
Private Sub mskPageBreakRange_Change(Index As Integer)
  Me.Changed = True
End Sub


Private Function GetDefinition() As Recordset

  Dim strSQL As String

  strSQL = "SELECT ASRSysCrossTab.*, " & _
           "CONVERT(integer,ASRSysCrossTab.TimeStamp) AS intTimeStamp, " & _
           "ASRSysPickListName.Name AS PickListName, " & _
           "ASRSysPickListName.Access AS PickListAccess, " & _
           "ASRSysExpressions.Name AS FilterName, " & _
           "ASRSysExpressions.Access AS FilterAccess " & _
           "FROM ASRSysCrossTab " & _
           "LEFT OUTER JOIN ASRSysExpressions ON ASRSysCrossTab.FilterID = ASRSysExpressions.ExprID " & _
           "LEFT OUTER JOIN ASRSysPickListName ON ASRSysCrossTab.PickListID = ASRSysPickListName.PickListID " & _
           "WHERE ASRSysCrossTab.CrossTabID = " & CStr(mlngCrossTabID)
  Set GetDefinition = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

End Function


Public Sub PrintDef(lCrossTabID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim iLoop As Integer
  Dim varBookmark As Variant

  Set datData = New HRProDataMgr.clsDataAccess
  
  mlngCrossTabID = lCrossTabID
  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Cross Tab"
    Exit Sub
  End If

  PopulateAccessGrid
  mobjOutputDef.ReadDefFromRecset rsTemp


  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader "Cross Tab : " & rsTemp!Name
        
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
        
        .PrintNormal "Base Table : " & GetItemName(True, rsTemp!TableID)
        'MH20000905 Check if picklist/filter has been deleted
        Select Case Val(rsTemp!Selection)
        Case 0: .PrintNormal "Records : All"
        Case 1: .PrintNormal "Records : Picklist '" & IIf(IsNull(rsTemp!PicklistName), "<Deleted>", rsTemp!PicklistName) & "'"
        Case 2: .PrintNormal "Records : Filter '" & IIf(IsNull(rsTemp!FilterName), "<Deleted>", rsTemp!FilterName) & "'"
        End Select
        
        .PrintNormal
        .PrintNormal "Display filter or picklist title in the report header : " & IIf(rsTemp!PrintFilterHeader = True, "Yes", "No")
        .PrintNormal
        
        '---------
        
        .PrintTitle "Columns"
    
        .PrintNormal "Horizontal Column : " & GetItemName(False, rsTemp!HorizontalColID)
        If Val(rsTemp!HorizontalStart) <> 0 Or Val(rsTemp!HorizontalStop) <> 0 Then
          .PrintNormal "Horizontal Start : " & datGeneral.ConvertNumberForDisplay(rsTemp!HorizontalStart)
          .PrintNormal "Horizontal Stop : " & datGeneral.ConvertNumberForDisplay(rsTemp!HorizontalStop)
          .PrintNormal "Horizontal Increment : " & datGeneral.ConvertNumberForDisplay(rsTemp!HorizontalStep)
        End If
        
        .PrintNormal
        
      
        .PrintNormal "Vertical Column : " & GetItemName(False, rsTemp!VerticalColID)
        If Val(rsTemp!VerticalStart) <> 0 Or Val(rsTemp!VerticalStop) <> 0 Then
          .PrintNormal "Vertical Start : " & datGeneral.ConvertNumberForDisplay(rsTemp!VerticalStart)
          .PrintNormal "Vertical Stop : " & datGeneral.ConvertNumberForDisplay(rsTemp!VerticalStop)
          .PrintNormal "Vertical Increment : " & datGeneral.ConvertNumberForDisplay(rsTemp!VerticalStep)
        End If
        
        .PrintNormal
    
    
        If rsTemp!PageBreakColID > 0 Then
          .PrintNormal "Page Break Column : " & GetItemName(False, rsTemp!PageBreakColID)
          If Val(rsTemp!PageBreakStart) <> 0 Or Val(rsTemp!PageBreakStop) <> 0 Then
            .PrintNormal "Page Break Start : " & datGeneral.ConvertNumberForDisplay(rsTemp!PageBreakStart)
            .PrintNormal "Page Break Stop : " & datGeneral.ConvertNumberForDisplay(rsTemp!PageBreakStop)
            .PrintNormal "Page Break Increment : " & datGeneral.ConvertNumberForDisplay(rsTemp!PageBreakStep)
          End If
        
        Else
          .PrintNormal "Page Break Column : <None>"
        
        End If
        
        .PrintNormal
      
        If rsTemp!IntersectionColID > 0 Then
          .PrintNormal "Intersection Column : " & GetItemName(False, rsTemp!IntersectionColID)
        Else
          .PrintNormal "Intersection Column : <None>"
        End If
        
        
        .PrintNormal "Intersection Type : " & _
            Choose(rsTemp!IntersectionType + 1, "Count", "Average", "Maximum", "Minimum", "Total")
        .PrintNormal
    
        .PrintNormal "Percentage of Type : " & IIf(rsTemp!Percentage, "True", "False")
        .PrintNormal "Percentage of Page : " & IIf(rsTemp!PercentageofPage, "True", "False")
        .PrintNormal "Suppress Zeros : " & IIf(rsTemp!SuppressZeros, "True", "False")
        .PrintNormal "Use 1000 separators (,) : " & IIf(rsTemp!ThousandSeparators, "True", "False")
        .PrintNormal
      
        '---------

        .PrintTitle "Output"

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
        .PrintConfirm "Cross Tab : " & rsTemp!Name, "Cross Tab Definition"
    
      End If
    End With
  
  End If
    
  rsTemp.Close
  Set rsTemp = Nothing
  Set datData = Nothing

Exit Sub

LocalErr:
  ErrorMsgbox "Printing Cross Tab Definition Failed"

End Sub

Private Sub CheckIfEnableRangeLabels()
    
  Dim blnNoRangesEnabled As Boolean
    
  blnNoRangesEnabled = (mskHorizontalRange(0).Enabled = False And _
                        mskVerticalRange(0).Enabled = False And _
                        mskPageBreakRange(0).Enabled = False)

  lblRange(0).ForeColor = IIf(blnNoRangesEnabled, vbApplicationWorkspace, vbWindowText)
  lblRange(1).ForeColor = IIf(blnNoRangesEnabled, vbApplicationWorkspace, vbWindowText)
  lblRange(2).ForeColor = IIf(blnNoRangesEnabled, vbApplicationWorkspace, vbWindowText)

End Sub

'*** OUTPUT OPTIONS ***
Private Sub optOutputFormat_Click(Index As Integer)
  mobjOutputDef.FormatClick Index
  Changed = True
End Sub

Private Sub chkDestination_Click(Index As Integer)
  mobjOutputDef.DestinationClick Index
  Changed = True
End Sub


