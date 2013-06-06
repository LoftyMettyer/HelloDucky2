VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmMatchDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Match Report Definition"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
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
   Icon            =   "frmMatchDef.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDocument 
      Height          =   510
      Left            =   1965
      Picture         =   "frmMatchDef.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   6060
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picNoDrop 
      Height          =   495
      Left            =   1380
      Picture         =   "frmMatchDef.frx":08D6
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   6075
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8450
      TabIndex        =   80
      Top             =   6000
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7185
      TabIndex        =   79
      Top             =   6000
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5805
      Left            =   50
      TabIndex        =   86
      Top             =   90
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   10239
      _Version        =   393216
      Style           =   1
      Tabs            =   5
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
      TabPicture(0)   =   "frmMatchDef.frx":11A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinition(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDefinition(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ta&bles"
      TabPicture(1)   =   "frmMatchDef.frx":11BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRelations"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Colu&mns"
      TabPicture(2)   =   "frmMatchDef.frx":11D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraFieldsSelected"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraFieldsAvailable"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraFieldButtons"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&Sort Order"
      TabPicture(3)   =   "frmMatchDef.frx":11F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraReportOrder"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "O&utput"
      TabPicture(4)   =   "frmMatchDef.frx":1210
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraReportOptions"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fraOutputFormat"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraOutputDestination"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   3990
         Left            =   -72240
         TabIndex        =   87
         Top             =   1665
         Width           =   6675
         Begin VB.CheckBox chkPreview 
            Caption         =   "P&review on screen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   100
            Top             =   400
            Width           =   3495
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   99
            Top             =   2720
            Width           =   1605
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   98
            Top             =   1820
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   97
            Top             =   1300
            Width           =   1605
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   96
            Top             =   850
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6240
            TabIndex        =   95
            Top             =   2660
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6240
            TabIndex        =   94
            Top             =   1760
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   2660
            Width           =   2880
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            TabIndex        =   92
            Top             =   3060
            Width           =   3180
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   2160
            Width           =   3180
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   1240
            Width           =   3180
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   89
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1760
            Width           =   2880
         End
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3360
            TabIndex        =   88
            Tag             =   "0"
            Top             =   3460
            Width           =   3180
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   106
            Top             =   1305
            Width           =   1455
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   105
            Top             =   2220
            Width           =   1350
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   104
            Top             =   2715
            Width           =   1200
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   103
            Top             =   3120
            Width           =   1305
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   102
            Top             =   1815
            Width           =   1050
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   1800
            TabIndex        =   101
            Top             =   3525
            Width           =   1065
         End
      End
      Begin VB.Frame fraFieldButtons 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "fraFieldButtons"
         Height          =   5300
         Left            =   -70900
         TabIndex        =   85
         Top             =   360
         Width           =   1535
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   100
            TabIndex        =   43
            Top             =   2535
            Width           =   1305
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add A&ll"
            Height          =   400
            Left            =   100
            TabIndex        =   41
            Top             =   1335
            Width           =   1305
         End
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   100
            TabIndex        =   45
            Top             =   3735
            Width           =   1305
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "&Up"
            Enabled         =   0   'False
            Height          =   400
            Left            =   100
            TabIndex        =   44
            Top             =   3240
            Width           =   1305
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   100
            TabIndex        =   42
            Top             =   2025
            Width           =   1305
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   400
            Left            =   100
            TabIndex        =   40
            Top             =   840
            Width           =   1305
         End
      End
      Begin VB.Frame fraOutputFormat 
         Caption         =   "Output Format :"
         Height          =   3990
         Left            =   -74880
         TabIndex        =   71
         Top             =   1665
         Width           =   2500
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "D&ata Only"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   72
            Top             =   400
            Value           =   -1  'True
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CS&V File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   73
            Top             =   800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   74
            Top             =   1200
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   75
            Top             =   1600
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   76
            Top             =   2000
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Char&t"
            Height          =   195
            Index           =   5
            Left            =   200
            TabIndex        =   77
            Top             =   2400
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel P&ivot Table"
            Height          =   195
            Index           =   6
            Left            =   200
            TabIndex        =   78
            Top             =   2800
            Width           =   1900
         End
      End
      Begin VB.Frame fraRelations 
         Height          =   5300
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   9350
         Begin VB.CommandButton cmdNewRelation 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7950
            TabIndex        =   32
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditRelation 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   33
            Top             =   900
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteRelation 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   34
            Top             =   1500
            Width           =   1200
         End
         Begin VB.CommandButton cmdClearRelations 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   35
            Top             =   2100
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdRelations 
            Height          =   4800
            Left            =   195
            TabIndex        =   31
            Top             =   300
            Width           =   7425
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   7
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
            SelectTypeRow   =   1
            BalloonHelp     =   0   'False
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   7
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "TableID"
            Columns(0).Name =   "TableID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   2619
            Columns(1).Caption=   "Table"
            Columns(1).Name =   "Table Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "MatchTableID"
            Columns(2).Name =   "MatchTableID"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   2619
            Columns(3).Caption=   "Match Table"
            Columns(3).Name =   "Match Table"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   2619
            Columns(4).Caption=   "Required Matches"
            Columns(4).Name =   "Required Matches"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   2619
            Columns(5).Caption=   "Preferred Matches"
            Columns(5).Name =   "Preferred Matches"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   2566
            Columns(6).Caption=   "Score Calculation"
            Columns(6).Name =   "Match Score"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13097
            _ExtentY        =   8467
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
      Begin VB.Frame fraReportOptions 
         Caption         =   "Matched Records :"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   61
         Top             =   400
         Width           =   9315
         Begin VB.CheckBox chkEqualGrade 
            Caption         =   "Allow progress to e&qual grade"
            Height          =   195
            Left            =   6315
            TabIndex        =   69
            Top             =   360
            Width           =   2925
         End
         Begin VB.OptionButton optLowest 
            Caption         =   "&Lowest Match Scores"
            Height          =   195
            Left            =   200
            TabIndex        =   66
            Top             =   760
            Width           =   2160
         End
         Begin COASpinner.COA_Spinner spnMaxRecords 
            Height          =   315
            Left            =   4320
            TabIndex        =   64
            Top             =   300
            Width           =   765
            _ExtentX        =   1349
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
            MaximumValue    =   999
            Text            =   "0"
         End
         Begin COASpinner.COA_Spinner spnLimit 
            Height          =   315
            Left            =   4320
            TabIndex        =   68
            Top             =   705
            Width           =   765
            _ExtentX        =   1349
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
            MaximumValue    =   10000
            Text            =   "0"
         End
         Begin VB.OptionButton optHighest 
            Caption         =   "Hi&ghest Match Scores"
            Height          =   195
            Left            =   200
            TabIndex        =   62
            Top             =   360
            Value           =   -1  'True
            Width           =   2160
         End
         Begin VB.CheckBox chkReportStructure 
            Caption         =   "Restrict by reporti&ng structure"
            Height          =   195
            Left            =   6315
            TabIndex        =   70
            Top             =   660
            Width           =   2940
         End
         Begin VB.CheckBox chkLimit 
            Caption         =   "Minimum Score :"
            Height          =   195
            Left            =   2550
            TabIndex        =   67
            Top             =   760
            Width           =   1800
         End
         Begin VB.Label lblMaxRecords 
            AutoSize        =   -1  'True
            Caption         =   "Matched Records :"
            Height          =   195
            Left            =   2550
            TabIndex        =   63
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblAllRecords 
            AutoSize        =   -1  'True
            Caption         =   "(All Records)"
            Height          =   195
            Left            =   5115
            TabIndex        =   65
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraReportOrder 
         Caption         =   "Sort Order :"
         Height          =   5180
         Left            =   -74850
         TabIndex        =   53
         Top             =   440
         Width           =   9315
         Begin VB.CommandButton cmdClearOrder 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7920
            TabIndex        =   58
            Top             =   1920
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveDownOrder 
            Caption         =   "Move Do&wn"
            Height          =   400
            Left            =   7900
            TabIndex        =   60
            Top             =   4515
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveUpOrder 
            Caption         =   "Move U&p"
            Height          =   400
            Left            =   7900
            TabIndex        =   59
            Top             =   3975
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditOrder 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   7900
            TabIndex        =   56
            Top             =   850
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteOrder 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   7900
            TabIndex        =   57
            Top             =   1400
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddOrder 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7900
            TabIndex        =   55
            Top             =   300
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdReportOrder 
            Height          =   4620
            Left            =   195
            TabIndex        =   54
            Top             =   300
            Width           =   7410
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            GroupHeadLines  =   0
            Col.Count       =   7
            stylesets.count =   2
            stylesets(0).Name=   "ssetDormant"
            stylesets(0).ForeColor=   0
            stylesets(0).BackColor=   16777215
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
            stylesets(0).Picture=   "frmMatchDef.frx":122C
            stylesets(1).Name=   "ssetActive"
            stylesets(1).ForeColor=   16777215
            stylesets(1).BackColor=   -2147483646
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
            stylesets(1).Picture=   "frmMatchDef.frx":1248
            MultiLine       =   0   'False
            AllowRowSizing  =   0   'False
            AllowGroupSizing=   0   'False
            AllowGroupMoving=   0   'False
            AllowColumnMoving=   0
            AllowGroupSwapping=   0   'False
            AllowColumnSwapping=   0
            AllowGroupShrinking=   0   'False
            AllowDragDrop   =   0   'False
            SelectTypeCol   =   0
            SelectTypeRow   =   1
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   7
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ColumnID"
            Columns(0).Name =   "ColExprID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   9975
            Columns(1).Caption=   "Column"
            Columns(1).Name =   "Column"
            Columns(1).CaptionAlignment=   0
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   2619
            Columns(2).Caption=   "Order"
            Columns(2).Name =   "Order"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   4
            Columns(2).Locked=   -1  'True
            Columns(2).Style=   3
            Columns(2).Row.Count=   2
            Columns(2).Col.Count=   2
            Columns(2).Row(0).Col(0)=   "Ascending"
            Columns(2).Row(1).Col(0)=   "Descending"
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "Break on Change"
            Columns(3).Name =   "ColType"
            Columns(3).CaptionAlignment=   2
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   11
            Columns(3).FieldLen=   1
            Columns(3).Style=   2
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "Page on Change"
            Columns(4).Name =   "Page"
            Columns(4).CaptionAlignment=   2
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   11
            Columns(4).FieldLen=   256
            Columns(4).Style=   2
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "Value on Change"
            Columns(5).Name =   "Value"
            Columns(5).CaptionAlignment=   2
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   11
            Columns(5).FieldLen=   256
            Columns(5).Style=   2
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "Suppress Repeated Values"
            Columns(6).Name =   "Hide"
            Columns(6).CaptionAlignment=   2
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   11
            Columns(6).FieldLen=   256
            Columns(6).Style=   2
            _ExtentX        =   13070
            _ExtentY        =   8149
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
      Begin VB.Frame fraFieldsAvailable 
         Caption         =   "Columns Available :"
         Height          =   5300
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   3615
         Begin VB.ComboBox cboTblAvailable 
            Height          =   315
            Left            =   200
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   300
            Width           =   3225
         End
         Begin ComctlLib.ListView ListView1 
            Height          =   4335
            Left            =   195
            TabIndex        =   38
            Top             =   720
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   7646
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
         Height          =   5300
         Left            =   -69100
         TabIndex        =   39
         Top             =   360
         Width           =   3580
         Begin VB.CheckBox chkProp_IsNumeric 
            Caption         =   "chkProp_IsNumeric"
            Height          =   255
            Left            =   2760
            TabIndex        =   84
            Top             =   4440
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtProp_ColumnHeading 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1170
            TabIndex        =   48
            Top             =   3960
            Width           =   2250
         End
         Begin COASpinner.COA_Spinner spnSize 
            Height          =   315
            Left            =   1170
            TabIndex        =   50
            Top             =   4395
            Width           =   1410
            _ExtentX        =   2487
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
            Height          =   3540
            Left            =   195
            TabIndex        =   46
            Top             =   300
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   6244
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Left            =   1170
            TabIndex        =   52
            Top             =   4785
            Width           =   1410
            _ExtentX        =   2487
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
            MaximumValue    =   9999
            Text            =   "0"
         End
         Begin VB.Label lblProp_ColumnHeading 
            BackStyle       =   0  'Transparent
            Caption         =   "Heading :"
            Height          =   195
            Left            =   195
            TabIndex        =   47
            Top             =   4050
            Width           =   1260
         End
         Begin VB.Label lblProp_Size 
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            Height          =   195
            Left            =   195
            TabIndex        =   49
            Top             =   4455
            Width           =   570
         End
         Begin VB.Label lblProp_Decimals 
            BackStyle       =   0  'Transparent
            Caption         =   "Decimals :"
            Height          =   195
            Left            =   195
            TabIndex        =   51
            Top             =   4845
            Width           =   1380
         End
      End
      Begin VB.Frame fraDefinition 
         Caption         =   "Data :"
         Height          =   3275
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2380
         Width           =   9360
         Begin VB.CheckBox chkPrintFilterHeader 
            Caption         =   "Display &title in the report header"
            Enabled         =   0   'False
            Height          =   240
            Left            =   5010
            TabIndex        =   19
            Top             =   1520
            Width           =   3420
         End
         Begin VB.Frame fraCriteriaSelection 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1095
            Left            =   4990
            TabIndex        =   81
            Top             =   2040
            Width           =   4275
            Begin VB.TextBox txtFilter 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   720
               Width           =   1950
            End
            Begin VB.TextBox txtPicklist 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   345
               Width           =   1950
            End
            Begin VB.OptionButton optFilter 
               Caption         =   "F&ilter"
               Height          =   195
               Index           =   1
               Left            =   900
               TabIndex        =   27
               Top             =   785
               Width           =   840
            End
            Begin VB.OptionButton optPicklist 
               Caption         =   "Pic&klist"
               Height          =   195
               Index           =   1
               Left            =   900
               TabIndex        =   24
               Top             =   405
               Width           =   885
            End
            Begin VB.OptionButton optAllRecords 
               Caption         =   "A&ll"
               Height          =   195
               Index           =   1
               Left            =   900
               TabIndex        =   23
               Top             =   25
               Value           =   -1  'True
               Width           =   540
            End
            Begin VB.CommandButton cmdPicklist 
               Caption         =   "..."
               DisabledPicture =   "frmMatchDef.frx":1264
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   3870
               TabIndex        =   26
               Top             =   345
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.CommandButton cmdFilter 
               Caption         =   "..."
               DisabledPicture =   "frmMatchDef.frx":15C5
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   3870
               TabIndex        =   29
               Top             =   720
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.Label lblTable2Records 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Records :"
               Height          =   195
               Left            =   0
               TabIndex        =   22
               Top             =   30
               Width           =   690
            End
         End
         Begin VB.ComboBox cboTable2 
            Height          =   315
            Left            =   1635
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1995
            Width           =   3000
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   8885
            TabIndex        =   18
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   8885
            TabIndex        =   15
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboTable1 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   315
            Width           =   3000
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Index           =   0
            Left            =   5910
            TabIndex        =   12
            Top             =   365
            Value           =   -1  'True
            Width           =   540
         End
         Begin VB.OptionButton optPicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Index           =   0
            Left            =   5910
            TabIndex        =   13
            Top             =   750
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   195
            Index           =   0
            Left            =   5910
            TabIndex        =   16
            Top             =   1120
            Width           =   840
         End
         Begin VB.TextBox txtPicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6930
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   705
            Width           =   1950
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6930
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1080
            Width           =   1950
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Match Table :"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Index           =   3
            Left            =   5010
            TabIndex        =   11
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame fraDefinition 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   400
         Width           =   9360
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5805
            MaxLength       =   30
            TabIndex        =   6
            Top             =   315
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   2
            Top             =   315
            Width           =   3000
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1620
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   705
            Width           =   3000
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1080
            Left            =   5805
            TabIndex        =   107
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
            stylesets(0).Picture=   "frmMatchDef.frx":1926
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
            stylesets(1).Picture=   "frmMatchDef.frx":1942
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Index           =   2
            Left            =   5010
            TabIndex        =   5
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   1
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   3
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Index           =   3
            Left            =   5010
            TabIndex        =   7
            Top             =   810
            Width           =   780
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   690
      Top             =   5955
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchDef.frx":195E
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchDef.frx":1D2A
            Key             =   "IMG_CALC"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchDef.frx":213A
            Key             =   "IMG_MATCH"
         EndProperty
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   2640
      Top             =   6000
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
      Bands           =   "frmMatchDef.frx":251E
   End
End
Attribute VB_Name = "frmMatchDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private datData As clsDataAccess
Private mblnLoading As Boolean
Private mblnDefinitionCreator As Boolean
Private mlngMatchReportID As Long
Private mlngMatchReportType As MatchReportType
Private mobjOutputDef As clsOutputDef

Private mstrTable1Name As String
Private mstrTable2Name As String
Private mblnFromCopy As Boolean
Private mblnFormPrint As Boolean
Private mblnColumnDrag As Boolean
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mlngTimeStamp As Long
Private mblnForceHidden As Boolean
Private mblnRecordSelectionInvalid As Boolean
Private mblnDeleted As Boolean
Private mbNeedsSave As Boolean

Private colRelatedTables As Collection
Private mcolMatchReportColDetails As Collection

Private mlngExprDeleteOnOK() As Long
Private mlngExprDeleteOnCancel() As Long

Public Property Get SelectedID() As Long
  SelectedID = mlngMatchReportID
End Property

Public Property Get MatchReportType() As MatchReportType
  MatchReportType = mlngMatchReportType
End Property

Public Property Let MatchReportType(lngMatchReportType As MatchReportType)
  
  mlngMatchReportType = lngMatchReportType
  chkEqualGrade.Visible = (mlngMatchReportType <> mrtNormal)
  chkReportStructure.Visible = (mlngMatchReportType <> mrtNormal)

  Select Case mlngMatchReportType
  Case mrtSucession
    'JPD 20030911 Fault 6359
    Me.Caption = "Succession Planning Definition"
    Me.HelpContextID = 1076
  Case mrtCareer
    'JPD 20030911 Fault 6359
    Me.Caption = "Career Progression Definition"
    Me.HelpContextID = 1078
  End Select

End Property

Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional plngMatchReportID As Long, Optional bPrint As Boolean) As Boolean

  ' This function is called from frmMain and prepares the form depending
  ' on whether the user is creating a new definition or editing an existing
  ' one.
  
  Screen.MousePointer = vbHourglass
  
  ' Set references to class modules
  Set datData = New HRProDataMgr.clsDataAccess
  
  mblnLoading = True

  If bNew Then
    mblnDefinitionCreator = True
    
    'Set ID to 0 to indicate new record
    mlngMatchReportID = 0

    'Set controls to defaults
    ClearForNew
    
    LoadTable1Combo
    LoadTable2Combo

'    UpdateDependantFields
    
    PopulateTableAvailable , True
    
    PopulateAccessGrid
    Changed = False
    
  Else

    mlngMatchReportID = plngMatchReportID
    
    ' Is is a copy of an existing one ?
    FromCopy = bCopy
    
    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint
    
    PopulateAccessGrid
    
    If Not RetrieveMatchReportDetails(plngMatchReportID) Then
      If mblnDeleted Or Cancelled Then
        Initialise = False
        Exit Function
      Else
        If MsgBox("HR Pro could not load all of the definition successfully. The recommendation is that" & vbCrLf & _
               "you delete the definition and create a new one, however, you may edit the existing" & vbCrLf & _
               "definition if you wish. Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
          Cancelled = True
          Initialise = False
          Exit Function
        End If
      End If
    End If

    PopulateTableAvailable , True
'    UpdateOrderButtons
'
    
    If bCopy = True Then
      mlngMatchReportID = 0
      Changed = True
    Else
      Changed = mblnRecordSelectionInvalid And (Not mblnReadOnly) ' False
    End If

  End If
'
'  If mblnForceHidden Then mblnForceHidden = True
'  Cancelled = False
  Screen.MousePointer = vbNormal
  mblnLoading = False
  
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
  Select Case mlngMatchReportType
    Case mrtSucession
      Set rsAccess = GetUtilityAccessRecords(utlSuccession, mlngMatchReportID, mblnFromCopy)
    Case mrtCareer
      Set rsAccess = GetUtilityAccessRecords(utlCareer, mlngMatchReportID, mblnFromCopy)
    Case Else
      Set rsAccess = GetUtilityAccessRecords(utlMatchReport, mlngMatchReportID, mblnFromCopy)
  End Select
  
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


Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOk.Enabled = pblnChanged
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

Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FromCopy(ByVal bCopy As Boolean)
  mblnFromCopy = bCopy
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property

Public Property Get DefinitionOwner() As Boolean
  DefinitionOwner = mblnDefinitionCreator
End Property

Private Sub cboPrinterName_Click()
  Changed = True
End Sub

Private Sub cboSaveExisting_Click()
  Changed = True
End Sub

'Private Function TableAlreadyAvailable(plngTableID As Long) As Boolean
'
'  Dim i As Integer
'
'  With cboTblAvailable
'    For i = 0 To .ListCount - 1 Step 1
'      If .ItemData(i) = plngTableID Then
'        TableAlreadyAvailable = True
'        Exit Function
'      End If
'    Next i
'  End With
'
'End Function


Private Sub cboTable1_Click()
  
  ' When the user changes the Base Table, check to see if the user
  ' has defined any columns in the report. If they have, check that
  ' they have selected a different table in the combo to the one that
  ' was there before. If so, then prompt user, otherwise, go ahead and
  ' clear the definition
  If mblnLoading = True Then Exit Sub
  If mstrTable1Name = cboTable1.Text And (mblnLoading = False) Then Exit Sub
  
  'If (mstrTable1Name <> cboTable1.Text) Or (mblnLoading = True) Then
    If ListView2.ListItems.Count > 0 Or grdRelations.Rows > 0 Then
      If MsgBox("Warning: Changing a table will result in all table/column " & _
            "specific aspects of this report definition being cleared." & vbCrLf & _
            "Are you sure you wish to continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
    
        mblnLoading = True
        ClearForNew
        mblnLoading = False
        Changed = True
        
      Else
        ' User opted to abort the base table change
        SetComboText cboTable1, mstrTable1Name
        Exit Sub
      
      End If
    Else
      Changed = True
    End If
  'End If
  
  mstrTable1Name = cboTable1.Text
  LoadTable2Combo


  '01/08/2001 MH Fault 2615
  optAllRecords(0).Value = True
  
  UpdateDependantFields
  PopulateTableAvailable , True
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub


Private Sub cboTable2_Click()

  ' When the user changes the Base Table, check to see if the user
  ' has defined any columns in the report. If they have, check that
  ' they have selected a different table in the combo to the one that
  ' was there before. If so, then prompt user, otherwise, go ahead and
  ' clear the definition
  If mstrTable2Name = cboTable2.Text And (mblnLoading = False) Then Exit Sub
  
  If mblnLoading = False Then

    'If (mstrTable2Name <> cboTable1.Text) Or (mblnLoading = True) Then
      If ListView2.ListItems.Count > 0 Or grdRelations.Rows > 0 Then
        If MsgBox("Warning: Changing the criteria table will result in all table/column " & _
              "specific aspects of this report definition being cleared." & vbCrLf & _
              "Are you sure you wish to continue?", _
              vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
      
          mblnLoading = True
          ClearForNew
          mblnLoading = False
          Changed = True
          
        Else
          ' User opted to abort the base table change
          SetComboText cboTable2, mstrTable2Name
          Exit Sub
        
        End If
      Else
        Changed = True
      End If
    'End If
    
    mstrTable2Name = cboTable2.Text
    
    UpdateDependantFields
    PopulateTableAvailable , True
    ForceDefinitionToBeHiddenIfNeeded

  End If

  lblTable2Records.Enabled = (cboTable2.ListIndex > 0)
  optAllRecords(1).Enabled = (cboTable2.ListIndex > 0)
  optPicklist(1).Enabled = (cboTable2.ListIndex > 0)
  optFilter(1).Enabled = (cboTable2.ListIndex > 0)
  'JPD 20040628 Fault 8854
  'If cboTable2.ListIndex = 0 Then
    optAllRecords(1).Value = True
  'End If

End Sub

Private Sub cboTblAvailable_Click()

  PopulateAvailable
  UpdateButtonStatus (SSTab1.Tab)

End Sub

Private Sub chkEqualGrade_Click()
  Changed = True
End Sub

Private Sub chkLimit_Click()
  spnLimit.Enabled = (chkLimit.Value = vbChecked)
  spnLimit.BackColor = IIf(chkLimit.Value = vbChecked, vbWindowBackground, vbButtonFace)
  If chkLimit.Value = vbUnchecked Then spnLimit.Value = 0
  Changed = True
End Sub

Private Sub chkPreview_Click()
  Changed = True
End Sub

Private Sub chkPrintFilterHeader_Click()
  Changed = True
End Sub

Private Sub chkReportStructure_Click()
  Changed = True
End Sub

Private Sub cmdAdd_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAdd.Picture = cmdAdd.Picture

End Sub

Private Sub cmdAddAll_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAddAll.Picture = cmdAddAll.Picture

End Sub

Private Sub cmdAddOrder_Click()

  grdReportOrder.Redraw = False
    
  If frmMailMergeOrder.Initialise(True, Me) = True Then
    frmMailMergeOrder.Icon = Me.Icon
    frmMailMergeOrder.Caption = Me.Caption
    frmMailMergeOrder.Show vbModal
    
    'AE20071025 Fault #6797
    If Not frmMailMergeOrder.UserCancelled Then
      grdReportOrder.MoveLast
      grdReportOrder.SelBookmarks.Add grdReportOrder.Bookmark
      UpdateOrderButtons
      Changed = True
    End If
  End If
  
  Unload frmMailMergeOrder
  Set frmMailMergeOrder = Nothing
  
  grdReportOrder.Redraw = True
  
End Sub

Private Sub UpdateOrderButtons()

  If mblnReadOnly Then
    Exit Sub
  End If
  
  grdReportOrder.Redraw = False
  
  If grdReportOrder.Rows = 0 Then
    cmdEditOrder.Enabled = False
    cmdDeleteOrder.Enabled = False
    cmdClearOrder.Enabled = False
    'grdReportOrder.Columns("Column").Width = 2340
  Else
    cmdEditOrder.Enabled = True
    cmdDeleteOrder.Enabled = True
    cmdClearOrder.Enabled = True
    If grdReportOrder.Rows > grdReportOrder.VisibleRows Then
      'grdReportOrder.Columns("Column").Width = grdReportOrder.Columns("Column").Width - 130
    End If
  End If

  With grdReportOrder
    If .AddItemRowIndex(.Bookmark) = 0 Then
      cmdMoveUpOrder.Enabled = False
      cmdMoveDownOrder.Enabled = .Rows > 1
    ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
      cmdMoveUpOrder.Enabled = .Rows > 1
      cmdMoveDownOrder.Enabled = False
    Else
      'TM20020809 Fault 4244 - only enable the move buttons if more than one row exists.
      cmdMoveUpOrder.Enabled = .Rows > 1
      cmdMoveDownOrder.Enabled = .Rows > 1
    End If
  End With

'  If grdReportOrder.Rows = 0 Then
'    cmdMoveUpOrder.Enabled = False
'    cmdMoveDownOrder.Enabled = False
'    Exit Sub
'  End If
  
'  If grdReportOrder.AddItemRowIndex(grdReportOrder.Bookmark) < 1 Then cmdMoveUpOrder.Enabled = False Else cmdMoveUpOrder.Enabled = True
'  If grdReportOrder.AddItemRowIndex(grdReportOrder.Bookmark) = (grdReportOrder.Rows - 1) Then cmdMoveDownOrder.Enabled = False Else cmdMoveDownOrder.Enabled = True

'  If grdReportOrder.SelBookmarks.Count = 0 Then
'    cmdMoveUpOrder.Enabled = False
'    cmdMoveDownOrder.Enabled = False
'  End If
  
  grdReportOrder.Columns("Order").Width = IIf(grdReportOrder.Rows > 16, 1485, 1730)
  
  grdReportOrder.Redraw = True

End Sub

Private Sub cmdCancel_Click()
  
  Unload Me

End Sub

Private Sub cmdClearOrder_Click()

  Dim varBookmark As Variant
  Dim lngTable1ID As Long
  Dim lngTable2ID As Long
  Dim lRow As Long
  Dim intMBResponse As Integer

  intMBResponse = MsgBox("Are you sure you wish to clear the sort order?", vbQuestion + vbYesNo, Me.Caption)
  If intMBResponse = vbYes Then
    If grdReportOrder.Rows = 1 Then
      grdReportOrder.RemoveItem 0
    Else
      grdReportOrder.RemoveAll
    End If

    Changed = True
    UpdateButtonStatus SSTab1.Tab
  End If

End Sub

Private Sub cmdClearRelations_Click()

  Dim varBookmark As Variant
  Dim lngTable1ID As Long
  Dim lngTable2ID As Long
  Dim lRow As Long
  Dim intMBResponse As Integer

  intMBResponse = MsgBox("Are you sure you want to clear the table relations?", vbExclamation + vbYesNo, Me.Caption)
  If intMBResponse = vbYes Then
  
    With grdRelations
      For lRow = 0 To .Rows - 1
        varBookmark = .AddItemBookmark(lRow)
        lngTable1ID = .Columns("TableID").CellValue(varBookmark)
        lngTable2ID = .Columns("MatchTableID").CellValue(varBookmark)
        ExprDeleteOnOK lngTable2ID, Val(.Columns("Required Matches").CellValue(varBookmark)), giEXPR_MATCHWHEREEXPRESSION
        ExprDeleteOnOK lngTable1ID, Val(.Columns("Preferred Matches").CellValue(varBookmark)), giEXPR_MATCHJOINEXPRESSION
        ExprDeleteOnOK lngTable1ID, Val(.Columns("Match Score").CellValue(varBookmark)), giEXPR_MATCHSCOREEXPRESSION
      Next
    
      If .Rows = 1 Then
        .RemoveItem 0
      Else
        .RemoveAll
      End If
    End With

    Set colRelatedTables = Nothing
    Set colRelatedTables = New Collection

    Changed = True
    UpdateButtonStatus (1)
  End If

End Sub

Private Sub cmdDeleteOrder_Click()

  Dim lRow As Long
  
  If grdReportOrder.Rows = 1 Then
    grdReportOrder.RemoveAll
  Else
    With grdReportOrder
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
  
  UpdateOrderButtons
  Changed = True

End Sub


'Private Sub cmdEditChild_Click()
'
'  Dim pstrRow As String
'  Dim plngRow As Long
'  Dim pfrmChild As New frmMatchReportChilds
'  Dim lngInitTableID As Long
'  Dim i2 As Integer
'  Dim bNeedRefreshAvail As Boolean
'
'  With grdChildren
'    plngRow = .AddItemRowIndex(.Bookmark)
'    lngInitTableID = .Columns("TableID").Value
'    pfrmChild.Initialize False _
'                , Me _
'                , .Columns("TableID").Value _
'                , .Columns("Table").Value _
'                , .Columns("FilterID").Value _
'                , .Columns("Filter").Value _
'                , IIf(.Columns("Records").Value = sALL_RECORDS, 0, .Columns("Records").Value)
'  End With
'
'  With pfrmChild
'    .Show vbModal
'
'    If Not .Cancelled Then
'      If .ChildTableID <> lngInitTableID Then
'        ' Check if any columns in the report definition are from the table that was
'        ' previously selected in the child combo box. If so, prompt user for action.
'        Select Case AnyChildColumnsUsed(lngInitTableID)
'        Case 2: ' child cols used and user wants to continue with the change
'          'TM20020424 Fault 3803
'          If ListView2.ListItems.Count > 0 Then
'            SelectLast ListView2
'            GetCurrentDetails ListView2.SelectedItem.Key
'          End If
'          pstrRow = .ChildTableID _
'                    & vbTab & .ChildTable _
'                    & vbTab & .FilterID _
'                    & vbTab & .Filter _
'                    & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)
'
'        Case 1: ' child cols used and user has aborted the change
'          With grdChildren
'            .Bookmark = .AddItemBookmark(plngRow)
'            .SelBookmarks.RemoveAll
'            .SelBookmarks.Add .AddItemBookmark(plngRow)
'          End With
'
'          Exit Sub
'
'        Case 0: ' no child cols used
'          pstrRow = .ChildTableID _
'                    & vbTab & .ChildTable _
'                    & vbTab & .FilterID _
'                    & vbTab & .Filter _
'                    & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)
'
'        End Select
'      Else
'        pstrRow = .ChildTableID _
'                  & vbTab & .ChildTable _
'                  & vbTab & .FilterID _
'                  & vbTab & .Filter _
'                  & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)
'      End If
'
'      With grdChildren
'        'TM20020424 Fault 3715
'        'Find and remove from Table Available
'        For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
'          If cboTblAvailable.ItemData(i2) = lngInitTableID Then
'
'            If cboTblAvailable.ListIndex = i2 Then
'              bNeedRefreshAvail = True
'            End If
'
'            cboTblAvailable.RemoveItem i2
'            Exit For
'          End If
'        Next i2
'        .RemoveItem plngRow
'        .AddItem pstrRow, plngRow
'      End With
'    Else
'      With grdChildren
'        .Bookmark = .AddItemBookmark(plngRow)
'        .SelBookmarks.RemoveAll
'        .SelBookmarks.Add .AddItemBookmark(plngRow)
'      End With
'
'      Exit Sub
'    End If
'  End With
'
'  Unload pfrmChild
'  Set pfrmChild = Nothing
'
'  PopulateTableAvailable , bNeedRefreshAvail
'
'  EnableDisableTabControls
'
'  ForceDefinitionToBeHiddenIfNeeded
'
'  With grdChildren
'    .Bookmark = .AddItemBookmark(plngRow)
'    .SelBookmarks.RemoveAll
'    .SelBookmarks.Add .AddItemBookmark(plngRow)
'  End With
'
'  UpdateButtonStatus (SSTab1.Tab)
'
'  Changed = True
'
'End Sub


Private Sub cmdDeleteRelation_Click()

  Dim lRow As Long
  Dim lngTable1ID As Long
  Dim lngTable2ID As Long
  Dim intMBResponse As Integer
  
  intMBResponse = MsgBox("Are you sure you want to delete the selected table relations?", vbExclamation + vbYesNo, Me.Caption)
  If intMBResponse = vbYes Then
  
    With grdRelations
      
      lngTable1ID = .Columns("TableID").CellValue(.Bookmark)
      lngTable2ID = .Columns("MatchTableID").CellValue(.Bookmark)
      ExprDeleteOnOK lngTable1ID, colRelatedTables("T" & CStr(lngTable1ID)).RequiredExprID, giEXPR_MATCHWHEREEXPRESSION
      ExprDeleteOnOK lngTable1ID, colRelatedTables("T" & CStr(lngTable1ID)).PreferredExprID, giEXPR_MATCHJOINEXPRESSION
      ExprDeleteOnOK lngTable1ID, colRelatedTables("T" & CStr(lngTable1ID)).MatchScoreID, giEXPR_MATCHSCOREEXPRESSION
      colRelatedTables.Remove "T" & CStr(lngTable1ID)

      If .Rows = 1 Then
        .RemoveAll
      Else
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
      End If
    
    End With

    Changed = True
    UpdateButtonStatus (1)
  End If

End Sub


Private Sub cmdEditOrder_Click()

  Dim lngColumnID As Long
  Dim strSortOrder As String

  With grdReportOrder
    .Redraw = False
    lngColumnID = .Columns("ColumnID").CellValue(.Bookmark)
    strSortOrder = .Columns("Order").CellText(.Bookmark)
  End With

  If frmMailMergeOrder.Initialise(False, Me, lngColumnID, strSortOrder) = True Then
    frmMailMergeOrder.Icon = Me.Icon
    frmMailMergeOrder.Caption = Me.Caption
    frmMailMergeOrder.Show vbModal
    
    'AE20071025 Fault #6797
    If Not frmMailMergeOrder.UserCancelled Then
      UpdateOrderButtons
      Changed = True
    End If
  End If

  Unload frmMailMergeOrder
  Set frmMailMergeOrder = Nothing

  grdReportOrder.Redraw = True

End Sub

Private Sub cmdMoveDown_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveDown.Picture = cmdMoveDown.Picture

End Sub

Private Sub cmdMoveDownOrder_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdReportOrder.AddItemRowIndex(grdReportOrder.Bookmark)
  strSourceRow = grdReportOrder.Columns(0).Text & vbTab & grdReportOrder.Columns(1).Text & vbTab & grdReportOrder.Columns(2).Text & vbTab & grdReportOrder.Columns(3).Text & vbTab & grdReportOrder.Columns(4).Text & vbTab & grdReportOrder.Columns(5).Text & vbTab & grdReportOrder.Columns(6).Text
  
  intDestinationRow = intSourceRow + 1
  grdReportOrder.MoveNext
  strDestinationRow = grdReportOrder.Columns(0).Text & vbTab & grdReportOrder.Columns(1).Text & vbTab & grdReportOrder.Columns(2).Text & vbTab & grdReportOrder.Columns(3).Text & vbTab & grdReportOrder.Columns(4).Text & vbTab & grdReportOrder.Columns(5).Text & vbTab & grdReportOrder.Columns(6).Text
  
  grdReportOrder.RemoveItem intDestinationRow
  grdReportOrder.RemoveItem intSourceRow
  
  grdReportOrder.AddItem strDestinationRow, intSourceRow
  grdReportOrder.AddItem strSourceRow, intDestinationRow
  
  grdReportOrder.SelBookmarks.RemoveAll
  grdReportOrder.MoveNext
  grdReportOrder.Bookmark = grdReportOrder.AddItemBookmark(intDestinationRow)
  grdReportOrder.SelBookmarks.Add grdReportOrder.AddItemBookmark(intDestinationRow)
  
  UpdateButtonStatus (SSTab1.Tab)

  Changed = True
  'mblnBaseTableSpecificChanged = True
  
End Sub

Private Sub cmdMoveUp_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveUp.Picture = cmdMoveUp.Picture

End Sub

Private Sub cmdMoveUpOrder_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdReportOrder.AddItemRowIndex(grdReportOrder.Bookmark)
  strSourceRow = grdReportOrder.Columns(0).Text & vbTab & grdReportOrder.Columns(1).Text & vbTab & grdReportOrder.Columns(2).Text & vbTab & grdReportOrder.Columns(3).Text & vbTab & grdReportOrder.Columns(4).Text & vbTab & grdReportOrder.Columns(5).Text & vbTab & grdReportOrder.Columns(6).Text
  
  intDestinationRow = intSourceRow - 1
  grdReportOrder.MovePrevious
  strDestinationRow = grdReportOrder.Columns(0).Text & vbTab & grdReportOrder.Columns(1).Text & vbTab & grdReportOrder.Columns(2).Text & vbTab & grdReportOrder.Columns(3).Text & vbTab & grdReportOrder.Columns(4).Text & vbTab & grdReportOrder.Columns(5).Text & vbTab & grdReportOrder.Columns(6).Text
  
  grdReportOrder.AddItem strSourceRow, intDestinationRow
  
  grdReportOrder.RemoveItem intSourceRow + 1
  
  grdReportOrder.SelBookmarks.RemoveAll
  grdReportOrder.MovePrevious
  grdReportOrder.Bookmark = grdReportOrder.AddItemBookmark(intDestinationRow)
  grdReportOrder.SelBookmarks.Add grdReportOrder.AddItemBookmark(intDestinationRow)

  UpdateButtonStatus (SSTab1.Tab)
  
  Changed = True
'  mblnBaseTableSpecificChanged = True
  
End Sub

'Private Sub cmdNewCalculation_Click()
'
'  Dim objExpr As New clsExprExpression
'  Set objExpr = New clsExprExpression
'  Dim strKey As String
'  Dim strKeyPrevSelected As String
'  Dim SelectedKeys() As String
'  Dim iCount As Integer
'  Dim sMessage As String
'
'  On Error GoTo NewCalc_ERROR
'
'  ReDim SelectedKeys(0)
'  Dim lst As ListItem
'  For Each lst In ListView1.ListItems
'    If lst.Selected = True Then
'      ReDim Preserve SelectedKeys(UBound(SelectedKeys) + 1)
'      SelectedKeys(UBound(SelectedKeys) - 1) = lst.Key
'    End If
'  Next lst
'
'  With objExpr
''    If .Initialise(cboTable1.ItemData(cboTable1.ListIndex), 0, giEXPR_RUNTIMECALCULATION, 0) Then
''      .SelectExpression True
''    End If
'
'    If .Initialise(cboTblAvailable.ItemData(cboTblAvailable.ListIndex), 0, giEXPR_RUNTIMECALCULATION, 0) Then
'      .SelectExpression True
'    End If
'
'  ' Refresh the listview to show the newly added calculation
'  PopulateAvailable
'
'  ' Refresh the names of selected calcs
'  RefreshSelectedCalcNames
'
'  UpdateButtonStatus (SSTab1.Tab)
'
'
'    If .ExpressionID > 0 Then
'
'      '02/08/2000 MH Fault 2386
'      If mblnReadOnly Then
'        MsgBox "Unable to select calculation as you are viewing a read only definition", vbExclamation, Me.Caption
'      Else
'
'        strKey = "E" & CStr(.ExpressionID)
'        If Not AlreadyUsed(strKey) Then
'
'          'De-select all the available items.
'          For Each lst In ListView1.ListItems
'              lst.Selected = False
'          Next lst
'
'          'Check for hidden elements within the calc.
'          sMessage = IsCalcValid(.ExpressionID)
'          If sMessage <> vbNullString Then
'            MsgBox "This calculation has been deleted or hidden by another user." & vbCrLf & _
'                   "It cannot be added to this definition", vbExclamation, App.Title
'          Else
'            ListView1.ListItems(strKey).Selected = True
'            CopyToSelected False
'          End If
'
'          ' RH 09/04/01 - leaves 2 things highlighted
'          For Each lst In ListView2.ListItems
'            If lst.Selected = True Then
'              lst.Selected = False
'            End If
'          Next lst
'          ListView2.ListItems(strKey).Selected = True
'
'          'De-select all the available items.
'          For Each lst In ListView1.ListItems
'            lst.Selected = False
'          Next lst
'        End If
'
'      End If
'    End If
'
'      ' RH 27/09/00 - BUG 1017
''If ListView1.ListItems.Count > 0 Then ListView1.ListItems(ListView1.ListItems.Count).Selected = True
'      'ListItems(ListView1.ListItems.Count).Selected = True
'
'  End With
'
'  Set objExpr = Nothing
'
'  ' Reselect the cols/calcs that were selected before the calc button was pressed
'  For iCount = 0 To (UBound(SelectedKeys) - 1)
'    For Each lst In ListView1.ListItems
'      If lst.Key = SelectedKeys(iCount) Then
'        lst.Selected = True
'        Exit For
'      End If
'    Next lst
'  Next iCount
'
'  ListView1.SetFocus
'
'  Exit Sub
'
'NewCalc_ERROR:
'
'  Select Case Err.Number
'
'    Case 35601:  ' Expression could not be selected because the copy was aborted - hidden calc
'                 ' selected, but user not the definition owner.
'    Case Else: MsgBox "Error : " & Err.Description, vbExclamation + vbOKOnly, App.Title
'
'  End Select
'
'  Resume Next
'
'
'End Sub

Private Sub RefreshSelectedCalcNames()

  On Error GoTo ErrTrap
  Dim lst As ListItem
  
  For Each lst In ListView2.ListItems
    If Left(lst.Key, 1) = "E" Then
      lst.Text = datGeneral.GetExpression(Right(lst.Key, Len(lst.Key) - 1))
      If lst.Text = vbNullString Then lst.Text = "<Invalid Calc>"
    End If
  Next lst
  
ErrTrap:
  
  Exit Sub

End Sub


Private Sub cmdNewRelation_Click()

  Dim objTemp As clsMatchRelation
  Dim objBookmark As Variant
  Dim lngBaseSelected() As Long
  Dim lngCriteriaSelected() As Long
  Dim lngCount As Long
  
  ReDim lngBaseSelected(colRelatedTables.Count) As Long
  ReDim lngCriteriaSelected(colRelatedTables.Count) As Long
  For Each objTemp In colRelatedTables
    lngBaseSelected(lngCount) = objTemp.Table1ID
    lngCriteriaSelected(lngCount) = objTemp.Table2ID
    lngCount = lngCount + 1
  Next

  frmMatchDefTable.NewRelation Me, lngBaseSelected(), lngCriteriaSelected()
    
  With frmMatchDefTable
  
    If Not .Cancelled Then

      Set objTemp = frmMatchDefTable.MatchRelation
      'grdRelations.Columns("TableID").CellText(objBookmark) = objTemp.TableID
      
      grdRelations.AddItem ( _
        .cboTables.ItemData(.cboTables.ListIndex) & vbTab & _
        .cboTables.List(.cboTables.ListIndex) & vbTab & _
        .cboMatchTables.ItemData(.cboMatchTables.ListIndex) & vbTab & _
        .cboMatchTables.List(.cboMatchTables.ListIndex) & vbTab & _
        .txtRequired.Text & vbTab & _
        .txtPreferred.Text & vbTab & _
        .txtScore.Text)
      grdRelations.SelBookmarks.Add grdRelations.GetBookmark(grdRelations.Rows - 1)

      objTemp.Table1ID = .cboTables.ItemData(.cboTables.ListIndex)
      objTemp.Table2ID = .cboMatchTables.ItemData(.cboMatchTables.ListIndex)
      objTemp.RequiredExprID = Val(.txtRequired.Tag)
      objTemp.PreferredExprID = Val(.txtPreferred.Tag)
      objTemp.MatchScoreID = Val(.txtScore.Tag)
      objTemp.BreakdownColumns = .MatchRelation.BreakdownColumns

      colRelatedTables.Add objTemp, "T" & CStr(objTemp.Table1ID)

      Changed = True
      UpdateButtonStatus (SSTab1.Tab)
    End If

  End With

  Set objBookmark = Nothing
  Set frmMatchDefTable = Nothing

End Sub

Private Sub cmdEditRelation_Click()

  Dim objTemp As clsMatchRelation
  Dim strKey As String
  Dim strTempName As String
  Dim lngRow As Long
  Dim lngCount As Long
  Dim lngBaseSelected() As Long
  Dim lngCriteriaSelected() As Long
  
  strKey = "T" & grdRelations.Columns("TableID").CellValue(grdRelations.Bookmark)
  
  ReDim lngBaseSelected(colRelatedTables.Count) As Long
  ReDim lngCriteriaSelected(colRelatedTables.Count) As Long
  For Each objTemp In colRelatedTables
    lngBaseSelected(lngCount) = objTemp.Table1ID
    lngCriteriaSelected(lngCount) = objTemp.Table2ID
    lngCount = lngCount + 1
  Next

  Set objTemp = colRelatedTables.Item(strKey)
  
  frmMatchDefTable.EditRelation Me, lngBaseSelected(), lngCriteriaSelected(), objTemp, mblnReadOnly
  
  With frmMatchDefTable
  
    If .ChangedName Or Not .Cancelled Then
  
      lngRow = grdRelations.AddItemRowIndex(grdRelations.Bookmark)
      
      grdRelations.RemoveItem lngRow
      grdRelations.AddItem .cboTables.ItemData(.cboTables.ListIndex) & vbTab & _
        .cboTables.List(.cboTables.ListIndex) & vbTab & _
        .cboMatchTables.ItemData(.cboMatchTables.ListIndex) & vbTab & _
        .cboMatchTables.List(.cboMatchTables.ListIndex) & vbTab & _
        .txtRequired.Text & vbTab & _
        .txtPreferred.Text & vbTab & _
        .txtScore.Text, lngRow
      grdRelations.Bookmark = grdRelations.AddItemBookmark(lngRow)
      grdRelations.SelBookmarks.Add grdRelations.Bookmark

      Set objTemp = New clsMatchRelation
      objTemp.Table1ID = .cboTables.ItemData(.cboTables.ListIndex)
      objTemp.Table2ID = .cboMatchTables.ItemData(.cboMatchTables.ListIndex)
      objTemp.RequiredExprID = Val(.txtRequired.Tag)
      objTemp.PreferredExprID = Val(.txtPreferred.Tag)
      objTemp.MatchScoreID = Val(.txtScore.Tag)
      objTemp.BreakdownColumns = .MatchRelation.BreakdownColumns
      
      colRelatedTables.Remove strKey
      colRelatedTables.Add objTemp, "T" & CStr(objTemp.Table1ID)

      Changed = True
      UpdateButtonStatus (SSTab1.Tab)
    End If
  
  End With
  
  Set frmMatchDefTable = Nothing

End Sub

Private Sub cmdOK_Click()

  If Changed = True Then
    Screen.MousePointer = vbHourglass

    If ValidateDefinition(mlngMatchReportID) Then
      If SaveDefinition Then
        mblnCancelled = False
        Me.Hide
      End If
    End If

  End If

  Screen.MousePointer = vbNormal

End Sub

Private Sub cmdRemove_LostFocus()
  'JPD 20031013 Fault 5827
  cmdRemove.Picture = cmdRemove.Picture

End Sub

Private Sub cmdRemoveAll_LostFocus()
  'JPD 20031013 Fault 5827
  cmdRemoveAll.Picture = cmdRemoveAll.Picture

End Sub

Private Sub Form_Activate()
  SSTab1_Click 0
  mblnCancelled = True
  
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture

End Sub

Private Sub Form_Load()

  SSTab1.Tab = 0
  ReDim mvarHiddenCount(2, 0)
  
  Set datData = New clsDataAccess
  Set colRelatedTables = New Collection
  Set mcolMatchReportColDetails = New Collection
  
  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos True, True, True
  
  ReDim mlngExprDeleteOnOK(2, 0)
  ReDim mlngExprDeleteOnCancel(2, 0)

  fraFieldButtons.BackColor = Me.BackColor

  ''MH20030507 TEMPORARY 3 LINES!
  'optHighest.Value = True
  'optLowest.Enabled = False
  'chkMinimum.Enabled = False

  grdAccess.RowHeight = 239

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer
    
    If Changed = True And Not FormPrint Then
      
      pintAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, Me.Caption)
        
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


Private Sub grdRelations_Click()
  grdRelations.SelBookmarks.RemoveAll
  grdRelations.SelBookmarks.Add grdRelations.Bookmark
  If cmdEditRelation.Enabled = False Then
    UpdateButtonStatus (SSTab1.Tab)
  End If
End Sub

Private Sub grdRelations_DblClick()
  If Not mblnReadOnly Then
    If grdRelations.Rows > 0 Then
      cmdEditRelation_Click
    Else
      cmdNewRelation_Click
    End If
  End If
End Sub

Private Sub grdRelations_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  If Not cmdEditRelation.Enabled Then
    UpdateButtonStatus (SSTab1.Tab)
  End If
End Sub

Private Sub grdReportOrder_AfterColUpdate(ByVal ColIndex As Integer)

  Changed = True
  
End Sub

Private Sub grdReportOrder_Change()

  'CheckGridBreaks

  'CheckRepeatOptions
  
  Changed = True
  
End Sub

Private Sub grdReportOrder_Click()

'  'TM20010821 Fault 2379
'  'Sets the original selection of the break.
'  'Then 'sOriginalSelection' is used in the CheckGridBreaks function.
'  If grdReportOrder.Columns("Break").Value Then
'    sOriginalSelection = "Break"
'  Else
'    sOriginalSelection = "Page"
'  End If

End Sub

Private Sub grdReportOrder_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  
  If Not mblnReadOnly Then
  
    With grdReportOrder
'      ' Set the styleSet of the rows to show which is selected.
'      For iLoop = 0 To .Rows
'        If iLoop = .Row Then
'          .Columns(1).CellStyleSet "ssetActive", iLoop
'        Else
'          .Columns(1).CellStyleSet "ssetDormant", iLoop
'        End If
'      Next iLoop
'
'      ' Activate the 'values' column.
      
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      
      If .Col = 1 Then
        .Col = 0
      End If

    End With
  
    With grdReportOrder
      If .AddItemRowIndex(.Bookmark) = 0 Then
        cmdMoveUpOrder.Enabled = False
        cmdMoveDownOrder.Enabled = .Rows > 1
      ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
        cmdMoveUpOrder.Enabled = .Rows > 1
        cmdMoveDownOrder.Enabled = False
      Else
        cmdMoveUpOrder.Enabled = True
        cmdMoveDownOrder.Enabled = True
      End If
    End With
  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ListView1_GotFocus()
  cmdAdd.Default = True
End Sub

Private Sub ListView1_LostFocus()
  cmdOk.Default = True
End Sub

Private Sub ListView2_GotFocus()
  cmdRemove.Default = True
End Sub

Private Sub ListView2_LostFocus()
  cmdOk.Default = True
End Sub

Private Sub optAllRecords_Click(Index As Integer)

  Changed = True

  With txtPicklist(Index)
    .Text = ""
    .Tag = 0
  End With
  
  cmdPicklist(Index).Enabled = False

  With txtFilter(Index)
    .Text = ""
    .Tag = 0
  End With

  cmdFilter(Index).Enabled = False
  
  If Index = 0 Then
    chkPrintFilterHeader.Enabled = False
    chkPrintFilterHeader.Value = vbUnchecked
  End If
  
  'JPD 20040628 Fault 8818
  If Not mblnLoading Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
End Sub

Private Sub optHighest_Click()
  chkLimit.Caption = "Minimum Score"
  Changed = True
End Sub

Private Sub optLowest_Click()
  chkLimit.Caption = "Maximum Score"
  Changed = True
End Sub

Private Sub optFilter_Click(Index As Integer)

  Changed = True

  cmdFilter(Index).Enabled = True

  With txtPicklist(Index)
    .Text = ""
    .Tag = 0
  End With

  txtFilter(Index).Text = "<None>"
  cmdPicklist(Index).Enabled = False

  If Index = 0 Then
    chkPrintFilterHeader.Enabled = True
  End If

  ForceDefinitionToBeHiddenIfNeeded

End Sub

Private Sub optPicklist_Click(Index As Integer)

  Changed = True

  cmdPicklist(Index).Enabled = True

  With txtFilter(Index)
    .Text = ""
    .Tag = 0
  End With

  txtPicklist(Index).Text = "<None>"
  cmdFilter(Index).Enabled = False
  
  If Index = 0 Then
    chkPrintFilterHeader.Enabled = True
  End If
  
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub

Private Sub spnLimit_Change()
  Changed = True
End Sub

Private Sub spnMaxRecords_Change()

  lblAllRecords.Visible = (spnMaxRecords.Value = 0)
  Changed = True

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  Dim ctl As Control

  If Not mblnReadOnly Then
    For Each ctl In Me.Controls
      If TypeOf ctl Is VB.Frame Then
        ctl.Enabled = ctl.Left >= 0
      End If
    Next
    UpdateButtonStatus SSTab1.Tab
  End If

End Sub


Private Sub cmdPicklist_Click(Index As Integer)

  If Index = 0 Then
    GetPicklist cboTable1, txtPicklist(Index)
  Else
    GetPicklist cboTable2, txtPicklist(Index)
  End If

End Sub

Private Sub cmdFilter_Click(Index As Integer)
  
  If Index = 0 Then
    GetFilter cboTable1, txtFilter(Index)
  Else
    GetFilter cboTable2, txtFilter(Index)
  End If

End Sub


'# DRAG AND DROP CODE ###########################################

Private Sub ListView1_DblClick()

  If mblnReadOnly Then
    Exit Sub
  End If

  ' Copy the item doubleclicked on to the 'Selected' Listview
  CopyToSelected False

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

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Capture the Ctl-A event. This does not trigger the itemclick
  ' event so have to force an updatebuttonstatus here
  
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
Private Sub ListView2_DblClick()

  If mblnReadOnly Then
    Exit Sub
  End If
  
  ' Remove the item doubleclicked on from the 'Selected' Listview
  CopyToAvailable False
  
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If mblnReadOnly Then
    Exit Sub
  End If
  
  ' Popup menu on right button.
  If Button = vbRightButton Then
    
    With ActiveBar1.Bands("ColSelPopup")
      ' Enable/disable the required tools.
      .Tools("ID_Add").Enabled = ((Not mblnReadOnly) And (cmdAdd.Enabled))
      .Tools("ID_AddAll").Enabled = ((Not mblnReadOnly) And (cmdAddAll.Enabled))
      .Tools("ID_Remove").Enabled = False
      .Tools("ID_RemoveAll").Enabled = False
      .Tools("ID_MoveUp").Enabled = False
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

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)

'TM20010822 Fault 2566
'Need to enter the 'UpdateButtonStatus' routine to populate the count,heading etc. controls.
'  If mblnReadOnly Then
'    Exit Sub
'  End If
  
  UpdateButtonStatus (SSTab1.Tab)

End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If mblnReadOnly Then
    Exit Sub
  End If

  ' Popup menu on right button.
  If Button = vbRightButton Then
    
    With ActiveBar1.Bands("ColSelPopup")
      ' Enable/disable the required tools.
      .Tools("ID_Add").Enabled = False
      .Tools("ID_AddAll").Enabled = False
      .Tools("ID_Remove").Enabled = ((Not mblnReadOnly) And (cmdRemove.Enabled))
      .Tools("ID_RemoveAll").Enabled = ((Not mblnReadOnly) And (cmdRemoveAll.Enabled))
      .Tools("ID_MoveUp").Enabled = ((Not mblnReadOnly) And (cmdMoveUp.Enabled))
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
  
  End Select

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If mblnReadOnly Then
    Exit Sub
  End If
  
  ' Start the drag operation
  Dim objItem As ComctlLib.ListItem

  If Button = vbLeftButton Then
    If ListView1.ListItems.Count > 0 Then
      mblnColumnDrag = True
      ListView1.Drag vbBeginDrag
    End If
  End If

End Sub

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
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

Private Sub ListView1_DragDrop(Source As Control, x As Single, y As Single)
  
  ' Perform the drop operation
  If Source Is ListView2 Then
    CopyToAvailable False
    ListView2.Drag vbCancel
  Else
    ListView2.Drag vbCancel
  End If

End Sub

Private Sub ListView2_DragDrop(Source As Control, x As Single, y As Single)
  
  ' Perform the drop operation - action depends on source and destination
  
  If Source Is ListView1 Then
    If ListView2.HitTest(x, y) Is Nothing Then
      CopyToSelected False
    Else
      CopyToSelected False, ListView2.HitTest(x, y).Index
    End If
    ListView1.Drag vbCancel
  Else
    If ListView2.HitTest(x, y) Is Nothing Then
      ChangeSelectedOrder
    Else
      ChangeSelectedOrder ListView2.HitTest(x, y).Index
    End If
    ListView2.Drag vbCancel
  End If

End Sub

Private Sub Frafieldsavailable_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture
  
End Sub

Private Sub Frafieldsselected_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture
  
End Sub

Private Sub ListView2_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument.Picture
  End If

  ' Set DropHighlight to the mouse's coordinates.
  Set ListView2.DropHighlight = ListView2.HitTest(x, y)

End Sub

Private Sub ListView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument.Picture
  End If

End Sub

Private Function CopyToSelected(bAll As Boolean, Optional intBeforeIndex As Integer)

  ' Copy items to the 'Selected' listview
  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  Dim fOK As Boolean
  Dim iItemSelectedCount As Integer

  Dim objCalcExpr As clsExprExpression
  
  Dim objTempItem As ListItem
  Dim objWorkingItem As ListItem
  Dim iSelectedCount As Integer
  Dim iItemToSelect As Integer
  Dim prstTemp As Recordset
  Dim iItemsToDelete() As Variant
  ReDim iItemsToDelete(0)
  Dim intTemp As Integer
  
  Dim sTempTableName As String
  Dim lngColumnID As Long
  Dim lngTableID As Long
  Dim bCheckIfHidden As Boolean
  
  bCheckIfHidden = False
  Screen.MousePointer = vbHourglass

  'If user has clicked ADD ALL then do this...
  If bAll Then
    For Each objTempItem In ListView1.ListItems
      If Left(objTempItem.Key, 1) = "C" Then
        '*******************************************************************
        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
        lngTableID = GetTableIDFromColumn(lngColumnID)
        sTempTableName = GetTableNameFromColumn(lngColumnID)
        
        'If (lngTableID = cboTable1.ItemData(cboTable1.ListIndex)) _
        '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
        '  grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
        'End If
        '*******************************************************************

        ListView2.ListItems.Add , objTempItem.Key, GetTableNameFromColumn(Right(objTempItem.Key, Len(objTempItem.Key) - 1)) & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
        fOK = True
      Else
        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
        
        'Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT A.Access, A.TableID, B.TableName " & _
                                                          "FROM AsrSysExpressions A " & _
                                                          "     INNER JOIN ASRSysTables B " & _
                                                          "     ON A.TableID = B.TableID " & _
                                                          "WHERE A.ExprID = " & lngColumnID)
        'If Not prstTemp.BOF And Not prstTemp.EOF Then
        '  If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
        '    MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
        '          " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, Me.Caption
        '    fOK = False
        '  Else
            fOK = True
            ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon

        '    lngTableID = prstTemp!TableID
            'If lngTableID = cboTable1.ItemData(cboTable1.ListIndex) _
            '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
            '  grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
            'End If
        '  End If
        'End If
      End If
      If fOK Then AddToCollection2 objTempItem
    Next objTempItem
    
    ListView1.ListItems.Clear
    SelectFirst ListView2
    UpdateButtonStatus (SSTab1.Tab)
    ForceDefinitionToBeHiddenIfNeeded
    Screen.MousePointer = vbNormal
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
    If intBeforeIndex = 0 Then
      
      If Left(objTempItem.Key, 1) = "C" Then
        '*******************************************************************
        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
        lngTableID = GetTableIDFromColumn(lngColumnID)
        sTempTableName = GetTableNameFromColumn(lngColumnID)
        
        'If (lngTableID = cboTable1.ItemData(cboTable1.ListIndex)) _
        '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
        '  grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
        'End If
        '*******************************************************************
        
        ListView2.ListItems.Add , objTempItem.Key, sTempTableName & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
        fOK = True
      Else
        bCheckIfHidden = True
'        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
'
'        Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT A.Access, A.TableID, B.TableName " & _
'                                                          "FROM AsrSysExpressions A " & _
'                                                          "     INNER JOIN ASRSysTables B " & _
'                                                          "     ON A.TableID = B.TableID " & _
'                                                          "WHERE A.ExprID = " & lngColumnID)
'
'        If prstTemp.BOF And prstTemp.EOF Then
'          MsgBox "The selected calculation has been deleted.", vbExclamation + vbOKOnly, "Match Reports"
'          fOK = True
'        ElseIf prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
'          MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
'                " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Match Reports"
'          fOK = False
'        Else
          ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
          fOK = True
'
'          lngTableID = prstTemp!TableID
'          'If lngTableID = cboTable1.ItemData(cboTable1.ListIndex) _
'          '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
'          '  grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
'          'End If
'        End If
      
      End If
        
      If fOK Then
        AddToCollection2 objTempItem
        ListView1.ListItems.Remove iItemToSelect
      End If
    
   Else
  
      ' Before index
      If Left(objTempItem.Key, 1) = "C" Then
        '*******************************************************************
        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
        lngTableID = GetTableIDFromColumn(lngColumnID)
        sTempTableName = GetTableNameFromColumn(lngColumnID)
        
        'If (lngTableID = cboTable1.ItemData(cboTable1.ListIndex)) _
        '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
        '  grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
        'End If
        '*******************************************************************
        
        ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, sTempTableName & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
        fOK = True
      Else
        bCheckIfHidden = True
        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
    
        Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT A.Access, A.TableID, B.TableName " & _
                                                          "FROM AsrSysExpressions A " & _
                                                          "     INNER JOIN ASRSysTables B " & _
                                                          "     ON A.TableID = B.TableID " & _
                                                          "WHERE A.ExprID = " & lngColumnID)
        
        If prstTemp.BOF And prstTemp.EOF Then
          MsgBox "The selected calculation has been deleted.", vbExclamation + vbOKOnly, Me.Caption
          fOK = True
        ElseIf prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
          MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, Me.Caption
          fOK = False
        Else
          ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
          fOK = True

          lngTableID = prstTemp!TableID
          'If lngTableID = cboTable1.ItemData(cboTable1.ListIndex) _
          '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
          '  grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
          'End If

        End If
      
      End If
        
      If fOK Then
        AddToCollection2 objTempItem
        ListView1.ListItems.Remove iItemToSelect
      End If
      
   End If

    If ListView1.ListItems.Count > 0 Then
      If iItemToSelect > ListView1.ListItems.Count Then
        iItemToSelect = ListView1.ListItems.Count
      End If
      ListView1.ListItems(iItemToSelect).Selected = True
    End If
    
    If intBeforeIndex = 0 Then
      SelectLast ListView2
    Else
      For Each objTempItem In ListView2.ListItems
        objTempItem.Selected = (objTempItem.Index = intBeforeIndex)
      Next objTempItem
      Set ListView2.DropHighlight = Nothing
    End If
    
    UpdateButtonStatus (SSTab1.Tab)
    If bCheckIfHidden Then
      ForceDefinitionToBeHiddenIfNeeded
    End If
    Screen.MousePointer = vbNormal
    Changed = True
    Exit Function
  End If

  'There are more than one item selected
  For Each objTempItem In ListView1.ListItems
    
    If objTempItem.Selected Then
    
      If intBeforeIndex = 0 Then
        
        If Left(objTempItem.Key, 1) = "C" Then
          '*******************************************************************
          lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
          lngTableID = GetTableIDFromColumn(lngColumnID)
          sTempTableName = GetTableNameFromColumn(lngColumnID)
          
          'If (lngTableID = cboTable1.ItemData(cboTable1.ListIndex)) _
          '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
          '  grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
          'End If
          '*******************************************************************
          
          ListView2.ListItems.Add , objTempItem.Key, sTempTableName & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
          fOK = True
        Else
          lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)

          Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT A.Access, A.TableID, B.TableName " & _
                                                          "FROM AsrSysExpressions A " & _
                                                          "     INNER JOIN ASRSysTables B " & _
                                                          "     ON A.TableID = B.TableID " & _
                                                          "WHERE A.ExprID = " & lngColumnID)
          If prstTemp.BOF And prstTemp.EOF Then
            MsgBox "One or more of the selected calculation(s) have been deleted.", vbExclamation + vbOKOnly, Me.Caption
            fOK = False
          End If
          If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
            MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                  " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, Me.Caption
            fOK = False
          Else
            ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
            fOK = True

            lngTableID = prstTemp!TableID
            'If lngTableID = cboTable1.ItemData(cboTable1.ListIndex) _
            '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
            '  grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
            'End If

          End If
        End If
    
      Else
      
        ' Before an existing item
          If Left(objTempItem.Key, 1) = "C" Then

            '*******************************************************************
            lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
            lngTableID = GetTableIDFromColumn(lngColumnID)
            sTempTableName = GetTableNameFromColumn(lngColumnID)
            
            'If (lngTableID = cboTable1.ItemData(cboTable1.ListIndex)) _
            '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
            '  grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
            'End If
            '*******************************************************************

            ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, sTempTableName & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
            fOK = True
          Else
            lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)

            Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT A.Access, A.TableID, B.TableName " & _
                                                          "FROM AsrSysExpressions A " & _
                                                          "     INNER JOIN ASRSysTables B " & _
                                                          "     ON A.TableID = B.TableID " & _
                                                          "WHERE A.ExprID = " & lngColumnID)
            If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
              MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                     " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, Me.Caption
              fOK = False
            Else
              fOK = True
              ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon

              lngTableID = prstTemp!TableID
              'If lngTableID = cboTable1.ItemData(cboTable1.ListIndex) _
              '  Or (lngTableID = TxtParent1.Tag) Or (lngTableID = TxtParent2.Tag) Then
              '  grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
              'End If

            End If
          End If
          intBeforeIndex = intBeforeIndex + 1
      
      End If
    
      If fOK = True Then
        AddToCollection2 objTempItem
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
  
  If intBeforeIndex = 0 Then
    SelectLast ListView2
  Else
    For Each objTempItem In ListView2.ListItems
      objTempItem.Selected = (objTempItem.Index = intBeforeIndex)
    Next objTempItem
    Set ListView2.DropHighlight = Nothing
  End If

  UpdateButtonStatus (SSTab1.Tab)
  ForceDefinitionToBeHiddenIfNeeded
  Screen.MousePointer = vbNormal
  Changed = True

'
'  For iLoop = ListView1.ListItems.Count To 1 Step -1
'
'    'If we are not inserting it before existing columns...
'    If intBeforeIndex = 0 Then
'
'      If ListView1.ListItems(iLoop).Selected Then
'
'        If Left(ListView1.ListItems(iLoop).Key, 1) = "C" Then
'          ListView2.ListItems.Add , ListView1.ListItems(iLoop).Key, GetTableNameFromColumn(Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1)) & "." & ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'          fOK = True
'        Else
'          Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1))
'          If prstTemp.BOF And prstTemp.EOF Then
'            MsgBox "One or more of the selected calculation(s) have been deleted.", vbExclamation + vbOKOnly, "Match Reports"
'            Exit For
'          End If
'          If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
'            MsgBox "Cannot include the '" & ListView1.ListItems(iLoop).Text & "' calculation." & vbCrLf & _
'                  " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Match Reports"
'            fOK = False
'          Else
'            fOK = True
'            ListView2.ListItems.Add , ListView1.ListItems(iLoop).Key, ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'          End If
'        End If

'        If intBeforeIndex Then
'          If Left(ListView1.ListItems(iLoop).Key, 1) = "C" Then
'            ListView2.ListItems.Add intBeforeIndex, ListView1.ListItems(iLoop).Key, GetTableNameFromColumn(Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1)) & "." & ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'            fOK = True
'          Else
'            Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1))
'            If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
'              MsgBox "Cannot include the '" & ListView1.ListItems(iLoop).Text & "' calculation." & vbCrLf & _
'                     " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Match Reports"
'              fOK = False
'            Else
'              fOK = True
'              ListView2.ListItems.Add intBeforeIndex, ListView1.ListItems(iLoop).Key, ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'            End If
'          End If
'          intBeforeIndex = intBeforeIndex + 1
'        Else
'
'          If Left(ListView1.ListItems(iLoop).Key, 1) = "C" Then
'            ListView2.ListItems.Add , ListView1.ListItems(iLoop).Key, GetTableNameFromColumn(Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1)) & "." & ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'            fOK = True
'          Else
'            Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1))
'            If prstTemp.BOF And prstTemp.EOF Then
'              MsgBox "One or more of the selected calculation(s) have been deleted.", vbExclamation + vbOKOnly, "Match Reports"
'              Exit For
'            End If
'            If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
'              MsgBox "Cannot include the '" & ListView1.ListItems(iLoop).Text & "' calculation." & vbCrLf & _
'                    " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Match Reports"
'              fOK = False
'            Else
'              fOK = True
'              ListView2.ListItems.Add , ListView1.ListItems(iLoop).Key, ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'            End If
'          End If
'
'        End If
'
'        If fOK Then AddToCollection iLoop
'
'        ListView1.ListItems.Remove ListView1.ListItems(iLoop).Key
'
'      End If`
'    Else
'
'      If Left(ListView1.ListItems(iLoop).Key, 1) = "C" Then
'        ListView2.ListItems.Add 1, ListView1.ListItems(iLoop).Key, GetTableNameFromColumn(Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1)) & "." & ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'        fOK = True
'      Else
'
'        Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1))
'        If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
'              MsgBox "Cannot include the '" & ListView1.ListItems(iLoop).Text & "' calculation." & vbCrLf & _
'                    " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Match Reports"
'          fOK = False
'        Else
'          fOK = True
'          ListView2.ListItems.Add 1, ListView1.ListItems(iLoop).Key, ListView1.ListItems(iLoop).Text, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon
'        End If
'      End If
'
'      If fOK Then AddToCollection iLoop
'
'      ListView1.ListItems.Remove ListView1.ListItems(iLoop).Key
'
'    End If
'
'  Next iLoop
'
'  If ListView1.ListItems.Count > 0 Then
'    If iItemSelectedCount > 1 Then
'      ListView1.ListItems(1).Selected = True
'    Else
'      If iItemToSelect > ListView1.ListItems.Count Then iItemToSelect = ListView1.ListItems.Count
'      ListView1.ListItems(iItemToSelect).Selected = True
'    End If
'  End If
'
'  'PopulateTableAvailable cboTblAvailable.Text
'  If bAll Then
'    SelectFirst ListView2
'  Else
'    SelectLast ListView2
''  End If
'
'  UpdateButtonStatus
'  ForceDefinitionToBeHiddenIfNeeded
'
'  Screen.MousePointer = vbNormal
'
'  Changed = True
  
End Function

Private Function CheckInRepetitionGrid(pstrKey As String) As Boolean

'  ' Loop through the sort order grid, checking if the specified column is
'  ' defined in the sort order.
'  Dim pvarbookmark As Variant
'  Dim pintLoop As Integer
'
'  With grdRepetition
'    .MoveFirst
'    Do Until pintLoop = .Rows
'      pvarbookmark = .GetBookmark(pintLoop)
'      If .Columns("ColumnID").CellText(pvarbookmark) = pstrKey Then
'        CheckInRepetitionGrid = True
'        Exit Function
'      End If
'      pintLoop = pintLoop + 1
'    Loop
'  End With
'
'  CheckInRepetitionGrid = False

End Function

Private Function CheckInSortOrder(plngKey As Long) As Boolean

  ' Loop through the sort order grid, checking if the specified column is
  ' defined in the sort order.
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  
  With grdReportOrder
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("ColumnID").CellText(pvarbookmark) = plngKey Then
        CheckInSortOrder = True
        Exit Function
      End If
      pintLoop = pintLoop + 1
    Loop
  End With
  
  CheckInSortOrder = False

End Function

Private Sub RemoveFromSortOrder(plngKey As Long)

  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  
  ' Remove the specified column from the sort order grid.
  With grdReportOrder
    .MoveFirst
    
    If .Rows = 1 Then
      .RemoveAll
      Exit Sub
    End If
    
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("ColumnID").CellText(pvarbookmark) = plngKey Then
        .RemoveItem pintLoop
        Exit Sub
      End If
      pintLoop = pintLoop + 1
    Loop
  End With
  
End Sub

Private Sub RemoveFromRepetition(pstrKey As String)

'  Dim pintLoop As Integer
'  Dim pvarbookmark As Variant
'
'  ' Remove the specified column from the sort order grid.
'  With grdRepetition
'    .MoveFirst
'
'    Do Until pintLoop = .Rows
'      pvarbookmark = .GetBookmark(pintLoop)
'      If .Columns("ColumnID").CellText(pvarbookmark) = pstrKey Then
'        .RemoveItem pintLoop
'
'        .SelBookmarks.RemoveAll
'        .MoveFirst
'        Exit Sub
'      End If
'      pintLoop = pintLoop + 1
'    Loop
'  End With
  
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
        If CheckInSortOrder(Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)) = True Then
          If MsgBox("Removing the following column will also remove it from the report sort order." & vbCrLf & vbCrLf & ListView2.ListItems(iLoop).Text & vbCrLf & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            iTempItemIndex = iLoop
'            If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
'              RemoveFromRepetition ListView2.ListItems(iLoop).Key
'            End If
            RemoveFromSortOrder Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)
            mcolMatchReportColDetails.Remove ListView2.ListItems(iLoop).Key
            ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
          End If
        Else
          iTempItemIndex = iLoop
'          If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
'            RemoveFromRepetition ListView2.ListItems(iLoop).Key
'          End If
          mcolMatchReportColDetails.Remove ListView2.ListItems(iLoop).Key
          ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
        End If
      End If
    Else
      If CheckInSortOrder(Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)) = True Then
'        If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
'          RemoveFromRepetition ListView2.ListItems(iLoop).Key
'        End If
        RemoveFromSortOrder Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)
        mcolMatchReportColDetails.Remove ListView2.ListItems(iLoop).Key
        ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
      Else
'        If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
'          RemoveFromRepetition ListView2.ListItems(iLoop).Key
'        End If
        mcolMatchReportColDetails.Remove ListView2.ListItems(iLoop).Key
        ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
      End If
      
    End If
  Next iLoop
  
  If ListView2.ListItems.Count > 0 Then
    If iTempItemIndex > ListView2.ListItems.Count Then iTempItemIndex = ListView2.ListItems.Count
    If iTempItemIndex > 0 Then ListView2.ListItems(iTempItemIndex).Selected = True
  End If
  
  PopulateAvailable
  
  UpdateButtonStatus (SSTab1.Tab)
  UpdateOrderButtons
  ForceDefinitionToBeHiddenIfNeeded

  Changed = True

  Screen.MousePointer = vbNormal

End Function


Private Function ChangeSelectedOrder(Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)

  ' This function changes the order of listitems in the selected listview.
  ' At the moment, different arrays are used depending on what information you
  ' need to store...change the array to a type if it would suit the purpose
  ' better
  
  ' Dimension arrays
  Dim iLoop As Integer, Key() As String, Text() As String, Icon() As Variant, SmallIcon() As Variant
  ReDim Key(0), Text(0), Icon(0), SmallIcon(0)
  
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

End Function


Private Function UpdateButtonStatus(iTab As Integer)
'
  On Error Resume Next

  Dim tempItem As ListItem, iCount As Integer

''  If Not mblnReadOnly Then
  Select Case iTab
  Case 1
    If grdRelations.Rows = 0 Then
      cmdNewRelation.Enabled = fraRelations.Enabled
      cmdEditRelation.Enabled = False
      cmdDeleteRelation.Enabled = False
      cmdClearRelations.Enabled = False
    Else
      With grdRelations
        If .SelBookmarks.Count = 0 Then
          .Bookmark = .AddItemBookmark(0)
          .SelBookmarks.Add .Bookmark
        End If
      End With
      cmdEditRelation.Enabled = Not mblnReadOnly
      cmdEditRelation.SetFocus
      cmdDeleteRelation.Enabled = Not mblnReadOnly
      cmdClearRelations.Enabled = Not mblnReadOnly
    End If

  Case 2
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
    End If

    ' If there are more than 1 items selected then disable the move buttons and exit
    For Each tempItem In ListView2.ListItems
      If tempItem.Selected Then iCount = iCount + 1
    Next tempItem

    'Debug.Print Now() & vbTab & " - " & icount

    'If iCount > 1 Then
    If iCount > 1 Or iCount < 1 Then
      cmdMoveUp.Enabled = False
      cmdMoveDown.Enabled = False
      EnableColProperties False
    Else
      If ListView2.SelectedItem.Index <> 1 Then cmdMoveUp.Enabled = Not mblnReadOnly Else cmdMoveUp.Enabled = False
      If ListView2.SelectedItem.Index <> ListView2.ListItems.Count Then cmdMoveDown.Enabled = Not mblnReadOnly Else cmdMoveDown.Enabled = False
      EnableColProperties Not mblnReadOnly
    End If

    'Call CheckListViewColWidth(ListView1)
    'Call CheckListViewColWidth(ListView2)

  Case 3
    UpdateOrderButtons

  End Select

  Call CheckListViewColWidth(ListView1)
  Call CheckListViewColWidth(ListView2)

End Function

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

'###########

Private Sub cmdAdd_Click()
  CopyToSelected False
End Sub

Private Sub cmdMoveDown_Click()
  ChangeSelectedOrder ListView2.SelectedItem.Index + 2, True
End Sub

Private Sub cmdMoveUp_Click()
  ChangeSelectedOrder ListView2.SelectedItem.Index - 1
End Sub

Private Sub cmdRemove_Click()
  CopyToAvailable False
  If ListView2.ListItems.Count = 0 Then EnableColProperties False
End Sub

Private Sub cmdAddAll_Click()
  CopyToSelected True
End Sub

Private Sub cmdRemoveAll_Click()

  ' Remove All items from the 'Selected' Listview
  If grdReportOrder.Rows > 0 Then
    If MsgBox("Removing all selected report columns will also clear the report sort order." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
      CopyToAvailable True
      EnableColProperties False
    End If
  Else
    If MsgBox("Are you sure you wish to remove all columns / calculations from this definition ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
      CopyToAvailable True
      EnableColProperties False
    End If
  End If
  
End Sub

Private Function EnableColProperties(bStatus As Boolean)

  mblnLoading = True
  
  If Not ListView2.SelectedItem Is Nothing Then
    If ListView2.ListItems.Count > 0 Then GetCurrentDetails ListView2.SelectedItem.Key
  End If
  
  lblProp_ColumnHeading.Enabled = bStatus
  txtProp_ColumnHeading.Enabled = bStatus
  txtProp_ColumnHeading.BackColor = IIf(bStatus, &H80000005, &H8000000F)
  spnSize.Enabled = bStatus
  lblProp_Size.Enabled = bStatus
  spnSize.BackColor = IIf(bStatus, &H80000005, &H8000000F)
  'chkProp_Count.Enabled = bStatus
  
  spnDec.Enabled = bStatus
  spnDec.BackColor = IIf(bStatus, &H80000005, &H8000000F)
  
  lblProp_Decimals.Enabled = bStatus
    
  If chkProp_IsNumeric.Value = vbUnchecked Then
    spnDec.Enabled = False
    lblProp_Decimals.Enabled = spnDec.Enabled
    spnDec.BackColor = &H8000000F
  End If
  
  chkProp_IsNumeric.Value = vbUnchecked

  If ListView2.SelectedItem Is Nothing Then
    txtProp_ColumnHeading.Text = vbNullString
    spnSize.Value = 0
    spnDec.Value = 0
  End If
   
  mblnLoading = False

End Function

Private Function AddToCollection(iLoop As Integer) As Boolean

  Dim sColType As String
  Dim lID As Long
  Dim sHeading As String
  Dim lSize As Long
  Dim iDecPlaces As Integer
  Dim bAverage As Boolean
  Dim bCount As Boolean
  Dim bTotal As Boolean
  Dim bIsNumeric As Boolean
  
  mblnLoading = True
  GetDefaultDetails ListView1.ListItems(iLoop).Key
  mblnLoading = False
  
  sColType = Left(ListView1.ListItems(iLoop).Key, 1)
  lID = Right(ListView1.ListItems(iLoop).Key, Len(ListView1.ListItems(iLoop).Key) - 1)
  sHeading = txtProp_ColumnHeading.Text
  'lSize = txtProp_Size.Text
  lSize = spnSize.Text
  
  
  ' RH 20/09/00 - BUG 968 - Crashes when boolean col picked as first column
  'iDecPlaces = IIf(txtProp_DecPlaces.Text = "", "0", txtProp_DecPlaces.Text)
  iDecPlaces = IIf(spnDec.Text = "", "0", spnDec.Text)
  
  ' 190600 - Fault fix 358
  
  mblnLoading = True
  'chkProp_Average.Value = 0
  'chkProp_Count.Value = 0
  'chkProp_Total.Value = 0
  mblnLoading = False
  
  'bAverage = IIf(chkProp_Average.Value = vbChecked, True, False)
  'bCount = IIf(chkProp_Count.Value = vbChecked, True, False)
  'bTotal = IIf(chkProp_Total.Value = vbChecked, True, False)
  
  bIsNumeric = IIf(chkProp_IsNumeric.Value = vbChecked, True, False)
  
End Function

Private Function AddToCollection2(objTempItem As ListItem) As Boolean

  Dim sColType As String
  Dim lID As Long
  Dim sHeading As String
  Dim lSize As Long
  Dim iDecPlaces As Integer
  Dim bAverage As Boolean
  Dim bCount As Boolean
  Dim bTotal As Boolean
  Dim bIsNumeric As Boolean
  Dim objTemp As clsColumn
  
  Set objTemp = New clsColumn

  mblnLoading = True
  GetDefaultDetails objTempItem.Key
  mblnLoading = False
  
  'sColType = Left(objTempItem.Key, 1)
  'lID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
  'sHeading = txtProp_ColumnHeading.Text
  'lSize = spnSize.Text
  'iDecPlaces = IIf(spnDec.Text = "", "0", spnDec.Text)

  ' 190600 - Fault fix 358

  'mblnLoading = True
  'chkProp_Average.Value = 0
  'chkProp_Count.Value = 0
  'chkProp_Total.Value = 0
  'mblnLoading = False
  
  'bAverage = IIf(chkProp_Average.Value = vbChecked, True, False)
  'bCount = IIf(chkProp_Count.Value = vbChecked, True, False)
  'bTotal = IIf(chkProp_Total.Value = vbChecked, True, False)
  
  'bIsNumeric = IIf(chkProp_IsNumeric.Value = vbChecked, True, False)

  With objTemp
    .ColType = Left(objTempItem.Key, 1)
    .ID = Val(Mid(objTempItem.Key, 2))
    .Size = spnSize.Value
    .DecPlaces = spnDec.Value
    .IsNumeric = (chkProp_IsNumeric.Value = vbChecked)
    .Heading = objTempItem.Text
  End With
  
  mcolMatchReportColDetails.Add objTemp, objTempItem.Key
  

End Function


Private Function GetDefaultDetails(sKey As String) As Boolean

  ' This function returns the default Col/Expr Name, Size and
  ' Decimal Places. These can then be edited by the user if desired.
  
  Dim rsTemp As Recordset
  
  If Left(sKey, 1) = "C" Then
    
    Set rsTemp = datGeneral.GetColumnDefinition(Right(sKey, Len(sKey) - 1))
    
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      txtProp_ColumnHeading.Text = rsTemp!ColumnName
      
      'spnSize.Text = rsTemp!Size
      spnSize.Text = rsTemp!DefaultDisplayWidth
      
      If (rsTemp!DataType = sqlNumeric) Then ' its numeric
        spnDec.Text = rsTemp!Decimals
        chkProp_IsNumeric.Value = vbChecked
      ElseIf (rsTemp!DataType = sqlInteger) Then
        spnSize.Text = rsTemp!DefaultDisplayWidth ' 10 '5
        spnDec.Text = rsTemp!Decimals
        chkProp_IsNumeric.Value = vbChecked
      ElseIf rsTemp!DataType = sqlDate Then ' its a date
        spnSize.Text = rsTemp!DefaultDisplayWidth '10
        spnDec.Text = 0
        chkProp_IsNumeric.Value = vbUnchecked
      ElseIf rsTemp!DataType = sqlBoolean Then ' its a logic
        spnSize.Text = 1
        chkProp_IsNumeric.Value = vbUnchecked
      ElseIf rsTemp!DataType = sqlLongVarChar Then      ' working pattern field
        spnSize.Text = 14
      Else                                               ' its not
        spnDec.Text = 0
        chkProp_IsNumeric.Value = vbUnchecked
      End If
    End If
    
    If spnSize.Text = 0 Then
      If Len(rsTemp!SpinnerMaximum) > 0 Then spnSize.Text = Len(rsTemp!SpinnerMaximum)
    End If
    
  Else

    'Only calc in Match Reports is the Match Score.
    'Default this to numeric 6.2
    spnSize.Text = "0"
    spnDec.Text = "2"
    chkProp_IsNumeric.Value = vbChecked


'    Set rsTemp = datGeneral.GetRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & (Right(sKey, Len(sKey) - 1)))
'
'    If Not rsTemp.BOF And Not rsTemp.EOF Then
'      txtProp_ColumnHeading.Text = rsTemp!Name
'      spnSize.Text = rsTemp!ReturnSize
'
'      Dim objExpression As clsExprExpression
'      Set objExpression = New clsExprExpression
'      objExpression.ExpressionID = (Right(sKey, Len(sKey) - 1))
'      objExpression.ConstructExpression
'      objExpression.ValidateExpression True
'      If objExpression.ReturnType = 2 Then ' its numeric
'        spnDec.Text = rsTemp!ReturnDecimals
'        chkProp_IsNumeric.Value = vbChecked
'      Else                                               ' its not
'        spnDec.Text = 0
'        chkProp_IsNumeric.Value = vbUnchecked
'      End If
'    End If
'
'    Set objExpression = Nothing
    
  End If
  
End Function


Private Function GetCurrentDetails(sKey As String) As Boolean

  ' This function returns the details held in the collection
  ' for the currently highlighted item in the 'selected'
  ' listview
  
  Dim objTemp As clsColumn
  
  Set objTemp = mcolMatchReportColDetails.Item(sKey)
  
    If objTemp Is Nothing Then
    
      txtProp_ColumnHeading = ""
      'txtProp_Size = 0
      spnSize.Text = 0
      'txtProp_DecPlaces = 0
      spnDec.Text = 0
      'chkProp_Average.Value = vbUnchecked
      'chkProp_Count.Value = vbUnchecked
      'chkProp_Total.Value = vbUnchecked
      EnableColProperties False
      
    Else
  
      txtProp_ColumnHeading.Text = objTemp.Heading
      'txtProp_Size.Text = objTemp.Size
      spnSize.Text = objTemp.Size
      
      'txtProp_DecPlaces.Text = objTemp.DecPlaces
      spnDec.Text = objTemp.DecPlaces
      
      'chkProp_Average.Value = IIf(objTemp.Average, vbChecked, vbUnchecked)
      'chkProp_Count.Value = IIf(objTemp.Count, vbChecked, vbUnchecked)
      'chkProp_Total.Value = IIf(objTemp.Total, vbChecked, vbUnchecked)
      chkProp_IsNumeric.Value = IIf(objTemp.IsNumeric, vbChecked, vbUnchecked)
  
    End If
    
  Set objTemp = Nothing
    
End Function
Private Sub txtDesc_GotFocus()
  
  With txtDesc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
  cmdOk.Default = False
  
End Sub

Private Sub txtDesc_LostFocus()

  cmdOk.Default = True

End Sub

Private Sub txtEmailGroup_Change()
  Changed = True
End Sub

Private Sub txtEmailSubject_Change()
  Changed = True
End Sub

Private Sub txtEmailAttachAs_Change()
  Changed = True
End Sub

Private Sub txtFilename_Change()
  Changed = True
End Sub

Private Sub txtName_Change()
  Changed = True
End Sub

Private Sub txtDesc_Change()
  Changed = True
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub spnDec_Change()

  If spnDec.Text = "" Then spnDec.Text = "0"
  
  If Not mblnLoading Then
    Dim objItem As clsColumn
    Set objItem = mcolMatchReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.DecPlaces = Val(spnDec.Text)
    Set objItem = Nothing
    Changed = True
  End If
  
End Sub

Private Sub spnSize_Change()

  If spnSize.Text = "" Then spnSize.Text = "0"
  
  If Not mblnLoading Then
    Dim objItem As clsColumn
    Set objItem = mcolMatchReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Size = Val(spnSize.Text)
    Set objItem = Nothing
    Changed = True
  End If

End Sub

'Private Sub txtProp_Size_Change()
'
'  If Not mblnLoading Then
'    Dim objItem As clsColumn
'    Set objItem = mcolMatchReportColDetails.Item(ListView2.SelectedItem.Key)
'    objItem.Size = txtProp_Size.Text
'    Set objItem = Nothing
'    Changed = True
'  End If
'
'End Sub



Private Function ValidateCollection() As Boolean

  Dim intTemp As Integer
  Dim intTemp2 As Integer
  Dim intDupCount As Integer
  Dim pstrColumnsWithSizeZero As String
  
  ' First check the number of cols in the listview is the same as the
  ' number of items in the collection
  If ListView2.ListItems.Count <> mcolMatchReportColDetails.Count Then
    MsgBox "A serious error has occurred. To rectify, please remove all columns from the report definition and try again." & vbCrLf & "Please contact support stating : The no. of columns does not match the no of items in the collection.", vbCritical + vbOKOnly, Me.Caption
    SSTab1.Tab = 2
    Exit Function
  End If
  
  ' Now check that they are all unique
  For intTemp = 1 To ListView2.ListItems.Count
    intDupCount = 0
    If mcolMatchReportColDetails.Item(ListView2.ListItems(intTemp).Key).Heading = "" Then
      MsgBox "The '" & ListView2.ListItems(intTemp).Text & "' column has a blank column heading.", vbExclamation + vbOKOnly, Me.Caption
      SSTab1.Tab = 2
      Exit Function
    End If
    
    If Left(mcolMatchReportColDetails.Item(ListView2.ListItems(intTemp).Key).Heading, 3) = "?ID" Then
      MsgBox "The '" & ListView2.ListItems(intTemp).Text & "' column has a heading beginning '?ID'." & vbCrLf & _
                    "'?ID' is a reserved word and cannot be used at the beginning of a column heading.", vbExclamation + vbOKOnly, Me.Caption
      SSTab1.Tab = 2
      Exit Function
    End If
    
    For intTemp2 = 1 To ListView2.ListItems.Count
      If UCase(mcolMatchReportColDetails.Item(ListView2.ListItems(intTemp2).Key).Heading) = UCase(mcolMatchReportColDetails.Item(ListView2.ListItems(intTemp).Key).Heading) Then
        intDupCount = intDupCount + 1
      End If
    Next intTemp2
    
    If intDupCount > 1 Then
      MsgBox "One or more columns in your report have a heading of '" & mcolMatchReportColDetails.Item(ListView2.ListItems(intTemp).Key).Heading & "'" & vbCrLf & "Column headings must be unique.", vbExclamation + vbOKOnly, Me.Caption
      SSTab1.Tab = 2
      Exit Function
    End If
     
    'MH20010511 Allow zero size columns
    'If mcolMatchReportColDetails.Item(ListView2.ListItems(intTemp).Key).Size = 0 Then
    '  pstrColumnsWithSizeZero = pstrColumnsWithSizeZero & "        " & ListView2.ListItems(intTemp).Text & vbCrLf
    'End If
     
  Next intTemp
   
  'MH20010511 Allow zero size columns
  'If pstrColumnsWithSizeZero <> "" Then
  '  MsgBox "The following columns have a size of 0:" & vbCrLf & vbCrLf & pstrColumnsWithSizeZero & vbCrLf & _
  '         "Either allocate a size for these columns or remove them from the report.", vbExclamation + vbOKOnly, "Match Reports"
  '  SSTab1.Tab = 2
  '  Exit Function
  'End If
  
  ' Got here, so the column definition is fine too #######
  ValidateCollection = True

End Function

Private Function SaveDefinition() As Boolean

  On Error GoTo Save_ERROR
  
  Dim sSQL As String
  'Dim iLoop As Integer
  'Dim sKey As String
  Dim objCol As clsColumn
  Dim objRelation As clsMatchRelation
  Dim iDefExportTo As Integer
  Dim lngMatchRelationID As Long
  Dim lngCount As Long
  
  If mlngMatchReportID > 0 Then

    sSQL = "UPDATE ASRSYSMatchReportName SET " & _
        "Name = '" & Trim(Replace(txtName.Text, "'", "''")) & "', " & _
        "Description = '" & Replace(txtDesc.Text, "'", "''") & "', "
        
    sSQL = sSQL & _
        "MatchReportType = " & CStr(mlngMatchReportType) & ", " & _
        "ScoreMode = " & IIf(optHighest.Value = True, "0, ", "1, ") & _
        "NumRecords = " & CStr(spnMaxRecords.Value) & ", " & _
        "ScoreCheck = " & IIf(chkLimit.Value = vbChecked, "1, ", "0, ") & _
        "ScoreLimit = " & CStr(spnLimit.Value) & ", " & _
        "EqualGrade = " & IIf(chkEqualGrade.Value = vbChecked, "1, ", "0, ") & _
        "ReportingStructure = " & IIf(chkReportStructure.Value = vbChecked, "1, ", "0, ")

    sSQL = sSQL & _
        "Table1ID = " & CStr(cboTable1.ItemData(cboTable1.ListIndex)) & ", " & _
        "Table1AllRecords = " & IIf(optAllRecords(0).Value, "1", "0") & ", " & _
        "Table1Picklist = " & IIf(optPicklist(0).Value, txtPicklist(0).Tag, "0") & ", " & _
        "Table1Filter = " & IIf(optFilter(0).Value, txtFilter(0).Tag, "0") & ", " & _
        "PrintFilterHeader = " & IIf(chkPrintFilterHeader.Value = vbChecked, "1", "0") & ", "

    sSQL = sSQL & _
        "Table2ID = " & CStr(cboTable2.ItemData(cboTable2.ListIndex)) & ", " & _
        "Table2AllRecords = " & IIf(optAllRecords(1).Value, "1", "0") & ", " & _
        "Table2Picklist = " & IIf(optPicklist(1).Value, txtPicklist(1).Tag, "0") & ", " & _
        "Table2Filter = " & IIf(optFilter(1).Value, txtFilter(1).Tag, "0") & ", "

    sSQL = sSQL & _
        "OutputPreview = " & IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        "OutputFormat = " & CStr(mobjOutputDef.GetSelectedFormatIndex) & ", " & _
        "OutputScreen = " & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", " & _
        "OutputPrinter = " & IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0") & ", " & _
        "OutputPrinterName = '" & Replace(cboPrinterName.Text, " '", "''") & "', "
        
'    If chkDestination(desSave).Value = vbChecked Then
'      sSQL = sSQL & _
'        "OutputSave = 1, " & _
'        "OutputSaveExisting = " & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
'    Else
'      sSQL = sSQL & _
'        "OutputSave = 0, " & _
'        "OutputSaveExisting = 0, "
'    End If
    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & _
        "OutputSave = 1, " & _
        "OutputSaveFormat = " & Val(txtFilename.Tag) & ", " & _
        "OutputSaveExisting = " & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
    Else
      sSQL = sSQL & _
        "OutputSave = 0, " & _
        "OutputSaveFormat = 0, " & _
        "OutputSaveExisting = 0, "
    End If
    
    If chkDestination(desEmail).Value = vbChecked Then
      sSQL = sSQL & _
          "OutputEmail = 1, " & _
          "OutputEmailAddr = " & txtEmailGroup.Tag & ", " & _
          "OutputEmailSubject = '" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "OutputEmailAttachAs = '" & Replace(txtEmailAttachAs.Text, "'", "''") & "', " & _
          "OutputEmailFileFormat = " & CStr(Val(txtEmailAttachAs.Tag)) & ", "
    Else
      sSQL = sSQL & _
          "OutputEmail = 0, " & _
          "OutputEmailAddr = 0, " & _
          "OutputEmailSubject = '', " & _
          "OutputEmailAttachAs = '', " & _
          "OutputEmailFileFormat = 0, "
    End If
    
    sSQL = sSQL & _
        "OutputFilename = '" & Replace(txtFilename.Text, "'", "''") & "'"

    sSQL = sSQL & " WHERE MatchReportID = " & CStr(mlngMatchReportID)


    If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
      SaveDefinition = False
      Exit Function
    End If
    
    datData.ExecuteSql (sSQL)
    
    Select Case mlngMatchReportType
    Case mrtNormal
      UtilUpdateLastSaved utlMatchReport, mlngMatchReportID
    Case mrtSucession
      UtilUpdateLastSaved utlSuccession, mlngMatchReportID
    Case mrtCareer
      UtilUpdateLastSaved utlCareer, mlngMatchReportID
    End Select
  Else

    sSQL = "INSERT ASRSYSMatchReportName (" & _
        "Name, Description, UserName, MatchReportType, " & _
        "ScoreMode, NumRecords, ScoreCheck, ScoreLimit, EqualGrade, ReportingStructure, " & _
        "Table1ID, Table1AllRecords, Table1Picklist, Table1Filter, PrintFilterHeader, " & _
        "Table2ID, Table2AllRecords, Table2Picklist, Table2Filter, " & _
        "OutputPreview, OutputFormat, OutputScreen, OutputPrinter, OutputPrinterName, OutputSave, OutputSaveFormat, " & _
        "OutputSaveExisting, OutputEmail, OutputEmailAddr, OutputEmailSubject, OutputEmailAttachAs, OutputEmailFileFormat, OutputFilename) "

    sSQL = sSQL & "VALUES(" & _
        "'" & Trim(Replace(txtName.Text, "'", "''")) & "', " & _
        "'" & Trim(Replace(txtDesc.Text, "'", "''")) & "', " & _
        "'" & datGeneral.UserNameForSQL & "', "

    sSQL = sSQL & _
      CStr(mlngMatchReportType) & ", " & _
      IIf(optHighest.Value = True, "0, ", "1, ") & _
      CStr(spnMaxRecords.Value) & ", " & _
      IIf(chkLimit.Value = vbChecked, "1, ", "0, ") & _
      CStr(spnLimit.Value) & ", " & _
      IIf(chkEqualGrade.Value = vbChecked, "1, ", "0, ") & _
      IIf(chkReportStructure.Value = vbChecked, "1, ", "0, ")

    sSQL = sSQL & _
        CStr(cboTable1.ItemData(cboTable1.ListIndex)) & ", " & _
        IIf(optAllRecords(0).Value, "1", "0") & ", " & _
        IIf(optPicklist(0).Value, txtPicklist(0).Tag, "0") & ", " & _
        IIf(optFilter(0).Value, txtFilter(0).Tag, "0") & ", " & _
        IIf(chkPrintFilterHeader.Value = vbChecked, "1", "0") & ", "

    sSQL = sSQL & _
        CStr(cboTable2.ItemData(cboTable2.ListIndex)) & ", " & _
        IIf(optAllRecords(1).Value, "1", "0") & ", " & _
        IIf(optPicklist(1).Value, txtPicklist(1).Tag, "0") & ", " & _
        IIf(optFilter(1).Value, txtFilter(1).Tag, "0") & ", "

    sSQL = sSQL & _
        IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        CStr(mobjOutputDef.GetSelectedFormatIndex) & ", " & _
        IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", " & _
        IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0") & ", " & _
        "'" & Replace(cboPrinterName.Text, "'", "''") & "', "

    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & "1, " & CStr(Val(txtFilename.Tag)) & ", " & _
        cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
    Else
      sSQL = sSQL & "0, 0, 0, "
    End If

    If chkDestination(desEmail).Value = vbChecked Then
      sSQL = sSQL & "1, " & _
          txtEmailGroup.Tag & ", " & _
          "'" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "'" & Replace(txtEmailAttachAs.Text, "'", "''") & ", " & _
          CStr(Val(txtEmailAttachAs.Tag)) & ", "      'OutputEmail, OutputEmailAddr, OutputEmailAttachAs, OutputEmailFileFormat, OutputEmailSubject
    Else
      sSQL = sSQL & "0, 0, '', '', 0, "               'OutputEmail, OutputEmailAddr, OutputEmailAttachAs, OutputEmailFileFormat, OutputEmailSubject
    End If

    sSQL = sSQL & _
        "'" & Replace(txtFilename.Text, "'", "''") & "')"

    If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
      SaveDefinition = False
      Exit Function
    End If
  
    mlngMatchReportID = InsertMatchReport(sSQL, "AsrSysMatchReportName", "MatchReportID")
    
    Select Case mlngMatchReportType
    Case mrtNormal
      UtilCreated utlMatchReport, mlngMatchReportID
    Case mrtSucession
      UtilCreated utlSuccession, mlngMatchReportID
    Case mrtCareer
      UtilCreated utlCareer, mlngMatchReportID
    End Select
  End If
  
  SaveAccess
  
  gADOCon.Execute "DELETE FROM ASRSysMatchReportDetails WHERE MatchReportID = " & CStr(mlngMatchReportID)
  gADOCon.Execute "DELETE FROM ASRSysMatchReportBreakdown WHERE MatchReportID = " & CStr(mlngMatchReportID)
  gADOCon.Execute "DELETE FROM ASRSysMatchReportTables WHERE MatchReportID = " & CStr(mlngMatchReportID)
  
  
  
  'Check the sequence of the selected columns
  With ListView2
    For lngCount = 1 To .ListItems.Count
      mcolMatchReportColDetails(ListView2.ListItems(lngCount).Key).Sequence = lngCount
    Next
  End With


  'COLUMNS
  For Each objCol In mcolMatchReportColDetails
    sSQL = _
      "INSERT ASRSysMatchReportDetails (MatchReportID, ColType, " & _
      "ColExprID, ColSize, ColDecs, ColHeading, ColSequence, SortOrderSeq, SortOrderDirection)" & _
      "VALUES (" & _
      CStr(mlngMatchReportID) & ", " & _
      "'" & objCol.ColType & "', " & _
      CStr(objCol.ID) & ", " & _
      CStr(objCol.Size) & ", " & _
      CStr(objCol.DecPlaces) & ", " & _
      "'" & Replace(objCol.Heading, "'", "''") & "', " & _
      CStr(objCol.Sequence) & "," & _
      GetSortOrder(objCol.ID) & ")"
    datData.ExecuteSql sSQL
  Next


  For Each objRelation In colRelatedTables
    sSQL = _
      "INSERT ASRSysMatchReportTables (MatchReportID, " & _
      "Table1ID, Table2ID, RequiredExprID, PreferredExprID, MatchScoreExprID) " & _
      "VALUES (" & _
      CStr(mlngMatchReportID) & ", " & _
      CStr(objRelation.Table1ID) & ", " & _
      CStr(objRelation.Table2ID) & ", " & _
      CStr(objRelation.RequiredExprID) & ", " & _
      CStr(objRelation.PreferredExprID) & ", " & _
      CStr(objRelation.MatchScoreID) & ")"

    lngMatchRelationID = InsertMatchReport(sSQL, "ASRSysMatchReportTables", "MatchRelationID")
  
    If lngMatchRelationID > 0 Then
      For Each objCol In objRelation.BreakdownColumns
        sSQL = _
          "INSERT ASRSysMatchReportBreakdown (MatchReportID, MatchRelationID, " & _
          "ColType, ColExprID, ColSize, ColDecs, ColHeading, ColSequence)" & _
          "VALUES (" & _
          CStr(mlngMatchReportID) & ", " & _
          CStr(lngMatchRelationID) & ", " & _
          "'" & objCol.ColType & "', " & _
          CStr(objCol.ID) & ", " & _
          CStr(objCol.Size) & ", " & _
          CStr(objCol.DecPlaces) & ", " & _
          "'" & Replace(objCol.Heading, "'", "''") & "', " & _
          CStr(objCol.Sequence) & ")"
        datData.ExecuteSql sSQL
      Next
    End If

  Next
  
  
  SaveDefinition = True
  Changed = False
  
  Exit Function

Save_ERROR:

  SaveDefinition = False
  MsgBox "Warning : An error has occurred whilst saving..." & vbCrLf & Err.Description & vbCrLf & "Please cancel and try again. If this error continues, delete the definition.", vbCritical + vbOKOnly, Me.Caption

End Function


Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysMatchReportAccess WHERE ID = " & mlngMatchReportID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysMatchReportAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngMatchReportID & ", sysusers.name," & _
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
      sSQL = "IF EXISTS (SELECT * FROM ASRSysMatchReportAccess" & _
        " WHERE ID = " & CStr(mlngMatchReportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysMatchReportAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngMatchReportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub





Private Function IsInSortOrder(sKey As String) As String

  ' This function is used when saving the column definition.
  ' It Checks if the column has been defined as a sort order.
  ' Either way, it returns the string which should be added to
  ' the sSQL string that is being created.
  '
  ' <sequence>, <sortorder> <break>, <page>, <value>, <suppress>
  '
  
  Dim bm As Variant
  Dim LLoop As Long
  
  With grdReportOrder
    .MoveFirst
      Do Until LLoop = .Rows
        bm = .GetBookmark(LLoop)
        If .Columns(0).CellText(bm) = Right(sKey, Len(sKey) - 1) Then
          IsInSortOrder = CStr(LLoop + 1) & ","
          IsInSortOrder = IsInSortOrder & "'" & .Columns(2).CellText(bm) & "',"
          IsInSortOrder = IsInSortOrder & IIf(.Columns(3).CellValue(bm), 1, 0) & ","
          IsInSortOrder = IsInSortOrder & IIf(.Columns(4).CellValue(bm), 1, 0) & ","
          IsInSortOrder = IsInSortOrder & IIf(.Columns(5).CellValue(bm), 1, 0) & ","
          IsInSortOrder = IsInSortOrder & IIf(.Columns(6).CellValue(bm), 1, 0) & ""
          Exit Function
        End If
        LLoop = LLoop + 1
      Loop
  
  
'  Dim iloop As Integer
'
'  With frmMatchReports.grdReportOrder
'
'    For iloop = 0 To .Rows - 1
'
'      .Row = iloop
'
'      If .Columns(0).Text = Right(skey, Len(skey) - 1) Then
'
'        IsInSortOrder = CStr(iloop + 1) & ","
'        IsInSortOrder = IsInSortOrder & "'" & .Columns(2).Text & "',"
'        IsInSortOrder = IsInSortOrder & IIf(.Columns(3).Value, 1, 0) & ","
'        IsInSortOrder = IsInSortOrder & IIf(.Columns(4).Value, 1, 0) & ","
'        IsInSortOrder = IsInSortOrder & IIf(.Columns(5).Value, 1, 0) & ","
'        IsInSortOrder = IsInSortOrder & IIf(.Columns(6).Value, 1, 0) & ")"
'
'        Exit Function
'
'      End If
      
'    Next iloop

  End With
  
  IsInSortOrder = "'',0,0,0,0,0"

End Function

Private Function InsertChildDetails() As String

'  Dim pvarbookmark  As Variant
'  Dim i As Integer
'  Dim sSQL As String
'
'  sSQL = "("
'
'  With grdChildren
'    .MoveFirst
'    For i = 0 To .Rows - 1 Step 1
'      pvarbookmark = .GetBookmark(i)
'
'      sSQL = "INSERT INTO ASRSysMatchReportsChildDetails "
'      sSQL = sSQL & "VALUES ("
'      sSQL = sSQL & mlngMatchReportID & ","
'      sSQL = sSQL & .Columns("TableID").CellValue(pvarbookmark) & ","
'      sSQL = sSQL & IIf(.Columns("FilterID").CellValue(pvarbookmark) = vbNullString, 0, .Columns("FilterID").CellValue(pvarbookmark)) & ","
'      sSQL = sSQL & IIf(.Columns("Records").CellValue(pvarbookmark) = sALL_RECORDS, 0, .Columns("Records").CellValue(pvarbookmark))
'      sSQL = sSQL & ")"
'
'      datData.ExecuteSql (sSQL)
'    Next i
'  End With

End Function

Private Function RetrieveMatchReportDetails(plngMatchReportID As Long) As Boolean

  Dim objRelation As clsMatchRelation
  Dim objColumn As clsColumn
  Dim objExpression As clsExprExpression
  Dim rsColumns As Recordset
  Dim rsDef As Recordset
  'Dim colBreakdownCols As Collection
  
  Dim iLoop As Integer
  Dim sText As String
  Dim sMessage As String
  Dim sSQL As String
  Dim fAlreadyNotified As Boolean
  
  Dim lngRequiredExprID As Long
  Dim lngPreferredExprID As Long
  Dim lngMatchScoreExprID As Long
  
  
  On Error GoTo Load_ERROR
  
  Set rsDef = datGeneral.GetRecords("SELECT ASRSysMatchReportName.*, " & _
                                     "CONVERT(integer, ASRSysMatchReportName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysMatchReportName WHERE MatchReportID = " & CStr(plngMatchReportID))

  If rsDef.BOF And rsDef.EOF Then
    MsgBox "This Report definition has been deleted by another user.", vbExclamation + vbOKOnly, Me.Caption
    Set rsDef = Nothing
    RetrieveMatchReportDetails = False
    mblnDeleted = True
    Exit Function
  End If
  
  txtName.Text = rsDef!Name
  txtDesc.Text = IIf(IsNull(rsDef!Description), "", rsDef!Description)
    
  If FromCopy Then
    txtName.Text = "Copy of " & rsDef!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsDef!Name
    txtUserName = StrConv(rsDef!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase(rsDef!UserName) = LCase(gsUserName))
  End If
  
  ' Set Base Table

  mblnLoading = True

  LoadTable1Combo
  SetComboText cboTable1, datGeneral.GetTableName(rsDef!Table1ID)
  mstrTable1Name = cboTable1.Text

  LoadTable2Combo
  SetComboText cboTable2, datGeneral.GetTableName(rsDef!Table2ID)
  mstrTable2Name = cboTable2.Text

  UpdateDependantFields

  'PopulateTableAvailable

  'JPD 20040312 Fault 8266
  'mblnLoading = False

  SetRecordSelection 0, rsDef!Table1AllRecords, rsDef!Table1Picklist, rsDef!Table1Filter
  SetRecordSelection 1, rsDef!Table2AllRecords, rsDef!Table2Picklist, rsDef!Table2Filter
  If Not IsNull(rsDef!PrintFilterHeader) Then
    chkPrintFilterHeader.Value = IIf(rsDef!PrintFilterHeader, vbChecked, vbUnchecked)
  End If

  spnMaxRecords.Value = IIf(IsNull(rsDef!NumRecords), 0, rsDef!NumRecords)
  
  If Not IsNull(rsDef!ScoreMode) Then
    optLowest.Value = (rsDef!ScoreMode = 1)
    chkLimit.Value = IIf(rsDef!ScoreCheck, vbChecked, vbUnchecked)
    spnLimit.Value = rsDef!ScoreLimit
    chkEqualGrade.Value = IIf(rsDef!EqualGrade, vbChecked, vbUnchecked)
    chkReportStructure.Value = IIf(rsDef!ReportingStructure, vbChecked, vbUnchecked)
  End If

  mobjOutputDef.ReadDefFromRecset rsDef
  mlngTimeStamp = rsDef!intTimestamp

  ' =========================

  mblnReadOnly = Not datGeneral.SystemPermission("MATCHREPORTS", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    Select Case mlngMatchReportType
      Case mrtSucession
        mblnReadOnly = (CurrentUserAccess(utlSuccession, plngMatchReportID) = ACCESS_READONLY)
      Case mrtCareer
        mblnReadOnly = (CurrentUserAccess(utlMatchReport, plngMatchReportID) = ACCESS_READONLY)
      Case Else
        mblnReadOnly = (CurrentUserAccess(utlCareer, plngMatchReportID) = ACCESS_READONLY)
    End Select
  End If

  If mblnReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  ' =========================
  
  
  ' Now load the relations
  Set rsDef = datData.OpenPersistentRecordset( _
    "SELECT ASRSysMatchReportTables.*, " & _
    "a.tablename as 'Table1Name', " & _
    "b.tablename as 'Table2Name', " & _
    "isnull(r.Name,'<None>') as 'RequiredExprName', " & _
    "isnull(p.Name,'<None>') as 'PreferredExprName', " & _
    "isnull(s.Name,'<None>') as 'MatchScoreExprName' " & _
    "From ASRSysMatchReportTables " & _
    "LEFT OUTER JOIN asrsystables AS a ON a.tableid = ASRSysMatchReportTables.Table1ID " & _
    "LEFT OUTER JOIN asrsystables AS b ON b.tableid = ASRSysMatchReportTables.Table2ID " & _
    "LEFT OUTER JOIN asrsysexpressions AS r ON r.exprid = ASRSysMatchReportTables.RequiredExprID " & _
    "LEFT OUTER JOIN asrsysexpressions AS p ON p.exprid = ASRSysMatchReportTables.PreferredExprID " & _
    "LEFT OUTER JOIN asrsysexpressions AS s ON s.exprid = ASRSysMatchReportTables.MatchScoreExprID " & _
    "WHERE MatchReportID = " & CStr(mlngMatchReportID) & _
    " ORDER BY ASRSysMatchReportTables.MatchRelationID", adOpenKeyset, adLockOptimistic)


  If rsDef.BOF And rsDef.EOF Then
    MsgBox "Cannot load the table relation information for this definition", vbExclamation + vbOKOnly, Me.Caption
    RetrieveMatchReportDetails = False
    Set rsDef = Nothing
    Exit Function
  End If
  
  
  Do While Not rsDef.EOF
    
    If mblnFromCopy Then
      lngRequiredExprID = CopyExpression(rsDef!Table2ID, rsDef!RequiredExprID, giEXPR_MATCHWHEREEXPRESSION)
      lngPreferredExprID = CopyExpression(rsDef!Table1ID, rsDef!PreferredExprID, giEXPR_MATCHJOINEXPRESSION)
      lngMatchScoreExprID = CopyExpression(rsDef!Table1ID, rsDef!MatchScoreExprID, giEXPR_MATCHSCOREEXPRESSION)
    Else
      lngRequiredExprID = rsDef!RequiredExprID
      lngPreferredExprID = rsDef!PreferredExprID
      lngMatchScoreExprID = rsDef!MatchScoreExprID
    End If
    
    
    grdRelations.AddItem ( _
      rsDef!Table1ID & vbTab & _
      rsDef!Table1Name & vbTab & _
      rsDef!Table2ID & vbTab & _
      rsDef!Table2Name & vbTab & _
      rsDef!RequiredExprName & vbTab & _
      rsDef!PreferredExprName & vbTab & _
      rsDef!MatchScoreExprName)

    Set objRelation = New clsMatchRelation
    'Set colBreakdownCols = New Collection
    
    objRelation.Table1ID = rsDef!Table1ID
    objRelation.Table1Name = rsDef!Table1Name
    objRelation.Table2ID = rsDef!Table2ID
    objRelation.Table2Name = IIf(IsNull(rsDef!Table2Name), vbNullString, rsDef!Table2Name)
    objRelation.RequiredExprID = lngRequiredExprID
    objRelation.PreferredExprID = lngPreferredExprID
    objRelation.MatchScoreID = lngMatchScoreExprID
    
    
    Set rsColumns = datGeneral.GetReadOnlyRecords( _
    "SELECT ASRSysColumns.ColumnName, ASRSysTables.TableName, " & _
    "ASRSysMatchReportBreakdown.* FROM ASRSysMatchReportBreakdown " & _
    "LEFT OUTER JOIN ASRSysColumns ON ASRSysMatchReportBreakdown.ColExprID = ASRSysColumns.ColumnID " & _
    "LEFT OUTER JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID " & _
    "WHERE MatchRelationID = " & CStr(rsDef!MatchRelationID))

    Do While Not rsColumns.EOF

      Set objColumn = New clsColumn
      objColumn.ColType = rsColumns!ColType
      objColumn.ID = rsColumns!ColExprID
      objColumn.Size = rsColumns!ColSize
      objColumn.DecPlaces = rsColumns!ColDecs
      objColumn.Heading = rsColumns!ColHeading
      objColumn.Sequence = rsColumns!ColSequence

      If rsColumns!ColType = "C" Then
        objColumn.ColumnName = rsColumns!TableName & "." & rsColumns!ColumnName
        objColumn.IsNumeric = datGeneral.NumericColumn("C", GetTableIDFromColumn(rsColumns!ColExprID), rsColumns!ColExprID)
      Else
        objColumn.IsNumeric = True   'Has to be match score which is numeric
      End If

      objRelation.BreakdownColumns.Add objColumn, rsColumns!ColType & CStr(rsColumns!ColExprID)

      rsColumns.MoveNext
    Loop

    rsColumns.Close
    Set rsColumns = Nothing

    colRelatedTables.Add objRelation, "T" & CStr(objRelation.Table1ID)

    rsDef.MoveNext
  Loop

  rsDef.Close
  Set rsDef = Nothing

  ' =========================

  sMessage = vbNullString

  ' Now load the columns guff
  Set rsColumns = datGeneral.GetRecords("SELECT * FROM ASRSysMatchReportDetails WHERE MatchReportID = " & plngMatchReportID & " ORDER BY [ColSequence]")

  If rsColumns.BOF And rsColumns.EOF Then
    MsgBox "Cannot load the column definition for this definition", vbExclamation + vbOKOnly, Me.Caption
    RetrieveMatchReportDetails = False
    Set rsColumns = Nothing
    Exit Function
  End If

  Do Until rsColumns.EOF

    If rsColumns!ColType = "C" Then
      sText = datGeneral.GetTableName(datGeneral.GetColumnTable(rsColumns!ColExprID)) & "." & datGeneral.GetColumnName(rsColumns!ColExprID)
      ListView2.ListItems.Add , rsColumns!ColType & CStr(rsColumns!ColExprID), sText, ImageList1.ListImages("IMG_TABLE").Index, ImageList1.ListImages("IMG_TABLE").Index
    Else
      sText = "Match Score"
      ListView2.ListItems.Add , "E" & CStr(rsColumns!ColExprID), sText, ImageList1.ListImages("IMG_MATCH").Index, ImageList1.ListImages("IMG_MATCH").Index
    End If

    ' Add to collection
    If sText <> vbNullString And sMessage = vbNullString Then
      'mcolMatchReportColDetails.Add rsColumns!Type, rsColumns!ColExprID, rsColumns!Heading, rsColumns!Size, rsColumns!dp, rsColumns!Avge, rsColumns!cnt, rsColumns!tot, rsColumns!IsNumeric

      Set objColumn = New clsColumn

      objColumn.ColumnName = sText
      objColumn.ColType = rsColumns!ColType
      objColumn.ID = rsColumns!ColExprID
      objColumn.Size = rsColumns!ColSize
      objColumn.DecPlaces = rsColumns!ColDecs
      objColumn.Heading = rsColumns!ColHeading
      objColumn.Sequence = rsColumns!ColSequence
    
      If rsColumns!ColType = "C" Then
        objColumn.IsNumeric = datGeneral.NumericColumn("C", GetTableIDFromColumn(rsColumns!ColExprID), rsColumns!ColExprID)
      Else
        If rsColumns!ColExprID = 0 Then
          objColumn.IsNumeric = True
        Else
          objColumn.IsNumeric = datGeneral.NumericColumn("E", rsColumns!BaseTable, rsColumns!ColExprID)
        End If
      End If

      mcolMatchReportColDetails.Add objColumn, objColumn.ColType & CStr(objColumn.ID)
      Set objColumn = Nothing

    End If
    
    sMessage = vbNullString
    rsColumns.MoveNext

  Loop
  
  
  
  mblnLoading = True
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    RetrieveMatchReportDetails = False
    Exit Function
  End If
  
  mblnLoading = False

'  'TM20020220 Fault 3372
'  'PopulateAvailable
'  PopulateTableAvailable , True

  UpdateButtonStatus (SSTab1.Tab)
  
  
  ' Now do the sort order guff
  Set rsDef = datGeneral.GetRecords("SELECT * FROM ASRSysMatchReportDetails WHERE MatchReportID = " & plngMatchReportID & " AND SortOrderSeq > 0 ORDER BY [SortOrderSeq]")

  If rsDef.BOF And rsDef.EOF Then
    MsgBox "Cannot load the sort order for this definition", vbExclamation + vbOKOnly, Me.Caption
    RetrieveMatchReportDetails = False
    Set rsDef = Nothing
    Exit Function
  End If

  ' Add to the sort order grid
  Do Until rsDef.EOF
    If rsDef!ColType = "C" Then
      grdReportOrder.AddItem _
            rsDef!ColExprID & vbTab & _
            GetTableNameFromColumn(rsDef!ColExprID) & "." & datGeneral.GetColumnName(rsDef!ColExprID) & vbTab & _
            IIf(rsDef!SortOrderDirection = "A", "Ascending", "Descending")
    Else
      grdReportOrder.AddItem _
            "0" & vbTab & "Match Score" & vbTab & _
            IIf(rsDef!SortOrderDirection = "A", "Ascending", "Descending")
    End If
    
    rsDef.MoveNext
  Loop
  Set rsDef = Nothing

  With grdReportOrder
    .SelBookmarks.RemoveAll
    .MoveFirst
    .SelBookmarks.Add (.Bookmark)
  End With
  UpdateOrderButtons


  RetrieveMatchReportDetails = True
  Exit Function

Load_ERROR:

  MsgBox "Warning : Error whilst retrieving the report definition." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, Me.Caption
  RetrieveMatchReportDetails = False
  Set rsDef = Nothing

End Function


Private Sub UpdateDependantFields()
'''
'''  ' This sub populates the parent/child combos depending
'''  ' on the base table selected
'''
'''  Dim rsParents As New Recordset
'''  Dim rsTables As New Recordset
'''  Dim rsChildren As New Recordset
'''  Dim sSQL As String
'''  Dim lngTableID As Long
'''
'''  lngTableID = 0
'''  If cboTable1.ListIndex <> -1 Then
'''    lngTableID = cboTable1.ItemData(cboTable1.ListIndex)
'''  End If
'''
'''  'If mblnLoading Then Exit Sub
'''
'''  ' Get the parent(s) of the selected base table
'''
'''  sSQL = "SELECT asrsystables.tablename, asrsystables.tableid " & _
'''         "FROM asrsystables " & _
'''         "WHERE asrsystables.tableid in " & _
'''         "(select parentid from asrsysrelations " & _
'''         "WHERE childid = " & CStr(lngTableID) & ") " & _
'''         "ORDER BY tablename"
'''
'''  Set rsParents = datData.OpenPersistentRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'''
'''  If Not rsParents.BOF And Not rsParents.EOF Then
'''    rsParents.MoveLast
'''    rsParents.MoveFirst
'''  End If
'''
'''  Select Case rsParents.RecordCount
'''
'''    Case 0
'''      txtParent1.Text = "" '"<None>"
'''      txtParent1.Tag = 0
'''      optParent1AllRecords.Value = True
'''      txtParent1Filter.Text = ""
'''      txtParent1Filter.Tag = 0
'''      txtParent1Picklist.Text = ""
'''      txtParent1Picklist.Tag = 0
'''      fraParent1.Enabled = False
'''
'''      txtParent2.Text = "" '"<None>"
'''      txtParent2.Tag = 0
'''      optParent2AllRecords.Value = True
'''      txtParent2Filter.Text = ""
'''      txtParent2Filter.Tag = 0
'''      txtParent2Picklist.Text = ""
'''      txtParent2Picklist.Tag = 0
'''      fraParent2.Enabled = False
'''
'''    Case 1
'''      txtParent1.Text = rsParents!TableName
'''      txtParent1.Tag = rsParents!TableID
'''      optParent1AllRecords.Value = True
'''      txtParent1Filter.Text = ""
'''      txtParent1Filter.Tag = 0
'''      txtParent1Picklist.Text = ""
'''      txtParent1Picklist.Tag = 0
'''      fraParent1.Enabled = True
'''
'''      txtParent2.Text = "" '"<None>"
'''      txtParent2.Tag = 0
'''      optParent2AllRecords.Value = True
'''      txtParent2Filter.Text = ""
'''      txtParent2Filter.Tag = 0
'''      txtParent2Picklist.Text = ""
'''      txtParent2Picklist.Tag = 0
'''      fraParent2.Enabled = False
'''
'''    Case 2
'''      txtParent1.Text = rsParents!TableName
'''      txtParent1.Tag = rsParents!TableID
'''      optParent1AllRecords.Value = True
'''      txtParent1Filter.Text = ""
'''      txtParent1Filter.Tag = 0
'''      txtParent1Picklist.Text = ""
'''      txtParent1Picklist.Tag = 0
'''      fraParent1.Enabled = True
'''
'''      rsParents.MoveNext
'''
'''      txtParent2.Text = rsParents!TableName
'''      txtParent2.Tag = rsParents!TableID
'''      optParent2AllRecords.Value = True
'''      txtParent2Filter.Text = ""
'''      txtParent2Filter.Tag = 0
'''      txtParent2Picklist.Text = ""
'''      txtParent2Picklist.Tag = 0
'''      fraParent2.Enabled = True
'''  End Select
'''
'''  ' Clear recordset reference
'''  Set rsParents = Nothing
'''
''''  ' Clear Child Combo and add <None> entry
''''
''''  With cboChild
''''    .Clear
''''    .AddItem "<None>"
''''    .ItemData(.NewIndex) = 0
''''    mblnLoading = True
''''    .ListIndex = 0
''''    mblnLoading = False
''''  End With
''''
'''  ' Get the children of the selected base table
'''  sSQL = "SELECT asrsystables.tablename, asrsystables.tableid " & _
'''         "FROM asrsystables " & _
'''         "WHERE asrsystables.tableid in " & _
'''         "(select childid from asrsysrelations " & _
'''         "WHERE parentid = " & CStr(lngTableID) & ") " & _
'''         "ORDER BY tablename"
'''
'''  Set rsChildren = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'''
'''  If rsChildren.BOF And rsChildren.EOF Then
'''    fraChild.Enabled = False
'''    cmdAddChild.Enabled = False
'''    cmdEditChild.Enabled = False
'''    cmdRemove.Enabled = False
'''    cmdRemoveAllChilds.Enabled = False
'''    grdChildren.Enabled = False
'''  Else
'''    fraChild.Enabled = True
'''    grdChildren.Enabled = True
'''  End If
'''  Set rsChildren = Nothing
'''
''''  txtChildFilter.Text = ""
''''  txtChildFilter.Tag = 0
''''  spnMaxRecords.Value = 0
''''
End Sub


Public Sub PopulateTableAvailable(Optional pstrTable As String, Optional pbSetToBase As Boolean)
  
  'TM20020424 Fault 3715 - have added optional pbSetToBase to the sub, so the changing of the
  'cboTblAvailable.listindex property is only called when this is true.
  
  ' Now populate the TableAvailable combo
  
  If pbSetToBase Then
    cboTblAvailable.Clear
    ' Clear the listview
    ListView1.ListItems.Clear
  End If
  
  ' Add the base table to the top of the combo
  If cboTable1.ItemData(cboTable1.ListIndex) > 0 Then
    cboTblAvailable.AddItem cboTable1.Text
    cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = cboTable1.ItemData(cboTable1.ListIndex)
  End If

  If cboTable2.ItemData(cboTable2.ListIndex) > 0 Then
    cboTblAvailable.AddItem cboTable2.Text
    cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = cboTable2.ItemData(cboTable2.ListIndex)
  End If


'  'TM20020108 Fault 3165
'  ' Add the child tables to the combo if selected
'  Dim i As Integer
'  Dim pvarbookmark As Variant
'  With grdChildren
'    If .Rows > 0 Then
'      For i = 0 To .Rows - 1 Step 1
'        pvarbookmark = .AddItemBookmark(i)
'        If Not TableAlreadyAvailable(CInt(.Columns("TableID").CellValue(pvarbookmark))) Or pbSetToBase Then
'          cboTblAvailable.AddItem .Columns("Table").CellValue(pvarbookmark)
'          cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = CInt(.Columns("TableID").CellValue(pvarbookmark))
'        End If
'      Next i
'    End If
'  End With
  
  
  If Not IsMissing(pbSetToBase) Then
    If pbSetToBase Then
      ' Select the base table in the combo by default
      If Len(pstrTable) > 0 Then
        SetComboText cboTblAvailable, pstrTable
      Else
        If cboTblAvailable.ListCount > 0 Then
          cboTblAvailable.ListIndex = 0
        End If
      End If
    End If
  End If
  
  ' If theres only 1 table, then disable the combo, otherwise enable it
  If cboTblAvailable.ListCount = 1 Then
    cboTblAvailable.Enabled = False
    cboTblAvailable.BackColor = vbButtonFace
  Else
    cboTblAvailable.Enabled = True
    cboTblAvailable.BackColor = vbWindowBackground
  End If

End Sub


Public Function AnyChildColumnsUsed(lngTableID As Long, Optional bAutoYes As Boolean) As Integer

  ' This sub checks if any columns from the Child table which has just
  ' been deselected have been used in the current report definition.
  ' If so, user is prompted if they wish to continue. Continuing will
  ' delete the columns in the report from the old Child table.

  Dim objItem As ListItem
  Dim tempColumnID As Long
  Dim tempKey As String
  Dim fUsed As Boolean
  Dim rsCalc As ADODB.Recordset
  Dim lngExpr As Long
  
  For Each objItem In ListView2.ListItems
    If Left(objItem.Key, 1) = "C" Then
      tempColumnID = CLng(Right(objItem.Key, Len(objItem.Key) - 1))
      If GetTableIDFromColumn(tempColumnID) = lngTableID Then
        fUsed = True
      End If
    ElseIf Left(objItem.Key, 1) = "E" Then
      lngExpr = Right(objItem.Key, Len(objItem.Key) - 1)
      Set rsCalc = datGeneral.GetReadOnlyRecords("SELECT ExprID, Name, TableID FROM ASRSysExpressions WHERE ExprID = " & CStr(lngExpr))
       
      If Not (rsCalc.BOF And rsCalc.EOF) Then
        If rsCalc!TableID = lngTableID Then
          fUsed = True
        End If
      End If
      Set rsCalc = Nothing
    End If
  Next objItem
  
  If Not fUsed Then
    AnyChildColumnsUsed = 0
    Set objItem = Nothing
    Exit Function
  End If
  
  If Not bAutoYes Then
    If MsgBox("One or more columns from the '" & datGeneral.GetTableName(lngTableID) & "' table have been included in the current report definition." & vbCrLf & _
              "Changing the child table will remove these columns from the report definition." & vbCrLf & _
              "Do you wish to continue ?" _
              , vbYesNo + vbQuestion, Me.Caption) = vbNo Then
      AnyChildColumnsUsed = 1
      Set objItem = Nothing
      Exit Function
    End If
  End If
  
  Dim iLoop As Integer
  
  For iLoop = ListView2.ListItems.Count To 1 Step -1
    Set objItem = ListView2.ListItems(iLoop)
    If Left(objItem.Key, 1) = "C" Then
      tempColumnID = CLng(Right(objItem.Key, Len(objItem.Key) - 1))
      If GetTableIDFromColumn(tempColumnID) = lngTableID Then
        tempKey = objItem.Key
        ListView2.ListItems.Remove tempKey
        mcolMatchReportColDetails.Remove tempKey
      
        ' also remove from the sort order if its there
        grdReportOrder.MoveFirst
        
        Dim i As Integer
        For i = 0 To (grdReportOrder.Rows - 1)
          If Right(tempKey, Len(tempKey) - 1) = grdReportOrder.Columns(0).CellValue(grdReportOrder.Bookmark) Then
            ' delete
            If grdReportOrder.Rows = 1 Then
              grdReportOrder.RemoveAll
            Else
              grdReportOrder.RemoveItem (grdReportOrder.AddItemRowIndex(grdReportOrder.Bookmark))
            End If
          End If
        grdReportOrder.MoveNext
        Next i
  
      End If
    ElseIf Left(objItem.Key, 1) = "E" Then
      lngExpr = Right(objItem.Key, Len(objItem.Key) - 1)
      Set rsCalc = datGeneral.GetReadOnlyRecords("SELECT ExprID, Name, TableID FROM ASRSysExpressions WHERE ExprID = " & CStr(lngExpr))
       
      If Not (rsCalc.BOF And rsCalc.EOF) Then
        If rsCalc!TableID = lngTableID Then
          tempKey = objItem.Key
          ListView2.ListItems.Remove tempKey
          mcolMatchReportColDetails.Remove tempKey
        End If
      End If
      Set rsCalc = Nothing
    End If
  Next iLoop
  
  Set objItem = Nothing
  AnyChildColumnsUsed = 2
  
End Function

Public Function ValidateDefinition(lngCurrentID As Long) As Boolean
'
'  Dim LLoop As Long
'  Dim bm As Variant
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim strRecSelStatus As String
  Dim strDuplicateHeading As String

  Dim lngIndex As Long

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
    SSTab1.Tab = 0
    MsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If

  'Check if this definition has been changed by another user
  Select Case mlngMatchReportType
    Case mrtSucession
      Call UtilityAmended(utlSuccession, mlngMatchReportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
    Case mrtCareer
      Call UtilityAmended(utlCareer, mlngMatchReportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
    Case Else
      Call UtilityAmended(utlMatchReport, mlngMatchReportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  End Select
  
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngMatchReportID = 0
    mblnReadOnly = False
    ForceAccess
  End If

  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngMatchReportID) Then
    SSTab1.Tab = 0
    MsgBox "A " & Me.Caption & " definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    Exit Function
  End If

  
  For lngIndex = 0 To 1
  
    If optPicklist(lngIndex).Value Then
      If txtPicklist(lngIndex).Text = "" Or txtPicklist(lngIndex).Tag = "0" Or txtPicklist(lngIndex).Tag = "" Then
        SSTab1.Tab = 0
        MsgBox "You must select a picklist, or change the record selection for your " & _
            IIf(lngIndex = 0, "base table.", "match table."), vbExclamation + vbOKOnly, Me.Caption
        cmdPicklist(lngIndex).SetFocus
        ValidateDefinition = False
        Exit Function
      End If
    End If
  
    ' BASE TABLE - If using a filter, check one has been selected
    If optFilter(lngIndex).Value Then
      If txtFilter(lngIndex).Text = "" Or txtFilter(lngIndex).Tag = "0" Or txtFilter(lngIndex).Tag = "" Then
        SSTab1.Tab = 0
        MsgBox "You must select a filter, or change the record selection for your " & _
            IIf(lngIndex = 0, "base table.", "match table."), vbExclamation + vbOKOnly, Me.Caption
        cmdFilter(lngIndex).SetFocus
        ValidateDefinition = False
        Exit Function
      End If
    End If

  Next
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
    
  ' Check that there are columns defined in the report definition
  If Me.grdRelations.Rows = 0 Then
    SSTab1.Tab = 1
    MsgBox "You must setup at least one table relation", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  
  
  ' Check that there are columns defined in the report definition
  If ListView2.ListItems.Count = 0 Then
    SSTab1.Tab = 2
    MsgBox "You must select at least 1 column for your report.", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If


  strDuplicateHeading = CheckForDuplicateHeadings(mcolMatchReportColDetails)
  If strDuplicateHeading <> vbNullString Then
    SSTab1.Tab = 2
    MsgBox "More than one column has a heading of '" & strDuplicateHeading & "'" & vbCrLf & "Column headings must be unique.", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If


  ' Check that at least 1 column has been defined as the report order
  With grdReportOrder
    If .Rows = 0 Then
      SSTab1.Tab = 3
      MsgBox "You must select at least one column to order the report by.", vbExclamation + vbOKOnly, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
  End With

'  ' Check the no. of items in the collection is the same as the
'  ' number of items in the list view.
'  If Not ValidateCollection Then
'    ValidateDefinition = False
'    Exit Function
'  End If
  If Not mobjOutputDef.ValidDestination Then
    SSTab1.Tab = SSTab1.Tabs - 1
    Exit Function
  End If
  
If mlngMatchReportID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    Select Case mlngMatchReportType
      Case mrtSucession
        CheckCanMakeHiddenInBatchJobs utlSuccession, _
          CStr(mlngMatchReportID), _
          txtUserName.Text, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups, _
          sHiddenGroups
      Case mrtCareer
        CheckCanMakeHiddenInBatchJobs utlCareer, _
          CStr(mlngMatchReportID), _
          txtUserName.Text, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups, _
          sHiddenGroups
      Case Else
        CheckCanMakeHiddenInBatchJobs utlMatchReport, _
          CStr(mlngMatchReportID), _
          txtUserName.Text, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups, _
          sHiddenGroups
    End Select

    If (Not fBatchJobsOK) Then
      If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
        MsgBox "This definition cannot be made hidden from the following user groups :" & vbCrLf & vbCrLf & sBatchJobScheduledUserGroups & vbCrLf & _
               "as it is used in the following batch jobs which are scheduled to be run by these user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, _
               vbExclamation + vbOKOnly, Me.Caption
      Else
        MsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , Me.Caption
      End If

      Screen.MousePointer = vbNormal
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If MsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
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





Private Function CheckUniqueName(sName As String, lngCurrentID As Long) As Boolean

  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSYSMatchReportName " & _
         " WHERE UPPER(Name) = '" & UCase(Replace(sName, "'", "''")) & "'" & _
         " AND MatchReportType = " & CStr(mlngMatchReportType) & _
         " AND MatchReportID <> " & CStr(lngCurrentID)
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckUniqueName = True
  Else
    CheckUniqueName = False
  End If
  
  Set rsTemp = Nothing

End Function


Public Sub PopulateAvailable()
  
  ' This function is called whenever a new table is selected in the
  ' table combo, or when cols/expressions are removed from the report
  ' definition. It checks through each item in the 'Selected'
  ' listview and if it doesnt find them, it adds them to the
  ' 'Available' listview.

  Dim objRelation As clsMatchRelation
  Dim rsColumns As New Recordset
  Dim rsCalculations As New Recordset
  Dim sSQL As String
  Dim intCount As Integer
  Dim fOK As Boolean
  
  If cboTable1.ListIndex = -1 Then Exit Sub

  ' Clear the contents of the Available Listview
  ListView1.ListItems.Clear
  
  
  ' Add the Columns of the selected table to the listview
  sSQL = "SELECT columnID, tableID, columnName" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & cboTblAvailable.ItemData(cboTblAvailable.ListIndex) & _
    " AND columnType <> " & Trim(Str(colSystem)) & _
    " AND columnType <> " & Trim(Str(colLink)) & _
    " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
    " AND dataType <> " & Trim(Str(sqlOle)) & _
    " ORDER BY columnName"
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  ' Check if the column has already been selected. If so, dont add it
  ' to the available listview
  Do While Not rsColumns.EOF
    If Not AlreadyUsed(CStr("C" & rsColumns!ColumnID)) Then
      ListView1.ListItems.Add , "C" & rsColumns!ColumnID, rsColumns!ColumnName, , ImageList1.ListImages("IMG_TABLE").Index
    End If
    rsColumns.MoveNext
  Loop
  ' Clear recordset reference
  Set rsColumns = Nothing

  
  'For Each objRelation In colRelatedTables
  '  If objRelation.MatchScoreID > 0 Then
      If Not AlreadyUsed("E0") Then
        ListView1.ListItems.Add , "E0", "Match Score", , ImageList1.ListImages("IMG_MATCH").Index
      End If
  '    Exit For
  '  End If
  'Next

  'UpdateButtonStatus SSTab1.Tab


'  ' Skip adding calcs and hide calc button if the table selected is not
'  ' the base table
'  If cboTblAvailable.ItemData(cboTblAvailable.ListIndex) <> cboTable1.ItemData(cboTable1.ListIndex) Then
'    cmdNewCalculation.Visible = False
'    cmdNewCalculation.Enabled = False
'    ListView1.Height = cmdNewCalculation.Top + cmdNewCalculation.Height - ListView1.Top
'    Exit Sub
'  End If
  
  ' Add the Expressions of the selected table to the listview
 
'  sSQL = "SELECT ExprID, Name FROM ASRSysExpressions " & _
'         "WHERE TableID = " & cboTblAvailable.ItemData(cboTblAvailable.ListIndex) & " " & _
'         " AND Type = " & Trim(Str(giEXPR_RUNTIMECALCULATION)) & _
'         " AND ParentComponentID = 0" & _
'         " AND ((Access <> 'HD') OR (Access = 'HD' AND Username = '" & gsUserName & "')) " & _
'         "ORDER BY Name"
'
'  Set rsCalculations = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  ' Check if the column has already been selected. If so, dont add it
'  ' to the available listview
'  Do While Not rsCalculations.EOF
'    If Not AlreadyUsed(CStr("E" & rsCalculations!ExprID)) Then
'      If IsCalcValid(rsCalculations!ExprID) = vbNullString Then
'        ListView1.ListItems.Add , "E" & rsCalculations!ExprID, rsCalculations!Name, , ImageList1.ListImages("IMG_MATCH").Index
'      End If
'    End If
'    rsCalculations.MoveNext
'  Loop
  
  
  ' Bug 688 - If deleted any expressions, remove em from the selected listview

'###########'
'  For intCount = ListView2.ListItems.Count To 1 Step -1
'    If Left(ListView2.ListItems(intCount).Key, 1) = "E" Then
'
'      If Not rsCalculations.BOF Or Not rsCalculations.EOF Then
'        rsCalculations.MoveFirst
'        Do Until rsCalculations.EOF
'          If CStr(Mid(ListView2.ListItems(intCount).Key, 2)) = CStr(rsCalculations.Fields("exprid")) Then
'            fOK = True
'          End If
'          rsCalculations.MoveNext
'        Loop
'      End If
'
'      If fOK = False Then
'        RemoveFromCollection ListView2.ListItems(intCount).Key
'        ListView2.ListItems.Remove ListView2.ListItems(intCount).Key
'      End If
'    End If
'
'    fOK = False
'
'  Next intCount
'###########'
   
'  ' Clear recordset reference
'  Set rsCalculations = Nothing

  ' We are viewing the base table, so adjust the listview height and make
  ' the New Calculation command button visible
'  cmdNewCalculation.Visible = True
'  cmdNewCalculation.Enabled = True
'  ListView1.Height = 4275
'  ListView1.Height = cmdNewCalculation.Top - 100 - ListView1.Top

End Sub

Private Function AlreadyUsed(strKey As String) As Boolean

  Dim objItem As ListItem
  
  For Each objItem In ListView2.ListItems
    If objItem.Key = strKey Then
      AlreadyUsed = True
      Set objItem = Nothing
      Exit Function
    End If
  Next objItem
  
  Set objItem = Nothing
  
End Function


Private Sub ClearForNew()
  
  Dim lngIndex As Long
  
  Set colRelatedTables = Nothing
  Set colRelatedTables = New Collection
  Set mcolMatchReportColDetails = Nothing
  Set mcolMatchReportColDetails = New Collection
  
  For lngIndex = 0 To 1
    optAllRecords(lngIndex).Value = True
    txtPicklist(lngIndex).Text = ""
    txtPicklist(lngIndex).Tag = 0
    txtFilter(lngIndex).Text = ""
    txtFilter(lngIndex).Tag = 0
  Next

  grdRelations.RemoveAll
  ListView2.ListItems.Clear
  grdReportOrder.RemoveAll
  cmdEditOrder.Enabled = False
  cmdDeleteOrder.Enabled = False
  
  If mblnDefinitionCreator Then
    txtUserName = gsUserName
  End If

  ' Default option bit
'  With cboExportTo
'    .Clear
'    .AddItem "Html"
'    .AddItem "Microsoft Excel"
'    .AddItem "Microsoft Word"
'    .ListIndex = 0
'  End With
'
'  optOutput(0).Value = True
'  txtExportFilename.Text = ""
'  chkSave.Value = vbUnchecked
'  chkCloseApplication.Value = vbUnchecked
    
End Sub


Private Sub PopulateTableCombo()

  ' If something has been selected as a base table, this function populates
  ' the Table combo on the Columns Tab with the base table, its parents
  ' and its children.
  
  Dim rsParents As New Recordset
  Dim rsTables As New Recordset
  Dim rsChildren As New Recordset
  Dim sSQL As String
  
    ' Clear the contents of the tables combo
    cboTblAvailable.Clear
    
    ' Clear the listview
    ListView1.ListItems.Clear
    
    ' Add the base table to the top of the combo
    cboTblAvailable.AddItem cboTable1.Text
    cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = cboTable1.ItemData(cboTable1.ListIndex)
  
    ' Add the parents of the base table to the combo
    sSQL = "SELECT ParentID FROM ASRSysRelations WHERE ChildID = " & cboTable1.ItemData(cboTable1.ListIndex)
    Set rsParents = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsParents.EOF
      sSQL = "SELECT TableName, TableID FROM ASRSysTables " & _
             "WHERE TableID = " & rsParents!ParentID & " " & _
             "ORDER BY TableName"
      Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      Do While Not rsTables.EOF
        cboTblAvailable.AddItem rsTables!TableName
        cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = rsTables!TableID
        rsTables.MoveNext
      Loop
      rsParents.MoveNext
    Loop
    
    ' Add the children of the base table to the combo
    sSQL = "SELECT ChildID FROM ASRSysRelations WHERE ParentID = " & cboTable1.ItemData(cboTable1.ListIndex)
    Set rsChildren = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsChildren.EOF
      sSQL = "SELECT TableName, TableID FROM ASRSysTables " & _
             "WHERE TableID = " & rsChildren!ChildID & " " & _
             "ORDER BY TableName"
      Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      Do While Not rsTables.EOF
        cboTblAvailable.AddItem rsTables!TableName
        cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = rsTables!TableID
        rsTables.MoveNext
      Loop
      rsChildren.MoveNext
    Loop
  
    ' Select the base table in the combo by default
    cboTblAvailable.ListIndex = 0
  
End Sub


Private Function GetTableNameFromColumn(lngColumnID As Long) As String

  Dim rsInfo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysTables.TableName " & _
           "FROM ASRSysColumns JOIN ASRSysTables " & _
           "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           "WHERE ColumnID = " & CStr(lngColumnID)

  Set rsInfo = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableNameFromColumn = rsInfo!TableName
  
  Set rsInfo = Nothing

End Function

Private Function GetTableIDFromColumn(lngColumnID As Long) As String

  Dim rsInfo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysTables.TableID " & _
           "FROM ASRSysColumns JOIN ASRSysTables " & _
           "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           "WHERE ColumnID = " & CStr(lngColumnID)

  Set rsInfo = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableIDFromColumn = rsInfo!TableID
  
  Set rsInfo = Nothing

End Function


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

  'set the sql to only include tables for the selected export base table
  'sSQL = "Select Name, PickListID From ASRSysPickListName"
  'sSQL = sSQL & " Where TableID = " & cboTable1.ItemData(cboTable1.ListIndex)
  'sSQL = "TableID = " & cboTable1.ItemData(cboTable1.ListIndex)

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

  If Val(ctlTarget.Tag) > 0 Then
    Set rsTemp = GetSelectionAccess(ctlTarget.Tag, "picklist")
    blnHiddenPicklist = (rsTemp.Fields("Access").Value = "HD")
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  Set frmDefSel = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded

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

Private Function InsertMatchReport(pstrSQL As String, strTableName As String, strIDColumn As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertMatchReport_ERROR

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
    pmADO.Value = strTableName
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = strIDColumn
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      'MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
      MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        gADOCon.Errors(gADOCon.Errors.Count - 1).Description, vbOKOnly + vbExclamation, App.ProductName
      InsertMatchReport = 0
      Set cmADO = Nothing
      Exit Function
    End If
    
    InsertMatchReport = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertMatchReport_ERROR:
  
  fSavedOK = False
  Resume Next

End Function

Private Sub ClearDetailTables(plngMatchReportID As Long)

  ' Delete all column information from the Details table.
  
  Dim sSQL As String
  
  sSQL = "Delete From ASRSysMatchReportDetails Where MatchReportID = " & plngMatchReportID
  datData.ExecuteSql sSQL

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
  If (Len(txtPicklist(0).Tag) > 0) And (Val(txtPicklist(0).Tag) <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtPicklist(0).Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        If fRemove Then
          sBigMessage = "The '" & cboTable1.List(cboTable1.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtPicklist(0).Tag = 0
      txtPicklist(0).Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Base Table Filter
  If Len(txtFilter(0).Tag) > 0 And Val(txtFilter(0).Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtFilter(0).Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)

        If fRemove Then
          sBigMessage = "The '" & cboTable1.List(cboTable1.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboTable1.List(cboTable1.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtFilter(0).Tag = 0
      txtFilter(0).Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Check Match Table Picklist
  If (Len(txtPicklist(1).Tag) > 0) And (Val(txtPicklist(1).Tag) <> 0) Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, CLng(txtPicklist(1).Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Picklist hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)
        If fRemove Then
          sBigMessage = "The '" & cboTable2.List(cboTable2.ListIndex) & "' table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table picklist"
        End If

      Case REC_SEL_VALID_DELETED
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table picklist"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If
      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table picklist"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

    End Select

    If fRemove Then
      ' Picklist invalid, deleted or hidden by another user. Remove it from this definition.
      txtPicklist(1).Tag = 0
      txtPicklist(1).Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Match Table Filter
  If Len(txtFilter(1).Tag) > 0 And Val(txtFilter(1).Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtFilter(1).Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly) And _
          (Not FormPrint)

        If fRemove Then
          sBigMessage = "The '" & cboTable2.List(cboTable2.ListIndex) & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table filter"
  
          fRemove = (Not mblnReadOnly) And _
            (Not FormPrint)
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & cboTable2.List(cboTable2.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtFilter(1).Tag = 0
      txtFilter(1).Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

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




'###################################
Public Sub PrintDef(lMatchReportID As Long)

  Dim objPrintDef As clsPrintDef
  Dim objRelation As clsMatchRelation
  Dim objColumn As clsColumn
  
  Dim sSQL As String
  'Dim lngTempX As Long
  'Dim lngTempY As Long
  Dim sTemp As String
  Dim lngCount As Long
  Dim iLoop As Integer
  Dim varBookmark As Variant

  'mlngMatchReportID = lMatchReportID
  
  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        
        ' First section --------------------------------------------------------
        Select Case mlngMatchReportType
          Case mrtNormal: .PrintHeader "Match Report : " & txtName.Text
          Case mrtSucession: .PrintHeader "Succession Planning : " & txtName.Text
          Case mrtCareer: .PrintHeader "Career Progression : " & txtName.Text
        End Select

        .PrintNormal "Description : " & txtDesc.Text
        .PrintNormal "Owner : " & txtUserName.Text
        
        ' Access section --------------------------------------------------------
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop
        
        ' Data section --------------------------------------------------------
        .PrintTitle "Data"
        
        .PrintNormal "Base Table : " & cboTable1.Text
        If optAllRecords(0).Value = True Then
          .PrintNormal "Records : All Records"
        ElseIf optPicklist(0).Value = True Then
          .PrintNormal "Records : '" & txtPicklist(0).Text & "' picklist"
        ElseIf optFilter(0).Value = True Then
          .PrintNormal "Records : '" & txtFilter(0).Text & "' filter"
        End If
        .PrintNormal
        .PrintNormal "Display filter or picklist title in the report header : " & IIf(chkPrintFilterHeader.Value = vbChecked, "Yes", "No")
        .PrintNormal
    
        .PrintNormal "Match Table : " & cboTable2.Text
        If optAllRecords(1).Value = True Then
          .PrintNormal "Records : All Records"
        ElseIf optPicklist(1).Value = True Then
          .PrintNormal "Records : '" & txtPicklist(1).Text & "' picklist"
        ElseIf optFilter(1).Value = True Then
          .PrintNormal "Records : '" & txtFilter(1).Text & "' filter"
        End If
        .PrintNormal
        
        '----------------------------------------------------------------
        
        .PrintTitle "Related Tables and Breakdown"
        
        For Each objRelation In colRelatedTables
        
          .PrintNormal "Table : " & objRelation.Table1Name
          .PrintNormal "Match Table : " & objRelation.Table2Name
          
          If objRelation.RequiredExprID > 0 Then
            .PrintNormal "Required Matches : " & datGeneral.GetExpression(objRelation.RequiredExprID)
          End If
          If objRelation.PreferredExprID > 0 Then
            .PrintNormal "Preferred Matches : " & datGeneral.GetExpression(objRelation.PreferredExprID)
          End If
          If objRelation.MatchScoreID > 0 Then
            .PrintNormal "Score Calculation : " & datGeneral.GetExpression(objRelation.MatchScoreID)
          End If
        
          .PrintNormal
        
          For Each objColumn In objRelation.BreakdownColumns
            .PrintNormal "Name : " & objColumn.ColumnName
            .PrintNormal "Heading : " & objColumn.Heading
            .PrintNormal "Size : " & objColumn.Size
            
            If objColumn.IsNumeric Then
              .PrintNormal "Decimal Places : " & objColumn.DecPlaces
            End If
          
            .PrintNormal
          Next
        
        Next
    
        '----------------------------------------------------------------
        
        .PrintTitle "Columns"
        
        For Each objColumn In mcolMatchReportColDetails
          
          If objColumn.ColType = "C" Then
            .PrintNormal "Type : Column"
            .PrintNormal "Name : " & objColumn.ColumnName
          Else
            .PrintNormal "Type : Calculation"
            .PrintNormal "Name : Match Score"
          End If
          
          .PrintNormal "Heading : " & objColumn.Heading
          .PrintNormal "Size : " & objColumn.Size
          If objColumn.IsNumeric Then
            .PrintNormal "Decimal Places : " & objColumn.DecPlaces
          End If
    
          .PrintNormal
    
        Next
        
        '----------------------------------------------------------------
      
        .PrintTitle "Sort Order"
          
        For lngCount = 0 To grdReportOrder.Rows - 1
          .PrintNormal "Name : " & grdReportOrder.Columns("Column").CellText(grdReportOrder.AddItemBookmark(lngCount))
          'If grdReportOrder.Columns("Column").CellText(grdReportOrder.AddItemBookmark(lngCount)) = "Ascending" Then
          If grdReportOrder.Columns("Order").CellText(grdReportOrder.AddItemBookmark(lngCount)) = "Ascending" Then
            .PrintNormal "Order : Ascending"
          Else
            .PrintNormal "Order : Descending"
          End If
        Next
      
        .PrintNormal
        
        '----------------------------------------------------------------
          
        .PrintTitle "Output"
        
        .PrintNormal "Matched Records : " & IIf(optHighest.Value, "Highest", "Lowest") & " Match Scores"
        .PrintNormal "Record Count : " & IIf(spnMaxRecords.Value = 0, "All Records", CStr(spnMaxRecords.Value))
        .PrintNormal chkLimit.Caption & " " & IIf(chkLimit.Value = vbChecked, "Yes", "No")
        If chkLimit.Value = vbChecked Then
          .PrintNormal chkLimit.Caption & " " & CStr(spnLimit.Value)
        End If
        .PrintNormal

        If mlngMatchReportType <> mrtNormal Then
          .PrintNormal "Allow progression to equal grade : " & IIf(chkEqualGrade.Value = vbChecked, "Yes", "No")
          .PrintNormal "Restrict by reporting structure : " & IIf(chkReportStructure.Value = vbChecked, "Yes", "No")
          .PrintNormal
        End If

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
        
        Select Case mlngMatchReportType
        Case mrtNormal: .PrintConfirm "Match Report : " & txtName.Text, Me.Caption
        Case mrtSucession: .PrintConfirm "Succession Planning : " & txtName.Text, Me.Caption
        Case mrtCareer: .PrintConfirm "Career Progression : " & txtName.Text, Me.Caption
        End Select
      
      End If
    
    End With
  
  End If
  
Exit Sub

LocalErr:
  MsgBox "Printing " & Me.Caption & "definition failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"

End Sub

Private Sub CheckListViewColWidth(lstvw As ListView)

'  Dim objItem As ListItem
'  Dim lngMax As Long
'  Dim lngLen As Long
'
'  lngMax = 0
'
'  'If lstvw.ListItems.Count = 0 Then Exit Sub
'
'  For Each objItem In lstvw.ListItems
'
'    lngLen = Me.TextWidth(objItem.Text)
'    If lngMax < lngLen Then
'      lngMax = lngLen
'    End If
'
'  Next objItem
'
'  lngMax = lngMax + 60
'  lstvw.ColumnHeaders(1).Width = lngMax

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


Private Sub txtProp_ColumnHeading_Change()
  
  Dim objItem As clsColumn
  
  If Not mblnLoading Then
    Changed = True
    Set objItem = mcolMatchReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Heading = txtProp_ColumnHeading
    Set objItem = Nothing
  End If

End Sub

'Private Sub txtProp_Size_GotFocus()
'
'  UI.txtSelText
'
'End Sub
'
'Private Sub txtProp_Size_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  If (KeyCode < 58 And KeyCode > 47) Or (KeyCode < 106 And KeyCode > 95) Or (KeyCode <> 38) Or (KeyCode <> 40) Then
'    KeyCode = 0
'  End If
'
'  If KeyCode = 38 Then 'up
'    txtProp_Size.Text = txtProp_Size.Text + 1
'    KeyCode = 0
'    UI.txtSelText
'  ElseIf KeyCode = 40 Then 'down
'    If txtProp_Size.Text > 0 Then
'      txtProp_Size.Text = txtProp_Size.Text - 1
'      KeyCode = 0
'      UI.txtSelText
'    End If
'  End If
'
'End Sub
'
'Private Sub txtProp_Size_KeyUp(KeyCode As Integer, Shift As Integer)
'
''  If KeyCode = 38 Then 'up
''    txtProp_Size.Text = txtProp_Size.Text + 1
''    KeyCode = 0
''    UI.txtSelText
''  ElseIf KeyCode = 40 Then 'down
''    txtProp_Size.Text = txtProp_Size.Text - 1
''    KeyCode = 0
''    UI.txtSelText
''  End If
'
'End Sub

Private Sub txtProp_ColumnHeading_GotFocus()

  UI.txtSelText
  
End Sub


Public Sub LoadTable1Combo()
  
  Dim objTableView As CTablePrivilege
  
  With cboTable1
    .Clear
  
    If mlngMatchReportType <> mrtNormal Then
      .AddItem gstrPostTableName
      .ItemData(.NewIndex) = glngPostTableID
      .ListIndex = 0
      .Enabled = False
      .BackColor = vbButtonFace
    Else
      For Each objTableView In gcoTablePrivileges.Collection
        If objTableView.IsTable Then
          .AddItem objTableView.TableName
          .ItemData(.NewIndex) = objTableView.TableID
        End If
      Next
      
      If .ListCount > 0 Then
        SetComboItem cboTable1, glngPersonnelTableID
        If .ListIndex < 0 Then
          .ListIndex = 0
        End If
      End If
    End If
    mstrTable1Name = .List(.ListIndex)
  End With

End Sub



Public Function LoadTable2Combo()

  Dim objTableView As CTablePrivilege
  Dim lngTable1ID As Long

  lngTable1ID = cboTable1.ItemData(cboTable1.ListIndex)

  With cboTable2
    .Clear

    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  
    If mlngMatchReportType <> mrtNormal Then
      .AddItem gsPersonnelTableName
      .ItemData(.NewIndex) = glngPersonnelTableID
      .ListIndex = 1
      .Enabled = False
      .BackColor = vbButtonFace
    Else
      For Each objTableView In gcoTablePrivileges.Collection
        If objTableView.IsTable And objTableView.TableID <> lngTable1ID Then
          .AddItem objTableView.TableName
          .ItemData(.NewIndex) = objTableView.TableID
        End If
      Next
      .ListIndex = 0
    End If

  End With

End Function


Private Sub SetRecordSelection(lngIndex As Integer, blnAllRecords As Boolean, lngPicklist As Long, lngFilter As Long)

  Dim sText As String
  Dim sMessage As String

  ' Set Base Table Record Select Options
  If blnAllRecords Then optAllRecords(lngIndex).Value = True

  If lngPicklist > 0 Then
    optPicklist(lngIndex).Value = True
    sText = IsPicklistValid(lngPicklist)
    If sText <> vbNullString _
      Or (GetPickListField(lngPicklist, "Access") = "HD" And Not mblnDefinitionCreator) Then
      'If Not fAlreadyNotified Then
        If sText = vbNullString Then
          sMessage = "The picklist used in this definition has been made hidden by another user."
        Else
          sMessage = sText
        End If
        
        'If FormPrint Then
        '  sMessage = "Custom Report print failed : " & vbCrLf & vbCrLf & sMessage
        '  MsgBox sMessage, vbExclamation + vbOKOnly, "Match Reports"
        '  Cancelled = True
        '  RetrieveMatchReportDetails = False
        '  Exit Sub
        'End If
        
        MsgBox sMessage & vbCrLf & _
                 "It will be removed from the definition.", vbExclamation + vbOKOnly, Me.Caption
        
      '  fAlreadyNotified = True
      'End If
      txtPicklist(lngIndex).Tag = 0
      txtPicklist(lngIndex).Text = "<None>"
      mblnRecordSelectionInvalid = True
    Else
      txtPicklist(lngIndex).Tag = lngPicklist
      txtPicklist(lngIndex).Text = datGeneral.GetPicklistName(lngPicklist)
    End If

  End If
  
  If lngFilter > 0 Then
    optFilter(lngIndex).Value = True
    sText = IsFilterValid(lngFilter)
    If sText <> vbNullString _
      Or (GetExprField(lngFilter, "Access") = "HD" And Not mblnDefinitionCreator) Then
      'If Not fAlreadyNotified Then
        If sText = vbNullString Then
          sMessage = "The filter used in this definition has been made hidden by another user."
        Else
          sMessage = sText & vbCrLf
        End If

        'If FormPrint Then
        '  sMessage = "Custom Report print failed : " & vbCrLf & vbCrLf & sMessage
        '  MsgBox sMessage, vbExclamation + vbOKOnly, "Match Reports"
        '  Cancelled = True
        '  RetrieveMatchReportDetails = False
        '  Exit Sub
        'End If
        
        MsgBox sMessage & vbCrLf & _
              "It will be removed from the definition.", vbExclamation + vbOKOnly, Me.Caption

        'fAlreadyNotified = True
      'End If
      txtFilter(lngIndex).Tag = 0
      txtFilter(lngIndex).Text = "<None>"
      mblnRecordSelectionInvalid = True

    Else
      txtFilter(lngIndex).Tag = lngFilter
      txtFilter(lngIndex).Text = datGeneral.GetFilterName(lngFilter)
    End If
  End If

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


Private Function GetSortOrder(lngColumnID As Long) As String

  Dim lngCount As Long

  GetSortOrder = "0,''"
  With grdReportOrder
    .Redraw = False
    For lngCount = 0 To .Rows - 1
      .Bookmark = .AddItemBookmark(lngCount)
      If lngColumnID = CLng(.Columns("ColumnID").CellValue(.Bookmark)) Then
        GetSortOrder = CStr(lngCount + 1) & ",'" & Left(.Columns("Order").CellValue(.Bookmark), 1) & "'"
        Exit For
      End If
    Next
    .Redraw = True
  End With

End Function


Private Function CopyExpression(lngBaseTable As Long, lngInputExprID As Long, intType As Integer) As Long
  
  Dim objExpression As clsExprExpression
  Dim lngOutputExprID As Long
  Dim lngIndex As Long
  
  lngOutputExprID = 0
  
  Set objExpression = New clsExprExpression
  With objExpression
    If lngInputExprID > 0 Then
      
      Select Case intType
      Case giEXPR_MATCHWHEREEXPRESSION
        .Initialise lngBaseTable, lngInputExprID, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC
      Case giEXPR_MATCHJOINEXPRESSION
        .Initialise lngBaseTable, lngInputExprID, giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC
      Case Else
        .Initialise lngBaseTable, lngInputExprID, giEXPR_MATCHSCOREEXPRESSION, giEXPRVALUE_NUMERIC
      End Select
      
      .QuickCopyExpression datGeneral.GetExpression(lngInputExprID), True
      If .ExpressionID = 0 Then
        MsgBox "Error copying expression" & vbCrLf & Err.Description, Me.Caption
        Exit Function
      End If
      
      lngOutputExprID = .ExpressionID
      
      ExprDeleteOnCancel lngBaseTable, lngOutputExprID, intType
    
    End If
  End With
  Set objExpression = Nothing
  
  CopyExpression = lngOutputExprID

End Function


Public Function ExprDeleteOnCancel(lngBaseTable As Long, lngExprID As Long, intType As Integer)
  
  Dim lngIndex As Long

  If lngExprID > 0 Then
    lngIndex = UBound(mlngExprDeleteOnCancel, 2) + 1
    ReDim Preserve mlngExprDeleteOnCancel(2, lngIndex)
    mlngExprDeleteOnCancel(0, lngIndex) = lngExprID
    mlngExprDeleteOnCancel(1, lngIndex) = lngBaseTable
    mlngExprDeleteOnCancel(2, lngIndex) = intType
  End If

End Function


Public Function ExprDeleteOnOK(lngBaseTable As Long, lngExprID As Long, intType As Integer)
  
  Dim lngIndex As Long

  If lngExprID > 0 Then
    lngIndex = UBound(mlngExprDeleteOnOK, 2) + 1
    ReDim Preserve mlngExprDeleteOnOK(2, lngIndex)
    mlngExprDeleteOnOK(0, lngIndex) = lngExprID
    mlngExprDeleteOnOK(1, lngIndex) = lngBaseTable
    mlngExprDeleteOnOK(2, lngIndex) = intType
  End If

End Function


Private Sub Form_Unload(Cancel As Integer)
  
  Dim objExpression As clsExprExpression
  Dim lngArray() As Long
  Dim lngCount As Long
  Dim fOK As Boolean
  
  'If user clicks OK then
  '  if any expressions have been cleared then delete them.
  '  if any relations have been deleted then delete the expressions.
  
  'If user clicks cancel then
  '  if any new expressions have been created then delete them.
  '  if any expressions have been copied then delete the copies.
  
  
  If Not mblnCancelled Then
    lngArray = mlngExprDeleteOnOK
  Else
    lngArray = mlngExprDeleteOnCancel
  End If
  
  
  For lngCount = 1 To UBound(lngArray, 2)
    Set objExpression = New clsExprExpression

    Select Case lngArray(2, lngCount)
    Case giEXPR_MATCHWHEREEXPRESSION
      fOK = objExpression.Initialise(lngArray(1, lngCount), lngArray(0, lngCount), giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC)
    Case giEXPR_MATCHJOINEXPRESSION
      fOK = objExpression.Initialise(lngArray(1, lngCount), lngArray(0, lngCount), giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC)
    Case Else
      fOK = objExpression.Initialise(lngArray(1, lngCount), lngArray(0, lngCount), giEXPR_MATCHSCOREEXPRESSION, giEXPRVALUE_NUMERIC)
    End Select

    If fOK Then
      objExpression.DeleteExpression
    End If

    Set objExpression = Nothing
  Next

  Set mobjOutputDef = Nothing

End Sub


Public Function CheckForDuplicateHeadings(colCols As Collection) As String

  Dim objColumn As clsColumn
  Dim strHeading As String
  Dim lngCount1 As Long
  Dim lngCount2 As Long

  CheckForDuplicateHeadings = vbNullString
  For lngCount1 = 1 To colCols.Count
    For lngCount2 = lngCount1 To colCols.Count

      If Not (colCols(lngCount1) Is colCols(lngCount2)) Then
        If LCase(colCols(lngCount1).Heading) = LCase(colCols(lngCount2).Heading) Then
          CheckForDuplicateHeadings = colCols(lngCount1).Heading
          Exit Function
        End If
      End If

    Next
  Next

End Function

