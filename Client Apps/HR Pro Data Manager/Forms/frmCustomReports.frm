VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmCustomReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Report Definition"
   ClientHeight    =   6600
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
   HelpContextID   =   1023
   Icon            =   "frmCustomReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDocument 
      DragIcon        =   "frmCustomReports.frx":000C
      Height          =   480
      Index           =   0
      Left            =   4590
      Picture         =   "frmCustomReports.frx":0596
      ScaleHeight     =   420
      ScaleWidth      =   480
      TabIndex        =   122
      Top             =   6075
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picDocument 
      DragIcon        =   "frmCustomReports.frx":0E60
      Height          =   465
      Index           =   1
      Left            =   5130
      Picture         =   "frmCustomReports.frx":13EA
      ScaleHeight     =   405
      ScaleWidth      =   465
      TabIndex        =   121
      Top             =   6090
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7050
      TabIndex        =   115
      Top             =   6100
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8330
      TabIndex        =   116
      Top             =   6100
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   50
      TabIndex        =   117
      Top             =   50
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   10504
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
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
      TabPicture(0)   =   "frmCustomReports.frx":1CB4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraInformation"
      Tab(0).Control(1)=   "fraBase"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Related Ta&bles"
      TabPicture(1)   =   "frmCustomReports.frx":1CD0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraChild"
      Tab(1).Control(1)=   "fraParent2"
      Tab(1).Control(2)=   "fraParent1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Colu&mns"
      TabPicture(2)   =   "frmCustomReports.frx":1CEC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraButtons"
      Tab(2).Control(1)=   "fraFieldsSelected"
      Tab(2).Control(2)=   "fraFieldsAvailable"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&Sort Order"
      TabPicture(3)   =   "frmCustomReports.frx":1D08
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraReportOrder"
      Tab(3).Control(1)=   "fraRepetition"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "O&utput"
      TabPicture(4)   =   "frmCustomReports.frx":1D24
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fraReportOptions"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fraOutputFormat"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraOutputDestination"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      Begin VB.Frame fraReportOrder 
         Caption         =   "Sort Order :"
         Enabled         =   0   'False
         Height          =   3435
         Left            =   -74850
         TabIndex        =   74
         Top             =   400
         Width           =   9180
         Begin VB.CommandButton cmdClearOrder 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   7800
            TabIndex        =   79
            Top             =   1740
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddOrder 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7800
            TabIndex        =   76
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteOrder 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   7800
            TabIndex        =   78
            Top             =   1260
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditOrder 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   7800
            TabIndex        =   77
            Top             =   780
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveUpOrder 
            Caption         =   "Move U&p"
            Height          =   400
            Left            =   7800
            TabIndex        =   80
            Top             =   2360
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveDownOrder 
            Caption         =   "Move Do&wn"
            Height          =   400
            Left            =   7800
            TabIndex        =   81
            Top             =   2840
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdReportOrder 
            Height          =   2940
            Left            =   195
            TabIndex        =   75
            Top             =   300
            Width           =   7470
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            GroupHeadLines  =   0
            HeadLines       =   2
            Col.Count       =   7
            stylesets.count =   6
            stylesets(0).Name=   "ssetSelected"
            stylesets(0).ForeColor=   -2147483634
            stylesets(0).BackColor=   -2147483635
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
            stylesets(0).Picture=   "frmCustomReports.frx":1D40
            stylesets(1).Name=   "ssetHeaderDisabled"
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
            stylesets(1).Picture=   "frmCustomReports.frx":1D5C
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
            stylesets(2).Picture=   "frmCustomReports.frx":1D78
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
            stylesets(3).Picture=   "frmCustomReports.frx":1D94
            stylesets(4).Name=   "ssetCheckBoxSelected"
            stylesets(4).ForeColor=   -2147483640
            stylesets(4).BackColor=   -2147483635
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
            stylesets(4).Picture=   "frmCustomReports.frx":1DB0
            stylesets(5).Name=   "ssetDisabled"
            stylesets(5).ForeColor=   -2147483631
            stylesets(5).BackColor=   -2147483633
            stylesets(5).HasFont=   -1  'True
            BeginProperty stylesets(5).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(5).Picture=   "frmCustomReports.frx":1DCC
            CheckBox3D      =   0   'False
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   0
            StyleSet        =   "ssetDisabled"
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   7
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ColumnID"
            Columns(0).Name =   "ColumnID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3598
            Columns(1).Caption=   "Column"
            Columns(1).Name =   "Column"
            Columns(1).CaptionAlignment=   2
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   1270
            Columns(2).Caption=   "Order"
            Columns(2).Name =   "Order"
            Columns(2).AllowSizing=   0   'False
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   4
            Columns(2).Locked=   -1  'True
            Columns(2).Style=   3
            Columns(2).Row.Count=   2
            Columns(2).Col.Count=   2
            Columns(2).Row(0).Col(0)=   "Asc"
            Columns(2).Row(1).Col(0)=   "Desc"
            Columns(3).Width=   1773
            Columns(3).Caption=   "Break on Change"
            Columns(3).Name =   "Break"
            Columns(3).CaptionAlignment=   2
            Columns(3).AllowSizing=   0   'False
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   11
            Columns(3).FieldLen=   1
            Columns(3).Style=   2
            Columns(4).Width=   1826
            Columns(4).Caption=   "Page on Change"
            Columns(4).Name =   "Page"
            Columns(4).CaptionAlignment=   2
            Columns(4).AllowSizing=   0   'False
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   11
            Columns(4).FieldLen=   256
            Columns(4).Style=   2
            Columns(5).Width=   1826
            Columns(5).Caption=   "Value on Change"
            Columns(5).Name =   "Value"
            Columns(5).CaptionAlignment=   2
            Columns(5).AllowSizing=   0   'False
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   11
            Columns(5).FieldLen=   256
            Columns(5).Style=   2
            Columns(6).Width=   2831
            Columns(6).Caption=   "Suppress Repeated Values"
            Columns(6).Name =   "Hide"
            Columns(6).CaptionAlignment=   2
            Columns(6).AllowSizing=   0   'False
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   11
            Columns(6).FieldLen=   256
            Columns(6).Style=   2
            _ExtentX        =   13176
            _ExtentY        =   5186
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
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   3975
         Left            =   2745
         TabIndex        =   95
         Top             =   1800
         Width           =   6555
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3270
            TabIndex        =   114
            Tag             =   "0"
            Top             =   3460
            Width           =   3120
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   103
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1760
            Width           =   2790
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3270
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   1240
            Width           =   3135
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3270
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3270
            TabIndex        =   112
            Top             =   3060
            Width           =   3120
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   109
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   2660
            Width           =   2790
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":1DE8
            Enabled         =   0   'False
            Height          =   315
            Left            =   6060
            Picture         =   "frmCustomReports.frx":2149
            TabIndex        =   104
            Top             =   1760
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":24AA
            Enabled         =   0   'False
            Height          =   315
            Left            =   6060
            Picture         =   "frmCustomReports.frx":280B
            TabIndex        =   110
            Top             =   2660
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   97
            Top             =   850
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   98
            Top             =   1300
            Width           =   1650
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   101
            Top             =   1820
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   107
            Top             =   2720
            Width           =   1560
         End
         Begin VB.CheckBox chkPreview 
            Caption         =   "P&review on screen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   96
            Top             =   400
            Width           =   3495
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   1800
            TabIndex        =   113
            Top             =   3525
            Width           =   1020
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   102
            Top             =   1815
            Width           =   1005
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   111
            Top             =   3120
            Width           =   1305
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   108
            Top             =   2715
            Width           =   1200
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
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   99
            Top             =   1305
            Width           =   1410
         End
      End
      Begin VB.Frame fraButtons 
         BorderStyle     =   0  'None
         Caption         =   "Buttons :"
         Height          =   5385
         Left            =   -71040
         TabIndex        =   53
         Top             =   400
         Width           =   1515
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   400
            Left            =   120
            TabIndex        =   54
            Top             =   1020
            Width           =   1275
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   120
            TabIndex        =   56
            Top             =   2220
            Width           =   1275
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "&Up"
            Enabled         =   0   'False
            Height          =   400
            Left            =   120
            TabIndex        =   58
            Top             =   3420
            Width           =   1275
         End
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   120
            TabIndex        =   59
            Top             =   3900
            Width           =   1275
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add A&ll"
            Height          =   400
            Left            =   120
            TabIndex        =   55
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   120
            TabIndex        =   57
            Top             =   2700
            Width           =   1275
         End
      End
      Begin VB.Frame fraChild 
         Caption         =   "Child Tables : "
         Height          =   2400
         Left            =   -74850
         TabIndex        =   41
         Top             =   3390
         Width           =   9180
         Begin VB.CommandButton cmdEditChild 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   7800
            TabIndex        =   44
            Top             =   790
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveAllChilds 
            Caption         =   "Remo&ve All "
            Height          =   400
            Left            =   7800
            TabIndex        =   46
            Top             =   1770
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddChild 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7800
            TabIndex        =   43
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveChild 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   7800
            TabIndex        =   45
            Top             =   1280
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdChildren 
            Height          =   1875
            Left            =   195
            TabIndex        =   42
            Top             =   300
            Width           =   7455
            ScrollBars      =   0
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   7
            stylesets.count =   5
            stylesets(0).Name=   "ssetSelected"
            stylesets(0).ForeColor=   -2147483634
            stylesets(0).BackColor=   -2147483635
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
            stylesets(0).Picture=   "frmCustomReports.frx":2B6C
            stylesets(1).Name=   "ssetHeaderDisabled"
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
            stylesets(1).Picture=   "frmCustomReports.frx":2B88
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
            stylesets(2).Picture=   "frmCustomReports.frx":2BA4
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
            stylesets(3).Picture=   "frmCustomReports.frx":2BC0
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
            stylesets(4).Picture=   "frmCustomReports.frx":2BDC
            AllowUpdate     =   0   'False
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
            SelectTypeRow   =   0
            BalloonHelp     =   0   'False
            RowNavigation   =   2
            MaxSelectedRows =   1
            StyleSet        =   "ssetDisabled"
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
            Columns(1).Width=   4233
            Columns(1).Caption=   "Table"
            Columns(1).Name =   "Table"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3466
            Columns(2).Caption=   "Order"
            Columns(2).Name =   "Order"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "OrderID"
            Columns(3).Name =   "OrderID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "FilterID"
            Columns(4).Name =   "FilterID"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3519
            Columns(5).Caption=   "Filter"
            Columns(5).Name =   "Filter"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   1879
            Columns(6).Caption=   "Records"
            Columns(6).Name =   "Records"
            Columns(6).Alignment=   2
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13150
            _ExtentY        =   3307
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
      Begin VB.Frame fraOutputFormat 
         Caption         =   "Output Format :"
         Height          =   3975
         Left            =   120
         TabIndex        =   87
         Top             =   1800
         Width           =   2500
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel P&ivot Table"
            Height          =   195
            Index           =   6
            Left            =   200
            TabIndex        =   94
            Top             =   2800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Char&t"
            Height          =   195
            Index           =   5
            Left            =   200
            TabIndex        =   93
            Top             =   2400
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   92
            Top             =   2000
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   91
            Top             =   1600
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   90
            Top             =   1200
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CS&V File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   89
            Top             =   800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "D&ata Only"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   88
            Top             =   400
            Value           =   -1  'True
            Width           =   1900
         End
      End
      Begin VB.Frame fraRepetition 
         Caption         =   "Repetition : "
         Height          =   1830
         Left            =   -74850
         TabIndex        =   83
         Top             =   3960
         Width           =   9180
         Begin SSDataWidgets_B.SSDBGrid grdRepetition 
            Height          =   1335
            Left            =   195
            TabIndex        =   82
            Top             =   300
            Width           =   8865
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            GroupHeadLines  =   0
            Col.Count       =   3
            stylesets.count =   5
            stylesets(0).Name=   "ssetSelected"
            stylesets(0).ForeColor=   -2147483634
            stylesets(0).BackColor=   -2147483635
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
            stylesets(0).Picture=   "frmCustomReports.frx":2BF8
            stylesets(1).Name=   "ssetHeaderDisabled"
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
            stylesets(1).Picture=   "frmCustomReports.frx":2C14
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
            stylesets(2).Picture=   "frmCustomReports.frx":2C30
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
            stylesets(3).Picture=   "frmCustomReports.frx":2C4C
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
            stylesets(4).Picture=   "frmCustomReports.frx":2C68
            CheckBox3D      =   0   'False
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
            SelectTypeRow   =   0
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            StyleSet        =   "ssetDisabled"
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ColumnID"
            Columns(0).Name =   "ColumnID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   4128
            Columns(1).Caption=   "Column"
            Columns(1).Name =   "Column"
            Columns(1).CaptionAlignment=   2
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   1852
            Columns(2).Caption=   "Repetition"
            Columns(2).Name =   "Repetition"
            Columns(2).CaptionAlignment=   2
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(2).Style=   2
            _ExtentX        =   15637
            _ExtentY        =   2355
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
         Height          =   3340
         Left            =   -74850
         TabIndex        =   9
         Top             =   2450
         Width           =   9180
         Begin VB.CheckBox chkPrintFilterHeader 
            Caption         =   "Display &title in the report header"
            Height          =   240
            Left            =   4770
            TabIndex        =   20
            Top             =   1560
            Width           =   3960
         End
         Begin VB.CommandButton cmdBaseFilter 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":2C84
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            Picture         =   "frmCustomReports.frx":2FE5
            TabIndex        =   19
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdBasePicklist 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":3346
            Enabled         =   0   'False
            Height          =   315
            Left            =   8685
            Picture         =   "frmCustomReports.frx":36A7
            TabIndex        =   16
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   3090
         End
         Begin VB.OptionButton optBaseAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   5715
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton optBasePicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   5715
            TabIndex        =   14
            Top             =   760
            Width           =   885
         End
         Begin VB.OptionButton optBaseFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   5715
            TabIndex        =   17
            Top             =   1160
            Width           =   795
         End
         Begin VB.TextBox txtBasePicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   15
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
            TabIndex        =   18
            Tag             =   "0"
            Top             =   1100
            Width           =   2000
         End
         Begin VB.Label lblBaseRecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   4770
            TabIndex        =   12
            Top             =   360
            Width           =   870
         End
         Begin VB.Label lblBaseTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   10
            Top             =   360
            Width           =   1110
         End
      End
      Begin VB.Frame fraParent2 
         Caption         =   "Parent 2 :"
         Height          =   1440
         Left            =   -74850
         TabIndex        =   31
         Top             =   1890
         Width           =   9180
         Begin VB.CommandButton cmdParent2Picklist 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":3A08
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            Picture         =   "frmCustomReports.frx":3D69
            TabIndex        =   37
            Top             =   535
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtParent2Picklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   36
            Tag             =   "0"
            Top             =   535
            Width           =   2000
         End
         Begin VB.OptionButton optParent2AllRecords 
            Caption         =   "All"
            Height          =   195
            Left            =   5760
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton optParent2Picklist 
            Caption         =   "Picklist"
            Height          =   195
            Left            =   5760
            TabIndex        =   35
            Top             =   595
            Width           =   885
         End
         Begin VB.OptionButton optParent2Filter 
            Caption         =   "Filter"
            Height          =   195
            Left            =   5760
            TabIndex        =   38
            Top             =   985
            Width           =   795
         End
         Begin VB.TextBox txtParent2 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1000
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   300
            Width           =   3000
         End
         Begin VB.TextBox txtParent2Filter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   925
            Width           =   2000
         End
         Begin VB.CommandButton cmdParent2Filter 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":40CA
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            Picture         =   "frmCustomReports.frx":442B
            TabIndex        =   40
            Top             =   925
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.Label lblParent2Records 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   4860
            TabIndex        =   120
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblParent2Table 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   200
            TabIndex        =   32
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraParent1 
         Caption         =   "Parent 1 :"
         Height          =   1440
         Left            =   -74850
         TabIndex        =   21
         Top             =   400
         Width           =   9180
         Begin VB.OptionButton optParent1Filter 
            Caption         =   "Filter"
            Height          =   195
            Left            =   5760
            TabIndex        =   28
            Top             =   985
            Width           =   840
         End
         Begin VB.OptionButton optParent1Picklist 
            Caption         =   "Picklist"
            Height          =   195
            Left            =   5760
            TabIndex        =   25
            Top             =   595
            Width           =   885
         End
         Begin VB.OptionButton optParent1AllRecords 
            Caption         =   "All"
            Height          =   195
            Left            =   5760
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   630
         End
         Begin VB.TextBox txtParent1Picklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   26
            Tag             =   "0"
            Top             =   535
            Width           =   2000
         End
         Begin VB.CommandButton cmdParent1Picklist 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":478C
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            Picture         =   "frmCustomReports.frx":4AED
            TabIndex        =   27
            Top             =   535
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtParent1 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1000
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   300
            Width           =   3000
         End
         Begin VB.TextBox txtParent1Filter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6700
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   925
            Width           =   2000
         End
         Begin VB.CommandButton cmdParent1Filter 
            Caption         =   "..."
            DisabledPicture =   "frmCustomReports.frx":4E4E
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            Picture         =   "frmCustomReports.frx":51AF
            TabIndex        =   30
            Top             =   925
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.Label lblParent1Records 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   4860
            TabIndex        =   119
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblParent1Table 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   200
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraFieldsSelected 
         Caption         =   "Columns / Calculations Selected :"
         Height          =   5385
         Left            =   -69300
         TabIndex        =   60
         Top             =   400
         Width           =   3615
         Begin VB.CheckBox chkProp_Group 
            Caption         =   "&Group with Next"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1545
            TabIndex        =   73
            Top             =   4950
            Width           =   1695
         End
         Begin VB.CheckBox chkProp_Hidden 
            Caption         =   "&Hidden"
            Enabled         =   0   'False
            Height          =   255
            Left            =   260
            TabIndex        =   72
            Top             =   4950
            Width           =   1095
         End
         Begin VB.TextBox txtProp_ColumnHeading 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   63
            Top             =   3390
            Width           =   2115
         End
         Begin VB.CheckBox chkProp_Average 
            Caption         =   "A&verage"
            Enabled         =   0   'False
            Height          =   195
            Left            =   260
            TabIndex        =   69
            Top             =   4650
            Width           =   1080
         End
         Begin VB.CheckBox chkProp_Count 
            Caption         =   "C&ount"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1545
            TabIndex        =   70
            Top             =   4650
            Width           =   840
         End
         Begin VB.CheckBox chkProp_Total 
            Caption         =   "&Total"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2640
            TabIndex        =   71
            Top             =   4650
            Width           =   825
         End
         Begin VB.CheckBox chkProp_IsNumeric 
            Enabled         =   0   'False
            Height          =   195
            Left            =   2880
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   3480
            Visible         =   0   'False
            Width           =   285
         End
         Begin COASpinner.COA_Spinner spnSize 
            Height          =   315
            Left            =   1260
            TabIndex        =   65
            Top             =   3795
            Width           =   1185
            _ExtentX        =   2090
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
            Height          =   3000
            Left            =   200
            TabIndex        =   61
            Top             =   300
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   5292
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
            Left            =   1260
            TabIndex        =   68
            Top             =   4185
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.Label lblProp_Decimals 
            BackStyle       =   0  'Transparent
            Caption         =   "Decimals :"
            Height          =   195
            Left            =   260
            TabIndex        =   67
            Top             =   4245
            Width           =   1140
         End
         Begin VB.Label lblProp_Size 
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            Height          =   195
            Left            =   255
            TabIndex        =   64
            Top             =   3855
            Width           =   615
         End
         Begin VB.Label lblProp_ColumnHeading 
            BackStyle       =   0  'Transparent
            Caption         =   "Heading :"
            Height          =   195
            Left            =   260
            TabIndex        =   62
            Top             =   3450
            Width           =   1260
         End
      End
      Begin VB.Frame fraFieldsAvailable 
         Caption         =   "Columns / Calculations Available :"
         Height          =   5385
         Left            =   -74850
         TabIndex        =   47
         Top             =   400
         Width           =   3615
         Begin VB.ComboBox cboTblAvailable 
            Height          =   315
            Left            =   200
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   300
            Width           =   3225
         End
         Begin VB.CommandButton cmdNewCalculation 
            Caption         =   "Calculat&ion Definitions..."
            Height          =   400
            Left            =   200
            TabIndex        =   52
            Top             =   4800
            Visible         =   0   'False
            Width           =   3225
         End
         Begin ComctlLib.ListView ListView1 
            Height          =   3650
            Left            =   200
            TabIndex        =   51
            Top             =   1050
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   6429
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   5644
            EndProperty
         End
         Begin VB.OptionButton optColumns 
            Caption         =   "&Columns"
            Height          =   255
            Left            =   390
            TabIndex        =   49
            Top             =   740
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton optCalc 
            Caption         =   "Calculatio&ns"
            Height          =   255
            Left            =   1750
            TabIndex        =   50
            Top             =   740
            Width           =   1335
         End
      End
      Begin VB.Frame fraReportOptions 
         Caption         =   "Report Options :"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   120
         TabIndex        =   84
         Top             =   400
         Width           =   9180
         Begin VB.CheckBox chkIgnoreZeros 
            Caption         =   "Ignore &zeros when calculating aggregates"
            Height          =   255
            Left            =   200
            TabIndex        =   86
            Top             =   720
            Width           =   3975
         End
         Begin VB.CheckBox chkSummaryReport 
            Caption         =   "Summary report"
            Height          =   240
            Left            =   200
            TabIndex        =   85
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame fraInformation 
         Height          =   1950
         Left            =   -74850
         TabIndex        =   0
         Top             =   400
         Width           =   9180
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1395
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   700
            Width           =   3090
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   2
            Top             =   300
            Width           =   3090
         End
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5625
            MaxLength       =   30
            TabIndex        =   6
            Top             =   300
            Width           =   3405
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1080
            Left            =   5625
            TabIndex        =   8
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
            stylesets(0).Picture=   "frmCustomReports.frx":5510
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
            stylesets(1).Picture=   "frmCustomReports.frx":552C
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
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   4770
            TabIndex        =   7
            Top             =   750
            Width           =   825
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   200
            TabIndex        =   3
            Top             =   750
            Width           =   900
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   200
            TabIndex        =   1
            Top             =   360
            Width           =   510
         End
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Left            =   4770
            TabIndex        =   5
            Top             =   360
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox picNoDrop 
      Height          =   495
      Left            =   4035
      Picture         =   "frmCustomReports.frx":5548
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   6075
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2910
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1005
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCustomReports.frx":5E12
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCustomReports.frx":6364
            Key             =   "IMG_CALC"
         EndProperty
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   375
      Top             =   6120
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
      Bands           =   "frmCustomReports.frx":68B6
   End
End
Attribute VB_Name = "frmCustomReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DataAccess Class
Private datData As HRProDataMgr.clsDataAccess
Private mobjOutputDef As clsOutputDef

' Collection Class (Holds column details such as heading, size etc)
Public mcolCustomReportColDetails As clsColumns

' Long to hold current ReportID
Private mlngCustomReportID As Long

' Variables to hold current (or previously) selected table details
Private mstrBaseTable As String
Private mstrChildTable As String
Private mlngChildTable As Long
Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mblnFormPrint As Boolean
Private mblnColumnDrag As Boolean
'Private mblnChanged As Boolean
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

Private mblnParent1Enabled As Boolean
Private mblnParent2Enabled As Boolean
Private mblnChildsEnabled As Boolean

Private mblnResizingColumn As Boolean

Private mblnIsChildColumnSelected As Boolean

Private mblnGridActionCancelled As Boolean
Private mblnGridChangeRecursive As Boolean


Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOK.Enabled = (pblnChanged) And (Not mblnReadOnly)
End Property





Private Sub ClearRepetition()
  
  Dim objColumn As clsColumn
  Dim iRow As Integer
  
  For Each objColumn In mcolCustomReportColDetails.Collection
    If Not objColumn Is Nothing Then
      objColumn.Repetition = False
    End If
  Next objColumn
  
  Set objColumn = Nothing
  
  With grdRepetition
    .Redraw = False
    If .Rows > 0 Then
      .MoveFirst
    End If
  
    For iRow = 0 To .Rows - 1
      .Columns("repetition").Value = vbUnchecked
      .MoveNext
    Next iRow
    
    If .Rows > 0 Then
      .MoveFirst
    End If
    .Redraw = True
  End With
  
End Sub

Private Function IsChildColumnSelected() As Boolean

  Dim intColCount As Integer
  Dim intChildCount As Integer
  Dim bm As Variant
  Dim lngTableID As Long
  Dim varOriginalBM As Variant
  
  If grdChildren.Rows < 1 Then
    IsChildColumnSelected = False
    Exit Function
  End If
  
  varOriginalBM = grdChildren.Bookmark
  
  For intColCount = 1 To ListView2.ListItems.Count Step 1
    
    'TM20030529 Fault - check for Child Calculations also.
    If UCase(Left(ListView2.ListItems(intColCount).Key, 1)) = "C" Then
      lngTableID = GetTableIDFromColumn(CLng(Right(ListView2.ListItems(intColCount).Key, Len(ListView2.ListItems(intColCount).Key) - 1)))
    Else
      lngTableID = GetExprField(CLng(Right(ListView2.ListItems(intColCount).Key, Len(ListView2.ListItems(intColCount).Key) - 1)), "TableID")
    End If
    
    grdChildren.Redraw = False
    grdChildren.MoveFirst
    For intChildCount = 0 To grdChildren.Rows - 1 Step 1
      bm = grdChildren.AddItemBookmark(intChildCount)
      If lngTableID = CLng(grdChildren.Columns("TableID").CellValue(bm)) Then
        IsChildColumnSelected = True

        grdChildren.SelBookmarks.RemoveAll
        grdChildren.Bookmark = varOriginalBM
        grdChildren.SelBookmarks.Add grdChildren.Bookmark

        grdChildren.Redraw = True
        Exit Function
      End If
'TM29012004 Fault 7208 fixed.
'      grdChildren.MoveNext
    Next intChildCount
    
  Next intColCount
  
  grdChildren.SelBookmarks.RemoveAll
  grdChildren.Bookmark = varOriginalBM
  grdChildren.SelBookmarks.Add grdChildren.Bookmark
  
  grdChildren.Redraw = True
  IsChildColumnSelected = False
  
End Function

Private Sub RefreshRepetitionGrid()
  
  Dim objItem As ListItem
  Dim iLoop As Integer
  Dim sKey As String
  Dim objColumm  As clsColumn
  
  If mblnLoading Then
    Exit Sub
  End If
  
  mblnIsChildColumnSelected = IsChildColumnSelected
  
  FormatGridColumnWidths
  
  With grdRepetition
    .Enabled = True
    .AllowUpdate = ((Not mblnReadOnly) And (mblnIsChildColumnSelected))
    .Columns(1).Locked = True
    .CheckBox3D = False
    
    If mblnReadOnly Or (Not mblnIsChildColumnSelected) Then
      If (Not mblnIsChildColumnSelected) Then
        ClearRepetition
      End If
      .HeadStyleSet = "ssetHeaderDisabled"
      .StyleSet = "ssetDisabled"
'      .ActiveRowStyleSet = "ssetDisabled"
      .RowNavigation = ssRowNavigationAllLock
      .SelectTypeRow = ssSelectionTypeNone
      .SelectByCell = False
      .SelectTypeCol = ssSelectionTypeNone
    Else
      .HeadStyleSet = "ssetHeaderEnabled"
      .StyleSet = "ssetEnabled"
'      .ActiveRowStyleSet = "ssetEnabled"
      .RowNavigation = ssRowNavigationLRLock
      .SelectTypeRow = ssSelectionTypeNone
      .SelectByCell = False
      .SelectTypeCol = ssSelectionTypeNone

      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End If
    
  End With
 
  grdRepetition_RowColChange 0, 0

End Sub

Private Sub RefreshChildrenGrid()
  
  With grdChildren
    .Enabled = True
    
    If mblnReadOnly Or (Not mblnChildsEnabled) Then
      .HeadStyleSet = "ssetHeaderDisabled"
      .StyleSet = "ssetDisabled"
      .ActiveRowStyleSet = "ssetDisabled"
      .SelectTypeRow = ssSelectionTypeNone
    Else
      .HeadStyleSet = "ssetHeaderEnabled"
      .StyleSet = "ssetEnabled"
      .ActiveRowStyleSet = "ssetSelected"
      .SelectTypeRow = ssSelectionTypeSingleSelect
      .RowNavigation = ssRowNavigationLRLock
      
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End If
    
  End With

  UpdateButtonStatus (SSTab1.Tab)
  
End Sub

Private Sub RefreshReportOrderGrid()
  
  With grdReportOrder
    .Refresh
    .Enabled = True
    .AllowUpdate = (Not mblnReadOnly)
    .Columns(1).Locked = True
    .CheckBox3D = False
    
    If mblnReadOnly Then
      .HeadStyleSet = "ssetHeaderDisabled"
      .StyleSet = "ssetDisabled"
      .ActiveRowStyleSet = "ssetDisabled"
      .SelectTypeRow = ssSelectionTypeNone
      .SelectTypeCol = ssSelectionTypeNone
      .RowNavigation = ssRowNavigationAllLock
    Else
      .HeadStyleSet = "ssetHeaderEnabled"
      .StyleSet = "ssetEnabled"
      .ActiveRowStyleSet = "ssetSelected"
      .SelectTypeRow = ssSelectionTypeSingleSelect
      .SelectByCell = False
      .SelectTypeCol = ssSelectionTypeNone
      .RowNavigation = ssRowNavigationLRLock

      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End If

  End With

  grdReportOrder_RowColChange 0, 0

  UpdateOrderButtons
  
End Sub

Private Function FormatGridColumnWidths() As Boolean

  Dim iColumn As Integer
  Dim iRow As Integer
  Dim lngTextWidth As Long
  Dim varBookmark As Variant
  Dim varOriginalPos As Variant

  lngTextWidth = 0
  With grdRepetition
    varOriginalPos = .Bookmark
    
    .Redraw = False
    .MoveFirst
    For iColumn = 0 To .Columns.Count - 1 Step 1
      
      lngTextWidth = Me.TextWidth(.Columns(iColumn).Caption)
      
      If .Columns(iColumn).Visible Then
        For iRow = 0 To .Rows - 1 Step 1
          varBookmark = .AddItemBookmark(iRow)
          
          If Me.TextWidth(Trim(.Columns(iColumn).CellText(varBookmark))) > lngTextWidth Then
            lngTextWidth = Me.TextWidth(Trim(.Columns(iColumn).CellText(varBookmark)))
          End If
        Next iRow
        
        .Columns(iColumn).Width = lngTextWidth + 195
      End If
      lngTextWidth = 0
    Next iColumn
    
    .Bookmark = varOriginalPos
    .Redraw = True
  End With
  
End Function




Public Property Get SelectedID() As Long
  SelectedID = mlngCustomReportID
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

Private Sub cboBaseTable_Click()
  
  ' When the user changes the Base Table, check to see if the user
  ' has defined any columns in the report. If they have, check that
  ' they have selected a different table in the combo to the one that
  ' was there before. If so, then prompt user, otherwise, go ahead and
  ' clear the definition
  If mblnLoading = True Then Exit Sub
  If mstrBaseTable = Me.cboBaseTable.Text And (mblnLoading = False) Then Exit Sub
  
  'If (mstrBaseTable <> Me.cboBaseTable.Text) Or (mblnLoading = True) Then
    If Me.ListView2.ListItems.Count > 0 Or Me.grdChildren.Rows > 0 _
      Or Me.optBaseFilter.Value Or optBaseFilter.Value Or optParent1Filter.Value _
      Or Me.optParent1Picklist.Value Or Me.optParent2Filter.Value Or Me.optParent2Picklist.Value Then
      If MsgBox("Warning: Changing the base table will result in all table/column " & _
            "specific aspects of this report definition being cleared." & vbCrLf & _
            "Are you sure you wish to continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Custom Reports") = vbYes Then
    
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
  'End If
  
  mstrBaseTable = Me.cboBaseTable.Text
  
  '01/08/2001 MH Fault 2615
  optBaseAllRecords.Value = True
  
  mcolCustomReportColDetails.RemoveAll ' this leaves 1 when u check the count prop!!!
  Set mcolCustomReportColDetails = New clsColumns
  
  UpdateDependantFields
  PopulateTableAvailable , True
  ForceDefinitionToBeHiddenIfNeeded
  
End Sub

Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional plngCustomReportID As Long, Optional bPrint As Boolean) As Boolean

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
    mlngCustomReportID = 0

    grdRepetition.RowHeight = lng_GRIDROWHEIGHT
    
    'Set controls to defaults
    ClearForNew
    
    'Load All Possible Base Tables into combo
    LoadBaseCombo

    UpdateDependantFields
    
    PopulateTableAvailable , True

    ' Set command button status
    UpdateButtonStatus (Me.SSTab1.Tab)
    
    PopulateAccessGrid
    
    Changed = False
  Else
    ' Make the CustomReportID visible to the rest of the module
    mlngCustomReportID = plngCustomReportID
    
    ' Is is a copy of an existing one ?
    FromCopy = bCopy
    
    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint
    
    PopulateAccessGrid
    
    If Not RetrieveCustomReportDetails(plngCustomReportID) Then
      If mblnDeleted Or Me.Cancelled Then
        Initialise = False
        Exit Function
      Else
        If MsgBox("HR Pro could not load all of the definition successfully. The recommendation is that" & vbCrLf & _
               "you delete the definition and create a new one, however, you may edit the existing" & vbCrLf & _
               "definition if you wish. Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, "Custom Reports") = vbNo Then
          Me.Cancelled = True
          Initialise = False
          Exit Function
        End If
      End If
    End If
    
    UpdateOrderButtons
    
    If bCopy = True Then
      mlngCustomReportID = 0
      Changed = True
    Else
      Changed = mblnRecordSelectionInvalid And (Not mblnReadOnly) ' False
    End If
    
  End If
  
  EnableDisableTabControls
    
  If mblnForceHidden Then
    mblnForceHidden = True
  End If
  Cancelled = False
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
  Set rsAccess = GetUtilityAccessRecords(utlCustomReport, mlngCustomReportID, mblnFromCopy)
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

Private Sub cboTblAvailable_Click()
  PopulateAvailable
  UpdateButtonStatus (Me.SSTab1.Tab)
End Sub

Private Sub chkIgnoreZeros_Click()
  Changed = True
End Sub

Private Sub chkPreview_Click()

  Changed = True

End Sub

Private Sub chkProp_Group_Click()
  ' Store the value in the collection for the current column
  If Not mblnLoading Then
    Changed = True
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.GroupWithNext = IIf(chkProp_Group.Value = vbChecked, True, False)
    Set objItem = Nothing
    
    EnableColProperties
  End If
End Sub

Private Sub cmdAdd_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAdd.Picture = cmdAdd.Picture

End Sub

Private Sub cmdAddAll_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAddAll.Picture = cmdAddAll.Picture

End Sub

Private Sub cmdClearOrder_Click()
  If MsgBox("Are you sure you wish to clear the sort order?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    grdReportOrder.RemoveAll
    grdReportOrder.SelBookmarks.RemoveAll
    EnableDisableTabControls
    Me.Changed = True
  End If
End Sub

Private Sub cmdMoveDown_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveDown.Picture = cmdMoveDown.Picture

End Sub

Private Sub cmdMoveUp_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveUp.Picture = cmdMoveUp.Picture

End Sub

Private Sub cmdRemove_LostFocus()
  'JPD 20031013 Fault 5827
  cmdRemove.Picture = cmdRemove.Picture

End Sub

Private Sub cmdRemoveAll_LostFocus()
  'JPD 20031013 Fault 5827
  cmdRemoveAll.Picture = cmdRemoveAll.Picture

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

Private Sub grdChildren_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  With grdChildren
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .AddItemBookmark(.Row)
  End With
  
  UpdateButtonStatus (SSTab1.Tab)
End Sub







Private Sub grdRepetition_Change()
  
  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String
  Dim sMessage As String
  
  sKey = grdRepetition.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item(sKey)
  sMessage = vbNullString
  
  If Not objColumn Is Nothing Then
  
    If mblnGridActionCancelled Then
      Select Case grdRepetition.Col
      Case 2
        If objColumn.Hidden And grdRepetition.Columns("repetition").Value Then
          mblnGridChangeRecursive = True
          grdRepetition.Columns("repetition").Value = vbUnchecked
          sMessage = "You cannot select 'Repetition' for a hidden column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        ElseIf objColumn.SurpressRepeatedValues And grdRepetition.Columns("repetition").Value Then
          mblnGridChangeRecursive = True
          grdRepetition.Columns("repetition").Value = vbUnchecked
          sMessage = "You cannot select both 'Suppress Repeated Values' and 'Repetition' for the same column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        End If
        
      End Select
  
    End If
    
    objColumn.Repetition = grdRepetition.Columns("repetition").Value
  
  End If
  
  Set objColumn = Nothing
  
  If (Not mblnGridActionCancelled) And (Not mblnGridChangeRecursive) Then
    Changed = True
  End If
    
  grdRepetition_RowColChange 0, 0
  grdReportOrder_RowColChange 0, 0
  
End Sub

Private Sub grdRepetition_Click()

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String

  If Not mblnIsChildColumnSelected And (grdRepetition.Rows > 0) Then
    MsgBox "Repetition cannot be selected until a child table column or calculation has been added to the report.", vbOKOnly + vbInformation, "Custom Reports"
    Exit Sub
  End If
  
  sKey = grdRepetition.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item(sKey)
  mblnGridActionCancelled = False
  mblnGridChangeRecursive = False
  
  If Not objColumn Is Nothing Then
  
    Select Case grdRepetition.Col
    Case 1
      Exit Sub
  
    Case 2
      If objColumn.Hidden Or objColumn.SurpressRepeatedValues Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    End Select
  
  End If
  
  Set objColumn = Nothing

End Sub


Private Sub grdRepetition_DblClick()

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String

  If Not mblnIsChildColumnSelected And (grdRepetition.Rows > 0) Then
    MsgBox "Repetition cannot be selected until a child table column or calculation has been added to the report.", vbOKOnly + vbInformation, "Custom Reports"
    Exit Sub
  End If

  sKey = grdRepetition.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item(sKey)
  mblnGridActionCancelled = False
  mblnGridChangeRecursive = False
  
  If Not objColumn Is Nothing Then
  
    Select Case grdRepetition.Col
    Case 1
      Exit Sub
  
    Case 2
      If objColumn.Hidden Or objColumn.SurpressRepeatedValues Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    End Select
  
  End If
  
  Set objColumn = Nothing

End Sub


Private Sub grdRepetition_GotFocus()
  If Me.ActiveControl Is grdRepetition Then
    If grdRepetition.Col = 1 Then grdRepetition.Col = 0
  End If
End Sub




Private Sub grdRepetition_KeyUp(KeyCode As Integer, Shift As Integer)

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String

  If Not mblnIsChildColumnSelected And (grdRepetition.Rows > 0) Then
    If (KeyCode = vbKeySpace) Then
      MsgBox "Repetition cannot be selected until a child table column or calculation has been added to the report.", vbOKOnly + vbInformation, "Custom Reports"
      Exit Sub
    End If
  End If
  
  sKey = grdRepetition.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item(sKey)
  mblnGridActionCancelled = False
  mblnGridChangeRecursive = False
  
  If Not objColumn Is Nothing Then
  
    Select Case grdRepetition.Col
    Case 1
      Exit Sub
  
    Case 2
      If objColumn.Hidden Or objColumn.SurpressRepeatedValues Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    End Select
  
  End If
  
  Set objColumn = Nothing

End Sub

Private Sub grdRepetition_RowLoaded(ByVal Bookmark As Variant)

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String
  
  With grdRepetition
  
    sKey = .Columns("ColumnID").CellValue(Bookmark)
    Set objColumn = mcolCustomReportColDetails.Item(sKey)
   
    If Not objColumn Is Nothing Then
    
      If objColumn.Hidden Or objColumn.SurpressRepeatedValues Or (Not mblnIsChildColumnSelected) Then
        .Columns("repetition").CellStyleSet "ssetDisabled", .AddItemRowIndex(Bookmark)
      Else
        .Columns("repetition").CellStyleSet "ssetEnabled", .AddItemRowIndex(Bookmark)
      End If
    
    End If
    
  End With
  
  Set objColumn = Nothing

  If Me.ActiveControl Is grdRepetition Then
    If grdRepetition.Col = 1 Then
      grdRepetition.Col = 0
    End If
  End If

End Sub

Private Sub grdReportOrder_Change()

  If mblnReadOnly Then
    Exit Sub
  End If
  
  Dim objColumn  As clsColumn
  Dim sKey As String
  Dim sMessage As String
  
  sMessage = vbNullString
  sKey = grdReportOrder.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item("C" & sKey)
  
  If Not objColumn Is Nothing Then

    If mblnGridActionCancelled Then
      Select Case grdReportOrder.Col
      Case 3
        If grdReportOrder.Columns("break").Value And objColumn.PageOnChange Then
          mblnGridChangeRecursive = True
          grdReportOrder.Columns("break").Value = vbUnchecked
          sMessage = "You cannot select both 'Break on Change' and 'Page on Change' for the same column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        End If
        
      Case 4
        If objColumn.BreakOnChange And grdReportOrder.Columns("page").Value Then
          mblnGridChangeRecursive = True
          grdReportOrder.Columns("page").Value = vbUnchecked
          sMessage = "You cannot select both 'Break on Change' and 'Page on Change' for the same column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        End If
      
      Case 5
        If objColumn.Hidden And grdReportOrder.Columns("value").Value Then
          mblnGridChangeRecursive = True
          grdReportOrder.Columns("value").Value = vbUnchecked
          sMessage = "You cannot select 'Value on Change' for a hidden column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        End If
        
      Case 6
        If objColumn.Hidden And grdReportOrder.Columns("hide").Value Then
          mblnGridChangeRecursive = True
          grdReportOrder.Columns("hide").Value = vbUnchecked
          sMessage = "You cannot select 'Suppress Repeated Values' for a hidden column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        ElseIf objColumn.Repetition And grdReportOrder.Columns("hide").Value Then
          mblnGridChangeRecursive = True
          grdReportOrder.Columns("hide").Value = vbUnchecked
          sMessage = "You cannot select both 'Suppress Repeated Values' and 'Repetition' for the same column."
          MsgBox sMessage, vbOKOnly + vbInformation, "Custom Reports"
        End If
        
      End Select
      
    End If
    
    objColumn.BreakOnChange = grdReportOrder.Columns("break").Value
    objColumn.PageOnChange = grdReportOrder.Columns("page").Value
    objColumn.ValueOnChange = grdReportOrder.Columns("value").Value
    objColumn.SurpressRepeatedValues = grdReportOrder.Columns("hide").Value
    
  End If
  
  Set objColumn = Nothing

  If (Not mblnGridActionCancelled) And (Not mblnGridChangeRecursive) Then
    Changed = True
  End If
  
  grdReportOrder_RowColChange 0, 0
  grdRepetition_RowColChange 0, 0
  
End Sub

Private Sub grdReportOrder_Click()

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String

  sKey = grdReportOrder.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item("C" & sKey)
  mblnGridActionCancelled = False
  mblnGridChangeRecursive = False

  If Not objColumn Is Nothing Then

    Select Case grdReportOrder.Col
    Case 1
      Exit Sub
  
    Case 3
      If objColumn.PageOnChange Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 4
      If objColumn.BreakOnChange Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 5
      If objColumn.Hidden Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 6
      If objColumn.Hidden Or objColumn.Repetition Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    End Select
  
  End If
  
  Set objColumn = Nothing

End Sub

Private Sub grdReportOrder_DblClick()

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String

  sKey = grdReportOrder.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item("C" & sKey)
  mblnGridActionCancelled = False
  mblnGridChangeRecursive = False

  If Not objColumn Is Nothing Then

    Select Case grdReportOrder.Col
    Case 1
      Exit Sub
  
    Case 3
      If objColumn.PageOnChange Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 4
      If objColumn.BreakOnChange Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 5
      If objColumn.Hidden Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 6
      If objColumn.Hidden Or objColumn.Repetition Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    End Select
  
  End If
  
  Set objColumn = Nothing

End Sub


Private Sub grdReportOrder_GotFocus()
  If Me.ActiveControl Is grdReportOrder Then
    If grdReportOrder.Col = 1 Then grdReportOrder.Col = 0
  End If
End Sub







Private Sub grdReportOrder_KeyUp(KeyCode As Integer, Shift As Integer)
  
  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String
  
  sKey = grdReportOrder.Columns("ColumnID").Value
  Set objColumn = mcolCustomReportColDetails.Item("C" & sKey)
  mblnGridActionCancelled = False
  mblnGridChangeRecursive = False

  If Not objColumn Is Nothing Then

    Select Case grdReportOrder.Col
    Case 1
      Exit Sub
  
    Case 3
      If objColumn.PageOnChange Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 4
      If objColumn.BreakOnChange Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 5
      If objColumn.Hidden Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    Case 6
      If objColumn.Hidden Or objColumn.Repetition Then
        mblnGridActionCancelled = True
        Exit Sub
      End If
  
    End Select
  
  End If
  
  Set objColumn = Nothing

End Sub


Private Sub grdReportOrder_RowLoaded(ByVal Bookmark As Variant)

  If mblnReadOnly Then
    Exit Sub
  End If

  Dim objColumn  As clsColumn
  Dim sKey As String
  
  With grdReportOrder
  
    sKey = .Columns("ColumnID").CellValue(Bookmark)
    Set objColumn = mcolCustomReportColDetails.Item("C" & sKey)
    
    If Not objColumn Is Nothing Then
    
      If objColumn.BreakOnChange Then
        .Columns("page").CellStyleSet "ssetDisabled", .AddItemRowIndex(Bookmark)
      Else
        .Columns("page").CellStyleSet "ssetEnabled", .AddItemRowIndex(Bookmark)
      End If
      
      If objColumn.PageOnChange Then
        .Columns("break").CellStyleSet "ssetDisabled", .AddItemRowIndex(Bookmark)
      Else
        .Columns("break").CellStyleSet "ssetEnabled", .AddItemRowIndex(Bookmark)
      End If
    
      If objColumn.Hidden Then
        .Columns("value").CellStyleSet "ssetDisabled", .AddItemRowIndex(Bookmark)
      Else
        .Columns("value").CellStyleSet "ssetEnabled", .AddItemRowIndex(Bookmark)
      End If
      
      If objColumn.Hidden Or objColumn.Repetition Then
        .Columns("hide").CellStyleSet "ssetDisabled", .AddItemRowIndex(Bookmark)
      Else
        .Columns("hide").CellStyleSet "ssetEnabled", .AddItemRowIndex(Bookmark)
      End If
      
    End If
      
  End With
  
  Set objColumn = Nothing

  If Me.ActiveControl Is grdReportOrder Then
    If grdReportOrder.Col = 1 Then
      grdReportOrder.Col = 0
    End If
  End If

End Sub

Private Sub optOutputFormat_Click(Index As Integer)

  mobjOutputDef.FormatClick Index

  Changed = True

End Sub

Private Sub chkDestination_Click(Index As Integer)

  mobjOutputDef.DestinationClick Index

  Changed = True

End Sub

Private Sub chkPrintFilterHeader_Click()

  Changed = True
  
End Sub

Private Sub chkProp_Average_Click()

  ' Store the value in the collection for the current column
  If Not mblnLoading Then
    Changed = True
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Average = IIf(chkProp_Average.Value = vbChecked, True, False)
    Set objItem = Nothing
    EnableColProperties
  End If
  
End Sub

Private Sub chkProp_Count_Click()
  
  ' Store the value in the collection for the current column
  If Not mblnLoading Then
    Changed = True
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Count = IIf(chkProp_Count.Value = vbChecked, True, False)
    Set objItem = Nothing
    EnableColProperties
  End If
  
End Sub

Private Sub chkProp_Hidden_Click()
  ' Store the value in the collection for the current column
  If Not mblnLoading Then
    Changed = True
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Hidden = IIf(chkProp_Hidden.Value = vbChecked, True, False)
    Set objItem = Nothing
    
    EnableColProperties
  End If
End Sub

Private Sub chkProp_Total_Click()

  ' Store the value in the collection for the current column
  If Not mblnLoading Then
    Changed = True
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Total = IIf(chkProp_Total.Value = vbChecked, True, False)
    Set objItem = Nothing
    EnableColProperties
  End If
  
End Sub

Private Sub chkSummaryReport_Click()

  Changed = True
  
End Sub


Private Sub cmdAddChild_Click()
  
  Dim plngRow As Long
  Dim pstrRow As String
  Dim pfrmChild As New frmCustomReportChilds
  
  If Me.grdChildren.Rows < 5 Then
    With pfrmChild
      .Initialize True _
                  , Me _
                  , _
                  , _
                  , _
                  , _
                  , 0
      
      If Not .Cancelled Then .Show vbModal
      
      If Not .Cancelled Then
        pstrRow = .ChildTableID _
                  & vbTab & .ChildTable _
                  & vbTab & .Order _
                  & vbTab & .OrderID _
                  & vbTab & .FilterID _
                  & vbTab & .Filter _
                  & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)
        
        With Me.grdChildren
          .AddItem pstrRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End With
        Changed = True
      End If
      
    End With

    Unload pfrmChild
    Set pfrmChild = Nothing
  
    PopulateTableAvailable , False
  
    EnableDisableTabControls
  
    ForceDefinitionToBeHiddenIfNeeded
  
    UpdateButtonStatus (Me.SSTab1.Tab)
    
    Changed = True

  Else
    MsgBox "The maximum of five child tables has been selected.", vbInformation + vbOKOnly, "Custom Reports"
  End If

End Sub

Private Sub cmdAddOrder_Click()

  If frmCustomReportsAddOrder.Initialise(mstrBaseTable, Me.ListView2.ListItems.Count, , Me) = True Then
    'Me.grdReportOrder.MoveLast
    'Me.grdReportOrder.SelBookmarks.Add grdReportOrder.Bookmark
    UpdateOrderButtons
    
    'AE20071025 Fault #6797
    If Not frmCustomReportsAddOrder.UserCancelled Then
      Changed = True
    End If
  End If
  
  Unload frmCustomReportsAddOrder
  Set frmCustomReportsAddOrder = Nothing
  
  Me.grdReportOrder.Redraw = True
  grdRepetition_RowColChange 0, 0

End Sub

Private Sub UpdateOrderButtons()

  If mblnReadOnly Then
    Exit Sub
  End If
  
  With grdReportOrder
    
    ''NHRD03102005 Fault 10364 Commented out this piece of old code after discussions with Tim
'    'TM20020809 Fault 4244 - if only one row exists then make sure that row is selected.
'    If (.Rows = 1) And (.SelBookmarks.Count < 1) Then
'      .MoveFirst
'      .SelBookmarks.RemoveAll
'      .SelBookmarks.Add .Bookmark
'    End If

    If .Rows = 0 Then
      cmdEditOrder.Enabled = False
      cmdDeleteOrder.Enabled = False
      cmdClearOrder.Enabled = False
    Else
      cmdEditOrder.Enabled = True
      cmdDeleteOrder.Enabled = True
      cmdClearOrder.Enabled = True
    End If

    If .AddItemRowIndex(.Bookmark) = 0 Then
      Me.cmdMoveUpOrder.Enabled = False
      Me.cmdMoveDownOrder.Enabled = .Rows > 1
    ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
      Me.cmdMoveUpOrder.Enabled = .Rows > 1
      Me.cmdMoveDownOrder.Enabled = False
    Else
      'TM20020809 Fault 4244 - only enable the move buttons if more than one row exists.
      Me.cmdMoveUpOrder.Enabled = .Rows > 1
      Me.cmdMoveDownOrder.Enabled = .Rows > 1
    End If
  
  End With

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdDeleteOrder_Click()

  Dim lRow As Long
  Dim objColumn As clsColumn
  
  If Me.grdReportOrder.Rows = 1 Then
    Me.grdReportOrder.RemoveAll
       
    For Each objColumn In mcolCustomReportColDetails.Collection
      objColumn.BreakOnChange = False
      objColumn.PageOnChange = False
      objColumn.ValueOnChange = False
      objColumn.SurpressRepeatedValues = False
    Next objColumn
    
  Else
    With grdReportOrder
      Set objColumn = mcolCustomReportColDetails.Item("C" & grdReportOrder.Columns("ColumnID").Value)
      objColumn.BreakOnChange = False
      objColumn.PageOnChange = False
      objColumn.ValueOnChange = False
      objColumn.SurpressRepeatedValues = False
      
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

  grdRepetition_RowColChange 0, 0

  Changed = True

  Set objColumn = Nothing
  
End Sub


Private Sub cmdEditChild_Click()

  Dim pstrRow As String
  Dim plngRow As Long
  Dim pfrmChild As New frmCustomReportChilds
  Dim lngInitTableID As Long
  Dim i2 As Integer
  Dim bNeedRefreshAvail As Boolean
  
  With Me.grdChildren
    plngRow = .AddItemRowIndex(.Bookmark)
    lngInitTableID = .Columns("TableID").Value
    pfrmChild.Initialize False _
                , Me _
                , .Columns("TableID").Value _
                , .Columns("Table").Value _
                , .Columns("FilterID").Value _
                , .Columns("Filter").Value _
                , IIf(.Columns("Records").Value = sALL_RECORDS, 0, .Columns("Records").Value) _
                , .Columns("OrderID").Value _
                , .Columns("Order").Value
                
  End With
  
  With pfrmChild
    .Show vbModal
    
    If Not .Cancelled Then
      If .ChildTableID <> lngInitTableID Then
        ' Check if any columns in the report definition are from the table that was
        ' previously selected in the child combo box. If so, prompt user for action.
        Select Case AnyChildColumnsUsed(lngInitTableID)
        Case 2: ' child cols used and user wants to continue with the change
          'TM20020424 Fault 3803
          If ListView2.ListItems.Count > 0 Then
            SelectLast ListView2
            GetCurrentDetails ListView2.SelectedItem.Key
          End If
          pstrRow = .ChildTableID _
                    & vbTab & .ChildTable _
                    & vbTab & .Order _
                    & vbTab & .OrderID _
                    & vbTab & .FilterID _
                    & vbTab & .Filter _
                    & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)
  
        Case 1: ' child cols used and user has aborted the change
          With Me.grdChildren
            .Bookmark = .AddItemBookmark(plngRow)
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(plngRow)
          End With

          Exit Sub
          
        Case 0: ' no child cols used
          pstrRow = .ChildTableID _
                    & vbTab & .ChildTable _
                    & vbTab & .Order _
                    & vbTab & .OrderID _
                    & vbTab & .FilterID _
                    & vbTab & .Filter _
                    & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)

        End Select
      Else
        pstrRow = .ChildTableID _
                  & vbTab & .ChildTable _
                  & vbTab & .Order _
                  & vbTab & .OrderID _
                  & vbTab & .FilterID _
                  & vbTab & .Filter _
                  & vbTab & IIf(.MaxRecords = 0, sALL_RECORDS, .MaxRecords)
      End If

      
      With grdChildren
        'TM20020424 Fault 3715
        'Find and remove from Table Available
        For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
          If cboTblAvailable.ItemData(i2) = lngInitTableID Then
            
            If cboTblAvailable.ListIndex = i2 Then
              bNeedRefreshAvail = True
            End If

            cboTblAvailable.RemoveItem i2
            Exit For
          End If
        Next i2
        .RemoveItem plngRow
        .AddItem pstrRow, plngRow
      End With
    Else
      With Me.grdChildren
        .Bookmark = .AddItemBookmark(plngRow)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .AddItemBookmark(plngRow)
      End With

      Exit Sub
    End If
  End With

  Unload pfrmChild
  Set pfrmChild = Nothing

  PopulateTableAvailable , bNeedRefreshAvail

  EnableDisableTabControls

  ForceDefinitionToBeHiddenIfNeeded
  
  With Me.grdChildren
    .Bookmark = .AddItemBookmark(plngRow)
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .AddItemBookmark(plngRow)
  End With
  
  UpdateButtonStatus (Me.SSTab1.Tab)
  
  Changed = True

End Sub

Private Sub cmdEditOrder_Click()
  
  frmCustomReportsAddOrder.Initialise mstrBaseTable, Me.ListView2.ListItems.Count, True, Me
  UpdateOrderButtons
  
  'AE20071025 Fault #6797
  If Not frmCustomReportsAddOrder.UserCancelled Then
    Changed = True
  End If
  
  Unload frmCustomReportsAddOrder
  Set frmCustomReportsAddOrder = Nothing

  Me.grdReportOrder.Redraw = True
  grdRepetition_RowColChange 0, 0

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
  
  UpdateButtonStatus (Me.SSTab1.Tab)

  Changed = True
  
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

  UpdateButtonStatus (Me.SSTab1.Tab)
  
  Changed = True
  
End Sub

Private Sub cmdNewCalculation_Click()

  Dim objExpr As New clsExprExpression
  Dim strKey As String
  Dim strKeyPrevSelected As String
  Dim SelectedKeys() As String
  Dim iCount As Integer
  Dim sMessage As String
  Dim objTempItem As ListItem
  
  On Error GoTo NewCalc_ERROR
  
  ReDim SelectedKeys(0)
  Dim lst As ListItem
  For Each lst In ListView1.ListItems
    If lst.Selected = True Then
      ReDim Preserve SelectedKeys(UBound(SelectedKeys) + 1)
      SelectedKeys(UBound(SelectedKeys) - 1) = lst.Key
    End If
  Next lst
  
  With objExpr
    If .Initialise(Me.cboTblAvailable.ItemData(Me.cboTblAvailable.ListIndex), 0, giEXPR_RUNTIMECALCULATION, 0) Then
      .SelectExpression True
    End If
    
    ' Refresh the listview to show the newly added calculation
    PopulateAvailable
    
    ' Refresh the names of selected calcs
    RefreshSelectedCalcNames
    
    UpdateButtonStatus (Me.SSTab1.Tab)

    If .ExpressionID > 0 Then
    
      '02/08/2000 MH Fault 2386
      If mblnReadOnly Then
        MsgBox "Unable to select calculation as you are viewing a read only definition", vbExclamation, "Custom Reports"
      
      Else
        strKey = "E" & CStr(.ExpressionID)
        If Not AlreadyUsed(strKey) Then
          
          'De-select all the available items.
          For Each lst In ListView1.ListItems
            lst.Selected = False
          Next lst
          
          'Check for hidden elements within the calc.
          sMessage = IsCalcValid(.ExpressionID)
          If sMessage <> vbNullString Then
            MsgBox "This calculation has been deleted or hidden by another user." & vbCrLf & _
                   "It cannot be added to this definition", vbExclamation, App.Title
          Else
            If optCalc.Value And (cboTblAvailable.ItemData(cboTblAvailable.ListIndex) = .BaseTableID) Then
              ListView1.ListItems(strKey).Selected = True
              Call CopyToSelected(False)
            Else
              Set objTempItem = ListView2.ListItems.Add(, strKey, .Name, , ImageList1.ListImages("IMG_CALC").Index)
              AddToCollection2 objTempItem
            End If
            
            If .BaseTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) _
              Or (.BaseTableID = Me.txtParent1.Tag) Or (.BaseTableID = Me.txtParent2.Tag) Then
              grdRepetition.AddItem "E" & .ExpressionID & vbTab & "<" & .BaseTableName & " Calc>" & .Name & vbTab & 0
            End If

          End If
        
          ' RH 09/04/01 - leaves 2 things highlighted
          For Each lst In ListView2.ListItems
            If lst.Selected = True Then
              lst.Selected = False
            End If
          Next lst
          ListView2.ListItems(strKey).Selected = True
          
          'De-select all the available items.
          For Each lst In ListView1.ListItems
            lst.Selected = False
          Next lst
        End If

      End If
    End If

  End With

  Set objExpr = Nothing

  ' Reselect the cols/calcs that were selected before the calc button was pressed
  For iCount = 0 To (UBound(SelectedKeys) - 1)
    For Each lst In ListView1.ListItems
      If lst.Key = SelectedKeys(iCount) Then
        lst.Selected = True
        Exit For
      End If
    Next lst
  Next iCount
  
  'JPD 20030728 Fault 6460
  UpdateButtonStatus (Me.SSTab1.Tab)
  ForceDefinitionToBeHiddenIfNeeded
  
  ListView1.SetFocus
  
  Exit Sub
  
NewCalc_ERROR:
  
  Select Case Err.Number
  
    Case 35601:  ' Expression could not be selected because the copy was aborted - hidden calc
                 ' selected, but user not the definition owner.
    Case Else: MsgBox "Error : " & Err.Description, vbExclamation + vbOKOnly, App.Title
  
  End Select
  
  Resume Next
  
  
End Sub

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


Private Sub cmdOK_Click()

  If Changed = True Then
    'NHRD24042002 Fault 3728 Switch on and off the hourglass when saving definitions
    Screen.MousePointer = vbHourglass
    
    'TM20020508 Fault 3839 - need to set the mouse pointer to default before exit sub is called.
    If Not ValidateDefinition(mlngCustomReportID) Then
      Screen.MousePointer = vbNormal
      Exit Sub
    End If
    
    If Not SaveDefinition Then
      Screen.MousePointer = vbNormal
      Exit Sub
    End If
    
    Screen.MousePointer = vbNormal
  End If
  
  Me.Hide
  
End Sub


Private Sub cmdParent1Filter_Click()
  
  GetFilter txtParent1, txtParent1Filter
  'Changed = True
  
End Sub

Private Sub cmdParent1Picklist_Click()
  GetPicklist txtParent1, txtParent1Picklist

End Sub

Private Sub cmdParent2Filter_Click()
  
  GetFilter txtParent2, txtParent2Filter
  'Changed = True

End Sub

Private Sub cmdParent2Picklist_Click()
  GetPicklist txtParent2, txtParent2Picklist

End Sub


Private Sub cmdRemoveAllChilds_Click()
  
  Dim i As Integer
  Dim i2 As Integer
  Dim pvarbookmark As Variant
  Dim bContinueRemoval As Boolean
  Dim lngSelectedChild As Long
  Dim lngRowCount As Long
  Dim bRemovedFromAvailable As Boolean
  Dim lRow As Long
  Dim bNeedRefreshAvail As Boolean
  
  If IsChildColumnSelected Then
    bContinueRemoval = (MsgBox("Removing all the child tables will remove all child table columns " & _
                              "included in the report definition. " & vbCrLf & _
                              "Do you wish to continue ?" _
                              , vbYesNo + vbQuestion, "Custom Reports") = vbYes)
  Else
    bContinueRemoval = True
  End If
  
  If Not bContinueRemoval Then Exit Sub

  With grdChildren
    lngRowCount = .Rows
    For i = 0 To lngRowCount - 1 Step 1
      .MoveFirst
      pvarbookmark = .GetBookmark(0)
      lRow = .AddItemRowIndex(pvarbookmark)
      
      lngSelectedChild = .Columns("TableID").CellValue(pvarbookmark)
      bRemovedFromAvailable = False
      'Find and remove from Table Available
      For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
        If Not bRemovedFromAvailable Then
          If cboTblAvailable.ItemData(i2) = lngSelectedChild Then
          
            If cboTblAvailable.ListIndex = i2 Then
              bNeedRefreshAvail = True
            End If

            cboTblAvailable.RemoveItem i2
            bRemovedFromAvailable = True
          End If
        End If
      Next i2

      AnyChildColumnsUsed lngSelectedChild, bContinueRemoval
      .RemoveItem lRow
    Next i
    .RemoveAll
  End With
  
  If bNeedRefreshAvail Then
    PopulateTableAvailable , bNeedRefreshAvail
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

  UpdateButtonStatus (Me.SSTab1.Tab)
  
  Changed = True

End Sub

Private Sub cmdRemoveChild_Click()
  
  Dim i2 As Integer
  Dim lRow As Long
  Dim lngSelectedChild As Long
  Dim bNeedRefreshAvail As Boolean
  Dim varBookmark As Variant
  
  With Me.grdChildren
    lRow = .AddItemRowIndex(.Bookmark)
'    lngSelectedChild = .Columns("TableID").CellValue(lRow)
    varBookmark = .Bookmark
    lngSelectedChild = .Columns("TableID").CellValue(varBookmark)
    
    ' Check if any columns in the report definition are from the table that was
    ' previously selected in the child combo box. If so, prompt user for action.
    Select Case AnyChildColumnsUsed(lngSelectedChild)
    Case 2: ' child cols used and user wants to continue with the change
      'TM20020716 Fault 4182
      If ListView2.ListItems.Count > 0 Then
        SelectLast ListView2
        GetCurrentDetails ListView2.SelectedItem.Key
      End If
      
      'Find and remove from Table Available
      For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
        If cboTblAvailable.ItemData(i2) = lngSelectedChild Then
        
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

    Case 1: ' child cols used and user has aborted the change
      Exit Sub
      
    Case 0: ' no child cols used
      
      'Find and remove from Table Available combo box.
      For i2 = 0 To cboTblAvailable.ListCount - 1 Step 1
        If cboTblAvailable.ItemData(i2) = lngSelectedChild Then

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
      
    End Select
  End With
  
  
  If bNeedRefreshAvail Then
    PopulateTableAvailable , bNeedRefreshAvail
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

  UpdateButtonStatus (Me.SSTab1.Tab)
  
  Changed = True
  
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  ' SUB COMPLETE 28/01/00
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture

End Sub

Private Sub Form_Load()

  ' SUB COMPLETE 28/01/00
  ' Instantiate collection class
  Set mcolCustomReportColDetails = New clsColumns
  SSTab1.Tab = 0
  ReDim mvarHiddenCount(2, 0)
  
  grdAccess.RowHeight = 239
  
  Set mobjOutputDef = New clsOutputDef
  mobjOutputDef.ParentForm = Me
  mobjOutputDef.PopulateCombos True, True, True
  
'    With cboExportTo
'      .Clear
'      .AddItem "Html"
'      .AddItem "Microsoft Excel"
'      .AddItem "Microsoft Word"
'      .ListIndex = 0
'    End With
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer
    
    If Changed = True And Not FormPrint Then
      
      pintAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, "Custom Reports")
        
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


Private Sub grdChildren_DblClick()
  If Not mblnReadOnly Then
    If grdChildren.Rows > 0 Then
      cmdEditChild_Click
    Else
      cmdAddChild_Click
    End If
  End If
End Sub



Private Sub grdRepetition_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  If mblnReadOnly Then
    Exit Sub
  End If
  
  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap
  
  Dim objColumn  As clsColumn
  Dim sKey As String
  Dim iRow As Integer
  
  With grdRepetition
  
    For iRow = 0 To .Rows - 1
      
      sKey = .Columns("ColumnID").CellValue(.AddItemBookmark(iRow))
      Set objColumn = mcolCustomReportColDetails.Item(sKey)
      
      If Not objColumn Is Nothing Then
      
        If objColumn.Hidden Or objColumn.SurpressRepeatedValues Or (Not mblnIsChildColumnSelected) Then
          .Columns("repetition").CellStyleSet "ssetDisabled", iRow
        Else
          .Columns("repetition").CellStyleSet "ssetEnabled", iRow
        End If
      
      End If
      
    Next iRow
    
    .Refresh
  End With
  
  Set objColumn = Nothing

  If Me.ActiveControl Is grdRepetition Then
    If grdRepetition.Col = 1 Then
      grdRepetition.Col = 0
    End If
  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub



Private Sub grdReportOrder_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  If mblnReadOnly Then
    Exit Sub
  End If

  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap
  
  Dim objColumn  As clsColumn
  Dim sKey As String
  Dim iRow As Integer
  
  With grdReportOrder
  
    For iRow = 0 To .Rows - 1
      
      sKey = .Columns("ColumnID").CellValue(.AddItemBookmark(iRow))
      Set objColumn = mcolCustomReportColDetails.Item("C" & sKey)
     
      If Not objColumn Is Nothing Then
       
        If objColumn.BreakOnChange Then
          .Columns("page").CellStyleSet "ssetDisabled", iRow
        Else
          .Columns("page").CellStyleSet "ssetEnabled", iRow
        End If
      
        If objColumn.PageOnChange Then
          .Columns("break").CellStyleSet "ssetDisabled", iRow
        Else
          .Columns("break").CellStyleSet "ssetEnabled", iRow
        End If
  
        If objColumn.Hidden Then
          .Columns("value").CellStyleSet "ssetDisabled", iRow
        Else
          .Columns("value").CellStyleSet "ssetEnabled", iRow
        End If
  
        If objColumn.Hidden Or objColumn.Repetition Then
          .Columns("hide").CellStyleSet "ssetDisabled", iRow
        Else
          .Columns("hide").CellStyleSet "ssetEnabled", iRow
        End If
        
      End If
      
    Next iRow
    
    .Refresh
  End With
  
  Set objColumn = Nothing

  If Me.ActiveControl Is grdReportOrder Then
    If grdReportOrder.Col = 1 Then
      grdReportOrder.Col = 0
    End If
  End If

  UpdateOrderButtons
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub
Private Sub ListView1_GotFocus()
  cmdAdd.Default = True
End Sub

Private Sub ListView1_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub ListView2_GotFocus()
  cmdRemove.Default = True
End Sub

Private Sub ListView2_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub optBaseAllRecords_Click()

  Changed = True

  cmdBasePicklist.Enabled = False
  With txtBasePicklist
    .Text = ""
    .Tag = 0
  End With
  
  cmdBaseFilter.Enabled = False
  With txtBaseFilter
    .Text = ""
    .Tag = 0
  End With

  chkPrintFilterHeader.Value = vbUnchecked
  
  'JPD 20040628 Fault 8818
  If Not mblnLoading Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
  EnableDisableTabControls
  
End Sub

Private Sub optBaseFilter_Click()

  Changed = True

  cmdBasePicklist.Enabled = False
  With txtBasePicklist
    .Text = ""
    .Tag = 0
  End With

  cmdBaseFilter.Enabled = True
  txtBaseFilter.Text = "<None>"
  
  ForceDefinitionToBeHiddenIfNeeded

  EnableDisableTabControls

End Sub

Private Sub optCalc_Click()
  PopulateAvailable
  UpdateButtonStatus (Me.SSTab1.Tab)
End Sub

Private Sub optColumns_Click()
  PopulateAvailable
  UpdateButtonStatus (Me.SSTab1.Tab)
End Sub

Private Sub optBasePicklist_Click()

  Changed = True

  cmdBaseFilter.Enabled = False
  With txtBaseFilter
    .Text = ""
    .Tag = 0
  End With

  cmdBasePicklist.Enabled = True
  txtBasePicklist.Text = "<None>"
  
  If Not mblnLoading Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
  EnableDisableTabControls
  
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





Private Sub SSTab1_Click(PreviousTab As Integer)
  EnableDisableTabControls
End Sub


Private Sub cmdBasePicklist_Click()

  GetPicklist cboBaseTable, txtBasePicklist

End Sub

Private Sub cmdBaseFilter_Click()
  
  GetFilter cboBaseTable, txtBaseFilter
'  Changed = True

End Sub



Private Sub ListView1_DblClick()

  If mblnReadOnly Then
    Exit Sub
  End If

  ' SUB COMPLETED 28/01/00
  ' Copy the item doubleclicked on to the 'Selected' Listview
  CopyToSelected False

End Sub


Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

  ' SUB COMPLETED 28/01/00
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

  ' SUB COMPLETED 28/01/00
  ' Capture the Ctl-A event. This does not trigger the itemclick
  ' event so have to force an updatebuttonstatus here
  
  Dim objItem As ListItem
  
  If Shift = vbCtrlMask And KeyCode = 65 Then
    For Each objItem In ListView2.ListItems
      objItem.Selected = True
    Next objItem
    Set objItem = Nothing
    UpdateButtonStatus (Me.SSTab1.Tab)
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
  
  ' SUB COMPLETED 28/01/00
  UpdateButtonStatus (Me.SSTab1.Tab)
  
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

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
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

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  ' SUB COMPLETED 28/01/00
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
  
  ' SUB COMPLETED 28/01/00
  ' Perform the drop operation
  If Source Is ListView2 Then
    CopyToAvailable False
    ListView2.Drag vbCancel
  Else
    ListView2.Drag vbCancel
  End If

End Sub

Private Sub ListView2_DragDrop(Source As Control, x As Single, y As Single)
  
  ' SUB COMPLETED 28/01/00
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
  
  ' SUB COMPLETED 28/01/00
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture
  
End Sub

Private Sub Frafieldsselected_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  ' SUB COMPLETED 28/01/00
  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture
  
End Sub

Private Sub ListView2_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  ' SUB COMPLETED 28/01/00
  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    If UCase(Left(Source.SelectedItem.Key, 1)) = "E" Then
      Source.DragIcon = picDocument(1).Picture
    Else
      Source.DragIcon = picDocument(0).Picture
    End If
  End If

  ' Set DropHighlight to the mouse's coordinates.
  Set ListView2.DropHighlight = ListView2.HitTest(x, y)

End Sub

Private Sub ListView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  ' SUB COMPLETED 28/01/00
  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    If UCase(Left(Source.SelectedItem.Key, 1)) = "E" Then
      Source.DragIcon = picDocument(1).Picture
    Else
      Source.DragIcon = picDocument(0).Picture
    End If
  End If

End Sub

Private Function CopyToSelected(bAll As Boolean, Optional intBeforeIndex As Integer)

  ' SUB COMPLETED 28/01/00
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
        
        If (lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex)) _
          Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
          grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
        End If
        '*******************************************************************

        ListView2.ListItems.Add , objTempItem.Key, GetTableNameFromColumn(Right(objTempItem.Key, Len(objTempItem.Key) - 1)) & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
        fOK = True
      Else
        lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
        
        Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT A.Access, A.TableID, B.TableName " & _
                                                          "FROM AsrSysExpressions A " & _
                                                          "     INNER JOIN ASRSysTables B " & _
                                                          "     ON A.TableID = B.TableID " & _
                                                          "WHERE A.ExprID = " & lngColumnID)
        If Not prstTemp.BOF And Not prstTemp.EOF Then
          If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
            MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                  " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Custom Reports"
            fOK = False
          Else
            fOK = True
            ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon

            lngTableID = prstTemp!TableID
            If lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) _
              Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
              grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
            End If
          End If
        End If
      End If
      If fOK Then AddToCollection2 objTempItem
    Next objTempItem
    
    ListView1.ListItems.Clear
    SelectFirst ListView2
    UpdateButtonStatus (Me.SSTab1.Tab)
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
        
        If (lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex)) _
          Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
          grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
        End If
        '*******************************************************************
        
        ListView2.ListItems.Add , objTempItem.Key, sTempTableName & "." & objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
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
          MsgBox "The selected calculation has been deleted.", vbExclamation + vbOKOnly, "Custom Reports"
          fOK = True
        ElseIf prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
          MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Custom Reports"
          fOK = False
        Else
          ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
          fOK = True
          
          lngTableID = prstTemp!TableID
          If lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) _
            Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
            grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
          End If
        End If
      
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
        
        If (lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex)) _
          Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
          grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
        End If
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
          MsgBox "The selected calculation has been deleted.", vbExclamation + vbOKOnly, "Custom Reports"
          fOK = True
        ElseIf prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
          MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Custom Reports"
          fOK = False
        Else
          ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
          fOK = True

          lngTableID = prstTemp!TableID
          If lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) _
            Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
            grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
          End If

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
    
    UpdateButtonStatus (Me.SSTab1.Tab)
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
          
          If (lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex)) _
            Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
            grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
          End If
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
            MsgBox "One or more of the selected calculation(s) have been deleted.", vbExclamation + vbOKOnly, "Custom Reports"
            fOK = False
          End If
          If prstTemp.Fields("Access") = "HD" And Not mblnDefinitionCreator Then
            MsgBox "Cannot include the '" & objTempItem.Text & "' calculation." & vbCrLf & _
                  " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Custom Reports"
            fOK = False
          Else
            ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
            fOK = True

            lngTableID = prstTemp!TableID
            If lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) _
              Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
              grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
            End If

          End If
        End If
    
      Else
      
        ' Before an existing item
          If Left(objTempItem.Key, 1) = "C" Then

            '*******************************************************************
            lngColumnID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
            lngTableID = GetTableIDFromColumn(lngColumnID)
            sTempTableName = GetTableNameFromColumn(lngColumnID)
            
            If (lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex)) _
              Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
              grdRepetition.AddItem "C" & lngColumnID & vbTab & sTempTableName & "." & objTempItem.Text & vbTab & 0
            End If
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
                     " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Custom Reports"
              fOK = False
            Else
              fOK = True
              ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon

              lngTableID = prstTemp!TableID
              If lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) _
                Or (lngTableID = Me.txtParent1.Tag) Or (lngTableID = Me.txtParent2.Tag) Then
                grdRepetition.AddItem "E" & lngColumnID & vbTab & "<" & prstTemp!TableName & " Calc>" & objTempItem.Text & vbTab & 0
              End If

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

  UpdateButtonStatus (Me.SSTab1.Tab)
  ForceDefinitionToBeHiddenIfNeeded
  Screen.MousePointer = vbNormal
  Changed = True
  
End Function

Private Function CheckInRepetitionGrid(pstrKey As String) As Boolean

  ' Loop through the sort order grid, checking if the specified column is
  ' defined in the sort order.
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  
  With grdRepetition
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("ColumnID").CellText(pvarbookmark) = pstrKey Then
        CheckInRepetitionGrid = True
        Exit Function
      End If
      pintLoop = pintLoop + 1
    Loop
  End With
  
  CheckInRepetitionGrid = False

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

  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  
  ' Remove the specified column from the sort order grid.
  With grdRepetition
    .MoveFirst
    
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("ColumnID").CellText(pvarbookmark) = pstrKey Then
        .RemoveItem pintLoop
        
        .SelBookmarks.RemoveAll
        .MoveFirst
        Exit Sub
      End If
      pintLoop = pintLoop + 1
    Loop
  End With
  
End Sub


Private Function CopyToAvailable(bAll As Boolean, Optional intBeforeIndex As Integer)

  ' SUB COMPLETED 28/01/00
  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  
  ' Dont add the to the first listview...just remove em and
  ' repopulate the available listview...much quicker
  
  Screen.MousePointer = vbHourglass
  
  For iLoop = ListView2.ListItems.Count To 1 Step -1
    If Not bAll Then
      If ListView2.ListItems(iLoop).Selected Then
        If CheckInSortOrder(Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)) = True Then
          If MsgBox("Removing the following column will also remove it from the report sort order." & vbCrLf & vbCrLf & ListView2.ListItems(iLoop).Text & vbCrLf & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, "Custom Reports") = vbYes Then
            iTempItemIndex = iLoop
            If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
              RemoveFromRepetition ListView2.ListItems(iLoop).Key
            End If
            RemoveFromSortOrder Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)
            RemoveFromCollection ListView2.ListItems(iLoop).Key
            ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
          End If
        Else
          iTempItemIndex = iLoop
          If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
            RemoveFromRepetition ListView2.ListItems(iLoop).Key
          End If
          RemoveFromCollection ListView2.ListItems(iLoop).Key
          ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
        End If
      End If
    Else
      If CheckInSortOrder(Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)) = True Then
        If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
          RemoveFromRepetition ListView2.ListItems(iLoop).Key
        End If
        RemoveFromSortOrder Right(ListView2.ListItems(iLoop).Key, Len(ListView2.ListItems(iLoop).Key) - 1)
        RemoveFromCollection ListView2.ListItems(iLoop).Key
        ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
      Else
        If CheckInRepetitionGrid(ListView2.ListItems(iLoop).Key) Then
          RemoveFromRepetition ListView2.ListItems(iLoop).Key
        End If
        RemoveFromCollection ListView2.ListItems(iLoop).Key
        ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
      End If
      
    End If
  Next iLoop
  
  If ListView2.ListItems.Count > 0 Then
    If iTempItemIndex > ListView2.ListItems.Count Then iTempItemIndex = ListView2.ListItems.Count
    If iTempItemIndex > 0 Then ListView2.ListItems(iTempItemIndex).Selected = True
  End If
  
  PopulateAvailable
  
  UpdateButtonStatus (Me.SSTab1.Tab)
  UpdateOrderButtons
  ForceDefinitionToBeHiddenIfNeeded

  Changed = True

  Screen.MousePointer = vbNormal

End Function


Private Function ChangeSelectedOrder(Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)

  ' SUB COMPLETED 28/01/00
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
  
  UpdateButtonStatus (Me.SSTab1.Tab)

End Function


Private Function UpdateButtonStatus(iTab As Integer)

  ' SUB COMPLETED 28/01/00
  On Error Resume Next
  
  Dim tempItem As ListItem, iCount As Integer
  
  Select Case iTab
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
      EnableColProperties
    Else
      cmdRemove.Enabled = Not mblnReadOnly
      cmdRemoveAll.Enabled = Not mblnReadOnly
    
    ' If there are more than 1 items selected then disable the move buttons and exit
    For Each tempItem In ListView2.ListItems
      If tempItem.Selected Then iCount = iCount + 1
    Next tempItem
    
    If iCount > 1 Or iCount < 1 Then
      cmdMoveUp.Enabled = False
      cmdMoveDown.Enabled = False
      EnableColProperties
    Else
      If ListView2.SelectedItem.Index <> 1 Then
        cmdMoveUp.Enabled = Not mblnReadOnly
      Else
        cmdMoveUp.Enabled = False
      End If
      If ListView2.SelectedItem.Index <> ListView2.ListItems.Count Then
        cmdMoveDown.Enabled = Not mblnReadOnly
      Else
        cmdMoveDown.Enabled = False
      End If
      EnableColProperties
    End If
    
   End If
  
  Case 1:
    If grdChildren.Rows = 0 Then
      cmdAddChild.Enabled = ((mblnChildsEnabled) And (Not mblnReadOnly))
      cmdEditChild.Enabled = False
      cmdRemoveChild.Enabled = False
      cmdRemoveAllChilds.Enabled = False
    Else
      cmdAddChild.Enabled = ((mblnChildsEnabled) And (Not mblnReadOnly))
      If grdChildren.SelBookmarks.Count > 0 Then
        cmdEditChild.Enabled = (Not mblnReadOnly)
        cmdRemoveChild.Enabled = (Not mblnReadOnly)
      Else
        cmdEditChild.Enabled = False
        cmdRemoveChild.Enabled = False
      End If
      cmdRemoveAllChilds.Enabled = (Not mblnReadOnly)
    End If
    
  End Select

  'TM20020508 Fault 3790
  Call CheckListViewColWidth(ListView1)
  Call CheckListViewColWidth(ListView2)

End Function

Private Function SelectLast(lvwCtl As ListView)

  ' SUB COMPLETED 28/01/00
  Dim objItem As ListItem
  
  For Each objItem In lvwCtl.ListItems
    objItem.Selected = IIf(objItem.Index = lvwCtl.ListItems.Count, True, False)
  Next objItem

End Function

Private Function SelectFirst(lvwCtl As ListView)

  ' SUB COMPLETED 28/01/00
  Dim objItem As ListItem
  
  For Each objItem In lvwCtl.ListItems
    objItem.Selected = IIf(objItem.Index = 1, True, False)
  Next objItem

End Function

'###########

Private Sub cmdAdd_Click()
  
  ' SUB COMPLETED 28/01/00
  ' Add the selected items to the 'Selected' Listview
  CopyToSelected False

End Sub

Private Sub cmdMoveDown_Click()

  ' SUB COMPLETED 28/01/00
  ChangeSelectedOrder ListView2.SelectedItem.Index + 2, True

End Sub

Private Sub cmdMoveUp_Click()

  ' SUB COMPLETED 28/01/00
  ChangeSelectedOrder ListView2.SelectedItem.Index - 1

End Sub

Private Sub cmdRemove_Click()

  ' Remove the selected items from the 'Selected' Listview
  CopyToAvailable False
  
End Sub

Private Sub cmdAddAll_Click()

  ' SUB COMPLETED 28/01/00
  ' Add All items from to the 'Selected' Listview
  CopyToSelected True
 
End Sub

Private Sub cmdRemoveAll_Click()

  ' Remove All items from the 'Selected' Listview
  If Me.grdReportOrder.Rows > 0 Then
    If MsgBox("Removing all selected report columns will also clear the report sort order." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, "Custom Reports") = vbYes Then
      CopyToAvailable True
      EnableColProperties
    End If
  Else
    If MsgBox("Are you sure you wish to remove all columns / calculations from this definition ?", vbYesNo + vbQuestion, "Custom Reports") = vbYes Then
      CopyToAvailable True
      EnableColProperties
    End If
  End If
  
End Sub

Private Function EnableColProperties()

  Dim blnAverageChecked As Boolean
  Dim blnCountChecked As Boolean
  Dim blnTotalChecked As Boolean
  Dim blnHiddenChecked As Boolean
  Dim blnGroupChecked As Boolean
  Dim blnIsNumeric As Boolean
  Dim iCount As Integer
  Dim blnStatus As Boolean
  Dim blnResetAll As Boolean
  Dim objItem As ListItem
  Dim blnLastColumn As Boolean
  Dim blnRep As Boolean
  Dim blnSRV As Boolean
  Dim blnVOC As Boolean
  Dim iPrevIndex As Integer
  Dim blnPrevGroupChecked As Boolean
  
  mblnLoading = True
  
  If Not ListView2.SelectedItem Is Nothing Then
    If ListView2.ListItems.Count > 0 Then
      GetCurrentDetails ListView2.SelectedItem.Key
    End If
    
    'If there are more than 1 items selected then disable the properties
    For Each objItem In ListView2.ListItems
      If objItem.Selected Then
        iCount = iCount + 1
      End If
    Next objItem

  End If

  Dim objTemp As clsColumn
  Dim objTemp2 As clsColumn
  
  If ListView2.ListItems.Count < 1 Then
    blnResetAll = True
    blnStatus = False
    
  Else
    blnResetAll = False
    If (iCount > 1) Or (iCount < 1) Then
      blnStatus = False
      blnResetAll = True
      blnLastColumn = True
    Else
      blnStatus = True
      blnLastColumn = (ListView2.SelectedItem.Index = ListView2.ListItems.Count)
      
      If (ListView2.SelectedItem.Index > 1) Then
        'Get the previous column's 'Group with Next' property value.
        iPrevIndex = (ListView2.SelectedItem.Index - 1)
        Set objTemp2 = mcolCustomReportColDetails.Item(ListView2.ListItems(iPrevIndex).Key)
        blnPrevGroupChecked = objTemp2.GroupWithNext
        Set objTemp2 = Nothing
      Else
        blnPrevGroupChecked = False
      End If
      
      Set objTemp = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
      With objTemp
        If blnPrevGroupChecked Then
          .Average = False
          .Count = False
          .Total = False
          .Hidden = False
        End If
        blnAverageChecked = .Average
        blnCountChecked = .Count
        blnTotalChecked = .Total
        blnHiddenChecked = .Hidden
        blnGroupChecked = (.GroupWithNext And Not blnLastColumn)
        chkProp_Group.Value = IIf(blnGroupChecked, vbChecked, vbUnchecked)
        objTemp.GroupWithNext = blnGroupChecked
        blnIsNumeric = .IsNumeric
        blnRep = .Repetition
        blnSRV = .SurpressRepeatedValues
        blnVOC = .ValueOnChange
      End With
      
    End If
  End If
  
  blnStatus = (blnStatus And (Not mblnReadOnly))
  
  If (blnResetAll And (Not mblnReadOnly)) Then
    'need to reset all the col property controls if there are no records.
    Me.txtProp_ColumnHeading.Text = ""
    Me.spnDec.Text = ""
    Me.spnSize.Text = ""
    Me.chkProp_Average.Value = vbUnchecked
    Me.chkProp_Count.Value = vbUnchecked
    Me.chkProp_IsNumeric.Value = vbUnchecked
    Me.chkProp_Total.Value = vbUnchecked
    Me.chkProp_Group.Value = vbUnchecked
    Me.chkProp_Hidden.Value = vbUnchecked
    
  ElseIf (blnPrevGroupChecked) Then
    Me.chkProp_Average.Value = vbUnchecked
    Me.chkProp_Count.Value = vbUnchecked
    Me.chkProp_Total.Value = vbUnchecked
    Me.chkProp_Hidden.Value = vbUnchecked
    
  End If

  lblProp_ColumnHeading.Enabled = blnStatus And (Not blnHiddenChecked)
  txtProp_ColumnHeading.Enabled = lblProp_ColumnHeading.Enabled
  txtProp_ColumnHeading.BackColor = IIf(txtProp_ColumnHeading.Enabled, vbWindowBackground, vbButtonFace)
  
  lblProp_Size.Enabled = blnStatus And (Not blnHiddenChecked)
  spnSize.Enabled = lblProp_Size.Enabled
  spnSize.BackColor = IIf(spnSize.Enabled, vbWindowBackground, vbButtonFace)
  
  lblProp_Decimals.Enabled = (blnStatus And blnIsNumeric) And (Not blnHiddenChecked)
  spnDec.Enabled = lblProp_Decimals.Enabled
  spnDec.BackColor = IIf(spnDec.Enabled, vbWindowBackground, vbButtonFace)
  
  chkProp_Average.Enabled = (blnStatus And blnIsNumeric And (Not blnHiddenChecked) And (Not blnGroupChecked) And (Not blnPrevGroupChecked))
  chkProp_Count.Enabled = (blnStatus And (Not blnHiddenChecked) And (Not blnGroupChecked) And (Not blnPrevGroupChecked))
  chkProp_Total.Enabled = (blnStatus And blnIsNumeric And (Not blnHiddenChecked) And (Not blnGroupChecked) And (Not blnPrevGroupChecked))
  
  chkProp_Hidden.Enabled = (blnStatus And (Not blnAverageChecked) And (Not blnCountChecked) And (Not blnTotalChecked) And (Not blnGroupChecked) And (Not blnRep) And (Not blnSRV) And (Not blnVOC) And (Not blnPrevGroupChecked))
  chkProp_Group.Enabled = (blnStatus And (Not blnAverageChecked) And (Not blnCountChecked) And (Not blnTotalChecked) And (Not blnHiddenChecked) And (Not blnLastColumn))
  
  RefreshReportOrderGrid
  RefreshRepetitionGrid
  
  mblnLoading = False

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
  Dim bHidden As Boolean
  Dim bGroupWithNext As Boolean
  Dim bIsNumeric As Boolean
  
  mblnLoading = True
  GetDefaultDetails objTempItem.Key
  mblnLoading = False
  
  sColType = Left(objTempItem.Key, 1)
  lID = Right(objTempItem.Key, Len(objTempItem.Key) - 1)
  sHeading = Me.txtProp_ColumnHeading.Text
  lSize = Me.spnSize.Text
  iDecPlaces = IIf(Me.spnDec.Text = "", "0", Me.spnDec.Text)
  
  ' 190600 - Fault fix 358
  
  mblnLoading = True
  Me.chkProp_Average.Value = 0
  Me.chkProp_Count.Value = 0
  Me.chkProp_Total.Value = 0
  Me.chkProp_Hidden.Value = 0
  Me.chkProp_Group = 0
  mblnLoading = False
  
  bAverage = IIf(Me.chkProp_Average.Value = vbChecked, True, False)
  bCount = IIf(Me.chkProp_Count.Value = vbChecked, True, False)
  bTotal = IIf(Me.chkProp_Total.Value = vbChecked, True, False)
  bHidden = IIf(Me.chkProp_Hidden.Value = vbChecked, True, False)
  bGroupWithNext = IIf(Me.chkProp_Group.Value = vbChecked, True, False)
  
  bIsNumeric = IIf(Me.chkProp_IsNumeric.Value = vbChecked, True, False)
  
  mcolCustomReportColDetails.Add sColType, lID, sHeading, lSize, iDecPlaces, bAverage, bCount, bTotal, bIsNumeric, bHidden, bGroupWithNext

  If Not mblnLoading Then Changed = True
  
End Function

Private Function RemoveFromCollection(sKey As String) As Boolean

  ' FUNCTION COMPLETED 28/01/00
  mcolCustomReportColDetails.Remove sKey
  
End Function


Private Function GetDefaultDetails(sKey As String) As Boolean

  ' FUNCTION COMPLETED 28/01/00
  ' This function returns the default Col/Expr Name, Size and
  ' Decimal Places. These can then be edited by the user if desired.
  
  Dim rsTemp As Recordset
  
  If Left(sKey, 1) = "C" Then
    
    Set rsTemp = datGeneral.GetColumnDefinition(Right(sKey, Len(sKey) - 1))
    
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      Me.txtProp_ColumnHeading.Text = rsTemp!ColumnName
      Me.spnSize.Text = rsTemp!DefaultDisplayWidth
      
      If (rsTemp!DataType = sqlNumeric) Then ' its numeric
        Me.spnDec.Text = rsTemp!Decimals
        Me.chkProp_IsNumeric.Value = vbChecked
      ElseIf (rsTemp!DataType = sqlInteger) Then
        Me.spnSize.Text = rsTemp!DefaultDisplayWidth ' 10 '5
        Me.spnDec.Text = rsTemp!Decimals
        Me.chkProp_IsNumeric.Value = vbChecked
      ElseIf rsTemp!DataType = sqlDate Then ' its a date
        Me.spnSize.Text = rsTemp!DefaultDisplayWidth '10
        Me.spnDec.Text = 0
        Me.chkProp_IsNumeric.Value = vbUnchecked
      ElseIf rsTemp!DataType = sqlBoolean Then ' its a logic
        Me.spnSize.Text = 1
        Me.chkProp_IsNumeric.Value = vbUnchecked
      ElseIf rsTemp!DataType = sqlLongVarChar Then      ' working pattern field
        Me.spnSize.Text = 14
        Me.chkProp_IsNumeric.Value = vbUnchecked
      Else                                               ' its not
        Me.spnDec.Text = 0
        Me.chkProp_IsNumeric.Value = vbUnchecked
      End If
    End If
    
    If Me.spnSize.Text = 0 Then
      If Len(rsTemp!SpinnerMaximum) > 0 Then Me.spnSize.Text = Len(rsTemp!SpinnerMaximum)
    End If
    
  Else
  
    Set rsTemp = datGeneral.GetRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & (Right(sKey, Len(sKey) - 1)))
    
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      Me.txtProp_ColumnHeading.Text = rsTemp!Name
      Me.spnSize.Text = rsTemp!ReturnSize
      
      Dim objExpression As clsExprExpression
      Set objExpression = New clsExprExpression
      objExpression.ExpressionID = (Right(sKey, Len(sKey) - 1))
      objExpression.ConstructExpression
      
      'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
      ' when all we need is the expression return type (time saving measure).
      objExpression.ValidateExpression True, True
      
      If objExpression.ReturnType = 2 Then ' its numeric
        Me.spnDec.Text = rsTemp!ReturnDecimals
        Me.chkProp_IsNumeric.Value = vbChecked
      Else                                               ' its not
        Me.spnDec.Text = 0
        Me.chkProp_IsNumeric.Value = vbUnchecked
      End If
    End If

    Set objExpression = Nothing
    
  End If
  
End Function


Private Function GetCurrentDetails(sKey As String) As Boolean

  ' FUNCTION COMPLETED 28/01/00
  ' This function returns the details held in the collection
  ' for the currently highlighted item in the 'selected'
  ' listview
  
  Dim objTemp As clsColumn
  
  Set objTemp = mcolCustomReportColDetails.Item(sKey)
    
  If objTemp Is Nothing Then
    txtProp_ColumnHeading = ""
    spnSize.Text = 0
    spnDec.Text = 0
    chkProp_Average.Value = vbUnchecked
    chkProp_Count.Value = vbUnchecked
    chkProp_Total.Value = vbUnchecked
    chkProp_Hidden.Value = vbUnchecked
    chkProp_Group.Value = vbChecked
    EnableColProperties
    
  Else
    txtProp_ColumnHeading.Text = objTemp.Heading
    spnSize.Text = objTemp.Size
    spnDec.Text = objTemp.DecPlaces
    chkProp_Average.Value = IIf(objTemp.Average, vbChecked, vbUnchecked)
    chkProp_Count.Value = IIf(objTemp.Count, vbChecked, vbUnchecked)
    chkProp_Total.Value = IIf(objTemp.Total, vbChecked, vbUnchecked)
    chkProp_Hidden.Value = IIf(objTemp.Hidden, vbChecked, vbUnchecked)
    chkProp_Group.Value = IIf(objTemp.GroupWithNext, vbChecked, vbUnchecked)
    chkProp_IsNumeric.Value = IIf(objTemp.IsNumeric, vbChecked, vbUnchecked)

  End If
    
  Set objTemp = Nothing
    
End Function
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

Private Sub txtProp_ColumnHeading_Change()

  ' SUB COMPLETED 28/01/00
  If Not mblnLoading Then
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Heading = txtProp_ColumnHeading.Text
    Set objItem = Nothing
    Changed = True
  End If

End Sub

'Private Sub txtProp_DecPlaces_Change()
'
'  ' SUB COMPLETED 28/01/00
'  If Not mblnLoading Then
'    Dim objItem As clsColumn
'    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
'    objItem.DecPlaces = txtProp_DecPlaces.Text
'    Set objItem = Nothing
'    Changed = True
'  End If
'
'End Sub

Private Sub spnDec_Change()

  ' SUB COMPLETED 28/01/00
  
  If Me.spnDec.Text = "" Then Me.spnDec.Text = "0"
  
  If Not mblnLoading Then
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.DecPlaces = spnDec.Text
    Set objItem = Nothing
    Changed = True
  End If
  
End Sub

Private Sub spnSize_Change()

  ' SUB COMPLETED 28/01/00
  
  If Me.spnSize.Text = "" Then Me.spnSize.Text = "0"
  
  If Not mblnLoading Then
    Dim objItem As clsColumn
    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
    objItem.Size = spnSize.Text
    Set objItem = Nothing
    Changed = True
  End If

End Sub

'Private Sub txtProp_Size_Change()
'
'  ' SUB COMPLETED 28/01/00
'  If Not mblnLoading Then
'    Dim objItem As clsColumn
'    Set objItem = mcolCustomReportColDetails.Item(ListView2.SelectedItem.Key)
'    objItem.Size = txtProp_Size.Text
'    Set objItem = Nothing
'    Changed = True
'  End If
'
'End Sub



Private Function ValidateCollection() As Boolean

  ' FUNCTION COMPLETED 28/01/00
  Dim intTemp As Integer
  Dim intTemp2 As Integer
  Dim intDupCount As Integer
  Dim pstrColumnsWithSizeZero As String
  Dim intHiddenCount As Integer
  Dim intAnswer As Integer
  Dim strMessage As String
  Dim blnHasAggregate As Boolean
  Dim blnHasNumericAggregate As Boolean
  
  intHiddenCount = 0
  strMessage = vbNullString
  blnHasAggregate = False
  blnHasNumericAggregate = False
  
  ' First check the number of cols in the listview is the same as the
  ' number of items in the collection
  If ListView2.ListItems.Count <> mcolCustomReportColDetails.Count Then
    MsgBox "A serious error has occurred. To rectify, please remove all columns from the report definition and try again." & vbCrLf & "Please contact support stating : The no. of columns does not match the no of items in the collection.", vbCritical + vbOKOnly, "Custom Reports"
    SSTab1.Tab = 2
    Exit Function
  End If
  
  ' Now check that they are all unique
  For intTemp = 1 To ListView2.ListItems.Count
    intDupCount = 0
    
    If (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Heading = "") And (Not mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Hidden) Then
      MsgBox "The '" & ListView2.ListItems(intTemp).Text & "' column has a blank column heading.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 2
      Exit Function
    End If
    
    If (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Average) _
      Or (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Count) _
      Or (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Total) Then
      blnHasAggregate = True
    End If
    
    If (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).IsNumeric) _
      And ((mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Average) _
          Or (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Count) _
          Or (mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Total)) Then
      blnHasNumericAggregate = True
    End If
    
    For intTemp2 = 1 To ListView2.ListItems.Count
      If UCase(Trim(mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp2).Key).Heading)) = UCase(Trim(mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Heading)) Then
        If Not mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp2).Key).Hidden Then
          intDupCount = intDupCount + 1
        End If
      End If
    Next intTemp2
    
    If intDupCount > 1 Then
      MsgBox "One or more columns / calculations in your report have a heading of '" & mcolCustomReportColDetails.Item(Me.ListView2.ListItems(intTemp).Key).Heading & "'" & vbCrLf & "Column headings must be unique.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 2
      Exit Function
    End If
    
    If mcolCustomReportColDetails.Item(ListView2.ListItems(intTemp).Key).Hidden Then
      intHiddenCount = intHiddenCount + 1
    End If
    
  Next intTemp
  
  If intHiddenCount = ListView2.ListItems.Count Then
    strMessage = "All columns / calculations selected in this definition are defined as hidden." & vbCrLf & vbCrLf & "Do you wish to continue?"
    
    intAnswer = MsgBox(strMessage, vbExclamation + vbYesNo, Me.Caption)
    
    If intAnswer = vbNo Then
      SSTab1.Tab = 2
      Exit Function
    End If
  End If
  
  ' Check that at least one column has VOC ticked if it is a summary report.
  If chkSummaryReport.Value Then
    If Not blnHasAggregate Then
      MsgBox "You have defined this report as a summary report but have not selected to show aggregates for any of the columns.", vbExclamation + vbOKOnly, "Custom Reports"
      ValidateCollection = False
      SSTab1.Tab = 2
      Exit Function
    End If
  End If
  
  ' Check that at least one numeric column has an aggregate ticked if 'Ignore Zeros' is checked.
  If chkIgnoreZeros.Value Then
    If Not blnHasNumericAggregate Then
      MsgBox "You have chosen to ignore zeros when calculating aggregates, but have not selected to show aggregates for any numeric columns.", vbExclamation + vbOKOnly, "Custom Reports"
      ValidateCollection = False
      SSTab1.Tab = 2
      Exit Function
    End If
  End If

  'MH20010511 Allow zero size columns
  'If pstrColumnsWithSizeZero <> "" Then
  '  MsgBox "The following columns have a size of 0:" & vbCrLf & vbCrLf & pstrColumnsWithSizeZero & vbCrLf & _
  '         "Either allocate a size for these columns or remove them from the report.", vbExclamation + vbOKOnly, "Custom Reports"
  '  SSTab1.Tab = 2
  '  Exit Function
  'End If
  
  ' Got here, so the column definition is fine too #######
  ValidateCollection = True

End Function

Private Function SaveDefinition() As Boolean

  ' FUNCTION COMPLETED 28/01/00
  On Error GoTo Save_ERROR
  
  Dim sSQL As String
  Dim iLoop As Integer
  Dim sKey As String
  Dim objCol As clsColumn
  
'########################### 1 Of 3 - SAVE THE BASIC DETAILS
  
'  Dim iDefExportTo As Integer
'  Dim iDefSave As Integer
'  Dim sDefSaveAs As String
'  Dim iDefCloseApp As Integer
'
'  If optOutput(0).Value = True Then
'    iDefExportTo = 0
'    iDefSave = 0
'    sDefSaveAs = vbNullString
'    iDefCloseApp = 0
'  Else
'    Select Case UCase(cboExportTo.Text)
'      Case "HTML": iDefExportTo = 0
'      Case "MICROSOFT EXCEL": iDefExportTo = 1
'      Case "MICROSOFT WORD": iDefExportTo = 2
'    End Select
'    If chkSave.Value = vbChecked Then iDefSave = 1 Else iDefSave = 0
'    sDefSaveAs = txtExportFilename.Text
'    If chkCloseApplication.Value = vbChecked Then iDefCloseApp = 1 Else iDefCloseApp = 0
'  End If
  
  
  If mlngCustomReportID > 0 Then
    
    ' Construct the SQL Update string (Editing an existing definition)
  
    sSQL = "UPDATE ASRSYSCustomReportsName SET " & _
           "Name = '" & Trim(Replace(Me.txtName.Text, "'", "''")) & "'," & _
           "Description = '" & Replace(Me.txtDesc.Text, "'", "''") & "'," & _
           "BaseTable = " & Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) & "," & _
           "AllRecords = " & IIf(Me.optBaseAllRecords.Value, 1, 0) & "," & _
           "Picklist = " & IIf(Me.optBasePicklist.Value, Me.txtBasePicklist.Tag, 0) & "," & _
           "Filter = " & IIf(Me.optBaseFilter.Value, Me.txtBaseFilter.Tag, 0) & "," & _
           "Parent1Table = " & Me.txtParent1.Tag & "," & _
           "Parent1Filter= " & Me.txtParent1Filter.Tag & "," & _
           "Parent2Table = " & Me.txtParent2.Tag & "," & _
           "Parent2Filter= " & Me.txtParent2Filter.Tag & ","
    sSQL = sSQL & "Summary = " & IIf(Me.chkSummaryReport.Value = vbChecked, 1, 0) & "," & _
           "PrintFilterHeader = " & IIf(Me.chkPrintFilterHeader.Value = vbChecked, 1, 0) & "," '& _
           "DefaultOutput = 0," & _
           "DefaultExportTo = 0," & _
           "DefaultSave = 0," & _
           "DefaultSaveAs = ''," & _
           "DefaultCloseApp = 0,"
           
    sSQL = sSQL & _
        " OutputPreview = " & IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        " OutputFormat = " & CStr(mobjOutputDef.GetSelectedFormatIndex) & ", " & _
        " OutputScreen = " & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", "

    If chkDestination(desPrinter).Value = vbChecked Then
      sSQL = sSQL & _
        " OutputPrinter = 1, " & _
        " OutputPrinterName = '" & Replace(cboPrinterName.Text, " '", "''") & "', "
    Else
      sSQL = sSQL & _
        " OutputPrinter = 0, " & _
        " OutputPrinterName = '', "
    End If
    
    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & _
        "OutputSave = 1, " & _
        "OutputSaveExisting = " & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
        '"OutputSaveFormat = " & Val(txtFilename.Tag) & ", " &
    Else
      sSQL = sSQL & _
        "OutputSave = 0, " & _
        "OutputSaveExisting = 0, "
        '"OutputSaveFormat = 0, " &
    End If
        
    If chkDestination(desEmail).Value = vbChecked Then
      sSQL = sSQL & _
          "OutputEmail = 1, " & _
          "OutputEmailAddr = " & txtEmailGroup.Tag & ", " & _
          "OutputEmailSubject = '" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "OutputEmailAttachAs = '" & Replace(txtEmailAttachAs.Text, "'", "''") & "', " '& _
          "OutputEmailFileFormat = " & CStr(Val(txtEmailAttachAs.Tag)) & ", "
    Else
      sSQL = sSQL & _
          "OutputEmail = 0, " & _
          "OutputEmailAddr = 0, " & _
          "OutputEmailSubject = '', " & _
          "OutputEmailAttachAs = '', " '& _
          "OutputEmailFileFormat = 0,"
    End If
    
    sSQL = sSQL & _
        "OutputFilename = '" & Replace(txtFilename.Text, "'", "''") & "', "
    
    If chkIgnoreZeros.Value = vbChecked Then
      sSQL = sSQL & _
        "IgnoreZeros = 1, "
    Else
      sSQL = sSQL & _
        "IgnoreZeros = 0, "
    End If
    
    sSQL = sSQL & _
          "Parent1AllRecords = " & IIf(Me.optParent1AllRecords.Value, 1, 0) & "," & _
          "Parent1Picklist = " & IIf(Me.optParent1Picklist.Value, Me.txtParent1Picklist.Tag, 0) & "," & _
          "Parent2AllRecords = " & IIf(Me.optParent2AllRecords.Value, 1, 0) & "," & _
          "Parent2Picklist = " & IIf(Me.optParent2Picklist.Value, Me.txtParent2Picklist.Tag, 0) & " " & _
          "WHERE ID = " & mlngCustomReportID
            
           'Don't update user !
           '"Username = '" & gsUserName & "' "
           
    datData.ExecuteSql (sSQL)
    
    Call UtilUpdateLastSaved(utlCustomReport, mlngCustomReportID)
    
  Else
  
    ' Construct the SQL Insert string (Adding a new definition)
    
    sSQL = "Insert ASRSYSCustomReportsName (" & _
           "Name, Description, BaseTable, " & _
           "AllRecords, Picklist, Filter, " & _
           "Parent1Table, Parent1Filter, " & _
           "Parent2Table, Parent2Filter, "
    sSQL = sSQL & "Summary, PrintFilterHeader, " & _
           "UserName," & _
           "Parent1AllRecords, Parent1Picklist, Parent2AllRecords, Parent2Picklist, IgnoreZeros, " & _
           "OutputPreview, OutputFormat, OutputScreen, OutputPrinter, " & _
           "OutputPrinterName, OutputSave, OutputSaveExisting, OutputEmail, " & _
           "OutputEmailAddr, OutputEmailSubject, OutputEmailAttachAs, OutputFileName " & _
           ") "
    
           '"DefaultOutput, DefaultExportTo, DefaultSave, DefaultSaveAs, DefaultCloseApp, " & _

    sSQL = sSQL & _
           "Values('" & _
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
    
    sSQL = sSQL & ", " & CStr(txtParent1.Tag)
    sSQL = sSQL & ", " & IIf(optParent1Filter, CStr(txtParent1Filter.Tag), "0")
    sSQL = sSQL & ", " & CStr(txtParent2.Tag)
    sSQL = sSQL & ", " & IIf(optParent2Filter, CStr(txtParent2Filter.Tag), "0")
    
    If Me.chkSummaryReport.Value Then sSQL = sSQL & ", 1" Else sSQL = sSQL & ", 0"
    If Me.chkPrintFilterHeader.Value Then sSQL = sSQL & ", 1" Else sSQL = sSQL & ", 0"
    sSQL = sSQL & ", '" & datGeneral.UserNameForSQL & "',"

    'sSQL = sSQL & "0,"
    'sSQL = sSQL & "0,"
    'sSQL = sSQL & "0,"
    'sSQL = sSQL & "'',"
    'sSQL = sSQL & "0,"
    
    sSQL = sSQL & IIf(optParent1AllRecords, "1", "0") & ","
    sSQL = sSQL & IIf(optParent1Picklist, CStr(txtParent1Picklist.Tag), "0") & ","
    sSQL = sSQL & IIf(optParent2AllRecords, "1", "0") & ","
    sSQL = sSQL & IIf(optParent2Picklist, CStr(txtParent2Picklist.Tag), "0") & ", "

    If chkIgnoreZeros.Value = vbChecked Then
      sSQL = sSQL & " 1, "
    Else
      sSQL = sSQL & " 0, "
    End If

    'Output Options
    sSQL = sSQL & CStr(IIf(chkPreview.Value = vbChecked, "1", "0")) & ","                 'OutputPreview
    sSQL = sSQL & CStr(mobjOutputDef.GetSelectedFormatIndex) & ","                        'OutputFormat
    sSQL = sSQL & CStr(IIf(chkDestination(desScreen).Value = vbChecked, "1", "0")) & ","  'OutputScreen
    
    If chkDestination(desPrinter).Value = vbChecked Then
      sSQL = sSQL & "1, "                                                                 'OutputPrinter
      sSQL = sSQL & "'" & Replace(cboPrinterName.Text, "'", "''") & "',"                  'OutputPrinterName
    Else
      sSQL = sSQL & "0, "                                                                 'OutputPrinter
      sSQL = sSQL & "'',"                                                                 'OutputPrinterName
    End If
    
    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & "1, " & _
        cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "    'OutputSave, OutputSaveFormat, OutputSaveExisting
    Else
      sSQL = sSQL & "0, 0, "                                       'OutputSave, OutputSaveFormat, OutputSaveExisting
    End If


    If chkDestination(desEmail).Value = vbChecked Then
      sSQL = sSQL & "1, " & _
          txtEmailGroup.Tag & ", " & _
          "'" & Replace(txtEmailSubject.Text, "'", "''") & "', " & _
          "'" & Replace(txtEmailAttachAs.Text, "'", "''") & "', " '& _
          CStr(Val(txtEmailAttachAs.Tag)) & ", "      'OutputEmail, OutputEmailAddr, OutputEmailSubject, OutputEmailFileFormat
    Else
      sSQL = sSQL & "0, 0, '', '', "  '0, "   'OutputEmail, OutputEmailAddr, OutputEmailSubject, OutputEmailFileFormat
    End If

    sSQL = sSQL & _
        "'" & Replace(txtFilename.Text, "'", "''") & "'"  'OutputFilename
        
    sSQL = sSQL & ")"

    mlngCustomReportID = InsertCustomReport(sSQL)
    
    Call UtilCreated(utlCustomReport, mlngCustomReportID)
  
  End If
  
  SaveAccess
  
  '########################### 2 Of 3 - SAVE THE CHILD DETAILS
  ' First, remove any records from the child detail tables.
  ClearChildTables mlngCustomReportID
  
  InsertChildDetails
  
  
  '########################### 3 Of 3 - SAVE THE COL DETAILS
  
  ' First, remove any records from the 2 detail tables.
  ClearDetailTables mlngCustomReportID
  
  For iLoop = 1 To ListView2.ListItems.Count
    
    sKey = ListView2.ListItems(iLoop).Key
    
    Set objCol = mcolCustomReportColDetails.Item(sKey)
    
    sSQL = "INSERT ASRSysCustomReportsDetails (" & _
           "CustomReportID, Sequence, Type, " & _
           "ColExprID, Heading, Size, DP, " & _
           "IsNumeric, Avge, Cnt, Tot, Hidden, GroupWithNextColumn, " & _
           "SortOrderSequence, SortOrder, " & _
           "Boc, Poc, Voc, Srv, Repetition) "
           
    sSQL = sSQL & _
           "VALUES(" & _
            mlngCustomReportID & "," & _
            iLoop & "," & _
            "'" & Left(sKey, 1) & "'," & _
            Right(sKey, Len(sKey) - 1) & "," & _
            "'" & Replace(objCol.Heading, "'", "''") & "'," & _
            objCol.Size & "," & _
            objCol.DecPlaces & "," & _
            IIf(objCol.IsNumeric, 1, 0) & "," & _
            IIf(objCol.Average, 1, 0) & "," & _
            IIf(objCol.Count, 1, 0) & "," & _
            IIf(objCol.Total, 1, 0) & "," & _
            IIf(objCol.Hidden, 1, 0) & "," & _
            IIf(objCol.GroupWithNext, 1, 0) & ","
            '"'" & objCol.Heading & "'," &

    sSQL = sSQL & IsInSortOrder(sKey)
    sSQL = sSQL & IsInRepetition(sKey)
    
    datData.ExecuteSql (sSQL)
    
  Next iLoop
  
  Set objCol = Nothing
  
  SaveDefinition = True
  Changed = False
  
  Exit Function

Save_ERROR:

  SaveDefinition = False
  MsgBox "Warning : An error has occurred whilst saving..." & vbCrLf & Err.Description & vbCrLf & "Please cancel and try again. If this error continues, delete the definition.", vbCritical + vbOKOnly, "Custom Reports"

End Function


Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysCustomReportAccess WHERE ID = " & mlngCustomReportID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysCustomReportAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngCustomReportID & ", sysusers.name," & _
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
      sSQL = "IF EXISTS (SELECT * FROM ASRSysCustomReportAccess" & _
        " WHERE ID = " & CStr(mlngCustomReportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysCustomReportAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngCustomReportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub




Private Function IsInSortOrder(sKey As String) As String

  ' FUNCTION COMPLETE 28/01/00
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
'  With frmCustomReports.grdReportOrder
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

Private Function IsInRepetition(sKey As String) As String
 
  Dim bm As Variant
  Dim LLoop As Long
  
  With grdRepetition
    .MoveFirst
      Do Until LLoop = .Rows
        bm = .GetBookmark(LLoop)
        If .Columns(0).CellText(bm) = sKey Then
          IsInRepetition = "," & IIf(.Columns(2).CellValue(bm), 1, 0) & ")"
          Exit Function
        End If
        LLoop = LLoop + 1
      Loop
  
  
  End With
  
  IsInRepetition = ",-1)"

End Function

Private Function InsertChildDetails() As String

  Dim pvarbookmark  As Variant
  Dim i As Integer
  Dim sSQL As String
  
  sSQL = "("
  
  With grdChildren
    .MoveFirst
    For i = 0 To .Rows - 1 Step 1
      pvarbookmark = .GetBookmark(i)
      
      sSQL = "INSERT INTO ASRSysCustomReportsChildDetails "
      sSQL = sSQL & "VALUES ("
      sSQL = sSQL & mlngCustomReportID & ","
      sSQL = sSQL & .Columns("TableID").CellValue(pvarbookmark) & ","
      sSQL = sSQL & IIf(.Columns("FilterID").CellValue(pvarbookmark) = vbNullString, 0, .Columns("FilterID").CellValue(pvarbookmark)) & ","
      sSQL = sSQL & IIf(.Columns("Records").CellValue(pvarbookmark) = sALL_RECORDS, 0, .Columns("Records").CellValue(pvarbookmark)) & ","
      sSQL = sSQL & IIf(.Columns("OrderID").CellValue(pvarbookmark) = vbNullString, 0, .Columns("OrderID").CellValue(pvarbookmark))
      sSQL = sSQL & ")"
      
      datData.ExecuteSql (sSQL)
    Next i
  End With

End Function

Private Function RetrieveCustomReportDetails(plngCustomReportID As Long) As Boolean

  Dim rsTemp As Recordset
  Dim iLoop As Integer
  Dim sText As String
  Dim fAlreadyNotified As Boolean
  Dim sMessage As String
  Dim rsChildren As ADODB.Recordset
  Dim sSQL As String
  Dim tmpColumn As clsColumn
  
  On Error GoTo Load_ERROR
  
  'Load the basic guff first
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysCustomReportsName.*, " & _
                                     "CONVERT(integer, ASRSysCustomReportsName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysCustomReportsName WHERE ID = " & plngCustomReportID)
  
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This Report definition has been deleted by another user.", vbExclamation + vbOKOnly, "Custom Reports"
    Set rsTemp = Nothing
    RetrieveCustomReportDetails = False
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
  
  'Set Base Table
  mblnLoading = True
  LoadBaseCombo
  
  SetComboText cboBaseTable, datGeneral.GetTableName(rsTemp!BaseTable)
  mstrBaseTable = cboBaseTable.Text
  UpdateDependantFields
  
'  mblnLoading = False
  
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
  
  'TM20020109 Fault 3165
  'Get the all the child table information.
  sSQL = "SELECT  A.CustomReportID, A.ChildTable, B.TableName, A.ChildOrder, O.Name AS 'OrderName', A.ChildFilter, X.Name, A.ChildMaxRecords " & _
          "FROM ASRSysCustomReportsChildDetails A " & _
          "       INNER JOIN ASRSysTables B " & _
          "       ON A.ChildTable = B.TableID " & _
          "       LEFT OUTER JOIN ASRSysExpressions X " & _
          "       ON A.ChildFilter = X.ExprID " & _
          "       LEFT OUTER JOIN ASRSysOrders O " & _
          "       ON A.ChildOrder = O.OrderID " & _
          "WHERE CustomReportID = " & plngCustomReportID & " " & _
          "ORDER BY B.TableName "
  
  Set rsChildren = datGeneral.GetRecords(sSQL)
  
  With rsChildren
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        If rsChildren!childFilter > 0 Then
          sText = IsFilterValid(rsChildren!childFilter)
      
          Me.grdChildren.AddItem !ChildTable & vbTab & _
                                  !TableName & vbTab & _
                                  !OrderName & vbTab & _
                                  !childorder & vbTab & _
                                  !childFilter & vbTab & _
                                  !Name & vbTab & _
                                  IIf(!ChildMaxRecords = 0, "All Records", !ChildMaxRecords)
        Else
            Me.grdChildren.AddItem !ChildTable & vbTab & _
                                    !TableName & vbTab & _
                                    !OrderName & vbTab & _
                                    !childorder & vbTab & _
                                    0 & vbTab & _
                                    vbNullString & vbTab & _
                                    IIf(!ChildMaxRecords = 0, "All Records", !ChildMaxRecords)
                                    
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsChildren = Nothing
  
  
  
  'Set Report options
  If rsTemp!Summary Then chkSummaryReport.Value = vbChecked
  If rsTemp!IgnoreZeros Then chkIgnoreZeros.Value = vbChecked
  
  If (rsTemp!PrintFilterHeader) And _
    ((rsTemp!Filter > 0) Or (rsTemp!picklist > 0)) Then chkPrintFilterHeader.Value = vbChecked
  
  'Set the default options
  optOutputFormat(rsTemp!OutputFormat).Value = True
  chkPreview.Value = IIf(rsTemp!OutputPreview, vbChecked, vbUnchecked)
'  chkDestination(desScreen).Value = IIf(rsTemp!OutputScreen, vbChecked, vbUnchecked)
'
'  chkDestination(desPrinter).Value = IIf(rsTemp!OutputPrinter, vbChecked, vbUnchecked)
'  SetComboText cboPrinterName, rsTemp!OutputPrinterName
'  If rsTemp!OutputPrinterName <> vbNullString Then
'    If cboPrinterName.Text <> rsTemp!OutputPrinterName Then
'      cboPrinterName.AddItem rsTemp!OutputPrinterName
'      cboPrinterName.ListIndex = cboPrinterName.NewIndex
'      MsgBox "This definition is set to output to printer " & rsTemp!OutputPrinterName & _
'             " which is not set up on your PC.", vbInformation, Me.Caption
'    End If
'  End If
'
'  chkDestination(desSave).Value = IIf(rsTemp!OutputSave, vbChecked, vbUnchecked)
'  If chkDestination(desSave).Value Then
'    txtFilename.Text = rsTemp!OutputFilename
'    txtFilename.Tag = rsTemp!OutputSaveFormat
'    SetComboItem cboSaveExisting, rsTemp!OutputSaveExisting
'  End If
'
'  chkDestination(desEmail).Value = IIf(rsTemp!OutputEmail, vbChecked, vbUnchecked)
'  If rsTemp!OutputEmail Then
'    txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
'    txtEmailGroup.Tag = rsTemp!OutputEmailAddr
'    txtEmailSubject.Text = rsTemp!OutputEmailSubject
'    txtEmailAttachAs.Text = IIf(IsNull(rsTemp!OutputEmailAttachAs), vbNullString, rsTemp!OutputEmailAttachAs)
'    txtEmailAttachAs.Tag = rsTemp!OutputEmailFileFormat
'  End If
  mobjOutputDef.PopulateOutputControls rsTemp
  
  mlngTimeStamp = rsTemp!intTimestamp
  
  '=========================
  
  mblnReadOnly = Not datGeneral.SystemPermission("CUSTOMREPORTS", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlCustomReport, mlngCustomReportID) = ACCESS_READONLY)
  End If
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    cboTblAvailable.Enabled = True
    cboTblAvailable.BackColor = vbButtonFace
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  '=========================
  
  sMessage = vbNullString
  
  ' Now load the columns guff
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & plngCustomReportID & " ORDER BY [Sequence]")
  
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "Cannot load the column definition for this Custom Report", vbExclamation + vbOKOnly, "Custom Reports"
    RetrieveCustomReportDetails = False
    Set rsTemp = Nothing
    Exit Function
  End If
  
  Do Until rsTemp.EOF
  
    If rsTemp!Type = "C" Then
      sText = datGeneral.GetTableName(datGeneral.GetColumnTable(rsTemp!ColExprID)) & "." & datGeneral.GetColumnName(rsTemp!ColExprID)
      ListView2.ListItems.Add , rsTemp!Type & CStr(rsTemp!ColExprID), sText, ImageList1.ListImages("IMG_TABLE").Index, ImageList1.ListImages("IMG_TABLE").Index
    Else
      sText = datGeneral.GetExpression(rsTemp!ColExprID)
      'TM20010807 Fault 2656
      sMessage = IsCalcValid(rsTemp!ColExprID)
      If sMessage <> vbNullString _
        Or (GetExprField(rsTemp!ColExprID, "Access") = "HD" And Not mblnDefinitionCreator) Then
        If Not fAlreadyNotified Then
          If sMessage = vbNullString Then
            sMessage = "The calculation used in this definition has been made hidden by another user."
          Else
            sMessage = sMessage
          End If
  
          If FormPrint Then
            sMessage = "Custom Report print failed : " & vbCrLf & vbCrLf & sMessage
            MsgBox sMessage, vbExclamation + vbOKOnly, "Custom Reports"
            Me.Cancelled = True
            RetrieveCustomReportDetails = False
            Exit Function
          End If
          
          MsgBox sMessage & vbCrLf & _
                   "It will be removed from the definition.", vbExclamation + vbOKOnly, "Custom Reports"
  
          fAlreadyNotified = True
        End If
        mblnRecordSelectionInvalid = True
      Else
        ListView2.ListItems.Add , rsTemp!Type & CStr(rsTemp!ColExprID), sText, ImageList1.ListImages("IMG_CALC").Index, ImageList1.ListImages("IMG_CALC").Index
      End If
      
    End If
  
    ' Add to collection
    If sText <> vbNullString And sMessage = vbNullString Then
      Set tmpColumn = mcolCustomReportColDetails.Add(rsTemp!Type, rsTemp!ColExprID, rsTemp!Heading, rsTemp!Size, rsTemp!dp, rsTemp!Avge, rsTemp!cnt, rsTemp!tot, rsTemp!IsNumeric, rsTemp!Hidden, rsTemp!GroupWithNextColumn)
      tmpColumn.BreakOnChange = rsTemp!boc
      tmpColumn.PageOnChange = rsTemp!poc
      tmpColumn.ValueOnChange = rsTemp!voc
      tmpColumn.SurpressRepeatedValues = rsTemp!srv
      tmpColumn.Repetition = False
    End If
    sMessage = vbNullString
    rsTemp.MoveNext
  
  Loop
  
  mblnLoading = True
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Cancelled = True
    RetrieveCustomReportDetails = False
    Exit Function
  End If
  'mblnLoading = False
  
  PopulateTableAvailable , True
  
  UpdateButtonStatus (Me.SSTab1.Tab)
 
 ' Now do the sort order guff
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & plngCustomReportID & " AND SortOrderSequence > 0 AND Type = 'C' ORDER BY [SortOrderSequence]")
  
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "Cannot load the sort order for this Custom Report", vbExclamation + vbOKOnly, "Custom Reports"
    RetrieveCustomReportDetails = False
    Set rsTemp = Nothing
    Exit Function
  End If
  
  ' Add to the sort order grid
  Do Until rsTemp.EOF
    Me.grdReportOrder.AddItem rsTemp!ColExprID & vbTab & _
                              GetTableNameFromColumn(rsTemp!ColExprID) & "." & datGeneral.GetColumnName(rsTemp!ColExprID) & vbTab & _
                              rsTemp!SortOrder & vbTab & _
                              rsTemp!boc & vbTab & _
                              rsTemp!poc & vbTab & _
                              rsTemp!voc & vbTab & _
                              rsTemp!srv
    rsTemp.MoveNext
  Loop
  ' Tidyup
  Set rsTemp = Nothing
  
  With Me.grdReportOrder
    .SelBookmarks.RemoveAll
    .MoveFirst
    .SelBookmarks.Add (.Bookmark)
  End With
  
  Me.grdRepetition.RowHeight = lng_GRIDROWHEIGHT
  
 ' Now get the repetition columns
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & plngCustomReportID & " AND Repetition >= 0 ORDER BY [Sequence]")
  
  'TM20020606 Fault 3962 - allow an empty repetition grid and therfore allow the user to select
  'no columns from the base table.
  If Not (rsTemp.BOF And rsTemp.EOF) Then
    rsTemp.MoveFirst
    
    ' Add to the repetition grid.
    Do Until rsTemp.EOF
      If rsTemp!Type = "C" Then
        Me.grdRepetition.AddItem rsTemp!Type & rsTemp!ColExprID & vbTab & _
                                  GetTableNameFromColumn(rsTemp!ColExprID) & "." & datGeneral.GetColumnName(rsTemp!ColExprID) & vbTab & _
                                  IIf(rsTemp!Repetition, 1, 0)
      Else
        Me.grdRepetition.AddItem rsTemp!Type & rsTemp!ColExprID & vbTab & _
                                  datGeneral.GetExpression(rsTemp!ColExprID) & vbTab & _
                                  IIf(rsTemp!Repetition, 1, 0)
      End If
      
      mcolCustomReportColDetails.Item(rsTemp!Type & rsTemp!ColExprID).Repetition = rsTemp!Repetition
      
      rsTemp.MoveNext
    Loop

  End If
  ' Tidyup
  rsTemp.Close
  Set rsTemp = Nothing

  With Me.grdRepetition
    .SelBookmarks.RemoveAll
    .MoveFirst
  End With

  RetrieveCustomReportDetails = True
  Exit Function

Load_ERROR:

  MsgBox "Warning : Error whilst retrieving the report definition." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Custom Reports"
  RetrieveCustomReportDetails = False
  Set rsTemp = Nothing

End Function






Private Sub UpdateDependantFields()

  ' SUB COMPLETED 28/01/00
  ' This sub populates the parent/child combos depending
  ' on the base table selected
  
  Dim rsParents As New Recordset
  Dim rsTables As New Recordset
  Dim rsChildren As New Recordset
  Dim sSQL As String
  Dim lngTableID As Long

  lngTableID = 0
  If cboBaseTable.ListIndex <> -1 Then
    lngTableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)
  End If

  'If mblnLoading Then Exit Sub
  
  ' Get the parent(s) of the selected base table

  sSQL = "SELECT asrsystables.tablename, asrsystables.tableid " & _
         "FROM asrsystables " & _
         "WHERE asrsystables.tableid in " & _
         "(select parentid from asrsysrelations " & _
         "WHERE childid = " & CStr(lngTableID) & ") " & _
         "ORDER BY tablename"
         
  Set rsParents = datData.OpenPersistentRecordset(sSQL, adOpenKeyset, adLockReadOnly)
  
  If Not rsParents.BOF And Not rsParents.EOF Then
    rsParents.MoveLast
    rsParents.MoveFirst
  End If
  
  Select Case rsParents.RecordCount
  
    Case 0
      mblnParent1Enabled = False
      mblnParent2Enabled = False
      
      txtParent1.Text = "" '"<None>"
      txtParent1.Tag = 0
      optParent1AllRecords.Value = True
      txtParent1Filter.Text = ""
      txtParent1Filter.Tag = 0
      txtParent1Picklist.Text = ""
      txtParent1Picklist.Tag = 0
      fraParent1.Enabled = False
      
      txtParent2.Text = "" '"<None>"
      txtParent2.Tag = 0
      optParent2AllRecords.Value = True
      txtParent2Filter.Text = ""
      txtParent2Filter.Tag = 0
      txtParent2Picklist.Text = ""
      txtParent2Picklist.Tag = 0
      fraParent2.Enabled = False
          
    Case 1
      mblnParent1Enabled = True
      mblnParent2Enabled = False
    
      txtParent1.Text = rsParents!TableName
      txtParent1.Tag = rsParents!TableID
      optParent1AllRecords.Value = True
      txtParent1Filter.Text = ""
      txtParent1Filter.Tag = 0
      txtParent1Picklist.Text = ""
      txtParent1Picklist.Tag = 0
      fraParent1.Enabled = True
      
      txtParent2.Text = "" '"<None>"
      txtParent2.Tag = 0
      optParent2AllRecords.Value = True
      txtParent2Filter.Text = ""
      txtParent2Filter.Tag = 0
      txtParent2Picklist.Text = ""
      txtParent2Picklist.Tag = 0
      fraParent2.Enabled = False
    
    Case 2
      mblnParent1Enabled = True
      mblnParent2Enabled = True
    
      txtParent1.Text = rsParents!TableName
      txtParent1.Tag = rsParents!TableID
      optParent1AllRecords.Value = True
      txtParent1Filter.Text = ""
      txtParent1Filter.Tag = 0
      txtParent1Picklist.Text = ""
      txtParent1Picklist.Tag = 0
      fraParent1.Enabled = True
      
      rsParents.MoveNext
      
      txtParent2.Text = rsParents!TableName
      txtParent2.Tag = rsParents!TableID
      optParent2AllRecords.Value = True
      txtParent2Filter.Text = ""
      txtParent2Filter.Tag = 0
      txtParent2Picklist.Text = ""
      txtParent2Picklist.Tag = 0
      fraParent2.Enabled = True
  End Select
    
  ' Clear recordset reference
  Set rsParents = Nothing
  
'  ' Clear Child Combo and add <None> entry
'
'  With cboChild
'    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'    mblnLoading = True
'    .ListIndex = 0
'    mblnLoading = False
'  End With
'
  ' Get the children of the selected base table
  sSQL = "SELECT asrsystables.tablename, asrsystables.tableid " & _
         "FROM asrsystables " & _
         "WHERE asrsystables.tableid in " & _
         "(select childid from asrsysrelations " & _
         "WHERE parentid = " & CStr(lngTableID) & ") " & _
         "ORDER BY tablename"

  Set rsChildren = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If rsChildren.BOF And rsChildren.EOF Then
    mblnChildsEnabled = False
    fraChild.Enabled = False
    fraButtons.Enabled = False
    cmdAddChild.Enabled = False
    cmdEditChild.Enabled = False
    cmdRemove.Enabled = False
    cmdRemoveAllChilds.Enabled = False
    grdChildren.Enabled = False
  Else
    mblnChildsEnabled = True
    fraChild.Enabled = True
    grdChildren.Enabled = True
  End If
  Set rsChildren = Nothing

End Sub


Public Sub PopulateTableAvailable(Optional pstrTable As String, Optional pbSetToBase As Boolean)
  
  'TM20020424 Fault 3715 - have added optional pbSetToBase to the sub, so the changing of the
  'cboTblAvailable.listindex property is only called when this is true.
  
  ' SUB COMPLETE 28/01/00
  ' Now populate the TableAvailable combo
  
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

  ' Add the parent1 table if it exists
  'If txtParent1.Text <> "<None>" Then
  If txtParent1.Text <> "" Then
    If Not TableAlreadyAvailable(txtParent1.Tag) Or pbSetToBase Then
      cboTblAvailable.AddItem txtParent1.Text
      cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = txtParent1.Tag
    End If
  End If
  
  ' Add the parent2 table if it exists
  'If txtParent2.Text <> "<None>" Then
  If txtParent2.Text <> "" Then
    If Not TableAlreadyAvailable(txtParent2.Tag) Or pbSetToBase Then
      cboTblAvailable.AddItem txtParent2.Text
      cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = txtParent2.Tag
    End If
  End If
  
  'TM20020108 Fault 3165
  ' Add the child tables to the combo if selected
  Dim i As Integer
  Dim pvarbookmark As Variant
  With grdChildren
    If .Rows > 0 Then
      For i = 0 To .Rows - 1 Step 1
        pvarbookmark = .AddItemBookmark(i)
        If Not TableAlreadyAvailable(CInt(.Columns("TableID").CellValue(pvarbookmark))) Or pbSetToBase Then
          cboTblAvailable.AddItem .Columns("Table").CellValue(pvarbookmark)
          cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = CInt(.Columns("TableID").CellValue(pvarbookmark))
        End If
      Next i
    End If
  End With
  
  
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
    cboTblAvailable.BackColor = IIf(mblnReadOnly, vbButtonFace, vbWindowBackground)
  End If

End Sub


Public Function AnyChildColumnsUsed(lngTableID As Long, Optional bAutoYes As Boolean) As Integer

  ' FUNCTION COMPLETE 28/01/00
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
              , vbYesNo + vbQuestion, "Custom Reports") = vbNo Then
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
        mcolCustomReportColDetails.Remove tempKey
      
        ' also remove from the sort order if its there
        Me.grdReportOrder.MoveFirst
        
        Dim i As Integer
        For i = 0 To (Me.grdReportOrder.Rows - 1)
          If Right(tempKey, Len(tempKey) - 1) = Me.grdReportOrder.Columns(0).CellValue(Me.grdReportOrder.Bookmark) Then
            ' delete
            If Me.grdReportOrder.Rows = 1 Then
              Me.grdReportOrder.RemoveAll
            Else
              Me.grdReportOrder.RemoveItem (Me.grdReportOrder.AddItemRowIndex(Me.grdReportOrder.Bookmark))
            End If
          End If
        Me.grdReportOrder.MoveNext
        Next i
  
      End If
    ElseIf Left(objItem.Key, 1) = "E" Then
      lngExpr = Right(objItem.Key, Len(objItem.Key) - 1)
      Set rsCalc = datGeneral.GetReadOnlyRecords("SELECT ExprID, Name, TableID FROM ASRSysExpressions WHERE ExprID = " & CStr(lngExpr))
       
      If Not (rsCalc.BOF And rsCalc.EOF) Then
        If rsCalc!TableID = lngTableID Then
          tempKey = objItem.Key
          ListView2.ListItems.Remove tempKey
          mcolCustomReportColDetails.Remove tempKey
        End If
      End If
      Set rsCalc = Nothing
    End If
  Next iLoop
  
  Set objItem = Nothing
  AnyChildColumnsUsed = 2
  
End Function

Public Function ValidateDefinition(lngCurrentID As Long) As Boolean

  Dim LLoop As Long
  Dim bm As Variant
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim strRecSelStatus As String
  
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
    MsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SetFocus
    Exit Function
  End If
  
  'Check if this definition has been changed by another user
  Call UtilityAmended(utlCustomReport, mlngCustomReportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName.Text = gsUserName
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngCustomReportID = 0
    mblnReadOnly = False
    ForceAccess
  End If
  
  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngCustomReportID) Then
    MsgBox "A Custom Report definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    Exit Function
  End If
  
  ' BASE TABLE - If using a picklist, check one has been selected
  If optBasePicklist.Value Then
    If txtBasePicklist.Text = "" Or txtBasePicklist.Tag = "0" Or txtBasePicklist.Tag = "" Then
      MsgBox "You must select a picklist, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 0
      cmdBasePicklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' BASE TABLE - If using a filter, check one has been selected
  If optBaseFilter.Value Then
    If txtBaseFilter.Text = "" Or txtBaseFilter.Tag = "0" Or txtBaseFilter.Tag = "" Then
      MsgBox "You must select a filter, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 0
      cmdBaseFilter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 1 TABLE - If using a picklist, check one has been selected
  If optParent1Picklist.Value Then
    If txtParent1Picklist.Text = "" Or txtParent1Picklist.Tag = "0" Or txtParent1Picklist.Tag = "" Then
      MsgBox "You must select a picklist, or change the record selection for your first parent table.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 1
      cmdParent1Picklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 1 TABLE - If using a filter, check one has been selected
  If optParent1Filter.Value Then
    If txtParent1Filter.Text = "" Or txtParent1Filter.Tag = "0" Or txtParent1Filter.Tag = "" Then
      MsgBox "You must select a filter, or change the record selection for your first parent table.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 1
      cmdParent1Filter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
    
  ' PARENT 2 TABLE - If using a picklist, check one has been selected
  If optParent2Picklist.Value Then
    If txtParent2Picklist.Text = "" Or txtParent2Picklist.Tag = "0" Or txtParent2Picklist.Tag = "" Then
      MsgBox "You must select a picklist, or change the record selection for your second parent table.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 1
      cmdParent2Picklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' PARENT 2 TABLE - If using a filter, check one has been selected
  If optParent2Filter.Value Then
    If txtParent2Filter.Text = "" Or txtParent2Filter.Tag = "0" Or txtParent2Filter.Tag = "" Then
      MsgBox "You must select a filter, or change the record selection for your second parent table.", vbExclamation + vbOKOnly, "Custom Reports"
      SSTab1.Tab = 1
      cmdParent2Filter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
  
  ' Check that there are columns defined in the report definition
  If ListView2.ListItems.Count = 0 Then
    MsgBox "You must select at least 1 column for your report.", vbExclamation + vbOKOnly, "Custom Reports"
    SSTab1.Tab = 2
    Exit Function
  End If
  
  ' Check that at least one column has VOC ticked if it is a summary report.
  If chkSummaryReport.Value Then
  
    Dim blnHasVOC As Boolean
    
    blnHasVOC = False
    
    With grdReportOrder
      .MoveFirst
      Do Until LLoop = .Rows
        bm = .GetBookmark(LLoop)
        If .Columns("Value").CellValue(bm) Then
          blnHasVOC = True
          Exit Do
        End If
        LLoop = LLoop + 1
      Loop
    End With
    
    If Not blnHasVOC Then
      If MsgBox("You have defined this report as a summary report but have not set a column as 'Value on Change'." & vbCrLf & vbCrLf & _
                 "Do you wish to continue?", vbQuestion + vbYesNo, "Custom Reports") = vbNo Then
        ValidateDefinition = False
        SSTab1.Tab = 3
        Exit Function
      End If
    End If
  End If
  
  ' Check that not BOC and POC on the same field
  With grdReportOrder
    .MoveFirst
    Do Until LLoop = .Rows
      bm = .GetBookmark(LLoop)
      If .Columns(3).CellValue(bm) And .Columns(4).CellValue(bm) Then
        MsgBox "You cannot select 'Break on Change' and 'Page Break' for the same column.", vbExclamation + vbOKOnly, "Custom Reports"
        ValidateDefinition = False
        Exit Function
      End If
      LLoop = LLoop + 1
    Loop
  End With
  
  ' Check that at least 1 column has been defined as the report order
  With grdReportOrder
    If .Rows = 0 Then
      MsgBox "You must select at least one column to order the report by.", vbExclamation + vbOKOnly, "Custom Reports"
      ValidateDefinition = False
      SSTab1.Tab = 3
      Exit Function
    End If
  End With
  
'  ' If exporting, and save box is checked, user must specify a filename
'  If optOutput(1).Value Then
'    If chkSave.Value And txtExportFilename = "" Then
'      MsgBox "You must select a filename if you opt to save the document !", vbExclamation + vbOKOnly, "Custom Reports"
'      ValidateDefinition = False
'      SSTab1.Tab = 4
'      Exit Function
'    End If
'  End If
  If Not mobjOutputDef.ValidDestination Then
    SSTab1.Tab = 4
    Exit Function
  End If
  
If mlngCustomReportID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlCustomReport, _
      CStr(mlngCustomReportID), _
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
               vbExclamation + vbOKOnly, "Custom Reports"
      Else
        MsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Custom Reports"
      End If

      Screen.MousePointer = vbNormal
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If MsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Custom Reports") = vbNo Then
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
  
  ' Check the no. of items in the collection is the same as the
  ' number of items in the list view.
  If Not ValidateCollection Then
    ValidateDefinition = False
    Exit Function
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

  ' FUNCTION COMPLETE 28/01/00
  ' Is there already a definition with the same name (that isnt the
  ' definition we are editing ?)
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSYSCustomReportsName " & _
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


Public Sub PopulateAvailable()
  
  ' This function is called whenever a new table is selected in the
  ' table combo, or when cols/expressions are removed from the report
  ' definition. It checks through each item in the 'Selected'
  ' listview and if it doesnt find them, it adds them to the
  ' 'Available' listview.

  Dim rsColumns As New Recordset
  Dim rsCalculations As New Recordset
  Dim sSQL As String
  Dim intCount As Integer
  Dim fOK As Boolean
  
  If cboBaseTable.ListIndex = -1 Then Exit Sub

  Screen.MousePointer = vbHourglass
  
  ' Clear the contents of the Available Listview
  ListView1.ListItems.Clear
  
  If optColumns.Value Then
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
  Else
 
    ' Add the Expressions of the selected table to the listview
    sSQL = "SELECT ExprID, Name FROM ASRSysExpressions " & _
           "WHERE TableID = " & cboTblAvailable.ItemData(cboTblAvailable.ListIndex) & " " & _
           " AND Type = " & Trim(Str(giEXPR_RUNTIMECALCULATION)) & _
           " AND ParentComponentID = 0" & _
           " AND ((Access <> 'HD') OR (Access = 'HD' AND Username = '" & datGeneral.UserNameForSQL & "')) " & _
           "ORDER BY Name"
    
    Set rsCalculations = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    ' Check if the column has already been selected. If so, dont add it
    ' to the available listview
    Do While Not rsCalculations.EOF
      If Not AlreadyUsed(CStr("E" & rsCalculations!ExprID)) Then
        If IsCalcValid(rsCalculations!ExprID) = vbNullString Then
          ListView1.ListItems.Add , "E" & rsCalculations!ExprID, rsCalculations!Name, , ImageList1.ListImages("IMG_CALC").Index
        End If
      End If
      rsCalculations.MoveNext
    Loop
  
    ' Clear recordset reference
    Set rsCalculations = Nothing
  
  End If
  
  ' We are viewing the base table, so adjust the listview height and make
  ' the New Calculation command button visible
  cmdNewCalculation.Visible = True
  cmdNewCalculation.Enabled = (Not mblnReadOnly)
'  ListView1.Height = 4275
'  ListView1.Height = cmdNewCalculation.Top - 100 - ListView1.Top

  Screen.MousePointer = vbDefault

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

Public Sub LoadBaseCombo()

  ' Loads the Base combo with all tables (even lookups)
  
  Dim sSQL As String
  Dim rsTables As New Recordset

  sSQL = "Select TableName, TableID From ASRSysTables" '- WHERE TableType='1' OR TableType='2'"  ' * UNCOMMENT TO EXCLUDE LOOKUPS
  sSQL = sSQL & " ORDER BY TableName"
  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With cboBaseTable
    .Clear
    Do While Not rsTables.EOF
      .AddItem rsTables!TableName
      .ItemData(.NewIndex) = rsTables!TableID
      rsTables.MoveNext
    Loop
    'If .ListCount > 0 Then .ListIndex = 0
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

Public Sub EnableDisableTabControls()

  Dim mblnWasNotChanged As Boolean
  Dim objItem As ListItem
  Dim iLoop As Integer
  
  Const lng_FILTERSCROLL = 2620
  Const lng_FILTERNOSCROLL = 2850
  
  
  Select Case SSTab1.Tab
    Case 1 'Related Tables Tab
      RefreshChildrenGrid
      
    Case 3 'Sort Order Tab
      RefreshReportOrderGrid
      RefreshRepetitionGrid
      'grdReportOrder.MoveFirst
    End Select
    
  If mblnReadOnly Then
    Exit Sub
  End If
    
   mblnWasNotChanged = Changed
     
  'NHRD28042003 Fault 5165 Added an extra frame around command buttons on Tab 2
  'then sorted out the enabling and disabling of frames depending on which frame is
  'selected.  This stops the taborder 'going off' to other tabs and then coming back.
  
  ' TAB 0 CONTROLS
  fraInformation.Enabled = (SSTab1.Tab = 0)
  fraBase.Enabled = (SSTab1.Tab = 0)

  ' TAB 1 CONTROLS
  fraParent1.Enabled = ((SSTab1.Tab = 1) And (mblnParent1Enabled))
  fraParent2.Enabled = ((SSTab1.Tab = 1) And (mblnParent2Enabled))
  fraChild.Enabled = ((SSTab1.Tab = 1) And (mblnChildsEnabled))
  
  lblParent1Table.Enabled = fraParent1.Enabled
  lblParent1Records.Enabled = fraParent1.Enabled
  optParent1AllRecords.Enabled = fraParent1.Enabled
  optParent1Picklist.Enabled = fraParent1.Enabled
  optParent1Filter.Enabled = fraParent1.Enabled
  
  fraParent2.Enabled = (Len(txtParent2.Text) > 0)
  lblParent2Table.Enabled = fraParent2.Enabled
  lblParent2Records.Enabled = fraParent2.Enabled
  optParent2AllRecords.Enabled = fraParent2.Enabled
  optParent2Picklist.Enabled = fraParent2.Enabled
  optParent2Filter.Enabled = fraParent2.Enabled

  UpdateButtonStatus (Me.SSTab1.Tab)
    
  ' TAB 2 CONTROLS
  fraFieldsAvailable.Enabled = (SSTab1.Tab = 2)
  fraFieldsSelected.Enabled = (SSTab1.Tab = 2)
  'NHRD Included this line to get rid of the frames border.
  'Otherwise it just looks silly.
  fraButtons.Enabled = (SSTab1.Tab = 2)
  fraButtons.BorderStyle = 0  'IIf(fraButtons.Enabled, 0, 1)

  ' TAB 3 CONTROLS
  fraReportOrder.Enabled = (SSTab1.Tab = 3)
  fraRepetition.Enabled = (SSTab1.Tab = 3)
  Me.grdRepetition.RowHeight = lng_GRIDROWHEIGHT
  
  ' TAB 4 CONTROLS
  fraReportOptions.Enabled = (SSTab1.Tab = 4)
  fraOutputFormat.Enabled = (SSTab1.Tab = 4)
  fraOutputDestination.Enabled = (SSTab1.Tab = 4)
  
  Select Case SSTab1.Tab
    Case 0 'definition tab
      chkPrintFilterHeader.Enabled = (Me.optBaseFilter.Value) Or (Me.optBasePicklist.Value)
       
    Case 1 'realted tables tab
      UpdateButtonStatus (SSTab1.Tab)
      
      'TM25092003 Fault 7056
'      If mblnChildsEnabled And cmdAddChild.Enabled Then
'        cmdAddChild.SetFocus
'      End If
      
    Case 2 ' columns tab
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
      
      UpdateButtonStatus (Me.SSTab1.Tab)
        
    Case 3 ' order tab
      UpdateOrderButtons
      
      cmdAddOrder.SetFocus
      
    Case 4 ' output tab
       
  End Select

  If mblnWasNotChanged = False Then Changed = False
  
End Sub

Private Sub ClearForNew()
  
  'Clear out all fields required to be blank for a new report definition
  
  optBaseAllRecords.Value = True
  txtBasePicklist.Text = ""
  txtBasePicklist.Tag = 0
  txtBaseFilter.Text = ""
  txtBaseFilter.Tag = 0
  
  txtParent1.Text = ""
  txtParent1.Tag = 0
  txtParent1Picklist.Text = ""
  txtParent1Picklist.Tag = 0
  txtParent1Filter.Text = ""
  txtParent1Filter.Tag = 0
  
  txtParent2.Text = ""
  txtParent2.Tag = 0
  txtParent2Picklist.Text = ""
  txtParent2Picklist.Tag = 0
  txtParent2Filter.Text = ""
  txtParent2Filter.Tag = 0
  
'  cboChild.Clear
'  cboChild.AddItem "<None>"
'  cboChild.ItemData(cboChild.NewIndex) = 0
'  cboChild.ListIndex = 0
'  txtChildFilter.Text = ""
'  txtChildFilter.Tag = 0
  
  grdChildren.RemoveAll
  
  ' Columns Tab
  txtProp_ColumnHeading = ""
  'txtProp_Size = 0
  spnSize.Text = 0
  
  'txtProp_DecPlaces = 0
  spnDec.Text = "0"
  chkProp_Average.Value = vbUnchecked
  chkProp_Count.Value = vbUnchecked
  chkProp_Total.Value = vbUnchecked
  ListView2.ListItems.Clear
  
  ' Order Tab
  grdReportOrder.RemoveAll
  cmdEditOrder.Enabled = False
  cmdDeleteOrder.Enabled = False
  chkSummaryReport.Value = vbUnchecked
  chkIgnoreZeros.Value = vbUnchecked
  chkPrintFilterHeader.Value = vbUnchecked
  grdRepetition.RemoveAll
  
  
  If mblnDefinitionCreator Then txtUserName = gsUserName
  
'  ' Default option bit
'  With cboExportTo
'    .Clear
'    .AddItem "Html"
'    .AddItem "Microsoft Excel"
'    .AddItem "Microsoft Word"
'    .ListIndex = 0
'  End With
  
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
    cboTblAvailable.AddItem cboBaseTable.Text
    cboTblAvailable.ItemData(cboTblAvailable.NewIndex) = cboBaseTable.ItemData(cboBaseTable.ListIndex)
  
    ' Add the parents of the base table to the combo
    sSQL = "SELECT ParentID FROM ASRSysRelations WHERE ChildID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
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
    sSQL = "SELECT ChildID FROM ASRSysRelations WHERE ParentID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
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
  'sSQL = sSQL & " Where TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
  'sSQL = "TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)

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

Private Function InsertCustomReport(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertCustomReport_ERROR

'  Dim rsCustomReport As Recordset
  
'  datData.ExecuteSql sSQL
      
'  sSQL = "Select Max(ID) From ASRSYSCustomReportsName"
'  Set rsCustomReport = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  InsertCustomReport = rsCustomReport(0)
  
'  rsCustomReport.Close
'  Set rsCustomReport = Nothing

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
    pmADO.Value = "AsrSysCustomReportsName"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "ID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertCustomReport = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertCustomReport = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertCustomReport_ERROR:
  
  fSavedOK = False
  Resume Next

End Function

Private Sub ClearDetailTables(plngCustomReportID As Long)

  ' Delete all column information from the Details table.
  
  Dim sSQL As String
  
  sSQL = "Delete From ASRSysCustomReportsDetails Where CustomReportID = " & plngCustomReportID
  datData.ExecuteSql sSQL

End Sub

Private Sub ClearChildTables(plngCustomReportID As Long)

  ' Delete all column information from the Details table.
  
  Dim sSQL As String
  
  sSQL = "Delete From ASRSysCustomReportsChildDetails Where CustomReportID = " & plngCustomReportID
  datData.ExecuteSql sSQL

End Sub

'Private Sub cboExportTo_Click()
'
'  ' If user changes exportto, and a filename has already been selected
'  ' then change the extension of the filename automatically
'
'  Dim sText As String
'
'  If txtExportFilename.Text <> "" Then
'    sText = txtExportFilename.Text
'    Select Case UCase(cboExportTo.Text)
'      Case "HTML":  Mid(sText, Len(txtExportFilename.Text) - 2, 3) = "htm"
'      Case "MICROSOFT EXCEL": Mid(sText, Len(txtExportFilename.Text) - 2, 3) = "xls"
'      Case "MICROSOFT WORD":  Mid(sText, Len(txtExportFilename.Text) - 2, 3) = "doc"
'    End Select
'    txtExportFilename.Text = sText
'  End If
'
'  Changed = True
'
'End Sub
'
'Private Sub chkSave_Click()
'
'  Changed = True
'
'  ' If save box unchecked, remove filename from textbox
'  cmdExportFilename.Enabled = chkSave.Value
'
'  If chkSave.Value = vbUnchecked Then
'    txtExportFilename.Text = ""
'    chkCloseApplication.Value = vbUnchecked
'    chkCloseApplication.Enabled = False
'  Else
'    chkCloseApplication.Enabled = True
'  End If
'
'End Sub
'
'Private Sub cmdExportFilename_Click()
'
'  ' Set flags depending on exportto and show the dialog
'  With CommonDialog1
'
'    Select Case UCase(cboExportTo.Text)
'
'      Case "MICROSOFT EXCEL": .DefaultExt = ".xls"
'                              .Filter = "Excel Spreadsheets (*.xls)|*.xls"
'      Case "HTML":            .DefaultExt = ".htm"
'                              .Filter = "HTML Documents (*.htm)|*.htm"
'      Case "MICROSOFT WORD":  .DefaultExt = ".doc"
'                              .Filter = "Word Documents (*.doc)|*.doc"
'    End Select
'
'    .CancelError = False
'    .DialogTitle = "Select a filename for your " & cboExportTo.Text & " document..."
'
'    If Len(Trim(txtExportFilename.Text)) = 0 Then
'      .InitDir = gsDocumentsPath
'    Else
'      .FileName = txtExportFilename.Text
'    End If
'
'    .FilterIndex = 1
'    .Flags = cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt + cdlOFNPathMustExist
'
'    .ShowSave
'
'    If .FileName = "" Then Exit Sub Else txtExportFilename.Text = .FileName
'    Changed = True
'  End With
'
'End Sub
'
'Private Sub optOutput_Click(Index As Integer)
'
'  Changed = True
'
'  'Enable/Disable as appropriate
'  cboExportTo.BackColor = IIf(Index = 1, vbWindowBackground, vbButtonFace)
'  cboExportTo.Enabled = (Index = 1)
'  chkSave.Enabled = (Index = 1)
'  chkCloseApplication.Enabled = (Index = 1) And (chkSave.Value = vbChecked)
'  lblFormat.Enabled = (Index = 1)
'
'  If Index = 0 Then
'    txtExportFilename.Text = ""
'    chkSave.Value = vbUnchecked
'    chkCloseApplication.Value = vbUnchecked
'  End If
'
'End Sub

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
  Dim sTableName As String
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
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
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
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
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
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
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
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
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

  ' Child Filters
  With grdChildren
    If .Rows > 0 Then
      For iLoop = .Rows - 1 To 0 Step -1
        varBookmark = .AddItemBookmark(iLoop)
        lngFilterID = .Columns("FilterID").CellValue(varBookmark)
        
        If lngFilterID > 0 Then
          fRemove = False
          iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
  
          sTableName = .Columns("Table").CellValue(varBookmark)

          Select Case iResult
            Case REC_SEL_VALID_HIDDENBYUSER
              ' Filter hidden by the current user.
              ' Only a problem if the current definition is NOT owned by the current user,
              ' or if the current definition is not already hidden.
              fRemove = (Not mblnDefinitionCreator) And _
                (Not mblnReadOnly) And _
                (Not FormPrint)

              If fRemove Then
                sBigMessage = "The '" & sTableName & "' table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
                MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
              Else
                fNeedToForceHidden = True
  
                ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
                asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & sTableName & "' table filter"
              End If

            Case REC_SEL_VALID_DELETED
              ' Filter deleted by another user.
              ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
              asDeletedParameters(UBound(asDeletedParameters)) = "'" & sTableName & "' table filter"

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)

            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                ' Calc hidden by another user.
                ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
                asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & sTableName & "' table filter"
  
                fRemove = (Not mblnReadOnly) And _
                  (Not FormPrint)
              End If
            Case REC_SEL_VALID_INVALID
              ' Calc invalid.
              ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
              asInvalidParameters(UBound(asInvalidParameters)) = "'" & sTableName & "' table filter"

              fRemove = (Not mblnReadOnly) And _
                (Not FormPrint)
          End Select
          
          
          If fRemove Then
            'TM25092003 Fault 7056
            ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
'            sRow = .Columns("TableID").CellValue(varBookmark) _
'              & vbTab & .Columns("Table").CellValue(varBookmark) _
'              & vbTab & 0 _
'              & vbTab & vbNullString _
'              & vbTab & .Columns("Records").CellValue(varBookmark)
            sRow = .Columns("TableID").CellValue(varBookmark) _
              & vbTab & .Columns("Table").CellValue(varBookmark) _
              & vbTab & .Columns("Order").CellValue(varBookmark) _
              & vbTab & .Columns("OrderID").CellValue(varBookmark) _
              & vbTab & 0 _
              & vbTab & vbNullString _
              & vbTab & .Columns("Records").CellValue(varBookmark)

            If .Rows > 1 Then
              .RemoveItem iLoop
            Else
              .RemoveAll
            End If
            .AddItem sRow, iLoop
            
            'TM25092003 Fault 7056
            If (Not FormPrint) And (Not mblnLoading) Then
              'JPD 20030728 Fault 6460
              SSTab1.Tab = 1
              .SetFocus
            End If

            mblnRecordSelectionInvalid = True
          End If
        End If
      Next iLoop
    End If
  End With

  ' Calcs
  With ListView2
    For iLoop = .ListItems.Count To 1 Step -1
      strColumnType = Left$(.ListItems(iLoop).Key, 1)
      lngColumnID = Val(Mid$(.ListItems(iLoop).Key, 2))
      
      If strColumnType = "E" Then
        fRemove = False
        iResult = ValidateCalculation(lngColumnID)
  
        sCalcName = .ListItems(iLoop).Text
  
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
              MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
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
          RemoveFromCollection .ListItems(iLoop).Key
          .ListItems.Remove iLoop
          EnableColProperties

          'TM25092003 Fault 7056
          If (Not FormPrint) And (Not mblnLoading) Then
            SSTab1.Tab = 2
            .SetFocus
          End If
  
          mblnRecordSelectionInvalid = True
        End If
      End If
    Next iLoop
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

    For iLoop = 1 To UBound(asHiddenByOtherParameters)
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
  
  RefreshRepetitionGrid
  
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
Public Sub PrintDef(lCustomReportID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsTemp2 As Recordset
  Dim rsColumns As Recordset
  Dim rsChildren As ADODB.Recordset
  Dim sSQL As String
  Dim lngTempX As Long
  Dim lngTempY As Long
  Dim sTemp As String
  Dim iLoop As Integer
  Dim fFirstLoop As Boolean
  Dim varBookmark As Variant

  mlngCustomReportID = lCustomReportID
  
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysCustomReportsName.*, " & _
                                     "CONVERT(integer, ASRSysCustomReportsName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysCustomReportsName WHERE ID = " & mlngCustomReportID)
                                        
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Print Definition"
    Set rsTemp = Nothing
    Exit Sub
  End If

  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader "Custom Report : " & rsTemp!Name
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
      
      If rsTemp!AllRecords Then
        .PrintNormal "Records : All Records"
      ElseIf rsTemp!picklist Then
        .PrintNormal "Records : '" & datGeneral.GetPicklistName(rsTemp!picklist) & "' picklist"
      ElseIf rsTemp!Filter Then
        .PrintNormal "Records : '" & datGeneral.GetFilterName(rsTemp!Filter) & "' filter"
      End If
      
      .PrintNormal
      .PrintNormal "Display filter or picklist title in the report header : " & IIf(rsTemp!PrintFilterHeader = True, "Yes", "No")
      .PrintNormal
      
      .PrintNormal "Parent 1 Table : " & IIf(rsTemp!parent1table > 0, datGeneral.GetTableName(rsTemp!parent1table), "<None>")
      If (rsTemp!parent1picklist > 0) Then
        .PrintNormal "Parent 1 Records : '" & datGeneral.GetPicklistName(rsTemp!parent1picklist) & "' picklist"
      ElseIf (rsTemp!parent1filter > 0) Then
        .PrintNormal "Parent 1 Records : '" & datGeneral.GetFilterName(rsTemp!parent1filter) & "' filter"
      Else
        If (rsTemp!parent1table > 0) Then
          .PrintNormal "Parent 1 Records : All Records"
        Else
          .PrintNormal "Parent 1 Records : N/A"
        End If
      End If
      
      .PrintNormal
      
      .PrintNormal "Parent 2 Table : " & IIf(rsTemp!parent2table > 0, datGeneral.GetTableName(rsTemp!parent2table), "<None>")
      If (rsTemp!parent2picklist > 0) Then
        .PrintNormal "Parent 2 Records : '" & datGeneral.GetPicklistName(rsTemp!parent2picklist) & "' picklist"
      ElseIf (rsTemp!parent2filter > 0) Then
        .PrintNormal "Parent 2 Records : '" & datGeneral.GetFilterName(rsTemp!parent2filter) & "' filter"
      Else
        If (rsTemp!parent2table > 0) Then
          .PrintNormal "Parent 2 Records : All Records"
        Else
          .PrintNormal "Parent 2 Records : N/A"
        End If
      End If
     
'      .PrintNormal "Child Table : " & IIf(rsTemp!ChildTable > 0, datGeneral.GetTableName(rsTemp!ChildTable), "<None>")
'      .PrintNormal "Child Filter : " & IIf(rsTemp!childFilter > 0, datGeneral.GetFilterName(rsTemp!childFilter), "<None>")
'      .PrintNormal "Max Child Records : " & IIf(rsTemp!ChildMaxRecords = 0, "All Records", rsTemp!ChildMaxRecords)
      
      Set rsChildren = datGeneral.GetRecords("SELECT * FROM ASRSysCustomReportsChildDetails WHERE CustomReportID = " & mlngCustomReportID & " ORDER BY ID")
      
      If Not (rsChildren.BOF And rsChildren.EOF) Then
        .PrintNormal
        .PrintBold "Child Tables"
        .PrintNormal
        Do While Not rsChildren.EOF
          .PrintNormal "Child Table : " & IIf(rsChildren!ChildTable > 0, datGeneral.GetTableName(rsChildren!ChildTable), "<None>")
          .PrintNormal "Child Filter : " & IIf(rsChildren!childFilter > 0, datGeneral.GetFilterName(rsChildren!childFilter), "<None>")
          '.PrintNormal "Child Order : " & IIf(rsChildren!childOrder > 0, datGeneral.GetOrder(rsChildren!childOrder), "<None>")
          
          If (rsChildren!childorder > 0) Then
            .PrintNormal "Child Order : " & GetChildOrder(rsChildren!childorder)
          Else
            .PrintNormal "Child Order : <None>"
          End If
          .PrintNormal "Child Records : " & IIf(rsChildren!ChildMaxRecords = 0, "All Records", rsChildren!ChildMaxRecords)
          .PrintNormal
          rsChildren.MoveNext
        Loop
        rsChildren.Close
        Set rsChildren = Nothing
      End If
      
      ' Now do the Columns Section

      Set rsColumns = datGeneral.GetRecords("SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & mlngCustomReportID & " ORDER BY ID")
    
      .PrintTitle "Columns"
      
      Do While Not rsColumns.EOF
          
        .PrintNormal "Type : " & IIf(rsColumns!Type = "C", "Column", "Calculation")
        
        If rsColumns!Type = "C" Then
          sTemp = datGeneral.GetColumnTableName(rsColumns!ColExprID) & "." & datGeneral.GetColumnName(rsColumns!ColExprID)
        Else
          Set rsTemp2 = datGeneral.GetRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & rsColumns!ColExprID)
          If Not rsTemp2.BOF And Not rsTemp2.EOF Then
            sTemp = rsTemp2.Fields("Name")
          End If
          Set rsTemp2 = Nothing
        End If
        
        .PrintNormal "Name : " & sTemp
        .PrintNormal "Heading : " & rsColumns!Heading
        .PrintNormal "Size : " & rsColumns!Size

        If rsColumns!IsNumeric Then
          .PrintNormal "Decimal Places : " & rsColumns!dp
        End If
        .PrintNormal "Count : " & rsColumns!cnt
        If rsColumns!IsNumeric Then
          .PrintNormal "Average : " & rsColumns!Avge
          .PrintNormal "Total : " & rsColumns!tot
        End If
        
        .PrintNormal "Hidden : " & rsColumns!Hidden
        .PrintNormal "Group with Next : " & rsColumns!GroupWithNextColumn
        
        Select Case rsColumns!Repetition
        Case 0
          .PrintNormal "Repetition : False"

        Case 1
          .PrintNormal "Repetition : True"
        Case Else
          'Don't print.
        End Select
        
        .PrintNormal " "
        
        rsColumns.MoveNext
      
      Loop
      
      ' Now do the Sort Order n Options Section
    
      Set rsColumns = datGeneral.GetRecords("SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & mlngCustomReportID & " AND SortOrderSequence > 0 AND Type = 'C' ORDER BY [SortOrderSequence]")
    
      .PrintTitle "Sort Order"
        
      Do While Not rsColumns.EOF
        .PrintNormal "Name : " & datGeneral.GetColumnName(rsColumns!ColExprID)
        .PrintNormal "Order : " & IIf(rsColumns!SortOrder = "Asc", "Ascending", "Descending")
        .PrintNormal "Break On Change : " & IIf(rsColumns!boc = True, "Yes", "No")
        .PrintNormal "Page On Change : " & IIf(rsColumns!poc = True, "Yes", "No")
        .PrintNormal "Value On Change : " & IIf(rsColumns!voc = True, "Yes", "No")
        .PrintNormal "Suppress Repeated Values : " & IIf(rsColumns!srv = True, "Yes", "No")
        .PrintNormal " "
        rsColumns.MoveNext
      Loop
        
        .PrintTitle "Output Options"
      
        .PrintNormal "Summary Report : " & IIf(rsTemp!Summary = True, "Yes", "No")
        .PrintNormal "Ignore zeros when calculating aggregates : " & IIf(rsTemp!IgnoreZeros = True, "Yes", "No")
    
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
        .PrintConfirm "Custom Report : " & rsTemp!Name, "Custom Report Definition"
      End If
    
    End With
  
  End If
  
  Set rsTemp = Nothing
  Set rsColumns = Nothing

Exit Sub

LocalErr:
  MsgBox "Printing Custom Report Definition Failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"

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

Private Sub txtProp_ColumnHeading_LostFocus()
  
  'TM20020906 Fault 4382 - Dont allow leading and trailing spaces on the column heading.
  txtProp_ColumnHeading.Text = Trim(txtProp_ColumnHeading.Text)
  
End Sub

Public Function GetChildOrder(lOrderID As Long) As String
'NHRD20030704 Fault 5693
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetOrder(lOrderID)", Array(lOrderID)
    
  Dim sSQL As String
  Dim rsOrder As ADODB.Recordset
      
  sSQL = "SELECT * " & _
      "FROM ASRSysOrders " & _
      "WHERE OrderID=" & lOrderID
  Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  GetChildOrder = rsOrder!Name
  
  Set rsOrder = Nothing
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function



