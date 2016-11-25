VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.Form frmConfiguration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1015
   Icon            =   "frmConfiguration.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Re&store Defaults"
      Height          =   400
      Left            =   90
      TabIndex        =   124
      Top             =   7290
      Width           =   1620
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7050
      Left            =   120
      TabIndex        =   127
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12435
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "&Display Defaults"
      TabPicture(0)   =   "frmConfiguration.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDisplay(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDisplay(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDisplay(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmReportsGeneral"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Reports && Utilities"
      TabPicture(1)   =   "frmConfiguration.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraReports(0)"
      Tab(1).Control(1)=   "fraReports(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Network Configuration"
      TabPicture(2)   =   "frmConfiguration.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmOutputs"
      Tab(2).Control(1)=   "fraNetwork(1)"
      Tab(2).Control(2)=   "fraNetwork(0)"
      Tab(2).Control(3)=   "frmAutoLogin"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Batch Login"
      TabPicture(3)   =   "frmConfiguration.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraBatch(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fraBatch(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "E&vent Log"
      TabPicture(4)   =   "frmConfiguration.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FraEventLog(0)"
      Tab(4).Control(1)=   "FraEventLog(1)"
      Tab(4).Control(2)=   "FraEventLog(2)"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Report Out&put"
      TabPicture(5)   =   "frmConfiguration.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).Control(1)=   "FraOutput(1)"
      Tab(5).Control(2)=   "FraOutput(0)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Tool&bars"
      TabPicture(6)   =   "frmConfiguration.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraToolbars"
      Tab(6).Control(1)=   "fraToolbarGeneral"
      Tab(6).ControlCount=   2
      Begin VB.Frame frmReportsGeneral 
         Caption         =   "Selection Screen :"
         Height          =   1335
         Left            =   120
         TabIndex        =   133
         Top             =   5310
         Width           =   6735
         Begin VB.CheckBox chkRememberDefSelID 
            Caption         =   "R&emember selection screen details"
            Height          =   195
            Left            =   200
            TabIndex        =   21
            Top             =   910
            Width           =   3795
         End
         Begin VB.CheckBox chkCloseDefsel 
            Caption         =   "C&lose selection screen after run"
            Height          =   285
            Left            =   200
            TabIndex        =   19
            Top             =   255
            Width           =   3100
         End
         Begin VB.CheckBox chkRunRecentImmediate 
            Caption         =   "Alwa&ys display selection screen for Favourites and Recent"
            Height          =   330
            Left            =   200
            TabIndex        =   20
            Top             =   540
            Width           =   6360
         End
      End
      Begin VB.Frame fraReports 
         Caption         =   "Report / Utility / Tool Selection && Access :"
         Height          =   3500
         Index           =   0
         Left            =   -74880
         TabIndex        =   22
         Top             =   400
         Width           =   6735
         Begin SSDataWidgets_B.SSDBGrid grdUtilityReport 
            Height          =   2950
            Left            =   210
            TabIndex        =   23
            Top             =   315
            Width           =   6330
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            HeadLines       =   2
            Col.Count       =   3
            stylesets.count =   4
            stylesets(0).Name=   "ssetEnabled"
            stylesets(0).ForeColor=   -2147483640
            stylesets(0).BackColor=   -2147483643
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
            stylesets(0).Picture=   "frmConfiguration.frx":00D0
            stylesets(1).Name=   "ssetDisabled"
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
            stylesets(1).Picture=   "frmConfiguration.frx":00EC
            stylesets(2).Name=   "ActiveCheckbox"
            stylesets(2).ForeColor=   -2147483640
            stylesets(2).BackColor=   -2147483635
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
            stylesets(2).Picture=   "frmConfiguration.frx":0108
            stylesets(3).Name=   "ActiveText"
            stylesets(3).ForeColor=   -2147483634
            stylesets(3).BackColor=   -2147483635
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
            stylesets(3).Picture=   "frmConfiguration.frx":0124
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   0
            StyleSet        =   "ssetEnabled"
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   5398
            Columns(0).Caption=   "Report / Utility / Tool"
            Columns(0).Name =   "ReportUtility"
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   2646
            Columns(1).Caption=   "Only show own definitions"
            Columns(1).Name =   "Selection"
            Columns(1).Alignment=   2
            Columns(1).AllowSizing=   0   'False
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Style=   2
            Columns(2).Width=   2646
            Columns(2).Caption=   "Default Access"
            Columns(2).Name =   "DefaultAccess"
            Columns(2).AllowSizing=   0   'False
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(2).Locked=   -1  'True
            Columns(2).Style=   3
            Columns(2).Row.Count=   3
            Columns(2).Col.Count=   2
            Columns(2).Row(0).Col(0)=   "Read / Write"
            Columns(2).Row(1).Col(0)=   "Read Only"
            Columns(2).Row(2).Col(0)=   "Hidden"
            _ExtentX        =   11165
            _ExtentY        =   5203
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
      Begin VB.Frame Frame1 
         Caption         =   "Excel Options :"
         Height          =   1750
         Left            =   -74880
         TabIndex        =   105
         Top             =   4920
         Width           =   6735
         Begin VB.CheckBox chkOmitTopRow 
            Caption         =   "Omi&t Empty Row"
            Height          =   255
            Left            =   4320
            TabIndex        =   114
            Top             =   1080
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CheckBox chkOmitSpacerCol 
            Caption         =   "O&mit Empty Column"
            Height          =   255
            Left            =   4320
            TabIndex        =   115
            Top             =   1380
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox cboExcelFormat 
            Height          =   315
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   700
            Width           =   4800
         End
         Begin VB.CommandButton cmdFileClear 
            Caption         =   "O"
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   6180
            MaskColor       =   &H000000FF&
            TabIndex        =   109
            ToolTipText     =   "Clear Path"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   5850
            TabIndex        =   108
            ToolTipText     =   "Select Path"
            Top             =   300
            Width           =   330
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1725
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   300
            Width           =   4125
         End
         Begin VB.CheckBox chkExcelHeaders 
            Caption         =   "Row && Column &Headings"
            Height          =   255
            Left            =   1725
            TabIndex        =   112
            Top             =   1080
            Width           =   2400
         End
         Begin VB.CheckBox chkExcelGridlines 
            Caption         =   "Gr&idlines"
            Height          =   255
            Left            =   1725
            TabIndex        =   113
            Top             =   1380
            Width           =   1100
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Format :"
            Height          =   195
            Left            =   195
            TabIndex        =   110
            Top             =   765
            Width           =   1410
         End
         Begin VB.Label lblExcel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Template :"
            Height          =   195
            Left            =   195
            TabIndex        =   106
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame frmOutputs 
         Caption         =   "Output :"
         Height          =   945
         Left            =   -74880
         TabIndex        =   129
         Top             =   2060
         Width           =   6735
         Begin VB.CommandButton cmdDocumentsPath 
            Caption         =   "..."
            Height          =   315
            Left            =   5900
            TabIndex        =   33
            ToolTipText     =   "Select Path"
            Top             =   390
            Width           =   330
         End
         Begin VB.CommandButton cmdDocumentsPathClear 
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
            Left            =   6200
            MaskColor       =   &H000000FF&
            TabIndex        =   34
            ToolTipText     =   "Clear Path"
            Top             =   390
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtDocumentsPath 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   390
            Width           =   3700
         End
         Begin VB.Label lblDocumentsPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Document default output path :"
            Height          =   480
            Left            =   135
            TabIndex        =   130
            Top             =   330
            Width           =   1590
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraToolbarGeneral 
         Caption         =   "Layout :"
         Height          =   680
         Left            =   -74880
         TabIndex        =   121
         Top             =   6000
         Width           =   6735
         Begin VB.ComboBox cboToolbarPosition 
            Height          =   315
            Left            =   1450
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   240
            Width           =   5070
         End
         Begin VB.Label lblToolbarPosition 
            Caption         =   "Position :"
            Height          =   270
            Left            =   200
            TabIndex        =   122
            Top             =   300
            Width           =   800
         End
      End
      Begin VB.Frame fraToolbars 
         Caption         =   "Buttons :"
         Height          =   5565
         Left            =   -74880
         TabIndex        =   116
         Top             =   400
         Width           =   6735
         Begin VB.ComboBox cboToolbars 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":0140
            Left            =   1450
            List            =   "frmConfiguration.frx":0142
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   315
            Width           =   5070
         End
         Begin VB.CommandButton cmdShowHide 
            Caption         =   "S&how"
            Height          =   400
            Left            =   5300
            TabIndex        =   118
            Top             =   800
            Width           =   1200
         End
         Begin ComctlLib.ListView lvwToolbars 
            Height          =   4575
            Index           =   0
            Left            =   195
            TabIndex        =   117
            Top             =   795
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   8070
            View            =   2
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            OLEDragMode     =   1
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
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
            OLEDragMode     =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Name"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Enable"
               Object.Width           =   794
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Description"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "toolid"
               Object.Width           =   0
            EndProperty
         End
         Begin XtremeSuiteControls.PushButton cmdMoveDown 
            Height          =   405
            Left            =   5300
            TabIndex        =   120
            Top             =   2000
            Width           =   1200
            _Version        =   851969
            _ExtentX        =   2117
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Do&wn"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdMoveUp 
            Height          =   405
            Left            =   5300
            TabIndex        =   119
            Top             =   1500
            Width           =   1200
            _Version        =   851969
            _ExtentX        =   2117
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "&Up"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label lblToolbar 
            Caption         =   "Toolbar :"
            Height          =   195
            Left            =   200
            TabIndex        =   132
            Top             =   375
            Width           =   800
         End
      End
      Begin VB.Frame FraOutput 
         Caption         =   "Colours && Fonts :"
         Height          =   3225
         Index           =   1
         Left            =   -74880
         TabIndex        =   79
         Top             =   400
         Width           =   6735
         Begin VB.CheckBox chkUnderLine 
            Caption         =   "&Underline"
            Height          =   195
            Left            =   4080
            TabIndex        =   87
            Top             =   540
            Width           =   1695
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bo&ld"
            Height          =   195
            Left            =   4080
            TabIndex        =   86
            Top             =   285
            Width           =   1005
         End
         Begin VB.CheckBox chkGridlines 
            Caption         =   "&Gridlines"
            Height          =   195
            Left            =   4080
            TabIndex        =   88
            Top             =   795
            Width           =   1095
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000005&
            Height          =   1500
            Left            =   1725
            ScaleHeight     =   1440
            ScaleWidth      =   4740
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1620
            Width           =   4800
            Begin VB.Label lblHeading 
               Appearance      =   0  'Flat
               BackColor       =   &H00800080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Heading 2"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   1
               Left            =   2260
               TabIndex        =   94
               Top             =   615
               Width           =   1800
            End
            Begin VB.Label lblData 
               Appearance      =   0  'Flat
               BackColor       =   &H0000CCFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   3
               Left            =   2260
               TabIndex        =   97
               Top             =   1185
               Width           =   1800
            End
            Begin VB.Label lblData 
               Appearance      =   0  'Flat
               BackColor       =   &H0000CCFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Data 2"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   2260
               TabIndex        =   96
               Top             =   900
               Width           =   1800
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Title"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1440
               TabIndex        =   92
               Top             =   120
               Width           =   570
            End
            Begin VB.Label lblData 
               Appearance      =   0  'Flat
               BackColor       =   &H0000CCFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   2
               Left            =   480
               TabIndex        =   134
               Top             =   1185
               Width           =   1800
            End
            Begin VB.Label lblHeading 
               Appearance      =   0  'Flat
               BackColor       =   &H00800080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Heading 1"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   0
               Left            =   480
               TabIndex        =   93
               Top             =   615
               Width           =   1800
            End
            Begin VB.Label lblData 
               Appearance      =   0  'Flat
               BackColor       =   &H0000CCFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Data 1"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   480
               TabIndex        =   95
               Top             =   900
               Width           =   1800
            End
         End
         Begin VB.ComboBox cboTextType 
            Height          =   315
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   300
            Width           =   2200
         End
         Begin MSComctlLib.ImageCombo cboForeColour 
            Height          =   330
            Left            =   1725
            TabIndex        =   83
            Top             =   700
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin MSComctlLib.ImageCombo cboBackColour 
            Height          =   330
            Left            =   1725
            TabIndex        =   85
            Top             =   1100
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   6000
            Top             =   240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            UseMaskColor    =   0   'False
            _Version        =   393216
         End
         Begin VB.PictureBox picColour 
            Height          =   255
            Left            =   6360
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   89
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview :"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   90
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label lblBackColour 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Backcolour :"
            Height          =   195
            Left            =   195
            TabIndex        =   84
            Top             =   1160
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forecolour :"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   82
            Top             =   760
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text :"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   80
            Top             =   360
            Width           =   510
         End
      End
      Begin VB.Frame FraEventLog 
         Caption         =   "Data Transfer :"
         Height          =   1650
         Index           =   2
         Left            =   -74880
         TabIndex        =   77
         Top             =   5020
         Width           =   6735
         Begin VB.CheckBox chkDataTransferSuccess 
            Caption         =   "Successful records from Data &Transfers"
            Height          =   240
            Left            =   195
            TabIndex        =   78
            Top             =   315
            Width           =   3700
         End
      End
      Begin VB.Frame FraEventLog 
         Caption         =   "Imports / Exports :"
         Height          =   1920
         Index           =   1
         Left            =   -74880
         TabIndex        =   74
         Top             =   3065
         Width           =   6735
         Begin VB.CheckBox chkImportSuccess 
            Caption         =   "Successful records from &Imports"
            Height          =   195
            Left            =   195
            TabIndex        =   75
            Top             =   315
            Width           =   3100
         End
         Begin VB.CheckBox chkExportSuccess 
            Caption         =   "Successful records from E&xports"
            Height          =   195
            Left            =   195
            TabIndex        =   76
            Top             =   645
            Width           =   3100
         End
      End
      Begin VB.Frame FraEventLog 
         Caption         =   "Globals :"
         Height          =   2580
         Index           =   0
         Left            =   -74880
         TabIndex        =   70
         Top             =   400
         Width           =   6735
         Begin VB.CheckBox chkGlobalUpdateSuccess 
            Caption         =   "Successful records from Global &Updates"
            Height          =   195
            Left            =   195
            TabIndex        =   72
            Top             =   645
            Width           =   3800
         End
         Begin VB.CheckBox chkGlobalAddSuccess 
            Caption         =   "Successful records from Global &Adds"
            Height          =   195
            Left            =   195
            TabIndex        =   71
            Top             =   315
            Width           =   3500
         End
         Begin VB.CheckBox chkGlobalDeleteSuccess 
            Caption         =   "Successful records from Global D&eletes"
            Height          =   195
            Left            =   195
            TabIndex        =   73
            Top             =   960
            Width           =   3700
         End
      End
      Begin VB.Frame fraNetwork 
         Caption         =   "OLE Locations :"
         Height          =   1575
         Index           =   1
         Left            =   -74880
         TabIndex        =   48
         Top             =   3070
         Width           =   6735
         Begin VB.CommandButton cmdLocalOLEPath 
            Caption         =   "..."
            Height          =   315
            Left            =   5900
            TabIndex        =   39
            ToolTipText     =   "Select Path"
            Top             =   705
            Width           =   330
         End
         Begin VB.CommandButton cmdPhotoPath 
            Caption         =   "..."
            Height          =   315
            Left            =   5900
            TabIndex        =   42
            ToolTipText     =   "Select Path"
            Top             =   1110
            Width           =   330
         End
         Begin VB.CommandButton cmdOLEPath 
            Caption         =   "..."
            Height          =   315
            Left            =   5900
            TabIndex        =   36
            ToolTipText     =   "Select Path"
            Top             =   300
            Width           =   330
         End
         Begin VB.CommandButton cmdCrystalPath 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   5900
            TabIndex        =   45
            Top             =   1515
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton cmdCrystalPathClear 
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
            Left            =   6200
            MaskColor       =   &H000000FF&
            TabIndex        =   46
            Top             =   1515
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton cmdPhotoPathClear 
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
            Left            =   6200
            MaskColor       =   &H000000FF&
            TabIndex        =   43
            ToolTipText     =   "Clear Path"
            Top             =   1110
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdLocalOLEPathClear 
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
            Left            =   6200
            MaskColor       =   &H000000FF&
            TabIndex        =   40
            ToolTipText     =   "Clear Path"
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdOLEPathClear 
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
            Left            =   6200
            MaskColor       =   &H000000FF&
            TabIndex        =   37
            ToolTipText     =   "Clear Path"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtCrystalPath 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1515
            Visible         =   0   'False
            Width           =   3700
         End
         Begin VB.TextBox txtOLEPath 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   300
            Width           =   3700
         End
         Begin VB.TextBox txtPhotoPath 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1110
            Width           =   3700
         End
         Begin VB.TextBox txtLocalOLEPath 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   705
            Width           =   3700
         End
         Begin VB.Label lblCrystalPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crystal Path :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   135
            TabIndex        =   52
            Top             =   1560
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblOLEPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server path :"
            Height          =   195
            Left            =   135
            TabIndex        =   50
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblPhotoPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Photograph path (non-linked) :"
            Height          =   390
            Left            =   135
            TabIndex        =   51
            Top             =   1080
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local path :"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   49
            Top             =   765
            Width           =   840
         End
      End
      Begin VB.Frame fraNetwork 
         Caption         =   "Printer :"
         Height          =   1600
         Index           =   0
         Left            =   -74880
         TabIndex        =   27
         Top             =   400
         Width           =   6735
         Begin VB.CheckBox chkPrintingPrompt 
            Caption         =   "Displa&y print options before printing (excluding Word and Excel output)"
            Height          =   420
            Left            =   135
            TabIndex        =   30
            Top             =   700
            Width           =   6400
         End
         Begin VB.CheckBox chkPrintingConfirm 
            Caption         =   "Confirm &after print job has been completed (excluding Output options)"
            Height          =   420
            Left            =   135
            TabIndex        =   31
            Top             =   1065
            Width           =   6400
         End
         Begin VB.ComboBox cboPrinter 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   315
            Width           =   4300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current default :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   28
            Top             =   375
            Width           =   1440
         End
      End
      Begin VB.Frame fraBatch 
         Caption         =   "Login Details :"
         Height          =   3160
         Index           =   0
         Left            =   -74880
         TabIndex        =   54
         Top             =   400
         Width           =   6735
         Begin VB.CheckBox chkUseWindowsAuthentication 
            Caption         =   "&Use Windows Authentication"
            Enabled         =   0   'False
            Height          =   210
            Left            =   480
            TabIndex        =   59
            Top             =   1440
            Width           =   2700
         End
         Begin VB.CommandButton cmdTestLogon 
            Caption         =   "&Test Login"
            Enabled         =   0   'False
            Height          =   400
            Left            =   5300
            TabIndex        =   64
            Top             =   2595
            Width           =   1200
         End
         Begin VB.TextBox txtServer 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            MaxLength       =   128
            TabIndex        =   63
            Top             =   2175
            Width           =   4300
         End
         Begin VB.TextBox txtDatabase 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            MaxLength       =   128
            TabIndex        =   61
            Top             =   1770
            Width           =   4300
         End
         Begin VB.TextBox txtPWD 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2200
            MaxLength       =   128
            PasswordChar    =   "*"
            TabIndex        =   58
            Top             =   1005
            Width           =   4300
         End
         Begin VB.TextBox txtUID 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            MaxLength       =   128
            TabIndex        =   56
            Top             =   600
            Width           =   4300
         End
         Begin VB.CheckBox chkBatchLogon 
            Caption         =   "&Enable Batch Login on this computer"
            Height          =   195
            Left            =   195
            TabIndex        =   53
            Top             =   315
            Width           =   3500
         End
         Begin VB.Label lblServer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   480
            TabIndex        =   62
            Top             =   2235
            Width           =   1125
         End
         Begin VB.Label lblDatabase 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   480
            TabIndex        =   60
            Top             =   1830
            Width           =   1335
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   480
            TabIndex        =   57
            Top             =   1065
            Width           =   1335
         End
         Begin VB.Label lblUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   480
            TabIndex        =   55
            Top             =   660
            Width           =   975
         End
      End
      Begin VB.Frame fraBatch 
         Caption         =   "Email Errors :"
         Height          =   1695
         Index           =   1
         Left            =   -74880
         TabIndex        =   65
         Top             =   3600
         Width           =   6735
         Begin VB.CommandButton cmdBatchEmail 
            Caption         =   "Te&st Email"
            Enabled         =   0   'False
            Height          =   400
            Left            =   5300
            TabIndex        =   69
            Top             =   1020
            Width           =   1200
         End
         Begin VB.CheckBox chkBatchEmail 
            Caption         =   "Send an e&mail to the administrator if the batch login fails"
            Height          =   195
            Left            =   195
            TabIndex        =   66
            Top             =   315
            Width           =   5250
         End
         Begin VB.TextBox txtBatchEmailAddr 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2200
            MaxLength       =   128
            TabIndex        =   68
            Top             =   600
            Width           =   4300
         End
         Begin VB.Label lblBatchEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   480
            TabIndex        =   67
            Top             =   660
            Width           =   1455
         End
      End
      Begin VB.Frame fraReports 
         Caption         =   "Warning Message :"
         Height          =   1950
         Index           =   1
         Left            =   -74880
         TabIndex        =   24
         Top             =   3960
         Width           =   6735
         Begin VB.ListBox lstWarningMsg 
            Height          =   1185
            ItemData        =   "frmConfiguration.frx":0144
            Left            =   195
            List            =   "frmConfiguration.frx":0146
            Style           =   1  'Checkbox
            TabIndex        =   26
            Top             =   600
            Width           =   6325
         End
         Begin VB.Label lblWarningMsg 
            Caption         =   "Only show warnings for the following utilities :"
            Height          =   195
            Left            =   195
            TabIndex        =   25
            Top             =   320
            Width           =   5175
         End
      End
      Begin VB.Frame fraDisplay 
         Caption         =   "Diary Options :"
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   4020
         Width           =   6735
         Begin VB.ComboBox cboDiaryView 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":0148
            Left            =   2200
            List            =   "frmConfiguration.frx":014A
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   300
            Width           =   4300
         End
         Begin VB.CheckBox chkDiaryConstantCheck 
            Caption         =   "Display alar&med events throughout the day"
            Height          =   255
            Left            =   195
            TabIndex        =   18
            Top             =   780
            Width           =   4100
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default View :"
            Height          =   195
            Index           =   6
            Left            =   195
            TabIndex        =   16
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame fraDisplay 
         Caption         =   "Record Editing :"
         Height          =   2380
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   400
         Width           =   6735
         Begin VB.CheckBox chkEmailRecDesc 
            Caption         =   "&Include record description when sending an email from within a record"
            Height          =   255
            Left            =   195
            TabIndex        =   9
            Top             =   1960
            Width           =   6400
         End
         Begin VB.ComboBox cboQuickAccess 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":014C
            Left            =   2200
            List            =   "frmConfiguration.frx":014E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1500
            Width           =   4300
         End
         Begin VB.ComboBox cboLookUp 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":0150
            Left            =   2200
            List            =   "frmConfiguration.frx":0152
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1100
            Width           =   4300
         End
         Begin VB.ComboBox cboHistory 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":0154
            Left            =   2200
            List            =   "frmConfiguration.frx":0156
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   700
            Width           =   4300
         End
         Begin VB.ComboBox cboPrimary 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":0158
            Left            =   2200
            List            =   "frmConfiguration.frx":015A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   315
            Width           =   4300
         End
         Begin VB.Label lblQuickAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quick Access :"
            Height          =   195
            Left            =   195
            TabIndex        =   7
            Top             =   1560
            Width           =   1530
         End
         Begin VB.Label lblLookup 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lookup Tables :"
            Height          =   195
            Left            =   195
            TabIndex        =   5
            Top             =   1155
            Width           =   1620
         End
         Begin VB.Label lblHistory 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Child Tables :"
            Height          =   195
            Left            =   195
            TabIndex        =   3
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label lblPrimary 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parent Tables :"
            Height          =   195
            Left            =   195
            TabIndex        =   1
            Top             =   375
            Width           =   1590
         End
      End
      Begin VB.Frame fraDisplay 
         Caption         =   "Filters / Calculations :"
         Height          =   1155
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   2820
         Width           =   6735
         Begin VB.ComboBox cboNodeSize 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":015C
            Left            =   2200
            List            =   "frmConfiguration.frx":015E
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   700
            Width           =   4300
         End
         Begin VB.ComboBox cboColours 
            Height          =   315
            ItemData        =   "frmConfiguration.frx":0160
            Left            =   2200
            List            =   "frmConfiguration.frx":0162
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   300
            Width           =   4300
         End
         Begin VB.Label lblExpandNodes 
            AutoSize        =   -1  'True
            Caption         =   "Expand Nodes :"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   765
            Width           =   1680
         End
         Begin VB.Label lblViewInColour 
            AutoSize        =   -1  'True
            Caption         =   "View In Colour :"
            Height          =   195
            Left            =   195
            TabIndex        =   11
            Top             =   360
            Width           =   1680
         End
      End
      Begin VB.Frame FraOutput 
         Caption         =   "Word Options :"
         Height          =   1200
         Index           =   0
         Left            =   -74880
         TabIndex        =   98
         Top             =   3680
         Width           =   6735
         Begin VB.ComboBox cboWordFormat 
            Height          =   315
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   700
            Width           =   4800
         End
         Begin VB.CommandButton cmdFileClear 
            Caption         =   "O"
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   6180
            MaskColor       =   &H000000FF&
            TabIndex        =   102
            ToolTipText     =   "Clear Path"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1725
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   300
            Width           =   4125
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   5850
            TabIndex        =   101
            ToolTipText     =   "Select Path"
            Top             =   300
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Format :"
            Height          =   195
            Left            =   195
            TabIndex        =   103
            Top             =   765
            Width           =   1410
         End
         Begin VB.Label lblWord 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Template :"
            Height          =   195
            Left            =   195
            TabIndex        =   99
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame frmAutoLogin 
         Caption         =   "Login :"
         Height          =   780
         Left            =   -74880
         TabIndex        =   128
         Top             =   4710
         Width           =   6735
         Begin VB.CheckBox chkBypassLogonDetails 
            Caption         =   "Bypass prompt for lo&gon details"
            Height          =   285
            Left            =   150
            TabIndex        =   47
            Top             =   285
            Width           =   3100
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5875
      TabIndex        =   126
      Top             =   7290
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   4550
      TabIndex        =   125
      Top             =   7290
      Width           =   1200
   End
   Begin ComctlLib.ImageList ilstToolbar 
      Index           =   0
      Left            =   2640
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483638
      MaskColor       =   -2147483638
      _Version        =   327682
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrRECORDEDITBAND = "bndRecord"
Private Const mstrFINDWINDOWBAND = "bndHistory"

'Dim Colors() As String
'Dim colColours As Collection

Private moutTitle As clsOutputStyle
Private moutHeading As clsOutputStyle
Private moutData As clsOutputStyle
Private moutCurrent As clsOutputStyle

Private mblnUserSettings As Boolean

Private mbLoading As Boolean

Private msPhotoPath As String
Private msOLEPath As String
Private msCrystalPath As String
Private msDocumentsPath As String
Private msLocalOlePath As String

Private mcPrimary As DefaultDisplay
Private mcHistory As DefaultDisplay
Private mcLookUp As DefaultDisplay
Private mcQuickAccess As DefaultDisplay

Private mstrDefs() As String
Private mstrWarning() As String
Private mbCloseDefSelAfterRun As Boolean
Private mbRecentDisplayDefSel As Boolean
Private mbRememberDefSelID As Boolean

Private mcExpressionColours As ExpressionColour
Private mcExpressionNodeSize As ExpressionSaveView

'Private mlngDiaryFilterEventType As Long
'Private mlngDiaryFilterAlarmStatus As Long
'Private mlngDiaryFilterPastPresent As Long
'Private mblnDiaryFilterOnlyMine As Boolean
Private mlngDiaryDefaultView As Long

Private mstrDefaultPrinter As String
Private mlngToolbarPosition As Long

Private mlngWordFormat As Long
Private mlngExcelFormat As Long

Private mblnRestoringDefaults As Boolean
Dim OpenedMe As Form

Private Sub cboColours_Click()

  If mcExpressionColours <> cboColours.ItemData(cboColours.ListIndex) Then
    mcExpressionColours = cboColours.ItemData(cboColours.ListIndex)
    Changed = Not mbLoading
  End If
    
End Sub
Public Property Let CallingForm(frm As Form)
Set OpenedMe = frm    'you can now idenify the calling Form and all its properties
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  If Not mbLoading Then
    cmdOK.Enabled = pblnChanged
  End If
End Property


Private Sub cboDiaryView_Click()
  If mlngDiaryDefaultView <> cboDiaryView.ItemData(cboDiaryView.ListIndex) Then
    mlngDiaryDefaultView = cboDiaryView.ItemData(cboDiaryView.ListIndex)
    Changed = Not mbLoading
  End If
End Sub


Private Sub cboHistory_Click()

  If mcHistory <> cboHistory.ItemData(cboHistory.ListIndex) Then
    mcHistory = cboHistory.ItemData(cboHistory.ListIndex)
    Changed = Not mbLoading
  End If

End Sub

Private Sub cboLookUp_Click()

  If mcLookUp <> cboLookUp.ItemData(cboLookUp.ListIndex) Then
    mcLookUp = cboLookUp.ItemData(cboLookUp.ListIndex)
    Changed = Not mbLoading
  End If

End Sub

Private Sub cboNodeSize_Click()

  If mcExpressionNodeSize <> cboNodeSize.ItemData(cboNodeSize.ListIndex) Then
    mcExpressionNodeSize = cboNodeSize.ItemData(cboNodeSize.ListIndex)
    Changed = Not mbLoading
  End If

End Sub


Private Sub cboPrimary_Click()

  If mcPrimary <> cboPrimary.ItemData(cboPrimary.ListIndex) Then
    mcPrimary = cboPrimary.ItemData(cboPrimary.ListIndex)
    Changed = Not mbLoading
  End If
  
End Sub

Private Sub cboPrinter_Click()
 
  If mstrDefaultPrinter <> cboPrinter.Text Then
    mstrDefaultPrinter = cboPrinter.ItemData(cboPrinter.ListIndex)
    Changed = Not mbLoading
  End If
    
End Sub

Private Sub cboQuickAccess_Click()
 
  If mcQuickAccess <> cboQuickAccess.ItemData(cboQuickAccess.ListIndex) Then
    mcQuickAccess = cboQuickAccess.ItemData(cboQuickAccess.ListIndex)
    Changed = Not mbLoading
  End If
  
End Sub

Private Sub cboToolbarPosition_Click()

  If mlngToolbarPosition <> cboToolbarPosition.ItemData(cboToolbarPosition.ListIndex) Then
    mlngToolbarPosition = cboToolbarPosition.ItemData(cboToolbarPosition.ListIndex)
    Changed = Not mbLoading
  End If

End Sub

Private Sub cboToolbars_Change()
  fraToolbars.Refresh
End Sub

Private Sub cboToolbars_Click()
  
  LoadToolBarOptions False
  lvwToolbars(CurrentToolBarIndex).Refresh

End Sub

Private Sub cboWordFormat_Click()
  Changed = Not mbLoading
End Sub

Private Sub cboExcelFormat_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkBatchEmail_Click()

  Dim blnBatchEmail As Boolean
  
  blnBatchEmail = (chkBatchEmail.Value = vbChecked)
  
  lblBatchEmail.Enabled = blnBatchEmail
  txtBatchEmailAddr.Enabled = blnBatchEmail
  txtBatchEmailAddr.BackColor = IIf(blnBatchEmail, vbWindowBackground, vbButtonFace)
  
  If Not blnBatchEmail Then
    txtBatchEmailAddr.Text = vbNullString
  End If

  Changed = Not mbLoading
  
End Sub

Private Sub chkBatchLogon_Click()

  Dim blnBatchLogon As Boolean

  blnBatchLogon = (chkBatchLogon.Value = vbChecked)

  lblUser.Enabled = blnBatchLogon
  lblPassword.Enabled = blnBatchLogon
  lblDatabase.Enabled = blnBatchLogon
  lblServer.Enabled = blnBatchLogon
  
  txtUID.Enabled = blnBatchLogon
  txtPWD.Enabled = blnBatchLogon
  chkUseWindowsAuthentication.Enabled = blnBatchLogon And glngSQLVersion > 7
  If Not blnBatchLogon Then
    chkUseWindowsAuthentication.Value = vbUnchecked
  End If
  txtDatabase.Enabled = blnBatchLogon
  txtServer.Enabled = blnBatchLogon
  
  txtUID.BackColor = IIf(blnBatchLogon, vbWindowBackground, vbButtonFace)
  txtPWD.BackColor = IIf(blnBatchLogon, vbWindowBackground, vbButtonFace)
  txtDatabase.BackColor = IIf(blnBatchLogon, vbWindowBackground, vbButtonFace)
  txtServer.BackColor = IIf(blnBatchLogon, vbWindowBackground, vbButtonFace)

  chkBatchEmail.Enabled = blnBatchLogon
  cmdTestLogon.Enabled = blnBatchLogon
  
  If blnBatchLogon Then
    txtUID.Text = gsUserName
    txtPWD.Text = vbNullString
    txtDatabase.Text = gsDatabaseName
    txtServer.Text = GetPCSetting("Login", "DataMgr_Server", vbNullString)
  Else
    txtUID.Text = vbNullString
    txtPWD.Text = vbNullString
    txtDatabase.Text = vbNullString
    txtServer.Text = vbNullString
    
    chkBatchEmail.Value = vbUnchecked
    txtBatchEmailAddr.Text = vbNullString
  End If

  Changed = Not mbLoading

End Sub


Private Sub chkBypassLogonDetails_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkCloseDefsel_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkDataTransferSuccess_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkDiaryConstantCheck_Click()
  Changed = Not mbLoading
End Sub

'MH20041104
Private Sub chkEmailRecDesc_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkExcelGridlines_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkExcelHeaders_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkExportSuccess_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkGlobalAddSuccess_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkGlobalDeleteSuccess_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkGlobalUpdateSuccess_Click()
  Changed = Not mbLoading
End Sub
Private Sub chkImportSuccess_Click()

  Changed = Not mbLoading

End Sub

Private Sub chkOmitSpacerCol_Click()
Changed = Not mbLoading
End Sub

Private Sub chkOmitTopRow_Click()
Changed = Not mbLoading
End Sub

Private Sub chkPrintingConfirm_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkPrintingPrompt_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkRememberDefSelID_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkRunRecentImmediate_Click()
  Changed = Not mbLoading
End Sub

Private Sub chkUseWindowsAuthentication_Click()

  txtUID.Text = ""
  txtUID.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)
  
  txtPWD.Text = ""
  txtPWD.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)

  ' Grey out controls
  txtUID.BackColor = IIf(txtUID.Enabled, vbWindowBackground, vbButtonFace)
  txtPWD.BackColor = IIf(txtPWD.Enabled, vbWindowBackground, vbButtonFace)
  cmdTestLogon.Enabled = IIf(chkUseWindowsAuthentication.Value = vbChecked, False, True)

  Changed = Not mbLoading

End Sub

Private Sub cmdBatchEmail_Click()
        
  Dim strMBText As String
  
  strMBText = "Are you sure that you would like to send a test message to '" & _
              txtBatchEmailAddr.Text & "' ?"
  
  If COAMsgBox(strMBText, vbYesNo + vbQuestion, "Test Message") = vbYes Then
    Screen.MousePointer = vbHourglass
    If frmEmailSel.SendEmail _
      (txtBatchEmailAddr.Text, "Test Message", _
          "This is a test message from OpenHR Batch Login configuration.  Please ignore.", True) Then
      Screen.MousePointer = vbDefault
      COAMsgBox "Message Sent.", vbInformation
    End If
    
    Unload frmEmailSel
    Set frmEmailSel = Nothing
    Screen.MousePointer = vbDefault
  End If

End Sub


Private Sub cmdCancel_Click()

  ' Restore the path settings and unload the form.
'  If Not mblnUserSettings Then
'    RestoreOriginalValues
'  End If
  Unload Me

End Sub

Private Sub cmdDefault_Click()

  If COAMsgBox("Are you sure you want to restore all default settings?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
''    'Set all the defaults for the currently selected flag.
'''    SetUserSettingDefaults
''    'Set the selected item of each of the drop-down boxes and check boxes.
'''    ReadUserSettings True
''
''    'TM20020111 Fault 2772
''    ReadUserSettingDefaults

    'gADOCon.Execute "DELETE FROM ASRSYSUserSettings WHERE username = System_User"
    cboWordFormat.ListIndex = -1
    cboExcelFormat.ListIndex = -1
    ReadUserSettings True
    Changed = True

    cboTextType_Click
   
    ' Clear recent menu history
    gADOCon.Execute "EXEC dbo.[spstat_clearrecentusage]"
   
    ' Clear favourites
    gADOCon.Execute "EXEC dbo.[spstat_clearfavourites]"

  End If

End Sub

Private Sub cmdDocumentsPathClear_Click()
  
  Me.txtDocumentsPath.Text = vbNullString
  cmdDocumentsPathClear.Enabled = False
  
  Changed = Not mbLoading

End Sub


Private Sub cmdFileClear_Click(Index As Integer)

  Dim strMBText As String

  strMBText = "Are you sure that you would like to clear the " & _
              IIf(Index = 0, "Excel", "Word") & " Template?"
  
  If COAMsgBox(strMBText, vbQuestion + vbYesNoCancel, Me.Caption) = vbYes Then
    txtFilename(Index).Text = vbNullString
    Changed = True
  End If

End Sub

Private Sub cmdFileName_Click(Index As Integer)
  
  On Local Error GoTo LocalErr

  With frmMain.CommonDialog1
    If Len(Trim(txtFilename(Index).Text)) = 0 Or txtFilename(Index).Text = "<None>" Then
      .InitDir = gsDocumentsPath
    Else
      .FileName = txtFilename(Index).Text
    End If

    .CancelError = True
    Select Case Index
    Case 0
      .DialogTitle = "Excel Template"
      '.Filter = gsOfficeTemplateFilter_Excel
      '.Filter = GetCommonDialogFormats("ExcelTemplate", GetOfficeExcelVersion)
      'InitialiseCommonDialogFormats frmMain.CommonDialog1, "ExcelTemplate", GetOfficeExcelVersion
      .Filter = "Excel Template (*.xlt;*.xltx;*.xls;*.xlsx)|*.xlt;*.xltx;*.xls;*.xlsx"
      .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
      .ShowOpen
    Case 1
      .DialogTitle = "Word Template"
      '.Filter = gsOfficeTemplateFilter_Word
      '.Filter = GetCommonDialogFormats("WordTemplate", GetOfficeWordVersion)
      .Filter = "Word Template (*.dot;*.dotx;*.doc;*.docx)|*.dot;*.dotx;*.doc;*.docx"
      .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
      .ShowOpen
    End Select

    If Len(.FileName) > 256 Then
      COAMsgBox "Path and file name must not exceed 256 characters in length", vbExclamation, Me.Caption
      Exit Sub
    End If

    If .FileName <> "" Then
      txtFilename(Index).Text = .FileName
      Changed = True
    End If
  
  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    COAMsgBox "Error selecting file", vbCritical, Me.Caption
    txtFilename(Index).Text = vbNullString
  End If

End Sub

Private Sub cmdMoveDown_LostFocus()
  cmdMoveDown.Picture = cmdMoveDown.Picture

End Sub

Private Sub cmdMoveUp_LostFocus()
  cmdMoveUp.Picture = cmdMoveUp.Picture
  
End Sub

Private Sub cmdShowHide_Click()

  Dim iCurrentListID As Integer
  iCurrentListID = CurrentToolBarIndex

  lvwToolbars(iCurrentListID).SelectedItem.SubItems(1) = IIf(lvwToolbars(iCurrentListID).SelectedItem.SubItems(1) = "Hidden", "", "Hidden")
  lvwToolbars(iCurrentListID).SetFocus

  UpdateToolbarButtonStatus
  
  Changed = Not mbLoading

End Sub

Private Sub cmdMoveDown_Click()

  ChangeSelectedToolOrder lvwToolbars(CurrentToolBarIndex), lvwToolbars(CurrentToolBarIndex).SelectedItem.Index + 2, True
  lvwToolbars(CurrentToolBarIndex).SetFocus

End Sub

Private Sub cmdMoveUp_Click()

  ChangeSelectedToolOrder lvwToolbars(CurrentToolBarIndex), lvwToolbars(CurrentToolBarIndex).SelectedItem.Index - 1
  lvwToolbars(CurrentToolBarIndex).SetFocus
  
End Sub

Private Sub cmdOLEPathClear_Click()
  
  Me.txtOLEPath.Text = vbNullString
  cmdOLEPathClear.Enabled = False

  Changed = Not mbLoading

End Sub

Private Sub cmdLocalOLEPathClear_Click()
  
  Me.txtLocalOLEPath.Text = vbNullString
  cmdLocalOLEPathClear.Enabled = False

  Changed = Not mbLoading

End Sub

Private Sub cmdPhotoPathClear_Click()
  
  Me.txtPhotoPath.Text = vbNullString
  cmdPhotoPathClear.Enabled = False

  Changed = Not mbLoading

End Sub

Private Sub cmdCrystalPathClear_Click()
  
  Me.txtCrystalPath.Text = vbNullString
  cmdCrystalPathClear.Enabled = False

  Changed = Not mbLoading

End Sub


Public Function Initialise(blnUserSettings As Boolean) As Boolean
  'JPD 20030915 Fault 6884
  Screen.MousePointer = vbHourglass
  ' NPG20100824 - As frmlogin can now call this form we don't want to always disable the menus...
  ' frmMain.DisableMenu
  If Not gcoTablePrivileges Is Nothing Then frmMain.DisableMenu
  

  mbLoading = True

  mblnUserSettings = blnUserSettings

  SSTab1.TabVisible(0) = mblnUserSettings
  SSTab1.TabVisible(1) = mblnUserSettings
  SSTab1.TabVisible(2) = Not mblnUserSettings
  SSTab1.TabVisible(3) = Not mblnUserSettings
  SSTab1.TabVisible(4) = mblnUserSettings
  SSTab1.TabVisible(5) = mblnUserSettings
  SSTab1.TabVisible(6) = mblnUserSettings
    
  SSTab1.Tab = IIf(mblnUserSettings, 0, 2)
  cmdDefault.Visible = mblnUserSettings
    
  If mblnUserSettings Then
    Me.Caption = "User Configuration"
    PopulateControlsUserSettings
    ReadUserSettings False
    Me.HelpContextID = 1113
  
  Else
    Me.Caption = "PC Configuration"
    ReadPCSettings
  
    ' JDM - 27/11/01 - Fault 3109 - Disable access to batch log on.
    SSTab1.TabVisible(3) = datGeneral.SystemPermission("CONFIGURATION", "PC")
 
    Me.HelpContextID = 1114
  End If

  Initialise = True

  Changed = False

  mbLoading = False

  'JPD 20030915 Fault 6884
  frmMain.EnableMenu Me
  Screen.MousePointer = vbDefault
   
End Function


Private Sub cmdOK_Click()
  
  On Error GoTo SaveError
  
  Screen.MousePointer = vbHourglass

  If mblnUserSettings Then
    SaveUserSettings
    Unload Me
  Else
    If SavePCSettings Then
      Unload Me
    End If
  End If

  Screen.MousePointer = vbDefault
  
  Exit Sub
  
SaveError:
  COAMsgBox "Error saving configuration settings." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, app.Title
  
End Sub


Private Sub cmdOLEPath_Click()
'  ' Display the path selection form.
'  Dim frmSelectPath As frmPathSel
'
'  Set frmSelectPath = New frmPathSel
'  With frmSelectPath
'    .QuietMode = True
'    .SelectionType = 2
'    .Show vbModal
'  End With
'  Set frmSelectPath = Nothing
'
'  If txtOLEPath.Text <> gsOLEPath Then
'    txtOLEPath.Text = gsOLEPath
'    Changed = Not mbLoading
'  End If
'
'  cmdOLEPathClear.Enabled = Not (gsOLEPath = "")
  
  Dim strFolder As String
      
  strFolder = BrowseFolders("You have opted to change the default storage folder.")
  
  If txtOLEPath.Text <> strFolder And strFolder <> vbNullString Then
    txtOLEPath.Text = strFolder
    Changed = Not mbLoading
  End If
    
  cmdOLEPathClear.Enabled = Not (txtOLEPath.Text = "")
    
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPhotoPath_Click()
'  ' Display the path selection form.
'  Dim frmSelectPath As frmPathSel
'
'  Set frmSelectPath = New frmPathSel
'  With frmSelectPath
'    .QuietMode = True
'    .SelectionType = 1
'    .Show vbModal
'  End With
'  Set frmSelectPath = Nothing
'
'  If txtPhotoPath.Text <> gsPhotoPath Then
'    txtPhotoPath.Text = gsPhotoPath
'    Changed = Not mbLoading
'  End If
'
'  cmdPhotoPathClear.Enabled = Not (gsPhotoPath = "")

  Dim strFolder As String
      
  strFolder = BrowseFolders("You have opted to change the default storage folder.")
  
  If txtPhotoPath.Text <> strFolder And strFolder <> vbNullString Then
    txtPhotoPath.Text = strFolder
    Changed = Not mbLoading
  End If

  cmdPhotoPathClear.Enabled = Not (txtPhotoPath.Text = "")
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdDocumentsPath_Click()

'  ' Display the path selection form.
'  Dim frmSelectPath As frmPathSel
'
'  Set frmSelectPath = New frmPathSel
'  With frmSelectPath
'    .QuietMode = True
'    .SelectionType = 8
'    .Show vbModal
'  End With
'  Set frmSelectPath = Nothing
'
'  If txtDocumentsPath.Text <> gsDocumentsPath Then
'    txtDocumentsPath.Text = gsDocumentsPath
'    Changed = Not mbLoading
'  End If
'
'  cmdDocumentsPathClear.Enabled = Not (gsDocumentsPath = "")
  
  Dim strFolder As String
  
  strFolder = BrowseFolders("You have opted to change the default storage folder.")
  
  If txtDocumentsPath.Text <> strFolder And strFolder <> vbNullString Then
    txtDocumentsPath.Text = strFolder
    Changed = Not mbLoading
  End If
    
  cmdDocumentsPathClear.Enabled = Not (txtDocumentsPath.Text = "")
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdCrystalPath_Click()
  ' Display the path selection form.
  Dim frmSelectPath As frmPathSel
    
  Set frmSelectPath = New frmPathSel
  With frmSelectPath
    .QuietMode = True
    .SelectionType = 4
    .Show vbModal
  End With
  Set frmSelectPath = Nothing
  
  If txtCrystalPath.Text <> gsCrystalPath Then
    txtCrystalPath.Text = gsCrystalPath
    Changed = Not mbLoading
  End If
  
  cmdCrystalPathClear.Enabled = Not (gsCrystalPath = "")

  Screen.MousePointer = vbDefault

End Sub


Private Sub cmdLocalOLEPath_Click()
'  ' Display the path selection form.
'  Dim frmSelectPath As frmPathSel
'
'  Set frmSelectPath = New frmPathSel
'  With frmSelectPath
'    .QuietMode = True
'    .SelectionType = 16
'    .Show vbModal
'  End With
'  Set frmSelectPath = Nothing
'
'  If txtLocalOLEPath.Text <> gsLocalOLEPath Then
'    txtLocalOLEPath.Text = gsLocalOLEPath
'    Changed = Not mbLoading
'  End If
'
'  cmdLocalOLEPathClear.Enabled = Not (gsLocalOLEPath = "")
  
  Dim strFolder As String
  
  strFolder = BrowseFolders("You have opted to change the default storage folder.")
  
  If txtLocalOLEPath.Text <> strFolder And strFolder <> vbNullString Then
    txtLocalOLEPath.Text = strFolder
    Changed = Not mbLoading
  End If
  
  cmdLocalOLEPathClear.Enabled = Not (txtLocalOLEPath.Text = "")

  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdTestLogon_Click()

  Dim objTestConn As ADODB.Connection
  Dim sConnect As String
  
  On Error GoTo LocalErr
  
  If Trim(txtUID.Text) = vbNullString Then
    COAMsgBox "You must enter a user name.", vbInformation, "Batch Login"
    Exit Sub
  End If
  
  If Trim(txtDatabase.Text) = vbNullString Then
    COAMsgBox "You must enter a Database name.", vbInformation, "Batch Login"
    Exit Sub
  End If
  
  If Trim(txtServer.Text) = vbNullString Then
    COAMsgBox "You must enter a server name.", vbInformation, "Batch Login"
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  sConnect = "Driver=SQL Server;" & _
             "Server=" & txtServer.Text & ";" & _
             "UID=" & txtUID.Text & ";" & _
             "PWD=" & txtPWD.Tag & ";" & _
             "Database=" & txtDatabase.Text & ";Pooling=false;App=Test OpenHR Batch;"

  Set objTestConn = New ADODB.Connection
  With objTestConn
    .ConnectionString = sConnect
    .Provider = "SQLOLEDB"
    .CommandTimeout = 10
    .ConnectionTimeout = 30   'MH20030911 Fault 6944
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Packet Size") = 32767
    .Open
    .Close
  End With
  
  Set objTestConn = Nothing
  
  Screen.MousePointer = vbDefault
  ' JPD20030211 Fault 5045
  COAMsgBox "Test completed successfully.", vbInformation, "Batch Login"

Exit Sub

LocalErr:
  Screen.MousePointer = vbDefault
  COAMsgBox "Error during batch login test." & vbCrLf & _
         ADOConError(objTestConn), vbInformation, "Batch Login"

End Sub


Private Sub DoPrinterTab()

  On Error GoTo InitERROR
  
  ' Retrieve the current default printer and set the combo
  Dim objPrinter As Printer
  Dim bFoundDefault As Boolean
  Dim strMsg As String
  
  If Printers.Count = 0 Then
    cboPrinter.Enabled = False
    cboPrinter.BackColor = vbButtonFace
    'TM20011114 Fault 3150 - disable printer check boxes if no printers detected.
    chkPrintingConfirm.Enabled = False
    chkPrintingConfirm.Value = vbUnchecked
    
    chkPrintingPrompt.Enabled = False
    chkPrintingPrompt.Value = vbUnchecked
    'Me.SSTab1.TabEnabled(2) = False
    Exit Sub
  End If
  
  'TM20020828 Fault 1432
  Printer.TrackDefault = True
  gstrDefaultPrinterName = Printer.DeviceName
  
  SavePCSetting "Printer", "DeviceName", gstrDefaultPrinterName
  
  For Each objPrinter In Printers
    cboPrinter.AddItem objPrinter.DeviceName
'    If LCase(objPrinter.DeviceName) = LCase(mstrDefaultPrinter) Then
'      bFoundDefault = True
'    End If
    
    If LCase(objPrinter.DeviceName) = LCase(gstrDefaultPrinterName) Then
      cboPrinter.ListIndex = cboPrinter.NewIndex
      bFoundDefault = True
    End If

  Next objPrinter
  
  If Not bFoundDefault Then
    strMsg = "Unable to find default printer !" & vbCrLf & vbCrLf
    strMsg = strMsg & "Installed Printers : " & vbCrLf
    For Each objPrinter In Printers
      strMsg = strMsg & vbCrLf & objPrinter.DeviceName
    Next objPrinter
    strMsg = strMsg & vbCrLf & vbCrLf & "Default Printer : " & vbCrLf & vbCrLf & Printer.DeviceName
    COAMsgBox strMsg, vbExclamation + vbOKOnly, app.Title
'  Else
'    cboPrinter.Text = mstrDefaultPrinter
  End If
  
  Exit Sub

InitERROR:

  COAMsgBox "Error whilst intialising default printer tab." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, app.Title
  
End Sub

Private Sub SetCombo(objControl As ComboBox, lItemData As Long)

  Dim lCount As Long
  
  For lCount = 0 To objControl.ListCount
    If objControl.ItemData(lCount) = lItemData Then
      objControl.ListIndex = lCount
      Exit Sub
    End If
  Next lCount

End Sub

Private Sub RestoreOriginalValues()
  
  ' Restore the original values.
  SavePCSetting "Datapaths", "photopath_" & gsDatabaseName, msPhotoPath
  gsPhotoPath = msPhotoPath
  
  SavePCSetting "Datapaths", "olepath_" & gsDatabaseName, msOLEPath
  gsOLEPath = msOLEPath
  
  SavePCSetting "Datapaths", "crystalpath_" & gsDatabaseName, msCrystalPath
  gsCrystalPath = msCrystalPath
  
  SavePCSetting "Datapaths", "documentspath_" & gsDatabaseName, msDocumentsPath
  gsDocumentsPath = msDocumentsPath
  
  SavePCSetting "Datapaths", "localolepath_" & gsDatabaseName, msLocalOlePath
  gsLocalOLEPath = msLocalOlePath
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    Unload Me
End Select

End Sub


Private Sub Form_Paint()
  'TM11092003 Fault 5887 + 6231
  If Not mbLoading And (Not Me.ActiveControl Is Nothing) Then
    Select Case Me.ActiveControl.Name
      Case "grdUtilityReport"
        grdUtilityReport.SelBookmarks.Add grdUtilityReport.Bookmark
        grdUtilityReport_RowColChange 0, 0
        grdUtilityReport.Refresh
        
      Case "lvwToolbars"
        lvwToolbars(CurrentToolBarIndex).Refresh
        lvwToolbars(CurrentToolBarIndex).SetFocus
    End Select
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim iAnswer As Integer
  
  If Changed = True Then
    iAnswer = COAMsgBox("You have changed the current configuration. Save changes?", vbYesNoCancel + vbExclamation, app.Title)
    
    Select Case iAnswer
      Case vbYes
        cmdOK_Click
        Exit Sub
      Case vbCancel
        Cancel = 1
        Exit Sub
    End Select
    
  End If
    
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub grdUtilityReport_Change()
  If Not mbLoading Then
    Changed = True
    
    ' The next line looks duff, but is required, honestly.
    If grdUtilityReport.Col = 1 Then
      grdUtilityReport.Col = 1
    End If
  End If

End Sub

Private Sub grdUtilityReport_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  Dim sAccessOptions As String
  Dim sOption As String
  Dim iLoop As Integer
  
  With grdUtilityReport
    ' Only display the required access options.
      sAccessOptions = mstrDefs(3, .AddItemRowIndex(.Bookmark))

      .Columns("DefaultAccess").RemoveAll

      Do While InStr(sAccessOptions, vbTab) > 0
        sOption = Left(sAccessOptions, InStr(sAccessOptions, vbTab) - 1)
        .Columns("DefaultAccess").AddItem AccessDescription(sOption)
        sAccessOptions = Mid(sAccessOptions, InStr(sAccessOptions, vbTab) + 1)
      Loop

      If Len(sAccessOptions) > 0 Then
        .Columns("DefaultAccess").AddItem AccessDescription(sAccessOptions)
      End If
   
    ' Set the styleSet of the rows to show which is selected.
    For iLoop = 0 To .Rows - 1
      If iLoop = .Row Then
        .Columns(0).CellStyleSet "ActiveText", iLoop
        .Columns(1).CellStyleSet "ActiveCheckBox", iLoop
        .Columns(2).CellStyleSet "ActiveText", iLoop
      Else
        .Columns(0).CellStyleSet "ssetEnabled", iLoop
        .Columns(1).CellStyleSet "ssetEnabled", iLoop
        .Columns(2).CellStyleSet "ssetEnabled", iLoop
      End If
    Next iLoop
    
'    If .Col >= 0 Then
'      If .Columns(.Col).Style = 2 Then
'        ' Checkbox cell
'        .ActiveCell.StyleSet = "ActiveCheckbox"
'      Else
'        ' Text cell
'        .ActiveCell.StyleSet = "ActiveText"
'      End If
'    End If
    
  End With

End Sub


Private Sub lblData_Click(Index As Integer)
  cboTextType.ListIndex = 2
End Sub

Private Sub lblHeading_Click(Index As Integer)
  cboTextType.ListIndex = 1
End Sub

Private Sub lblTitle_Click()
  cboTextType.ListIndex = 0
End Sub


Private Sub lstWarningMsg_Click()

  Changed = Not mbLoading

End Sub

Private Sub lvwToolbars_Click(Index As Integer)

  UpdateToolbarButtonStatus

End Sub

Private Sub lvwToolbars_DblClick(Index As Integer)

  lvwToolbars(Index).SelectedItem.SubItems(1) = IIf(lvwToolbars(Index).SelectedItem.SubItems(1) = "Hidden", "", "Hidden")
  UpdateToolbarButtonStatus
  Changed = Not mbLoading

End Sub

Private Sub lvwToolbars_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  UpdateToolbarButtonStatus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  Dim ctl As Control

  For Each ctl In Me.Controls
    If TypeOf ctl Is VB.Frame Then
      ctl.Enabled = (ctl.Left >= 0)
    End If
  Next


  'Only show the 'Set Defaults' button if either the Display Defaults
  'or Reports/Utilities tab is selected.
  Select Case SSTab1.Tab
  Case 0
    If cboPrimary.Visible And cboPrimary.Enabled Then
      cboPrimary.SetFocus
    End If
  Case 1
    If grdUtilityReport.Visible And grdUtilityReport.Enabled Then
      grdUtilityReport.SetFocus
    End If
  Case 2
    If cboPrinter.Visible And cboPrinter.Enabled Then
      cboPrinter.SetFocus
    End If
  Case 3
    If chkBatchLogon.Visible And chkBatchLogon.Enabled Then
      chkBatchLogon.SetFocus
    End If
  Case 6
    UpdateToolbarButtonStatus
    DoEvents
  End Select

End Sub
Private Sub VisibleDefsSetUpArray()
  ' Setup an array of the reports/utilities/tools that appear in the
  ' reports/utilities config grid.
  ' Column 1 = Utility/report name
  ' Column 2 = ASRSysUserSettings table key
  ' Column 3 = Tab delimited string of the access settings applicable to this report/utility
  Dim iLoop As Integer
  
  Const ACCESS_FULL = ACCESS_READWRITE & vbTab & ACCESS_READONLY & vbTab & ACCESS_HIDDEN
  Const ACCESS_NOTHIDDEN = ACCESS_READWRITE & vbTab & ACCESS_READONLY
  
  iLoop = 0
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Batch Jobs"
  mstrDefs(2, iLoop) = "Batch Jobs"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Calculations"
  mstrDefs(2, iLoop) = "Calculations"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Calendar Reports"
  mstrDefs(2, iLoop) = "Calendar Reports"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  If glngPersonnelTableID > 0 Then      'MH20030905 Fault 6587
    ReDim Preserve mstrDefs(3, iLoop)
    mstrDefs(1, iLoop) = "Career Progression"
    mstrDefs(2, iLoop) = "Career Progression"       'MH20030916 Fault 6370
    mstrDefs(3, iLoop) = ACCESS_FULL
    iLoop = iLoop + 1
  End If
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Cross Tabs"
  mstrDefs(2, iLoop) = "Cross Tabs"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Custom Reports"
  mstrDefs(2, iLoop) = "Custom Reports"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Data Transfer"
  mstrDefs(2, iLoop) = "Data Transfer"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  If gbVersion1Enabled Then
    ReDim Preserve mstrDefs(3, iLoop)
    mstrDefs(1, iLoop) = "Document Types"
    mstrDefs(2, iLoop) = "version1"
    mstrDefs(3, iLoop) = ACCESS_NOTHIDDEN
    iLoop = iLoop + 1
  End If
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Email Groups"
  mstrDefs(2, iLoop) = "Email Groups"
  mstrDefs(3, iLoop) = ACCESS_NOTHIDDEN
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Envelopes & Labels"
  mstrDefs(2, iLoop) = "Labels"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Envelope & Label Templates"
  mstrDefs(2, iLoop) = "Label Definition"
  mstrDefs(3, iLoop) = ACCESS_NOTHIDDEN
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Export"
  mstrDefs(2, iLoop) = "Export"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Filters"
  mstrDefs(2, iLoop) = "Filters"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1
  
  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Global Add"
  mstrDefs(2, iLoop) = "Global Add"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Global Update"
  mstrDefs(2, iLoop) = "Global Update"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Global Delete"
  mstrDefs(2, iLoop) = "Global Delete"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Import"
  mstrDefs(2, iLoop) = "Import"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Mail Merge"
  mstrDefs(2, iLoop) = "Mail Merge"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Match Reports"
  mstrDefs(2, iLoop) = "Match Reports"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Picklists"
  mstrDefs(2, iLoop) = "Picklists"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Report Packs"
  mstrDefs(2, iLoop) = "Report Packs"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  ReDim Preserve mstrDefs(3, iLoop)
  mstrDefs(1, iLoop) = "Record Profile"
  mstrDefs(2, iLoop) = "Record Profile"
  mstrDefs(3, iLoop) = ACCESS_FULL
  iLoop = iLoop + 1

  If glngPersonnelTableID > 0 Then      'MH20030905 Fault 6587
    ReDim Preserve mstrDefs(3, iLoop)
    mstrDefs(1, iLoop) = "Succession Planning"
    mstrDefs(2, iLoop) = "Succession Planning"   'MH20030916 Fault 6370
    mstrDefs(3, iLoop) = ACCESS_FULL
    iLoop = iLoop + 1
  End If

End Sub

Private Sub ShowWarningSetUpArray()

  ReDim mstrWarning(4) As String

  mstrWarning(0) = "Data Transfer"
  mstrWarning(1) = "Global Add"
  mstrWarning(2) = "Global Update"
  mstrWarning(3) = "Global Delete"
  mstrWarning(4) = "Import"

End Sub

Private Sub SaveUtilityReportSettings()

  Dim lngCount As Long
  Dim varBookmark As Variant
  
  With grdUtilityReport
    ' Use the Update method to ensure that the cell's Text/Value properties
    ' are passed to the CellText/CellValue methods.
    .Update
    
    For lngCount = 0 To (.Rows - 1)
      varBookmark = .AddItemBookmark(lngCount)
      
      SaveUserSetting "defsel", "onlymine " & Replace(mstrDefs(2, lngCount), " ", ""), IIf(.Columns("Selection").CellValue(varBookmark), 1, 0)
      SaveUserSetting "utils&reports", "dfltaccess " & Replace(mstrDefs(2, lngCount), " ", ""), AccessCode(.Columns("DefaultAccess").CellText(varBookmark))
    Next lngCount
  End With
  
End Sub

Private Sub ShowWarningSave()

  Dim lngCount As Long

  With lstWarningMsg
    For lngCount = 0 To UBound(mstrWarning)
      SaveUserSetting "warningmsg", "warning " & Replace(mstrWarning(lngCount), " ", ""), IIf(.Selected(lngCount), 1, 0)
    Next
  End With

End Sub

Private Sub txtBatchEmailAddr_Change()
  cmdBatchEmail.Enabled = (Trim(txtBatchEmailAddr.Text) <> vbNullString)
  Changed = Not mbLoading
End Sub

Private Sub txtDatabase_Change()
  Changed = Not mbLoading
End Sub

Private Sub txtFilename_Change(Index As Integer)
  cmdFileClear(Index).Enabled = (Trim(txtFilename(Index).Text) <> vbNullString)
End Sub

Private Sub txtPWD_Change()
  txtPWD.Tag = txtPWD.Text
  Changed = Not mbLoading
End Sub

Private Sub txtServer_Change()
  Changed = Not mbLoading
End Sub

Private Sub txtUID_Change()
  Changed = Not mbLoading
End Sub

Private Sub txtUID_GotFocus()
  With txtUID
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtPWD_GotFocus()
  With txtPWD
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDatabase_GotFocus()
  With txtDatabase
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtServer_GotFocus()
  With txtServer
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub


Private Function SavePCSettings() As Boolean

  Dim blnBatchLogon As Boolean
  Dim pstrPathErrors As String
  Dim objDefPrinter As cSetDfltPrinter

  SavePCSettings = True

  DebugOutput "MDIForm_Configuration", "PCSettingSaveStarted"

  'JPD 20030828 Fault 4287
  ' Tell the user that certain functionality will not be available if the
  ' Photo, OLE, Crystal, Documents or LocalOLE paths are not defined.
  If Len(Trim(txtDocumentsPath.Text)) = 0 Then
    pstrPathErrors = pstrPathErrors & "No default path will be stored if a documents path is not entered." & vbCrLf
  End If
  
  If Len(Trim(txtLocalOLEPath.Text)) = 0 Then
    pstrPathErrors = pstrPathErrors & "You will not be able to use local OLE fields if the path is not defined." & vbCrLf
  End If
  
  If Len(Trim(txtOLEPath.Text)) = 0 Then
    pstrPathErrors = pstrPathErrors & "You will not be able to use server OLE fields if the path is not defined." & vbCrLf
  End If
  
  If Len(Trim(txtPhotoPath.Text)) = 0 Then
    pstrPathErrors = pstrPathErrors & "You will not be able to use non-linked photo fields if the path is not defined." & vbCrLf
  End If
 
  If Len(Trim(txtCrystalPath.Text)) = 0 And (txtCrystalPath.Visible = True) Then
    pstrPathErrors = pstrPathErrors & "No default path will be stored if an Crystal path is not entered." & vbCrLf
  End If
  
  If Len(pstrPathErrors) > 0 Then
    Screen.MousePointer = vbDefault
    If COAMsgBox(pstrPathErrors & vbCrLf & vbCrLf & "Do you wish to continue ?", vbQuestion + vbYesNo, "Configuration") = vbNo Then
      SavePCSettings = False
      Exit Function
    End If
    Screen.MousePointer = vbHourglass
  End If
  
  SavePCSetting "Printer", "DeviceName", cboPrinter.Text
  gstrDefaultPrinterName = cboPrinter.Text

  ' NPG20110105 Fault HRPRO-1089
  ' NHRD04052011 JIRA OpenHR-1533 If you are coming from frmMain
  gblnStartupPrinter = False 'Don't run this bit of code for now (InStr(LCase(Command$), "/printer=true") > 0)
  If gblnStartupPrinter Then 'Or OpenedMe.Name = "frmMain" Then
  'If gblnStartupPrinter Or OpenedMe.Name = "frmMain" Then

    DebugOutput "MDIForm_Configuration", "SetPrinterAsDefault"

    Set objDefPrinter = New cSetDfltPrinter
    objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
    Set objDefPrinter = Nothing
  End If
  '******************************************************************************
   
  gbPrinterPrompt = (chkPrintingPrompt.Value = vbChecked)
  SavePCSetting "Printer", "Prompt", gbPrinterPrompt
  
  gbPrinterConfirm = (chkPrintingConfirm.Value = vbChecked)
  SavePCSetting "Printer", "Confirm", gbPrinterConfirm
  
  SavePCSetting "Datapaths", "photopath_" & gsDatabaseName, txtPhotoPath.Text
  gsPhotoPath = txtPhotoPath.Text
  
  SavePCSetting "Datapaths", "olepath_" & gsDatabaseName, txtOLEPath.Text
  gsOLEPath = txtOLEPath.Text
  
  SavePCSetting "Datapaths", "crystalpath_" & gsDatabaseName, txtCrystalPath.Text
  gsCrystalPath = txtCrystalPath.Text
  
  SavePCSetting "Datapaths", "documentspath_" & gsDatabaseName, txtDocumentsPath.Text
  gsDocumentsPath = txtDocumentsPath.Text
  
  SavePCSetting "Datapaths", "localolepath_" & gsDatabaseName, txtLocalOLEPath.Text
  gsLocalOLEPath = txtLocalOLEPath.Text
    
  blnBatchLogon = (chkBatchLogon.Value = vbChecked)
  SavePCSetting "BatchLogon", "Enabled", blnBatchLogon
  If blnBatchLogon Then
    SaveBatchLogon txtUID.Text, txtPWD.Tag, txtDatabase.Text, txtServer.Text
    SavePCSetting "BatchLogon", "TrustedConnection", IIf(chkUseWindowsAuthentication.Value = vbChecked, True, False)
    SavePCSetting "BatchLogon", "Email", txtBatchEmailAddr.Text
  Else
    SavePCSetting "BatchLogon", "Data", vbNullString
    SavePCSetting "BatchLogon", "TrustedConnection", False
    SavePCSetting "BatchLogon", "Email", vbNullString
  End If

  ' Automatic logon
  
  SavePCSetting "Login", "DataMgr_Bypass", chkBypassLogonDetails.Value

  DebugOutput "MDIForm_Configuration", "PCSettingSaveCompleted"


  Changed = False

End Function
Private Sub SaveUserSettings()

  SaveUserSetting "RecordEditing", "Primary", mcPrimary
  gcPrimary = mcPrimary
  
  SaveUserSetting "RecordEditing", "History", mcHistory
  gcHistory = mcHistory
  
  SaveUserSetting "RecordEditing", "Lookup", mcLookUp
  gcLookUp = mcLookUp
  
  SaveUserSetting "RecordEditing", "QuickAccess", mcQuickAccess
  gcQuickAccess = mcQuickAccess

  ' Expression builder configuration
  SaveUserSetting "ExpressionBuilder", "ViewColours", mcExpressionColours
  SaveUserSetting "ExpressionBuilder", "NodeSize", mcExpressionNodeSize

  SaveUtilityReportSettings
  
  ' Close defsel after run
  gbCloseDefSelAfterRun = IIf(chkCloseDefsel.Value = vbChecked, True, False)
  gbRecentDisplayDefSel = IIf(chkRunRecentImmediate.Value = vbChecked, True, False)
  gbRememberDefSelID = IIf(chkRememberDefSelID.Value = vbChecked, True, False)
  
  SaveUserSetting "DefSel", "CloseAfterRun", gbCloseDefSelAfterRun
  SaveUserSetting "DefSel", "RecentDisplayDefSel", gbRecentDisplayDefSel
  SaveUserSetting "DefSel", "RememberLastID", gbRememberDefSelID
  
  'TM20010727 Fault 1607 (Suggestion)
  ShowWarningSave

  'MH20000423 Diary Stuff
  SaveUserSetting "Diary", "ViewMode", cboDiaryView.ItemData(cboDiaryView.ListIndex)

  gblnDiaryConstCheck = (chkDiaryConstantCheck.Value = vbChecked)
  SaveUserSetting "Diary", "ConstantCheck", gblnDiaryConstCheck

  'MH20041104
  SaveUserSetting "Email", "IncludeRecDesc", (chkEmailRecDesc.Value = vbChecked)


  'If gblnDiaryConstCheck Then
    With frmMain
      .tmrDiary.Enabled = (gblnDiaryConstCheck And datGeneral.SystemPermission("DIARY", "MANUALEVENTS"))
      .tmrDiary.Interval = 1   'Force diary check
      .RefreshMainForm Me, True
    End With
  'End If

  ' Event Log Settings
  SaveUserSetting "LogEvents", "Global_Update_Success", (chkGlobalUpdateSuccess.Value = vbChecked)
  SaveUserSetting "LogEvents", "Global_Add_Success", (chkGlobalAddSuccess.Value = vbChecked)
  SaveUserSetting "LogEvents", "Global_Delete_Success", (chkGlobalDeleteSuccess.Value = vbChecked)
  SaveUserSetting "LogEvents", "Import_Success", (chkImportSuccess.Value = vbChecked)
  SaveUserSetting "LogEvents", "Export_Success", (chkExportSuccess.Value = vbChecked)
  SaveUserSetting "LogEvents", "Data_Transfer_Success", (chkDataTransferSuccess = vbChecked)
  
  
  With moutTitle
    SaveUserSetting "Output", "TitleCol", .StartCol
    SaveUserSetting "Output", "TitleRow", .StartRow
    SaveUserSetting "Output", "TitleGridLines", IIf(.Gridlines, 1, 0)
    SaveUserSetting "Output", "TitleBold", IIf(.Bold, 1, 0)
    SaveUserSetting "Output", "TitleUnderline", IIf(.Underline, 1, 0)
    SaveUserSetting "Output", "TitleBackcolour", .BackCol
    SaveUserSetting "Output", "TitleForecolour", .ForeCol
  End With

  With moutHeading
    SaveUserSetting "Output", "HeadingCol", .StartCol
    SaveUserSetting "Output", "HeadingRow", .StartRow
    SaveUserSetting "Output", "HeadingGridLines", IIf(.Gridlines, 1, 0)
    SaveUserSetting "Output", "HeadingBold", IIf(.Bold, 1, 0)
    SaveUserSetting "Output", "HeadingUnderline", IIf(.Underline, 1, 0)
    SaveUserSetting "Output", "HeadingBackcolour", .BackCol
    SaveUserSetting "Output", "HeadingForecolour", .ForeCol
  End With

  With moutData
    SaveUserSetting "Output", "DataCol", .StartCol
    SaveUserSetting "Output", "DataRow", .StartRow
    SaveUserSetting "Output", "DataGridLines", IIf(.Gridlines, 1, 0)
    SaveUserSetting "Output", "DataBold", IIf(.Bold, 1, 0)
    SaveUserSetting "Output", "DataUnderline", IIf(.Underline, 1, 0)
    SaveUserSetting "Output", "DataBackcolour", .BackCol
    SaveUserSetting "Output", "DataForecolour", .ForeCol
  End With

  SaveUserSetting "Output", "ExcelTemplate", txtFilename(0).Text
  SaveUserSetting "Output", "WordTemplate", txtFilename(1).Text
  SaveUserSetting "Output", "ExcelHeaders", IIf(chkExcelHeaders.Value = vbChecked, 1, 0)
  SaveUserSetting "Output", "ExcelGridlines", IIf(chkExcelGridlines.Value = vbChecked, 1, 0)
  SaveUserSetting "Output", "ExcelOmitSpacerRow", IIf(chkOmitTopRow.Value = vbChecked, 1, 0)
  SaveUserSetting "Output", "ExcelOmitSpacerCol", IIf(chkOmitSpacerCol.Value = vbChecked, 1, 0)
  
  SaveUserSetting "Toolbar", "Position", mlngToolbarPosition
  SaveToolBarOptions

  SaveUserSetting "Output", "WordFormat", cboWordFormat.ItemData(cboWordFormat.ListIndex)
  SaveUserSetting "Output", "ExcelFormat", cboExcelFormat.ItemData(cboExcelFormat.ListIndex)

  Changed = False
  
End Sub


Private Sub PopulateControlsUserSettings()

  Dim objControl As Control
  Dim lngCount As Long
  Dim iItemX As ListItem

  ' Add colour options
  With cboColours
    .Clear
    .AddItem "Black"
    .ItemData(cboColours.NewIndex) = EXPRESSIONBUILDER_COLOUROFF
    .AddItem "Colour Levels"
    .ItemData(cboColours.NewIndex) = EXPRESSIONBUILDER_COLOURON
  End With

  ' Node statuses
  With cboNodeSize
    .Clear
    .AddItem "Minimized"
    .ItemData(cboNodeSize.NewIndex) = EXPRESSIONBUILDER_NODESMINIMIZE
    .AddItem "Expand All"
    .ItemData(cboNodeSize.NewIndex) = EXPRESSIONBUILDER_NODESEXPAND
    .AddItem "Expand Top Level"
    .ItemData(cboNodeSize.NewIndex) = EXPRESSIONBUILDER_NODESTOPLEVEL
'    .AddItem "As Last Save"
'    .ItemData(cboNodeSize.NewIndex) = EXPRESSIONBUILDER_NODESLASTSAVE
  End With

  With cboDiaryView
    .Clear
    .AddItem "Day View"
    .ItemData(.NewIndex) = 0
    .AddItem "Week View"
    .ItemData(.NewIndex) = 1
    .AddItem "Six Month View"
    .ItemData(.NewIndex) = 2
    .AddItem "List View"
    .ItemData(.NewIndex) = 3
  End With
  
  
  For Each objControl In Me.Controls
    If objControl.Name = "cboPrimary" Or _
       objControl.Name = "cboHistory" Or _
       objControl.Name = "cboLookUp" Or _
       objControl.Name = "cboQuickAccess" Then
    
      With objControl
        .Clear
        .AddItem "New Record"
        .ItemData(objControl.NewIndex) = disRecEdit_New
      
        .AddItem "First Record"
        .ItemData(objControl.NewIndex) = disRecEdit_First
      
        .AddItem "Find Window"
        .ItemData(objControl.NewIndex) = disFindWindow
      End With

    End If
  Next objControl
  
  ' Load the toolbar types
  With cboToolbars
    .Clear
    .AddItem "Record Editing"
    .ItemData(.NewIndex) = 1
    .AddItem "Find Window"
    .ItemData(.NewIndex) = 2
  End With
  
  ' Load toolbar locations
  With cboToolbarPosition
    .Clear
    .AddItem "Top"
    .ItemData(.NewIndex) = giTOOLBAR_TOP
    .AddItem "Left"
    .ItemData(.NewIndex) = giTOOLBAR_LEFT
    .AddItem "Right"
    .ItemData(.NewIndex) = giTOOLBAR_RIGHT
    .AddItem "Bottom"
    .ItemData(.NewIndex) = giTOOLBAR_BOTTOM
    .AddItem "None"
    .ItemData(.NewIndex) = giTOOLBAR_NONE
  End With
  
  Call VisibleDefsSetUpArray
  With grdUtilityReport
    .RemoveAll
    
    For lngCount = 0 To UBound(mstrDefs, 2)
      .AddItem mstrDefs(1, lngCount) & _
        vbTab & GetUserSettingOrDefault("defsel", "onlymine " & Replace(mstrDefs(2, lngCount), " ", ""), 0) & _
        vbTab & AccessDescription(GetUserSettingOrDefault("utils&reports", "dfltaccess " & Replace(mstrDefs(2, lngCount), " ", ""), ACCESS_READWRITE))
    Next
    
    .MoveFirst
    .Col = 1
  End With

  Call ShowWarningSetUpArray
  With lstWarningMsg
    For lngCount = 0 To UBound(mstrWarning)
      If .ListCount <= lngCount Then
        .AddItem mstrWarning(lngCount)
      End If
    Next
    .ListIndex = 0
  End With

  mlngWordFormat = GetOutputFormats("Word", cboWordFormat, GetOfficeWordVersion)
  mlngExcelFormat = GetOutputFormats("Excel", cboExcelFormat, GetOfficeExcelVersion)

End Sub

Private Sub ReadUserSettings(blnRestoreDefaults As Boolean)

  Dim objControl As Control
  Dim lngCount As Long

  Dim bLoggingGUSuccess As Boolean
  Dim bLoggingGASuccess As Boolean
  Dim bLoggingGDSuccess As Boolean
  Dim bLoggingImportSuccess As Boolean
  Dim bLoggingExportSuccess As Boolean
  Dim bLoggingDTSuccess As Boolean

  mblnRestoringDefaults = blnRestoreDefaults

  Set moutTitle = New clsOutputStyle
  Set moutHeading = New clsOutputStyle
  Set moutData = New clsOutputStyle
  
  
  With moutTitle
    .StartCol = Val(GetUserSettingOrDefault("Output", "TitleCol", "3"))
    .StartRow = Val(GetUserSettingOrDefault("Output", "TitleRow", "2"))
    .Gridlines = (GetUserSettingOrDefault("Output", "TitleGridLines", "0") = "1")
    .Bold = (GetUserSettingOrDefault("Output", "TitleBold", "1") = "1")
    .Underline = (GetUserSettingOrDefault("Output", "TitleUnderline", "0") = "1")
    .BackCol = Val(GetUserSettingOrDefault("Output", "TitleBackcolour", vbWhite))
    .ForeCol = Val(GetUserSettingOrDefault("Output", "TitleForecolour", GetColour("Midnight Blue")))
  
    RefreshGridlines 0, .Gridlines
    RefreshBold 0, .Bold
    RefreshUnderLine 0, .Underline
    RefreshBackColour 0, .BackCol
    RefreshForeColour 0, .ForeCol
  End With

  With moutHeading
    .StartCol = Val(GetUserSettingOrDefault("Output", "HeadingCol", "2"))
    .StartRow = Val(GetUserSettingOrDefault("Output", "HeadingRow", "4"))
    .Gridlines = (GetUserSettingOrDefault("Output", "HeadingGridLines", "1") = "1")
    .Bold = (GetUserSettingOrDefault("Output", "HeadingBold", "1") = "1")
    .Underline = (GetUserSettingOrDefault("Output", "HeadingUnderline", "0") = "1")
    .BackCol = Val(GetUserSettingOrDefault("Output", "HeadingBackcolour", GetColour("Dolphin Blue")))
    .ForeCol = Val(GetUserSettingOrDefault("Output", "HeadingForecolour", GetColour("Midnight Blue")))
  
    RefreshGridlines 1, .Gridlines
    RefreshBold 1, .Bold
    RefreshUnderLine 1, .Underline
    RefreshBackColour 1, .BackCol
    RefreshForeColour 1, .ForeCol
  End With

  With moutData
    .StartCol = Val(GetUserSettingOrDefault("Output", "DataCol", "2"))
    .StartRow = Val(GetUserSettingOrDefault("Output", "DataRow", "5"))
    .Gridlines = (GetUserSettingOrDefault("Output", "DataGridLines", "1") = "1")
    .Bold = (GetUserSettingOrDefault("Output", "DataBold", "0") = "1")
    .Underline = (GetUserSettingOrDefault("Output", "DataUnderline", "0") = "1")
    .BackCol = Val(GetUserSettingOrDefault("Output", "DataBackcolour", GetColour("Pale Grey")))
    .ForeCol = Val(GetUserSettingOrDefault("Output", "DataForecolour", GetColour("Midnight Blue")))
  
    RefreshGridlines 2, .Gridlines
    RefreshBold 2, .Bold
    RefreshUnderLine 2, .Underline
    RefreshBackColour 2, .BackCol
    RefreshForeColour 2, .ForeCol
  
  End With

  
  txtFilename(0).Text = GetUserSettingOrDefault("Output", "ExcelTemplate", vbNullString)
  txtFilename(1).Text = GetUserSettingOrDefault("Output", "WordTemplate", vbNullString)
  chkExcelHeaders.Value = IIf(GetUserSettingOrDefault("Output", "ExcelHeaders", 0) = 1, vbChecked, vbUnchecked)
  chkExcelGridlines.Value = IIf(GetUserSettingOrDefault("Output", "ExcelGridlines", 0) = 1, vbChecked, vbUnchecked)
  chkOmitSpacerCol.Value = IIf(GetUserSettingOrDefault("Output", "ExcelOmitSpacerCol", 0) = 1, vbChecked, vbUnchecked)
  chkOmitTopRow.Value = IIf(GetUserSettingOrDefault("Output", "ExcelOmitSpacerRow", 0) = 1, vbChecked, vbUnchecked)
  
  mcPrimary = GetUserSettingOrDefault("RecordEditing", "Primary", disFindWindow)
  SetCombo cboPrimary, mcPrimary
  
  mcHistory = GetUserSettingOrDefault("RecordEditing", "History", disFindWindow)
  SetCombo cboHistory, mcHistory
  
  'JPD20011005 Fault 2721 Default now set to Find Window
  mcLookUp = GetUserSettingOrDefault("RecordEditing", "LookUp", disFindWindow)
  SetCombo cboLookUp, mcLookUp
  
  mcQuickAccess = GetUserSettingOrDefault("RecordEditing", "QuickAccess", disRecEdit_New)
  SetCombo cboQuickAccess, mcQuickAccess

  'JDM - 16/03/01 - Fault 1935 - Defaults on expressions
  'JDM - Removed colour as last saved
  mcExpressionColours = GetUserSettingOrDefault("ExpressionBuilder", "ViewColours", EXPRESSIONBUILDER_COLOUROFF)
  If mcExpressionColours = EXPRESSIONBUILDER_COLOURLASTSAVE Then
    SetCombo cboColours, EXPRESSIONBUILDER_COLOUROFF
  Else
    SetCombo cboColours, mcExpressionColours
  End If
  
  ' JDM - 03/01/02 - Fault 3316 - Remove last save as an option
  mcExpressionNodeSize = GetUserSettingOrDefault("ExpressionBuilder", "NodeSize", EXPRESSIONBUILDER_NODESMINIMIZE)
  If mcExpressionNodeSize = EXPRESSIONBUILDER_NODESLASTSAVE Then
    SetCombo cboNodeSize, EXPRESSIONBUILDER_NODESEXPAND
  Else
    SetCombo cboNodeSize, mcExpressionNodeSize
  End If
  
  mlngDiaryDefaultView = GetUserSettingOrDefault("Diary", "ViewMode", 1)
  SetComboItem cboDiaryView, mlngDiaryDefaultView
  
  chkDiaryConstantCheck.Value = IIf(gblnDiaryConstCheck, vbChecked, vbUnchecked)
  
  'MH20041104
  chkEmailRecDesc.Value = IIf(CBool(GetUserSettingOrDefault("Email", "IncludeRecDesc", True)), vbChecked, vbUnchecked)
 
  With lstWarningMsg
    For lngCount = 0 To UBound(mstrWarning)
      .Selected(lngCount) = GetUserSettingOrDefault("warningmsg", "warning " & Replace(mstrWarning(lngCount), " ", ""), 1)
    Next
    .ListIndex = 0
  End With

  mbCloseDefSelAfterRun = CBool(GetUserSettingOrDefault("DefSel", "CloseAfterRun", False))
  chkCloseDefsel.Value = IIf(mbCloseDefSelAfterRun, vbChecked, vbUnchecked)

  mbRecentDisplayDefSel = CBool(GetUserSettingOrDefault("DefSel", "RecentDisplayDefSel", False))
  chkRunRecentImmediate.Value = IIf(mbRecentDisplayDefSel, vbChecked, vbUnchecked)

  chkRememberDefSelID.Value = IIf(gbRememberDefSelID, vbChecked, vbUnchecked)

  ' Log successful Global Adds
  bLoggingGASuccess = CBool(GetUserSettingOrDefault("LogEvents", "Global_Add_Success", False))
  chkGlobalAddSuccess.Value = IIf(bLoggingGASuccess, vbChecked, vbUnchecked)
  
  ' Log successful Global Updates
  bLoggingGUSuccess = CBool(GetUserSettingOrDefault("LogEvents", "Global_Update_Success", False))
  chkGlobalUpdateSuccess.Value = IIf(bLoggingGUSuccess, vbChecked, vbUnchecked)
  
  ' Log successful Global Deletes
  bLoggingGDSuccess = CBool(GetUserSettingOrDefault("LogEvents", "Global_Delete_Success", False))
  chkGlobalDeleteSuccess.Value = IIf(bLoggingGDSuccess, vbChecked, vbUnchecked)
  
  ' Log successful Imports
  bLoggingImportSuccess = CBool(GetUserSettingOrDefault("LogEvents", "Import_Success", False))
  chkImportSuccess.Value = IIf(bLoggingImportSuccess, vbChecked, vbUnchecked)
  
  ' Log successful Exports
  bLoggingExportSuccess = CBool(GetUserSettingOrDefault("LogEvents", "Export_Success", False))
  chkExportSuccess.Value = IIf(bLoggingExportSuccess, vbChecked, vbUnchecked)
  
  ' Log successful Data Transfers
  bLoggingDTSuccess = CBool(GetUserSettingOrDefault("LogEvents", "Data_Transfer_Success", False))
  chkDataTransferSuccess.Value = IIf(bLoggingDTSuccess, vbChecked, vbUnchecked)

  ' Load the toolbar options
  cboToolbars.ListIndex = 1

  ' Load the toolbar position
  mlngToolbarPosition = GetUserSettingOrDefault("Toolbar", "Position", giTOOLBAR_TOP)
  SetCombo cboToolbarPosition, mlngToolbarPosition

  mlngWordFormat = Val(GetUserSettingOrDefault("Output", "WordFormat", mlngWordFormat))     'WdSaveFormat.wdFormatDocument97
  SetComboItem cboWordFormat, mlngWordFormat
  If cboWordFormat.ListIndex < 0 And cboWordFormat.ListCount > 0 Then
    cboWordFormat.ListIndex = 0
  End If

  mlngExcelFormat = Val(GetUserSettingOrDefault("Output", "ExcelFormat", mlngExcelFormat))  'XlFileFormat.xlExcel8
  SetComboItem cboExcelFormat, mlngExcelFormat
  If cboExcelFormat.ListIndex < 0 And cboExcelFormat.ListCount > 0 Then
    cboExcelFormat.ListIndex = 0
  End If

  cboTextType.ListIndex = 0

End Sub

'Private Sub ReadUserSettingDefaults()
'
'  Dim objControl As Control
'  Dim lngCount As Long
'
'  mcPrimary = disFindWindow
'  SetCombo cboPrimary, mcPrimary
'
'  mcHistory = disFindWindow
'  SetCombo cboHistory, mcHistory
'
'  'JPD20011005 Fault 2721 Default now set to Find Window
'  mcLookUp = disFindWindow
'  SetCombo cboLookUp, mcLookUp
'
'  mcQuickAccess = disRecEdit_New
'  SetCombo cboQuickAccess, mcQuickAccess
'
'  'JDM - 16/03/01 - Fault 1935 - Defaults on expressions
'  'JDM - Removed colour as last saved
'  mcExpressionColours = EXPRESSIONBUILDER_COLOUROFF
'  If mcExpressionColours = EXPRESSIONBUILDER_COLOURLASTSAVE Then
'    SetCombo cboColours, EXPRESSIONBUILDER_COLOUROFF
'  Else
'    SetCombo cboColours, mcExpressionColours
'  End If
'
'  ' JDM - 03/01/02 - Fault 3316 - Remove last save as an option
'  mcExpressionNodeSize = GetUserSettingOrDefault("ExpressionBuilder", "NodeSize", EXPRESSIONBUILDER_NODESMINIMIZE)
'  If mcExpressionNodeSize = EXPRESSIONBUILDER_NODESLASTSAVE Then
'    SetCombo cboNodeSize, EXPRESSIONBUILDER_NODESEXPAND
'  Else
'    SetCombo cboNodeSize, mcExpressionNodeSize
'  End If
'
'  mlngDiaryDefaultView = 1
'  SetComboItem cboDiaryView, mlngDiaryDefaultView
'
'  'chkDiaryConstantCheck.Value = IIf(gblnDiaryConstCheck, vbChecked, vbUnchecked)
'  chkDiaryConstantCheck.Value = vbChecked
'
'  'MH20041104
'  chkEmailRecDesc.Value = vbChecked
'
'  With grdUtilityReport
'    For lngCount = 0 To (.Rows - 1)
'      .Bookmark = .AddItemBookmark(lngCount)
'
'      .Columns("Selection").Text = False
'      .Columns("DefaultAccess").Text = AccessDescription(ACCESS_READWRITE)
'    Next lngCount
'
'    .MoveFirst
'    .Col = 0
'    cmdDefault.SetFocus
'  End With
'
'  With lstWarningMsg
'    For lngCount = 0 To UBound(mstrWarning)
'      .Selected(lngCount) = 1
'    Next
'    .ListIndex = 0
'  End With
'
'  chkCloseDefsel.Value = vbUnchecked
'  chkGlobalAddSuccess.Value = vbUnchecked
'  chkGlobalUpdateSuccess.Value = vbUnchecked
'  chkGlobalDeleteSuccess.Value = vbUnchecked
'  chkDataTransferSuccess.Value = vbUnchecked
'  chkExportSuccess.Value = vbUnchecked
'  chkImportSuccess.Value = vbUnchecked
'
'  'Title
'  With moutTitle
'    .Gridlines = False
'    .Bold = True
'    .Underline = False
'    .BackCol = vbWhite
'    .ForeCol = GetColour("Midnight Blue")
'
'    RefreshGridlines 0, .Gridlines
'    RefreshBold 0, .Bold
'    RefreshUnderLine 0, .Underline
'    RefreshBackColour 0, .BackCol
'    RefreshForeColour 0, .ForeCol
'  End With
'
'  With moutHeading
'    .Gridlines = True
'    .Bold = True
'    .Underline = False
'    .BackCol = GetColour("Dolphin Blue")
'    .ForeCol = GetColour("Midnight Blue")
'
'    RefreshGridlines 1, .Gridlines
'    RefreshBold 1, .Bold
'    RefreshUnderLine 1, .Underline
'    RefreshBackColour 1, .BackCol
'    RefreshForeColour 1, .ForeCol
'  End With
'
'  With moutData
'    .Gridlines = True
'    .Bold = False
'    .Underline = False
'    .BackCol = GetColour("Pale Grey")
'    .ForeCol = GetColour("Midnight Blue")
'
'    RefreshGridlines 2, .Gridlines
'    RefreshBold 2, .Bold
'    RefreshUnderLine 2, .Underline
'    RefreshBackColour 2, .BackCol
'    RefreshForeColour 2, .ForeCol
'  End With
'
'  txtFilename(0).Text = vbNullString
'  txtFilename(1).Text = vbNullString
'  chkExcelHeaders.Value = False
'  chkExcelGridlines.Value = False
'
'  'NHRD04062003 Fault 5784
'  ' reseting the listindex kicks off the
'  'click events and gets the default we want
'  'cboTextType_Click
'  cboTextType.ListIndex = 0
'
'  ' Re-read the default toolbar settings
'  LoadToolBarOptions True
'  SetCombo cboToolbarPosition, giTOOLBAR_TOP
'
'  Changed = True
'
'End Sub


Private Sub ReadPCSettings()
  
  Dim lngBatchLogon As Boolean
  Dim strUserName As String
  Dim strPassword As String
  Dim strDatabase As String
  Dim strServer As String
  Dim bBypassLoginScreen As Boolean


  lngBatchLogon = GetPCSetting("BatchLogon", "Enabled", False)
  chkBatchLogon.Value = IIf(lngBatchLogon, vbChecked, vbUnchecked)

  If lngBatchLogon Then
    GetBatchLogon strUserName, strPassword, strDatabase, strServer
    txtUID.Text = strUserName
    txtPWD.Text = Space(20)         'Don't show the actual length of the password!
    txtPWD.Tag = strPassword
    txtDatabase = strDatabase
    txtServer = strServer
    
    chkUseWindowsAuthentication.Value = IIf(GetPCSetting("BatchLogon", "TrustedConnection", False), vbChecked, vbUnchecked)
    txtBatchEmailAddr.Text = GetPCSetting("BatchLogon", "Email", vbNullString)
    chkBatchEmail.Value = IIf(txtBatchEmailAddr.Text <> vbNullString, vbChecked, vbUnchecked)

  End If


  ' Initialise the form controls with the configuration details.
  msPhotoPath = gsPhotoPath
  txtPhotoPath.Text = gsPhotoPath
  cmdPhotoPathClear.Enabled = Not (gsPhotoPath = "")

  msOLEPath = gsOLEPath
  txtOLEPath.Text = gsOLEPath
  cmdOLEPathClear.Enabled = Not (gsOLEPath = "")

  msCrystalPath = gsCrystalPath
  txtCrystalPath.Text = gsCrystalPath
  cmdCrystalPathClear.Enabled = Not (gsCrystalPath = "")

  msDocumentsPath = gsDocumentsPath
  txtDocumentsPath.Text = gsDocumentsPath
  cmdDocumentsPathClear.Enabled = Not (gsDocumentsPath = "")

  msLocalOlePath = gsLocalOLEPath
  txtLocalOLEPath.Text = gsLocalOLEPath
  cmdLocalOLEPathClear.Enabled = Not (gsLocalOLEPath = "")

  'Printing options
  'TM20020911 Fault 4401
  If Printers.Count > 0 Then
    'mstrDefaultPrinter = GetPCSetting( "Printer", "DeviceName", "")
    Printer.TrackDefault = True
    mstrDefaultPrinter = Printer.DeviceName
  End If
  
  chkPrintingPrompt.Value = IIf(gbPrinterPrompt, vbChecked, vbUnchecked)
  chkPrintingConfirm.Value = IIf(gbPrinterConfirm, vbChecked, vbUnchecked)

  DoPrinterTab

  ' Automatic logon
  bBypassLoginScreen = GetPCSetting("Login", "DataMgr_Bypass", False)
  chkBypassLogonDetails.Value = IIf(bBypassLoginScreen, vbChecked, vbUnchecked)
  chkBypassLogonDetails.Enabled = gbUseWindowsAuthentication

End Sub


Private Function ADOConError(objTestConn As ADODB.Connection) As String

  Dim strErrorDesc As String
  Dim lngCount As Long

  strErrorDesc = vbNullString
  If Not objTestConn Is Nothing Then
    If Not objTestConn.Errors Is Nothing Then
      For lngCount = 0 To objTestConn.Errors.Count - 1
        strErrorDesc = objTestConn.Errors(lngCount).Description
      Next
      strErrorDesc = Mid(strErrorDesc, InStrRev(strErrorDesc, "]") + 1)
    End If
  End If

  ADOConError = strErrorDesc

End Function



Private Sub Form_Load()

  Dim i As Long
  Dim colorUpper As String
  Dim colorProper As String
  
  SSTab1_Click 0
  
  PopulateColourCombos
  
  CallingForm = Me
  
  With cboTextType
    .Clear
    .AddItem "Title"
    .AddItem "Heading"
    .AddItem "Data"
  End With
  
  cboForeColour.Refresh
  cboBackColour.Refresh
  
  grdUtilityReport.RowHeight = 239
  cmdOK.Enabled = False

  ' Resize the configuration form
  'SSTab1.Width = (fraDisplay(0).Left * 2) + fraDisplay(0).Width
  'Me.Width = SSTab1.Width + (SSTab1.Left * 2) + 90

End Sub

'Private Sub ImageCombo1_Click()
'
'  picSelected.BackColor = GetColorFromString(ImageCombo.SelectedItem.Key)
'  lblSelected.Caption = ImageCombo.SelectedItem.Text
'
'End Sub

Private Sub CreateColorImage(strColDesc As String, lngColValue)
  
  'sColor = StrConv(sColor, vbUpperCase)
  

End Sub


Private Sub PopulateColourCombos()
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strColour As String
  Dim lngColour As Long
  
  On Local Error GoTo LocalErr

  strSQL = "SELECT ColValue, ColDesc " & _
           "FROM ASRSysColours " & _
           "ORDER BY ColOrder"
           
  
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  Do While Not rsTemp.EOF
 
    With picColour
      .AutoRedraw = True
      .Width = picColour.ScaleX(12, vbPixels, vbTwips)
      .Height = picColour.ScaleY(12, vbPixels, vbTwips)
      .BorderStyle = 0
      .Appearance = 0
      .BackColor = rsTemp!ColValue
      .ForeColor = vbBlack
      picColour.Line (0, 0)-(picColour.Width - picColour.ScaleX(1, vbPixels, vbTwips) _
                 , picColour.Height - picColour.ScaleY(1, vbPixels, vbTwips)), , B
      .Picture = picColour.Image
      ImageList1.ListImages.Add , rsTemp!ColDesc, .Picture
      .Cls
      .Picture = Nothing
    End With

    rsTemp.MoveNext
  Loop

  If Not rsTemp.BOF Or Not rsTemp.EOF Then
    rsTemp.MoveFirst
  
    cboForeColour.ComboItems.Clear
    cboBackColour.ComboItems.Clear

    cboForeColour.ImageList = ImageList1
    cboBackColour.ImageList = ImageList1
    
    Do While Not rsTemp.EOF

      strColour = rsTemp!ColDesc
      lngColour = rsTemp!ColValue

      cboForeColour.Refresh
      cboForeColour.ComboItems.Add , "C" & CStr(lngColour), strColour, strColour
      cboForeColour.ComboItems.Item(cboForeColour.ComboItems.Count).Tag = lngColour
      cboBackColour.Refresh
      cboBackColour.ComboItems.Add , "C" & CStr(lngColour), strColour, strColour
      cboBackColour.ComboItems.Item(cboBackColour.ComboItems.Count).Tag = lngColour

      rsTemp.MoveNext
    Loop
  End If
  
  rsTemp.Close

  Set rsTemp = Nothing

Exit Sub

LocalErr:
  COAMsgBox Err.Description, vbExclamation, Me.Caption

End Sub

Private Sub cboTextType_Click()
  
  Dim lngCount As Long
  
  'NHRD04062003 Fault 5785 Added this control variable.
  'Helps with the Refreshxxxx Subs i.e.
  'it will not enable the OK button too soon.
  mbLoading = True
  
  Select Case cboTextType.ListIndex
  Case 0
    Set moutCurrent = moutTitle
    chkGridlines.Enabled = False
    lblBackColour.Enabled = False
    cboBackColour.Enabled = False
    cboBackColour.BackColor = vbButtonFace
  Case 1
    Set moutCurrent = moutHeading
    chkGridlines.Enabled = True
    lblBackColour.Enabled = True
    cboBackColour.Enabled = True
    cboBackColour.BackColor = vbWindowBackground
  Case 2
    Set moutCurrent = moutData
    chkGridlines.Enabled = True
    lblBackColour.Enabled = True
    cboBackColour.Enabled = True
    cboBackColour.BackColor = vbWindowBackground
  End Select
  
  For lngCount = 1 To cboBackColour.ComboItems.Count
    If cboBackColour.ComboItems.Item(lngCount).Tag = moutCurrent.BackCol Then
      cboBackColour.SelectedItem = cboBackColour.ComboItems.Item(lngCount)
      Exit For
    End If
  Next
  
  For lngCount = 1 To cboForeColour.ComboItems.Count
    If cboForeColour.ComboItems.Item(lngCount).Tag = moutCurrent.ForeCol Then
      cboForeColour.SelectedItem = cboForeColour.ComboItems.Item(lngCount)
      Exit For
    End If
  Next
  
  chkBold.Value = IIf(moutCurrent.Bold, vbChecked, vbUnchecked)
  chkUnderLine.Value = IIf(moutCurrent.Underline, vbChecked, vbUnchecked)
  chkGridlines.Value = IIf(moutCurrent.Gridlines, vbChecked, vbUnchecked)
  
  mbLoading = False
  
End Sub

Private Sub cboBackColour_Click()
  moutCurrent.BackCol = cboBackColour.ComboItems.Item(cboBackColour.SelectedItem.Index).Tag
  RefreshBackColour cboTextType.ListIndex, moutCurrent.BackCol
End Sub

Private Sub RefreshBackColour(lngIndex As Long, lngBackCol)

  Select Case lngIndex
  Case 0
    lblTitle.BackColor = lngBackCol
  Case 1
    lblHeading(0).BackColor = lngBackCol
    lblHeading(1).BackColor = lngBackCol
  Case 2
    lblData(0).BackColor = lngBackCol
    lblData(1).BackColor = lngBackCol
    lblData(2).BackColor = lngBackCol
    lblData(3).BackColor = lngBackCol
  End Select

  Changed = True

End Sub

Private Sub cboForeColour_Click()
  moutCurrent.ForeCol = cboForeColour.ComboItems.Item(cboForeColour.SelectedItem.Index).Tag
  RefreshForeColour cboTextType.ListIndex, moutCurrent.ForeCol
End Sub

Private Sub RefreshForeColour(lngIndex As Long, lngForeCol As Long)

  Select Case lngIndex
  Case 0
    lblTitle.ForeColor = lngForeCol
  Case 1
    lblHeading(0).ForeColor = lngForeCol
    lblHeading(1).ForeColor = lngForeCol
  Case 2
    lblData(0).ForeColor = lngForeCol
    lblData(1).ForeColor = lngForeCol
    lblData(2).ForeColor = lngForeCol
    lblData(3).ForeColor = lngForeCol
  End Select

  Changed = True

End Sub

Private Sub chkBold_Click()
  moutCurrent.Bold = (chkBold.Value = vbChecked)
  RefreshBold cboTextType.ListIndex, moutCurrent.Bold
End Sub

Private Sub RefreshBold(lngIndex As Long, blnBold As Boolean)

  Select Case lngIndex
  Case 0
    lblTitle.FontBold = blnBold
  Case 1
    lblHeading(0).FontBold = blnBold
    lblHeading(1).FontBold = blnBold
  Case 2
    lblData(0).FontBold = blnBold
    lblData(1).FontBold = blnBold
    lblData(2).FontBold = blnBold
    lblData(3).FontBold = blnBold
  End Select

  Changed = Not mbLoading

End Sub

Private Sub chkUnderLine_Click()
  moutCurrent.Underline = (chkUnderLine.Value = vbChecked)
  RefreshUnderLine cboTextType.ListIndex, moutCurrent.Underline
End Sub

Private Sub RefreshUnderLine(lngIndex, blnUnderline As Boolean)

  Select Case lngIndex
  Case 0
    lblTitle.FontUnderline = blnUnderline
  Case 1
    lblHeading(0).FontUnderline = blnUnderline
    lblHeading(1).FontUnderline = blnUnderline
  Case 2
    lblData(0).FontUnderline = blnUnderline
    lblData(1).FontUnderline = blnUnderline
    lblData(2).FontUnderline = blnUnderline
    lblData(3).FontUnderline = blnUnderline
  End Select

  Changed = Not mbLoading

End Sub

Private Sub chkGridlines_Click()
  moutCurrent.Gridlines = (chkGridlines.Value = vbChecked)
  RefreshGridlines cboTextType.ListIndex, moutCurrent.Gridlines
End Sub

Private Sub RefreshGridlines(lngIndex As Long, blnGridlines As Boolean)
  
  Dim lngGridlines As Long
  
  lngGridlines = IIf(blnGridlines, 1, 0)
  
  Select Case lngIndex
  'Can't have gridlines on title...
  Case 0
    lblTitle.BorderStyle = 0  'lngGridlines
  Case 1
    lblHeading(0).BorderStyle = lngGridlines
    lblHeading(1).BorderStyle = lngGridlines
  Case 2
    lblData(0).BorderStyle = lngGridlines
    lblData(1).BorderStyle = lngGridlines
    lblData(2).BorderStyle = lngGridlines
    lblData(3).BorderStyle = lngGridlines
  End Select

  Changed = Not mbLoading

End Sub

' Setup the listview for the toolbar options
Private Sub LoadToolBarOptions(pbReadDefaultValues As Boolean)

  Dim frmForm As Form
  Dim objActiveBar As ActiveBarLibraryCtl.ActiveBar
  Dim objTool As New ActiveBarLibraryCtl.Tool
  Dim iBandCount As Integer
  Dim iToolCount As Integer
  Dim strKey As String
  Dim strName As String
  Dim strCaption As String
  Dim bEnabled As Boolean
  Dim strToolID As String
  Dim lngToolID As Long
  Dim strBandName As String
  Dim strToolbarToLoad As String
  Dim strActiveBarToLoad As String
  Dim objListView As ListView
  Dim iDisplayListViewID As Integer
  Dim bAddThisTool As Boolean

  ' Trick the activebars to read default values of not
  gbReadToolbarDefaults = pbReadDefaultValues

  ' Load tools from the appropriate screen into local activebar control.
  Select Case cboToolbars.Text
    
    Case "Record Editing"
      strToolbarToLoad = mstrRECORDEDITBAND
      Set frmForm = New DataMgr.frmRecEdit4
      strActiveBarToLoad = frmForm.ActiveBar1.Tag   'Triggers form_load and the reshuffle event
      Set objActiveBar = frmForm.ActiveBar1
      
    Case "Find Window"
      strToolbarToLoad = mstrFINDWINDOWBAND
      Set frmForm = New DataMgr.frmFind2
      strActiveBarToLoad = frmForm.ActiveBar1.Tag   'Triggers form_load and the reshuffle event
      Set objActiveBar = frmForm.ActiveBar1
  
  End Select

  ' Reset reading default values
  gbReadToolbarDefaults = False

  ' Has a listview for this control already been loaded
  iDisplayListViewID = 0
  For Each objListView In lvwToolbars
    objListView.Visible = False
    iDisplayListViewID = IIf(objListView.Tag = strActiveBarToLoad & "%%" & strToolbarToLoad, objListView.Index, iDisplayListViewID)
  Next objListView
  
  ' Are we re-reading default values
  If iDisplayListViewID > 0 And pbReadDefaultValues Then
    Unload lvwToolbars(iDisplayListViewID)
    Unload ilstToolbar(iDisplayListViewID)
    iDisplayListViewID = 0
  End If
  
  If iDisplayListViewID > 0 Then
    ' Show required listview
    lvwToolbars(iDisplayListViewID).Visible = True
    DoEvents
  Else
    
    ' Load new listview and imagelists
    iDisplayListViewID = lvwToolbars.UBound + 1
    Load lvwToolbars(iDisplayListViewID)
    lvwToolbars(iDisplayListViewID).TabIndex = lvwToolbars(0).TabIndex 'MH20030908
    Load ilstToolbar(iDisplayListViewID)

    ' Set the template for the listview
    lvwToolbars(iDisplayListViewID).ColumnHeaders.Clear
    lvwToolbars(iDisplayListViewID).ColumnHeaders.Add , "name", "Button Name", 2200
    lvwToolbars(iDisplayListViewID).ColumnHeaders.Add , "enable", "Enabled/Disabled", 500
    lvwToolbars(iDisplayListViewID).ColumnHeaders.Add , "description", "Description", 0
    lvwToolbars(iDisplayListViewID).ColumnHeaders.Add , "toolid", "ToolID", 0
    lvwToolbars(iDisplayListViewID).View = lvwReport
    lvwToolbars(iDisplayListViewID).Tag = strActiveBarToLoad & "%%" & strToolbarToLoad

    ' Add images to imagelist
    For iBandCount = 0 To objActiveBar.Bands.Count - 1
      For iToolCount = 0 To objActiveBar.Bands(iBandCount).Tools.Count - 1
        strKey = "B" & Trim(Str(iBandCount)) & "T" & Trim(Str(iToolCount))
        ilstToolbar(iDisplayListViewID).ListImages.Add , strKey, objActiveBar.Bands(iBandCount).Tools(iToolCount).GetPicture(0)
      Next iToolCount
    Next iBandCount
    
    ' Attach imagelist to listview
    lvwToolbars(iDisplayListViewID).Icons = ilstToolbar(iDisplayListViewID)
    lvwToolbars(iDisplayListViewID).SmallIcons = ilstToolbar(iDisplayListViewID)
          
    ' Load the icons into the listview
    For iBandCount = 0 To objActiveBar.Bands.Count - 1
      ' Force it just to look at the record band
      If objActiveBar.Bands(iBandCount).Name = strToolbarToLoad Then
        For iToolCount = 0 To objActiveBar.Bands(iBandCount).Tools.Count - 1
          
          ' Only load tools if we have the correct license
          Select Case objActiveBar.Bands(iBandCount).Tools(iToolCount).Tag
            Case 0  ' Generic
              bAddThisTool = True
            Case 1  'Personnel module
              bAddThisTool = gfPersonnelEnabled
            Case 2  'Absence module
              bAddThisTool = gfAbsenceEnabled
            Case 3  'Training module
              bAddThisTool = gfTrainingBookingEnabled
          End Select
          
          If bAddThisTool Then
            strToolID = objActiveBar.Bands(iBandCount).Tools(iToolCount).ToolID
            strKey = "B" & Trim(Str(iBandCount)) & "T" & Trim(Str(iToolCount))
            strCaption = Replace(objActiveBar.Bands(iBandCount).Tools(iToolCount).Caption, "&", "")
            bEnabled = objActiveBar.Bands(iBandCount).Tools(iToolCount).Visible
            lvwToolbars(iDisplayListViewID).ListItems.Add , strKey, strCaption, strKey, strKey
            lvwToolbars(iDisplayListViewID).ListItems.Item(strKey).SubItems(1) = IIf(bEnabled, "", "Hidden")
            lvwToolbars(iDisplayListViewID).ListItems.Item(strKey).SubItems(2) = objActiveBar.Bands(iBandCount).Tools(iToolCount).ToolTipText
            lvwToolbars(iDisplayListViewID).ListItems.Item(strKey).SubItems(3) = strToolID
          End If
        Next iToolCount
      End If
    Next iBandCount

    lvwToolbars(iDisplayListViewID).Visible = True

  End If
  
  ' Tidy up
  Unload frmForm
  Set frmForm = Nothing
  Set objActiveBar = Nothing

End Sub

Private Sub SaveToolBarOptions()

  Dim iCount As Integer
  Dim objListView As ListView
  Dim strToolID As String

  ' Has a listview for this control already been loaded
  For Each objListView In lvwToolbars
    For iCount = 1 To objListView.ListItems.Count
      strToolID = objListView.Tag & "%%" & objListView.ListItems.Item(iCount).SubItems(3)
      SaveUserSetting "toolbar_showtool", strToolID, IIf(objListView.ListItems.Item(iCount).SubItems(1) = "Hidden", False, True)
      SaveUserSetting "toolbar_order", objListView.Tag & "%%" & Trim(Str(iCount - 1)), objListView.ListItems.Item(iCount).SubItems(3)
    Next iCount
  Next objListView

End Sub

Private Sub UpdateToolbarButtonStatus()

  Dim iCurrentListID As Integer
  iCurrentListID = CurrentToolBarIndex

  cmdShowHide.Caption = IIf(lvwToolbars(iCurrentListID).SelectedItem.SubItems(1) = "Hidden", "S&how", "&Hide")
  cmdShowHide.Enabled = lvwToolbars(iCurrentListID).SelectedItem.Selected
  cmdMoveUp.Enabled = lvwToolbars(iCurrentListID).SelectedItem.Selected And lvwToolbars(iCurrentListID).SelectedItem.Index > 1
  cmdMoveDown.Enabled = lvwToolbars(iCurrentListID).SelectedItem.Selected And lvwToolbars(iCurrentListID).SelectedItem.Index < lvwToolbars(iCurrentListID).ListItems.Count

End Sub


Private Function ChangeSelectedToolOrder(pobjListView As ListView, Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)
 
  ' Dimension arrays
  Dim iLoop As Integer, Key() As String, Text() As String, Icon() As Variant, SmallIcon() As Variant
  
  Dim SubItem1() As Variant, SubItem2() As Variant, SubItem3() As Variant
  
  ReDim Key(0), Text(0), Icon(0), SmallIcon(0)
  ReDim SubItem1(0), SubItem2(0), SubItem3(0)
  
  Dim itmX As ListItem
  Dim iCurrentListID As Integer
  
  ' Clear the highlight
  Set pobjListView.DropHighlight = Nothing
  
  ' If drop point is below all other items, then fix the intbeforeindex var
  If intBeforeIndex = 0 Then intBeforeIndex = lvwToolbars(iCurrentListID).ListItems.Count + 1
  
  ' First get all the items that are above the drop point that arent selected
  For iLoop = 1 To (intBeforeIndex - 1)
    If pobjListView.ListItems(iLoop).Selected = False Then
      ReDim Preserve Key(UBound(Key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)

      Key(UBound(Key) - 1) = pobjListView.ListItems(iLoop).Key
      Text(UBound(Text) - 1) = pobjListView.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = pobjListView.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = pobjListView.ListItems(iLoop).SmallIcon
    
      ReDim Preserve SubItem1(UBound(SubItem1) + 1)
      ReDim Preserve SubItem2(UBound(SubItem2) + 1)
      ReDim Preserve SubItem3(UBound(SubItem3) + 1)

      SubItem1(UBound(SubItem1) - 1) = pobjListView.ListItems(iLoop).SubItems(1)
      SubItem2(UBound(SubItem2) - 1) = pobjListView.ListItems(iLoop).SubItems(2)
      SubItem3(UBound(SubItem3) - 1) = pobjListView.ListItems(iLoop).SubItems(3)

    End If
  Next iLoop
  
  ' Now get all the items that are selected
  For iLoop = 1 To pobjListView.ListItems.Count
    If pobjListView.ListItems(iLoop).Selected = True Then
      ReDim Preserve Key(UBound(Key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
      Key(UBound(Key) - 1) = pobjListView.ListItems(iLoop).Key
      Text(UBound(Text) - 1) = pobjListView.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = pobjListView.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = pobjListView.ListItems(iLoop).SmallIcon
    
      ReDim Preserve SubItem1(UBound(SubItem1) + 1)
      ReDim Preserve SubItem2(UBound(SubItem2) + 1)
      ReDim Preserve SubItem3(UBound(SubItem3) + 1)

      SubItem1(UBound(SubItem1) - 1) = pobjListView.ListItems(iLoop).SubItems(1)
      SubItem2(UBound(SubItem2) - 1) = pobjListView.ListItems(iLoop).SubItems(2)
      SubItem3(UBound(SubItem3) - 1) = pobjListView.ListItems(iLoop).SubItems(3)

    End If
  Next iLoop
  
  ' Now get all the items below the drop point that arent selected
  If intBeforeIndex <> 0 Then
    For iLoop = (intBeforeIndex) To pobjListView.ListItems.Count
      If pobjListView.ListItems(iLoop).Selected = False Then
        ReDim Preserve Key(UBound(Key) + 1)
        ReDim Preserve Text(UBound(Text) + 1)
        ReDim Preserve Icon(UBound(Icon) + 1)
        ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
        Key(UBound(Key) - 1) = pobjListView.ListItems(iLoop).Key
        Text(UBound(Text) - 1) = pobjListView.ListItems(iLoop).Text
        Icon(UBound(Icon) - 1) = pobjListView.ListItems(iLoop).Icon
        SmallIcon(UBound(SmallIcon) - 1) = pobjListView.ListItems(iLoop).SmallIcon
      
        ReDim Preserve SubItem1(UBound(SubItem1) + 1)
        ReDim Preserve SubItem2(UBound(SubItem2) + 1)
        ReDim Preserve SubItem3(UBound(SubItem3) + 1)

        SubItem1(UBound(SubItem1) - 1) = pobjListView.ListItems(iLoop).SubItems(1)
        SubItem2(UBound(SubItem2) - 1) = pobjListView.ListItems(iLoop).SubItems(2)
        SubItem3(UBound(SubItem3) - 1) = pobjListView.ListItems(iLoop).SubItems(3)

      End If
    Next iLoop
  End If
  
  ' Clear all items from the listview
  pobjListView.ListItems.Clear
  
  ' Add items in the right order from the array
  For iLoop = LBound(Key) To (UBound(Key) - 1)
    
    Set itmX = pobjListView.ListItems.Add(, Key(iLoop), Text(iLoop), Icon(iLoop), SmallIcon(iLoop))
  
    itmX.SubItems(1) = SubItem1(iLoop)
    itmX.SubItems(2) = SubItem2(iLoop)
    itmX.SubItems(3) = SubItem3(iLoop)
  
  Next iLoop
  
  If mfFromButtons = True Then
    pobjListView.ListItems(intBeforeIndex - 1).Selected = True
  Else
    If intBeforeIndex < pobjListView.ListItems.Count Then pobjListView.ListItems(intBeforeIndex).Selected = True Else pobjListView.ListItems(pobjListView.ListItems.Count).Selected = True
  End If
  
  mfFromButtons = False
  
  Changed = Not mbLoading
  
  UpdateToolbarButtonStatus
  
End Function

' Returns the index of the currently selected toolbar listview
Private Function CurrentToolBarIndex() As Integer

  Dim objListView As ListView
  Dim iDisplayListViewID As Integer

  ' Has a listview for this control already been loaded
  iDisplayListViewID = 0
  For Each objListView In lvwToolbars
    iDisplayListViewID = IIf(objListView.Visible, objListView.Index, iDisplayListViewID)
  Next objListView

  CurrentToolBarIndex = iDisplayListViewID

End Function


Private Function GetOutputFormats(strDestin As String, cboTemp As ComboBox, intOfficeVersion As Integer) As Long

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strOutput As String
  Dim lngDefault As Long
  Dim strFormatField As String
  
  On Local Error GoTo LocalErr
  
  strFormatField = IIf(intOfficeVersion < 12, "Office2003", "Office2007")
  
  strSQL = "SELECT * FROM ASRSysFileFormats " & _
           "WHERE Destination = '" & Replace(strDestin, "'", "''") & "' " & _
           "  AND  NOT " & strFormatField & " IS NULL " & _
           "ORDER BY ID"
  Set rsTemp = datGeneral.GetRecords(strSQL)
    
  lngDefault = 0
  With cboTemp
    .Clear
  
    Do While Not rsTemp.EOF
      .AddItem rsTemp.Fields("Description").Value
      .ItemData(.NewIndex) = rsTemp.Fields(strFormatField).Value
      
      If rsTemp.Fields("Default").Value = True Then
        lngDefault = rsTemp.Fields(strFormatField).Value
      End If
      
      rsTemp.MoveNext
    Loop
  
    If .ListCount > 0 Then
      .ListIndex = 0
    End If

  End With
  
  GetOutputFormats = lngDefault
   
LocalErr:
  If Not rsTemp Is Nothing Then
    If rsTemp.State <> adStateClosed Then
      rsTemp.Close
    End If
    Set rsTemp = Nothing
  End If
    
End Function


Public Function GetUserSettingOrDefault(strSection As String, strKey As String, varDefault As Variant) As Variant
  If mblnRestoringDefaults Then
    GetUserSettingOrDefault = varDefault
  Else
    GetUserSettingOrDefault = GetUserSetting(strSection, strKey, varDefault)
  End If
End Function

Public Function BrowseFolders(Optional odtvTitle As String) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = odtvTitle
    With tBrowseInfo
           .hwndOwner = Me.hWnd
           .lpszTitle = lstrcat(szTitle, "")
           .ulFlags = BIF_RETURNONLYFSDIRS
        End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseFolders = sBuffer
    End If
End Function

