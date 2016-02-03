VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmMailMerge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail Merge Definition"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1046
   Icon            =   "frmMailMerge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDocument 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   6630
      Picture         =   "frmMailMerge.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   86
      Top             =   5055
      Visible         =   0   'False
      Width           =   525
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   75
      TabIndex        =   84
      Top             =   90
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   4
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
      TabPicture(0)   =   "frmMailMerge.frx":08D6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinition(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDefinition(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Colu&mns"
      TabPicture(1)   =   "frmMailMerge.frx":08F2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraColumns(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraColumns(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraColumns(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Sort Order"
      TabPicture(2)   =   "frmMailMerge.frx":090E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSort(0)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ou&tput"
      TabPicture(3)   =   "frmMailMerge.frx":092A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraOutputFormat"
      Tab(3).Control(1)=   "fraOutputOptions"
      Tab(3).Control(2)=   "fraOutput(2)"
      Tab(3).Control(3)=   "fraOutput(0)"
      Tab(3).Control(4)=   "fraOutput(1)"
      Tab(3).ControlCount=   5
      Begin VB.Frame fraDefinition 
         Height          =   2355
         Index           =   0
         Left            =   135
         TabIndex        =   91
         Top             =   405
         Width           =   9600
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6030
            MaxLength       =   30
            TabIndex        =   4
            Top             =   300
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   1
            Top             =   300
            Width           =   3090
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1620
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1110
            Width           =   3090
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   3090
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1485
            Left            =   6030
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
            stylesets(0).Picture=   "frmMailMerge.frx":0946
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
            stylesets(1).Picture=   "frmMailMerge.frx":0962
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
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Left            =   5175
            TabIndex        =   96
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   95
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   195
            TabIndex        =   94
            Top             =   1155
            Width           =   1080
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   5175
            TabIndex        =   93
            Top             =   765
            Width           =   825
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category :"
            Height          =   240
            Left            =   195
            TabIndex        =   92
            Top             =   765
            Width           =   1005
         End
      End
      Begin VB.Frame fraOutputFormat 
         Caption         =   "Output Format :"
         Height          =   3330
         Left            =   -74880
         TabIndex        =   50
         Top             =   1440
         Width           =   2745
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   51
            Top             =   400
            Width           =   2400
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Indi&vidual Emails"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   52
            Top             =   800
            Width           =   2400
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Document Mana&gement"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   53
            Top             =   1200
            Width           =   2400
         End
      End
      Begin VB.Frame fraOutputOptions 
         Caption         =   "Options :"
         Height          =   1005
         Left            =   -74880
         TabIndex        =   44
         Top             =   360
         Width           =   9600
         Begin VB.CheckBox chkSuppressBlank 
            Caption         =   "S&uppress blank lines"
            Height          =   195
            Left            =   7080
            TabIndex        =   49
            Top             =   600
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkPauseBeforeMerge 
            Caption         =   "Pause &before merge"
            Height          =   195
            Left            =   7080
            TabIndex        =   48
            Top             =   300
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "<None>"
            Top             =   315
            Width           =   5245
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   6555
            Picture         =   "frmMailMerge.frx":097E
            TabIndex        =   47
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.CommandButton cmdLabelType 
            Caption         =   "..."
            DisabledPicture =   "frmMailMerge.frx":09F6
            Height          =   315
            Left            =   6555
            TabIndex        =   85
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Label lblPrimary 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Template :"
            Height          =   195
            Left            =   225
            TabIndex        =   45
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame fraSort 
         Caption         =   "Sort Order :"
         Height          =   4365
         Index           =   0
         Left            =   -74865
         TabIndex        =   36
         Top             =   405
         Width           =   9600
         Begin VB.CommandButton cmdClearOrder 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8000
            TabIndex        =   41
            Top             =   1935
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteOrder 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8000
            TabIndex        =   40
            Top             =   1395
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditOrder 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   8000
            TabIndex        =   39
            Top             =   840
            Width           =   1200
         End
         Begin VB.CommandButton cmdNewOrder 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   8000
            TabIndex        =   38
            Top             =   315
            Width           =   1200
         End
         Begin VB.CommandButton cmdSortMoveUp 
            Caption         =   "Move &Up"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8000
            TabIndex        =   42
            Top             =   3285
            Width           =   1200
         End
         Begin VB.CommandButton cmdSortMoveDown 
            Caption         =   "Move Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8000
            TabIndex        =   43
            Top             =   3810
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdReportOrder 
            Height          =   3885
            Left            =   180
            TabIndex        =   37
            Top             =   315
            Width           =   7410
            ScrollBars      =   0
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            stylesets.count =   5
            stylesets(0).Name=   "ssetHeaderDisabled"
            stylesets(0).ForeColor=   -2147483631
            stylesets(0).BackColor=   -2147483633
            stylesets(0).Picture=   "frmMailMerge.frx":0D57
            stylesets(1).Name=   "ssetSelected"
            stylesets(1).ForeColor=   -2147483634
            stylesets(1).BackColor=   -2147483635
            stylesets(1).Picture=   "frmMailMerge.frx":0D73
            stylesets(2).Name=   "ssetEnabled"
            stylesets(2).ForeColor=   -2147483640
            stylesets(2).BackColor=   -2147483643
            stylesets(2).Picture=   "frmMailMerge.frx":0D8F
            stylesets(3).Name=   "ssetHeaderEnabled"
            stylesets(3).ForeColor=   -2147483630
            stylesets(3).BackColor=   -2147483633
            stylesets(3).Picture=   "frmMailMerge.frx":0DAB
            stylesets(4).Name=   "ssetDisabled"
            stylesets(4).ForeColor=   -2147483631
            stylesets(4).BackColor=   -2147483633
            stylesets(4).Picture=   "frmMailMerge.frx":0DC7
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
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   0
            StyleSet        =   "ssetDisabled"
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ActiveRowStyleSet=   "ssetSelected"
            Columns.Count   =   3
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ColExprID"
            Columns(0).Name =   "ColExprID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   10848
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
            _ExtentX        =   13070
            _ExtentY        =   6853
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
      Begin VB.Frame fraColumns 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4365
         Index           =   2
         Left            =   -71400
         TabIndex        =   81
         Top             =   405
         Width           =   2265
         Begin VB.CommandButton cmdAddHeading 
            Caption         =   "Add &Heading"
            Height          =   405
            Left            =   540
            TabIndex        =   23
            Top             =   1080
            Width           =   1395
         End
         Begin VB.CommandButton cmdAddSeparator 
            Caption         =   "Add Se&parator"
            Height          =   405
            Left            =   540
            TabIndex        =   24
            Top             =   1560
            Width           =   1395
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   540
            TabIndex        =   27
            Top             =   3465
            Width           =   1395
         End
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   540
            TabIndex        =   28
            Top             =   3930
            Width           =   1395
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   405
            Left            =   540
            TabIndex        =   21
            Top             =   90
            Width           =   1395
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "R&emove"
            Height          =   405
            Left            =   540
            TabIndex        =   25
            Top             =   2310
            Width           =   1395
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add A&ll"
            Height          =   405
            Left            =   540
            TabIndex        =   22
            Top             =   570
            Width           =   1395
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remo&ve All"
            Height          =   405
            Left            =   540
            TabIndex        =   26
            Top             =   2775
            Width           =   1395
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns / Calculations Selected :"
         Height          =   4350
         Index           =   1
         Left            =   -68850
         TabIndex        =   35
         Top             =   405
         Width           =   3570
         Begin VB.Frame fraSizeDecimals 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   735
            Left            =   240
            TabIndex        =   88
            Top             =   3140
            Width           =   2055
            Begin COASpinner.COA_Spinner spnSize 
               Height          =   300
               Left            =   945
               TabIndex        =   31
               Top             =   0
               Width           =   1000
               _ExtentX        =   1773
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
               MaximumValue    =   2147483647
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnDec 
               Height          =   300
               Left            =   945
               TabIndex        =   32
               Top             =   405
               Width           =   1000
               _ExtentX        =   1773
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
               MaximumValue    =   999
               Text            =   "0"
            End
            Begin VB.Label lblProp_Size 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Size :"
               Height          =   195
               Left            =   0
               TabIndex        =   90
               Top             =   60
               Width           =   615
            End
            Begin VB.Label lblProp_Decimals 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Decimals :"
               Height          =   195
               Left            =   0
               TabIndex        =   89
               Top             =   465
               Width           =   945
            End
         End
         Begin VB.TextBox txtProp_ColumnHeading 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1185
            MaxLength       =   50
            TabIndex        =   30
            Top             =   2740
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.CheckBox chkStartColumnOnNewLine 
            Caption         =   "Sta&rt column on new line"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   3960
            Width           =   2940
         End
         Begin ComctlLib.ListView ListView2 
            Height          =   2355
            Left            =   180
            TabIndex        =   29
            Top             =   300
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   4154
            SortKey         =   1
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Column"
               Object.Tag             =   "Column"
               Text            =   "Column"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "SortKey"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblProp_ColumnHeading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Heading :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   2805
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns / Calculations Available :"
         Height          =   4350
         Index           =   0
         Left            =   -74865
         TabIndex        =   34
         Top             =   405
         Width           =   3360
         Begin ComctlLib.ListView ListView1 
            Height          =   2745
            Left            =   180
            TabIndex        =   19
            Top             =   1050
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   4842
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   0
            EndProperty
         End
         Begin VB.OptionButton optCalc 
            Caption         =   "Calculat&ions"
            Height          =   255
            Left            =   1750
            TabIndex        =   18
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optColumns 
            Caption         =   "Colum&ns"
            Height          =   255
            Left            =   390
            TabIndex        =   17
            Top             =   720
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.CommandButton cmdCalculations 
            Caption         =   "Calculation De&finitions..."
            Height          =   390
            Left            =   180
            TabIndex        =   20
            Top             =   3780
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.ComboBox cboTblAvailable 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   300
            Width           =   3000
         End
      End
      Begin VB.Frame fraDefinition 
         Caption         =   "Data :"
         Height          =   1935
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   2835
         Width           =   9600
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1080
            Width           =   1965
         End
         Begin VB.TextBox txtPicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   705
            Width           =   1965
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   6135
            TabIndex        =   13
            Top             =   1120
            Width           =   840
         End
         Begin VB.OptionButton optPicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   6135
            TabIndex        =   10
            Top             =   750
            Width           =   975
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   6135
            TabIndex        =   9
            Top             =   365
            Value           =   -1  'True
            Width           =   540
         End
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   315
            Width           =   3090
         End
         Begin VB.CommandButton cmdPicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   9135
            TabIndex        =   12
            Top             =   705
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   9135
            TabIndex        =   15
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   6
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Index           =   3
            Left            =   5190
            TabIndex        =   8
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Document Management :"
         Height          =   3330
         Index           =   2
         Left            =   -72015
         TabIndex        =   71
         Top             =   1440
         Width           =   6735
         Begin VB.CheckBox chkDocManManualHeader 
            Caption         =   "Manual document &header"
            Height          =   195
            Left            =   225
            TabIndex        =   77
            Top             =   1585
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.CheckBox chkDocManScreen 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Left            =   225
            TabIndex        =   78
            Top             =   405
            Width           =   2685
         End
         Begin VB.TextBox txtDocumentMap 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   75
            Text            =   "<None>"
            Top             =   1135
            Visible         =   0   'False
            Width           =   4355
         End
         Begin VB.CommandButton cmdDocumentMap 
            Caption         =   "..."
            DisabledPicture =   "frmMailMerge.frx":0DE3
            Height          =   315
            Left            =   6240
            TabIndex        =   76
            Top             =   1135
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.ComboBox cboDocManEngine 
            Height          =   315
            Left            =   1885
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   755
            Width           =   4690
         End
         Begin VB.Label lblDocumentMap 
            Caption         =   "Document Type : "
            Height          =   285
            Left            =   225
            TabIndex        =   74
            Top             =   1180
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblDocManEngine 
            AutoSize        =   -1  'True
            Caption         =   "Engine :"
            Height          =   195
            Left            =   225
            TabIndex        =   72
            Top             =   800
            Width           =   840
         End
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Word Document :"
         Height          =   3330
         Index           =   0
         Left            =   -72015
         TabIndex        =   54
         Top             =   1440
         Width           =   6735
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   59
            Top             =   1335
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   56
            Top             =   825
            Width           =   1605
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   375
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6245
            TabIndex        =   62
            Top             =   1275
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3495
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   765
            Width           =   3090
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1275
            Width           =   2750
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2025
            TabIndex        =   57
            Top             =   825
            Width           =   1590
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2025
            TabIndex        =   60
            Top             =   1335
            Width           =   1185
         End
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Individual Emails :"
         Height          =   3330
         Index           =   1
         Left            =   -72015
         TabIndex        =   63
         Top             =   1440
         Width           =   6735
         Begin VB.TextBox txtEmailAttachmentName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1875
            MaxLength       =   255
            TabIndex        =   70
            Top             =   1515
            Width           =   4710
         End
         Begin VB.TextBox txtEmailSubject 
            Height          =   315
            Left            =   1875
            MaxLength       =   255
            TabIndex        =   67
            Top             =   715
            Width           =   4710
         End
         Begin VB.ComboBox cboEMailField 
            Height          =   315
            ItemData        =   "frmMailMerge.frx":1144
            Left            =   1875
            List            =   "frmMailMerge.frx":1146
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   315
            Width           =   4710
         End
         Begin VB.CheckBox chkEMailAttachment 
            Caption         =   "Se&nd as attachment"
            Height          =   195
            Left            =   240
            TabIndex        =   68
            Top             =   1165
            Width           =   2115
         End
         Begin VB.Label lblEmailAttachAs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   525
            TabIndex        =   69
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblEMailSubject 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subject :"
            Height          =   195
            Left            =   225
            TabIndex        =   66
            Top             =   765
            Width           =   915
         End
         Begin VB.Label lblEMailField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address :"
            Height          =   195
            Left            =   240
            TabIndex        =   64
            Top             =   375
            Width           =   1410
         End
      End
   End
   Begin VB.PictureBox picDocument 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   5550
      Picture         =   "frmMailMerge.frx":1148
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   83
      Top             =   5055
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picNoDrop 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4515
      Picture         =   "frmMailMerge.frx":16D2
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   82
      Top             =   5070
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   6135
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   7440
      TabIndex        =   79
      Top             =   5100
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8775
      TabIndex        =   80
      Top             =   5100
      Width           =   1200
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   5040
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
            Picture         =   "frmMailMerge.frx":1F9C
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMailMerge.frx":22EE
            Key             =   "IMG_CALC"
         EndProperty
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   3840
      Top             =   5040
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
      Bands           =   "frmMailMerge.frx":2840
   End
End
Attribute VB_Name = "frmMailMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Enum OutputType
'  Document = 0
'  Printer = 1
'  Email = 2
'  Version1 = 3
'  HRProDocumentManagement = 4
'End Enum

Private Const mstrSQLTableDef As String = "ASRSysMailMergeName"
Private Const mstrSQLTableCol As String = "ASRSysMailMergeColumns"

Private Const sDFLTTEXT_HEADING = "<Heading>"
Private Const sDFLTTEXT_SEPARATOR = "<Separator>"
Private Const sTYPECODE_HEADING = "H"
Private Const sTYPECODE_SEPARATOR = "S"
Private Const sTYPECODE_COLUMN = "C"
Private Const sTYPECODE_EXPRESSION = "E"

Private fOK As Boolean
Private rsTables As New ADODB.Recordset
Private datData As DataMgr.clsDataAccess          'DataAccess Class
Private mlngTimeStamp As Long

Private mlngMailMergeID As Long
Private mstrPrimaryTable As String
Private mblnLoading As Boolean
Private mblnReadOnly As Boolean
Private mblnFromCopy As Boolean
Private mblnFormPrint As Boolean
Private mblnDefinitionCreator As Boolean

Private mlngColumnDragIndex As Long
Private mblnHiddenPicklistOrFilter As Boolean
Private msHiddenCalcsSelected As String
Private miHiddenCalcs As Integer
Private mblnRecordSelectionInvalid As Boolean
Private mblnDeletedCalc As Boolean
Private mbNeedsSave As Boolean
Private mblnCancelled As Boolean
Private mblnEmailPermission As Boolean
Private mblnWarnedNoEmail As Boolean  'Only need this message once.

Private mblnForceHidden As Boolean

' Flag for whether we are currently in a drag operation
Private mfColumnDrag As Boolean

' Flag for whether this mail merge is a labels/envelope setting
Private mbIsLabel As Boolean

' Holds the ID for currently selected label template
Private mlngLabelTypeID As Long
Private miNumberOfRowsInLabel As Integer




Private Sub cboDocManEngine_Click()
 Me.Changed = True
End Sub

Private Sub cboPrinterName_Click()
  Me.Changed = True
End Sub

Private Sub chkCloseAfterDocManInsert_Click()
  Me.Changed = True
End Sub

Private Sub chkDestination_Click(Index As Integer)
  Select Case Index
  Case 1
    lblPrinter.Enabled = (chkDestination(1).Value = vbChecked)
    cboPrinterName.Enabled = (chkDestination(1).Value = vbChecked)
    cboPrinterName.ListIndex = IIf(chkDestination(1).Value = vbChecked, 0, -1)
    cboPrinterName.BackColor = IIf(chkDestination(1).Value = vbChecked, vbWindowBackground, vbButtonFace)
  Case 2
    lblFileName.Enabled = (chkDestination(2).Value = vbChecked)
    cmdFilename(0).Enabled = (chkDestination(2).Value = vbChecked)
    If chkDestination(2).Value = vbUnchecked Then
      txtFilename(0).Text = vbNullString
    End If
  End Select
  Me.Changed = True
End Sub

Private Sub chkManualDocManHeader_Click()
  Me.Changed = True
End Sub

Private Sub chkDocManManualHeader_Click()
 Me.Changed = True
End Sub

Private Sub chkDocManScreen_Click()
 Me.Changed = True
End Sub

Private Sub cmdDocumentMap_Click()

  Dim frmDefinition As frmDocumentMap
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean
  Dim blnOK As Boolean
  Dim strSelectedName As String
  Dim lngSelectedID As Long

  Set frmSelection = New frmDefSel
  blnExit = False
   
  With frmSelection
    Do While Not blnExit
        
      .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtSelect + edtPrint + edtDeselect + edtProperties
      'mlngOptions = edtAdd + edtDelete + edtEdit + edtCopy + edtSelect + edtPrint + edtDeselect + edtProperties
      
      .EnableRun = False
            
      If mlngLabelTypeID > 0 Then
        .SelectedID = mlngLabelTypeID
      End If
      
      If .ShowList(utlDocumentMapping) Then
        
              
        .Show vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmDocumentMap
          frmDefinition.Initialise True, .FromCopy, , False
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtEdit
          Set frmDefinition = New frmDocumentMap
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
           
        Case edtSelect
          lngSelectedID = .SelectedID
          strSelectedName = .SelectedText
          
          txtDocumentMap.Text = strSelectedName
          txtDocumentMap.Tag = lngSelectedID
                    
          blnExit = True
          Me.Changed = True

        Case edtPrint
          Set frmDefinition = New frmDocumentMap
          frmDefinition.PrintDefinition .SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case edtDeselect
          txtDocumentMap.Text = vbNullString
          blnExit = True
          Me.Changed = True
          
        Case 0
          blnExit = True  'cancel

        End Select
      
        ' Store the ID of selected label type
        mlngLabelTypeID = .SelectedID
      
      End If
    
    Loop
  End With

  UpdateButtonStatus

  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub grdReportOrder_Change()
  Changed = True
End Sub

Private Sub grdReportOrder_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  
  With grdReportOrder

    ' Set the styleSet of the rows to show which is selected.
    For iLoop = 0 To .Rows - 1
      If (mblnReadOnly) Then
        .Columns(0).CellStyleSet "ssetDisabled", iLoop
        .Columns(1).CellStyleSet "ssetDisabled", iLoop
        .Columns(2).CellStyleSet "ssetDisabled", iLoop
      Else
        If iLoop = .Row Then
          .Columns(0).CellStyleSet "ssetSelected", iLoop
          .Columns(1).CellStyleSet "ssetSelected", iLoop
          .Columns(2).CellStyleSet "ssetSelected", iLoop
        Else
          .Columns(0).CellStyleSet "ssetEnabled", iLoop
          .Columns(1).CellStyleSet "ssetEnabled", iLoop
          .Columns(2).CellStyleSet "ssetEnabled", iLoop
        End If
      End If
    Next iLoop
    
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
  
    If (mblnReadOnly) Then
      Me.cmdSortMoveUp.Enabled = False
      Me.cmdSortMoveDown.Enabled = False
    Else
      If .AddItemRowIndex(.Bookmark) = 0 Then
        Me.cmdSortMoveUp.Enabled = False
        Me.cmdSortMoveDown.Enabled = .Rows > 1
      ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
        Me.cmdSortMoveUp.Enabled = .Rows > 1
        Me.cmdSortMoveDown.Enabled = False
      Else
        Me.cmdSortMoveUp.Enabled = .Rows > 1
        Me.cmdSortMoveDown.Enabled = .Rows > 1
      End If
    End If
  End With

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub grdReportOrder_RowLoaded(ByVal Bookmark As Variant)
  
  With grdReportOrder

    If (mblnReadOnly) Then
      .Columns(0).CellStyleSet "ssetDisabled"
      .Columns(1).CellStyleSet "ssetDisabled"
      .Columns(2).CellStyleSet "ssetDisabled"
    Else
      .Columns(0).CellStyleSet "ssetEnabled"
      .Columns(1).CellStyleSet "ssetEnabled"
      .Columns(2).CellStyleSet "ssetEnabled"
    End If
   
  End With

End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property

Private Function IsDefinitionCreator(lngID As Long) As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = vbNullString
  sSQL = "SELECT * FROM ASRSysMailMergeName WHERE MailMergeID = " & lngID
  
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.RecordCount > 0 Then
    IsDefinitionCreator = (LCase(rsTemp!userName) = LCase(gsUserName))
  Else
    IsDefinitionCreator = False
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing
  
End Function

Private Function MyPointsToCentimeters(ByVal psngInput As Single) As Single

  MyPointsToCentimeters = psngInput * 0.03527778

End Function

Public Property Get SelectedID() As Long
  SelectedID = mlngMailMergeID
End Property

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
  Call PrimaryTableClick
End Sub

Private Function ErrorCOAMsgBox(strMessage As String) As Boolean
  
  Dim sCaption As String
  sCaption = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge")
  
  Screen.MousePointer = vbDefault
  COAMsgBox strMessage & vbCrLf & Err.Description, vbCritical, sCaption

End Function

Private Function WarningCOAMsgBox(strMessage As String) As Boolean
  
  Dim sCaption As String
  sCaption = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge")
  
  If Screen.MousePointer = vbHourglass Then
    Screen.MousePointer = vbDefault
    COAMsgBox strMessage, vbExclamation, sCaption
    Screen.MousePointer = vbHourglass
  Else
    COAMsgBox strMessage, vbExclamation, sCaption
  End If
End Function

Private Sub cboEMailField_Click()
  Me.Changed = True
End Sub

Private Sub cboTblAvailable_Click()
  Call PopulateAvailable
  Call UpdateButtonStatus
End Sub


Private Sub chkEMailAttachment_Click()

'  If mblnEmailPermission Then
    lblEmailAttachAs.Enabled = chkEMailAttachment.Value
    With txtEmailAttachmentName
      .Enabled = (chkEMailAttachment.Value)
      .BackColor = IIf(chkEMailAttachment.Value, vbWindowBackground, vbButtonFace)
    
      If chkEMailAttachment.Value Then
        'Get everything to the right of the last "\" in template name
        .Text = Mid(txtFilename(1), InStrRev(txtFilename(1), "\") + 1)
        'NHRD31082004 Fault 6653
        If .Text = "<None>" Then .Text = ""
      Else
        .Text = vbNullString
      End If
  
    End With
'  Else
'    With txtEmailAttachmentName
'      .Enabled = False
'      .BackColor = vbButtonFace
'    End With
'  End If
  
  Me.Changed = True
End Sub

Private Sub chkStartColumnOnNewLine_Click()

  Dim lst As ListItem

  For Each lst In ListView2.ListItems
    If lst.Selected Then
      If Not (lst.SubItems(5) = IIf(chkStartColumnOnNewLine.Value = 1, True, False)) Then
        lst.SubItems(5) = IIf(chkStartColumnOnNewLine.Value = 1, True, False)
        Me.Changed = True
      End If
    End If
  Next

End Sub

Private Sub chkPauseBeforeMerge_Click()
  Me.Changed = True
End Sub

Private Sub chkSuppressBlank_Click()
  Me.Changed = True
End Sub

Private Sub cmdAdd_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAdd.Picture = cmdAdd.Picture

End Sub

Private Sub cmdAddAll_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAddAll.Picture = cmdAddAll.Picture

End Sub

Private Sub cmdAddHeading_Click()

  Dim sKey As String
  Dim strText As String
  Dim lngID As Long
  Dim itmX As ListItem
  
  For Each itmX In ListView2.ListItems
    itmX.Selected = False
  Next
  ListView2.SelectedItem = Nothing
  ListView2.Refresh
  
  sKey = UniqueKey(sTYPECODE_HEADING)
  strText = sDFLTTEXT_HEADING
  
  Set itmX = ListView2.ListItems.Add(, sKey, strText) ', ListView1.ListItems(0).Icon, ListView1.ListItems(0).SmallIcon)
  itmX.SubItems(1) = sTYPECODE_HEADING
  itmX.SubItems(2) = 0        ' Size
  itmX.SubItems(3) = 0        ' Decimals
  itmX.SubItems(4) = False    ' Is Numeric
  itmX.SubItems(5) = True     ' Start on new line
  itmX.SubItems(6) = ""
  itmX.Selected = True
  
  UpdateButtonStatus
  Screen.MousePointer = vbDefault
  Changed = True

End Sub

Private Sub cmdAddHeading_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAddHeading.Picture = cmdAddHeading.Picture
End Sub


Private Sub cmdAddSeparator_Click()

  Dim sKey As String
  Dim strText As String
  Dim lngID As Long
  Dim itmX As ListItem
  
  For Each itmX In ListView2.ListItems
    itmX.Selected = False
  Next
  ListView2.SelectedItem = Nothing
  ListView2.Refresh
  
  sKey = UniqueKey(sTYPECODE_SEPARATOR)
  strText = sDFLTTEXT_SEPARATOR
  
  Set itmX = ListView2.ListItems.Add(, sKey, strText) ', ListView1.ListItems(0).Icon, ListView1.ListItems(0).SmallIcon)
  itmX.SubItems(1) = sTYPECODE_SEPARATOR
  itmX.SubItems(2) = 0        ' Size
  itmX.SubItems(3) = 0        ' Decimals
  itmX.SubItems(4) = False    ' Is Numeric
  itmX.SubItems(5) = True    ' Start on new line
  itmX.SubItems(6) = ""
  itmX.Selected = True
  
  UpdateButtonStatus
  Screen.MousePointer = vbDefault
  Changed = True

End Sub

Private Sub cmdAddSeparator_LostFocus()
  'JPD 20031013 Fault 5827
  cmdAddSeparator.Picture = cmdAddSeparator.Picture

End Sub


Private Sub cmdCancel_Click()

  Dim strSQL As String
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  If Me.Changed And Not mblnReadOnly Then
    
    'strMBText = "Mail Merge definition has changed.  Save changes ?"
    strMBText = "You have changed the current definition. Save changes ?"
    intMBButtons = vbQuestion + vbYesNoCancel + vbDefaultButton1
    strMBTitle = Me.Caption
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    
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

Private Sub cmdClearOrder_Click()

  If COAMsgBox("Are you sure you wish to clear the sort order?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    grdReportOrder.RemoveAll
    UpdateOrderButtonStatus
    Me.Changed = True
  End If

End Sub

Private Sub cmdDeleteOrder_Click()
  
  Dim lRow As Long
  
  Me.Changed = True
    
  With grdReportOrder
  
    If .Rows = 1 Then
      .RemoveAll
    Else
      lRow = .AddItemRowIndex(.Bookmark)
      .RemoveItem lRow
      If lRow < .Rows Then
        .Bookmark = lRow
      Else
        .Bookmark = (.Rows - 1)
      End If
      .SelBookmarks.Add .Bookmark
    End If

    'If .Rows = 0 Then
      UpdateOrderButtonStatus
    'End If

  End With

End Sub

Private Sub cmdEditOrder_Click()

  Dim pfrmOrderEdit As New frmMailMergeOrder
  Dim lngColumnID As Long
  Dim strSortOrder As String

  With grdReportOrder
  
    lngColumnID = .Columns("ColExprID").CellValue(.Bookmark)
    strSortOrder = .Columns("Sort Order").CellText(.Bookmark)

    If lngColumnID > 0 Then
      'JPD 20030911 Fault 6359
      'pfrmOrderEdit.Caption = Me.Caption & " Order"
      pfrmOrderEdit.Caption = IIf(mbIsLabel, "Envelope & Label", "Mail Merge") & " Order"
      If pfrmOrderEdit.Initialise(False, Me, lngColumnID, strSortOrder) = True Then
        pfrmOrderEdit.Show vbModal
      End If
    End If
  
  End With
  
  'AE20071025 Fault #6797
  If Not pfrmOrderEdit.UserCancelled Then
    Unload pfrmOrderEdit
    Set pfrmOrderEdit = Nothing
    UpdateOrderButtonStatus
    Me.Changed = True
  End If

End Sub

Private Sub cmdFileName_Click(Index As Integer)

  Dim wrdApp As Word.Application
  Dim wrdDoc As Word.Document
  Dim strFormat As String

  On Local Error GoTo LocalErr

  'With CDialog
  With frmMain.CommonDialog1
    'TM20011031 Fault 3035
'    If Len(txtFilename(Index).Text) > 0 And txtFilename(Index).Text <> "<None>" Then
'      .FileName = txtFilename(Index).Text
'    Else
'      .FileName = vbNullString
'    End If
    If Len(Trim(txtFilename(Index).Text)) = 0 Or txtFilename(Index).Text = "<None>" Then
      .InitDir = gsDocumentsPath
      .FileName = vbNullString
    Else
      .FileName = txtFilename(Index).Text
    End If

    .CancelError = True
    Select Case Index
    Case 0
      .DialogTitle = Me.Caption & " Output Document"
      .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNOverwritePrompt
      InitialiseCommonDialogFormats frmMain.CommonDialog1, "Word", GetOfficeWordVersion, DirectionOutput
      .ShowSave
    Case 1
      'Word template
      .DialogTitle = Me.Caption & " Template"
      .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNCreatePrompt
      .Filter = "Word Template (*.dot;*.dotx;*.doc;*.docx)|*.dot;*.dotx;*.doc;*.docx"
      .ShowOpen
    End Select

    If Len(.FileName) > 256 Then
      WarningCOAMsgBox "Path and file name must not exceed 256 characters in length"
      Exit Sub
    End If

    If .FileName <> "" Then
      If Dir(frmMain.CommonDialog1.FileName) = vbNullString And Index = 1 Then  'Only show for templates
        If COAMsgBox("Template file does not exist.  Create it now?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then

          On Error GoTo WordErr

          txtFilename(Index).Text = frmMain.CommonDialog1.FileName
          strFormat = GetOfficeSaveAsFormat(frmMain.CommonDialog1.FileName, GetOfficeWordVersion, oaWord)
          
          Screen.MousePointer = vbHourglass
          gobjProgress.Caption = "Creating Word Document"
          gobjProgress.MainCaption = Me.Caption
          gobjProgress.AVI = dbWord
          gobjProgress.NumberOfBars = 0
          gobjProgress.Cancel = False
          gobjProgress.OpenProgress

          Set wrdApp = CreateObject("Word.Application")
          Set wrdDoc = wrdApp.Documents.Add
          wrdDoc.SaveAs frmMain.CommonDialog1.FileName, Val(strFormat)
          wrdDoc.Close False
          wrdApp.Quit False
        
          Set wrdDoc = Nothing
          Set wrdApp = Nothing
        
          gobjProgress.CloseProgress
          Screen.MousePointer = vbDefault
        
        End If
      Else
        txtFilename(Index).Text = frmMain.CommonDialog1.FileName

      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    On Local Error Resume Next
    wrdDoc.Close False
    wrdApp.Quit False
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
  
    gobjProgress.CloseProgress
    If Err.Number = 429 Then
      ErrorCOAMsgBox "Error opening Word application"
    Else
      ErrorCOAMsgBox "Error selecting file"
    End If
    txtFilename(Index).Text = vbNullString
  End If

Exit Sub

WordErr:
  gobjProgress.CloseProgress
  Screen.MousePointer = vbDefault
  If Err.Number = 429 Then
    COAMsgBox "Error opening Word application", vbCritical, Me.Caption
  Else
    ErrorCOAMsgBox "Error creating template file"
  End If

End Sub

Private Sub cmdCalculations_Click()

  Dim objExpr As clsExprExpression
  Dim strKey As String
  Dim SelectedKeys() As String
  Dim iCount As Integer
  Dim sMessage As String
  Dim itmX As ListItem
  
  On Error GoTo LocalErr
  
  ReDim SelectedKeys(0)
  Dim lst As ListItem
  For Each lst In ListView1.ListItems
    If lst.Selected = True Then
      ReDim Preserve SelectedKeys(UBound(SelectedKeys) + 1)
      SelectedKeys(UBound(SelectedKeys) - 1) = lst.Key
    End If
  Next lst
  
  
  Set objExpr = New clsExprExpression
  
  With objExpr
    If .Initialise(cboBaseTable.ItemData(cboBaseTable.ListIndex), 0, giEXPR_RUNTIMECALCULATION, 0) Then
      .SelectExpression True
    
      ' Refresh the listview to show the newly added calculation
      Call PopulateAvailable

      If .ExpressionID > 0 Then
      
        '08/08/2000 MH Fault 2664
        If mblnReadOnly Then
          COAMsgBox "Unable to select calculation as you are viewing a read only definition", vbExclamation, Me.Caption
        Else

          strKey = "E" & CStr(.ExpressionID)
          If Not AlreadyUsed(strKey) Then
            sMessage = IsCalcValid(.ExpressionID)
            If sMessage <> vbNullString Then
              COAMsgBox "This calculation has been deleted or hidden by another user." & vbCrLf & _
                     "It cannot be added to this definition", vbExclamation, app.title
            Else
              If optCalc.Value And (cboTblAvailable.ItemData(cboTblAvailable.ListIndex) = .BaseTableID) Then
                ListView1.ListItems(strKey).Selected = True
                Call CopyToSelected(False)

              Else

                'MH20050524 Faults 10089 & 10090
                .ConstructExpression
                .ValidateExpression True, True

                Set itmX = ListView2.ListItems.Add(, strKey, .Name, , ImageList1.ListImages("IMG_CALC").Index)
                itmX.Tag = "*" & .Access
                itmX.SubItems(1) = strKey
                itmX.SubItems(2) = 0
                itmX.SubItems(3) = 0
                itmX.SubItems(4) = (.ReturnType = giEXPRVALUE_NUMERIC Or _
                                   .ReturnType = giEXPRVALUE_BYREF_NUMERIC)
                
                ' Is this column to start on a new line
                itmX.SubItems(5) = True
 
                ' JDM - Fault 6310 - Enable ok button
                If Not mblnLoading Then
                  Me.Changed = True
                End If
 
              End If
            End If
          End If

          ' RH 09/04/01 - leaves 2 things highlighted
          For Each lst In ListView2.ListItems
            If lst.Selected = True Then
              lst.Selected = False
            End If
          Next lst
          ListView2.ListItems(strKey).Selected = True
        
        End If
        
'        ListView2.ListItems(strKey).Selected = True
      End If

      'Call UpdateButtonStatus
      'With ListView1.ListItems
      '  If .Count > 0 Then
      '    .Item(.Count).Selected = True
      '  End If
      'End With

    End If
  End With

  Set objExpr = Nothing

  'also need to refresh selected in case any renamed or deleted

  ' Reselect the cols/calcs that were selected before the calc button was pressed
  For iCount = 0 To (UBound(SelectedKeys) - 1)
    For Each lst In ListView1.ListItems
      If lst.Key = SelectedKeys(iCount) Then
        lst.Selected = True
        Exit For
      End If
    Next lst
  Next iCount
  
  ListView1.SetFocus
  
  'JPD 20030728 Fault 6476
  ForceDefinitionToBeHiddenIfNeeded
  Call UpdateButtonStatus

Exit Sub

LocalErr:
  
  Select Case Err.Number
  
    Case 35601:  ' Calc selected was hidden but user not definition owner
    Case Else: ErrorCOAMsgBox "Error selecting calculations"

  End Select
  
End Sub

Private Sub cmdLabelType_Click()

  Dim frmDefinition As frmLabelTypeDefinition
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean
  Dim blnOK As Boolean
  Dim strSelectedName As String
  Dim lngSelectedID As Long

  Set frmSelection = New frmDefSel
  blnExit = False
   
  With frmSelection
    Do While Not blnExit
      
      .EnableRun = False
      
      If mlngLabelTypeID > 0 Then
        .SelectedID = mlngLabelTypeID
      End If
      
      If .ShowList(utlLabelType) Then
        
        .Show vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmLabelTypeDefinition
          frmDefinition.Initialise True, .FromCopy, , False
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
                    
        Case edtEdit
          Set frmDefinition = New frmLabelTypeDefinition
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing
           
        Case edtSelect
          lngSelectedID = .SelectedID
          strSelectedName = .SelectedText
          
          txtFilename(1).Text = strSelectedName
          txtFilename(1).Tag = mlngLabelTypeID
                    
          blnExit = True
          Me.Changed = True

        Case edtPrint
          Set frmDefinition = New frmLabelTypeDefinition
          frmDefinition.PrintDefinition .SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing
        
        Case 0
          blnExit = True  'cancel

        End Select
      
        ' Store the ID of selected label type
        mlngLabelTypeID = .SelectedID
      
      End If
    
    Loop
  End With

  UpdateButtonStatus

  Unload frmSelection
  Set frmSelection = Nothing

'  If Not mblnLoading Then
'    Me.Changed = True
'  End If
  
End Sub

Private Sub cmdMoveDown_Click()

  ChangeSelectedOrder ListView2.SelectedItem.Index + 2, True

End Sub

Private Sub cmdMoveDown_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveDown.Picture = cmdMoveDown.Picture

End Sub


Private Sub cmdMoveUp_Click()

  ChangeSelectedOrder ListView2.SelectedItem.Index - 1

End Sub

Private Sub cmdMoveUp_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveUp.Picture = cmdMoveUp.Picture

End Sub


Private Sub cmdNewOrder_Click()

  Dim pfrmOrderEdit As New frmMailMergeOrder
  Dim pIntOldRowcount As Integer
  
  'NHRD07102004 Fault 6749
  pIntOldRowcount = grdReportOrder.Rows
  
  pfrmOrderEdit.Caption = IIf(mbIsLabel, "Envelope & Label", "Mail Merge") & " Order"
  If pfrmOrderEdit.Initialise(True, Me, 0, "") = True Then
    pfrmOrderEdit.Show vbModal
  End If
  
  'AE20071025 Fault #6797
  If Not pfrmOrderEdit.UserCancelled Then
    Me.Changed = True
    Unload pfrmOrderEdit
    Set pfrmOrderEdit = Nothing
    UpdateOrderButtonStatus
End If

  'MH20050104 Fault 9522
  'Me.Changed = grdReportOrder.Rows > pIntOldRowcount
  
  'AE20071025 Fault #6797
  'Me.Changed = (Me.Changed Or grdReportOrder.Rows > pIntOldRowcount)

End Sub

Private Sub cmdOK_Click()
  
  Dim fOK As Boolean
  
  If ValidateDefinition = False Then
    Exit Sub
  End If
      
  Screen.MousePointer = vbHourglass
    
  If SaveDefinition Then
    If SaveColumns Then
      Me.Hide
    End If
  End If
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdRemove_LostFocus()
  'JPD 20031013 Fault 5827
  cmdRemove.Picture = cmdRemove.Picture

End Sub

Private Sub cmdRemoveAll_LostFocus()
  'JPD 20031013 Fault 5827
  cmdRemoveAll.Picture = cmdRemoveAll.Picture

End Sub

Private Sub cmdSortMoveDown_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  With grdReportOrder
  
    intSourceRow = .AddItemRowIndex(.Bookmark)
    strSourceRow = .Columns(0).Text & vbTab & _
                   .Columns(1).Text & vbTab & _
                   .Columns(2).Text
  
    intDestinationRow = intSourceRow + 1
    .MoveNext
    strDestinationRow = .Columns(0).Text & vbTab & _
                        .Columns(1).Text & vbTab & _
                        .Columns(2).Text
  
    .RemoveItem intDestinationRow
    .RemoveItem intSourceRow
  
    .AddItem strDestinationRow, intSourceRow
    .AddItem strSourceRow, intDestinationRow
  
    .SelBookmarks.RemoveAll
    .MoveNext
    .Bookmark = .AddItemBookmark(intDestinationRow)
    .SelBookmarks.Add .AddItemBookmark(intDestinationRow)
  
    UpdateOrderButtonStatus

    Me.Changed = True

  End With

End Sub

Private Sub cmdSortMoveUp_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  With grdReportOrder
  
    intSourceRow = .AddItemRowIndex(.Bookmark)
    strSourceRow = .Columns(0).Text & vbTab & _
                   .Columns(1).Text & vbTab & _
                   .Columns(2).Text
  
    intDestinationRow = intSourceRow - 1
    .MovePrevious
    strDestinationRow = .Columns(0).Text & vbTab & _
                        .Columns(1).Text & vbTab & _
                        .Columns(2).Text
  
    .AddItem strSourceRow, intDestinationRow
  
    .RemoveItem intSourceRow + 1
  
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .AddItemBookmark(intDestinationRow)
    .MovePrevious
    .MovePrevious
    UpdateOrderButtonStatus
  
    Me.Changed = True

  End With

End Sub

Private Sub Form_Load()
  
  SSTab1.Tab = 0
  SSTab1_Click (0)
  fraSizeDecimals.BackColor = Me.BackColor
  grdAccess.RowHeight = 239
  mblnEmailPermission = datGeneral.SystemPermission("EMAILADDRESSES", "VIEW")
  mblnWarnedNoEmail = False
  
End Sub


Private Sub cmdPicklist_Click()

  Dim sSQL As String
  Dim lParent As Long
  Dim fExit As Boolean
  Dim frmPick As frmPicklists
  Dim rsTemp As Recordset

  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass

  fExit = False

  'set the sql to only include tables for the selected MailMerge base table
  'sSQL = "Select Name, PickListID From ASRSysPickListName"
  'sSQL = sSQL & " Where TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
  'sSQL = "TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)

  With frmDefSel
    .SelectedUtilityType = utlPicklist
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
          fExit = True
        End Select
      End If

    Loop

  End With
  
  Set frmDefSel = Nothing
  
  ForceDefinitionToBeHiddenIfNeeded
  
Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error selecting picklist"

End Sub


Private Sub cmdFilter_Click()

  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim objExpression As clsExprExpression
  Dim rsTemp As Recordset

  'If cboBaseTable.Text = "<None>" Then
  '  COAMsgBox "No primary table specified", vbExclamation, Me.Caption
  '  Exit Sub
  'End If

  On Error GoTo LocalErr
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression

  With objExpression
    ' Initialise the expression object.
    If .Initialise(cboBaseTable.ItemData(cboBaseTable.ListIndex), Val(txtFilter.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
  
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
  ErrorCOAMsgBox "Error selecting filter"

End Sub



Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional lMailMergeID As Long, Optional bPrint As Boolean) As Boolean

  Dim iUtilityType As utilityType

  On Error GoTo LocalErr

  Set datData = New DataMgr.clsDataAccess

  Screen.MousePointer = vbHourglass

  mblnLoading = True
  
  LoadTableCombo cboBaseTable
    
  iUtilityType = IIf(mbIsLabel, utlLabel, utlMailMerge)
  fOK = True
  
  ' Display/Hide the relevant controls for mail merge/labels
  DisplayLabelSpecifics
  
  optOutputFormat(2).Visible = (IsModuleEnabled(modVersionOne) And Not mbIsLabel)

  ' Populate print combos
  PopulatePrintCombo cboPrinterName
  PopulatePrintCombo cboDocManEngine
  cboDocManEngine.ListIndex = 0
  
  If bNew Then
    
    optOutputFormat(0).Value = True
    LoadPrimaryDependantCombos
    
    ' this MUST be before optallrecordsclick !!! RH 14/06/00
    mblnDefinitionCreator = True
    
    optAllRecords_Click  'Default to all records

    'Set ID to 0 to indicate new record
    mlngMailMergeID = 0
    txtUserName = gsUserName
    mblnDefinitionCreator = True

    PopulateAccessGrid

    GetObjectCategories cboCategory, iUtilityType, 0, cboBaseTable.ItemData(cboBaseTable.ListIndex)
    SetComboItem cboCategory, IIf(glngCurrentCategoryID = -1, 0, glngCurrentCategoryID)

    'Defaults
    chkPauseBeforeMerge.Value = IIf(mbIsLabel, vbUnchecked, vbChecked)

    Me.Changed = False

  Else
    mlngMailMergeID = lMailMergeID
    mblnFromCopy = bCopy
  
    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint
      
    PopulateAccessGrid
    
    Call RetrieveColumns
    Call RetreiveDefinition

    If fOK And Not Me.Cancelled Then
      
      If mblnFromCopy Then
        mlngMailMergeID = 0
        Me.Changed = True
      Else
        Me.Changed = ((mblnRecordSelectionInvalid Or mblnDeletedCalc) And Not mblnReadOnly)
      End If
      
    End If

  End If

  If mbIsLabel Then
    Me.HelpContextID = 1083
  End If

  UpdateButtonStatus

  Screen.MousePointer = vbDefault
  mblnLoading = False
  Initialise = fOK

  Screen.MousePointer = vbDefault

Exit Function

LocalErr:
  ErrorCOAMsgBox "Error with " & Me.Caption & " definition"

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
  If mbIsLabel Then
    Set rsAccess = GetUtilityAccessRecords(utlLabel, mlngMailMergeID, mblnFromCopy)
  Else
    Set rsAccess = GetUtilityAccessRecords(utlMailMerge, mlngMailMergeID, mblnFromCopy)
  End If
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'If form is visible then assume that unload is the same as pressing
  'the cancel button.  Do not unload !!!!!
  If Me.Visible And Not FormPrint Then
    Cancel = True
    Call cmdCancel_Click
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.RefreshMainForm Me, True
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


Private Sub grdReportOrder_DblClick()

  If cmdEditOrder.Enabled Then
    cmdEditOrder_Click
  End If

End Sub



'Private Sub ListView1_GotFocus()
'  Debug.Print "ListView1_GotFocus"
'  cmdAdd.Default = True
'End Sub
'
'Private Sub ListView1_LostFocus()
'  Debug.Print "ListView1_LostFocus"
'  cmdOK.Default = True
'End Sub
'
'Private Sub ListView2_GotFocus()
'  Debug.Print "ListView2_GotFocus"
'  cmdRemove.Default = True
'End Sub
'
'Private Sub ListView2_LostFocus()
'  Debug.Print "ListView2_LostFocus"
'  cmdOK.Default = True
'End Sub

Private Sub optCalc_Click()
  Call PopulateAvailable
  Call UpdateButtonStatus
End Sub

Private Sub optColumns_Click()
  Call PopulateAvailable
  Call UpdateButtonStatus
End Sub


Private Sub optAllRecords_Click()
  Call RecordSelectionClick(False, False)
End Sub

Private Sub optOutputFormat_Click(Index As Integer)
  
  chkDestination.Item(0).Value = vbChecked
  chkDestination.Item(1).Value = vbUnchecked
  chkDestination.Item(2).Value = vbUnchecked
  
  cboEMailField.ListIndex = IIf(cboEMailField.ListCount > 0, 0, -1)
  txtEmailSubject.Text = vbNullString
  chkEMailAttachment.Value = IIf(mbIsLabel, vbChecked, vbUnchecked)
  
  cboDocManEngine.ListIndex = IIf(cboDocManEngine.ListCount > 0, 0, -1)
  txtDocumentMap.Text = vbNullString
  txtDocumentMap.Tag = 0
  chkDocManManualHeader.Value = vbUnchecked
  chkDocManScreen.Value = vbUnchecked
  
  fraOutput(0).Visible = (Index = 0)
  fraOutput(1).Visible = (Index = 1)
  fraOutput(2).Visible = (Index = 2)
  
  Me.Changed = True

End Sub

Private Sub optPicklist_Click()
  Call RecordSelectionClick(True, False)
End Sub

Private Sub optFilter_Click()
  Call RecordSelectionClick(False, True)
End Sub


Private Sub RecordSelectionClick(blnPicklist As Boolean, blnFilter As Boolean)

  If mblnLoading Then
    Exit Sub
  End If

  On Error GoTo LocalErr
  
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
  
  ForceDefinitionToBeHiddenIfNeeded
  
  Me.Changed = True

Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error with record selection criteria"

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
    UpdateButtonStatus
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
    If mfColumnDrag Then
      ListView1.Drag vbCancel
      mfColumnDrag = False
    End If
  End If
  
End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)

  'If mblnReadOnly Then
  '  Exit Sub
  'End If

  UpdateButtonStatus

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
    If mfColumnDrag Then
      ListView2.Drag vbCancel
      mfColumnDrag = False
    End If
  End If
  
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Start the drag operation
  Dim objItem As ComctlLib.ListItem

  'Don't do drag drop if this is read only...
  If mblnReadOnly = True Then
    Exit Sub
  End If

  If Button = vbLeftButton Then
    If ListView1.ListItems.Count > 0 Then
      mfColumnDrag = True
      ListView1.Drag vbBeginDrag
    End If
  End If

End Sub

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  'Start the drag operation
  Dim objItem As ComctlLib.ListItem

  If mblnReadOnly Then
    Exit Sub
  End If

  If Button = vbLeftButton Then
    If ListView2.ListItems.Count > 0 Then
      mfColumnDrag = True
      ListView2.Drag vbBeginDrag
    End If
  End If

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

Private Sub ListView2_DragDrop(Source As Control, X As Single, Y As Single)
  
  ' Perform the drop operation - action depends on source and destination
  
  If Source Is ListView1 Then
    'If ListView2.HitTest(x, y) Is Nothing Then
      CopyToSelected False
    'Else
    '  CopyToSelected False, ListView2.HitTest(x, y).Index
    'End If
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


Private Sub fraColumns_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture
  
End Sub


Private Sub ListView2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    If UCase(Left(Source.SelectedItem.Key, 1)) = "E" Then
      Source.DragIcon = picDocument(1).Picture
    Else
      Source.DragIcon = picDocument(0).Picture
    End If
  End If

  ' Set DropHighlight to the mouse's coordinates.
  Set ListView2.DropHighlight = ListView2.HitTest(X, Y)

End Sub

Private Sub ListView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    If UCase(Left(Source.SelectedItem.Key, 1)) = "E" Then
      Source.DragIcon = picDocument(1).Picture
    Else
      Source.DragIcon = picDocument(0).Picture
    End If
  End If

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

Private Function CopyToSelected(bAll As Boolean)

  ' Copy items to the 'Selected' listview
  Dim objExpr As clsExprExpression
  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  
  Dim strText As String
  Dim strType As String
  Dim itmX As ListItem
  Dim bCheckIfHidden As Boolean
  
  bCheckIfHidden = False

  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass

  For Each itmX In ListView2.ListItems
    itmX.Selected = False
  Next
  ListView2.SelectedItem = Nothing
  ListView2.Refresh

  
  For iLoop = 1 To ListView1.ListItems.Count

    If bAll Or ListView1.ListItems(iLoop).Selected Then

      strType = Left(ListView1.ListItems(iLoop).Key, 1)
      strText = ListView1.ListItems(iLoop).Text
    
      If strType = "C" Then
        'Prefix column names with table name
        strText = GetTableNameFromColumn( _
                  Mid$(ListView1.ListItems(iLoop).Key, 2) _
                  ) & "." & strText
      End If
  
      Set itmX = ListView2.ListItems.Add(, ListView1.ListItems(iLoop).Key, strText, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon)
      itmX.Tag = ListView1.ListItems(iLoop).Tag
      itmX.SubItems(1) = strType & strText
      
      '01/08/2000 MH Fault 2010
      itmX.SubItems(2) = ListView1.ListItems(iLoop).SubItems(2)
      itmX.SubItems(3) = ListView1.ListItems(iLoop).SubItems(3)
      If strType = "C" Then
        itmX.SubItems(4) = ListView1.ListItems(iLoop).SubItems(4)
      Else
        bCheckIfHidden = True
        Set objExpr = New clsExprExpression
        objExpr.ExpressionID = Val(Mid(ListView1.ListItems(iLoop).Key, 2))
        objExpr.ConstructExpression
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        objExpr.ValidateExpression True, True
      
        itmX.SubItems(4) = (objExpr.ReturnType = giEXPRVALUE_NUMERIC Or _
                            objExpr.ReturnType = giEXPRVALUE_BYREF_NUMERIC)

        Set objExpr = Nothing
      End If
      
      ' Is this column to start on a new line
      itmX.SubItems(5) = True

      'MH20001102 Fault 1256
      'itmX.Selected = true
      itmX.Selected = Not bAll
      
      Changed = True
  
    End If
  
  Next iLoop
  
  
  For iLoop = ListView1.ListItems.Count To 1 Step -1
    If bAll Or ListView1.ListItems(iLoop).Selected Then
      
      If mblnDefinitionCreator Or ListView1.ListItems(iLoop).Tag <> "HD" Then
        iTempItemIndex = iLoop
        ListView1.ListItems.Remove ListView1.ListItems(iLoop).Key
      End If
    
    End If
  Next iLoop


  'Put this back in so that you can press enter to transfer column
  'and the next in the list is highlighted...
  If ListView1.ListItems.Count > 0 Then
    If iTempItemIndex > ListView1.ListItems.Count Then
      iTempItemIndex = ListView1.ListItems.Count
    End If
    If iTempItemIndex > 0 Then
      ListView1.ListItems(iTempItemIndex).Selected = True
    End If
  End If
  
  If bCheckIfHidden Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
  UpdateButtonStatus

  'TM20020104 Fault 3145 - Reset 'module' level hidden calc vaiables.
  miHiddenCalcs = 0
  msHiddenCalcsSelected = vbNullString
  
  Me.Changed = True
  Screen.MousePointer = vbDefault

Exit Function

LocalErr:
  ErrorCOAMsgBox "Error selecting columns"

End Function

Private Function CopyToAvailable(bAll As Boolean)

  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  Dim lngID As Long
  Dim strType As String
  
  ' Dont add the to the first listview...just remove em and
  ' repopulate the available listview...much quicker
  
  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass
  
  With ListView2.ListItems

    For iLoop = .Count To 1 Step -1
      If bAll Or .Item(iLoop).Selected Then
      
        strType = Left(.Item(iLoop).Key, 1)
        lngID = Val(Mid(.Item(iLoop).Key, 2))
      
        If strType = sTYPECODE_HEADING Or strType = sTYPECODE_SEPARATOR Then
          iTempItemIndex = iLoop
          .Remove .Item(iLoop).Key
        Else
          If Not IsInSortOrder(lngID) Then
            iTempItemIndex = iLoop
            .Remove .Item(iLoop).Key
          End If
        End If
      End If
    Next iLoop

    If .Count > 0 Then
      If iTempItemIndex > .Count Then
        iTempItemIndex = .Count
      End If
      If iTempItemIndex > 0 Then
        .Item(iTempItemIndex).Selected = True
        ListView2.Refresh
        DoEvents
      End If
    End If

  End With
  
  Call PopulateAvailable
  
  ForceDefinitionToBeHiddenIfNeeded
  Call UpdateButtonStatus

  Me.Changed = True
  Screen.MousePointer = vbDefault

Exit Function

LocalErr:
  ErrorCOAMsgBox "Error deselecting columns"

End Function

Private Function UpdateButtonStatus()

  Dim lst As ListItem
  Dim blnFoundNumeric As Boolean
  Dim intSelCount As Integer
  Dim iSelectedRow As Integer
  Dim iCount As Integer
  Dim bEnableSize As Boolean
  
  Call CheckListViewColWidth(ListView1)
  Call CheckListViewColWidth(ListView2)
  
  blnFoundNumeric = False
  intSelCount = 0
  iCount = 0
  bEnableSize = True
  
  For Each lst In ListView2.ListItems
    iCount = iCount + 1
    If lst.Selected Then
      iSelectedRow = iCount
      intSelCount = intSelCount + 1
      If ListView2.SelectedItem.SubItems(4) Then
        blnFoundNumeric = Not mblnReadOnly
        Exit For
      End If
    End If
  Next

  If intSelCount > 0 Then
    If intSelCount = 1 Then
      spnSize.Value = ListView2.SelectedItem.SubItems(2)
      spnDec.Value = ListView2.SelectedItem.SubItems(3)
      'MH20030417 Fault 5359
      If mbIsLabel Then
        chkStartColumnOnNewLine.Value = IIf(ListView2.SelectedItem.SubItems(5), 1, 0)
        txtProp_ColumnHeading.Enabled = IIf(Left(ListView2.SelectedItem.SubItems(1), 1) = "H", Not mblnReadOnly, False)
        txtProp_ColumnHeading.Text = ListView2.SelectedItem.SubItems(6)
        lblProp_ColumnHeading.Enabled = txtProp_ColumnHeading.Enabled
        bEnableSize = Not (Left(ListView2.SelectedItem.SubItems(1), 1) = sTYPECODE_HEADING _
            Or Left(ListView2.SelectedItem.SubItems(1), 1) = sTYPECODE_SEPARATOR)
        chkStartColumnOnNewLine.Enabled = IIf(Left(ListView2.SelectedItem.SubItems(1), 1) = "S", False, Not mblnReadOnly)
      
      End If
    Else
      'NHRD20092005 Fault 10281
      If mbIsLabel Then
        chkStartColumnOnNewLine.Value = IIf(ListView2.SelectedItem.SubItems(5), 1, 0)
      Else
        chkStartColumnOnNewLine.Value = vbUnchecked
      End If
      
      spnSize.Text = vbNullString
      spnDec.Text = vbNullString
      'chkStartColumnOnNewLine.Value = vbUnchecked
      txtProp_ColumnHeading.Enabled = False
      txtProp_ColumnHeading.Text = ""
    End If
    
    txtProp_ColumnHeading.BackColor = IIf(txtProp_ColumnHeading.Enabled, vbWhite, vbButtonFace)
    
    spnSize.Enabled = (bEnableSize And Not mblnReadOnly)
    lblProp_Size.Enabled = spnSize.Enabled
    spnSize.BackColor = IIf(spnSize.Enabled, vbWindowBackground, vbButtonFace)
    
    lblProp_Decimals.Enabled = blnFoundNumeric
    spnDec.Enabled = blnFoundNumeric
    spnDec.BackColor = IIf(blnFoundNumeric, vbWindowBackground, vbButtonFace)
    
    'chkStartColumnOnNewLine.Enabled = True
  
  Else
    spnSize.Value = 0
    lblProp_Size.Enabled = False
    spnSize.Enabled = False
    spnSize.BackColor = vbButtonFace
    
    spnDec.Value = 0
    lblProp_Decimals.Enabled = False
    spnDec.Enabled = False
    spnDec.BackColor = vbButtonFace
    
    chkStartColumnOnNewLine.Enabled = False
    
    'NHRD15072004 Fault 8675 Disabling Heding field and label.
    txtProp_ColumnHeading.Enabled = False
    txtProp_ColumnHeading.Text = ""
    lblProp_ColumnHeading.Enabled = txtProp_ColumnHeading.Enabled
    txtProp_ColumnHeading.BackColor = IIf(txtProp_ColumnHeading.Enabled, vbWhite, vbButtonFace)
  End If

  If mblnReadOnly Then
    Exit Function
  End If
  
  cmdAddAll.Enabled = (ListView1.ListItems.Count > 0)
  cmdAdd.Enabled = (ListView1.ListItems.Count > 0)  ' And Not (IsNull(ListView1.SelectedItem)))
  
  cmdRemoveAll.Enabled = (ListView2.ListItems.Count > 0)
  cmdRemove.Enabled = (ListView2.ListItems.Count > 0)   ' And Not (IsNull(ListView2.SelectedItem)))
    
  cmdMoveUp.Enabled = (ListView2.ListItems.Count > 0) And (intSelCount = 1) And iSelectedRow > 1
  cmdMoveDown.Enabled = (ListView2.ListItems.Count > 0) And (intSelCount = 1) And iSelectedRow < ListView2.ListItems.Count
  
  'If mbIsLabel Then
  '  lblTooManyColumnsForLabel.Visible = (NumberOfRowsSelected > miNumberOfRowsInLabel) And (mlngLabelTypeID > 0)
  'End If
  
End Function

Private Sub UpdateOrderButtonStatus()

  If mblnReadOnly Then
    Exit Sub
  End If
  
  With grdReportOrder
    
    If (.Rows = 1) Then
      .MoveFirst
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End If
    
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
      Me.cmdSortMoveUp.Enabled = False
      Me.cmdSortMoveDown.Enabled = .Rows > 1
    ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
      Me.cmdSortMoveUp.Enabled = .Rows > 1
      Me.cmdSortMoveDown.Enabled = False
    Else
      Me.cmdSortMoveUp.Enabled = .Rows > 1
      Me.cmdSortMoveDown.Enabled = .Rows > 1
    End If
  
  End With
  
End Sub

Private Sub RefreshReportOrderGrid()
  
  With grdReportOrder
    .Enabled = True
    .AllowUpdate = (False)
    
    If mblnReadOnly Then
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

  UpdateOrderButtonStatus
  
End Sub

'Private Function SelectLast(lvwCtl As ListView)
'
'  Dim objItem As ListItem
'
'  For Each objItem In lvwCtl.ListItems
'    objItem.Selected = IIf(objItem.Index = lvwCtl.ListItems.Count, True, False)
'  Next objItem
'
'End Function

Private Sub cmdAdd_Click()
  
  ' Add the selected items to the 'Selected' Listview
  CopyToSelected False

End Sub

Private Sub cmdRemove_Click()

  ' Remove the selected items from the 'Selected' Listview
  CopyToAvailable False
  'UpdateButtonStatus
  'If ListView2.ListItems.Count = 0 Then EnableColProperties False
  
End Sub

Private Sub cmdAddAll_Click()

  ' Add All items from to the 'Selected' Listview
  CopyToSelected True
  
End Sub

Private Sub cmdRemoveAll_Click()

  ' Remove All items from the 'Selected' Listview
  If grdReportOrder.Rows > 0 Then
    If COAMsgBox("Removing all selected columns will also clear the sort order." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
      grdReportOrder.RemoveAll
      CopyToAvailable True
    End If
  Else
    If COAMsgBox("Are you sure you wish to remove all columns / calculations from this definition ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
      CopyToAvailable True
    End If
  End If
  
End Sub


Private Sub spnDec_Change()

  Dim lst As ListItem
  
  If spnDec.Text <> vbNullString Then
    For Each lst In ListView2.ListItems
      If lst.Selected Then
        If Not (lst.SubItems(3) = spnDec.Value) Then
          lst.SubItems(3) = spnDec.Value
          Me.Changed = True
        End If
      End If
    Next
  End If

End Sub

Private Sub spnSize_Change()
  
  Dim lst As ListItem
  
  If spnSize.Text <> vbNullString Then
    For Each lst In ListView2.ListItems
      If lst.Selected Then
        If Not (lst.SubItems(2) = spnSize.Value) Then
          lst.SubItems(2) = spnSize.Value
          Me.Changed = True
        End If
      End If
    Next
  
  End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  EnableDisableTabControls

End Sub

Private Sub txtDesc_Change()
  Me.Changed = True
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

Private Sub txtEmailAttachmentName_Change()
  Me.Changed = True
End Sub

Private Sub txtEmailAttachmentName_GotFocus()
  With txtEmailAttachmentName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmailSubject_Change()
  Me.Changed = True
End Sub

Private Sub txtFilename_Change(Index As Integer)
  Me.Changed = True
End Sub

Private Sub txtName_Change()
  Me.Changed = True
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub



Private Function ValidateDefinition()

  Dim strRecSelStatus As String
  Dim blnContinueSave As Boolean
  Dim blnSaveAsNew As Boolean
  Dim strName As String
  Dim strMessage As String
  
  Dim iCount_Owner As Integer
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobScheduledUserGroups As String
  Dim sHiddenGroups As String
  
  fBatchJobsOK = True
  
  On Error GoTo LocalErr
  
  'Check that all required information has been completed before attempting to save
  ValidateDefinition = False
  strName = Trim(txtName.Text)
  
  If Len(strName) = 0 Then
    SSTab1.Tab = 0
    COAMsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If
  
  If optFilter Then
    If Val(txtFilter.Tag) = 0 Then
      SSTab1.Tab = 0
      WarningCOAMsgBox "No filter entered for the base table."
      cmdFilter.SetFocus
      Exit Function
    End If
  ElseIf optPicklist Then
    If Val(txtPicklist.Tag) = 0 Then
      SSTab1.Tab = 0
      WarningCOAMsgBox "No picklist entered for the base table."
      cmdPicklist.SetFocus
      Exit Function
    End If
  End If
  
  If ListView2.ListItems.Count = 0 Then
    SSTab1.Tab = 1
    'strMessage = "No " & IIf(mbIsLabel, "label", "merge") & " columns selected."
    strMessage = "No columns selected."
    WarningCOAMsgBox strMessage
    ListView1.SetFocus
    Exit Function
  End If

  If ListView2.ListItems.Count > 250 Then
    SSTab1.Tab = 1
    strMessage = "A maximum of 250 columns are allowed."
    WarningCOAMsgBox strMessage
    ListView1.SetFocus
    Exit Function
  End If

'  If mbIsLabel Then
'    If Not DoesJobFitOnLabel Then
'      SSTab1.Tab = 1
'      'WarningCOAMsgBox "There are too many columns for the format of the label type."
'      WarningCOAMsgBox "There are too many columns for this template."
'      ListView1.SetFocus
'      Exit Function
'    End If
'  End If

  If grdReportOrder.Rows = 0 Then
    SSTab1.Tab = 2
    'strMessage = "You must select at least 1 column to order the " & IIf(mbIsLabel, "label", "mail") & " merge by."
    strMessage = "You must select at least 1 column to order by."
    WarningCOAMsgBox strMessage
    cmdNewOrder.SetFocus
    Exit Function
  End If
  
  If txtFilename(1).Text = vbNullString Or txtFilename(1).Text = "<None>" Then
    SSTab1.Tab = 3
    If Not mbIsLabel Then
      WarningCOAMsgBox "No Template selected."
      cmdFilename(1).SetFocus
    Else
      WarningCOAMsgBox "No Template selected."
    End If
    Exit Function
  End If
  
  If mbIsLabel Then
    If Not DoesJobFitOnLabel Then
      SSTab1.Tab = 1
      'WarningCOAMsgBox "There are too many columns for the format of the label type."
      WarningCOAMsgBox "There are too many columns for this template."
      ListView1.SetFocus
      Exit Function
    End If
  End If
  
  On Error Resume Next
  
  If Not mbIsLabel Then
    If Dir(txtFilename(1).Text) = vbNullString Then
      SSTab1.Tab = 3
      WarningCOAMsgBox "Template file not found."
      cmdFilename(1).SetFocus
      Exit Function
    End If
  End If
  
  
  On Error GoTo LocalErr

  If optOutputFormat(0).Value = True Then
    ' Word Doc validation
    If txtFilename(1).Text = txtFilename(0).Text Then
      SSTab1.Tab = 3
      WarningCOAMsgBox "Word cannot give the save document the same name as the template document." + vbCrLf + "Enter a different name for the document you want to save."
      cmdFilename(0).SetFocus
      Exit Function
    End If

    If chkDestination(0).Value = vbUnchecked And _
       chkDestination(1).Value = vbUnchecked And _
       chkDestination(2).Value = vbUnchecked Then
          SSTab1.Tab = 3
          WarningCOAMsgBox "You must select a destination."
          Exit Function
    End If

    If chkDestination(2).Value = vbChecked And Len(txtFilename(0).Text) = 0 Then
      SSTab1.Tab = 3
      WarningCOAMsgBox "You must enter a file name."
      cmdFilename(0).SetFocus
      Exit Function
    End If

  ElseIf optOutputFormat(1).Value = True Then
    ' Email validation
    If cboEMailField.Enabled And cboEMailField.Text = "<None>" Then
      SSTab1.Tab = 3
      WarningCOAMsgBox "No email column selected."
      cboEMailField.SetFocus
      Exit Function
    End If

    If chkEMailAttachment.Value = vbChecked Then
      If Trim(txtEmailAttachmentName.Text) = vbNullString Then
        SSTab1.Tab = 3
        WarningCOAMsgBox "You must enter an attachment file name."
        txtEmailAttachmentName.SetFocus
        Exit Function
      End If
      
      If InStr(txtEmailAttachmentName, "/") > 0 Or _
         InStr(txtEmailAttachmentName, ":") > 0 Or _
         InStr(txtEmailAttachmentName, "?") > 0 Or _
         InStr(txtEmailAttachmentName, Chr(34)) > 0 Or _
         InStr(txtEmailAttachmentName, "<") > 0 Or _
         InStr(txtEmailAttachmentName, ">") > 0 Or _
         InStr(txtEmailAttachmentName, "|") > 0 Or _
         InStr(txtEmailAttachmentName, "\") > 0 Or _
         InStr(txtEmailAttachmentName, "*") Then
        SSTab1.Tab = 3
        WarningCOAMsgBox "The attachment file name cannot contain any of the following characters:" & vbCrLf & _
                      "/  :  ?  " & Chr(34) & "  <  >  |  \  *"
        txtEmailAttachmentName.SetFocus
        Exit Function
      End If
    End If
  
  ElseIf optOutputFormat(2).Value = True Then
    ' Document management validation
'    If Val(txtDocumentMap.Tag) = 0 Then
'      WarningCOAMsgBox "No document management type is defined."
'      SSTab1.Tab = 3
'      cmdDocumentMap.SetFocus
'      Exit Function
'    End If
  
  End If
  
  
  'MH20020801 Fault 4227
  'If InStr(txtEmailAttachmentName, "\") > 0 Or InStr(txtEmailAttachmentName, "*") Then
  '  SSTab1.Tab = 3
  '  WarningCOAMsgBox "Invalid 'Attach as' name." & vbCrLf & _
  '                "Please remove all '\' and '*' signs"
  '  txtEmailAttachmentName.SetFocus
  '  Exit Function
  'End If
  
  
  'Check if this definition has been changed by another user
  If mbIsLabel Then
    Call UtilityAmended(utlLabel, mlngMailMergeID, mlngTimeStamp, blnContinueSave, blnSaveAsNew)
  Else
    Call UtilityAmended(utlMailMerge, mlngMailMergeID, mlngTimeStamp, blnContinueSave, blnSaveAsNew)
  End If
  
  If blnContinueSave = False Then
    Exit Function
  ElseIf blnSaveAsNew Then
    txtUserName = gsUserName
    mblnDefinitionCreator = True
    mblnReadOnly = False
    ForceAccess
    mlngMailMergeID = 0
  End If

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
    
  If UniqueName(strName) = False Then
    SSTab1.Tab = 0
    strMessage = "A definition called '" & Trim(txtName.Text) & "' already exists."
    WarningCOAMsgBox strMessage
    txtName.SetFocus
    Exit Function
  End If
  
If mlngMailMergeID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    If mbIsLabel Then
      CheckCanMakeHiddenInBatchJobs utlLabel, _
        CStr(mlngMailMergeID), _
        txtUserName.Text, _
        iCount_Owner, _
        sBatchJobDetails_Owner, _
        sBatchJobIDs, _
        sBatchJobDetails_NotOwner, _
        fBatchJobsOK, _
        sBatchJobDetails_ScheduledForOtherUsers, _
        sBatchJobScheduledUserGroups, _
        sHiddenGroups
    Else
        CheckCanMakeHiddenInBatchJobs utlMailMerge, _
          CStr(mlngMailMergeID), _
          txtUserName.Text, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups, _
          sHiddenGroups
    End If

    If (Not fBatchJobsOK) Then
      If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
        COAMsgBox "This definition cannot be made hidden from the following user groups :" & vbCrLf & vbCrLf & sBatchJobScheduledUserGroups & vbCrLf & _
               "as it is used in the following batch jobs which are scheduled to be run by these user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, _
               vbExclamation + vbOKOnly, Me.Caption
      Else
        COAMsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , Me.Caption
      End If

      Screen.MousePointer = vbDefault
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      
      
      If COAMsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                
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

LocalErr:
  ErrorCOAMsgBox "Error validating " & Me.Caption & " definition"
  ValidateDefinition = False

End Function


'Private Function DefChangedSinceEdit() As Boolean
'
'  Dim strMBText As String
'  Dim intMBButtons As Integer
'  Dim strMBTitle As String
'  Dim intMBResponse As Integer
'
'  Dim blnAmended As Boolean
'  Dim blnDeleted As Boolean
'
'  DefChangedSinceEdit = False
'
'  'Check to see if another user has changed this definition
'  'whilst this user has been editting this definition
'  If mlngMailMergeID > 0 Then
'    blnAmended = datGeneral.RecordAmended(mlngMailMergeID, mlngTimeStamp, mstrSQLTableDef, blnDeleted, "MailMergeID")
'
'    If blnDeleted Then
'      strMBText = "The current definition has been deleted by another user. " & _
'                  "Would you still like to save this definition?"
'      intMBButtons = vbExclamation + vbYesNo
'      strMBTitle = "Mail Merge"
'
'      intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
'
'      Select Case intMBResponse
'      Case vbYes      'If deleted then save as new record
'        mlngMailMergeID = 0
'
'      Case vbNo       'Do not save
'        DefChangedSinceEdit = True
'
'      End Select
'
'    Else
'      strMBText = "The current definition has been amended by another user. " & _
'                  "click 'YES' to overwrite definition, 'NO' to save as a new definition or 'CANCEL' to continue to edit this definition"
'      intMBButtons = vbExclamation + vbYesNoCancel
'      strMBTitle = "Mail Merge"
'
'      intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
'
'      Select Case intMBResponse
'      Case vbYes      'overwrite existing definition and any changes
''Check if read only
'      Case vbNo       'save as new (but this may cause duplicate name message)
'        mlngMailMergeID = 0
'
'      Case vbCancel   'Do not save
'        DefChangedSinceEdit = True
'
'      End Select
'
'    End If
'  End If
'
'End Function


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






Private Function UniqueName(sName As String) As Boolean

  Dim rsName As Recordset
  Dim sSQL As String
    
  sSQL = "SELECT * FROM " & mstrSQLTableDef & _
         " WHERE Name = '" & Replace(sName, "'", "''") & "' AND MailMergeID <> " & mlngMailMergeID
    
  ' JDM - 30/06/03 - Fault 6147 - Separate labels and mail merges
  If mbIsLabel Then
    sSQL = sSQL & " AND IsLabel = 1"
  Else
    sSQL = sSQL & " AND IsLabel = 0"
  End If
    
  Set rsName = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  UniqueName = (rsName.BOF And rsName.EOF)
  rsName.Close
    
  Set rsName = Nothing

End Function


Private Function SaveDefinition() As Boolean

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strName As String
  Dim strDesc As String
  Dim strTableID As String
  'Dim strOrderID As String
  Dim strSelection As Integer
  Dim strPicklist As String
  Dim strFilter As String
  Dim strTemplate As String
  Dim strSuppressBlanks As String
  Dim strPauseBeforeMerge As String
  Dim strIsLabel As String
  Dim strUserName As String
  Dim blnAmended As Boolean
  Dim blnDeleted As Boolean

  Dim strOutput As String
  Dim strOutputScreen As String
  Dim strOutputPrinter As String
  Dim strOutputPrinterName As String
  Dim strOutputSave As String
  Dim strOutputFilename As String
  
  Dim strEmailAttachment As String
  Dim strEmailAttachmentName As String
  Dim strEmailAddrID As String
  Dim strEmailSubject As String
  
  Dim strDocManMapID As String
  Dim strDocManManualHeader As String
  Dim iUtilityType As utilityType
  
  On Error GoTo LocalErr

  SaveDefinition = True
  Screen.MousePointer = vbHourglass


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

  iUtilityType = IIf(mbIsLabel, utlLabel, utlMailMerge)

  If optOutputFormat(2).Value = True Then        'Document Management
    strOutput = "2"
    strOutputScreen = IIf(chkDocManScreen.Value = vbChecked, "1", "0")
    strOutputPrinter = "1"
    strOutputPrinterName = IIf(cboDocManEngine.Text <> "<Default Printer>", "'" & Replace(cboDocManEngine.Text, "'", "''") & "'", "''")
    strOutputSave = "0"
    strOutputFilename = "''"
    strEmailAddrID = "0"
    strEmailSubject = "''"
    strEmailAttachment = "0"
    strEmailAttachmentName = "''"
    strDocManMapID = txtDocumentMap.Tag
    strDocManManualHeader = IIf(chkDocManManualHeader.Value = vbChecked, "1", "0")
  
  ElseIf optOutputFormat(1).Value = True Then    'Email
    strOutput = "1"
    strOutputScreen = "0"
    strOutputPrinter = "0"
    strOutputPrinterName = "''"
    strOutputSave = "0"
    strOutputFilename = "''"
    strEmailAddrID = CStr(cboEMailField.ItemData(cboEMailField.ListIndex))
    strEmailSubject = "'" & Replace(txtEmailSubject, "'", "''") & "'"
    strEmailAttachment = IIf(chkEMailAttachment.Value = vbChecked, "1", "0")
    strEmailAttachmentName = "'" & Replace(txtEmailAttachmentName.Text, "'", "''") & "'"
    strDocManMapID = "0"
    strDocManManualHeader = "0"
  
  Else                                          'Word Document
    strOutput = "0"
    strOutputScreen = IIf(chkDestination(0).Value = vbChecked, "1", "0")
    strOutputPrinter = IIf(chkDestination(1).Value = vbChecked, "1", "0")
    strOutputPrinterName = IIf(cboPrinterName.Text <> "<Default Printer>", "'" & Replace(cboPrinterName.Text, "'", "''") & "'", "''")
    strOutputSave = IIf(chkDestination(2).Value = vbChecked, "1", "0")
    strOutputFilename = "'" & Replace(txtFilename(0).Text, "'", "''") & "'"
    strEmailAddrID = "0"
    strEmailSubject = "''"
    strEmailAttachment = "0"
    strEmailAttachmentName = "''"
    strDocManMapID = "0"
    strDocManManualHeader = "0"
  
  End If
  
  strName = "'" & Replace(Trim(txtName.Text), "'", "''") & "'"
  strDesc = "'" & Replace(txtDesc.Text, "'", "''") & "'"
  strTableID = CStr(cboBaseTable.ItemData(cboBaseTable.ListIndex))
  strTemplate = "'" & Replace(txtFilename(1).Text, "'", "''") & "'"
  strSuppressBlanks = CStr(Abs(chkSuppressBlank <> 0))
  strPauseBeforeMerge = CStr(Abs(chkPauseBeforeMerge <> 0))

  strUserName = "'" & datGeneral.UserNameForSQL & "'"
  strIsLabel = CStr(Abs(mbIsLabel <> 0))

  If mlngMailMergeID > 0 Then
    strSQL = "UPDATE " & mstrSQLTableDef & " SET " & _
             "Name = " & strName & ", " & _
             "Description = " & strDesc & ", " & _
             "TableID = " & strTableID & ", " & _
             "Selection = " & strSelection & ", " & _
             "PickListID = " & strPicklist & ", " & _
             "FilterID = " & strFilter & ", " & _
             "TemplateFileName = " & strTemplate & ", " & _
             "SuppressBlanks = " & strSuppressBlanks & ", " & _
             "PauseBeforeMerge = " & strPauseBeforeMerge & ", " & _
             "IsLabel = " & strIsLabel & ", " & _
             "LabelTypeID = " & Str(mlngLabelTypeID) & ", " & _
             "OutputFormat = " & strOutput & ", " & _
             "OutputScreen = " & strOutputScreen & ", " & _
             "OutputPrinter = " & strOutputPrinter & ", " & _
             "OutputPrinterName = " & strOutputPrinterName & ", " & _
             "OutputSave = " & strOutputSave & ", " & _
             "OutputFilename = " & strOutputFilename & ", " & _
             "EmailAddrID = " & strEmailAddrID & ", " & _
             "EMailSubject = " & strEmailSubject & ", " & _
             "EMailAsAttachment = " & strEmailAttachment & ", " & _
             "EmailAttachmentName = " & strEmailAttachmentName & ", " & _
             "DocumentMapID = " & strDocManMapID & ", " & _
             "ManualDocManHeader = " & strDocManManualHeader & " " & _
             "WHERE MailMergeID = " & CStr(mlngMailMergeID)
    gADOCon.Execute strSQL, , adCmdText


    Call UtilUpdateLastSaved(iUtilityType, mlngMailMergeID)

  Else
    strSQL = "INSERT " & mstrSQLTableDef & " (" & _
                "Name, Description, TableID, " & _
                "Selection, PicklistID, FilterID, " & _
                "UserName, IsLabel, LabelTypeID, " & _
                "TemplateFileName, SuppressBlanks, PauseBeforeMerge, " & _
                "OutputFormat, OutputScreen, " & _
                "OutputPrinter, OutputPrinterName, " & _
                "OutputSave, OutputFileName, " & _
                "EmailAddrID, EmailSubject, " & _
                "EMailAsAttachment, EmailAttachmentName, " & _
                "DocumentMapID, ManualDocManHeader) " & _
             "VALUES(" & _
                strName & ", " & strDesc & ", " & strTableID & ", " & _
                strSelection & ", " & strPicklist & ", " & strFilter & ", " & _
                strUserName & ", " & strIsLabel & ", " & CStr(mlngLabelTypeID) & ", " & _
                strTemplate & ", " & strSuppressBlanks & ", " & strPauseBeforeMerge & ", " & _
                strOutput & ", " & strOutputScreen & ", " & _
                strOutputPrinter & ", " & strOutputPrinterName & ", " & _
                strOutputSave & ", " & strOutputFilename & ", " & _
                strEmailAddrID & ", " & strEmailSubject & ", " & _
                strEmailAttachment & ", " & strEmailAttachmentName & ", " & _
                strDocManMapID & ", " & strDocManManualHeader & ")"

    ' RH 04/09/00 - Use the new stored procedure for inserting util defs
    mlngMailMergeID = InsertMailMerge(strSQL)

    Call UtilCreated(iUtilityType, mlngMailMergeID)

  End If

  SaveAccess
  
  SaveObjectCategories cboCategory, iUtilityType, mlngMailMergeID
  
  
Exit Function

LocalErr:
  ErrorCOAMsgBox "Error saving " & Me.Caption & " definition"
  SaveDefinition = False

End Function

Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysMailMergeAccess WHERE ID = " & mlngMailMergeID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysMailMergeAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngMailMergeID & ", sysusers.name," & _
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
      sSQL = "IF EXISTS (SELECT * FROM ASRSysMailMergeAccess" & _
        " WHERE ID = " & CStr(mlngMailMergeID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysMailMergeAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngMailMergeID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub






Private Function InsertMailMerge(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertMerge_ERROR

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
    pmADO.Value = mstrSQLTableDef
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "MailMergeID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, app.ProductName
        InsertMailMerge = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertMailMerge = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertMerge_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function


Private Function SaveColumns() As Boolean

  Dim strSQL As String
  Dim strColumnType As String
  Dim lngColumnID As Long
  Dim lngSequence As Long
  Dim strAscDesc As String
  Dim strHeading As String
  
  'Dim lngEmailAddrID As Long
  Dim intCount As Integer
  
  On Error GoTo LocalErr
  
  SaveColumns = True
  
  strSQL = "DELETE FROM " & mstrSQLTableCol & " WHERE MailMergeID = " & CStr(mlngMailMergeID)
  gADOCon.Execute strSQL, , adCmdText


  'lngEmailAddrID = 0
  'If cboEMailField.Enabled Then
  '  lngEmailAddrID = cboEMailField.ItemData(cboEMailField.ListIndex)
  'End If


  With ListView2

    For intCount = 1 To .ListItems.Count

      strColumnType = Left$(.ListItems(intCount).Key, 1)
      lngColumnID = Val(Mid$(.ListItems(intCount).Key, 2))

      'If lngEmailAddrID > 0 And lngEmailAddrID = lngColumnID Then
      '  'Don't need to add email column if it has already been selected
      '  lngEmailAddrID = 0
      'End If

      lngSequence = 0
      strAscDesc = vbNullString
      Call GetSortOrder(lngColumnID, lngSequence, strAscDesc)

      strHeading = Replace(.ListItems(intCount).SubItems(6), "'", "''")

      strSQL = "INSERT " & mstrSQLTableCol & _
               " (MailMergeID, Type, ColumnID, SortOrderSequence, SortOrder, Size, Decimals, StartOnNewLine, ColumnOrder, HeadingText) " & _
               "VALUES( " & _
               CStr(mlngMailMergeID) & ", " & _
               "'" & strColumnType & "', " & _
               CStr(lngColumnID) & ", " & _
               CStr(lngSequence) & ", " & _
               "'" & strAscDesc & "', " & _
               CStr(.ListItems(intCount).SubItems(2)) & ", " & _
               CStr(.ListItems(intCount).SubItems(3)) & ", " & _
               CStr(IIf(.ListItems(intCount).SubItems(5) = "True", "1", "0")) & ", " & _
               CStr(intCount) & ", '" & strHeading & "')"
      gADOCon.Execute strSQL, , adCmdText
    Next
  
  End With

  'If lngEMailColumnID > 0 Then
  '  'Insert additional record (Type 'X') for email field
  '  strSQL = "INSERT " & mstrSQLTableCol & _
  '           " (MailMergeID, Type, ColumnID, SortOrderSequence, SortOrder) " & _
  '           "VALUES( " & _
  '           CStr(mlngMailMergeID) & ", " & _
  '           "'X', " & _
  '           CStr(lngEMailColumnID) & ", 0, '')"
  '  gADOCon.Execute strSQL, , adCmdText
  'End If

Exit Function

LocalErr:
  ErrorCOAMsgBox "Error saving " & Me.Caption & " columns"
  SaveColumns = False

End Function


Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property

Private Sub RetreiveDefinition()

  Dim rsTemp As Recordset
  Dim strRecSelStatus As String
  Dim sMessage As String
  Dim fAlreadyNotified As Boolean
  Dim strPrinterName As String
  Dim iUtilityType As utilityType

  On Error GoTo LocalErr

  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, Me.Caption
    fOK = False
    Exit Sub
  End If

  SetComboText cboBaseTable, GetItemName(True, rsTemp!TableID)
  Let mstrPrimaryTable = cboBaseTable.Text

  Call LoadPrimaryDependantCombos

  ' Number of rows in a label type
  If mbIsLabel Then
    miNumberOfRowsInLabel = datGeneral.HowManyRowsInALabel(rsTemp!LabelTypeID)
    UpdateButtonStatus
    iUtilityType = utlLabel
  Else
    miNumberOfRowsInLabel = 0
    iUtilityType = utlMailMerge
  End If

  GetObjectCategories cboCategory, iUtilityType, mlngMailMergeID

  optOutputFormat(Val(rsTemp!OutputFormat)).Value = True
  
 
  
  Select Case Val(rsTemp!OutputFormat)
    Case 0  'Word Document

      chkDestination(0).Value = IIf(rsTemp!OutputScreen, vbChecked, vbUnchecked)
      chkDestination(1).Value = IIf(rsTemp!OutputPrinter, vbChecked, vbUnchecked)
      chkDestination(2).Value = IIf(rsTemp!OutputSave, vbChecked, vbUnchecked)
      
      SetPrinterCombo cboPrinterName, IIf(IsNull(rsTemp!OutputPrinterName), "", rsTemp!OutputPrinterName)
      
      txtFilename(0).Text = rsTemp!OutputFilename

    Case 1  'Individual Email
      'OutputClick Val(rsTemp!Output)
      optOutputFormat(1).Visible = True   'Make sure in case its a label

      chkEMailAttachment = Abs(rsTemp!EMailAsAttachment Or mbIsLabel)
      txtEmailAttachmentName = IIf(IsNull(rsTemp!EmailAttachmentName), "", rsTemp!EmailAttachmentName)

      txtEmailSubject.Text = rsTemp!EmailSubject

      If IIf(IsNull(rsTemp!EmailAddrID), 0, rsTemp!EmailAddrID) = 0 Then
        COAMsgBox "Please select a destination email address for this merge.", vbExclamation, Me.Caption
        mblnRecordSelectionInvalid = True
      Else
        SetComboItem cboEMailField, rsTemp!EmailAddrID
      End If

    Case 2  'Document Management
      'OutputClick Val(rsTemp!Output)

      optOutputFormat(2).Visible = True   'Make sure in case its not licenced
      SetPrinterCombo cboDocManEngine, IIf(IsNull(rsTemp!OutputPrinterName), "", rsTemp!OutputPrinterName)
      
'      If rsTemp!DocumentMapID > 0 Then
'        txtDocumentMap.Tag = rsTemp!DocumentMapID
'        txtDocumentMap.Text = rsTemp!DocumentMapName
'      End If
'
'      chkDocManManualHeader.Value = IIf(rsTemp!ManualDocManHeader, vbChecked, vbUnchecked)
      chkDocManScreen.Value = IIf(rsTemp!OutputScreen, vbChecked, vbUnchecked)

  End Select

  ' === Label specific stuff
  mlngLabelTypeID = IIf(IsNull(rsTemp!LabelTypeID), 0, rsTemp!LabelTypeID)

  txtFilename(1).Text = rsTemp!TemplateFileName

  chkSuppressBlank.Value = Abs(rsTemp!SuppressBlanks)
  chkPauseBeforeMerge.Value = Abs(rsTemp!PauseBeforeMerge)

  ' === Standard access stuff ===

  txtDesc.Text = IIf(rsTemp!Description <> vbNullString, rsTemp!Description, vbNullString)

  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!userName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!userName) = LCase$(gsUserName))
  End If

  'mblnHiddenPicklistOrFilter = False
  'mblnRecordSelectionInvalid = False

  If rsTemp!PicklistID > 0 Then
    optPicklist.Value = True
    txtPicklist.Tag = rsTemp!PicklistID
    txtPicklist.Text = rsTemp!PicklistName

    cmdPicklist.Enabled = (Not mblnReadOnly)
    cmdFilter.Enabled = False

  ElseIf rsTemp!FilterID > 0 Then
    optFilter.Value = True
    txtFilter.Tag = rsTemp!FilterID
    txtFilter.Text = rsTemp!FilterName

    cmdPicklist.Enabled = False
    cmdFilter.Enabled = (Not mblnReadOnly)

  Else
    optAllRecords.Value = True

    cmdPicklist.Enabled = False
    cmdFilter.Enabled = False

  End If

  If mbIsLabel Then
    mblnReadOnly = Not datGeneral.SystemPermission("LABELS", "EDIT")
  Else
    mblnReadOnly = Not datGeneral.SystemPermission("MAILMERGE", "EDIT")
  End If

  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    If mbIsLabel Then
      mblnReadOnly = (CurrentUserAccess(utlLabel, mlngMailMergeID) = ACCESS_READONLY)
    Else
      mblnReadOnly = (CurrentUserAccess(utlMailMerge, mlngMailMergeID) = ACCESS_READONLY)
    End If
  End If

  If mblnReadOnly Then
    ControlsDisableAll Me
    cmdCalculations.Enabled = True
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True

  mlngTimeStamp = rsTemp!intTimestamp

  ' =============================
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    fOK = False
    Exit Sub
  End If
  
Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error retrieving " & Me.Caption & " definition"

End Sub


Private Function SetPrinterCombo(cboTemp As ComboBox, strValue As String) As Boolean

  If strValue <> vbNullString Then
    SetComboText cboTemp, strValue
    If cboTemp.Text <> strValue Then
      cboTemp.AddItem strValue
      cboTemp.ListIndex = cboPrinterName.NewIndex
      If Not mblnFormPrint Then
        COAMsgBox "This definition is set to output to printer " & strValue & _
               " which is not set up on your PC.", vbInformation, Me.Caption
      End If
    End If
  End If
  
  If cboTemp.ListIndex < 0 And cboTemp.ListCount > 0 Then
    cboTemp.ListIndex = 0
  End If

End Function



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


Private Function GetOrderName(lItemID As Long) As String

  Dim sSQL As String
  Dim rsItem As Recordset
    
  sSQL = "Select Name From ASRSysOrders Where OrderID = " & lItemID
    
  Set rsItem = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  GetOrderName = rsItem(0)

  rsItem.Close
  Set rsItem = Nothing

End Function


Private Sub GetEmailFieldDefault()

  Dim rsTemp As Recordset
  Dim strSQL As String

  strSQL = "SELECT DefaultEmailID FROM ASRSysTables " & _
           "WHERE TableID = 0 OR TableID = " & CStr(cboBaseTable.ItemData(cboBaseTable.ListIndex))
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

'  With cboOutput
'    If .ListIndex <> -1 Then
'      If .ItemData(.ListIndex) = 2 Then
'        If Not rsTemp.BOF And Not rsTemp.EOF Then
'          SetComboItem cboEMailField, rsTemp!DefaultEmailID
'        End If
'      End If
'    End If
'  End With

  With cboEMailField
    If .ListIndex = -1 And .ListCount > 0 Then
      .ListIndex = 0
    End If
  End With

End Sub


Private Sub RetrieveColumns()

  Dim objExpr As clsExprExpression
  Dim rsColumns As Recordset
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strKey As String
  
  Dim strType As String
  Dim strText As String
  Dim itmX As ListItem
  Dim lngTableID As Long
  
  Dim sMessage As String
  Dim fAlreadyNotified As Boolean

  On Error GoTo LocalErr

  'Merge Column Type 'X' is a hidden column which is required by the mail merge
  '(Current only used for the email column)

  ' Do not sort the columns if we are a label
  If mbIsLabel Then
    strSQL = "SELECT * FROM " & mstrSQLTableCol & _
             " WHERE MailMergeID = " & CStr(mlngMailMergeID) & " AND Type <> 'X'" & _
             " ORDER BY ColumnOrder"
  Else
    strSQL = "SELECT * FROM " & mstrSQLTableCol & _
             " WHERE MailMergeID = " & CStr(mlngMailMergeID) & " AND Type <> 'X'" & _
             " ORDER BY SortOrderSequence"
  End If
  Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  On Error Resume Next
    
  mblnDefinitionCreator = IsDefinitionCreator(mlngMailMergeID)
  
  Do While Not rsColumns.EOF
    
    strType = rsColumns!Type
    
    Select Case strType
    
    ' Headings
    Case sTYPECODE_HEADING
      strKey = UniqueKey(strType)
      strText = sDFLTTEXT_HEADING
      Set itmX = ListView2.ListItems.Add(, strKey, strText)
      itmX.SubItems(1) = strKey
      itmX.SubItems(2) = 0
      itmX.SubItems(3) = 0
      itmX.SubItems(4) = False
      itmX.SubItems(5) = rsColumns!StartOnNewLine
      itmX.SubItems(6) = rsColumns!HeadingText
    
    ' Seperators
    Case sTYPECODE_SEPARATOR
      strKey = UniqueKey(strType)
      strText = sDFLTTEXT_SEPARATOR
      Set itmX = ListView2.ListItems.Add(, strKey, strText)
      itmX.SubItems(1) = strKey
      itmX.SubItems(2) = 0
      itmX.SubItems(3) = 0
      itmX.SubItems(4) = False
      itmX.SubItems(5) = rsColumns!StartOnNewLine
      itmX.SubItems(6) = " "
    
    ' Columns
    Case sTYPECODE_COLUMN
      strKey = strType & CStr(rsColumns!ColumnID)
      lngTableID = datGeneral.GetColumnTable(rsColumns!ColumnID)
      strText = datGeneral.GetTableName(lngTableID) & "." & datGeneral.GetColumnName(rsColumns!ColumnID)
      Set itmX = ListView2.ListItems.Add(, strKey, strText, ImageList1.ListImages("IMG_TABLE").Index, ImageList1.ListImages("IMG_TABLE").Index)
      itmX.SubItems(1) = strType & strText
      itmX.SubItems(2) = rsColumns!Size
      itmX.SubItems(3) = rsColumns!Decimals
      itmX.SubItems(4) = (datGeneral.GetDataType(lngTableID, rsColumns!ColumnID) = sqlNumeric)
      itmX.SubItems(5) = IIf(Not IsNull(rsColumns!StartOnNewLine), rsColumns!StartOnNewLine, False)
      itmX.SubItems(6) = ""

    ' Expressions
    Case sTYPECODE_EXPRESSION
      
'      If rsTemp.BOF And rsTemp.EOF Then
      
      'JPD 20031211 Fault 7679 - Already do these checks when ForceDefinitionToBeHiddenIfNeeded
      ' is called in the RetreiveDefinition method.
'      'TM20010807 Fault 2656
'      strKey = strType & CStr(rsColumns!ColumnID)
'      sMessage = IsCalcValid(rsColumns!ColumnID)
'      If sMessage <> vbNullString _
'        Or (GetExprField(rsColumns!ColumnID, "Access") = "HD" And Not mblnDefinitionCreator) Then
'        If Not fAlreadyNotified Then
'          If sMessage = vbNullString Then
'            sMessage = "The calculation used in this definition has been made hidden by another user."
'          End If
'
'          If FormPrint Then
'            sMessage = Me.Caption & " print failed : " & vbCrLf & vbCrLf & sMessage
'            COAMsgBox sMessage, vbExclamation + vbOKOnly, App.Title
'            Me.Cancelled = True
'            Exit Sub
'          End If
'
'          COAMsgBox sMessage & vbCrLf & _
'                 "It will be removed from the definition.", vbExclamation + vbOKOnly, App.Title
'
'          fAlreadyNotified = True
'        End If
'        mblnDeletedCalc = True
'        mbNeedsSave = True
'
''      ElseIf sMessage <> vbNullString And Not mblnDefinitionCreator Then
''          COAMsgBox "This definition contains hidden calculation(s) and is owned by another user." & vbCrLf & "This definition will now be made hidden.", vbExclamation + vbOKOnly, App.Title
''          SetUtilityAccess mlngMailMergeID, "MAILMERGE", "HD"
''          Me.Changed = False
''          Me.Cancelled = True
''          mblnDeletedCalc = False
''          Exit Sub
'
'      Else
        strKey = strType & CStr(rsColumns!ColumnID)
        strSQL = "Select * From ASRSysExpressions Where ExprID = " & rsColumns!ColumnID
        Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
        Set itmX = ListView2.ListItems.Add(, strKey, rsTemp!Name, ImageList1.ListImages("IMG_CALC").Index, ImageList1.ListImages("IMG_CALC").Index)

        itmX.SubItems(1) = strType & strText
        itmX.SubItems(2) = rsColumns!Size
        itmX.SubItems(3) = rsColumns!Decimals
        
        Set objExpr = New clsExprExpression
        objExpr.ExpressionID = rsColumns!ColumnID
        objExpr.ConstructExpression
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        objExpr.ValidateExpression True, True

        itmX.SubItems(4) = (objExpr.ReturnType = giEXPRVALUE_NUMERIC Or _
                            objExpr.ReturnType = giEXPRVALUE_BYREF_NUMERIC)
        Set objExpr = Nothing
        
        itmX.SubItems(5) = IIf(Not IsNull(rsColumns!StartOnNewLine), rsColumns!StartOnNewLine, False)
        itmX.SubItems(6) = ""
  
        
        rsTemp.Close
        Set rsTemp = Nothing

'      End If
      
    
    End Select
    
    rsColumns.MoveNext
  Loop

  ' JDM - 01/04/04 - Fault 8436 - Moved sort order stuff down here because of differences between mail merge and envelopes
  grdReportOrder.RemoveAll

  strSQL = "SELECT * FROM " & mstrSQLTableCol & _
           " WHERE MailMergeID = " & CStr(mlngMailMergeID) & " AND Type <> 'X'" & _
           " ORDER BY SortOrderSequence"

  Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  rsColumns.MoveFirst
  Do While Not rsColumns.EOF
    If rsColumns!SortOrderSequence > 0 Then
      strText = datGeneral.GetTableName(lngTableID) & "." & datGeneral.GetColumnName(rsColumns!ColumnID)
      grdReportOrder.AddItem CStr(rsColumns!ColumnID) & vbTab & _
                                strText & vbTab & _
                                IIf(Left(rsColumns!SortOrder, 1) = "A", "Ascending", "Descending")
    End If
    rsColumns.MoveNext
  Loop


  If grdReportOrder.Rows > 0 Then
    grdReportOrder.SelBookmarks.RemoveAll
    grdReportOrder.MoveFirst
    grdReportOrder.SelBookmarks.Add grdReportOrder.Bookmark
  End If
  
  On Error GoTo 0

Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error retrieving " & Me.Caption & " definition"

End Sub

Private Sub PrimaryTableClick()
  
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  
  'If mblnLoading Or (cboBaseTable.Text = mstrPrimaryTable) Then
  If cboBaseTable.Text = mstrPrimaryTable Then
    Exit Sub
  End If
  
  
  On Error GoTo LocalErr
  
  intMBResponse = vbYes

  If mstrPrimaryTable <> vbNullString And ListView2.ListItems.Count > 0 And Not mblnLoading Then

    SSTab1.Tab = 0
    strMBText = "Warning: Changing the base table will result in all table/column " & _
            "specific aspects of this " & LCase(Me.Caption) & " being cleared." & vbCrLf & _
            "Are you sure you wish to continue?"
    intMBButtons = vbQuestion + vbYesNo
    strMBTitle = Me.Caption
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)

  End If

  If intMBResponse = vbYes Then
    Me.Changed = True
    
    If Not mblnLoading And mstrPrimaryTable <> vbNullString Then
      ListView2.ListItems.Clear
      grdReportOrder.RemoveAll
      UpdateButtonStatus
      UpdateOrderButtonStatus
      
    End If

    mstrPrimaryTable = cboBaseTable.Text
    Call LoadPrimaryDependantCombos
    optAllRecords.Value = True
    
'    With cboOutput
'      If .ListIndex <> -1 Then
'        If .ItemData(.ListIndex) = 2 Then
'          Call GetEmailFieldDefault
'        End If
'      End If
'    End With
  Else
    SetComboText cboBaseTable, mstrPrimaryTable
  
  End If

  If Not mblnLoading Then
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error changing base table"

End Sub


Private Sub LoadPrimaryDependantCombos()

  Dim sSQL As String
  Dim rsTables As New Recordset
  Dim lngDefaultOrderID As Long
  Dim fOriginalLoading As Boolean
  
  On Error GoTo LocalErr
  
  fOriginalLoading = mblnLoading
  mblnLoading = True

  Call PopulateTableCombo

  'sSQL = "SELECT DefaultOrderID FROM ASRSysTables WHERE TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
  'Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  'lngDefaultOrderID = rsTables!DefaultOrderID

  'sSQL = "SELECT Name,OrderID FROM ASRSysOrders WHERE TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
  'Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)


  'With cboOrder
  '  .Clear
  '  '.AddItem "<None>"
  '  '.ItemData(.NewIndex) = 0
  '  Do While Not rsTables.EOF
  '    .AddItem rsTables!Name
  '    .ItemData(.NewIndex) = rsTables!OrderID
  '    rsTables.MoveNext
  '  Loop
  '  SetComboItem cboOrder, lngDefaultOrderID
  'End With


  'sSQL = "SELECT ColumnName,ColumnID FROM ASRSysColumns WHERE TableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
'  If mblnEmailPermission Then
    sSQL = "SELECT Name, EmailID FROM ASRSysEmailAddress " & _
           "WHERE tableID = 0 OR tableID = " & cboBaseTable.ItemData(cboBaseTable.ListIndex)
    Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
    With cboEMailField
      .Clear
      Do While Not rsTables.EOF
        '.AddItem rsTables!ColumnName
        '.ItemData(.NewIndex) = rsTables!ColumnID
        .AddItem rsTables!Name
        .ItemData(.NewIndex) = rsTables!EmailID
        rsTables.MoveNext
      Loop
      If .ListCount > 0 Then
        '.ListIndex = 0
        Call GetEmailFieldDefault
      Else
        .AddItem "<None>"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
      End If
    End With
  
    rsTables.Close
    Set rsTables = Nothing
'  End If

  mblnLoading = fOriginalLoading

Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error populating email column combo box"

End Sub


Private Sub PopulateAvailable()

  Dim rsColumns As New Recordset
  Dim rsCalculations As New Recordset
  Dim sSQL As String
  Dim strKey As String
  Dim intCount As Integer
  Dim objListItem As ListItem
  
  On Error GoTo LocalErr

  Screen.MousePointer = vbHourglass

  ' Clear the contents of the Available Listview
  ListView1.ListItems.Clear
  
  If optColumns.Value Then
    ' Add the Columns of the selected table to the listview
    sSQL = "SELECT columnID, tableID, columnName, size, decimals, DefaultDisplayWidth, dataType " & _
      " FROM ASRSysColumns" & _
      " WHERE tableID = " & cboTblAvailable.ItemData(cboTblAvailable.ListIndex) & " " & _
      " AND columnType <> " & Trim(Str(colSystem)) & _
      " AND columnType <> " & Trim(Str(colLink)) & _
      " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
      " AND dataType <> " & Trim(Str(sqlOle)) & _
      " ORDER BY ColumnName"
    Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    ' Check if the column has already been selected. If so, dont add it
    ' to the available listview
    Do While Not rsColumns.EOF
      strKey = "C" & CStr(rsColumns!ColumnID)
      If Not AlreadyUsed(strKey) Then
  
        '01/08/2000 MH Fault 2010
        'ListView1.ListItems.Add , strKey, rsColumns!ColumnName, , ImageList1.ListImages("IMG_TABLE").Index
        Set objListItem = ListView1.ListItems.Add(, strKey, rsColumns!ColumnName, , ImageList1.ListImages("IMG_TABLE").Index)
        objListItem.SubItems(2) = rsColumns!DefaultDisplayWidth
        objListItem.SubItems(3) = rsColumns!Decimals
        objListItem.SubItems(4) = (rsColumns!DataType = sqlNumeric Or rsColumns!DataType = sqlInteger)
  
      End If
      rsColumns.MoveNext
    Loop
    ' Clear recordset reference
    Set rsColumns = Nothing
    
  Else
    ' Add the Expressions of the selected table to the listview
    ' Only add Type 1 expressions (CALCS)
    sSQL = "SELECT ExprID, Name, Access FROM ASRSysExpressions " & _
           "WHERE TableID = " & cboTblAvailable.ItemData(cboTblAvailable.ListIndex) & " " & _
           "AND Type = " & Trim(Str(giEXPR_RUNTIMECALCULATION)) & " " & _
           "AND ParentComponentID = 0 " & _
           "AND (Username = '" & datGeneral.UserNameForSQL & "' OR Access <> 'HD') " & _
           "ORDER BY Name"
    Set rsCalculations = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    ' Check if the column has already been selected. If so, dont add it
    Do While Not rsCalculations.EOF
      strKey = "E" & CStr(rsCalculations!ExprID)
      If Not AlreadyUsed(strKey) Then
        If IsCalcValid(rsCalculations!ExprID) = vbNullString Then
        ListView1.ListItems.Add , strKey, rsCalculations!Name, , ImageList1.ListImages("IMG_CALC").Index
        With ListView1.ListItems(strKey)
          .Tag = rsCalculations!Access
          .SubItems(2) = 0
          .SubItems(3) = 0
          .SubItems(4) = 0
      
        End With
        End If
      Else
        ListView2.ListItems(strKey).Text = rsCalculations!Name
        ListView2.ListItems(strKey).Tag = "*" & rsCalculations!Access
      End If
      rsCalculations.MoveNext
    Loop
    ' Clear recordset reference
    Set rsCalculations = Nothing
  End If

  ' Skip adding calcs if the table selected is not the base table
  If cboTblAvailable.ItemData(cboTblAvailable.ListIndex) <> cboBaseTable.ItemData(cboBaseTable.ListIndex) Then
    'not base table selected
    optCalc.Value = False
    optColumns.Value = True
    optCalc.Enabled = False
    optColumns.Enabled = False
    
    cmdCalculations.Visible = False
    ListView1.Height = 3100 '2565
'    Exit Sub
  Else
    optCalc.Enabled = Not mblnReadOnly
    optColumns.Enabled = Not mblnReadOnly

    'base table selected
    cmdCalculations.Visible = True
    ListView1.Height = 2600 '2685 '2100
    
  End If

'  If any calculations tags in listview2 are not "OK" then
'  they must have been deleted so remove them !
'  intCount = 1
'  Do While intCount <= ListView2.ListItems.Count
'    With ListView2.ListItems(intCount)
'      If Left$(.Key, 1) = "E" Then
'        If Left(.Tag, 1) = "*" Then
'          .Tag = Mid$(.Tag, 2)
'          .Selected = False
'        Else
'          ListView2.ListItems.Remove intCount
'          intCount = intCount - 1
'        End If
'      End If
'    End With
'    intCount = intCount + 1
'  Loop

  Screen.MousePointer = vbDefault

Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error populating columns available for selection"

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


Private Sub PopulateTableCombo()

  ' If something has been selected as a base table, this function populates
  ' the Table combo on the Columns Tab with the base table, its parents
  ' and its children.
  
  Dim rsParents As New Recordset
  Dim rsTables As New Recordset
  Dim rsChildren As New Recordset
  Dim sSQL As String
  
  On Error GoTo LocalErr
  
  ' Clear the contents of the tables combo
  cboTblAvailable.Clear
    
  ' Clear the listview
  ListView1.ListItems.Clear
    
  ' If no base table is selected, dont populate with anything
  'If cboBaseTable.ListIndex = 0 Then Exit Sub
    
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
    
  ' Select the base table in the combo by default
  With cboTblAvailable
    .ListIndex = 0
    .Enabled = (.ListCount > 1)
    .BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
  End With
  
Exit Sub

LocalErr:
  ErrorCOAMsgBox "Error populating tables available for selection"
  
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


Private Function GetSortOrder(lngID As Long, lngSeq As Long, strAscDesc As String) As Boolean

  Dim blnFound As Boolean
  Dim intLoop As Integer
  Dim pvarbookmark As Variant
  
  On Error GoTo LocalErr
  
  With grdReportOrder
  
    .MoveFirst
    blnFound = False
    intLoop = 0
    Do Until intLoop = .Rows
      
      pvarbookmark = .GetBookmark(intLoop)
      If Val(.Columns("ColExprID").CellText(pvarbookmark)) = lngID Then
        lngSeq = intLoop + 1
        strAscDesc = IIf(Left(.Columns("Sort Order").CellText(pvarbookmark), 1) = "A", "Asc", "Desc")
        Exit Function
      End If
      
      intLoop = intLoop + 1
    Loop

  End With

  GetSortOrder = blnFound

Exit Function

LocalErr:
  ErrorCOAMsgBox "Error checking sort order"
  
End Function


Private Function IsInSortOrder(plngColExprID As Long) As Boolean

  'Purpose : Removes the specified column from the sort order.
  'Input   : ColExprID
  'Output  : None
  
  Dim varBookmark As Variant
  Dim intLoop As Integer
  Dim lngColumnID As Long
  Dim lRow As Long
  
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  IsInSortOrder = False
  
  With grdReportOrder
    intLoop = 0
    Do While intLoop < .Rows
      'lngColumnID = Val(.Columns("ColExprID").CellText(.GetBookmark(intLoop)))
      lngColumnID = Val(.Columns("ColExprID").CellText(.AddItemBookmark(intLoop)))
      
      If lngColumnID = plngColExprID Or lngColumnID = 0 Then
        
        strMBText = "Removing the following column will also remove it from the " & Me.Caption & " sort order." & vbCrLf & vbCrLf & _
                    .Columns("Column").CellText(.AddItemBookmark(intLoop)) & vbCrLf & vbCrLf & _
                    "Do you wish to continue ?"
        intMBButtons = vbExclamation + vbYesNo
        strMBTitle = Me.Caption
        intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)

        If intMBResponse = vbYes Then
          If .Rows = 1 Then
            .RemoveAll
          Else
            lRow = intLoop    '.AddItemRowIndex(.Bookmark)
            .RemoveItem lRow
            If lRow < .Rows Then
              .Bookmark = lRow
            Else
              .Bookmark = (.Rows - 1)
            End If
            .SelBookmarks.Add .Bookmark
          End If
        Else
          IsInSortOrder = True  'Its still in sort order !
        End If

        UpdateOrderButtonStatus   'MH20021120 Fault 4706
        Exit Function
      End If
      intLoop = intLoop + 1
    Loop
  End With

End Function


Private Function GetDefinition() As Recordset

  Dim strSQL As String

  strSQL = "SELECT " & _
           mstrSQLTableDef & ".*, " & _
           "CONVERT(integer," & mstrSQLTableDef & ".TimeStamp) AS intTimeStamp, " & _
           "ASRSysPickListName.Name AS PickListName, " & _
           "ASRSysPickListName.Access AS PickListAccess, " & _
           "ASRSysExpressions.Name AS FilterName, " & _
           "ASRSysExpressions.Access AS FilterAccess, " & _
           "ASRSysDocumentManagementTypes.Name AS DocumentMapName " & _
           "FROM " & mstrSQLTableDef & " " & _
           "LEFT OUTER JOIN ASRSysExpressions ON " & mstrSQLTableDef & ".FilterID = ASRSysExpressions.ExprID " & _
           "LEFT OUTER JOIN ASRSysPickListName ON " & mstrSQLTableDef & ".PickListID = ASRSysPickListName.PickListID " & _
           "LEFT OUTER JOIN ASRSysDocumentManagementTypes ON " & mstrSQLTableDef & ".DocumentMapID = ASRSysDocumentManagementTypes.DocumentMapID  " & _
           "WHERE " & mstrSQLTableDef & ".MailMergeID = " & CStr(mlngMailMergeID)
  Set GetDefinition = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

End Function


Public Sub PrintDef(lMailMergeID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsColumns As Recordset
  Dim strSQL As String
  Dim iLoop As Integer
  Dim fFirstLoop As Boolean
  Dim varBookmark As Variant
  Dim strType As String
  
  Set datData = New DataMgr.clsDataAccess
  
  mlngMailMergeID = lMailMergeID
  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation, Me.Caption
    Exit Sub
  End If
  
  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      .TabsOnPage = 8
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        If mbIsLabel Then
          .PrintHeader "Envelopes & Labels : " & rsTemp!Name
          .PrintNormal "Category : " & GetObjectCategory(utlLabel, mlngMailMergeID)
        Else
          .PrintHeader "Mail Merge : " & rsTemp!Name
          .PrintNormal "Category : " & GetObjectCategory(utlMailMerge, mlngMailMergeID)
        End If
    
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal "Owner : " & rsTemp!userName
        
        ' Access section --------------------------------------------------------
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop

        .PrintNormal
        
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
        
        '--------
        
        .PrintTitle "Columns"
        .PrintBold "Data" & vbTab & vbTab & vbTab & "Type" & vbTab & "Size" & vbTab & "Decimals" & IIf(mbIsLabel, vbTab & "New Line", "")
      
        strSQL = "SELECT " & mstrSQLTableCol & ".*, " & _
                 "       ASRSysExpressions.Name as ExprName, " & _
                 "       UPPER(ASRSysExpressions.Name) as UCASEExprName, " & _
                 "       ASRSysColumns.ColumnName, " & _
                 "       ASRSysTables.TableName " & _
                 "FROM " & mstrSQLTableCol & _
                 " LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.ColumnID = " & mstrSQLTableCol & ".ColumnID " & _
                 " LEFT OUTER JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID " & _
                 " LEFT OUTER JOIN ASRSysExpressions ON ASRSysExpressions.ExprID = " & mstrSQLTableCol & ".ColumnID " & _
                 " WHERE MailMergeID = " & CStr(mlngMailMergeID) & " AND " & mstrSQLTableCol & ".Type <> 'X'"
        If mbIsLabel Then
          strSQL = strSQL & _
              " ORDER BY " & mstrSQLTableCol & ".ColumnOrder"
        Else
          strSQL = strSQL & _
              " ORDER BY " & mstrSQLTableCol & ".Type, TableName, ColumnName, UCASEExprName"
        End If
        Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
        
        Do While Not rsColumns.EOF
          
          Select Case rsColumns!Type
          Case sTYPECODE_COLUMN
            strType = IIf(rsColumns!Type = "C", "Column", "Calculation")
            .PrintNonBold rsColumns!TableName & "." & rsColumns!ColumnName & _
                vbTab & vbTab & vbTab & strType & vbTab & rsColumns!Size & vbTab & rsColumns!Decimals & _
                vbTab & IIf(mbIsLabel, IIf(rsColumns!StartOnNewLine = True, "Yes", "No"), "")
          
          Case sTYPECODE_EXPRESSION
            strType = IIf(rsColumns!Type = "C", "Column", "Calculation")
            .PrintNonBold IIf(IsNull(rsColumns!exprName), "<Deleted Calculation>", rsColumns!exprName) & _
                vbTab & vbTab & vbTab & strType & vbTab & rsColumns!Size & vbTab & rsColumns!Decimals & _
                vbTab & IIf(mbIsLabel, IIf(rsColumns!StartOnNewLine = True, "Yes", "No"), "")
          
          Case sTYPECODE_HEADING
            .PrintNonBold "<Header Text> : " & rsColumns!HeadingText & _
                vbTab & vbTab & vbTab & "" & vbTab & rsColumns!Size & vbTab & rsColumns!Decimals & _
                vbTab & IIf(mbIsLabel, IIf(rsColumns!StartOnNewLine = True, "Yes", "No"), "")
          
          Case sTYPECODE_SEPARATOR
            .PrintNonBold "<Separator>" & vbTab & vbTab & vbTab & "" & vbTab & rsColumns!Size & vbTab & rsColumns!Decimals & _
                vbTab & IIf(mbIsLabel, IIf(rsColumns!StartOnNewLine = True, "Yes", "No"), "")
          
          End Select
          
          rsColumns.MoveNext
        Loop
        .PrintNormal
    
        '--------
        
        .PrintTitle "Sort Order"
        .PrintBold "Column" & vbTab & vbTab & vbTab & vbTab & "Sort Order"
        
        strSQL = "SELECT " & mstrSQLTableCol & ".SortOrder, " & _
                 "       ASRSysColumns.ColumnName, " & _
                 "       ASRSysTables.TableName " & _
                 "FROM " & mstrSQLTableCol & _
                 " JOIN ASRSysColumns ON ASRSysColumns.ColumnID = " & mstrSQLTableCol & ".ColumnID " & _
                 " JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID " & _
                 " WHERE MailMergeID = " & CStr(mlngMailMergeID) & " AND Type <> 'X'" & _
                 " AND SortOrderSequence > 0 " & _
                 " ORDER BY SortOrderSequence"
        Set rsColumns = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

        Do While Not rsColumns.EOF
          .PrintNonBold rsColumns!TableName & "." & rsColumns!ColumnName & vbTab & vbTab & vbTab & vbTab & _
              IIf(Left(rsColumns!SortOrder, 1) = "A", "Ascending", "Descending")
          rsColumns.MoveNext
        Loop
        .PrintNormal
    
        '--------
        
        .PrintTitle "Output Options"
    
        If mbIsLabel Then
          .PrintNormal "Label Type : " & rsTemp!TemplateFileName
        Else
          .PrintNormal "Template File Name : " & rsTemp!TemplateFileName
        End If
           
        'MH20050105 Fault 9663
        'If mbIsLabel Then
        '  .PrintNormal "Pause Before Label Merge : " & IIf(rsTemp!PauseBeforeMerge, "True", "False")
        'Else
        '  .PrintNormal "Pause Before Merge : " & IIf(rsTemp!PauseBeforeMerge, "True", "False")
        'End If
        .PrintNormal "Pause Before Merge : " & IIf(rsTemp!PauseBeforeMerge, "True", "False")
        
        .PrintNormal "Suppress Blank Lines : " & IIf(rsTemp!SuppressBlanks, "True", "False")
        .PrintNormal
    
'        Select Case Val(rsTemp!Output)
'        Case OutputType.Document
'          .PrintNormal "Destination : New Document"
'          .PrintNormal "Save Output : " & IIf(rsTemp!DocSave, "True", "False")
'          If Abs(rsTemp!DocSave) <> 0 Then
'            .PrintNormal "Save File Name : " & rsTemp!DocFileName
'            .PrintNormal "Close Word After Save : " & IIf(rsTemp!CloseDoc, "True", "False")
'          End If
'        Case OutputType.Printer
'          .PrintNormal "Destination : Printer (" & cboPrinterName.Text & ")"
'        Case OutputType.Email
'          .PrintNormal "Destination : Email"
'          '.PrintNormal "Email Column : " & GetItemName(False, rsTemp!EmailColumnID)
'          .PrintNormal "Email Address : " & GetEmailName(rsTemp!EmailAddrID)
'          .PrintNormal "Email Subject : " & rsTemp!EmailSubject
'          .PrintNormal "Email as Attachment : " & IIf(rsTemp!EMailAsAttachment, "True", "False")
'          If rsTemp!EMailAsAttachment Then
'            .PrintNormal "Attach As : " & IIf(IsNull(rsTemp!EmailAttachmentName), "", rsTemp!EmailAttachmentName)
'          End If
'        Case OutputType.Version1
'          .PrintNormal "Destination : Version 1 (" & cboPrinterName.Text & ")"""
'          If Abs(rsTemp!AddV1Header) Then
'            .PrintNormal "Add Version 1 Header Information : True"
'          End If
'
'        End Select
        .PrintTitle "Output Options"
  
        Select Case Val(rsTemp!OutputFormat)
        Case 0: .PrintNormal "Output Format : Word Document"
        Case 1: .PrintNormal "Output Format : Individual Emails"
        Case 2: .PrintNormal "Output Format : Document Management"
        End Select
  
        Select Case Val(rsTemp!OutputFormat)
        Case 0  'Word Document

            'If chkPreview.Value = vbChecked Then
            '  .PrintNormal "Output Destination : Preview on screen prior to output"
            'End If

            If chkDestination(0).Value = vbChecked Then
              .PrintNormal "Output Destination : Display on screen"
            End If

            If chkDestination(1).Value = vbChecked Then
              .PrintNormal "Output Destination : Send to printer"
              .PrintNormal "Printer Location : " & cboPrinterName.List(cboPrinterName.ListIndex)
            End If

            If chkDestination(2).Value = vbChecked Then
              .PrintNormal "Output Destination : Save to file"
              .PrintNormal "File Name : " & txtFilename(0).Text
              '.PrintNormal "File Options : " & cboSaveExisting.List(cboSaveExisting.ListIndex)
            End If

        Case 1  'Individual Email

            .PrintNormal "Output Destination : Send to email"
            If rsTemp!EmailAddrID > 0 Then
              .PrintNormal "Email Address : " & GetEmailName(rsTemp!EmailAddrID)
            End If
            .PrintNormal "Email Subject : " & rsTemp!EmailSubject
            .PrintNormal "Email Attach As : " & IIf(IsNull(rsTemp!EmailAttachmentName), "", rsTemp!EmailAttachmentName)

        Case 2  'Document Management

          .PrintNormal "Engine : " & IIf(IsNull(rsTemp!OutputPrinterName), "", rsTemp!OutputPrinterName)
          '.PrintNormal "Document Type : " & IIf(IsNull(rsTemp!DocumentMapName), "", rsTemp!DocumentMapName)
          '.PrintNormal "Manual document header : " & IIf(IsNull(rsTemp!ManualDocManHeader), "", rsTemp!ManualDocManHeader)
          .PrintNormal "Display output on screen : " & IIf(IsNull(rsTemp!OutputScreen), "", rsTemp!OutputScreen)

        End Select


        '--------
    
        .PrintEnd
        If mbIsLabel Then
          .PrintConfirm "Labels & Envelopes : " & rsTemp!Name, "Labels & Envelopes Definition"
        Else
          .PrintConfirm "Mail Merge : " & rsTemp!Name, "Mail Merge Definition"
        End If
      End If
  
    End With
    
  End If
  
  If Not rsColumns Is Nothing Then rsColumns.Close
  If Not rsTemp Is Nothing Then rsTemp.Close
  
  Set rsColumns = Nothing
  Set rsTemp = Nothing
  Set datData = Nothing

Exit Sub

LocalErr:
  ErrorCOAMsgBox "Printing " & Me.Caption & " Definition Failed"

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
  
  ' Return false if some of the filters/picklists need to be removed from the definition,
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
        ' Picklist deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & cboBaseTable.List(cboBaseTable.ListIndex) & "' table filter"

        fRemove = (Not mblnReadOnly) And _
          (Not FormPrint)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Picklist hidden by another user.
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
      txtFilter.Tag = 0
      txtFilter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If
  
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
          .ListItems.Remove iLoop
  
          If Not FormPrint Then
            SSTab1.Tab = 1
            
            ' JDM - Fault 8973 - 27/07/04 - Commented out line below - was problematic when messagebox is displayed.
            '.SetFocus
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


Private Function HiddenCalcSelected() As Boolean

  Dim objItem As ListItem
  Dim rsTemp As Recordset
  Dim strSQL As String
  
  Set rsTemp = New Recordset
  
  HiddenCalcSelected = False

  For Each objItem In ListView2.ListItems
    If Left(objItem.Key, 1) = "E" Then
      
      strSQL = "SELECT * FROM AsrSysExpressions " & _
               "WHERE ExprID = " & Mid(objItem.Key, 2)
      Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
      
      HiddenCalcSelected = (rsTemp.Fields("Access") = "HD")
      If HiddenCalcSelected = True Then
        Exit For
      End If
    
    End If

  Next objItem
  
  If rsTemp.State = adStateOpen Then
    rsTemp.Close
  End If
  
  Set rsTemp = Nothing
  Set objItem = Nothing
  
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


Private Function GetEmailName(lngEmailID As Long) As String

  Dim rsTemp As Recordset
  Dim strSQL As String

  strSQL = "SELECT Name FROM ASRSysEmailAddress WHERE EmailID = " & CStr(lngEmailID)
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  GetEmailName = rsTemp(0)
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function

Private Sub EnableDisableTabControls()

  ' Definition tab page controls
  fraDefinition(0).Enabled = (SSTab1.Tab = 0)
  fraDefinition(1).Enabled = (SSTab1.Tab = 0)
  
  'Columns tab page controls
  fraColumns(0).Enabled = (SSTab1.Tab = 1)
  fraColumns(1).Enabled = (SSTab1.Tab = 1)
  
  'NHRD09112004 Fault 5830
  optCalc.TabStop = (Not SSTab1.Tab = 1)
  
  If (SSTab1.Tab = 1) Then
    UpdateButtonStatus
  End If
  
  'Sort order tab page controls
  cmdAdd.Enabled = (SSTab1.Tab = 1 And ListView1.ListItems.Count > 0 And Not mblnReadOnly)
  cmdAddAll.Enabled = (SSTab1.Tab = 1 And ListView1.ListItems.Count > 0 And Not mblnReadOnly)
  cmdRemove.Enabled = (SSTab1.Tab = 1 And ListView2.ListItems.Count > 0 And Not mblnReadOnly)
  cmdRemoveAll.Enabled = (SSTab1.Tab = 1 And ListView2.ListItems.Count > 0 And Not mblnReadOnly)
  
  If (SSTab1.Tab = 2) Then
    RefreshReportOrderGrid
    'UpdateOrderButtonStatus
  End If
  
  ' Output tab page controls
  fraSort(0).Enabled = (SSTab1.Tab = 2)
  'fraOutput(0).Enabled = (SSTab1.Tab = 3)
  'fraOutput(1).Enabled = (SSTab1.Tab = 3)
  'fraOutput(2).Enabled = (SSTab1.Tab = 3)
  
End Sub

Public Property Let IsLabel(ByVal bIsLabel As Boolean)
  mbIsLabel = bIsLabel
End Property

Private Sub DisplayLabelSpecifics()
  
  chkEMailAttachment.Enabled = Not mbIsLabel
  
  ' Display relevent tabs
  If mbIsLabel Then
  
    lblPrimary.Caption = "Template :"
    cmdLabelType.Visible = True
    cmdFilename(1).Visible = False
    chkStartColumnOnNewLine.Visible = True
    'chkPromptForPrintStart.Visible = True

    cmdMoveUp.Visible = True
    cmdMoveDown.Visible = True

    ' Don't automatically sort the selected columns
    ListView2.Sorted = False

    'MH20050105 Fault 9663
    'chkPauseBeforeMerge.Caption = "Pause &before label merge"
    Me.Caption = "Envelope & Label Definition"

    lblProp_ColumnHeading.Visible = True
    txtProp_ColumnHeading.Visible = True

    chkEMailAttachment.Value = vbChecked

  Else

    lblPrimary.Caption = "Template :"
    cmdLabelType.Visible = False
    cmdFilename(1).Visible = True
    chkStartColumnOnNewLine.Visible = False
    'chkPromptForPrintStart.Visible = False

    fraSizeDecimals.Top = 3480
    ListView2.Height = 3050

    ' Hide the move up down buttons
    cmdMoveUp.Visible = False
    cmdMoveDown.Visible = False

    ' Hide the add heading/separator buttons
    cmdAddHeading.Visible = False
    cmdAddSeparator.Visible = False

    cmdAdd.Top = cmdAddHeading.Top
    cmdAddAll.Top = cmdAddSeparator.Top

    ' Automatically sort the selected columns
    ListView2.Sorted = True

    'MH20050105 Fault 9663
    'chkPauseBeforeMerge.Caption = "Pause &before mail merge"
    Me.Caption = "Mail Merge Definition"

  End If

End Sub

Private Function ChangeSelectedOrder(Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)

  ' SUB COMPLETED 28/01/00
  ' This function changes the order of listitems in the selected listview.
  ' At the moment, different arrays are used depending on what information you
  ' need to store...change the array to a type if it would suit the purpose
  ' better
  
  ' Dimension arrays
  Dim iLoop As Integer, Key() As String, Text() As String, Icon() As Variant, SmallIcon() As Variant
  
  Dim SubItem1() As Variant, SubItem2() As Variant, SubItem3() As Variant, SubItem4() As Variant, SubItem5() As Variant, SubItem6() As Variant
  
  ReDim Key(0), Text(0), Icon(0), SmallIcon(0)
  ReDim SubItem1(0), SubItem2(0), SubItem3(0), SubItem4(0), SubItem5(0), SubItem6(0)
  
  Dim itmX As ListItem
  
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
    
      ReDim Preserve SubItem1(UBound(SubItem1) + 1)
      ReDim Preserve SubItem2(UBound(SubItem2) + 1)
      ReDim Preserve SubItem3(UBound(SubItem3) + 1)
      ReDim Preserve SubItem4(UBound(SubItem4) + 1)
      ReDim Preserve SubItem5(UBound(SubItem5) + 1)
      ReDim Preserve SubItem6(UBound(SubItem6) + 1)
            
      SubItem1(UBound(SubItem1) - 1) = ListView2.ListItems(iLoop).SubItems(1)
      SubItem2(UBound(SubItem2) - 1) = ListView2.ListItems(iLoop).SubItems(2)
      SubItem3(UBound(SubItem3) - 1) = ListView2.ListItems(iLoop).SubItems(3)
      SubItem4(UBound(SubItem4) - 1) = ListView2.ListItems(iLoop).SubItems(4)
      SubItem5(UBound(SubItem5) - 1) = ListView2.ListItems(iLoop).SubItems(5)
      SubItem6(UBound(SubItem6) - 1) = ListView2.ListItems(iLoop).SubItems(6)
        
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
    
      ReDim Preserve SubItem1(UBound(SubItem1) + 1)
      ReDim Preserve SubItem2(UBound(SubItem2) + 1)
      ReDim Preserve SubItem3(UBound(SubItem3) + 1)
      ReDim Preserve SubItem4(UBound(SubItem4) + 1)
      ReDim Preserve SubItem5(UBound(SubItem5) + 1)
      ReDim Preserve SubItem6(UBound(SubItem6) + 1)
            
      SubItem1(UBound(SubItem1) - 1) = ListView2.ListItems(iLoop).SubItems(1)
      SubItem2(UBound(SubItem2) - 1) = ListView2.ListItems(iLoop).SubItems(2)
      SubItem3(UBound(SubItem3) - 1) = ListView2.ListItems(iLoop).SubItems(3)
      SubItem4(UBound(SubItem4) - 1) = ListView2.ListItems(iLoop).SubItems(4)
      SubItem5(UBound(SubItem5) - 1) = ListView2.ListItems(iLoop).SubItems(5)
      SubItem6(UBound(SubItem6) - 1) = ListView2.ListItems(iLoop).SubItems(6)
    
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
      
        ReDim Preserve SubItem1(UBound(SubItem1) + 1)
        ReDim Preserve SubItem2(UBound(SubItem2) + 1)
        ReDim Preserve SubItem3(UBound(SubItem3) + 1)
        ReDim Preserve SubItem4(UBound(SubItem4) + 1)
        ReDim Preserve SubItem5(UBound(SubItem5) + 1)
        ReDim Preserve SubItem6(UBound(SubItem6) + 1)

        SubItem1(UBound(SubItem1) - 1) = ListView2.ListItems(iLoop).SubItems(1)
        SubItem2(UBound(SubItem2) - 1) = ListView2.ListItems(iLoop).SubItems(2)
        SubItem3(UBound(SubItem3) - 1) = ListView2.ListItems(iLoop).SubItems(3)
        SubItem4(UBound(SubItem4) - 1) = ListView2.ListItems(iLoop).SubItems(4)
        SubItem5(UBound(SubItem5) - 1) = ListView2.ListItems(iLoop).SubItems(5)
        SubItem6(UBound(SubItem6) - 1) = ListView2.ListItems(iLoop).SubItems(6)
      
      End If
    Next iLoop
  End If
  
  ' Clear all items from the listview
  ListView2.ListItems.Clear
  
  ' Add items in the right order from the array
  For iLoop = LBound(Key) To (UBound(Key) - 1)
    
    Set itmX = ListView2.ListItems.Add(, Key(iLoop), Text(iLoop), Icon(iLoop), SmallIcon(iLoop))
  
    itmX.SubItems(1) = SubItem1(iLoop)
    itmX.SubItems(2) = SubItem2(iLoop)
    itmX.SubItems(3) = SubItem3(iLoop)
    itmX.SubItems(4) = SubItem4(iLoop)
    itmX.SubItems(5) = SubItem5(iLoop)
    itmX.SubItems(6) = SubItem6(iLoop)
  
  Next iLoop
  
  If mfFromButtons = True Then
    ListView2.ListItems(intBeforeIndex - 1).Selected = True
  Else
    If intBeforeIndex < ListView2.ListItems.Count Then ListView2.ListItems(intBeforeIndex).Selected = True Else ListView2.ListItems(ListView2.ListItems.Count).Selected = True
  End If
  
  mfFromButtons = False
  
  Changed = True
  
  UpdateButtonStatus
  
End Function

Private Function UniqueKey(psType As String) As String
  Dim objItem As ListItem
  Dim iKey As Integer
  Dim iNewKey As Integer
  
  iNewKey = 1
  
  For Each objItem In ListView2.ListItems
    If Left(objItem.Key, 1) = psType Then
      iKey = Val(Right(objItem.Key, Len(objItem.Key) - 1))
    
      If iKey >= iNewKey Then
        iNewKey = iKey + 1
      End If
    End If
  Next objItem
  Set objItem = Nothing
  
  UniqueKey = psType & Trim(Str(iNewKey))
  
End Function

Private Sub txtProp_ColumnHeading_Change()

  Dim lst As ListItem

  'If txtProp_ColumnHeading.Text <> vbNullString Then
    For Each lst In ListView2.ListItems
      If lst.Selected Then
        If Not (lst.SubItems(6) = txtProp_ColumnHeading.Text) Then
          lst.SubItems(6) = txtProp_ColumnHeading.Text
          
          If spnSize.Enabled Then
            spnSize.Value = Len(txtProp_ColumnHeading.Text)
          End If
          Me.Changed = True
        End If
      End If
    Next
  'End If



End Sub

' The amount of rows that this label definition uses.
Private Function NumberOfRowsSelected() As Integer

  Dim iCount As Integer
  Dim iNewLines As Integer

  iNewLines = 0
  With ListView2.ListItems
    For iCount = 1 To .Count
      iNewLines = iNewLines + IIf(.Item(iCount).SubItems(5) = True Or iCount = 1, 1, 0)
      iNewLines = iNewLines + IIf(Left(.Item(iCount).SubItems(1), 1) = sTYPECODE_HEADING And iCount > 1, 1, 0)
    Next iCount
  End With
  NumberOfRowsSelected = iNewLines

End Function

Private Sub txtProp_ColumnHeading_KeyPress(KeyAscii As Integer)

  If KeyAscii = 34 Then   ' Double Quote
    KeyAscii = 0
  End If

End Sub


Public Function DoesJobFitOnLabel() As Boolean

  ' Returns how many rows a label type can show
  Dim sSQL As String
  Dim rsData As ADODB.Recordset
  Dim sngRequiredLabelSize As Single
  Dim iHeadingRows As Integer
  Dim iStandardRows As Integer
  Dim iSpaceBetweenRows As Integer
  Dim lst As ListItem

  iHeadingRows = 0
  iStandardRows = 0
  iSpaceBetweenRows = 3
  
  ' Count the types of row output we have
  For Each lst In ListView2.ListItems
    
    'MH20040128 Fault 7552 - Check if its "New Line"...
    If lst.SubItems(5) = True Then
      
      If Left(lst.SubItems(1), 1) = sTYPECODE_HEADING Then
        iHeadingRows = iHeadingRows + 1
      End If
      If Left(lst.SubItems(1), 1) <> sTYPECODE_HEADING Then
        iStandardRows = iStandardRows + 1
      End If
    
    End If
  
  Next lst

  ' Calculate required label size
  sSQL = "SELECT LabelHeight, PageHeight, IsEnvelope, StandardFontSize, HeadingFontSize" _
    & " FROM ASRSysLabelTypes" _
    & " WHERE LabelTypeID = " & Trim(Str(mlngLabelTypeID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
      
    sngRequiredLabelSize = MyPointsToCentimeters((rsData!StandardFontSize + iSpaceBetweenRows) * iStandardRows) _
        + MyPointsToCentimeters((rsData!HeadingFontSize + iSpaceBetweenRows) * iHeadingRows)
        
    If rsData!IsEnvelope Then
      DoesJobFitOnLabel = (sngRequiredLabelSize <= rsData!PageHeight)
    Else
      DoesJobFitOnLabel = (sngRequiredLabelSize <= rsData!LabelHeight)
    End If
  Else
    DoesJobFitOnLabel = False
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Private Sub PopulatePrintCombo(cboTemp As ComboBox)

  Dim objPrinter As Printer

  With cboTemp
    .Clear
    .AddItem "<Default Printer>"
    For Each objPrinter In Printers
      .AddItem objPrinter.DeviceName
    Next
  End With

End Sub

Private Sub cboCategory_Click()
  Changed = True
End Sub

