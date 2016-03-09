VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Definition"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1044
   Icon            =   "frmImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabImport 
      Height          =   5625
      Left            =   105
      TabIndex        =   46
      Top             =   75
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   9922
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "frmImport.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraInformation"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Colu&mns"
      TabPicture(1)   =   "frmImport.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraColumns"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "O&ptions"
      TabPicture(2)   =   "frmImport.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraFileDetails"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraTableDetails"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraOptions"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraSource"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame fraSource 
         Caption         =   "Data Source :"
         Height          =   1485
         Left            =   -74850
         TabIndex        =   53
         Top             =   400
         Width           =   9400
         Begin VB.CommandButton cmdFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4650
            TabIndex        =   58
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.OptionButton optAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   1575
            TabIndex        =   55
            Top             =   705
            Value           =   -1  'True
            Width           =   800
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "&Filter"
            Height          =   315
            Left            =   1575
            TabIndex        =   56
            Top             =   960
            Width           =   800
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   57
            Tag             =   "0"
            Top             =   960
            Width           =   2250
         End
         Begin VB.ComboBox cboFileFormat 
            Height          =   315
            ItemData        =   "frmImport.frx":0060
            Left            =   2415
            List            =   "frmImport.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   285
            Width           =   2550
         End
         Begin VB.Label lblFileRecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter :"
            Height          =   195
            Left            =   195
            TabIndex        =   64
            Top             =   705
            Width           =   555
         End
         Begin VB.Label lblFileType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Format :"
            Height          =   195
            Left            =   195
            TabIndex        =   59
            Top             =   345
            Width           =   900
         End
      End
      Begin VB.Frame fraInformation 
         Height          =   2355
         Left            =   150
         TabIndex        =   47
         Top             =   400
         Width           =   9405
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5805
            MaxLength       =   30
            TabIndex        =   4
            Top             =   300
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   1
            Top             =   300
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
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   3090
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1485
            Left            =   5805
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
            stylesets(0).Picture=   "frmImport.frx":0064
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
            stylesets(1).Picture=   "frmImport.frx":0080
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
            Left            =   4950
            TabIndex        =   52
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   51
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   195
            TabIndex        =   50
            Top             =   1155
            Width           =   1080
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   4950
            TabIndex        =   49
            Top             =   765
            Width           =   825
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category :"
            Height          =   240
            Left            =   195
            TabIndex        =   48
            Top             =   765
            Width           =   1005
         End
      End
      Begin VB.Frame fraColumns 
         Height          =   5055
         Left            =   -74850
         TabIndex        =   8
         Top             =   400
         Width           =   9400
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Move Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   15
            Top             =   3315
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "Move &Up"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   14
            Top             =   2775
            Width           =   1200
         End
         Begin VB.CommandButton cmdClearColumn 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   13
            Top             =   1935
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteColumn 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   12
            Top             =   1395
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditColumn 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7950
            TabIndex        =   11
            Top             =   855
            Width           =   1200
         End
         Begin VB.CommandButton cmdNewColumn 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7950
            TabIndex        =   10
            Top             =   315
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdColumns 
            Height          =   4590
            Left            =   210
            TabIndex        =   9
            Top             =   315
            Width           =   7575
            ScrollBars      =   0
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   8
            stylesets.count =   2
            stylesets(0).Name=   "ssetDormant"
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
            stylesets(0).Picture=   "frmImport.frx":009C
            stylesets(1).Name=   "ssetActive"
            stylesets(1).ForeColor=   16777215
            stylesets(1).BackColor=   8388608
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
            stylesets(1).Picture=   "frmImport.frx":00B8
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
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   8
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
            Columns(3).Width=   4551
            Columns(3).Caption=   "Table Name"
            Columns(3).Name =   "Table Name"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(3).Locked=   -1  'True
            Columns(4).Width=   4763
            Columns(4).Caption=   "Column Name"
            Columns(4).Name =   "Column Name"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(4).Locked=   -1  'True
            Columns(5).Width=   820
            Columns(5).Caption=   "Size"
            Columns(5).Name =   "Size"
            Columns(5).Alignment=   2
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(5).Locked=   -1  'True
            Columns(6).Width=   794
            Columns(6).Caption=   "Key"
            Columns(6).Name =   "Key"
            Columns(6).Alignment=   2
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   11
            Columns(6).FieldLen=   1
            Columns(6).Style=   2
            Columns(7).Width=   2090
            Columns(7).Caption=   "Create Lookup"
            Columns(7).Name =   "LookupEntries"
            Columns(7).Alignment=   2
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   11
            Columns(7).FieldLen=   1
            Columns(7).Style=   2
            TabNavigation   =   1
            _ExtentX        =   13361
            _ExtentY        =   8096
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
      Begin VB.Frame fraOptions 
         Caption         =   "Records :"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   35
         Top             =   4125
         Width           =   9400
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1110
            Left            =   5550
            TabIndex        =   43
            Top             =   120
            Width           =   3600
            Begin VB.OptionButton optDontUpdateAny 
               Caption         =   "Update &None"
               Height          =   195
               Left            =   1320
               TabIndex        =   41
               Top             =   825
               Width           =   2430
            End
            Begin VB.OptionButton optUpdateAll 
               Caption         =   "&Update All"
               Height          =   195
               Left            =   1320
               TabIndex        =   40
               Top             =   510
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.Label lblDupRecordsReturned 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "If key field(s) return multiple records :"
               Height          =   195
               Left            =   -15
               TabIndex        =   39
               Top             =   195
               Width           =   3330
            End
         End
         Begin VB.OptionButton optImportType 
            Caption         =   "Create new records onl&y (ignore key fields on base table)"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   945
            Width           =   5760
         End
         Begin VB.OptionButton optImportType 
            Caption         =   "Update &records and create new where not matched"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   315
            Value           =   -1  'True
            Width           =   4815
         End
         Begin VB.OptionButton optImportType 
            Caption         =   "Update &existing records only"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   630
            Width           =   3375
         End
      End
      Begin VB.Frame fraData 
         Caption         =   "Data :"
         Height          =   2565
         Left            =   150
         TabIndex        =   0
         Top             =   2895
         Width           =   9405
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   3000
         End
         Begin VB.Label lblBaseTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Left            =   225
            TabIndex        =   6
            Top             =   420
            Width           =   885
         End
      End
      Begin VB.Frame fraTableDetails 
         Caption         =   "Source Table :"
         Height          =   2070
         Left            =   -74850
         TabIndex        =   65
         Top             =   1965
         Width           =   9400
         Begin VB.CheckBox chkUseUpdateBlob 
            Caption         =   "Mar&k import records as processed"
            Height          =   285
            Left            =   240
            TabIndex        =   63
            Top             =   1635
            Width           =   3495
         End
         Begin VB.TextBox txtLinkedServer 
            Height          =   315
            Left            =   1620
            TabIndex        =   60
            Top             =   315
            Width           =   3350
         End
         Begin VB.TextBox txtLinkedCatalog 
            Height          =   315
            Left            =   1620
            TabIndex        =   61
            Top             =   750
            Width           =   3350
         End
         Begin VB.TextBox txtLinkedTable 
            Height          =   315
            Left            =   1620
            TabIndex        =   62
            Top             =   1200
            Width           =   3350
         End
         Begin VB.Label lblLinkedServer 
            Caption         =   "Server :"
            Height          =   300
            Left            =   255
            TabIndex        =   68
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label lblLinkedCatalog 
            Caption         =   "Catalog :"
            Height          =   255
            Left            =   255
            TabIndex        =   67
            Top             =   825
            Width           =   1170
         End
         Begin VB.Label lblLinkedTable 
            Caption         =   "Table :"
            Height          =   165
            Left            =   255
            TabIndex        =   66
            Top             =   1275
            Width           =   960
         End
      End
      Begin VB.Frame fraFileDetails 
         Caption         =   "File :"
         Height          =   2070
         Left            =   -74850
         TabIndex        =   16
         Top             =   1965
         Width           =   9400
         Begin VB.ComboBox cboDateSeparator 
            Height          =   315
            ItemData        =   "frmImport.frx":00D4
            Left            =   6795
            List            =   "frmImport.frx":00E4
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1500
            Width           =   1275
         End
         Begin VB.ComboBox cboDelimiter 
            Height          =   315
            ItemData        =   "frmImport.frx":00F9
            Left            =   6795
            List            =   "frmImport.frx":0106
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   300
            Width           =   2250
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Height          =   315
            Left            =   4650
            TabIndex        =   19
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtDelimiter 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   8895
            MaxLength       =   1
            TabIndex        =   27
            Top             =   300
            Width           =   300
         End
         Begin VB.TextBox txtEncapsulator 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   6795
            MaxLength       =   1
            TabIndex        =   29
            Text            =   """"
            Top             =   700
            Width           =   300
         End
         Begin VB.ComboBox cboDateFormat 
            Height          =   315
            ItemData        =   "frmImport.frx":011D
            Left            =   6795
            List            =   "frmImport.frx":012D
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1100
            Width           =   1275
         End
         Begin COASpinner.COA_Spinner spnHeaderLines 
            Height          =   315
            Left            =   2400
            TabIndex        =   21
            Top             =   1080
            Width           =   810
            _ExtentX        =   1429
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
         Begin COASpinner.COA_Spinner spnFooterLines 
            Height          =   315
            Left            =   2400
            TabIndex        =   23
            Top             =   1500
            Width           =   810
            _ExtentX        =   1429
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
         Begin VB.Label lblFooterLines 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Footer Lines to Ignore :"
            Height          =   195
            Left            =   225
            TabIndex        =   22
            Top             =   1560
            Width           =   1710
         End
         Begin VB.Label lblHeaderLines 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Header Lines to Ignore :"
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   1155
            Width           =   2070
         End
         Begin VB.Label lblDateSeparator 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separator :"
            Height          =   195
            Left            =   5175
            TabIndex        =   33
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label lblOtherDelimiter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other :"
            Height          =   195
            Left            =   8205
            TabIndex        =   26
            Top             =   360
            Width           =   525
         End
         Begin VB.Label lblFilename 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File name :"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblDelimiter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delimiter :"
            Height          =   195
            Left            =   5175
            TabIndex        =   24
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblEncapsulator 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data enclosed in :"
            Height          =   195
            Left            =   5175
            TabIndex        =   28
            Top             =   765
            Width           =   1605
         End
         Begin VB.Label lblEncapsulatorHint 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   7400
            TabIndex        =   42
            Top             =   750
            Width           =   45
         End
         Begin VB.Label lblDateFormat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date format :"
            Height          =   195
            Left            =   5175
            TabIndex        =   30
            Top             =   1155
            Width           =   1110
         End
         Begin VB.Label lblNoSep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(No Separator)"
            Height          =   195
            Left            =   8050
            TabIndex        =   32
            Top             =   1155
            Visible         =   0   'False
            Width           =   1350
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7335
      TabIndex        =   44
      Top             =   5835
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8590
      TabIndex        =   45
      Top             =   5835
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   75
      Top             =   5730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdatData As DataMgr.clsDataAccess
Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mlngImportID As Long
Private mblnReadOnly As Boolean
Private mlngTimeStamp As Long
Private mblnDefinitionCreator As Boolean
Private mstrBaseTable As String
Private mintCurrentFileFormat As ImportType
Private mlngOriginalFilterID As Long

Private Function ColumnInFilter(plngColumnID, plngFilterID As Long) As Boolean
  ' Check that the column is not used as a field component in an expression.
  Dim fUsed As Boolean
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  
  fUsed = False
  
  sSQL = "SELECT ASRSysExprComponents.componentID" & _
    " FROM ASRSysExprComponents" & _
    " WHERE ASRSysExprComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
    " AND ASRSysExprComponents.fieldColumnID = " & Trim(Str(plngColumnID)) & _
    " AND ASRSysExprComponents.exprID = " & Trim(Str(plngFilterID))

  Set rsTemp = datGeneral.GetRecords(sSQL)

  fUsed = Not (rsTemp.BOF And rsTemp.EOF)

  rsTemp.Close

  If Not fUsed Then
    ' Column not in this level of the filter expression. Try the next level down.
    sSQL = "SELECT ASRSysExpressions.exprID" & _
      " FROM ASRSysExpressions" & _
      " WHERE ASRSysExpressions.parentComponentID IN (" & _
      "   SELECT ASRSysExprComponents.componentID" & _
      "   FROM ASRSysExprComponents" & _
      "   WHERE ASRSysExprComponents.exprID = " & Trim(Str(plngFilterID)) & ")"
  
    Set rsTemp = datGeneral.GetRecords(sSQL)
    Do Until rsTemp.EOF
      fUsed = ColumnInFilter(plngColumnID, rsTemp!ExprID)
      
      If fUsed Then
        Exit Do
      End If
      
      rsTemp.MoveNext
    Loop
  
    rsTemp.Close
  End If
  
  ColumnInFilter = fUsed
  
End Function

Public Property Get SelectedID() As Long
  SelectedID = mlngImportID
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property
Private Sub ForceAccess(Optional pvAccess As Variant)
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    For iLoop = 0 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      .Bookmark = varBookmark
      
      If iLoop = 0 Then
        .Columns("Access").Text = ""
      Else
        If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
          If Not IsMissing(pvAccess) Then
            .Columns("Access").Text = AccessDescription(CStr(pvAccess))
          End If
        End If
      End If
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow

End Sub



Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOk.Enabled = pblnChanged
End Property

Private Sub cboBaseTable_Click()

  If (grdColumns.Rows = 0) Or (mstrBaseTable = cboBaseTable.Text) Then
    mstrBaseTable = cboBaseTable.Text
    Exit Sub
  End If

  If COAMsgBox("Warning: Changing the base table will result in all table/column " & _
            "specific aspects of this import definition being cleared." & vbCrLf & _
            "Are you sure you wish to continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Import") = vbYes Then
         
    Me.grdColumns.RemoveAll
    UpdateButtonStatus
    mstrBaseTable = cboBaseTable.Text
  
    ' Clear the file filter as it is tied to the selected columns.
    optAllRecords.Value = True
  Else
    SetComboText cboBaseTable, mstrBaseTable
  End If
  
End Sub


Private Sub cboDateFormat_Click()
  Changed = True
  'TM20020726 Fault 2123 - no need to show the "No Separator" label.
'  If Left(cboDateFormat.Text, 1) = "y" Then lblNoSep.Visible = True Else lblNoSep.Visible = False
  
End Sub

Private Sub cboDateSeparator_Click()

  Changed = True
  
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

Private Sub cboFileFormat_Click()

  ' Exit if the user has selected the file format that was already selected in the combo
  If mintCurrentFileFormat = cboFileFormat.ItemData(cboFileFormat.ListIndex) Then Exit Sub
  
  ' Set the delimiter/encapsulator as required
  Select Case cboFileFormat.ItemData(cboFileFormat.ListIndex)
    Case ImportType.DelimitedFile:
      'If selected then Delimited and Encapsulation needs to be visible/enabled
      cboDelimiter.Enabled = True
      cboDelimiter.BackColor = &H80000005
      
      lblDelimiter.Enabled = True
      lblDelimiter.BackColor = &H80000005
      
      lblEncapsulator.Enabled = True
      txtEncapsulator.Enabled = True
      txtEncapsulator.BackColor = &H80000005
      
      If cboDelimiter.Text = "<Other>" Then
        'If <Other> is selected as a delimiter choice...
        lblOtherDelimiter.Enabled = True
        txtDelimiter.Enabled = True
        txtDelimiter.BackColor = &H80000005
      End If
      
    Case ImportType.FixedLengthFile
      SetComboText cboDelimiter, ","
      cboDelimiter.Enabled = False
      cboDelimiter.BackColor = &H8000000F
      
      lblDelimiter.Enabled = False
      lblDelimiter.BackColor = &H8000000F

      txtEncapsulator.Enabled = False
      lblEncapsulator.Enabled = False
      txtEncapsulator.BackColor = &H8000000F
      
      If cboDelimiter.Enabled = False Then
        lblOtherDelimiter.Enabled = False
        txtDelimiter.Enabled = False
        txtDelimiter.BackColor = &H8000000F
        txtDelimiter.Text = ""
      End If
      
    Case ImportType.ExcelWorksheet
      SetComboText cboDelimiter, ","
      cboDelimiter.Enabled = False
      cboDelimiter.BackColor = &H8000000F
      
      lblDelimiter.Enabled = False
      lblDelimiter.BackColor = &H8000000F
      
      txtEncapsulator.Enabled = False
      lblEncapsulator.Enabled = False
      txtEncapsulator.BackColor = &H8000000F
      
      If cboDelimiter.Enabled = False Then
        lblOtherDelimiter.Enabled = False
        txtDelimiter.Enabled = False
        txtDelimiter.BackColor = &H8000000F
        txtDelimiter.Text = ""
      End If
        
  End Select
 
  ' Store the selected fileformat - needed for the first line of this sub
  mintCurrentFileFormat = cboFileFormat.ItemData(cboFileFormat.ListIndex)
      
  'AE20071105 Fault #12553
  txtFilename.Text = vbNullString

  ResetFillers (mintCurrentFileFormat)
  RefreshOptionsFrame
  
  Changed = True
  
End Sub

Private Sub chkUseUpdateBlob_Click()
  Changed = True
End Sub

Private Sub cmdCancel_Click()
  Dim objExpression As clsExprExpression

  ' If the definition was a copy and the user is cancelling then
  ' we need to delete the filter expression copy that was made.
  If mblnFromCopy And (mlngOriginalFilterID > 0) Then
    ' Instantiate a new expression object.
    Set objExpression = New clsExprExpression
    
    With objExpression
      ' Initialise the expression object.
      If .Initialise(0, mlngOriginalFilterID, giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
        .DeleteExpression
      End If
    End With
  
    Set objExpression = Nothing
  End If

Unload Me

End Sub

Private Sub cmdClearColumn_Click()
  
  'Purpose : Check the user really wishes to clear the column grid.
  'Input   : None
  'Output  : None
  
  If COAMsgBox("Clear all import columns/fillers." & vbCrLf & "Are you sure ?", vbOKCancel + vbQuestion, "Import") = vbOK Then
    grdColumns.RemoveAll
    UpdateButtonStatus
    Changed = True
  
    ' Clear the file filter as it is tied to the selected columns.
    optAllRecords.Value = True
  End If
  
End Sub


Private Sub cmdDeleteColumn_Click()
  
  'Purpose : Remove the selected row from the import grid.
  'Input   : None
  'Output  : None
   Dim lRow As Long

  lRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
  If lRow = -1 Then
    grdColumns.MoveLast
    grdColumns.SelBookmarks.Add grdColumns.Bookmark
  End If


  If (txtFilter.Tag > 0) And (grdColumns.Columns("ColExprID").CellText(grdColumns.Bookmark) > 0) Then
    If ColumnInFilter(grdColumns.Columns("ColExprID").CellText(grdColumns.Bookmark), txtFilter.Tag) Then
      COAMsgBox "The column cannot be removed." & vbCr & _
        "It is used in the File Filter definition.", _
        vbExclamation + vbOKOnly, Application.Name
      Exit Sub
    End If
  End If
  
  If Me.grdColumns.Rows = 1 Then
    Me.grdColumns.RemoveAll
  Else
    ' Store the row to be deleted
    lRow = grdColumns.AddItemRowIndex(grdColumns.Bookmark)
    
    ' Delete the row
    grdColumns.RemoveItem grdColumns.AddItemRowIndex(grdColumns.Bookmark)

    If grdColumns.Rows > 0 Then
      If lRow < grdColumns.Rows Then
        grdColumns.SelBookmarks.Add grdColumns.Bookmark ' grdColumns.GetBookmark(lRow)
      ElseIf lRow = grdColumns.Rows Then
        grdColumns.MoveLast
        grdColumns.SelBookmarks.Add grdColumns.Bookmark
      End If
    End If
  
  End If
  
  UpdateButtonStatus
    
  Changed = True
  
End Sub

Private Sub cmdEditColumn_Click()

  'Purpose : Edit the current row in the column grid.
  'Input   : None
  'Output  : None
  
  Dim pstrRow As String
  Dim plngRow As Long
  Dim pfrmColumnEdit As frmImportColumns
  Dim lngOriginalColumnID  As Long
  Dim sColumnName As String
  Dim blnColumnRemoved As Boolean
  
  Screen.MousePointer = vbHourglass
  Set pfrmColumnEdit = New frmImportColumns
  
  lngOriginalColumnID = 0
  sColumnName = ""
  
  With grdColumns
    plngRow = .AddItemRowIndex(.Bookmark)
        
    Select Case .Columns("Type").Text
      'JPD20010907 No Fault - changed the Key column from text to checkbox.
      'Case "C": pfrmColumnEdit.Initialise False, "C", .Columns("TableID").Value, .Columns("ColExprID").Value, IIf(.Columns("Key").Value = "Y", True, False), .Columns("Size").Value, Me
      Case "C":
        'pfrmColumnEdit.Initialise False, "C", .Columns("TableID").Value, .Columns("ColExprID").Value, IIf(UCase(.Columns("Key").Value) = "TRUE", True, False), .Columns("Size").Value
        pfrmColumnEdit.Initialise False, "C", .Columns("TableID").Value, .Columns("ColExprID").Value, IIf(UCase(.Columns("Key").Value) = "TRUE", True, False), .Columns("Size").Value, IIf(UCase(.Columns("LookupEntries").Value) = "TRUE", True, False), Me
        lngOriginalColumnID = .Columns("ColExprID").Value
        sColumnName = .Columns("Table Name").Value & "." & .Columns("Column Name").Value
      Case "F":
        pfrmColumnEdit.Initialise False, "F", 0, 0, 0, .Columns("Size").Value, False, Me
    End Select
  End With
  
  With pfrmColumnEdit
    .Show vbModal
    
    If Not .Cancelled Then
      Changed = True

      'MH20010926 Fault 2867
      'Prevent runtime error when editting a filler and clicking ok...
      If lngOriginalColumnID > 0 Then
        blnColumnRemoved = .optFiller
        If .cboColumn.ListIndex >= 0 Then
          blnColumnRemoved = (blnColumnRemoved Or _
              (lngOriginalColumnID <> .cboColumn.ItemData(.cboColumn.ListIndex)))
        End If

        If blnColumnRemoved And ColumnInFilter(lngOriginalColumnID, txtFilter.Tag) Then
          COAMsgBox "The '" & sColumnName & "' column cannot be removed." & vbCr & _
            "It is used in the File Filter definition.", _
            vbExclamation + vbOKOnly, Application.Name
          Exit Sub
        End If
      End If


      If .optFiller Then
        'JPD20010907 No Fault - changed the Key column from text to checkbox.
        'pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>" & vbTab & vbTab & "N" & vbTab & .txtLength.Text
        'pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>" & vbTab & vbTab & "false" & vbTab & .txtLength.Text & vbTab & "false"
        'NHRD23072003 Fault 6257 Swapped Key and Size columns
        pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>" & vbTab & vbTab & .txtLength.Text & vbTab & "false" & vbTab & "false"
      Else
        'JPD20010907 No Fault - changed the Key column from text to checkbox.
        'pstrRow = "C" & vbTab & .cboTable.ItemData(.cboTable.ListIndex) & vbTab & .cboColumn.ItemData(.cboColumn.ListIndex) & vbTab & .cboTable.Text & vbTab & .cboColumn.Text & vbTab & IIf(.chkKey.Value, "Y", "N") & vbTab & .txtLength.Text
        'pstrRow = "C" & vbTab & .cboTable.ItemData(.cboTable.ListIndex) & vbTab & .cboColumn.ItemData(.cboColumn.ListIndex) & vbTab & .cboTable.Text & vbTab & .cboColumn.Text & vbTab & IIf(.chkKey.Value, "true", "false") & vbTab & .txtLength.Text & vbTab & IIf(.chkLookupEntries.Value, "true", "false")
        'NHRD23072003 Fault 6257 Swapped Key and Size columns
        pstrRow = "C" & vbTab & .cboTable.ItemData(.cboTable.ListIndex) & vbTab & .cboColumn.ItemData(.cboColumn.ListIndex) & vbTab & .cboTable.Text & vbTab & .cboColumn.Text & vbTab & .txtLength.Text & vbTab & IIf(.chkKey.Value, "true", "false") & vbTab & IIf(.chkLookupEntries.Value, "true", "false")
      End If

      With grdColumns
        
        .RemoveItem plngRow
        .AddItem pstrRow, plngRow
        .SelBookmarks.Add .AddItemBookmark(plngRow)
        
        ' RH 29/08/00 - BUG 858
        .Bookmark = .AddItemBookmark(plngRow)
        
      End With
    
    End If
  
  End With
  
  Unload pfrmColumnEdit
  Set pfrmColumnEdit = Nothing
  UpdateButtonStatus

End Sub

Private Sub cmdFileName_Click()

  'Purpose : Show common dialog box to allow user to select a file to import
  With CDialog
  
    ' If there is a filename already select it, try and set the cdialog directory and filename property to match it.
    If Len(Trim(txtFilename.Text)) = 0 Then
      .InitDir = gsDocumentsPath
    Else
      .FileName = txtFilename.Text
    End If
    
    ' Set flags
    .CancelError = False
    .DialogTitle = "File To Import..."
    .Flags = &H200806

    Select Case cboFileFormat.ItemData(cboFileFormat.ListIndex)
      Case 0
        .Filter = "Comma Separated Values (*.csv)|*.csv|Text (*.txt)|*.txt|All Files|*.*"
      Case 1
        .Filter = "Text (*.txt)|*.txt|All Files|*.*"
      Case 2

        InitialiseCommonDialogFormats CDialog, "Excel", GetOfficeExcelVersion, DirectionInput
    End Select
    
    
    ' Show the dialog
    .ShowOpen
    
    ' Saftey first !
    If Len(.FileName) > 256 Then
      COAMsgBox "Path & file name must not exceed 256 characters in length.", vbExclamation + vbOKOnly, "Import"
      Exit Sub
    End If
    
    ' If they select an Excel file but not Excel Worksheet in the list disallow it
    If cboFileFormat.ItemData(cboFileFormat.ListIndex) <> 2 And (Right(.FileName, 3) = "xls" Or Right(.FileName, 3) = "xlsx") Then
      COAMsgBox "Cannot select an Excel file unless Excel Worksheet is selected as the file format.", vbExclamation + vbOKOnly, "Import"
      Exit Sub
    End If
    
    'AE20071105 Fault #12553
    ' If something was selected, then update the text box with the selected filename
    If .FileName <> "" Then
      txtFilename.Text = .FileName
    End If

  End With
  
  Changed = True
  
End Sub

Public Function Initialise(pblnNew As Boolean, pblnCopy As Boolean, Optional plngImportID As Long) As Boolean
  
  ' Set reference to data access class module
  Set mdatData = New DataMgr.clsDataAccess
  
  Screen.MousePointer = vbHourglass

  If pblnNew Then
    mlngImportID = 0
    
    'Clear fields and set username
    ClearForNew False
    
    'Load All Possible Base Tables into combo
    LoadCombos
  
    GetObjectCategories cboCategory, utlExport, 0, cboBaseTable.ItemData(cboBaseTable.ListIndex)
    SetComboItem cboCategory, IIf(glngCurrentCategoryID = -1, 0, glngCurrentCategoryID)
  
    mblnDefinitionCreator = True
  
    PopulateAccessGrid
    Changed = False
  Else
    ' Make the ImportID visible to the rest of the module
    mlngImportID = plngImportID
    
    ' Is is a copy of an existing one ?
    mblnFromCopy = pblnCopy
    
    If Not RetrieveImportDetails(mlngImportID) Then
      If COAMsgBox("OpenHR could not load all of the definition successfully. The recommendation is that" & vbCrLf & _
             "you delete the definition and create a new one, however, you may edit the existing" & vbCrLf & _
             "definition if you wish. Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, "Import") = vbNo Then
        Initialise = False
        Changed = False
        Exit Function
      End If
    End If
  End If
    
  mlngOriginalFilterID = txtFilter.Tag
  
  PopulateAccessGrid
  RefreshOptionsFrame
    
  'Reset pointer so copy will be saved as new
  If mblnFromCopy Then
    mlngImportID = 0
    Changed = True
  Else
    Changed = False
  End If
    
  mblnLoading = False
  Screen.MousePointer = vbDefault
  Initialise = True
  
End Function


Private Sub LoadCombos()

  'Purpose : Populate the base table combo with all tables, then the fileformat combo
  
  On Error GoTo LoadCombos_ERROR
  
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
    '.ListIndex = 0
    If .ListCount > 0 Then
      If gsPersonnelTableName <> "" Then
        SetComboText cboBaseTable, gsPersonnelTableName
      Else
        .ListIndex = 0
      End If
    End If
  End With

  With cboFileFormat
    
    .AddItem "Delimited File"
    .ItemData(.NewIndex) = 0

    .AddItem "Fixed Length File"
    .ItemData(.NewIndex) = 1
    
    .AddItem "Excel Worksheet"
    .ItemData(.NewIndex) = 2
        
    .AddItem "Linked Server"
    .ItemData(.NewIndex) = 3
    
    .ListIndex = 0
  End With
  
  pstrSQL = vbNullString
  Set prstTables = Nothing
  
  Exit Sub
  
LoadCombos_ERROR:
  
  pstrSQL = vbNullString
  Set prstTables = Nothing
  COAMsgBox "Error populating the combo boxes." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Import"
 
End Sub

Private Sub cmdFilter_Click()
  GetFilter txtFilter

End Sub

Private Sub GetFilter(ctlTarget As Control)
  'Purpose : Show the expression.dll form and populate the relevant control
  'Input   : Target control - used to know which tags/text/listindex
  '          properties to set once an expression has been selected/cleared.
  'Output  : None

  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  Dim intLoop As Integer
  Dim alngColumns() As Long
  Dim varBookmark As Variant
  Dim fCancelled As Boolean
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    fOK = .Initialise(0, Val(ctlTarget.Tag), giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC)
      
    If fOK Then
      ' Construct an array of the columns in the import definition.
      ReDim alngColumns(0)
      grdColumns.MoveFirst
    
      Do Until intLoop = grdColumns.Rows
        varBookmark = grdColumns.GetBookmark(intLoop)
    
        If grdColumns.Columns("ColExprID").CellText(varBookmark) > 0 Then
          ReDim Preserve alngColumns(UBound(alngColumns) + 1)
          alngColumns(UBound(alngColumns)) = grdColumns.Columns("ColExprID").CellText(varBookmark)
        End If
      
        intLoop = intLoop + 1
      Loop
    
      .ColumnList = alngColumns
    End If

    If fOK Then
      If Val(ctlTarget.Tag) > 0 Then
        'MH20050809 Fault 9991
        .AccessOverride = IIf(mblnReadOnly, "RO", "RW")
        .EditExpression fCancelled
      Else
        .NewExpression fCancelled
      End If

      ' Read the selected expression info.
      ctlTarget.Text = IIf(.ExpressionID > 0, .Name, "<None>")
      
      If (ctlTarget.Tag <> .ExpressionID) Or (Not fCancelled) Then
        'Changed = True
        Changed = Not mblnReadOnly
      End If

      ctlTarget.Tag = .ExpressionID
      mlngOriginalFilterID = .ExpressionID
    End If
  End With
  
  Set objExpression = Nothing
  
End Sub



Private Sub cmdMoveDown_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  Dim lngCount As Long

  With grdColumns
  
    intSourceRow = .AddItemRowIndex(.Bookmark)
    'strSourceRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text & vbTab & .Columns(5).Text & vbTab & .Columns(6).Text
    strSourceRow = vbNullString
    For lngCount = 0 To grdColumns.Columns.Count - 1
      strSourceRow = strSourceRow & _
          IIf(lngCount > 0, vbTab, "") & _
          grdColumns.Columns(lngCount).Text
    Next
    
    intDestinationRow = intSourceRow + 1
    .MoveNext
    'strDestinationRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text & vbTab & .Columns(5).Text & vbTab & .Columns(6).Text
    strDestinationRow = vbNullString
    For lngCount = 0 To grdColumns.Columns.Count - 1
      strDestinationRow = strDestinationRow & _
          IIf(lngCount > 0, vbTab, "") & _
          grdColumns.Columns(lngCount).Text
    Next
    
    .RemoveItem intDestinationRow
    .RemoveItem intSourceRow
    
    .AddItem strDestinationRow, intSourceRow
    .AddItem strSourceRow, intDestinationRow
    
    .Bookmark = .AddItemBookmark(intDestinationRow)
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
    '.SelBookmarks.RemoveAll
    '.MoveNext
    '.SelBookmarks.Add .AddItemBookmark(intDestinationRow)
  
  End With
  
  UpdateButtonStatus
  Changed = True
  
End Sub

Private Sub cmdMoveUp_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  Dim lngCount As Long
  
  
  With grdColumns
  
    intSourceRow = .AddItemRowIndex(.Bookmark)
    'strSourceRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text & vbTab & .Columns(5).Text & vbTab & .Columns(6).Text
    strSourceRow = vbNullString
    For lngCount = 0 To grdColumns.Columns.Count - 1
      strSourceRow = strSourceRow & _
          IIf(lngCount > 0, vbTab, "") & _
          grdColumns.Columns(lngCount).Text
    Next
    
    intDestinationRow = intSourceRow - 1
    .MovePrevious
    'strDestinationRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text & vbTab & .Columns(5).Text & vbTab & .Columns(6).Text
    strDestinationRow = vbNullString
    For lngCount = 0 To grdColumns.Columns.Count - 1
      strDestinationRow = strDestinationRow & _
          IIf(lngCount > 0, vbTab, "") & _
          grdColumns.Columns(lngCount).Text
    Next
    
    .AddItem strSourceRow, intDestinationRow
    
    .RemoveItem intSourceRow + 1

    .Bookmark = .AddItemBookmark(intDestinationRow)
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
    '.MovePrevious
    '.MovePrevious
  
  End With
  
  UpdateButtonStatus
  Changed = True
  
End Sub

'Private Sub DisableLengthColumn(iFileType As Integer)
'
'  'Greys out the length column if not fixed length file type.
'  If iFileType <> 1 Then
'    Me.grdColumns.Columns(5).ForeColor = vbGrayText
'    Me.grdColumns.Columns(5).HeadForeColor = vbGrayText
'  Else
'    Me.grdColumns.Columns(5).ForeColor = vbWindowText
'    Me.grdColumns.Columns(5).HeadForeColor = vbWindowText
'  End If
'
'End Sub

Private Sub cmdNewColumn_Click()
  
  'Purpose : Add new row in the column grid.
  'Input   : None
  'Output  : None

  Dim pstrRow As String
  Dim pfrmColumnEdit As frmImportColumns
  
  Set pfrmColumnEdit = New frmImportColumns
  
  With pfrmColumnEdit
    
    If .Initialise(True, "", 0, 0, 0, 0, 0, Me) = True Then
    
      .Show vbModal
      
      If Not .Cancelled Then
        
        Changed = True
        
        If .optFiller Then
          'JPD20010907 No Fault - changed the Key column from text to checkbox.
          'pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>" & vbTab & vbTab & "N" & vbTab & .txtLength.Text
          pstrRow = "F" & vbTab & "0" & vbTab & "0" & vbTab & "<Filler>" & vbTab & vbTab & .txtLength.Text & vbTab & "false" & vbTab & "false"
        Else
          'JPD20010907 No Fault - changed Key column from text to checkbox
          'pstrRow = "C" & vbTab & .cboTable.ItemData(.cboTable.ListIndex) & vbTab & .cboColumn.ItemData(.cboColumn.ListIndex) & vbTab & .cboTable.Text & vbTab & .cboColumn.Text & vbTab & IIf(.chkKey.Value, "Y", "N") & vbTab & .txtLength.Text
          pstrRow = "C" & vbTab & .cboTable.ItemData(.cboTable.ListIndex) & vbTab & .cboColumn.ItemData(.cboColumn.ListIndex) & vbTab & .cboTable.Text & vbTab & .cboColumn.Text & vbTab & .txtLength.Text & vbTab & IIf(.chkKey.Value, "true", "false") & vbTab & IIf(.chkLookupEntries.Value, "true", "false")
        End If
        
        With grdColumns
          .AddItem pstrRow
          .MoveLast
          .SelBookmarks.Add .Bookmark
        End With
        
        cmdEditColumn.Enabled = True
        cmdDeleteColumn.Enabled = True
        cmdClearColumn.Enabled = True
      
      End If
  
     End If
     
  End With
  
  Unload pfrmColumnEdit
  Set pfrmColumnEdit = Nothing
  UpdateButtonStatus
  
End Sub

Private Sub cmdOK_Click()

  If Not ValidateDefinition Then Exit Sub
  If Not SaveDefinition Then Exit Sub
  
  Me.Hide
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    Case KeyCode = 192
        KeyCode = 0
  End Select
End Sub

Private Sub Form_Load()

  tabImport.Tab = 0
  grdAccess.RowHeight = 239
  'DisableLengthColumn 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer
  Dim objExpression As clsExprExpression
  
  If Changed = True Then
    
    pintAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, "Import")
      
    If pintAnswer = vbYes Then
      cmdOK_Click
      Cancel = True
      Exit Sub
    ElseIf pintAnswer = vbCancel Then
      Cancel = True
      Exit Sub
    Else
      ' If the definition was a copy and the user is cancelling then
      ' we need to delete the filter expression copy that was made.
      If mblnFromCopy And (mlngOriginalFilterID > 0) Then
        ' Instantiate a new expression object.
        Set objExpression = New clsExprExpression
        
        With objExpression
          ' Initialise the expression object.
          If .Initialise(0, mlngOriginalFilterID, giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
            .DeleteExpression
          End If
        End With
      
        Set objExpression = Nothing
      End If
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

    If ((Not mblnDefinitionCreator) Or mblnReadOnly) Or _
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
    If (Not mblnDefinitionCreator) Or mblnReadOnly Then
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


Private Sub grdColumns_AfterColUpdate(ByVal ColIndex As Integer)
  Changed = True
End Sub




Private Sub grdColumns_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub


Private Sub grdColumns_DblClick()
  If cmdEditColumn.Enabled Then
    cmdEditColumn_Click
  End If
End Sub

Private Sub grdColumns_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  Dim iLoop As Integer
  
  If Not mblnReadOnly Then
  
    With grdColumns
      ' Set the styleSet of the rows to show which is selected.
      For iLoop = 0 To .Rows
        If iLoop = .Row Then
          .Columns(0).CellStyleSet "ssetActive", iLoop
          .Columns(1).CellStyleSet "ssetActive", iLoop
          .Columns(2).CellStyleSet "ssetActive", iLoop
          .Columns(3).CellStyleSet "ssetActive", iLoop
          .Columns(4).CellStyleSet "ssetActive", iLoop
          .Columns(5).CellStyleSet "ssetActive", iLoop
'          .Columns(6).CellStyleSet "ssetActive", iLoop
        Else
          .Columns(0).CellStyleSet "ssetDormant", iLoop
          .Columns(1).CellStyleSet "ssetDormant", iLoop
          .Columns(2).CellStyleSet "ssetDormant", iLoop
          .Columns(3).CellStyleSet "ssetDormant", iLoop
          .Columns(4).CellStyleSet "ssetDormant", iLoop
          .Columns(5).CellStyleSet "ssetDormant", iLoop
 '         .Columns(6).CellStyleSet "ssetDormant", iLoop
        End If
      Next iLoop

      If .AddItemRowIndex(.Bookmark) = 0 Then
        Me.cmdMoveUp.Enabled = False
        Me.cmdMoveDown.Enabled = (.Rows > 1)
      ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
        Me.cmdMoveUp.Enabled = (.Rows > 1)
        Me.cmdMoveDown.Enabled = False
      Else
        Me.cmdMoveUp.Enabled = True
        Me.cmdMoveDown.Enabled = True
      End If
      
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End With
    
  End If
    
  UpdateButtonStatus

End Sub




Private Sub optAllRecords_Click()
  Changed = True

  With txtFilter
    .Text = ""
    .Tag = 0
  End With

  cmdFilter.Enabled = False

End Sub


Private Sub optDontUpdateAny_Click()

  Changed = True

End Sub

Private Sub optFilter_Click()
  Changed = True

  With txtFilter
    .Text = ""
    .Tag = 0
  End With

  cmdFilter.Enabled = True

End Sub

Private Sub optImportType_Click(Index As Integer)

  If optImportType(0).Value = True Then
    lblDupRecordsReturned.Enabled = False
    optUpdateAll.Enabled = False
    optDontUpdateAny.Enabled = False
  Else
    lblDupRecordsReturned.Enabled = True
    optUpdateAll.Enabled = True
    optDontUpdateAny.Enabled = True
  End If

  If Not mblnLoading Then Changed = True

End Sub



Private Sub optUpdateAll_Click()

  Changed = True
  
End Sub

Private Sub spnFooterLines_Change()
  Changed = True
End Sub

Private Sub spnHeaderLines_Change()
  Changed = True
End Sub

Private Sub tabImport_Click(PreviousTab As Integer)

  Dim ctl As Control

  If Not mblnReadOnly Then
    For Each ctl In Me.Controls
      If TypeOf ctl Is VB.Frame Then
        ctl.Enabled = ctl.Left >= 0
      End If
    Next
    UpdateButtonStatus
  End If

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

Private Sub txtDesc_Change()
  Changed = True
End Sub

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

Private Sub txtEncapsulator_Change()

  Changed = True
  If txtEncapsulator = " " Then
    lblEncapsulatorHint.Caption = "(Space)"
  ElseIf Len(txtEncapsulator) = 0 Then
    lblEncapsulatorHint.Caption = "(None)"
  Else
    lblEncapsulatorHint.Caption = ""
  End If
  
End Sub

Private Sub txtEncapsulator_GotFocus()

  With txtEncapsulator
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
End Sub

Private Sub txtLinkedCatalog_Change()
  Changed = True
End Sub

Private Sub txtLinkedServer_Change()
  Changed = True
End Sub

Private Sub txtLinkedTable_Change()
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

Private Sub UpdateButtonStatus()

  'Purpose : Updates all command button status depending on grid.rows etc
  'Input   : None
  'Output  : None
  
  Dim lngWidth As Long
  Dim lngRow As Long
  
  If mblnReadOnly Then Exit Sub

  With grdColumns
    
    lngWidth = .Width
    lngWidth = lngWidth - .Columns("Create Lookup").Width
    
    'MH20030815
    'Visible rows is set to zero when editting a definition???
    'If .VisibleRows < .Rows Then
    If 13 < .Rows Then
      .ScrollBars = ssScrollBarsVertical
      lngWidth = lngWidth - 256

      lngRow = .Rows - .VisibleRows
      If .AddItemRowIndex(.FirstRow) > lngRow Then
        .FirstRow = .AddItemBookmark(lngRow)
      End If
    Else
      .ScrollBars = ssScrollBarsNone
      .FirstRow = .AddItemBookmark(0)
    End If

    If .Columns("Size").Visible = True Then
      .Columns("Size").Width = 630 '600
      lngWidth = lngWidth - .Columns("Size").Width
    End If

    .Columns("Key").Width = 509
    lngWidth = lngWidth - .Columns("Key").Width

    .Columns("Table Name").Width = lngWidth / 2
    .Columns("Column Name").Width = lngWidth / 2

  End With

  If grdColumns.Rows = 0 Then
    cmdEditColumn.Enabled = False
    cmdDeleteColumn.Enabled = False
    cmdClearColumn.Enabled = False
    cmdMoveUp.Enabled = False
    cmdMoveDown.Enabled = False
    Exit Sub
  Else
    If grdColumns.AddItemRowIndex(grdColumns.Bookmark) > 0 Or grdColumns.SelBookmarks.Count > 0 Then
      cmdEditColumn.Enabled = True
      cmdDeleteColumn.Enabled = True
      cmdClearColumn.Enabled = True
    End If
  End If

  If grdColumns.AddItemRowIndex(grdColumns.Bookmark) = 0 Then
    Me.cmdMoveUp.Enabled = False
    Me.cmdMoveDown.Enabled = (grdColumns.Rows > 1)
  ElseIf grdColumns.AddItemRowIndex(grdColumns.Bookmark) = (grdColumns.Rows - 1) Then
    Me.cmdMoveUp.Enabled = (grdColumns.Rows > 1)
    Me.cmdMoveDown.Enabled = False
  Else
    Me.cmdMoveUp.Enabled = True
    Me.cmdMoveDown.Enabled = True
  End If
  
  grdColumns.SelBookmarks.RemoveAll
  grdColumns.SelBookmarks.Add grdColumns.Bookmark

End Sub

Private Function ValidateDefinition() As Boolean

  'Purpose : Check all mandatory information is entered and also check that
  '          If there is a problem with validation, the program will display
  '          the tab containing the problem to the user.
  'Input   : None
  'Output  : True/False
  
  On Error GoTo ValidateDefinition_ERROR
  
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean

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
    tabImport.Tab = 0
    COAMsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  'Check if this definition has been changed by another user
  Call UtilityAmended(utlImport, mlngImportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngImportID = 0
    mblnReadOnly = False
    ForceAccess
  End If
  
  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngImportID) Then
    tabImport.Tab = 0
    COAMsgBox "An Import definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Check a filename has been selected
  If Len(Trim(txtFilename.Text)) = 0 And cboFileFormat.ItemData(cboFileFormat.ListIndex) <> ImportType.SQLTable Then
    tabImport.Tab = 2
    COAMsgBox "You must select the file you wish to use with this import definition.", vbExclamation + vbOKOnly, "Import"
    cmdFilename.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Check that there are columns defined in the definition
  If grdColumns.Rows = 0 Then
    tabImport.Tab = 1
    COAMsgBox "You must select at least 1 column for your import.", vbExclamation + vbOKOnly, "Import"
    ValidateDefinition = False
    Exit Function
  End If
  
  '  Check that a delimiter is specified if the file format is ASCII Delimited
  If cboFileFormat.ItemData(cboFileFormat.ListIndex) = ImportType.DelimitedFile Then
    If cboDelimiter.Text = "<Other>" And Trim(txtDelimiter.Text) = "" Then
      tabImport.Tab = 2
      COAMsgBox "You must specify a delimiter for delimited files.", vbExclamation + vbOKOnly, "Import"
      ValidateDefinition = False
      Exit Function
    End If
  End If

  ' Check that if the file format is defined as Fixed Length, that there is a size
  ' defined for each column in the import definition.
  If cboFileFormat.ItemData(cboFileFormat.ListIndex) = ImportType.FixedLengthFile Then
    ' Loop thru the import grid, checking for size
    With grdColumns
      .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("Size").CellText(pvarbookmark) = "" Then
          tabImport.Tab = 1
          COAMsgBox "You have specified this import definition as a fixed file format import, but" & vbCrLf & _
            "not all columns defined have been given a size. Please ensure all" & vbCrLf & _
            "columns in the definition have sizes for fixed length imports.", vbExclamation + vbOKOnly, "Import"
          ValidateDefinition = False
          Exit Function
        End If
        pintLoop = pintLoop + 1
      Loop
    End With
  End If

  ' SQL Table specific
  If cboFileFormat.ItemData(cboFileFormat.ListIndex) = ImportType.SQLTable Then
    If txtLinkedServer.Text = "" Or txtLinkedCatalog.Text = "" Or txtLinkedTable.Text = "" Then
      COAMsgBox "You must specify all options for importing from a linked table.", vbExclamation + vbOKOnly, "Import"
      tabImport.Tab = 2
      ValidateDefinition = False
      Exit Function
    End If
  End If

  ' Now check that there is at least 1 key field defined.
  If CheckForKeyFields = False Then
    tabImport.Tab = 1
    COAMsgBox "You must define at least one key field on the '" & cboBaseTable.Text & "' table to update records. " & _
           "If you are creating new records only, then please select " & _
           "the Create New Records Only option.", vbExclamation + vbOKOnly, "Import"
    ValidateDefinition = False
    Exit Function
  End If
    
  ' Check that they are not all key fields in the definition
  If CheckNotAllKeyFields = False Then
    COAMsgBox "You must define at least one field that is not a key field.", vbExclamation + vbOKOnly, "Import"
    ValidateDefinition = False
    Exit Function
  End If
    
  ' If using a filter, check one has been selected
  If optFilter.Value Then
    If txtFilter.Text = "" Or txtFilter.Tag = "0" Then
      COAMsgBox "You must select a filter, or change the record selection for your file.", vbExclamation + vbOKOnly, "Import"
      tabImport.Tab = 2
      cmdFilter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  If optImportType(1).Value = False Then  'not "Update record Only"
    If Not CheckMandatoryColumns Then
      ValidateDefinition = False
      Exit Function
    End If
  End If

If mlngImportID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlImport, _
      CStr(mlngImportID), _
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
               vbExclamation + vbOKOnly, "Import"
      Else
        COAMsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Import"
      End If

      Screen.MousePointer = vbDefault
      tabImport.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If COAMsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Import") = vbNo Then
        Screen.MousePointer = vbDefault
        tabImport.Tab = 0
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
  
  COAMsgBox "Error whilst validating import definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Import"
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





Private Function CheckForKeyFields() As Boolean
  
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim pblnNoKeys As Boolean
  Dim lngBaseTableID As Long
  
  'MH20010816 Fault 2017
  'If Me.chkCreateNewOnly.Value Then
  If optImportType(0).Value Then
    CheckForKeyFields = True
    Exit Function
  End If
  
  pintLoop = 0
  lngBaseTableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)

  With grdColumns
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)

      'MH20040419 S000047 - Check for keyed fields on the base table...
      If Val(.Columns("tableid").CellValue(pvarbookmark)) = lngBaseTableID Then

      'MH20030822 Fault 6820 - When referencing the "Key" column use cellvalue rather than celltext
      ''''JPD20010907 No Fault - changed Key column from text to checkbox.
      ''''If .Columns("key").CellText(pvarbookmark) = "Y" Then
      '''If .Columns("key").CellText(pvarbookmark) = True Then
      If .Columns("key").CellValue(pvarbookmark) = True Then
        CheckForKeyFields = True
        Exit Function
      End If
      
      End If
      
      pintLoop = pintLoop + 1
    Loop
  End With

  CheckForKeyFields = False

End Function

Private Function CheckNotAllKeyFields() As Boolean
  
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim pblnNoKeys As Boolean
  
  pintLoop = 0
  
  With grdColumns
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      
      'MH20060926 Fault
      If .Columns("type").CellValue(pvarbookmark) = "C" Then
      
      'MH20030822 Fault 6820 - When referencing the "Key" column use cellvalue rather than celltext
      ''''JPD20010907 No Fault - changed the Key column from text to checkbox.
      ''''If .Columns("key").CellText(pvarbookmark) = "N" Then
      '''If .Columns("key").CellText(pvarbookmark) = False Then
        If .Columns("key").CellValue(pvarbookmark) = False Then
          CheckNotAllKeyFields = True
          Exit Function
        End If
      
      End If
      pintLoop = pintLoop + 1
    Loop
  End With

  CheckNotAllKeyFields = False

End Function

Private Function SaveDefinition() As Boolean

  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  Dim sSQL As String
  Dim lCount As Long
  Dim lImportID As Long
  Dim rsImport As New Recordset
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim lngImportType As Long
  
  On Error GoTo Err_Trap
    
  'Set Import Definition access string
  If optImportType(1).Value = True Then
    lngImportType = 1
  ElseIf optImportType(2).Value = True Then
    lngImportType = 2
  Else
    lngImportType = 0
  End If
  
  
  If mlngImportID > 0 Then

    'We are updating an existing import definition

    'First save the basic definition

    sSQL = "UPDATE ASRSysImportName SET " & _
             "Name = '" & Trim(Replace(Me.txtName.Text, "'", "''")) & "'," & _
             "Description = '" & Replace(Me.txtDesc.Text, "'", "''") & "'," & _
             "BaseTable = " & Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) & "," & _
             "FileType = " & Me.cboFileFormat.ItemData(cboFileFormat.ListIndex) & "," & _
             "FileName = '" & Replace(Me.txtFilename.Text, "'", "''") & "'," & _
             "Delimiter = '" & cboDelimiter.Text & "'," & _
             "OtherDelimiter = '" & Replace(Me.txtDelimiter.Text, "'", "''") & "'," & _
             "DateFormat = '" & cboDateFormat.Text & "'," & _
             "Encapsulator = " & IIf(Len(Me.txtEncapsulator.Text) = 0, "Null", "'" & Replace(Me.txtEncapsulator.Text, "'", "''") & "'") & "," & _
             "MultipleRecordAction = " & IIf(Me.optUpdateAll.Value, 1, 0) & "," & _
             "HeaderLines = " & CStr(spnHeaderLines.Value) & "," & _
             "FooterLines = " & CStr(spnFooterLines.Value) & "," & _
             "FilterID = " & IIf(optFilter.Value, txtFilter.Tag, 0) & ","
            
    sSQL = sSQL & "LinkedServer = '" & Replace(txtLinkedServer.Text, "'", "''") & "'," & _
             "LinkedCatalog = '" & Replace(txtLinkedCatalog.Text, "'", "''") & "'," & _
             "LinkedTable = '" & Replace(txtLinkedTable.Text, "'", "''") & "', " & _
             "UseUpdateBlob = " & IIf(chkUseUpdateBlob.Value, "1", "0")
            
    'TM20020726 Fault 2123 - Upadte the DateSeparator column.
    'MH20010816 Fault 2017
    sSQL = sSQL & "," & _
             "ImportType = " & CStr(lngImportType) & "," & _
              "DateSeparator = '" & cboDateSeparator.Text & "' " & _
              "WHERE ID = " & mlngImportID

     mdatData.ExecuteSql (sSQL)
    
    Call UtilUpdateLastSaved(utlImport, mlngImportID)

  Else

    ' Adding a new import definition
    sSQL = "Insert ASRSysImportName (" & _
           "Name, Description, BaseTable, " & _
           "FileType, FileName, Delimiter, OtherDelimiter, DateFormat, " & _
           "Encapsulator, MultipleRecordAction, " & _
           "HeaderLines, FooterLines, ImportType, " & _
           "UserName, FilterID, DateSeparator, LinkedServer, LinkedCatalog, LinkedTable, UseUpdateBlob) "
    
    sSQL = sSQL & _
           "Values('" & _
           Trim(Replace(txtName.Text, "'", "''")) & "','" & _
           Replace(txtDesc.Text, "'", "''") & "'," & _
           cboBaseTable.ItemData(cboBaseTable.ListIndex)

    sSQL = sSQL & ", " & Me.cboFileFormat.ItemData(cboFileFormat.ListIndex)
    sSQL = sSQL & ", '" & Replace(Me.txtFilename.Text, "'", "''") & "'"
    sSQL = sSQL & ", '" & cboDelimiter.Text & "'"
    sSQL = sSQL & ", '" & Replace(Me.txtDelimiter.Text, "'", "''") & "'"
    sSQL = sSQL & ", '" & Me.cboDateFormat.Text & "'"
    sSQL = sSQL & ", " & IIf(Len(Me.txtEncapsulator.Text) = 0, "Null", "'" & Replace(Me.txtEncapsulator.Text, "'", "''") & "'")
    sSQL = sSQL & ", " & IIf(Me.optUpdateAll.Value, 1, 0)
    sSQL = sSQL & ", " & CStr(spnHeaderLines.Value)
    sSQL = sSQL & ", " & CStr(spnFooterLines.Value)
    sSQL = sSQL & ", " & CStr(lngImportType)
    sSQL = sSQL & ", '" & datGeneral.UserNameForSQL & "'"
    sSQL = sSQL & ", " & txtFilter.Tag
    sSQL = sSQL & ", '" & cboDateSeparator.Text & "'"
    sSQL = sSQL & ", '" & Replace(txtLinkedServer.Text, "'", "''") & "'"
    sSQL = sSQL & ", '" & Replace(txtLinkedCatalog.Text, "'", "''") & "'"
    sSQL = sSQL & ", '" & Replace(txtLinkedTable.Text, "'", "''") & "'"
    sSQL = sSQL & ", " & IIf(chkUseUpdateBlob.Value, "1", "0") & ")"

    mlngImportID = InsertImport(sSQL)

    If mlngImportID = 0 Then
      SaveDefinition = False
      Exit Function
    End If

    Call UtilCreated(utlImport, mlngImportID)

  End If

  SaveAccess
  SaveObjectCategories cboCategory, utlImport, mlngImportID
  
  ' Now save the column details

  ' First, remove any records from the detail table with the specified ID
  ClearDetailTables mlngImportID

  ' Loop through the details grid, and also the sortorder grid
  With grdColumns

    .MoveFirst

    Do Until pintLoop = .Rows

      pvarbookmark = .GetBookmark(pintLoop)

      sSQL = "INSERT ASRSysImportDetails (" & _
             "ImportID, " & _
             "Type, " & _
             "TableID, " & _
             "ColExprID, " & _
             "KeyField, " & _
             "Size, " & _
             "LookupEntries) "

      sSQL = sSQL & "VALUES(" & mlngImportID & ", "

      sSQL = sSQL & "'" & .Columns("Type").CellText(pvarbookmark) & "', "
      sSQL = sSQL & .Columns("TableID").CellText(pvarbookmark) & ", "
      sSQL = sSQL & .Columns("ColExprID").CellText(pvarbookmark) & ", "
      
      
      
      
      'MH20030822 Fault 6820 - When referencing the "Key" column use cellvalue rather than celltext
      '''''JPD20010907 No Fault - changed the Key column from text to checkbox.
      '''''sSQL = sSQL & IIf(.Columns("Key").CellText(pvarbookmark) = "Y", 1, 0) & ", "
      ''''sSQL = sSQL & IIf(UCase(.Columns("Key").CellText(pvarbookmark)) = "TRUE", 1, 0) & ", "
      '''sSQL = sSQL & IIf(.Columns("Key").CellText(pvarbookmark) = True, 1, 0) & ", "
      sSQL = sSQL & IIf(.Columns("Key").CellValue(pvarbookmark) = True, 1, 0) & ", "
      
      
      
      'TM20011219 Fault 3039 - don't save the size if not fixed length filetype.
      sSQL = sSQL & IIf((.Columns("Size").CellText(pvarbookmark) = "" Or _
                        cboFileFormat.ItemData(cboFileFormat.ListIndex) <> 1), "Null", .Columns("Size").CellText(pvarbookmark)) & ", "

      sSQL = sSQL & IIf(.Columns("LookupEntries").CellText(pvarbookmark) = True, 1, 0) & ")"
      
      pintLoop = pintLoop + 1

      mdatData.ExecuteSql (sSQL)

    Loop

  End With
  
  SaveDefinition = True
  Changed = False

  If (mlngOriginalFilterID > 0) And (mlngOriginalFilterID <> txtFilter.Tag) Then
    ' Delete the filter if there is one.
  
    ' Instantiate a new expression object.
    Set objExpression = New clsExprExpression
    
    With objExpression
      ' Initialise the expression object.
      fOK = .Initialise(0, mlngOriginalFilterID, giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC)
      
      If fOK Then
        .DeleteExpression
      End If
    End With
  
    Set objExpression = Nothing
  End If
  
  Exit Function

Err_Trap:
  
  COAMsgBox "Error whilst saving Import definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Import"
  SaveDefinition = False

End Function

Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysImportAccess WHERE ID = " & mlngImportID
  mdatData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysImportAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngImportID & ", sysusers.name," & _
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
  mdatData.ExecuteSql (sSQL)

  ' Update the new access records with the real access values.
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      sSQL = "IF EXISTS (SELECT * FROM ASRSysImportAccess" & _
        " WHERE ID = " & CStr(mlngImportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysImportAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngImportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      mdatData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub





Private Function CheckUniqueName(sName As String, lngCurrentID As Long) As Boolean
  
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  
  sSQL = "SELECT * FROM ASRSysImportName " & _
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



Private Function InsertImport(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertImport_ERROR
  
'  Dim rsImport As ADODB.Recordset
'  mdatData.ExecuteSql pstrSQL
'  pstrSQL = "Select Max(ID) From ASRSysImportName"
'  Set rsImport = mdatData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)
'  InsertImport = rsImport(0)
'  rsImport.Close
'  Set rsImport = Nothing
'  Exit Function
  

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
    pmADO.Value = "AsrSysImportName"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "ID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, app.ProductName
        InsertImport = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertImport = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertImport_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function

Private Sub ClearDetailTables(plngImportID As Long)

  ' Delete all column information from the Details table.
  
  Dim pstrSQL As String
  
  pstrSQL = "Delete From ASRSysImportDetails Where ImportID = " & plngImportID
  mdatData.ExecuteSql pstrSQL

End Sub

Private Sub ClearForNew(Optional bPartialClear As Boolean)
  
  'Purpose : Clear out all fields required to be blank for a new definition
  'Input   : Optional True/False for if its a partial or complete clear up
  'Output  : None
  
  With Me
    
    .grdColumns.RemoveAll
    
    If bPartialClear Then Exit Sub
    
    .txtName = vbNullString
    .txtDesc = vbNullString
    .txtUserName = gsUserName
    .txtFilename.Text = ""
    .optUpdateAll.Value = True
    'AE20071004
'    .chkIgnoreFirstLine.Value = 0
'    .chkIgnoreLastLine.Value = 0
    .spnHeaderLines.Value = 0
    .spnFooterLines.Value = 0
    '.chkCreateNewOnly.Value = 0
    optImportType(2).Value = True
    .optAllRecords.Value = True
    .txtFilter.Tag = 0
    .txtFilter.Text = ""
    cboDateFormat.ListIndex = 0
    'AE20071005 Fault #9615
    'cboDateSeparator.ListIndex = 0
    SetComboText cboDateSeparator, UI.GetSystemDateSeparator
    If cboDateSeparator.ListIndex = -1 Then cboDateSeparator.ListIndex = 0
    
    'AE20071004 Fault 12487
    .cboDelimiter.ListIndex = 0

  End With
  
End Sub

Private Function RetrieveImportDetails(plngImportID As Long) As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim pintLoop As Integer
  Dim pstrText As String
  Dim objExpression As clsExprExpression
  
  On Error GoTo Load_ERROR
  
  'Load the basic guff first
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysImportName.*, " & _
                                     "CONVERT(integer, ASRSysImportName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysImportName WHERE ID = " & plngImportID)

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "Cannot load the definition for this import." & vbCrLf & "(" & Err.Description & ")", vbCritical + vbOKOnly, "Import"
    Set rsTemp = Nothing
    RetrieveImportDetails = False
    Exit Function
  End If
  
  ' Set Definition Description
  txtDesc.Text = IIf(IsNull(rsTemp!Description), "", rsTemp!Description)
    
  Changed = False
  
  ' Set Base Table
  LoadCombos
  SetComboText cboBaseTable, datGeneral.GetTableName(rsTemp!BaseTable)
  mstrBaseTable = cboBaseTable.Text
  
  ' Set the categories combo
  GetObjectCategories cboCategory, utlImport, plngImportID
  
  Select Case rsTemp!filetype
    Case ImportType.DelimitedFile: SetComboText cboFileFormat, "Delimited File"
    Case ImportType.FixedLengthFile: SetComboText cboFileFormat, "Fixed Length File"
    Case ImportType.ExcelWorksheet: SetComboText cboFileFormat, "Excel Worksheet"
    Case ImportType.SQLTable: SetComboText cboFileFormat, "Linked Server"
  End Select
  
  ''TM20011219 Fault 3039 - disable the length column if required.
  'DisableLengthColumn (rsTemp!filetype)

  txtFilename.Text = rsTemp!FileName
  
  'AE20071004 Fault 12488
  'SetComboText cboDelimiter, rsTemp!delimiter
  SetComboText cboDelimiter, rsTemp!delimiter, True
  
  Me.txtDelimiter.Text = rsTemp!otherdelimiter
  
  txtEncapsulator.Text = IIf(IsNull(rsTemp!encapsulator), "", rsTemp!encapsulator)
   
  If Not IsNull(rsTemp!DateFormat) Then SetComboText cboDateFormat, rsTemp!DateFormat
  
  'TM20020726 Fault 2123
  If Not IsNull(rsTemp!DateFormat) Then
    SetComboText cboDateSeparator, rsTemp!dateseparator
  Else
    SetComboText cboDateFormat, "<None>"
  End If
  
  ' Options
  If rsTemp!MultipleRecordAction Then optUpdateAll.Value = True Else optDontUpdateAny.Value = True
  'If rsTemp!ignorefirstline Then chkIgnoreFirstLine.Value = 1
  'If rsTemp!ignorelastline Then chkIgnoreLastLine.Value = 1
  spnHeaderLines.Value = IIf(IsNull(rsTemp!HeaderLines), 0, rsTemp!HeaderLines)
  spnFooterLines.Value = IIf(IsNull(rsTemp!FooterLines), 0, rsTemp!FooterLines)


  'MH20010816 Fault 2017
  'If rsTemp!CreateNewOnly Then chkCreateNewOnly.Value = 1
  optImportType(rsTemp!ImportType).Value = True
  
  txtLinkedServer.Text = IIf(IsNull(rsTemp!LinkedServer), "", rsTemp!LinkedServer)
  txtLinkedCatalog.Text = IIf(IsNull(rsTemp!LinkedCatalog), "", rsTemp!LinkedCatalog)
  txtLinkedTable.Text = IIf(IsNull(rsTemp!LinkedTable), "", rsTemp!LinkedTable)
  
  If Not IsNull(rsTemp!UseUpdateBlob) Then
    If rsTemp!UseUpdateBlob Then chkUseUpdateBlob.Value = vbChecked
  End If
  
  ' Set name, username, access etc
  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If
  
  mblnReadOnly = Not datGeneral.SystemPermission("IMPORT", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlImport, plngImportID) = ACCESS_READONLY)
  End If

  If mblnReadOnly Then
    ControlsDisableAll Me
    grdColumns.Enabled = True
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  mlngTimeStamp = rsTemp!intTimestamp
    
  If IsNull(rsTemp!FilterID) Then
    optAllRecords.Value = True
  Else
    If rsTemp!FilterID > 0 Then
      optFilter.Value = True
    Else
      optAllRecords.Value = True
    End If
  End If
  cmdFilter.Enabled = optFilter.Value
  txtFilter.Tag = IIf(IsNull(rsTemp!FilterID), 0, rsTemp!FilterID)
  txtFilter.Text = IIf(txtFilter.Tag > 0, datGeneral.GetFilterName(txtFilter.Tag), "")
    
  If mblnFromCopy And (txtFilter.Tag > 0) Then
    ' Need to copy the filter expression.
    ' Instantiate a new expression object.
    Set objExpression = New clsExprExpression
    
    With objExpression
      ' Initialise the expression object.
      .Initialise 0, Val(txtFilter.Tag), giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC
      .QuickCopyExpression (txtFilter.Text)
      
      txtFilter.Tag = .ExpressionID
    End With
    Set objExpression = Nothing
  End If
  
  ' =========================
  
  ' Now load the details
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysImportDetails WHERE ImportID = " & plngImportID & " ORDER BY ID")
  
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "Error loading the column definition for this import." & _
           IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", vbNullString), vbCritical + vbOKOnly, "Import"
    Set rsTemp = Nothing
    RetrieveImportDetails = False
    Exit Function
  End If

  Do Until rsTemp.EOF

    pstrText = rsTemp!Type & vbTab & rsTemp!TableID & vbTab & rsTemp!ColExprID & vbTab & IIf(rsTemp!TableID > 0, datGeneral.GetTableName(rsTemp!TableID), "<Filler>") & vbTab

    If rsTemp!ColExprID = 0 Then
      pstrText = pstrText & ""
    Else
      pstrText = pstrText & datGeneral.GetColumnName(rsTemp!ColExprID)
    End If

    'TM20011219 Fault 3039 - If the column does has a tableid it must be a
    'column so show the size specified (!size), if the import file type is not fixed
    'length there will be not size saved so get the default column size.
    If rsTemp!TableID > 0 Then
      pstrText = pstrText & vbTab & IIf(IsNull(rsTemp!Size), datGeneral.GetDataSize(rsTemp!ColExprID, True), rsTemp!Size)
    Else
      pstrText = pstrText & vbTab & IIf(IsNull(rsTemp!Size), "", rsTemp!Size)
    End If

    'JPD20010907 No Fault - change the Key column to be a checkbox rather than a Y/N column.
    'pstrText = pstrText & vbTab & IIf(rsTemp!KeyField, "Y", "N") & vbTab & IIf(IsNull(rsTemp!Size), "", rsTemp!Size)
    pstrText = pstrText & vbTab & IIf(rsTemp!KeyField, "true", "false")

    pstrText = pstrText & vbTab & IIf(rsTemp!LookupEntries, "true", "false")

    Me.grdColumns.AddItem pstrText

    rsTemp.MoveNext
  Loop

  UpdateButtonStatus

  ' Tidyup
  Set rsTemp = Nothing
  RetrieveImportDetails = True
  Exit Function

Load_ERROR:

  'COAMsgBox "Error whilst retrieving the import definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Import"
  RetrieveImportDetails = False
  Set rsTemp = Nothing

End Function

Private Sub RefreshOptionsFrame()

  Dim outType As ImportType
  outType = cboFileFormat.ItemData(cboFileFormat.ListIndex)

  fraTableDetails.Visible = (outType = ImportType.SQLTable)
  fraFileDetails.Visible = Not fraTableDetails.Visible
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
  Set rsAccess = GetUtilityAccessRecords(utlImport, mlngImportID, mblnFromCopy)
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


Private Sub ResetFillers(iFileType As Integer)

  'Sets the filler columns length to "", as the length is meaningless for non fixed
  'length files.
  Dim iCount As Integer
    
  If iFileType <> 1 Then
    grdColumns.MoveFirst
    For iCount = 0 To grdColumns.Rows Step 1
      If UCase(grdColumns.Columns(0).Text) = "F" Then
        grdColumns.Columns(6).Text = ""
      End If
      grdColumns.MoveNext
    Next iCount
  End If
  
End Sub

Private Function CheckMandatoryColumns() As Boolean
        
  Dim rsColumns As ADODB.Recordset
  Dim strSQL As String

  Dim lngTableIDs() As Long
  Dim intIndex As Integer
  Dim lngRow As Long
  Dim pvarbookmark As Variant
  Dim lngTableID As Integer
  Dim blnFoundTable As Boolean

  Dim strMandatoryColumns As String
  Dim strTableIDs As String
  Dim strColumnIDs As String

  Dim strMBText As String
  
  strMandatoryColumns = vbNullString
  
  ReDim lngTableIDs(0) As Long
  strTableIDs = vbNullString
  strColumnIDs = vbNullString
  
  With grdColumns
    '.Row = 0
    .MoveFirst
    For lngRow = 0 To .Rows - 1

      'pvarbookmark = .GetBookmark(lngRow)
      
      'If .Columns("Key").CellText(pvarBookmark) = "N" Then
      
        lngTableID = Val(.Columns("TableID").CellText(pvarbookmark))
        'If lngTableID > 0 Then
        If lngTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex) Then
        
          strColumnIDs = strColumnIDs & _
            IIf(strColumnIDs <> "", ", ", "") & .Columns("ColExprID").CellText(pvarbookmark)
  
          'Loop though all of the columns in the array and
          'check to see if this table is already in the array
          blnFoundTable = False
          For intIndex = 0 To UBound(lngTableIDs)
            blnFoundTable = (lngTableIDs(intIndex) = lngTableID)
            If blnFoundTable Then
              Exit For
            End If
          Next
  
          'If this table is not in the array then
          'add this table to the array
          If blnFoundTable = False Then
            intIndex = IIf(lngTableIDs(0) = 0, 0, UBound(lngTableIDs) + 1)
            ReDim Preserve lngTableIDs(intIndex) As Long
            lngTableIDs(intIndex) = lngTableID
          
            strTableIDs = strTableIDs & _
              IIf(strTableIDs <> "", ", ", "") & CStr(lngTableID)
          End If
        
        End If
      
      'End If
    
      .MoveNext
    Next
  End With

  'MH20000814
  'Allow save if mandatory ommitted if it has a default value
  'This is to get around the staff number on a applicants to personnel transfer
  
  'MH20000904
  'Allow save if mandatory ommitted and it is a calculated column

  '******************************************************************************
  ' TM20010719 Fault 2242 - ColumnType <> 4 clause added to ignore all linked   *
  ' columns. (It doesn't need to validate the linked columns because this is    *
  ' done using the Vaidate SP.                                                  *
  '******************************************************************************

  If strTableIDs <> vbNullString Then
    strSQL = "SELECT ASRSysTables.TableName, ASRSysColumns.ColumnName " & _
             "FROM ASRSysColumns " & _
             "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
             "WHERE ASRSysColumns.TableID IN (" & strTableIDs & ") " & _
             "  AND ASRSysColumns.ColumnID NOT IN (" & strColumnIDs & ") " & _
             "  AND " & SQLWhereMandatoryColumn & _
             " ORDER BY ASRSysTables.TableName, ASRSysColumns.ColumnName"
             '"  AND Mandatory = '1' " & _
             "  AND Rtrim(DefaultValue) = '' AND Convert(int,dfltValueExprID) = 0 " & _
             "  AND CalcExprID = 0 " & _
             "  AND ColumnType <> 4 "
    Set rsColumns = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
    
    While Not rsColumns.EOF
      strMandatoryColumns = strMandatoryColumns & _
        rsColumns!TableName & "." & rsColumns!ColumnName & vbCrLf
      rsColumns.MoveNext
    Wend

  End If

  CheckMandatoryColumns = (strMandatoryColumns = vbNullString)
  
'  If CheckMandatoryColumns = False Then
'    strMBText = "Unable to save definition as the following mandatory" & vbCrLf & _
'                "columns have not been populated:" & vbCrLf & vbCrLf & _
'                strMandatoryColumns & vbCrLf & _
'                "Please add these columns to the import definition."
'
'    COAMsgBox strMBText, vbExclamation + vbOKOnly, "Import"
'
'  End If

' RH 15/03/01 - Bug 1967
  If CheckMandatoryColumns = False Then
    strMBText = "The following columns are mandatory but have not been" & vbCrLf & _
                "included in the definition:" & vbCrLf & vbCrLf & _
                strMandatoryColumns & vbCrLf & _
                "New records will not be imported without these columns." & vbCrLf & _
                "Do you wish to continue ?"

    If COAMsgBox(strMBText, vbQuestion + vbYesNo, "Import") = vbYes Then CheckMandatoryColumns = True

  End If

End Function


'###################################
Public Sub PrintDef(lImportID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As ADODB.Recordset
  Dim rsColumns As ADODB.Recordset
  Dim sSQL As String
  Dim lngTempX As Long
  Dim lngTempY As Long
  Dim sTemp As String
  Dim iLoop As Integer
  Dim varBookmark As Variant

  mlngImportID = lImportID
  
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysImportName.*, " & _
                                     "CONVERT(integer, ASRSysImportName.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysImportName WHERE ID = " & mlngImportID)
                                        
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Print Definition"
    Set rsTemp = Nothing
    Exit Sub
  End If

  PopulateAccessGrid
  
  ' JDM - 20/08/01 - Fault 2701 - Asking user to save after printing definition
  Me.Changed = False

  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader "Import : " & rsTemp!Name
    
        .PrintNormal "Category : " & GetObjectCategory(utlImport, mlngImportID)
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
        
        ' Now do the Columns Section
      
        Set rsColumns = datGeneral.GetRecords("SELECT * FROM ASRSysImportDetails WHERE ImportID = " & mlngImportID & " ORDER BY ID")
        .PrintTitle "Columns"
        
        Do While Not rsColumns.EOF
            
          .PrintNormal "Type : " & IIf(rsColumns!Type = "C", "Column", "Filler")
          
          If rsColumns!Type = "C" Then
            .PrintNormal "Table : " & datGeneral.GetTableName(rsColumns!TableID)
            .PrintNormal "Column : " & datGeneral.GetColumnName(rsColumns!ColExprID)
          Else
            .PrintNormal "Table : N/A"
            .PrintNormal "Column : N/A"
          End If
          
          .PrintNormal "Key : " & IIf(rsColumns!KeyField = True, "Yes", "No")
          .PrintNormal "Length : " & IIf(IsNull(rsColumns!Size), "N/A", rsColumns!Size)
          'AE20071005 Fault #9069
          .PrintNormal "Create missing lookup table entries : " & IIf(rsColumns!LookupEntries = True, "Yes", "No")

          .PrintNormal
          
          rsColumns.MoveNext
        
        Loop
        
        ' Now do the Options Section
      
        .PrintTitle "Options"
        
        Select Case rsTemp!filetype
          Case ImportType.DelimitedFile: .PrintNormal "Import Type : Delimited File"
          Case ImportType.FixedLengthFile: .PrintNormal "Import Type : Fixed Length File"
          Case ImportType.ExcelWorksheet: .PrintNormal "Import Type : Excel Worksheet"
          Case ImportType.SQLTable: .PrintNormal "Import Type : Linked Server"
        End Select
        

        If rsTemp!filetype = ImportType.SQLTable Then
          .PrintNormal "Server : " & rsTemp!LinkedServer
          .PrintNormal "Catalog : " & rsTemp!LinkedCatalog
          .PrintNormal "Table : " & rsTemp!LinkedTable
        Else
        
          .PrintNormal "File Name : " & rsTemp!FileName
          
          If rsTemp!filetype = 0 Then
            .PrintNormal "Delimiter : " & IIf(rsTemp!delimiter = "<Other>", rsTemp!otherdelimiter, rsTemp!delimiter)
          Else
            .PrintNormal "Delimiter : " & "<None>"
          End If
          
          If Len(rsTemp!encapsulator) > 0 Then
            If rsTemp!encapsulator = Chr(34) Then
              If rsTemp!filetype = 0 Then
                .PrintNormal "Data Enclosed In : " & Chr(34)
              Else
                .PrintNormal "Data Enclosed In : <None>"
              End If
            Else
              If rsTemp!filetype = 0 Then
                .PrintNormal "Data Enclosed In : " & IIf(rsTemp!encapsulator = "", "<None>", rsTemp!encapsulator)
              Else
                .PrintNormal "Data Enclosed In : <None>"
              End If
            End If
          Else
            .PrintNormal "Data Enclosed In : <None>"
          End If
          
          .PrintNormal "Date Format : " & IIf(IsNull(rsTemp!DateFormat), "(Not Specified)", rsTemp!DateFormat)
          
          .PrintNormal "Date Separator : " & IIf(IsNull(rsTemp!dateseparator), "(Not Specified)", rsTemp!dateseparator)
          
          'MH20071003
          .PrintNormal "Header Lines To Ignore : " & IIf(IsNull(rsTemp!HeaderLines), 0, rsTemp!HeaderLines)
          .PrintNormal "Footer Lines To Ignore : " & IIf(IsNull(rsTemp!FooterLines), 0, rsTemp!FooterLines)
        
        End If
        
        If rsTemp!FilterID > 0 Then
          .PrintNormal "File Records : '" & datGeneral.GetFilterName(rsTemp!FilterID) & "' filter"
        Else
          .PrintNormal "File Records : All"
        End If
        
        'MH20010816 Fault 2017
        '.PrintNormal "Create New Records Only : " & IIf(rsTemp!CreateNewOnly = True, "Yes", "No")
        '.PrintNormal "Mulitple Record Action : " & IIf(rsTemp!MultipleRecordAction = True, "Create/Update all records found", "Dont create/update any records")
        .PrintNormal "Type : " & Replace(Me.optImportType(rsTemp!ImportType).Caption, "&", "")
        If rsTemp!ImportType <> 0 Then
          .PrintNormal "Mulitple Record Action : " & IIf(rsTemp!MultipleRecordAction = True, "Update All", "Update None")
        End If
    
        .PrintEnd
        .PrintConfirm "Import : " & rsTemp!Name, "Import Definition"
      End If
  
    End With
    
  End If
  
  Set rsTemp = Nothing
  Set rsColumns = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing Import Definition Failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"

End Sub

Private Sub cboCategory_Click()
  Changed = True
End Sub




