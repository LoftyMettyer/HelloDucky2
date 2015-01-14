VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmTabEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Table"
   ClientHeight    =   6015
   ClientLeft      =   315
   ClientTop       =   1665
   ClientWidth     =   8130
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
   HelpContextID   =   5033
   Icon            =   "frmTabEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSort 
      Height          =   450
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   58
      Top             =   5505
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTabTableProperties 
      Height          =   5295
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "De&finition"
      TabPicture(0)   =   "frmTabEdit.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRecordDescription"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblOrder"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTableName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblEmail"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdRecordDescription"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtRecordDescription"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOrder"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOrder"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraTableType"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTableName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdEmail"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtEmail"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkCopyWhenParentRecordIsCopied"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Su&mmary"
      TabPicture(1)   =   "frmTabEdit.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblParentTable"
      Tab(1).Control(1)=   "fraColumns"
      Tab(1).Control(2)=   "cmdInsert"
      Tab(1).Control(3)=   "cmdDown"
      Tab(1).Control(4)=   "cmdUp"
      Tab(1).Control(5)=   "cmdRemove"
      Tab(1).Control(6)=   "cmdAdd"
      Tab(1).Control(7)=   "fraSummaryFields"
      Tab(1).Control(8)=   "cmdInsertBreak"
      Tab(1).Control(9)=   "cboParentTable"
      Tab(1).Control(10)=   "cmdColumnBreak"
      Tab(1).Control(11)=   "chkManualColumnBreak"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Ema&ils"
      TabPicture(2)   =   "frmTabEdit.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraEmail"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Calen&dars"
      TabPicture(3)   =   "frmTabEdit.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraCalendarLinks"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Workflows"
      TabPicture(4)   =   "frmTabEdit.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraWorkflowLinks"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Audi&t"
      TabPicture(5)   =   "frmTabEdit.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraAudit"
      Tab(5).Control(1)=   "fraTableStats"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "&Validation"
      TabPicture(6)   =   "frmTabEdit.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraTableValidations"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Tri&ggers"
      TabPicture(7)   =   "frmTabEdit.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraTableTriggers"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "fraSystemTriggers"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      Begin VB.CheckBox chkCopyWhenParentRecordIsCopied 
         Caption         =   "Cop&y when parent record is copied"
         Enabled         =   0   'False
         Height          =   315
         Left            =   195
         TabIndex        =   75
         Top             =   3840
         Width           =   4035
      End
      Begin VB.Frame fraSystemTriggers 
         Caption         =   "System Triggers :"
         Height          =   1515
         Left            =   -74820
         TabIndex        =   71
         Top             =   3570
         Width           =   7560
         Begin VB.CheckBox chkDisableDelete 
            Caption         =   "Disable Instead of Delete"
            Height          =   195
            Left            =   225
            TabIndex        =   74
            Top             =   1110
            Width           =   2505
         End
         Begin VB.CheckBox chkDisableUpdate 
            Caption         =   "Disable Update"
            Height          =   210
            Left            =   225
            TabIndex        =   73
            Top             =   750
            Width           =   2175
         End
         Begin VB.CheckBox chkDisableInsert 
            Caption         =   "Disable Instead of Insert"
            Height          =   195
            Left            =   225
            TabIndex        =   72
            Top             =   390
            Width           =   2610
         End
      End
      Begin VB.Frame fraTableTriggers 
         Caption         =   "Backbone Code :"
         Height          =   3060
         Left            =   -74800
         TabIndex        =   65
         Top             =   405
         Width           =   7560
         Begin VB.CommandButton cmdTableTriggerEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   400
            Left            =   2190
            TabIndex        =   69
            Top             =   2340
            Width           =   1200
         End
         Begin VB.CommandButton cmdTableTriggerDeleteAll 
            Caption         =   "Delete &All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   6030
            TabIndex        =   68
            Top             =   2340
            Width           =   1200
         End
         Begin VB.CommandButton cmdTableTriggerDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   400
            Left            =   4140
            TabIndex        =   67
            Top             =   2340
            Width           =   1200
         End
         Begin VB.CommandButton cmdTableTriggerNew 
            Caption         =   "&New"
            Enabled         =   0   'False
            Height          =   400
            Left            =   330
            TabIndex        =   66
            Top             =   2340
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid ssGrdTableTriggers 
            Height          =   1860
            Left            =   330
            TabIndex        =   70
            Top             =   300
            Width           =   6915
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            DefColWidth     =   26458
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   3
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   26458
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "TriggerID"
            Columns(0).Name =   "TriggerID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   12197
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "TriggerName"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   26458
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "Type"
            Columns(2).Name =   "TriggerType"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   12197
            _ExtentY        =   3281
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
      Begin VB.Frame fraTableValidations 
         Caption         =   "Overlapping Column Validations :"
         Height          =   4710
         Left            =   -74800
         TabIndex        =   59
         Top             =   400
         Width           =   7560
         Begin VB.CommandButton cmdTableValidationNew 
            Caption         =   "&New"
            Enabled         =   0   'False
            Height          =   400
            Left            =   330
            TabIndex        =   63
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdTableValidationDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   400
            Left            =   4140
            TabIndex        =   62
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdTableValidationDeleteAll 
            Caption         =   "Delete &All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   6030
            TabIndex        =   61
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdTableValidationEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   400
            Left            =   2190
            TabIndex        =   60
            Top             =   4125
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid ssGrdTableValidations 
            Height          =   3660
            Left            =   330
            TabIndex        =   64
            Top             =   300
            Width           =   6915
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   2
            DefColWidth     =   26458
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   3
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   26458
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ValidationID"
            Columns(0).Name =   "ValidationID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   26458
            Columns(1).Caption=   "Description"
            Columns(1).Name =   "Description"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   12197
            _ExtentY        =   6456
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
      Begin VB.Frame fraEmail 
         Caption         =   "Email Links :"
         Height          =   4710
         Left            =   -74800
         TabIndex        =   30
         Top             =   400
         Width           =   7560
         Begin VB.CommandButton cmdEmailLinkProperties 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   2190
            TabIndex        =   33
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveAllEmailLinks 
            Caption         =   "Delete &All"
            Height          =   400
            Left            =   6030
            TabIndex        =   35
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveEmailLink 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   4140
            TabIndex        =   34
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddEmailLink 
            Caption         =   "&New"
            Enabled         =   0   'False
            Height          =   400
            Left            =   330
            TabIndex        =   32
            Top             =   4125
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid ssGrdEmailLinks 
            Height          =   3660
            Left            =   330
            TabIndex        =   31
            Top             =   300
            Width           =   6915
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   4
            DefColWidth     =   26458
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   3
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   4
            Columns(0).Width=   5027
            Columns(0).Caption=   "Name"
            Columns(0).Name =   "colTitle"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   6694
            Columns(1).Caption=   "Email Activation"
            Columns(1).Name =   "colOffset"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   26458
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "Subject"
            Columns(2).Name =   "colSubject"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   26458
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "colLinkID"
            Columns(3).Name =   "colLinkID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   12197
            _ExtentY        =   6456
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
      Begin VB.Frame fraAudit 
         Caption         =   "Audit Log :"
         Height          =   1050
         Left            =   -74800
         TabIndex        =   48
         Top             =   400
         Width           =   7340
         Begin VB.CheckBox chkAuditInsertion 
            Caption         =   "Audit &Addition"
            Height          =   240
            Left            =   200
            TabIndex        =   49
            Top             =   300
            Width           =   1500
         End
         Begin VB.CheckBox chkAuditDeletion 
            Caption         =   "Audit &Deletion"
            Height          =   300
            Left            =   200
            TabIndex        =   50
            Top             =   600
            Width           =   1500
         End
      End
      Begin VB.Frame fraWorkflowLinks 
         Caption         =   "Workflow Links :"
         Height          =   4710
         Left            =   -74800
         TabIndex        =   42
         Top             =   400
         Width           =   7560
         Begin VB.CommandButton cmdRemoveWorkflowLink 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   400
            Left            =   4140
            TabIndex        =   46
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddWorkflowLink 
            Caption         =   "&New"
            Enabled         =   0   'False
            Height          =   400
            Left            =   330
            TabIndex        =   44
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveAllWorkflowLinks 
            Caption         =   "Delete &All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   6030
            TabIndex        =   47
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdWorkflowLinkProperties 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   400
            Left            =   2190
            TabIndex        =   45
            Top             =   4125
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid ssGrdWorkflowLinks 
            Height          =   3660
            Left            =   330
            TabIndex        =   43
            Top             =   300
            Width           =   6915
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            DefColWidth     =   26458
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   3
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   26458
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "LinkID"
            Columns(0).Name =   "LinkID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   10319
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Name"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   1402
            Columns(2).Caption=   "Enabled"
            Columns(2).Name =   "Enabled"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   11
            Columns(2).FieldLen=   256
            Columns(2).Locked=   -1  'True
            Columns(2).Style=   2
            Columns(2).HasBackColor=   -1  'True
            Columns(2).BackColor=   -2147483633
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   12197
            _ExtentY        =   6456
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
      Begin VB.Frame fraTableStats 
         Caption         =   "Table Information : "
         Height          =   3015
         Left            =   -74800
         TabIndex        =   51
         Top             =   1550
         Width           =   7340
         Begin ComctlLib.ListView lstOLEColumns 
            Height          =   1600
            Left            =   200
            TabIndex        =   55
            Top             =   1200
            Width           =   6910
            _ExtentX        =   12197
            _ExtentY        =   2831
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Column Name"
               Object.Width           =   7497
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Size"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label lblEmbeddedOLEInfo 
            Caption         =   "Embedded OLE Information : "
            Height          =   255
            Left            =   180
            TabIndex        =   54
            Top             =   900
            Width           =   3195
         End
         Begin VB.Label lblStatsRows 
            Caption         =   "Rows :"
            Height          =   255
            Left            =   200
            TabIndex        =   52
            Top             =   300
            Width           =   7000
         End
         Begin VB.Label lblDataSize 
            Caption         =   "Table Size : "
            Height          =   255
            Left            =   200
            TabIndex        =   53
            Top             =   600
            Width           =   7000
         End
      End
      Begin VB.Frame fraCalendarLinks 
         Caption         =   "Calendar Links :"
         Height          =   4710
         Left            =   -74800
         TabIndex        =   36
         Top             =   400
         Width           =   7560
         Begin VB.CommandButton cmdOutlookLinkProperties 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   400
            Left            =   2190
            TabIndex        =   39
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveAllOutlookLinks 
            Caption         =   "Delete &All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   6030
            TabIndex        =   41
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveOutlookLink 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   400
            Left            =   4140
            TabIndex        =   40
            Top             =   4125
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddOutlookLink 
            Caption         =   "&New"
            Enabled         =   0   'False
            Height          =   400
            Left            =   330
            TabIndex        =   38
            Top             =   4125
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid ssGrdOutlookLinks 
            Height          =   3660
            Left            =   330
            TabIndex        =   37
            Top             =   300
            Width           =   6915
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
            DefColWidth     =   26458
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   3
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   26458
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "LinkID"
            Columns(0).Name =   "LinkID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   5027
            Columns(1).Caption=   "Name"
            Columns(1).Name =   "Title"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   6694
            Columns(2).Caption=   "Subject"
            Columns(2).Name =   "Subject"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            UseDefaults     =   0   'False
            TabNavigation   =   1
            _ExtentX        =   12197
            _ExtentY        =   6456
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
      Begin VB.CheckBox chkManualColumnBreak 
         Caption         =   "Manual column br&eaks"
         Height          =   240
         Left            =   -69975
         TabIndex        =   18
         Top             =   645
         Width           =   2385
      End
      Begin VB.CommandButton cmdColumnBreak 
         Caption         =   "Colum&n Break >"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -71850
         TabIndex        =   25
         Top             =   3090
         Width           =   1630
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3200
         Width           =   5355
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "..."
         Height          =   315
         Left            =   7455
         TabIndex        =   15
         Top             =   3200
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.ComboBox cboParentTable 
         Height          =   315
         Left            =   -73500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2500
      End
      Begin VB.CommandButton cmdInsertBreak 
         Caption         =   "&Break >"
         Height          =   360
         Left            =   -71850
         TabIndex        =   24
         Top             =   2655
         Width           =   1630
      End
      Begin VB.Frame fraSummaryFields 
         Caption         =   "Summary Columns :"
         Height          =   3500
         Left            =   -69975
         TabIndex        =   28
         Top             =   1000
         Width           =   2745
         Begin VB.ListBox lstSummaryFields 
            Height          =   2985
            Index           =   1
            Left            =   200
            TabIndex        =   29
            Top             =   300
            Width           =   2340
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add  >"
         Height          =   360
         Left            =   -71850
         TabIndex        =   21
         Top             =   1110
         Width           =   1630
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove  <"
         Height          =   360
         Left            =   -71850
         TabIndex        =   23
         Top             =   1980
         Width           =   1630
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Move Up"
         Height          =   360
         Left            =   -71850
         TabIndex        =   26
         Top             =   3675
         UseMaskColor    =   -1  'True
         Width           =   1630
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Move Down"
         Height          =   360
         Left            =   -71850
         Picture         =   "frmTabEdit.frx":00EC
         TabIndex        =   27
         Top             =   4110
         UseMaskColor    =   -1  'True
         Width           =   1630
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert  >"
         Height          =   360
         Left            =   -71850
         TabIndex        =   22
         Top             =   1545
         Width           =   1630
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns :"
         Height          =   3500
         Left            =   -74800
         TabIndex        =   19
         Top             =   1000
         Width           =   2790
         Begin VB.ListBox lstColumns 
            Height          =   2985
            Index           =   1
            Left            =   200
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   300
            Width           =   2340
         End
      End
      Begin VB.TextBox txtTableName 
         Height          =   315
         Left            =   1000
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "txtTabName"
         Top             =   600
         Width           =   6765
      End
      Begin VB.Frame fraTableType 
         Caption         =   "Type :"
         Height          =   800
         Left            =   200
         TabIndex        =   3
         Top             =   1100
         Width           =   7560
         Begin VB.OptionButton optTableType 
            Caption         =   "&Parent"
            Height          =   315
            Index           =   0
            Left            =   200
            TabIndex        =   4
            Top             =   300
            Width           =   900
         End
         Begin VB.OptionButton optTableType 
            Caption         =   "C&hild"
            Height          =   315
            Index           =   1
            Left            =   1700
            TabIndex        =   5
            Top             =   300
            Width           =   800
         End
         Begin VB.OptionButton optTableType 
            Caption         =   "&Lookup"
            Height          =   315
            Index           =   2
            Left            =   3200
            TabIndex        =   6
            Top             =   300
            Width           =   1000
         End
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "..."
         Height          =   315
         Left            =   7455
         TabIndex        =   9
         Top             =   2200
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtOrder 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2200
         Width           =   5355
      End
      Begin VB.TextBox txtRecordDescription 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2700
         Width           =   5355
      End
      Begin VB.CommandButton cmdRecordDescription 
         Caption         =   "..."
         Height          =   315
         Left            =   7455
         TabIndex        =   12
         Top             =   2700
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Email :"
         Height          =   195
         Left            =   200
         TabIndex        =   13
         Top             =   3255
         Width           =   1440
      End
      Begin VB.Label lblParentTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Table :"
         Height          =   195
         Left            =   -74800
         TabIndex        =   16
         Top             =   660
         Width           =   1245
      End
      Begin VB.Label lblTableName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   200
         TabIndex        =   1
         Top             =   660
         Width           =   510
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Order :"
         Height          =   195
         Left            =   200
         TabIndex        =   7
         Top             =   2265
         Width           =   1515
      End
      Begin VB.Label lblRecordDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Description :"
         Height          =   195
         Left            =   200
         TabIndex        =   10
         Top             =   2760
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   5520
      TabIndex        =   56
      Top             =   5490
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6810
      TabIndex        =   57
      Top             =   5490
      Width           =   1200
   End
End
Attribute VB_Name = "frmTabEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Table definition variables.
Private mobjTable As Table
Private mlngOrderID As Long
Private mlngRecDescExprID As Long
Private mlngEmailID As Long
'Private mblnChanged As Boolean
Private mblnUnloadForm As Boolean
Private mbManualColumnBreaks As Boolean
Private mbManualColumnInserted As Boolean

'Private mlngEmailNotification(1) As Long

' Form handling variables.
Private gfCancelled As Boolean

Private mblnNotLoading As Boolean

' Private constants.
Private Const miBREAKID = -1
Private Const miCOLUMNBREAKID = -2
Private Const msBREAKSTRING = "  <Break>"
Private Const msCOLUMNBREAKSTRING = "  <Column Break>"

Private mblnTableViewExists As Boolean
Private mblnReadOnly As Boolean

Private mavColumnInfo() As Variant
Private mfLoading As Boolean

Private Enum TableProperties_TabStrips
  iTABLEPROPERTYTTAB_DEFINITION = 0
  iTABLEPROPERTYTTAB_SUMMARYCOLUMNS = 1
  iTABLEPROPERTYTTAB_EMAILLINKS = 2
  iTABLEPROPERTYTTAB_CALENDARLINKS = 3
  iTABLEPROPERTYTTAB_WORKFLOWLINKS = 4
  iTABLEPROPERTYTTAB_AUDIT = 5
  iTABLEPROPERTYTTAB_VALIDATION = 6
  iTABLEPROPERTYTTAB_TRIGGERS = 7
End Enum

Private mfRebuildWorkflowLinks As Boolean
Private mvarEmailLinks As Collection

Private mblnEmailSortByActivation As Boolean
Private mblnEmailSortDesc As Boolean

Private Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Private Property Let Changed(ByVal blnNewValue As Boolean)
  cmdOK.Enabled = blnNewValue
End Property


Private Sub RefreshSummaryFieldControls()
  ' Refesh the Summary Field controls whose status is variable.
  Dim iLoop As Integer
  Dim iListboxIndex As Integer
  
  iListboxIndex = CurrentListboxIndex
  
  'If (mobjTable.TableType = iTabChild) And _
    (iListboxIndex > 0) Then
  If (mobjTable.TableType = iTabChild) And _
    (iListboxIndex > 0) And Not mblnReadOnly Then
    cmdAdd.Enabled = (lstColumns(iListboxIndex).ListIndex >= 0)
    cmdInsert.Enabled = cmdAdd.Enabled
    cmdRemove.Enabled = (lstSummaryFields(iListboxIndex).ListIndex >= 0)
    cmdUp.Enabled = (lstSummaryFields(iListboxIndex).ListIndex > 0) And (lstSummaryFields(iListboxIndex).ListCount > 1)
    cmdDown.Enabled = (lstSummaryFields(iListboxIndex).ListIndex < (lstSummaryFields(iListboxIndex).ListCount - 1)) And (lstSummaryFields(iListboxIndex).ListCount > 1)
    cmdColumnBreak.Enabled = mbManualColumnBreaks And (mbManualColumnInserted = False)
  Else
    ' Disable all controls for non-child tables.
    For iLoop = 1 To lstColumns.UBound
      lstColumns(iLoop).Enabled = False
    Next iLoop
    For iLoop = 1 To lstSummaryFields.UBound
      lstSummaryFields(iLoop).Enabled = False
    Next iLoop
    cmdAdd.Enabled = False
    cmdInsert.Enabled = False
    cmdInsertBreak.Enabled = False
    cmdRemove.Enabled = False
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    cmdColumnBreak.Enabled = False
  End If

End Sub


Public Property Get Cancelled() As Boolean
  Cancelled = gfCancelled
End Property
Private Sub GetOrderDetails()
  ' Get the Order details.
  Dim objOrder As Order
  
  If mlngOrderID > 0 Then
  
    ' Instantiate a new Order object.
    Set objOrder = New Order
    objOrder.OrderID = mlngOrderID
    
    ' Read the name of the current order.
    If objOrder.ConstructOrder Then
      txtOrder.Text = objOrder.OrderName
    End If
    
    ' Disassociate object variables.
    Set objOrder = Nothing
  Else
    txtOrder.Text = ""
  End If
  
End Sub
Private Sub GetRecordDescriptionDetails()
  ' Get the Record Description details.
  Dim sExprName As String
  
  ' Initialize the default expression name.
  sExprName = ""
  
  If mlngRecDescExprID > 0 Then
  
    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", mlngRecDescExprID, False
    
      ' Read the expression's name from the recordset.
      If Not .NoMatch Then
        sExprName = !Name
      End If
      
    End With
  End If
  
  txtRecordDescription.Text = sExprName

End Sub


Private Sub GetEmailAddressDetails()

  Dim strEmailName As String

  strEmailName = vbNullString

  If mlngEmailID > 0 Then

    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", mlngEmailID

      ' Read the expression's name from the recordset.
      If Not .NoMatch Then
        strEmailName = !Name
      End If

    End With
  End If

  txtEmail = strEmailName

End Sub


Public Property Set Table(pObjTable As Object)
  
  ' Set the edit table property of the form.
  Dim objOutlookLink As clsOutlookLink
  Dim objWorkflowLink As clsWorkflowTriggeredLink
  Dim objTableValidation As clsTableValidation
  Dim objTableTrigger As clsTableTrigger
  Dim fSaved As Boolean
  Dim sAddString As String

  Set mobjTable = pObjTable
  
  ' Determine if the table has been saved yet.
  fSaved = (mobjTable.IsChanged Or Not mobjTable.IsNew)
  mfLoading = True
  mfRebuildWorkflowLinks = False

  If mobjTable.IsNew And Not fSaved Then
    Me.Caption = "New Table Properties"
    txtTableName = vbNullString
    optTableType(1).value = True
    mlngOrderID = 0
    mlngRecDescExprID = 0
    mlngEmailID = 0
    mbManualColumnBreaks = False
  Else
    Me.Caption = "'" & mobjTable.TableName & "' Table Properties" + IIf(mobjTable.Locked, " (Locked)", "")
    txtTableName.Text = mobjTable.TableName
    optTableType(mobjTable.TableType - 1).value = True
    mlngOrderID = mobjTable.PrimaryOrderID
    mlngRecDescExprID = mobjTable.RecordDescriptionID
    mlngEmailID = mobjTable.PrimaryEmailID
    mbManualColumnBreaks = mobjTable.ManualSummaryColumnBreaks
    chkAuditInsertion.value = IIf(mobjTable.AuditInsert = True, vbChecked, vbUnchecked)
    chkAuditDeletion.value = IIf(mobjTable.AuditDelete = True, vbChecked, vbUnchecked)
    'mlngEmailNotification(0) = mobjTable.EmailInsertID
    'mlngEmailNotification(1) = mobjTable.EmailDeleteID


    'MH20090528
  
    Dim iLoop As Integer
    
    For iLoop = 1 To mobjTable.EmailLinks.Count
      mvarEmailLinks.Add mobjTable.EmailLinks.Item(iLoop), "ID" & mobjTable.EmailLinks.Item(iLoop).LinkID
    Next
    PopulateEmailLinks 0
    
    ssGrdOutlookLinks.RemoveAll
    For Each objOutlookLink In mobjTable.OutlookLinks
      With objOutlookLink
        ssGrdOutlookLinks.AddItem .LinkID & vbTab & .Title & vbTab & GetExpressionName(.Subject)
      End With
    Next

    ssGrdWorkflowLinks.RemoveAll
    For Each objWorkflowLink In mobjTable.WorkflowTriggeredLinks
      With objWorkflowLink
        ssGrdWorkflowLinks.AddItem .LinkID & vbTab & GetWorkflowName(.WorkflowID) & vbTab & GetWorkflowEnabled(.WorkflowID)
      End With
    Next objWorkflowLink
    Set objWorkflowLink = Nothing
       
    ssGrdTableValidations.RemoveAll
    For Each objTableValidation In mobjTable.TableValidations
      With objTableValidation
        sAddString = GetValidationString(objTableValidation)
        ssGrdTableValidations.AddItem sAddString
      End With
    Next objTableValidation
    Set objTableValidation = Nothing
    
    ssGrdTableTriggers.RemoveAll
    For Each objTableTrigger In mobjTable.TableTriggers
      With objTableTrigger
        sAddString = GetTriggerString(objTableTrigger)
        ssGrdTableTriggers.AddItem sAddString
      End With
    Next objTableTrigger
    Set objTableTrigger = Nothing
    
   chkDisableInsert.value = IIf(mobjTable.InsertTriggerDisabled, vbChecked, vbUnchecked)
   chkDisableUpdate.value = IIf(mobjTable.UpdateTriggerDisabled, vbChecked, vbUnchecked)
   chkDisableDelete.value = IIf(mobjTable.DeleteTriggerDisabled, vbChecked, vbUnchecked)
    
   chkCopyWhenParentRecordIsCopied.value = IIf(mobjTable.CopyWhenParentRecordIsCopied, vbChecked, vbUnchecked)
    
  End If

  RefreshTableValidationsButtons
  RefreshTableTriggerButtons
  RefreshOutlookLinksButtons
  cboParentTable_Initialise
  mfLoading = False
  
End Property

Private Sub cboParentTable_Click()
  ' Display the required Columns and Summary Fields listboxes.
  Dim fVisible As Boolean
  Dim iLoop As Integer
    
  For iLoop = 1 To lstColumns.UBound
    fVisible = (cboParentTable.ItemData(cboParentTable.ListIndex) = val(lstColumns(iLoop).Tag))
    
    lstColumns(iLoop).Visible = fVisible
    lstSummaryFields(iLoop).Visible = fVisible
  Next iLoop

  RefreshSummaryFieldControls

End Sub

Private Sub chkAuditDeletion_Click()
  Changed = True
End Sub

Private Sub chkAuditInsertion_Click()
  Changed = True
End Sub

Private Sub chkCopyWhenParentRecordIsCopied_Click()
  Changed = True
End Sub

Private Sub chkDisableDelete_Click()
  Changed = True
End Sub

Private Sub chkDisableInsert_Click()
  Changed = True
End Sub

Private Sub chkDisableUpdate_Click()
  Changed = True
End Sub

Private Sub chkManualColumnBreak_Click()

  Dim i As Integer
  
  ' Enable/disable the column break button
  mbManualColumnBreaks = IIf(chkManualColumnBreak.value = vbChecked, True, False)
  
  ' Remove column breaks from the summary columns list
  ' INSERT CODE HERE
  'TM20020501 Fault 3246 - Remove column breaks when tick box unchecked.
  If Not mbManualColumnBreaks Then
    For i = 0 To Me.lstSummaryFields(1).ListCount - 1 Step 1
      If Trim(Me.lstSummaryFields(1).List(i)) = Trim(msCOLUMNBREAKSTRING) Then
        
        If Me.lstSummaryFields(1).ListIndex = i Then
          Me.lstSummaryFields(1).ListIndex = 0
        End If

        Me.lstSummaryFields(1).RemoveItem (i)
        
        MsgBox "The Column Break has been removed from the Summary Columns list.", vbOKOnly + vbInformation, App.Title
        
        mbManualColumnInserted = False
        Exit For
      End If
    Next i
  End If
  
  Changed = True
  
  ' Refresh the display
  RefreshSummaryFieldControls
    
End Sub

Private Sub cmdAdd_Click()
  ' Add the selected column to the Summary Fields listbox in the last position.
  Dim iListboxIndex As Integer
  Dim iListIndex As Integer
  Dim objColumn As Column
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
  
    iListIndex = lstColumns(iListboxIndex).ListIndex
  
    If iListIndex >= 0 Then
      
      ' Get the table and column name of the summary field
      Set objColumn = New Column
      objColumn.ColumnID = lstColumns(iListboxIndex).ItemData(iListIndex)
      
      If objColumn.ReadColumn Then
      
        ' Add the column to the summary fields listbox, and select it.
        With lstSummaryFields(iListboxIndex)
          .AddItem objColumn.Properties("columnName")
          .ItemData(.NewIndex) = objColumn.ColumnID
          .ListIndex = .NewIndex
        End With
        
        ' Remove the column from the Columns listbox, and select the next item.
        With lstColumns(iListboxIndex)
          .RemoveItem iListIndex
          If .ListCount > 0 Then
            .ListIndex = IIf(.ListCount > iListIndex, iListIndex, .ListCount - 1)
          End If
        End With

        Changed = True

      End If
      Set objColumn = Nothing
      
      RefreshSummaryFieldControls
      
    End If
  End If
  
End Sub

Private Sub cmdAddWorkflowLink_Click()
  WorkflowLink True

End Sub

Private Sub cmdCancel_Click()

  Dim lngResponse As Long

  If Changed Then
    lngResponse = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
  Else
    lngResponse = vbNo
  End If
  
  
  mblnUnloadForm = True
  If lngResponse = vbYes Then
    cmdOK_Click
  ElseIf lngResponse = vbNo Then
    gfCancelled = True
    UnLoad Me
  Else
    mblnUnloadForm = False
  End If
  
End Sub

Private Sub cmdColumnBreak_Click()

  ' Insert a column break in the Summary Fields listbox in the selected position, and select it.
  Dim iListboxIndex As Integer
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
    With lstSummaryFields(iListboxIndex)
      If .ListIndex >= 0 Then
        .AddItem msCOLUMNBREAKSTRING, .ListIndex
      Else
        .AddItem msCOLUMNBREAKSTRING
      End If
      .ItemData(.NewIndex) = miCOLUMNBREAKID
      .ListIndex = .NewIndex
    End With
    
    mbManualColumnInserted = True
    Changed = True
    
    RefreshSummaryFieldControls
    
  End If

End Sub

Private Sub cmdTableTriggerDelete_Click()

  Dim objTrigger As clsTableTrigger
  Dim lngTriggerID As Long
  Dim lngCount As Long
  
  lngTriggerID = ssGrdTableTriggers.Columns("TriggerID").value
  
  If DeleteRow("Table Trigger", ssGrdTableTriggers) Then
    lngCount = 1
    Do While lngCount <= mobjTable.TableTriggers.Count
      Set objTrigger = mobjTable.TableTriggers(lngCount)
      
      If objTrigger.TriggerID = lngTriggerID Then
        objTrigger.Deleted = True
      End If
      
      lngCount = lngCount + 1
    Loop

    Changed = True
    
  End If

  RefreshTableTriggerButtons

End Sub

Private Sub cmdTableTriggerDeleteAll_Click()

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  strMBText = "Remove all table triggers for this table, are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Remove All"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbYes Then

    ssGrdTableTriggers.RemoveAll
    mobjTable.TableTriggers = Nothing
    mobjTable.TableTriggers = New Collection

    RefreshTableTriggerButtons
    Changed = True
  End If
  
End Sub

Private Sub cmdTableTriggerEdit_Click()
  TableTrigger False
End Sub

Private Sub cmdTableTriggerNew_Click()
  TableTrigger True
End Sub

Private Sub cmdDown_Click()
  ' Move the selected Summary Field UP one position.
  Dim sCurrentItemText As String
  Dim iCurrentItemData As Integer
  Dim iListboxIndex As Integer
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
    With lstSummaryFields(iListboxIndex)
      If (.ListIndex >= 0) And _
        (.ListIndex < (.ListCount - 1)) Then
      
        ' Swap the current Summary Field item with the one above it, keeping it selected.
        sCurrentItemText = .List(.ListIndex)
        iCurrentItemData = .ItemData(.ListIndex)
      
        .List(.ListIndex) = .List(.ListIndex + 1)
        .ItemData(.ListIndex) = .ItemData(.ListIndex + 1)
        
        .List(.ListIndex + 1) = sCurrentItemText
        .ItemData(.ListIndex + 1) = iCurrentItemData
    
        .ListIndex = .ListIndex + 1
      End If
    End With
  End If
  
  Changed = True
  
  ' Refresh the display
  RefreshSummaryFieldControls
  
End Sub

Private Sub cmdEmail_Click()
  ' Display the Email selection form.
  Dim objEmail As clsEmailAddr
  
  ' Create a new Email object.
  Set objEmail = New clsEmailAddr
  
  ' Initialize the Email object.
  With objEmail
    .EmailID = mlngEmailID
    .TableID = mobjTable.TableID
  
    ' Instruct the Email object to handle the selection.
    If .SelectEmail Then
      mlngEmailID = .EmailID
      txtEmail.Text = .EmailName
    
      Changed = True
    Else
      ' Check in case the original Email has been deleted.
      With recEmailAddrEdit
        .Index = "idxID"
        .Seek "=", mlngEmailID

        If .NoMatch Then
          mlngEmailID = 0
          txtEmail.Text = vbNullString
        Else
          If !Deleted Then
            mlngEmailID = 0
            txtEmail.Text = vbNullString
          End If
        End If
      End With
    End If
  End With
  
  ' Disassociate object variables.
  Set objEmail = Nothing

End Sub

'Private Sub cmdEmailNotification_Click(Index As Integer)
'
'  ' Display the Email selection form.
'  Dim objEmail As clsEmailAddr
'
'  ' Create a new Email object.
'  Set objEmail = New clsEmailAddr
'
'  ' Initialize the Email object.
'  With objEmail
'    .EmailID = mlngEmailNotification(Index)
'    .TableID = mobjTable.TableID
'
'    ' Instruct the Email object to handle the selection.
'    If .SelectEmail Then
'      mlngEmailNotification(Index) = .EmailID
'      txtEmailNotification(Index).Text = .EmailName
'
'      Changed = True
'    Else
'      ' Check in case the original Email has been deleted.
'      With recEmailAddrEdit
'        .Index = "idxID"
'        .Seek "=", mlngEmailNotification(Index)
'
'        If .NoMatch Then
'          mlngEmailNotification(Index) = 0
'          txtEmailNotification(Index).Text = vbNullString
'        Else
'          If !Deleted Then
'            mlngEmailNotification(Index) = 0
'            txtEmailNotification(Index).Text = vbNullString
'          End If
'        End If
'      End With
'    End If
'  End With
'
'  ' Disassociate object variables.
'  Set objEmail = Nothing
'
'End Sub

Private Sub cmdInsert_Click()
  ' Insert the selected column in the Summary Fields listbox in the selected position.
  Dim iListIndex As Integer
  Dim iListboxIndex As Integer
  Dim objColumn As Column
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
  
    iListIndex = lstColumns(iListboxIndex).ListIndex
    
    If iListIndex >= 0 Then
      
      ' Get the table and column name of the summary field.
      Set objColumn = New Column
      objColumn.ColumnID = lstColumns(iListboxIndex).ItemData(iListIndex)
      
      If objColumn.ReadColumn Then
      
        ' Insert the Column in the Summary Fields listbox, and select it.
        With lstSummaryFields(iListboxIndex)
          If .ListIndex >= 0 Then
            .AddItem objColumn.Properties("columnName"), .ListIndex
          Else
            .AddItem objColumn.Properties("columnName")
          End If
          .ItemData(.NewIndex) = objColumn.ColumnID
          .ListIndex = .NewIndex
        End With
        
        With lstColumns(iListboxIndex)
          .RemoveItem .ListIndex
          If .ListCount > 0 Then
            .ListIndex = IIf(.ListCount > iListIndex, iListIndex, .ListCount - 1)
          End If
        End With
          
        Changed = True
      
      End If
      Set objColumn = Nothing
    
      RefreshSummaryFieldControls
    End If
  End If

End Sub

Private Sub cmdInsertBreak_Click()
  ' Insert a Break in the Summary Fields listbox in the selected position, and select it.
  Dim iListboxIndex As Integer
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
    With lstSummaryFields(iListboxIndex)
      If .ListIndex >= 0 Then
        .AddItem msBREAKSTRING, .ListIndex
      Else
        .AddItem msBREAKSTRING
      End If
      .ItemData(.NewIndex) = miBREAKID
      .ListIndex = .NewIndex
    End With
    
    Changed = True
    
    RefreshSummaryFieldControls
  End If
  
End Sub


Private Sub cmdOK_Click()
  ' Save the changes and unload the form.
  On Error GoTo ErrorTrap
  
  Dim fBreak As Boolean
  Dim iLoop As Integer
  Dim iListboxIndex As Integer
  Dim iSequence As Integer
  Dim sName As String
  Dim objSummaryField As cSummaryField
  Dim objSummaryFields As Collection
  Dim bBreakColumn As Boolean
  Dim objOutlookLink As clsOutlookLink
  Dim frmUse As frmUsage
  Dim frmPermissions As frmDefaultPermissions2
  
  mblnTableViewExists = False

  ' Format the table name.
  sName = Trim(FormatName(txtTableName.Text))
  If Len(sName) < 1 Then
    MsgBox "You must enter a table name.", _
      vbOKOnly + vbExclamation, Application.Name
    If txtTableName.Enabled Then
      txtTableName.SetFocus
    End If
    Exit Sub
  End If
  
  ' Ensure that the table name is unique.
  'If mobjTable.IsNew Then
    With recTabEdit
      .Index = "idxName"
      .Seek "=", sName, False
      
      If Not .NoMatch Then
        If !TableID <> mobjTable.TableID Then
          MsgBox "A table named '" & sName & "' already exists.", _
            vbOKOnly + vbExclamation, Application.Name
            mblnTableViewExists = True
          If txtTableName.Enabled Then
            txtTableName.SetFocus
          End If
          Exit Sub
        End If
      End If
    End With
  
    With recViewEdit
      .Index = "idxViewName"
      .Seek "=", sName, False
      
      If Not .NoMatch Then
        MsgBox "A view named '" & sName & "' already exists.", _
          vbOKOnly + vbExclamation, Application.Name
          mblnTableViewExists = True
        If txtTableName.Enabled Then
          txtTableName.SetFocus
        End If
        Exit Sub
      End If
    End With
  
    ' Ensure that the table name is not a keyword.
    If IsKeyword(sName) Then
      MsgBox "'" & sName & "' cannot be used as a table name" & _
        vbCr & "as it is a reserved word.", _
        vbOKOnly + vbExclamation, Application.Name
        mblnTableViewExists = True
      If txtTableName.Enabled Then
        txtTableName.SetFocus
      End If
      Exit Sub
    End If
  
    ' Ensure that the table name is not a system database name.
    If UCase(Left(sName, 6)) = "ASRSYS" Or UCase(Left(sName, 6)) = "TBSTAT" Or UCase(Left(sName, 6)) = "TBUSER" Then
      MsgBox "'" & sName & "' cannot be used as a table name" & _
        vbCr & "as the prefix '" & UCase(Left(sName, 6)) & "' is reserved for system tables.", _
        vbOKOnly + vbExclamation, Application.Name
        mblnTableViewExists = True
      If txtTableName.Enabled Then
        txtTableName.SetFocus
      End If
      Exit Sub
    End If
    
    If UCase(Left(sName, 5)) = "TBSYS" Then
      MsgBox "'" & sName & "' cannot be used as a table name" & _
        vbCr & "as the prefix '" & UCase(Left(sName, 5)) & "' is reserved for system tables.", _
        vbOKOnly + vbExclamation, Application.Name
        mblnTableViewExists = True
      If txtTableName.Enabled Then
        txtTableName.SetFocus
      End If
      Exit Sub
    End If
    
  'JPD 20060213 Fault 10781
  ' If the table type has changed, check it does not contravene any rules
  ' dependent on the tsable type..
  If (mobjTable.TableID > 0) And _
    ((optTableType(0).value And mobjTable.TableType <> iTabParent) Or _
    (optTableType(1).value And mobjTable.TableType <> iTabChild) Or _
    (optTableType(2).value And mobjTable.TableType <> iTabLookup)) Then
    ' Existing table AND table type has changed.
    
    ' If the table is now a Lookup table, make sure it has no relationships set up.
    If optTableType(2).value Then
      With recRelEdit
        .Index = "idxParentID"
        .Seek ">=", mobjTable.TableID

        If Not .NoMatch Then
          MsgBox "This table cannot be made a 'Lookup' type table" & _
            " as it already has relationships configured.", _
            vbOKOnly + vbExclamation, Application.Name
          Exit Sub
        End If
      End With
      
      With recRelEdit
        .Index = "idxChildID"
        .Seek "=", mobjTable.TableID

        If Not .NoMatch Then
          MsgBox "This table cannot be made a 'Lookup' type table" & _
            " as it already has relationships configured.", _
            vbOKOnly + vbExclamation, Application.Name
          Exit Sub
        End If
      End With
    End If
  
    ' If the table is now a Parent table, make sure it is NOT a child in any configured relationships.
    If optTableType(0).value Then
      With recRelEdit
        .Index = "idxChildID"
        .Seek "=", mobjTable.TableID

        If Not .NoMatch Then
          MsgBox "This table cannot be made a 'Parent' type table" & _
            " as there is already a relationship configured with it as the child table.", _
            vbOKOnly + vbExclamation, Application.Name
          Exit Sub
        End If
      End With
    End If
    
    ' If the table 'was' a Parent table, make sure that no associated views defined.
    If mobjTable.TableType = iTabParent Then
      Set frmUse = New frmUsage
      frmUse.ResetList
      With recViewEdit
        .Index = "idxViewTableID"
        .Seek "=", mobjTable.TableID
        
        If Not .NoMatch Then

          Do While Not .EOF
            If (!ViewTableID <> mobjTable.TableID) Then
              Exit Do
            End If
              
            If (Not !Deleted) Then
              frmUse.AddToList ("View : " & !ViewName)
            End If
          .MoveNext
          Loop
        End If
      End With
      
      If (frmUse.lstUsage.ListCount > 0) Then
        Screen.MousePointer = vbDefault
        frmUse.ShowMessage sName & " Table", "The type of this table cannot be changed" & _
                " as it is associated with the following view definitions :", UsageCheckObject.Table
        UnLoad frmUse
        Set frmUse = Nothing
        Exit Sub
      End If
      UnLoad frmUse
      Set frmUse = Nothing
    End If
    
  End If
  
  'NPG20080207 Fault 12874
  Set frmPermissions = New frmDefaultPermissions2
  
  'JDM - Fault 551 - Apply default permissions to existing security groups
  'NPG20080206 Fault 12874 - references new frmdefaultpermissions2 form now.
  If mobjTable.PermissionsPrompted = False And mobjTable.IsNew = True Then
    If optTableType(2).value = True Then
      frmPermissions.SetType "new", giTABLELOOKUP, Me.Icon
    Else
      frmPermissions.SetType "new", giTABLEPARENT, Me.Icon
    End If
    frmPermissions.Show vbModal
    If frmPermissions.OkCancel = vbOK Then
      mobjTable.GrantRead = frmPermissions.GrantRead
      mobjTable.GrantEdit = frmPermissions.GrantEdit
      mobjTable.GrantNew = frmPermissions.GrantNew
      mobjTable.GrantDelete = frmPermissions.GrantDelete
      mobjTable.PermissionsPrompted = True
    Else
      gfCancelled = True
      Exit Sub
    End If

  End If
  
  ' Write the table definition values to the table object.
  If Not mobjTable.IsNew Then
    If mobjTable.TableName <> sName Then
      'Name has changed so refresh stored procedures
      Application.ChangedTableName = True
      Call MarkViewsAndExpressionsChanged
    End If
  End If

  mobjTable.TableName = sName
  mobjTable.TableType = IIf(optTableType(0).value = True, 1, _
    IIf(optTableType(1).value = True, 2, 3))
  mobjTable.PrimaryOrderID = mlngOrderID
  mobjTable.RecordDescriptionID = mlngRecDescExprID
  mobjTable.PrimaryEmailID = mlngEmailID


  mobjTable.EmailLinks = mvarEmailLinks


  ' Notifications
  mobjTable.AuditInsert = (chkAuditInsertion.value = vbChecked)
  mobjTable.AuditDelete = (chkAuditDeletion.value = vbChecked)
  'mobjTable.EmailInsertID = mlngEmailNotification(0)
  'mobjTable.EmailDeleteID = mlngEmailNotification(1)
  
  ' Write the summary field definition to the table object.
  mobjTable.ManualSummaryColumnBreaks = mbManualColumnBreaks
  
  daoDb.Execute "DELETE FROM tmpSummary WHERE historytableID=" & mobjTable.TableID
  mobjTable.ClearSummaryFields
  Set objSummaryFields = New Collection
  
  For iListboxIndex = 1 To cboParentTable.ListCount
  
    iSequence = 1
    fBreak = False
    bBreakColumn = False
    For iLoop = 0 To (lstSummaryFields(iListboxIndex).ListCount - 1)
      If lstSummaryFields(iListboxIndex).ItemData(iLoop) = miCOLUMNBREAKID And mbManualColumnBreaks Then
        bBreakColumn = True
      Else
        If lstSummaryFields(iListboxIndex).ItemData(iLoop) = miBREAKID Then
          fBreak = True
        Else
          Set objSummaryField = New cSummaryField
          With objSummaryField
            .HistoryTableID = mobjTable.TableID
            .Sequence = iSequence
            .StartOfGroup = fBreak
            .StartOfColumn = bBreakColumn
            .SummaryColumnID = lstSummaryFields(iListboxIndex).ItemData(iLoop)
          End With
          objSummaryFields.Add objSummaryField
          iSequence = iSequence + 1
          fBreak = False
        End If
        bBreakColumn = False
      End If
    Next iLoop

  Next iListboxIndex

  Set mobjTable.SummaryFields = objSummaryFields
  Set objSummaryFields = Nothing


  'MH20040322
  For Each objOutlookLink In mobjTable.OutlookLinks
    objOutlookLink.WriteLink
  Next

  If mfRebuildWorkflowLinks Then
    Application.ChangedWorkflowLink = True
  End If
  
  mobjTable.InsertTriggerDisabled = (chkDisableInsert.value = vbChecked)
  mobjTable.UpdateTriggerDisabled = (chkDisableUpdate.value = vbChecked)
  mobjTable.DeleteTriggerDisabled = (chkDisableDelete.value = vbChecked)
  mobjTable.CopyWhenParentRecordIsCopied = (chkCopyWhenParentRecordIsCopied.value = vbChecked)

  gfCancelled = False
  
  UnLoad Me
  Exit Sub

ErrorTrap:
  MsgBox Err.Description, _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  
End Sub

Private Sub cmdOrder_Click()
  ' Display the Order selection form.
  Dim fCanSelectOrder As Boolean
  Dim objOrder As Order
  Dim sSQL As String
  Dim rsInfo As DAO.Recordset
  
  ' Create a new order object.
  Set objOrder = New Order
  
  'JPD 20040114 Fault 7912
  fCanSelectOrder = False
  
  If (mobjTable.TableID > 0) Then
    sSQL = "SELECT tmpColumns.columnID" & _
      " FROM tmpColumns" & _
      " WHERE tmpColumns.tableID = " & Trim(Str(mobjTable.TableID)) & _
      " AND tmpColumns.deleted = FALSE" & _
      " AND tmpColumns.ColumnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
      " AND tmpColumns.ColumnType <> " & Trim(Str(giCOLUMNTYPE_LINK)) & _
      " AND tmpColumns.DataType <> " & Trim(Str(dtLONGVARBINARY)) & _
      " AND tmpColumns.DataType <> " & Trim(Str(dtVARBINARY))
    Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      
    fCanSelectOrder = Not (rsInfo.BOF And rsInfo.EOF)
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If
  
  If Not fCanSelectOrder Then
    MsgBox "Unable to select an order." & vbCrLf & vbCrLf & _
      "You must first define some suitable columns for this table.", vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
  
  ' Initialize the order object.
  With objOrder
    .OrderID = mlngOrderID
    .TableID = mobjTable.TableID
    .OrderType = giORDERTYPE_DYNAMIC
  
    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      mlngOrderID = .OrderID
      txtOrder.Text = .OrderName
      
      Changed = True

      DoEvents
    
    Else
      ' Check in case the original order has been deleted.
      With recOrdEdit
        .Index = "idxID"
        .Seek "=", mlngOrderID

        If .NoMatch Then
          mlngOrderID = 0
        Else
          If !Deleted Then
            mlngOrderID = 0
          End If
        End If
      End With
    End If
  End With
  
  ' Disassociate object variables.
  Set objOrder = Nothing

End Sub

Private Sub cmdRecordDescription_Click()
  ' Display the Record Description selection form.
  Dim fOK As Boolean
  Dim objExpr As CExpression
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    fOK = .Initialise(mobjTable.TableID, mlngRecDescExprID, giEXPR_RECORDDESCRIPTION, giEXPRVALUE_CHARACTER)
  
    If fOK Then
      ' Instruct the expression object to display the
      ' expression selection form.
      If .SelectExpression Then
        mlngRecDescExprID = .ExpressionID
        
        'TM20020724 Fault 3884 - if the record desc exprid is less than 0, then set to 0.
        If mlngRecDescExprID < 0 Then
          mlngRecDescExprID = 0
        End If
        
        ' Read the selected expression info.
        GetRecordDescriptionDetails

        Changed = True
        
        'MH20051110 Fault 8801
        If TableHasDiaryLinks(mobjTable.TableID) Then
          Application.ChangedDiaryLink = True
        End If
      Else
        ' Check in case the original expression has been deleted.
        With recExprEdit
          .Index = "idxExprID"
          .Seek "=", mlngRecDescExprID, False
  
          If .NoMatch Then
            ' Read the selected expression info.
            mlngRecDescExprID = 0
            GetRecordDescriptionDetails
          End If
        End With
      End If
    End If
  End With
  
  ' Disassociate object variables.
  Set objExpr = Nothing

End Sub


Private Sub cmdRemove_Click()
  ' Remove the selected Summary Field from the Summary Field listbox and put it back
  ' in the Columns listbox.
  Dim iListIndex As Integer
  Dim iListboxIndex As Integer
  Dim objColumn As Column
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
  
    iListIndex = lstSummaryFields(iListboxIndex).ListIndex
    
    If iListIndex >= 0 Then
      
      If lstSummaryFields(iListboxIndex).ItemData(iListIndex) = miCOLUMNBREAKID Then
        mbManualColumnInserted = False
      Else
      
        If lstSummaryFields(iListboxIndex).ItemData(iListIndex) <> miBREAKID Then
        
          ' Get the table and column name of the summary field
          Set objColumn = New Column
          objColumn.ColumnID = lstSummaryFields(iListboxIndex).ItemData(iListIndex)
          
          If objColumn.ReadColumn Then
          
            ' Add the column to the Columns listbox, and select it.
            With lstColumns(iListboxIndex)
              .AddItem objColumn.Properties("columnName")
              .ItemData(.NewIndex) = objColumn.ColumnID
              .ListIndex = .NewIndex
            End With
            
          End If
          
          Set objColumn = Nothing
        End If
      End If
      
      ' Remove the item from the Summary Fields listbox, and select the next one.
      With lstSummaryFields(iListboxIndex)
        .RemoveItem iListIndex
        If .ListCount > 0 Then
          .ListIndex = IIf(.ListCount > iListIndex, iListIndex, .ListCount - 1)
        End If
      End With
    
      Changed = True
      
      RefreshSummaryFieldControls
    End If
  End If
  
End Sub

Private Sub cmdRemoveAllWorkflowLinks_Click()
  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  Dim objLink As clsWorkflowTriggeredLink
  
  strMBText = "Remove all Workflow links for this table, are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Remove All"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbYes Then

    ssGrdWorkflowLinks.RemoveAll
    
    For Each objLink In mobjTable.WorkflowTriggeredLinks
      If objLink.LinkType = WORKFLOWTRIGGERLINKTYPE_DATE Then
        mfRebuildWorkflowLinks = True
      End If
    Next objLink
    Set objLink = Nothing
    
    mobjTable.WorkflowTriggeredLinks = Nothing
    mobjTable.WorkflowTriggeredLinks = New Collection

    RefreshWorkflowLinksButtons
    Changed = True
  End If

End Sub

Private Sub cmdRemoveWorkflowLink_Click()
  Dim objLink As clsWorkflowTriggeredLink
  Dim lngLinkID As Long
  Dim lngCount As Long
  Dim iLinkType As WorkflowTriggerLinkType
  
  lngLinkID = ssGrdWorkflowLinks.Columns("LinkID").value
  iLinkType = WORKFLOWTRIGGERLINKTYPE_COLUMN
  
  If DeleteRow("Workflow link", ssGrdWorkflowLinks) Then
    lngCount = 1
    Do While lngCount <= mobjTable.WorkflowTriggeredLinks.Count
      Set objLink = mobjTable.WorkflowTriggeredLinks(lngCount)
      
      If objLink.LinkID = lngLinkID Then
        objLink.Deleted = True
        iLinkType = objLink.LinkType
      End If
      
      lngCount = lngCount + 1
    Loop

    Changed = True
    
    If iLinkType = WORKFLOWTRIGGERLINKTYPE_DATE Then
      mfRebuildWorkflowLinks = True
    End If
  End If

  RefreshWorkflowLinksButtons


End Sub

Private Sub cmdTableValidationDelete_Click()

  Dim objValidation As clsTableValidation
  Dim lngValidationID As Long
  Dim lngCount As Long
  Dim iValidationType As clsTableValidation
  
  lngValidationID = ssGrdTableValidations.Columns("ValidationID").value
  
  If DeleteRow("Table Validation", ssGrdTableValidations) Then
    lngCount = 1
    Do While lngCount <= mobjTable.TableValidations.Count
      Set objValidation = mobjTable.TableValidations(lngCount)
      
      If objValidation.ValidationID = lngValidationID Then
        objValidation.Deleted = True
      End If
      
      lngCount = lngCount + 1
    Loop

    Changed = True
    
  End If

  RefreshTableValidationsButtons

End Sub

Private Sub cmdTableValidationDeleteAll_Click()

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  Dim objLink As clsWorkflowTriggeredLink
  
  strMBText = "Remove all table validations for this table, are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Remove All"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbYes Then

    ssGrdTableValidations.RemoveAll
    mobjTable.TableValidations = Nothing
    mobjTable.TableValidations = New Collection

    RefreshTableValidationsButtons
    Changed = True
  End If


End Sub

Private Sub cmdTableValidationEdit_Click()
  TableValidation False
End Sub

Private Sub cmdTableValidationNew_Click()
  TableValidation True
End Sub

Private Sub cmdUp_Click()
  ' Move the selected Summary Field UP one position.
  Dim sCurrentItemText As String
  Dim iCurrentItemData As Integer
  Dim iListboxIndex As Integer
  
  iListboxIndex = CurrentListboxIndex
  
  If iListboxIndex > 0 Then
    With lstSummaryFields(iListboxIndex)
      If .ListIndex > 0 Then
      
        ' Swap the current Summary Field item with the one above it, keeping it selected.
        sCurrentItemText = .List(.ListIndex)
        iCurrentItemData = .ItemData(.ListIndex)
      
        .List(.ListIndex) = .List(.ListIndex - 1)
        .ItemData(.ListIndex) = .ItemData(.ListIndex - 1)
        
        .List(.ListIndex - 1) = sCurrentItemText
        .ItemData(.ListIndex - 1) = iCurrentItemData
        
        .ListIndex = .ListIndex - 1
      End If
    End With
  End If
  
  Changed = True
  
  ' Refresh the display
  RefreshSummaryFieldControls
  
End Sub


Private Sub cmdWorkflowLinkProperties_Click()
  WorkflowLink False

End Sub

Private Sub Form_Activate()

  ' Only enable the table name and type controls if the
  ' table is new.
'  If Not ASRDEVELOPMENT Then
'    If Not mobjTable.IsNew Then
'      txtTableName.Enabled = False
'      fraTableType.Enabled = False
'    End If
'
'    'MH20000814
'    txtTableName.BackColor = IIf(mobjTable.IsNew, vbWindowBackground, vbButtonFace)
'  End If

  'Set mvarEmailLinks = New Collection


  optTableType(0).Enabled = fraTableType.Enabled And Not mblnReadOnly
  optTableType(1).Enabled = fraTableType.Enabled And Not mblnReadOnly
  optTableType(2).Enabled = fraTableType.Enabled And Not mblnReadOnly
  
  ' Only enable the Order and Record Description controls
  ' if a valid table is selected.
  If mobjTable.TableID <= 0 Then
    'JPD 20040114 Fault 7912
    'cmdOrder.Enabled = False
    cmdRecordDescription.Enabled = False
    cmdEmail.Enabled = False    'MH20010117
  End If
  
  ' Get the the Order details.
  GetOrderDetails
  
  ' Get the Record Description details.
  GetRecordDescriptionDetails
  
  ' Get the Email Address details.
  GetEmailAddressDetails
  
  '' Get the notification alerts
  'GetNotifications
  
  ' Get table stats
  GetTableStats
  
  ' Set the manual column break setting
  chkManualColumnBreak.value = IIf(mobjTable.ManualSummaryColumnBreaks, vbChecked, vbUnchecked)
  
  ' Set focus to table name textbox.
  If txtTableName.Enabled Then
    txtTableName.SetFocus
  Else
    If cmdOrder.Enabled Then
      cmdOrder.SetFocus
    End If
  End If

  ' RH 20/02/01 - this shouldnt be here...it causes the OK
  '               button to disable when it should be enabled, eg
  '               when a primary order has been set, frm activate runs
  '               again
  If mblnNotLoading = False Then
    Changed = False
    mblnNotLoading = True
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case vbKeyU
    If (Shift And vbAltMask) > 0 Then
      cmdUp_Click
    End If
  Case vbKeyD
    If (Shift And vbAltMask) > 0 Then
      cmdDown_Click
    End If
End Select

End Sub


Private Sub Form_Load()
  Const GRIDROWHEIGHT = 239
  
  Set mvarEmailLinks = New Collection
  
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  ssGrdEmailLinks.RowHeight = GRIDROWHEIGHT
  ssGrdOutlookLinks.RowHeight = GRIDROWHEIGHT
  ssGrdWorkflowLinks.RowHeight = GRIDROWHEIGHT
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode) Or _
                  mobjTable.Locked
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    ssGrdOutlookLinks.Enabled = True
    cmdOutlookLinkProperties.Caption = "Vi&ew"
    cmdRecordDescription.Enabled = True
    cmdOrder.Enabled = True
    cmdEmail.Enabled = True
  
    ssGrdWorkflowLinks.Enabled = True
    cmdWorkflowLinkProperties.Caption = "Vi&ew"
  End If

  ssTabTableProperties.TabVisible(iTABLEPROPERTYTTAB_WORKFLOWLINKS) = Application.WorkflowModule

  ' Set the maximum table name length.
  txtTableName.MaxLength = MaxTableNameLength
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

  ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    If mblnTableViewExists Then
      Cancel = mblnUnloadForm
    Else
      Cancel = Not mblnUnloadForm
    End If
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Terminate()
  Set mvarEmailLinks = Nothing
End Sub

Private Sub lstColumns_DblClick(Index As Integer)
  cmdAdd_Click
  
End Sub


Private Sub lstSummaryFields_Click(Index As Integer)
  'JPD 20030902 Fault 6066
  If Not mfLoading Then
    RefreshSummaryFieldControls
  End If
  
End Sub

Private Sub lstSummaryFields_DblClick(Index As Integer)
  cmdRemove_Click

End Sub


Private Sub optTableType_Click(Index As Integer)
  
  chkCopyWhenParentRecordIsCopied.Enabled = (Index = 1)
  
  If Not chkCopyWhenParentRecordIsCopied.Enabled Then
    chkCopyWhenParentRecordIsCopied.value = vbUnchecked
  End If
  
  Changed = True
  
End Sub

Private Sub ssGrdEmailLinks_HeadClick(ByVal ColIndex As Integer)
  
  If ColIndex = 0 Then
    If mblnEmailSortByActivation Then
      mblnEmailSortByActivation = False
      mblnEmailSortDesc = False
    Else
      mblnEmailSortDesc = Not mblnEmailSortDesc
    End If
  
  Else
    If Not mblnEmailSortByActivation Then
      mblnEmailSortByActivation = True
      mblnEmailSortDesc = False
    Else
      mblnEmailSortDesc = Not mblnEmailSortDesc
    End If
  End If

  PopulateEmailLinks 0

End Sub

Private Sub ssGrdTableTriggers_DblClick()
  If cmdTableTriggerEdit.Enabled Then
    TableTrigger False
  End If
End Sub

Private Sub ssGrdTableValidations_DblClick()
  If cmdTableValidationEdit.Enabled Then
    TableValidation False
  End If
End Sub

Private Sub ssGrdWorkflowLinks_DblClick()
  If cmdWorkflowLinkProperties.Enabled Then
    cmdWorkflowLinkProperties_Click
  End If
  
End Sub

Private Sub ssTabTableProperties_Click(PreviousTab As Integer)
  Dim lstBoxTemp As ListBox
  
  ' --------------
  ' DEFINITION tab
  ' --------------
  fraTableType.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION)
  optTableType(0).Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION)
  optTableType(1).Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION)
  optTableType(2).Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION)
  cmdOrder.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION)
  'JPD 20050830 Fault 10284
  cmdRecordDescription.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION) And (mobjTable.TableID > 0)
  cmdEmail.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_DEFINITION) And (mobjTable.TableID > 0)

  ' -------------------
  ' SUMMARY COLUMNS tab
  ' -------------------
  cboParentTable.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  chkManualColumnBreak.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdAdd.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdInsert.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdRemove.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdInsertBreak.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdColumnBreak.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdUp.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  cmdDown.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  
  For Each lstBoxTemp In lstColumns
    lstBoxTemp.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  Next lstBoxTemp
  For Each lstBoxTemp In lstSummaryFields
    lstBoxTemp.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_SUMMARYCOLUMNS)
  Next lstBoxTemp
  Set lstBoxTemp = Nothing
  
  RefreshSummaryFieldControls
  
  
  ' --------------------------
  ' CALENDAR LINKS tab
  ' --------------------------
  fraEmail.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_EMAILLINKS)
  RefreshEmailLinksButtons
  
  ' --------------------------
  ' CALENDAR LINKS tab
  ' --------------------------
  fraCalendarLinks.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_CALENDARLINKS)
  RefreshOutlookLinksButtons
  
  ' --------------------------
  ' WORKFLOW LINKS tab
  ' --------------------------
  fraWorkflowLinks.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_WORKFLOWLINKS)
  RefreshWorkflowLinksButtons
  
  ' --------------------------
  ' TABLE VALIDATIONS tab
  ' --------------------------
  fraTableValidations.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_VALIDATION)
  RefreshTableValidationsButtons
  
  ' --------------------------
  ' TABLE TRIGGERS tab
  ' --------------------------
  fraTableTriggers.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_TRIGGERS)
  fraSystemTriggers.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_TRIGGERS)
  RefreshTableTriggerButtons
   
  ' --------------------------
  ' STATISTICS tab
  ' --------------------------
  fraAudit.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_AUDIT)
  fraTableStats.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_AUDIT)
  'fraOLEStats.Enabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_AUDIT)
  
End Sub

Private Sub txtTableName_Change()
  Dim sValidatedName As String
  Dim iSelStart As Integer
  Dim iSelLen As Integer
  
  'JPD 20090102 Fault 13484
  sValidatedName = ValidateName(txtTableName.Text)
  
  If sValidatedName <> txtTableName.Text Then
    iSelStart = txtTableName.SelStart
    iSelLen = txtTableName.SelLength
    
    txtTableName.Text = sValidatedName
    
    txtTableName.SelStart = iSelStart
    txtTableName.SelLength = iSelLen
  End If
  
  Changed = True

End Sub

Private Sub txtTableName_GotFocus()
  With txtTableName
    If .Locked = False Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With
End Sub

Private Sub txtTableName_KeyPress(KeyAscii As Integer)

  ' Validate the character entered.
  KeyAscii = ValidNameChar(KeyAscii, txtTableName.SelStart)
  
End Sub

Private Sub lstColumns_Initialise()
  ' Populate the Columns listbox with the columns from the Parent Table(s) of the Child Table.
  Dim fIsSummaryField As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim rsColumns As DAO.Recordset
  Dim objSummaryField As cSummaryField
            
  ' For each table in the parent table combo ...
  For iLoop = 1 To cboParentTable.ListCount
  
    ' Create a new Columns listbox if required.
    If iLoop > lstColumns.UBound Then
      Load lstColumns(iLoop)
    End If
    lstColumns(iLoop).Tag = cboParentTable.ItemData(iLoop - 1)
 
    With lstColumns(iLoop)
      .Clear
      .Visible = False
      .Tag = cboParentTable.ItemData(iLoop - 1)
      
      ' Get the column details for the history table's parent tables.
      sSQL = "SELECT tmpColumns.columnID, tmpColumns.columnName " & _
        " FROM tmpColumns" & _
        " WHERE tmpColumns.tableID = " & Trim(Str(cboParentTable.ItemData(iLoop - 1))) & _
        " AND tmpColumns.deleted = FALSE" & _
        " AND tmpColumns.controlType <> " & Trim(Str(giCTRL_OLE)) & _
        " AND tmpColumns.controlType <> " & Trim(Str(giCTRL_PHOTO)) & _
        " AND tmpColumns.controlType <> " & Trim(Str(giCTRL_LINK)) & _
        " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM))
      
      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
      Do While Not rsColumns.EOF
        ' Check if the column is already in the summary fields collection.
        fIsSummaryField = False
        For Each objSummaryField In mobjTable.SummaryFields
          If objSummaryField.SummaryColumnID = rsColumns!ColumnID Then
            fIsSummaryField = True
            Exit For
          End If
        Next objSummaryField
        Set objSummaryField = Nothing
          
        ' Add the column if it is not already in the summary fields collection.
        If Not fIsSummaryField Then
          .AddItem rsColumns!ColumnName
          .ItemData(.NewIndex) = rsColumns!ColumnID
        End If
          
        rsColumns.MoveNext
      Loop
      rsColumns.Close
      Set rsColumns = Nothing
    
      ' Select the first item (if there are any).
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
  Next iLoop
  
End Sub
Private Sub lstSummaryFields_Initialise()
  ' Populate the Summary Fields listbox with the Summary Fields of the table.
  Dim iLoop As Integer
  Dim objColumn As Column
  Dim objSummaryField As cSummaryField
  
  ' For each table in the parent table combo ...
  For iLoop = 1 To cboParentTable.ListCount
  
    ' Create a new Summary Field listbox if required.
    If iLoop > lstSummaryFields.UBound Then
      Load lstSummaryFields(iLoop)
    End If
    lstSummaryFields(iLoop).Tag = cboParentTable.ItemData(iLoop - 1)
 
    With lstSummaryFields(iLoop)
      .Clear
      .Visible = False
      .Tag = cboParentTable.ItemData(iLoop - 1)
    
      For Each objSummaryField In mobjTable.SummaryFields
      
        ' Get the table and column name of the summary field
        Set objColumn = New Column
        objColumn.ColumnID = objSummaryField.SummaryColumnID
        
        If objColumn.ReadColumn Then
        
          If objColumn.TableID = cboParentTable.ItemData(iLoop - 1) Then
            
            If objSummaryField.StartOfColumn Then
              .AddItem msCOLUMNBREAKSTRING
              .ItemData(.NewIndex) = miCOLUMNBREAKID
              mbManualColumnInserted = True
            End If
            
            If objSummaryField.StartOfGroup Then
              .AddItem msBREAKSTRING
              .ItemData(.NewIndex) = miBREAKID
            End If
            
            .AddItem objColumn.Properties("columnName")
            .ItemData(.NewIndex) = objColumn.ColumnID
          End If
        End If
        
        Set objColumn = Nothing
        
      Next objSummaryField
      Set objSummaryField = Nothing
      
      ' Select the first item (if there are any).
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
  Next iLoop
  
End Sub


Private Sub cboParentTable_Initialise()
  ' Populate the Parent Table combo with the parent table(s) of the current table.
  
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  
  cboParentTable.Clear
  
  If mobjTable.TableType = iTabChild Then
    ' Get the history table's parent tables.
    sSQL = "SELECT tmpTables.tableName, tmpTables.tableID" & _
      " FROM tmpRelations, tmpTables" & _
      " WHERE tmpRelations.childID=" & Trim(Str(mobjTable.TableID)) & _
      " AND tmpTables.tableID = tmpRelations.parentid" & _
      " AND tmpTables.deleted = FALSE"
    
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
    With rsTables
      Do While Not .EOF
        cboParentTable.AddItem !TableName
        cboParentTable.ItemData(cboParentTable.NewIndex) = !TableID
        
        .MoveNext
      Loop
              
      .Close
    End With
    Set rsTables = Nothing
  
  End If

  lstSummaryFields_Initialise
  lstColumns_Initialise
  
  If cboParentTable.ListCount > 0 Then
    cboParentTable.ListIndex = 0
  Else
    cboParentTable.Enabled = False
  End If

  RefreshSummaryFieldControls
  RefreshTableValidationsButtons
  
End Sub

Private Function CurrentListboxIndex() As Integer
  ' Return the index of the listboxes associated with the current table
  ' in the cboParentTable combo.
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  iIndex = 0
  
  If cboParentTable.ListIndex >= 0 Then
    For iLoop = 1 To lstColumns.UBound
      If (cboParentTable.ItemData(cboParentTable.ListIndex) = val(lstColumns(iLoop).Tag)) Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
  CurrentListboxIndex = iIndex
  
End Function


Private Function MarkViewsAndExpressionsChanged()

  Dim rsChangedDefs As DAO.Recordset
  Dim sSQL As String

  'Mark parent views as changed...
  sSQL = "SELECT ViewID FROM tmpViews " & _
         "WHERE ViewTableID = " & CStr(mobjTable.TableID)
  Set rsChangedDefs = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  Do While Not rsChangedDefs.EOF
    With recViewEdit
      .Index = "idxViewID"
      .Seek "=", rsChangedDefs!ViewID
      If Not .NoMatch Then
        .Edit
        .Fields("Changed") = True
        .Update
      End If
    End With
    rsChangedDefs.MoveNext
  Loop
  rsChangedDefs.Close
  Set rsChangedDefs = Nothing
  

'This isn't quite enough and finding exactly which expressions need to be
'rebuilt could be tricky.  Probably best to rebuild all expressions if
'renaming a table (for the moment).

'  'Also mark expressions as changed...
'  sSQL = "SELECT ExprID FROM tmpExpressions " & _
'         "WHERE ParentComponentID = 0 AND TableID = " & CStr(mobjTable.TableID)
'  Set rsChangedDefs = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  Do While Not rsChangedDefs.EOF
'    With recExprEdit
'      .Index = "idxExprID"
'      If !TableID = mobjTable.TableID And !ParentComponentID = 0 Then
'        .Edit
'        .Fields("Changed") = True
'        .Update
'      End If
'
'    End With
'    rsChangedDefs.MoveNext
'  Loop
'  rsChangedDefs.Close
'  Set rsChangedDefs = Nothing


  'Also mark column email addresses as changed
  sSQL = "UPDATE tmpEmailAddresses" & _
    " SET changed = TRUE" & _
    " WHERE columnID > 0 and TableID = " & CStr(mobjTable.TableID)
  daoDb.Execute sSQL, dbFailOnError

End Function

Public Sub PrintDefinition()

  Dim objPrinter As SystemMgr.clsPrintDef
  Dim bOK As Boolean
  Dim iCount As Integer
  Dim strSize As String
  Dim strDecimals As String
  Dim strColumnType As String
  Dim strControlType As String
  Dim strDataType As String
  Dim iCountCols As Integer
  Dim strType As String
  Dim iColumns As Integer
  Dim iParents As Integer
  
  bOK = True

  ' Get the the Order details.
  GetOrderDetails

  ' Get the Record Description details.
  GetRecordDescriptionDetails

  ' Get the Email Address details.
  GetEmailAddressDetails

  ' Load the printer object
  Set objPrinter = New SystemMgr.clsPrintDef
  With objPrinter
    If .IsOK Then
      If .PrintStart(True) Then
  
        .TabsOnPage = 6
        .PrintHeader "Table Name : " & txtTableName.Text
    
        .PrintTitle "Definition"
        
        For iCount = 0 To optTableType.Count - 1
          If optTableType(iCount).value = True Then strType = Replace(optTableType(iCount).Caption, "&", "")
        Next iCount
        .PrintNormal "Table Type : " & strType
        
        .PrintNormal "Primary Order : " & IIf(Len(txtOrder.Text) = 0, "<None>", txtOrder.Text)
        .PrintNormal "Record Description : " & IIf(Len(txtRecordDescription.Text) = 0, "<None>", txtRecordDescription.Text)
        .PrintNormal "Default Email : " & IIf(Len(txtEmail.Text) = 0, "<None>", txtEmail.Text)
        
        .PrintTitle "Columns"
    
        iColumns = 0
        With recColEdit
          .Index = "idxName"
          .Seek ">=", mobjTable.TableID
      
          If Not .NoMatch Then
            Do While Not .EOF
              ' Ignore any columns for tables other than the one used by the specified
              ' node.
              If .Fields("tableID") <> mobjTable.TableID Then
                Exit Do
              End If
      
              ' Ignore deleted and system columns.
              If (Not .Fields("deleted")) And _
                (Not !columntype = giCOLUMNTYPE_SYSTEM) Then
      
                  ' Size
                  If ColumnHasSize(.Fields("DataType").value) Then
                    If .Fields("Size").value = VARCHAR_MAX_Size Then
                      strSize = vbNullString
                    Else
                      strSize = Trim(Str(.Fields("Size").value))
                    End If
                    If ColumnHasScale(.Fields("DataType").value) Then
                      strDecimals = Trim(Str(.Fields("Decimals").value))
                    Else
                      strDecimals = ""
                    End If
                  Else
                    strSize = ""
                    strDecimals = ""
                  End If
      
                  ' Column Type
                  Select Case .Fields("ColumnType").value
                    Case giCOLUMNTYPE_SYSTEM
                     strColumnType = "System"
                    Case giCOLUMNTYPE_DATA
                      strColumnType = "Data"
                    Case giCOLUMNTYPE_LOOKUP
                      strColumnType = "Lookup"
                    Case giCOLUMNTYPE_CALCULATED
                      strColumnType = "Calculated"
                    Case giCOLUMNTYPE_LINK
                      strColumnType = "Link"
                  End Select
      
                  Select Case .Fields("ControlType").value
                    Case giCTRL_CHECKBOX
                      strControlType = "Check Box"
                    Case giCTRL_COMBOBOX
                      strControlType = "Dropdown List"
                    Case giCTRL_OPTIONGROUP
                      strControlType = "Option Group"
                    Case giCTRL_SPINNER
                      strControlType = "Spinner"
                    Case giCTRL_TEXTBOX
                      strControlType = "Text Box"
                    Case giCTRL_WORKINGPATTERN
                      strControlType = "Working Pattern"
                    Case Else
                      strControlType = ""
                  End Select
      
                  ' Data type
                  strDataType = GetDataDesc(.Fields("DataType").value)
                  
                  If Len(strSize) > 0 Then
                    If Len(strDecimals) > 0 Then
                      strDataType = strDataType & " (" & strSize & "," & strDecimals & ")"
                    Else
                      strDataType = strDataType & " (" & strSize & ")"
                    End If
                  End If
    
                  ' Output the column info
                  If iColumns = 0 Then
                    objPrinter.PrintBold "Column Name" & vbTab & vbTab & vbTab & "Data Type" & vbTab & "Control" & vbTab & "Type" '& vbTab & "Size"
                  End If
                  
                  objPrinter.PrintNonBold .Fields("ColumnName").value & vbTab & vbTab & vbTab _
                        & strDataType & vbTab & strControlType & vbTab & strColumnType
                  iColumns = iColumns + 1
              End If
      
              .MoveNext
            Loop
          End If
        End With
  
        If iColumns = 0 Then .PrintNonBold "<None>"
        
        ' Summary Columns
        .PrintTitle "Summary Columns"
      
        iParents = 0
        For iCount = 1 To cboParentTable.ListCount
          .PrintNormal cboParentTable.List(iCount - 1)
          iParents = iParents + 1
          
          iColumns = 0
          For iCountCols = 0 To lstSummaryFields(iCount).ListCount - 1
            .PrintNonBold Trim(lstSummaryFields(1).List(iCountCols))
            iColumns = iColumns + 1
          Next iCountCols
        
          If iColumns = 0 Then .PrintNonBold "<None>"
        Next iCount
    
        If iParents = 0 Then .PrintNonBold "<None>"
        
        ' Audit Log
        .PrintTitle "Audit Log"
        .PrintNormal "Insertions : " & IIf(chkAuditInsertion.value = vbChecked, "Yes", "No")
        .PrintNormal "Deletions : " & IIf(chkAuditDeletion.value = vbChecked, "Yes", "No")
  
        '' Email Notifications
        '.PrintTitle "Email Notification"
        '.PrintNormal "Insertion : " & IIf(Len(txtEmailNotification(0).Text) = 0, "<None>", txtEmailNotification(0).Text)
        '.PrintNormal "Deletion : " & IIf(Len(txtEmailNotification(1).Text) = 0, "<None>", txtEmailNotification(1).Text)
  
  
        ' Email Links Tab
        .PrintTitle "Email Links"
        If ssGrdEmailLinks.Rows > 0 Then
          .TabsOnPage = 2

          .PrintBold "Name" & vbTab & "Email Activation"

          ssGrdEmailLinks.MoveFirst
          For iCount = 1 To ssGrdEmailLinks.Rows
            .PrintNonBold ssGrdEmailLinks.Columns(0).value & vbTab & ssGrdEmailLinks.Columns(1).value
            ssGrdEmailLinks.MoveNext
          Next iCount
        Else
          .PrintNonBold "<None>"
        End If


        ' Calendar Links Tab
        .PrintTitle "Outlook Calendar Links"
        If ssGrdOutlookLinks.Rows > 0 Then
          .TabsOnPage = 2
  
          .PrintBold "Name" & vbTab & "Subject"
  
          ssGrdOutlookLinks.MoveFirst
          For iCount = 1 To ssGrdOutlookLinks.Rows
            .PrintNonBold ssGrdOutlookLinks.Columns(1).value & vbTab & ssGrdOutlookLinks.Columns(2).value
            ssGrdOutlookLinks.MoveNext
          Next iCount
        Else
          .PrintNonBold "<None>"
        End If
  
  
        ' Workflow Links Tab
        If Application.WorkflowModule Then
          .PrintTitle "Workflow Links"
          If ssGrdWorkflowLinks.Rows > 0 Then
            .TabsOnPage = 2
    
            .PrintBold "Name" & vbTab & "Enabled"
    
            ssGrdWorkflowLinks.MoveFirst
            For iCount = 1 To ssGrdWorkflowLinks.Rows
              .PrintNonBold ssGrdWorkflowLinks.Columns("Name").value & vbTab & IIf(ssGrdWorkflowLinks.Columns("Enabled").value, "Yes", "No")
              ssGrdWorkflowLinks.MoveNext
            Next iCount
          Else
            .PrintNonBold "<None>"
          End If
        End If
  
  
        ' Table validation stuff
        .PrintTitle "Validations"
        If ssGrdTableValidations.Rows > 0 Then
          .TabsOnPage = 1
  
          .PrintBold "Description"
  
          ssGrdTableValidations.MoveFirst
          For iCount = 1 To ssGrdTableValidations.Rows
            .PrintNonBold ssGrdTableValidations.Columns(1).value
            ssGrdTableValidations.MoveNext
          Next iCount
        Else
          .PrintNonBold "<None>"
        End If
  
  
        ' Print the footer
        .PrintEnd       ' this adds the correct footer
      End If
    End If
  End With
  Set objPrinter = Nothing

TidyUpAndExit:
  If Not bOK Then
    MsgBox "Unable to print the table definition." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Sub

Public Sub CopyDefinitionToClipboard()

  Dim bOK As Boolean
  Dim iCount As Integer
  Dim iCountCols As Integer
  Dim strColumnType As String
  Dim strControlType As String
  Dim strClipboardText As String
  Dim strType As String
  Dim strSize As String
  Dim strDecimals As String
  Dim strDataType As String
  Dim iColumns As Integer
  Dim iParents As Integer
  
  bOK = True
 
  ' Get the the Order details.
  GetOrderDetails
  
  ' Get the Record Description details.
  GetRecordDescriptionDetails
  
  ' Get the Email Address details.
  GetEmailAddressDetails
  
  strClipboardText = "Definition" & vbCrLf
  strClipboardText = strClipboardText & "----------" & vbCrLf & vbCrLf
  
  ' Table Info
  strClipboardText = strClipboardText & "Table Name : " & txtTableName.Text & vbCrLf
  
  ' Table type
  For iCount = 0 To optTableType.Count - 1
    If optTableType(iCount).value = True Then strType = Replace(optTableType(iCount).Caption, "&", "")
  Next iCount
  strClipboardText = strClipboardText & "Table Type : " & strType & vbCrLf
  
  ' Primary order
  strClipboardText = strClipboardText & "Primary Order : " & IIf(Len(txtOrder.Text) = 0, "<None>", txtOrder.Text) & vbCrLf
  
  ' Record Description
  strClipboardText = strClipboardText & "Record Description : " & IIf(Len(txtRecordDescription.Text) = 0, "<None>", txtRecordDescription.Text) & vbCrLf
 
  ' Default Email
  strClipboardText = strClipboardText & "Default Email : " & IIf(Len(txtEmail.Text) = 0, "<None>", txtEmail.Text) & vbCrLf & vbCrLf
 
  ' Column info
  strClipboardText = strClipboardText & "Columns" & vbCrLf
  strClipboardText = strClipboardText & "-------" & vbCrLf & vbCrLf
 
  iColumns = 0
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mobjTable.TableID
  
    If Not .NoMatch Then
      Do While Not .EOF
        ' Ignore any columns for tables other than the one used by the specified
        ' node.
        If .Fields("tableID") <> mobjTable.TableID Then
          Exit Do
        End If
          
        ' Ignore deleted and system columns.
        If (Not .Fields("deleted")) And _
          (Not !columntype = giCOLUMNTYPE_SYSTEM) Then
    
            ' Size
            If ColumnHasSize(.Fields("DataType").value) Then
              strSize = Trim(Str(.Fields("Size").value))
              If ColumnHasScale(.Fields("DataType").value) Then
                strDecimals = Trim(Str(.Fields("Decimals").value))
              Else
                strDecimals = ""
              End If
            Else
              strSize = ""
              strDecimals = ""
            End If
          
            ' Column Type
            Select Case .Fields("ColumnType").value
              Case giCOLUMNTYPE_SYSTEM
               strColumnType = "System"
              Case giCOLUMNTYPE_DATA
                strColumnType = "Data"
              Case giCOLUMNTYPE_LOOKUP
                strColumnType = "Lookup"
              Case giCOLUMNTYPE_CALCULATED
                strColumnType = "Calculated"
              Case giCOLUMNTYPE_LINK
                strColumnType = "Link"
            End Select
          
            Select Case .Fields("ControlType").value
              Case giCTRL_CHECKBOX
                strControlType = "Check Box"
              Case giCTRL_COMBOBOX
                strControlType = "Dropdown List"
              Case giCTRL_OPTIONGROUP
                strControlType = "Option Group"
              Case giCTRL_SPINNER
                strControlType = "Spinner"
              Case giCTRL_TEXTBOX
                strControlType = "Text Box"
              Case giCTRL_WORKINGPATTERN
                strControlType = "Working Pattern"
              Case Else
                strControlType = ""
            End Select
          
            ' Data type
            strDataType = GetDataDesc(.Fields("DataType").value)
          
            If iColumns = 0 Then
              strClipboardText = strClipboardText & "Column Name" & vbTab & "Data Type" & vbTab & "Control" & vbTab & "Type" & vbTab & "Size" & vbCrLf
            End If
            
            strClipboardText = strClipboardText & .Fields("ColumnName").value & vbTab _
                  & strDataType & vbTab & strControlType & vbTab & strColumnType & vbTab _
                  & strSize & vbTab & strDecimals & vbCrLf
          
            iColumns = iColumns + 1
        End If
          
        .MoveNext
      Loop
    End If
  End With

  If iColumns = 0 Then strClipboardText = strClipboardText & "<None>" & vbCrLf

  ' Summary Columns
  strClipboardText = strClipboardText & vbCrLf & "Summary Columns" & vbCrLf
  strClipboardText = strClipboardText & "---------------" & vbCrLf & vbCrLf

  iParents = 0
  For iCount = 1 To cboParentTable.ListCount
    strClipboardText = strClipboardText & cboParentTable.List(iCount - 1) & vbCrLf
    iParents = iParents + 1
    
    iColumns = 0
    For iCountCols = 0 To lstSummaryFields(iCount).ListCount - 1
      strClipboardText = strClipboardText & "     " & Trim(lstSummaryFields(1).List(iCountCols)) & vbCrLf
      iColumns = iColumns + 1
    Next iCountCols
  
    If iColumns = 0 Then strClipboardText = strClipboardText & "     <None>" & vbCrLf
  Next iCount

  If iParents = 0 Then strClipboardText = strClipboardText & "<None>" & vbCrLf


  ' Email Links Tab
  strClipboardText = strClipboardText & vbCrLf & "Email Links" & vbCrLf
  strClipboardText = strClipboardText & "---------------" & vbCrLf & vbCrLf
  If ssGrdEmailLinks.Rows > 0 Then

    strClipboardText = strClipboardText & "Name          Email Activation" & vbCrLf

    ssGrdEmailLinks.MoveFirst
    For iCount = 1 To ssGrdEmailLinks.Rows
      strClipboardText = strClipboardText & ssGrdEmailLinks.Columns(0).value & "          " & ssGrdEmailLinks.Columns(1).value & vbCrLf
      ssGrdEmailLinks.MoveNext
    Next iCount
  Else
    strClipboardText = strClipboardText & "<None>" & vbCrLf
  End If


  ' Calendar Links Tab
  strClipboardText = strClipboardText & vbCrLf & "Outlook Calendar Links" & vbCrLf
  strClipboardText = strClipboardText & "---------------" & vbCrLf & vbCrLf
  If ssGrdOutlookLinks.Rows > 0 Then
    
    strClipboardText = strClipboardText & "Title          Subject" & vbCrLf
    
    ssGrdOutlookLinks.MoveFirst
    For iCount = 1 To ssGrdOutlookLinks.Rows
      strClipboardText = strClipboardText & ssGrdOutlookLinks.Columns(1).value & "          " & ssGrdOutlookLinks.Columns(2).value & vbCrLf
      ssGrdOutlookLinks.MoveNext
    Next iCount
  Else
    strClipboardText = strClipboardText & "<None>" & vbCrLf
  End If

  ' Workflow Links Tab
  If Application.WorkflowModule Then
    strClipboardText = strClipboardText & vbCrLf & "Workflow Links" & vbCrLf
    strClipboardText = strClipboardText & "---------------" & vbCrLf & vbCrLf
    If ssGrdWorkflowLinks.Rows > 0 Then
  
      strClipboardText = strClipboardText & "Name          Enabled" & vbCrLf
  
      ssGrdWorkflowLinks.MoveFirst
      For iCount = 1 To ssGrdWorkflowLinks.Rows
        strClipboardText = strClipboardText & ssGrdWorkflowLinks.Columns("Name").value & "          " & IIf(ssGrdWorkflowLinks.Columns("Enabled").value, "Yes", "No") & vbCrLf
        ssGrdWorkflowLinks.MoveNext
      Next iCount
    Else
      strClipboardText = strClipboardText & "<None>" & vbCrLf
    End If
  End If

  ' Table validation stuff
  strClipboardText = strClipboardText & vbCrLf & "Validations" & vbCrLf
  strClipboardText = strClipboardText & "---------------" & vbCrLf & vbCrLf
  If ssGrdTableValidations.Rows > 0 Then
    strClipboardText = strClipboardText & "Description" & vbCrLf

    ssGrdTableValidations.MoveFirst
    For iCount = 1 To ssGrdTableValidations.Rows
      strClipboardText = strClipboardText & ssGrdTableValidations.Columns(1).value
      ssGrdTableValidations.MoveNext
    Next iCount
  Else
    strClipboardText = strClipboardText & "<None>" & vbCrLf
  End If


  ' Put the info in the clipboard
  Clipboard.Clear
  Clipboard.SetText strClipboardText

End Sub


'Private Sub GetNotifications()
'
'  Dim strEmailName As String
'  Dim iCount As Integer
'
'  strEmailName = vbNullString
'
'  For iCount = 0 To 1
'    If mlngEmailNotification(iCount) > 0 Then
'
'      With recEmailAddrEdit
'        .Index = "idxID"
'        .Seek "=", mlngEmailNotification(iCount)
'
'        ' Read the expression's name from the recordset.
'        If Not .NoMatch Then
'          strEmailName = !Name
'        End If
'
'      End With
'
'      txtEmailNotification(iCount) = strEmailName
'
'    End If
'  Next iCount
'
'End Sub



Private Sub cmdRemoveAllOutlookLinks_Click()

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  strMBText = "Remove all Outlook Calendar links for this table, are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Remove All"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbYes Then

    ssGrdOutlookLinks.RemoveAll
    
    mobjTable.OutlookLinks = Nothing
    mobjTable.OutlookLinks = New Collection

    RefreshOutlookLinksButtons
    Changed = True

  End If

End Sub


Private Sub RefreshEmailLinksButtons()
  
  Dim blnEnabled As Boolean
  
  With ssGrdEmailLinks
    If .Rows > 0 And .SelBookmarks.Count = 0 Then
      .SelBookmarks.Add .AddItemBookmark(0)
    End If

    blnEnabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_EMAILLINKS And Not mblnReadOnly)

    cmdAddEmailLink.Enabled = blnEnabled
    cmdEmailLinkProperties.Enabled = (.SelBookmarks.Count > 0 And ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_EMAILLINKS)
    cmdRemoveEmailLink.Enabled = (.SelBookmarks.Count > 0 And blnEnabled)
    cmdRemoveAllEmailLinks.Enabled = (.Rows > 0 And blnEnabled)
  End With

End Sub

Private Sub RefreshOutlookLinksButtons()
  
  Dim blnEnabled As Boolean
  
  With ssGrdOutlookLinks
    If .Rows > 0 And .SelBookmarks.Count = 0 Then
      .SelBookmarks.Add .AddItemBookmark(0)
    End If

    blnEnabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_CALENDARLINKS And Not mblnReadOnly)

    cmdAddOutlookLink.Enabled = blnEnabled
    cmdOutlookLinkProperties.Enabled = (.SelBookmarks.Count > 0 And ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_CALENDARLINKS)
    cmdRemoveOutlookLink.Enabled = (.SelBookmarks.Count > 0 And blnEnabled)
    cmdRemoveAllOutlookLinks.Enabled = (.Rows > 0 And blnEnabled)
  End With

End Sub

Private Sub RefreshWorkflowLinksButtons()
  
  Dim blnEnabled As Boolean

  With ssGrdWorkflowLinks
    If .Rows > 0 And .SelBookmarks.Count = 0 Then
      .SelBookmarks.Add .AddItemBookmark(0)
    End If

    blnEnabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_WORKFLOWLINKS And Not mblnReadOnly)

    cmdAddWorkflowLink.Enabled = blnEnabled
    cmdWorkflowLinkProperties.Enabled = (.SelBookmarks.Count > 0 And ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_WORKFLOWLINKS)
    cmdRemoveWorkflowLink.Enabled = (.SelBookmarks.Count > 0 And blnEnabled)
    cmdRemoveAllWorkflowLinks.Enabled = (.Rows > 0 And blnEnabled)
  End With

End Sub

Private Sub RefreshTableValidationsButtons()
  
  Dim blnEnabled As Boolean
  Dim bOnlyOneParent As Boolean
  
  bOnlyOneParent = cboParentTable.ListCount < 2
  ControlsDisableAll fraTableValidations, bOnlyOneParent
  
  If bOnlyOneParent Then
    With ssGrdTableValidations
      If .Rows > 0 And .SelBookmarks.Count = 0 Then
        .SelBookmarks.Add .AddItemBookmark(0)
      End If
  
      blnEnabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_VALIDATION And Not mblnReadOnly)
  
      cmdTableValidationNew.Enabled = blnEnabled
      cmdTableValidationEdit.Enabled = (.SelBookmarks.Count > 0 And ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_VALIDATION)
      cmdTableValidationDelete.Enabled = (.SelBookmarks.Count > 0 And blnEnabled)
      cmdTableValidationDeleteAll.Enabled = (.Rows > 0 And blnEnabled)
    End With
  End If

End Sub

Private Sub cmdOutlookLinkProperties_Click()
  OutlookLink False
End Sub

Private Sub cmdAddOutlookLink_Click()
  OutlookLink True
End Sub


Private Sub OutlookLink(blnNew As Boolean)

  Dim frmOutlook As frmOutlookCalendarLink
  Dim objNewLink As clsOutlookLink
  Dim strKey As String
  Dim lngID As Long
  Dim iLoop As Integer
  Dim lngRow As Long

  Set frmOutlook = New frmOutlookCalendarLink
  Set objNewLink = New clsOutlookLink


  'Set up defaults for new link
  If blnNew Then
    With objNewLink

      .LinkID = GetNewLinkID
      .TableID = mobjTable.TableID

      .Title = vbNullString
      .FilterID = 0
      .BusyStatus = 0
  
      .StartDate = 0
      .EndDate = 0
  
      .TimeRange = 0
      .FixedStartTime = vbNullString
      .FixedEndTime = vbNullString
      .ColumnStartTime = 0
      .ColumnEndTime = 0
  
      .Reminder = 0
      .ReminderOffset = 0
      .ReminderPeriod = 0
  
      .Subject = 0
      .content = vbNullString
    
    End With
  
    'frmOutlook.TableID = mobjTable.TableID
    frmOutlook.Locked = mobjTable.Locked
    frmOutlook.OutlookLink = objNewLink
    frmOutlook.OutlookLink.Destinations = objNewLink.Destinations
    frmOutlook.OutlookLink.LinkColumns = objNewLink.LinkColumns

  Else
    lngID = val(ssGrdOutlookLinks.Columns(0).CellText(ssGrdOutlookLinks.Bookmark))
    frmOutlook.OutlookLink = mobjTable.OutlookLinks("ID" & CStr(lngID))
'    frmOutlook.OutlookLink.Recipients = mvarOutlookLinks("ID" & CStr(lngID)).Recipients
  
  End If

  If frmOutlook.PopulateControls Then
    frmOutlook.Show vbModal
  End If

  If Not frmOutlook.Cancelled Then

    Set objNewLink = frmOutlook.OutlookLink
    strKey = "ID" & objNewLink.LinkID
    If Not blnNew Then
      mobjTable.OutlookLinks.Remove strKey
    End If
    mobjTable.OutlookLinks.Add objNewLink, strKey

    With ssGrdOutlookLinks
      .Redraw = False

      If Not blnNew Then
        lngRow = .AddItemRowIndex(.Bookmark)
        .RemoveItem lngRow
      Else
        lngRow = .Rows
      End If

      .AddItem objNewLink.LinkID & vbTab & objNewLink.Title & vbTab & GetExpressionName(objNewLink.Subject), lngRow
      .Redraw = True
      
      .Redraw = False
      .Bookmark = .AddItemBookmark(lngRow)
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      .Redraw = True
    End With


    RefreshOutlookLinksButtons
    Changed = True

  End If

  Set objNewLink = Nothing

  UnLoad frmOutlook
  Set frmOutlook = Nothing

End Sub


Private Sub WorkflowLink(blnNew As Boolean)

  Dim frmLink As frmWorkflowLink
  Dim objNewLink As clsWorkflowTriggeredLink
  Dim strKey As String
  Dim lngID As Long
  Dim lngRow As Long
  Dim iOriginalLinkType As WorkflowTriggerLinkType
  
  Set frmLink = New frmWorkflowLink
  Set objNewLink = New clsWorkflowTriggeredLink

  'Set up defaults for new link
  If blnNew Then
    With objNewLink
      .LinkID = GetNewWorkflowLinkID
      .WorkflowID = 0
      .TableID = mobjTable.TableID
      .FilterID = 0
      .EffectiveDate = Date
      .LinkType = WORKFLOWTRIGGERLINKTYPE_COLUMN
      .RecordInsert = False
      .RecordUpdate = False
      .RecordDelete = False
      .DateColumnID = 0
      .DateOffset = 0
      .DateOffsetPeriod = WORKFLOWTRIGGERLINKOFFESTPERIOD_DAY
    End With

    frmLink.WorkflowLink = objNewLink
  Else
    lngID = val(ssGrdWorkflowLinks.Columns(0).CellText(ssGrdWorkflowLinks.Bookmark))
    frmLink.WorkflowLink = mobjTable.WorkflowTriggeredLinks("ID" & CStr(lngID))
  End If

  iOriginalLinkType = frmLink.WorkflowLink.LinkType

  frmLink.Locked = mobjTable.Locked
  If frmLink.PopulateControls Then
    frmLink.Show vbModal
  End If

  If Not frmLink.Cancelled Then
    Set objNewLink = frmLink.WorkflowLink
    strKey = "ID" & objNewLink.LinkID
    If Not blnNew Then
      mobjTable.WorkflowTriggeredLinks.Remove strKey
    End If
    mobjTable.WorkflowTriggeredLinks.Add objNewLink, strKey

    With ssGrdWorkflowLinks
      .Redraw = False

      If Not blnNew Then
        lngRow = .AddItemRowIndex(.Bookmark)
        .RemoveItem lngRow
      Else
        lngRow = .Rows
      End If

      .AddItem objNewLink.LinkID & vbTab & GetWorkflowName(objNewLink.WorkflowID) & vbTab & GetWorkflowEnabled(objNewLink.WorkflowID), lngRow
      .Redraw = True

      .Redraw = False
      .Bookmark = .AddItemBookmark(lngRow)
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      .Redraw = True
    End With

    RefreshWorkflowLinksButtons
    Changed = True
    
    If iOriginalLinkType = WORKFLOWTRIGGERLINKTYPE_DATE _
      Or frmLink.WorkflowLink.LinkType = WORKFLOWTRIGGERLINKTYPE_DATE Then

      mfRebuildWorkflowLinks = True
    End If
  End If

  Set objNewLink = Nothing

  UnLoad frmLink
  Set frmLink = Nothing

End Sub



Private Function GetNewLinkID()

  Dim objNewLink As clsOutlookLink
  Dim lngNewID As Long

  On Local Error GoTo LocalErr

  Set objNewLink = New clsOutlookLink

  lngNewID = UniqueColumnValue("tmpOutlookLinks", "LinkID")

  Do While lngNewID < 99999
    Set objNewLink = mobjTable.OutlookLinks("ID" & lngNewID)
    lngNewID = lngNewID + 1
  Loop

Exit Function

LocalErr:
  GetNewLinkID = lngNewID

End Function

Private Function GetNewWorkflowLinkID() As Long
  Dim objNewLink As clsWorkflowTriggeredLink
  Dim lngNewID As Long
  
  On Local Error GoTo LocalErr
  
  lngNewID = UniqueColumnValue("tmpWorkflowTriggeredLinks", "LinkID")

  Do While lngNewID < 99999
    Set objNewLink = mobjTable.WorkflowTriggeredLinks("ID" & lngNewID)
    lngNewID = lngNewID + 1
  Loop

Exit Function

LocalErr:
  GetNewWorkflowLinkID = lngNewID

End Function

Private Function GetNewTableValidationID() As Long
  Dim objNewValidation As clsTableValidation
  Dim lngNewID As Long
  
  On Local Error GoTo LocalErr
  
  lngNewID = UniqueColumnValue("tmpTableValidations", "ValidationID")

  Do While lngNewID < 99999
    Set objNewValidation = mobjTable.TableValidations("ID" & lngNewID)
    lngNewID = lngNewID + 1
  Loop

  GetNewTableValidationID = lngNewID

Exit Function

LocalErr:
  GetNewTableValidationID = lngNewID

End Function


Private Sub cmdRemoveOutlookLink_Click()

  Dim objOutlookLink As clsOutlookLink
  Dim lngLinkID As Long
  Dim lngCount As Long

  lngLinkID = ssGrdOutlookLinks.Columns(0).value

  If DeleteRow("outlook link", ssGrdOutlookLinks) Then
    'mvaroutlookLinks.Remove "C" & lngLinkID

    lngCount = 1
    Do While lngCount <= mobjTable.OutlookLinks.Count
      Set objOutlookLink = mobjTable.OutlookLinks(lngCount)
      If objOutlookLink.LinkID = lngLinkID Then
        'objOutlookLink.Remove lngCount
        objOutlookLink.Deleted = True
      End If
      lngCount = lngCount + 1
    Loop

    Changed = True
    Application.ChangedOutlookLink = True
  End If

  RefreshOutlookLinksButtons

End Sub

Private Function DeleteRow(strType As String, ssgrd As SSDBGrid) As Boolean

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  Dim lRow As Long

  strMBText = "Delete this " & strType & ", are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Delete"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbNo Then
    DeleteRow = False
    Exit Function
  End If
  
  DeleteRow = True

  With ssgrd
    .Redraw = False
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
    .Redraw = True
    
  End With
 
  If Not mfLoading Then Changed = True
 

 
End Function

Private Sub GetTableStats()

  Dim sSQL As String
  Dim rsDetails As New ADODB.Recordset
  Dim rsDetails2 As New ADODB.Recordset
  Dim objListItem As ListItem
  Dim strSize As String
  Dim strOriginalName As String
  
  strOriginalName = vbNullString
  
  lstOLEColumns.ListItems.Clear

  If Not mobjTable.IsNew = True Then

    ' Get basic table stats
    'JPD 20040924 Fault 9224
    sSQL = "SELECT tableName" & _
      " FROM ASRSysTables" & _
      " WHERE tableID = " & CStr(mobjTable.TableID)
    rsDetails.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    With rsDetails
      If Not (.BOF And .EOF) Then
        strOriginalName = .Fields("tableName").value
      End If
      .Close
    End With
    
    If LenB(strOriginalName) = 0 Then
      lblStatsRows.Caption = "Rows : 0"
      lblDataSize.Caption = "Table Size : 0"
    Else
      sSQL = "exec sp_spaceused 'tbuser_" & strOriginalName & "'"
      rsDetails.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
      With rsDetails
        If Not (.BOF And .EOF) Then
          lblStatsRows.Caption = "Rows : " & .Fields("Rows").value
          
          strSize = Replace(.Fields("Data").value, "KB", "") * 1000
          lblDataSize.Caption = "Table Size : " & NiceSize(strSize)
        End If
        .Close
      End With
    
    End If
    
    ' Get OLE stats
    sSQL = "SELECT ColumnName FROM ASRSysColumns WHERE TableID = " & mobjTable.TableID & " AND DataType = -4 AND MaxOLESizeEnabled = 1"
    rsDetails.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    Do While Not (rsDetails.EOF)
  
      sSQL = "SELECT SUM(DATALENGTH(" & rsDetails.Fields(0).value & "))" _
          & " FROM " & strOriginalName _
          & " WHERE DATALENGTH(" & rsDetails.Fields(0).value & ") > 300"
      rsDetails2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      
      Do While Not (rsDetails2.EOF)
        strSize = IIf(IsNull(rsDetails2.Fields(0).value), 0, rsDetails2.Fields(0).value)
    
        Set objListItem = lstOLEColumns.ListItems.Add(, , rsDetails.Fields(0).value)
        objListItem.SubItems(1) = NiceSize(strSize)
      
        rsDetails2.MoveNext
      Loop
      rsDetails2.Close
  
      rsDetails.MoveNext
    Loop
    rsDetails.Close

  Else
  
    lblStatsRows.Caption = "Rows : 0"
    lblDataSize.Caption = "Table Size : 0"
    
  End If

  Set rsDetails = Nothing

End Sub

Private Function NiceSize(pstrSize As String) As String

  Select Case Len(pstrSize)
    Case Is < 5
      NiceSize = pstrSize & " bytes"
    
    Case Is < 7
      NiceSize = Mid(pstrSize, 1, Len(pstrSize) - 3) & " KB"
    
    Case 7
      NiceSize = Mid(pstrSize, 1, 1) & "." & Mid(pstrSize, 2, 2) & " MB"
    
    Case Is < 10
      NiceSize = Mid(pstrSize, 1, Len(pstrSize) - 6) & " MB"
  
  End Select

End Function


Private Sub cmdAddEmailLink_Click()

  Dim frmEmail As frmEmailLink
  Dim objNewLink As clsEmailLink
  Dim strKey As String
  'Dim iLoop As Integer

  Set frmEmail = New frmEmailLink
  Set objNewLink = New clsEmailLink
  
  
  'Set up defaults for new link
  With objNewLink
  
    .LinkID = .GetNewLinkID(mvarEmailLinks)
    '.ColumnID = mobjTable.ColumnID

    .Title = vbNullString
    .FilterID = 0
    .DateOffset = 0
    .DatePeriod = iTimePeriodDays
    .DateAmendment = True
    .EffectiveDate = Date
  
    '.Subject = vbNullString
    '.Importance = 1
    '.Sensitivity = 0
    '.IncRecordDesc = True
    '.IncColumnDetails = True
    '.IncUsername = True
  
    '.Text = vbNullString
    .Attachment = vbNullString
  
  End With
  
  frmEmail.TableID = mobjTable.TableID
  frmEmail.EmailLink = objNewLink
  frmEmail.EmailLink.RecipientsTo = objNewLink.RecipientsTo
  frmEmail.EmailLink.RecipientsCc = objNewLink.RecipientsCc
  frmEmail.EmailLink.RecipientsBcc = objNewLink.RecipientsBcc
  'frmEmail.AllowOffset = (miDataType = dtTIMESTAMP)
  frmEmail.PopulateControls
  frmEmail.Show vbModal

  If Not frmEmail.Cancelled Then

    Set objNewLink = frmEmail.EmailLink

    'With objNewLink.Recipients
      strKey = "ID" & objNewLink.LinkID
      mvarEmailLinks.Add objNewLink, strKey
      'For iLoop = 1 To objNewLink.Recipients.Count
      '  mvarEmailLinks.Item(strKey).Recipients.Add objNewLink.Recipients(iLoop)
      'Next
    
    
'    With ssGrdEmailLinks
'
'      .AddItem _
'        objNewLink.Title & vbTab & _
'        GetActivationDesc(objNewLink) & vbTab & _
'        "" & vbTab & _
'        objNewLink.LinkID
'
'      .Redraw = False
'      .Bookmark = .AddItemBookmark(.Rows - 1)
'      .SelBookmarks.RemoveAll
'      .SelBookmarks.Add .Bookmark
'      .Redraw = True
'
'    End With

    PopulateEmailLinks objNewLink.LinkID


    RefreshEmailLinksButtons
    Changed = True

  End If

  Set objNewLink = Nothing

  UnLoad frmEmail
  Set frmEmail = Nothing



'  ' Define a new email Link.
'  Dim frmEmail As frmEmailLink
'  Dim sDfltComment As String
'  Dim sOffset As String
'
'  ' Create a default comment for the email link.
'  sDfltComment = ""
'  With recTabEdit
'    .Index = "idxTableID"
'    .Seek "=", IIf(IsNull(mobjColumn.Properties("tableID")), 0, mobjColumn.Properties("tableID"))
'
'    If Not .NoMatch Then
'      sDfltComment = .Fields("tableName")
'    End If
'  End With
'  sDfltComment = sDfltComment & IIf(Len(sDfltComment) > 0, ".", "")
'  sDfltComment = sDfltComment & Trim(txtColumnName.Text)
'
'  Set frmEmail = New frmEmailLink
'
'  With frmEmail
'    ' Initialise the email link.
'    .emailComment = sDfltComment
'    .emailOffset = "0"
'    .emailPeriod = iTimePeriodDays
'    .emailReminder = False
'
'    ' Display the email link form.
'    .Show vbModal
'
'    ' Read the new email link.
'    If Not .Cancelled Then
'      sOffset = GetActivationDesc(frmEmail.emailOffset, frmEmail.emailPeriod)
'
'      ' Add the email link to the grid.
'      ssGrdEmailLinks.AddItem frmEmail.emailComment & _
'        vbTab & sOffset & _
'        vbTab & frmEmail.emailReminder & _
'        vbTab & frmEmail.emailOffset & _
'        vbTab & frmEmail.emailPeriod
'
'      ' Select the new row.
'      ssGrdEmailLinks.Bookmark = (ssGrdEmailLinks.Rows - 1)
'      ssGrdEmailLinks.SelBookmarks.Add ssGrdEmailLinks.Bookmark
'
'      ' Refesh the email link page controls.
'      RefreshEmailLinksButtons
'    End If
'  End With
'
'  ' Disassociate object variables.
'  Set frmEmail = Nothing
'
End Sub




Private Sub cmdEmailLinkProperties_Click()

  Dim frmEmail As frmEmailLink
  Dim objNewLink As clsEmailLink
  Dim lngOldLinkID As Long

  Dim strRow As String
  Dim lngRow As Long

  Set frmEmail = New frmEmailLink
  Set objNewLink = New clsEmailLink

  On Error GoTo LocalErr

  'Get existing object
  Set objNewLink = mvarEmailLinks.Item("ID" & ssGrdEmailLinks.Columns(3).value)
  lngOldLinkID = objNewLink.LinkID
  
  frmEmail.Locked = mobjTable.Locked
  Load frmEmail   'Required!
  frmEmail.TableID = mobjTable.TableID
  frmEmail.EmailLink = objNewLink
  'frmEmail.AllowOffset = (miDataType = dtTIMESTAMP)
  frmEmail.PopulateControls
  frmEmail.Show vbModal

  If Not frmEmail.Cancelled Then

    mvarEmailLinks.Remove "ID" & lngOldLinkID
    Set objNewLink = frmEmail.EmailLink
    mvarEmailLinks.Add objNewLink, "ID" & objNewLink.LinkID

'    strRow = objNewLink.Title & vbTab & _
'             GetActivationDesc(objNewLink) & vbTab & _
'             "" & vbTab & _
'             CStr(objNewLink.LinkID)
'
'    With ssGrdEmailLinks
'
'      lngRow = .AddItemRowIndex(.Bookmark)
'      .RemoveItem lngRow
'      .AddItem strRow, lngRow
'
'      .Redraw = False
'
'      .Bookmark = .AddItemBookmark(lngRow)
'      .SelBookmarks.RemoveAll
'      .SelBookmarks.Add .Bookmark
'
'      .Redraw = True
'
'    End With
    
    
    PopulateEmailLinks objNewLink.LinkID

    
    RefreshEmailLinksButtons
    Changed = True

  End If

  Set objNewLink = Nothing
  
  UnLoad frmEmail
  Set frmEmail = Nothing

Exit Sub

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASR DEVELOPMENT"
    Stop
  End If

End Sub



Private Sub cmdRemoveAllEmailLinks_Click()
  
  If RemoveAllRows("email links", ssGrdEmailLinks) Then
  
    Do While mvarEmailLinks.Count > 0
      mvarEmailLinks.Remove 1
    Loop
    
    ' Refesh the diary link page controls.
    RefreshEmailLinksButtons
    Changed = True
    Application.ChangedEmailLink = True   '15/08/2001 MH Fault 2679

  End If

End Sub



Private Sub cmdRemoveEmailLink_Click()

  Dim lngLinkID As Long
  Dim iLoop As Long

  lngLinkID = ssGrdEmailLinks.Columns(3).value

  If DeleteRow("email link", ssGrdEmailLinks) Then
    'On Error Resume Next
    
    iLoop = 1
    Do While iLoop <= mvarEmailLinks.Count
      If mvarEmailLinks(iLoop).LinkID = lngLinkID Then
        mvarEmailLinks.Remove iLoop
      Else
        iLoop = iLoop + 1
      End If
    Loop
    
    Changed = True
    Application.ChangedEmailLink = True     '15/08/2001 MH Fault 2679

  End If

  ' Refesh the email link page controls.
  RefreshEmailLinksButtons
  
End Sub







Private Sub ssGrdEmailLinks_DblClick()
  
  If cmdEmailLinkProperties.Enabled Then
    cmdEmailLinkProperties_Click
  ElseIf cmdAddEmailLink.Enabled Then
    cmdAddEmailLink_Click
  End If

End Sub


Private Function GetActivationDesc(objEmailLink As clsEmailLink) As String
  
'  If blnImmediate = True Then
'    GetActivationDesc = "Immediate"
'  Else
'    If intOffset = 0 Then
'      GetActivationDesc = "No offset"
'    Else
'      GetActivationDesc = _
'        CStr(Abs(intOffset)) & " " & _
'        TimePeriod(intTimePeriod) & _
'        IIf(Abs(intOffset) = 1, "", "s") & _
'        IIf(intOffset < 0, " before", " after")
'    End If
'  End If
  
  Dim strOutput As String
  
  Select Case objEmailLink.LinkType
  Case 0
    If objEmailLink.Columns.Count = 1 Then
      strOutput = GetColumnName(objEmailLink.Columns(1), True) & " changes"
    Else
      strOutput = "Any of " & CStr(objEmailLink.Columns.Count) & " specified columns change"
    End If
  Case 1
    strOutput = vbNullString
    
    If objEmailLink.RecordDelete Then
      strOutput = "deleted"
    End If
    
    If objEmailLink.RecordUpdate Then
      strOutput = "updated" & IIf(strOutput <> vbNullString, " or " & strOutput, "")
                  
    End If
    
    If objEmailLink.RecordInsert Then
      strOutput = "inserted" & IIf(strOutput <> vbNullString, " or " & strOutput, "")
    End If
    
    strOutput = "Record is " & strOutput
    
  Case 2
    If objEmailLink.DateOffset = 0 Then
      'strOutput = "At " & GetColumnName(objEmailLink.DateColumnID, True)
      strOutput = GetColumnName(objEmailLink.DateColumnID, True)
    Else
      'strOutput = CStr(Abs(objEmailLink.DateOffset)) & " " & _
        TimePeriod(objEmailLink.DatePeriod) & _
        IIf(Abs(objEmailLink.DateOffset) = 1, "", "s") & _
        IIf(objEmailLink.DateOffset < 0, " before ", " after ") & _
        GetColumnName(objEmailLink.DateColumnID, True)
      strOutput = GetColumnName(objEmailLink.DateColumnID, True) & " (" & _
        CStr(Abs(objEmailLink.DateOffset)) & " " & _
        TimePeriod(objEmailLink.DatePeriod) & _
        IIf(Abs(objEmailLink.DateOffset) = 1, "", "s") & _
        IIf(objEmailLink.DateOffset < 0, " before)", " after)")
    End If
  End Select
  
  GetActivationDesc = strOutput

End Function


Public Function RemoveAllRows(strType As String, ssgrd As SSDBGrid) As Boolean

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  strMBText = "Remove all " & strType & " for this column, are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Remove All"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbNo Then
    RemoveAllRows = False
    Exit Function
  End If

  RemoveAllRows = True
  ssgrd.RemoveAll
  If Not mfLoading Then Changed = True

End Function

Private Sub PopulateEmailLinks(intSelectedID As Integer)

  Dim objEmailLink As clsEmailLink
  Dim intLoop As Integer
  Dim intID As Integer
  
  
  With lstSort
    .Clear
    For Each objEmailLink In mvarEmailLinks
      If mblnEmailSortByActivation Then
        .AddItem GetActivationDesc(objEmailLink)
      Else
        .AddItem objEmailLink.Title
      End If
      .ItemData(.NewIndex) = objEmailLink.LinkID
    Next
  End With
    
    
  With ssGrdEmailLinks
    
    If intSelectedID = 0 Then
      intSelectedID = val(.Columns(3).CellText(.Bookmark))
    End If
    
    .RemoveAll
    For intLoop = 0 To lstSort.ListCount - 1
      
      If mblnEmailSortDesc Then
        intID = lstSort.ItemData(lstSort.ListCount - (intLoop + 1))
      Else
        intID = lstSort.ItemData(intLoop)
      End If
      
      For Each objEmailLink In mvarEmailLinks
        If intID = objEmailLink.LinkID Then
          .AddItem objEmailLink.Title & vbTab & _
                GetActivationDesc(objEmailLink) & vbTab & _
                "" & vbTab & objEmailLink.LinkID

          If intSelectedID = intID Then
            .Redraw = False
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .SelBookmarks.Add .Bookmark
            .Redraw = True
          End If

          Exit For
        End If
      Next

    Next
    
    Set objEmailLink = Nothing
  End With

End Sub


Private Sub TableValidation(blnNew As Boolean)

  Dim frmValidation As frmTableValidation
  Dim objValidation As clsTableValidation
  Dim strKey As String
  Dim lngID As Long
  Dim iLoop As Integer
  Dim lngRow As Long
  Dim sAddString As String
  
  Set frmValidation = New frmTableValidation
  Set objValidation = New clsTableValidation
  
  'Set up defaults for new link
  With objValidation
    If blnNew Then
      .ValidationID = GetNewTableValidationID
      .TableID = mobjTable.TableID
      .ValidationType = 3
      .Message = vbNullString
    Else
      lngID = val(ssGrdTableValidations.Columns(0).CellText(ssGrdTableValidations.Bookmark))
      Set objValidation = mobjTable.TableValidations("ID" & CStr(lngID))
    End If
  End With

  frmValidation.Locked = mobjTable.Locked
  frmValidation.ValidationObject = objValidation

  If frmValidation.PopulateControls Then
    frmValidation.Show vbModal
  End If

  If Not frmValidation.Cancelled Then

    Set objValidation = frmValidation.ValidationObject
    strKey = "ID" & objValidation.ValidationID
    If Not blnNew Then
      mobjTable.TableValidations.Remove strKey
    End If
    mobjTable.TableValidations.Add objValidation, strKey

    With ssGrdTableValidations
      .Redraw = False

      If Not blnNew Then
        lngRow = .AddItemRowIndex(.Bookmark)
        .RemoveItem lngRow
      Else
        lngRow = .Rows
      End If

      sAddString = GetValidationString(objValidation)
      .AddItem sAddString, lngRow
      .Redraw = True

      .Redraw = False
      .Bookmark = .AddItemBookmark(lngRow)
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      .Redraw = True
    End With

    RefreshTableValidationsButtons
    Changed = True

  End If

  Set objValidation = Nothing
  UnLoad frmValidation
  Set frmValidation = Nothing

End Sub

Private Function GetValidationString(ByRef objValidation As clsTableValidation) As String

  GetValidationString = objValidation.ValidationID & vbTab & objValidation.Message & vbTab

End Function

Private Sub RefreshTableTriggerButtons()
  
  Dim blnEnabled As Boolean
  Dim bOnlyOneParent As Boolean
  
  ControlsDisableAll fraTableTriggers, bOnlyOneParent
  
    With ssGrdTableTriggers
      If .Rows > 0 And .SelBookmarks.Count = 0 Then
        .SelBookmarks.Add .AddItemBookmark(0)
      End If
  
      blnEnabled = (ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_TRIGGERS And Not mblnReadOnly)
  
      cmdTableTriggerNew.Enabled = blnEnabled
      cmdTableTriggerEdit.Enabled = (.SelBookmarks.Count > 0 And ssTabTableProperties.Tab = iTABLEPROPERTYTTAB_TRIGGERS)
      cmdTableTriggerDelete.Enabled = (.SelBookmarks.Count > 0 And blnEnabled)
      cmdTableTriggerDeleteAll.Enabled = (.Rows > 0 And blnEnabled)
    End With

End Sub

Private Function GetTriggerString(ByRef objTrigger As clsTableTrigger) As String
  GetTriggerString = objTrigger.TriggerID & vbTab & objTrigger.Name & vbTab
End Function

Private Function GetNewTableTriggerID() As Long
  Dim objNewTrigger As clsTableTrigger
  Dim lngNewID As Long
  
  On Local Error GoTo LocalErr
  
  lngNewID = UniqueColumnValue("tmpTableTrigger", "TriggerID")

  Do While lngNewID < 99999
    Set objNewTrigger = mobjTable.TableTriggers("ID" & lngNewID)
    lngNewID = lngNewID + 1
  Loop

Exit Function

LocalErr:
  GetNewTableTriggerID = lngNewID

End Function

Private Sub TableTrigger(blnNew As Boolean)

  Dim frmTrigger As frmTrigger
  Dim objTrigger As clsTableTrigger
  Dim strKey As String
  Dim lngID As Long
'  Dim iLoop As Integer
  Dim lngRow As Long
  Dim sAddString As String

  Set frmTrigger = New frmTrigger
  Set objTrigger = New clsTableTrigger

  ' Set up defaults for new trigger
  With objTrigger
    If blnNew Then
      .TriggerID = GetNewTableTriggerID
      .TableID = mobjTable.TableID
      .Name = vbNullString
      .IsSystem = False
      .CodePosition = TriggerCodePosition.AfterU02Update
    Else
      lngID = val(ssGrdTableTriggers.Columns(0).CellText(ssGrdTableTriggers.Bookmark))
      Set objTrigger = mobjTable.TableTriggers("ID" & CStr(lngID))
    End If
  End With

  frmTrigger.Locked = mobjTable.Locked
  frmTrigger.TriggerObject = objTrigger

  If frmTrigger.PopulateControls Then
    frmTrigger.Show vbModal
  End If

  If Not frmTrigger.Cancelled Then

    Set objTrigger = frmTrigger.TriggerObject
    strKey = "ID" & objTrigger.TriggerID
    If Not blnNew Then
      mobjTable.TableTriggers.Remove strKey
    End If
    mobjTable.TableTriggers.Add objTrigger, strKey

    With ssGrdTableTriggers
      .Redraw = False

      If Not blnNew Then
        lngRow = .AddItemRowIndex(.Bookmark)
        .RemoveItem lngRow
      Else
        lngRow = .Rows
      End If

      sAddString = GetTriggerString(objTrigger)
      .AddItem sAddString, lngRow
      .Redraw = True

      .Redraw = False
      .Bookmark = .AddItemBookmark(lngRow)
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      .Redraw = True
    End With

    RefreshTableTriggerButtons
    Changed = True

  End If

  Set objTrigger = Nothing
  UnLoad frmTrigger
  Set frmTrigger = Nothing

End Sub

