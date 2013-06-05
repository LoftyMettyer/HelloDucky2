VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSSIntranetSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Module"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5040
   Icon            =   "frmSSIntranetSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6040
      TabIndex        =   45
      Top             =   5700
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7300
      TabIndex        =   46
      Top             =   5700
      Width           =   1200
   End
   Begin TabDlg.SSTab ssTabStrip 
      Height          =   5535
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frmSSIntranetSetup.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraViews"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Hypertext Links"
      TabPicture(1)   =   "frmSSIntranetSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraHypertextLinks"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dash&board"
      TabPicture(2)   =   "frmSSIntranetSetup.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraButtonLinks"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Dropdown List Links"
      TabPicture(3)   =   "frmSSIntranetSetup.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraDropdownListLinks"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "On-screen Docume&nt Display"
      TabPicture(4)   =   "frmSSIntranetSetup.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDocuments"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraDocuments 
         Caption         =   "On-screen Document Display :"
         Enabled         =   0   'False
         Height          =   4935
         Left            =   -74850
         TabIndex        =   54
         Top             =   400
         Width           =   8120
         Begin VB.CommandButton cmdAddDocument 
            Caption         =   "&Add ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   38
            Top             =   750
            Width           =   1245
         End
         Begin VB.CommandButton cmdEditDocument 
            Caption         =   "&Edit ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   39
            Top             =   1250
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveDocument 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   6720
            TabIndex        =   41
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveAllDocuments 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   6720
            TabIndex        =   42
            Top             =   2745
            Width           =   1245
         End
         Begin VB.CommandButton cmdCopyDocument 
            Caption         =   "Cop&y ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   40
            Top             =   1750
            Width           =   1245
         End
         Begin VB.ComboBox cboDocumentView 
            Height          =   315
            Left            =   2100
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   300
            Width           =   4515
         End
         Begin SSDataWidgets_B.SSDBGrid grdDocuments 
            Height          =   3945
            Index           =   0
            Left            =   195
            TabIndex        =   37
            Top             =   750
            Visible         =   0   'False
            Width           =   6400
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   15
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   79
            Columns.Count   =   15
            Columns(0).Width=   10821
            Columns(0).Caption=   "Document Title"
            Columns(0).Name =   "Text"
            Columns(0).CaptionAlignment=   2
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "URL"
            Columns(1).Name =   "URL"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "HRProScreenID"
            Columns(2).Name =   "HRProScreenID"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "PageTitle"
            Columns(3).Name =   "PageTitle"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "startMode"
            Columns(4).Name =   "startMode"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "UtilityType"
            Columns(5).Name =   "UtilityType"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "UtilityID"
            Columns(6).Name =   "UtilityID"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "HiddenGroups"
            Columns(7).Name =   "HiddenGroups"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   32000
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "NewWindow"
            Columns(8).Name =   "NewWindow"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "EMailAddress"
            Columns(9).Name =   "EMailAddress"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "EMailSubject"
            Columns(10).Name=   "EMailSubject"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "AppFilePath"
            Columns(11).Name=   "AppFilePath"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   3200
            Columns(12).Visible=   0   'False
            Columns(12).Caption=   "AppParameters"
            Columns(12).Name=   "AppParameters"
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   8
            Columns(12).FieldLen=   256
            Columns(13).Width=   3200
            Columns(13).Visible=   0   'False
            Columns(13).Caption=   "ReportOutputFilePath"
            Columns(13).Name=   "DocumentFilePath"
            Columns(13).DataField=   "Column 13"
            Columns(13).DataType=   8
            Columns(13).FieldLen=   256
            Columns(14).Width=   3200
            Columns(14).Visible=   0   'False
            Columns(14).Caption=   "ShowDocumentHyperlink"
            Columns(14).Name=   "DisplayDocumentHyperlink"
            Columns(14).DataField=   "Column 14"
            Columns(14).DataType=   11
            Columns(14).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   11289
            _ExtentY        =   6959
            _StockProps     =   79
            Enabled         =   0   'False
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdMoveDocumentUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   6720
            TabIndex        =   43
            Top             =   3255
            Width           =   1245
         End
         Begin VB.CommandButton cmdMoveDocumentDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   6720
            TabIndex        =   44
            Top             =   3750
            Width           =   1245
         End
         Begin VB.Label lblDocumentView 
            Caption         =   "Table (View) :"
            Height          =   195
            Left            =   195
            TabIndex        =   55
            Top             =   360
            Width           =   1560
         End
      End
      Begin VB.Frame fraHypertextLinks 
         Caption         =   "Hypertext Links :"
         Enabled         =   0   'False
         Height          =   4935
         Left            =   -74850
         TabIndex        =   47
         Top             =   405
         Width           =   8120
         Begin VB.ComboBox cboHypertextLinkView 
            Height          =   315
            Left            =   2100
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   300
            Width           =   4515
         End
         Begin VB.CommandButton cmdCopyHypertextLink 
            Caption         =   "Cop&y ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   12
            Top             =   1750
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveAllHypertextLinks 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   6720
            TabIndex        =   14
            Top             =   2775
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveHypertextLink 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   6720
            TabIndex        =   13
            Top             =   2265
            Width           =   1245
         End
         Begin VB.CommandButton cmdEditHypertextLink 
            Caption         =   "&Edit ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   11
            Top             =   1250
            Width           =   1245
         End
         Begin VB.CommandButton cmdAddHyperTextLink 
            Caption         =   "&Add ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   10
            Top             =   750
            Width           =   1245
         End
         Begin SSDataWidgets_B.SSDBGrid grdHypertextLinks 
            Height          =   3945
            Index           =   0
            Left            =   195
            TabIndex        =   9
            Top             =   750
            Visible         =   0   'False
            Width           =   6400
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   13
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   265
            Columns.Count   =   13
            Columns(0).Width=   10821
            Columns(0).Caption=   "Text"
            Columns(0).Name =   "Text"
            Columns(0).CaptionAlignment=   2
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "URL"
            Columns(1).Name =   "URL"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "UtilityType"
            Columns(2).Name =   "UtilityType"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "UtilityID"
            Columns(3).Name =   "UtilityID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "HiddenGroups"
            Columns(4).Name =   "HiddenGroups"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   32000
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "NewWindow"
            Columns(5).Name =   "NewWindow"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "EMailAddress"
            Columns(6).Name =   "EMailAddress"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "EMailSubject"
            Columns(7).Name =   "EMailSubject"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "AppFilePath"
            Columns(8).Name =   "AppFilePath"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "AppParameters"
            Columns(9).Name =   "AppParameters"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "Element_Type"
            Columns(10).Name=   "Element_Type"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "SeparatorOrientation"
            Columns(11).Name=   "SeparatorOrientation"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   3200
            Columns(12).Visible=   0   'False
            Columns(12).Caption=   "PictureID"
            Columns(12).Name=   "PictureID"
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   8
            Columns(12).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   11298
            _ExtentY        =   6959
            _StockProps     =   79
            Enabled         =   0   'False
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdMoveHypertextLinkDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   6720
            TabIndex        =   16
            Top             =   3765
            Width           =   1245
         End
         Begin VB.CommandButton cmdMoveHypertextLinkUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   6720
            TabIndex        =   15
            Top             =   3270
            Width           =   1245
         End
         Begin VB.Label lblHypertextLinkView 
            Caption         =   "Table (View) :"
            Height          =   195
            Left            =   195
            TabIndex        =   48
            Top             =   360
            Width           =   1740
         End
      End
      Begin VB.Frame fraViews 
         Caption         =   "Tables (Views) :"
         Height          =   4935
         Left            =   -74850
         TabIndex        =   53
         Top             =   400
         Width           =   8120
         Begin VB.CommandButton cmdRemoveAllTableViews 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   6720
            TabIndex        =   5
            Top             =   1800
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveTableView 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   6720
            TabIndex        =   4
            Top             =   1300
            Width           =   1245
         End
         Begin VB.CommandButton cmdEditTableView 
            Caption         =   "&Edit ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   3
            Top             =   800
            Width           =   1245
         End
         Begin VB.CommandButton cmdAddTableView 
            Caption         =   "&Add ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   2
            Top             =   300
            Width           =   1245
         End
         Begin SSDataWidgets_B.SSDBGrid grdTableViews 
            Height          =   4380
            Left            =   195
            TabIndex        =   1
            Top             =   300
            Width           =   6400
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   14
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   265
            Columns.Count   =   14
            Columns(0).Width=   8837
            Columns(0).Caption=   "Table (View)"
            Columns(0).Name =   "TableView"
            Columns(0).CaptionAlignment=   2
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "TableID"
            Columns(1).Name =   "TableID"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "ViewID"
            Columns(2).Name =   "ViewID"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "ButtonLinkPromptText"
            Columns(3).Name =   "ButtonLinkPromptText"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "ButtonLinkButtonText"
            Columns(4).Name =   "ButtonLinkButtonText"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "HypertextLinkText"
            Columns(5).Name =   "HypertextLinkText"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   32000
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "DropdownListLinkText"
            Columns(6).Name =   "DropdownListLinkText"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "ButtonLink"
            Columns(7).Name =   "ButtonLink"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "HypertextLink"
            Columns(8).Name =   "HypertextLink"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "DropdownListLink"
            Columns(9).Name =   "DropdownListLink"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "LinksLinkText"
            Columns(10).Name=   "LinksLinkText"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "PageTitle"
            Columns(11).Name=   "PageTitle"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   1958
            Columns(12).Caption=   "Single Record"
            Columns(12).Name=   "SingleRecord"
            Columns(12).Alignment=   2
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   11
            Columns(12).FieldLen=   256
            Columns(12).Style=   2
            Columns(13).Width=   3200
            Columns(13).Visible=   0   'False
            Columns(13).Caption=   "WFOutOfOffice"
            Columns(13).Name=   "WFOutOfOffice"
            Columns(13).DataField=   "Column 13"
            Columns(13).DataType=   11
            Columns(13).FieldLen=   256
            Columns(13).Style=   2
            TabNavigation   =   1
            _ExtentX        =   11289
            _ExtentY        =   7726
            _StockProps     =   79
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdMoveTableViewDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   6720
            TabIndex        =   7
            Top             =   2805
            Width           =   1245
         End
         Begin VB.CommandButton cmdMoveTableViewUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   6720
            TabIndex        =   6
            Top             =   2295
            Width           =   1245
         End
      End
      Begin VB.Frame fraButtonLinks 
         Caption         =   "Dashboard Links :"
         Enabled         =   0   'False
         Height          =   4935
         Left            =   150
         TabIndex        =   49
         Top             =   405
         Width           =   8120
         Begin VB.CommandButton cmdPreview 
            Caption         =   "Preview..."
            Height          =   400
            Left            =   6720
            TabIndex        =   57
            Top             =   615
            Width           =   1245
         End
         Begin VB.ComboBox cboSecurityGroup 
            Height          =   315
            Left            =   2100
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   675
            Width           =   4515
         End
         Begin VB.ComboBox cboButtonLinkView 
            Height          =   315
            Left            =   2100
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   300
            Width           =   4515
         End
         Begin VB.CommandButton cmdCopyButtonLink 
            Caption         =   "Cop&y ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   22
            Top             =   2100
            Width           =   1245
         End
         Begin VB.CommandButton cmdAddButtonLink 
            Caption         =   "&Add ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   20
            Top             =   1095
            Width           =   1245
         End
         Begin VB.CommandButton cmdEditButtonLink 
            Caption         =   "&Edit ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   21
            Top             =   1590
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveButtonLink 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   6720
            TabIndex        =   23
            Top             =   2610
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveAllButtonLinks 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   6720
            TabIndex        =   24
            Top             =   3120
            Width           =   1245
         End
         Begin SSDataWidgets_B.SSDBGrid grdButtonLinks 
            Height          =   3585
            Index           =   0
            Left            =   195
            TabIndex        =   19
            Top             =   1110
            Visible         =   0   'False
            Width           =   6420
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   55
            stylesets.count =   2
            stylesets(0).Name=   "ssEnabled"
            stylesets(0).ForeColor=   0
            stylesets(0).BackColor=   16777215
            stylesets(0).HasFont=   -1  'True
            BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            stylesets(0).Picture=   "frmSSIntranetSetup.frx":0098
            stylesets(1).Name=   "ssDisabled"
            stylesets(1).ForeColor=   0
            stylesets(1).BackColor=   16777215
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
            stylesets(1).Picture=   "frmSSIntranetSetup.frx":00B4
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   265
            Columns.Count   =   55
            Columns(0).Width=   4392
            Columns(0).Caption=   "Prompt"
            Columns(0).Name =   "Prompt"
            Columns(0).CaptionAlignment=   0
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   11
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   7223
            Columns(1).Caption=   "Element Text"
            Columns(1).Name =   "ButtonText"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "URL"
            Columns(2).Name =   "URL"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "HRProScreenID"
            Columns(3).Name =   "HRProScreenID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "PageTitle"
            Columns(4).Name =   "PageTitle"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "startMode"
            Columns(5).Name =   "startMode"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "UtilityType"
            Columns(6).Name =   "UtilityType"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "UtilityID"
            Columns(7).Name =   "UtilityID"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "HiddenGroups"
            Columns(8).Name =   "HiddenGroups"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   32000
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "NewWindow"
            Columns(9).Name =   "NewWindow"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "EMailAddress"
            Columns(10).Name=   "EMailAddress"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "EMailSubject"
            Columns(11).Name=   "EMailSubject"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   3200
            Columns(12).Visible=   0   'False
            Columns(12).Caption=   "AppFilePath"
            Columns(12).Name=   "AppFilePath"
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   8
            Columns(12).FieldLen=   256
            Columns(13).Width=   3200
            Columns(13).Visible=   0   'False
            Columns(13).Caption=   "AppParameters"
            Columns(13).Name=   "AppParameters"
            Columns(13).DataField=   "Column 13"
            Columns(13).DataType=   8
            Columns(13).FieldLen=   256
            Columns(14).Width=   3200
            Columns(14).Visible=   0   'False
            Columns(14).Caption=   "Element_Type"
            Columns(14).Name=   "Element_Type"
            Columns(14).DataField=   "Column 14"
            Columns(14).DataType=   8
            Columns(14).FieldLen=   256
            Columns(15).Width=   3200
            Columns(15).Visible=   0   'False
            Columns(15).Caption=   "SeparatorOrientation"
            Columns(15).Name=   "SeparatorOrientation"
            Columns(15).DataField=   "Column 15"
            Columns(15).DataType=   8
            Columns(15).FieldLen=   256
            Columns(16).Width=   3200
            Columns(16).Visible=   0   'False
            Columns(16).Caption=   "PictureID"
            Columns(16).Name=   "PictureID"
            Columns(16).DataField=   "Column 16"
            Columns(16).DataType=   8
            Columns(16).FieldLen=   256
            Columns(17).Width=   3200
            Columns(17).Visible=   0   'False
            Columns(17).Caption=   "ChartShowLegend"
            Columns(17).Name=   "ChartShowLegend"
            Columns(17).DataField=   "Column 17"
            Columns(17).DataType=   8
            Columns(17).FieldLen=   256
            Columns(18).Width=   3200
            Columns(18).Visible=   0   'False
            Columns(18).Caption=   "ChartType"
            Columns(18).Name=   "ChartType"
            Columns(18).DataField=   "Column 18"
            Columns(18).DataType=   8
            Columns(18).FieldLen=   256
            Columns(19).Width=   3200
            Columns(19).Visible=   0   'False
            Columns(19).Caption=   "ChartShowGrid"
            Columns(19).Name=   "ChartShowGrid"
            Columns(19).DataField=   "Column 19"
            Columns(19).DataType=   8
            Columns(19).FieldLen=   256
            Columns(20).Width=   3200
            Columns(20).Visible=   0   'False
            Columns(20).Caption=   "ChartStackSeries"
            Columns(20).Name=   "ChartStackSeries"
            Columns(20).DataField=   "Column 20"
            Columns(20).DataType=   8
            Columns(20).FieldLen=   256
            Columns(21).Width=   3200
            Columns(21).Visible=   0   'False
            Columns(21).Caption=   "ChartViewID"
            Columns(21).Name=   "ChartViewID"
            Columns(21).DataField=   "Column 21"
            Columns(21).DataType=   8
            Columns(21).FieldLen=   256
            Columns(22).Width=   3200
            Columns(22).Visible=   0   'False
            Columns(22).Caption=   "ChartTableID"
            Columns(22).Name=   "ChartTableID"
            Columns(22).DataField=   "Column 22"
            Columns(22).DataType=   8
            Columns(22).FieldLen=   256
            Columns(23).Width=   3200
            Columns(23).Visible=   0   'False
            Columns(23).Caption=   "ChartColumnID"
            Columns(23).Name=   "ChartColumnID"
            Columns(23).DataField=   "Column 23"
            Columns(23).DataType=   8
            Columns(23).FieldLen=   256
            Columns(24).Width=   3200
            Columns(24).Visible=   0   'False
            Columns(24).Caption=   "ChartFilterID"
            Columns(24).Name=   "ChartFilterID"
            Columns(24).DataField=   "Column 24"
            Columns(24).DataType=   8
            Columns(24).FieldLen=   256
            Columns(25).Width=   3200
            Columns(25).Visible=   0   'False
            Columns(25).Caption=   "ChartAggregateType"
            Columns(25).Name=   "ChartAggregateType"
            Columns(25).DataField=   "Column 25"
            Columns(25).DataType=   8
            Columns(25).FieldLen=   256
            Columns(26).Width=   3200
            Columns(26).Visible=   0   'False
            Columns(26).Caption=   "ChartShowValues"
            Columns(26).Name=   "ChartShowValues"
            Columns(26).DataField=   "Column 26"
            Columns(26).DataType=   8
            Columns(26).FieldLen=   256
            Columns(27).Width=   3200
            Columns(27).Visible=   0   'False
            Columns(27).Caption=   "UseFormatting"
            Columns(27).Name=   "UseFormatting"
            Columns(27).DataField=   "Column 27"
            Columns(27).DataType=   8
            Columns(27).FieldLen=   256
            Columns(28).Width=   3200
            Columns(28).Visible=   0   'False
            Columns(28).Caption=   "Formatting_DecimalPlaces"
            Columns(28).Name=   "Formatting_DecimalPlaces"
            Columns(28).DataField=   "Column 28"
            Columns(28).DataType=   8
            Columns(28).FieldLen=   256
            Columns(29).Width=   3200
            Columns(29).Visible=   0   'False
            Columns(29).Caption=   "Formatting_Use1000Separator"
            Columns(29).Name=   "Formatting_Use1000Separator"
            Columns(29).DataField=   "Column 29"
            Columns(29).DataType=   8
            Columns(29).FieldLen=   256
            Columns(30).Width=   3200
            Columns(30).Visible=   0   'False
            Columns(30).Caption=   "Formatting_Prefix"
            Columns(30).Name=   "Formatting_Prefix"
            Columns(30).DataField=   "Column 30"
            Columns(30).DataType=   8
            Columns(30).FieldLen=   256
            Columns(31).Width=   3200
            Columns(31).Visible=   0   'False
            Columns(31).Caption=   "Formatting_Suffix"
            Columns(31).Name=   "Formatting_Suffix"
            Columns(31).DataField=   "Column 31"
            Columns(31).DataType=   8
            Columns(31).FieldLen=   256
            Columns(32).Width=   3200
            Columns(32).Visible=   0   'False
            Columns(32).Caption=   "UseConditionalFormatting"
            Columns(32).Name=   "UseConditionalFormatting"
            Columns(32).DataField=   "Column 32"
            Columns(32).DataType=   8
            Columns(32).FieldLen=   256
            Columns(33).Width=   3200
            Columns(33).Visible=   0   'False
            Columns(33).Caption=   "ConditionalFormatting_Operator_1"
            Columns(33).Name=   "ConditionalFormatting_Operator_1"
            Columns(33).DataField=   "Column 33"
            Columns(33).DataType=   8
            Columns(33).FieldLen=   256
            Columns(34).Width=   3200
            Columns(34).Visible=   0   'False
            Columns(34).Caption=   "ConditionalFormatting_Value_1"
            Columns(34).Name=   "ConditionalFormatting_Value_1"
            Columns(34).DataField=   "Column 34"
            Columns(34).DataType=   8
            Columns(34).FieldLen=   256
            Columns(35).Width=   3200
            Columns(35).Visible=   0   'False
            Columns(35).Caption=   "ConditionalFormatting_Style_1"
            Columns(35).Name=   "ConditionalFormatting_Style_1"
            Columns(35).DataField=   "Column 35"
            Columns(35).DataType=   8
            Columns(35).FieldLen=   256
            Columns(36).Width=   3200
            Columns(36).Visible=   0   'False
            Columns(36).Caption=   "ConditionalFormatting_Colour_1"
            Columns(36).Name=   "ConditionalFormatting_Colour_1"
            Columns(36).DataField=   "Column 36"
            Columns(36).DataType=   8
            Columns(36).FieldLen=   256
            Columns(37).Width=   3200
            Columns(37).Visible=   0   'False
            Columns(37).Caption=   "ConditionalFormatting_Operator_2"
            Columns(37).Name=   "ConditionalFormatting_Operator_2"
            Columns(37).DataField=   "Column 37"
            Columns(37).DataType=   8
            Columns(37).FieldLen=   256
            Columns(38).Width=   3200
            Columns(38).Visible=   0   'False
            Columns(38).Caption=   "ConditionalFormatting_Value_2"
            Columns(38).Name=   "ConditionalFormatting_Value_2"
            Columns(38).DataField=   "Column 38"
            Columns(38).DataType=   8
            Columns(38).FieldLen=   256
            Columns(39).Width=   3200
            Columns(39).Visible=   0   'False
            Columns(39).Caption=   "ConditionalFormatting_Style_2"
            Columns(39).Name=   "ConditionalFormatting_Style_2"
            Columns(39).DataField=   "Column 39"
            Columns(39).DataType=   8
            Columns(39).FieldLen=   256
            Columns(40).Width=   3200
            Columns(40).Visible=   0   'False
            Columns(40).Caption=   "ConditionalFormatting_Colour_2"
            Columns(40).Name=   "ConditionalFormatting_Colour_2"
            Columns(40).DataField=   "Column 40"
            Columns(40).DataType=   8
            Columns(40).FieldLen=   256
            Columns(41).Width=   3200
            Columns(41).Visible=   0   'False
            Columns(41).Caption=   "ConditionalFormatting_Operator_3"
            Columns(41).Name=   "ConditionalFormatting_Operator_3"
            Columns(41).DataField=   "Column 41"
            Columns(41).DataType=   8
            Columns(41).FieldLen=   256
            Columns(42).Width=   3200
            Columns(42).Visible=   0   'False
            Columns(42).Caption=   "ConditionalFormatting_Value_3"
            Columns(42).Name=   "ConditionalFormatting_Value_3"
            Columns(42).DataField=   "Column 42"
            Columns(42).DataType=   8
            Columns(42).FieldLen=   256
            Columns(43).Width=   3200
            Columns(43).Visible=   0   'False
            Columns(43).Caption=   "ConditionalFormatting_Style_3"
            Columns(43).Name=   "ConditionalFormatting_Style_3"
            Columns(43).DataField=   "Column 43"
            Columns(43).DataType=   8
            Columns(43).FieldLen=   256
            Columns(44).Width=   3200
            Columns(44).Visible=   0   'False
            Columns(44).Caption=   "ConditionalFormatting_Colour_3"
            Columns(44).Name=   "ConditionalFormatting_Colour_3"
            Columns(44).DataField=   "Column 44"
            Columns(44).DataType=   8
            Columns(44).FieldLen=   256
            Columns(45).Width=   3200
            Columns(45).Visible=   0   'False
            Columns(45).Caption=   "SeparatorColour"
            Columns(45).Name=   "SeparatorColour"
            Columns(45).DataField=   "Column 45"
            Columns(45).DataType=   8
            Columns(45).FieldLen=   256
            Columns(46).Width=   3200
            Columns(46).Visible=   0   'False
            Columns(46).Caption=   "InitialDisplayMode"
            Columns(46).Name=   "InitialDisplayMode"
            Columns(46).DataField=   "Column 46"
            Columns(46).DataType=   8
            Columns(46).FieldLen=   256
            Columns(47).Width=   3200
            Columns(47).Visible=   0   'False
            Columns(47).Caption=   "Chart_TableID_2"
            Columns(47).Name=   "Chart_TableID_2"
            Columns(47).DataField=   "Column 47"
            Columns(47).DataType=   8
            Columns(47).FieldLen=   256
            Columns(48).Width=   3200
            Columns(48).Visible=   0   'False
            Columns(48).Caption=   "Chart_ColumnID_2"
            Columns(48).Name=   "Chart_ColumnID_2"
            Columns(48).DataField=   "Column 48"
            Columns(48).DataType=   8
            Columns(48).FieldLen=   256
            Columns(49).Width=   3200
            Columns(49).Visible=   0   'False
            Columns(49).Caption=   "Chart_TableID_3"
            Columns(49).Name=   "Chart_TableID_3"
            Columns(49).DataField=   "Column 49"
            Columns(49).DataType=   8
            Columns(49).FieldLen=   256
            Columns(50).Width=   3200
            Columns(50).Visible=   0   'False
            Columns(50).Caption=   "Chart_ColumnID_3"
            Columns(50).Name=   "Chart_ColumnID_3"
            Columns(50).DataField=   "Column 50"
            Columns(50).DataType=   8
            Columns(50).FieldLen=   256
            Columns(51).Width=   3200
            Columns(51).Visible=   0   'False
            Columns(51).Caption=   "Chart_SortOrderID"
            Columns(51).Name=   "Chart_SortOrderID"
            Columns(51).DataField=   "Column 51"
            Columns(51).DataType=   8
            Columns(51).FieldLen=   256
            Columns(52).Width=   3200
            Columns(52).Visible=   0   'False
            Columns(52).Caption=   "Chart_SortDirection"
            Columns(52).Name=   "Chart_SortDirection"
            Columns(52).DataField=   "Column 52"
            Columns(52).DataType=   8
            Columns(52).FieldLen=   256
            Columns(53).Width=   3200
            Columns(53).Visible=   0   'False
            Columns(53).Caption=   "Chart_ColourID"
            Columns(53).Name=   "Chart_ColourID"
            Columns(53).DataField=   "Column 53"
            Columns(53).DataType=   8
            Columns(53).FieldLen=   256
            Columns(54).Width=   3200
            Columns(54).Visible=   0   'False
            Columns(54).Caption=   "ChartShowPercentages"
            Columns(54).Name=   "ChartShowPercentages"
            Columns(54).DataField=   "Column 54"
            Columns(54).DataType=   8
            Columns(54).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   11324
            _ExtentY        =   6324
            _StockProps     =   79
            Enabled         =   0   'False
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdMoveButtonLinkUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   6720
            TabIndex        =   25
            Top             =   3615
            Width           =   1245
         End
         Begin VB.CommandButton cmdMoveButtonLinkDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   6720
            TabIndex        =   26
            Top             =   4110
            Width           =   1245
         End
         Begin VB.Label lblSecurityGroup 
            Caption         =   "User Group :"
            Height          =   195
            Left            =   195
            TabIndex        =   56
            Top             =   705
            Width           =   1695
         End
         Begin VB.Label lblButtonLinkView 
            Caption         =   "Table (View) :"
            Height          =   195
            Left            =   195
            TabIndex        =   50
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraDropdownListLinks 
         Caption         =   "Dropdown List Links :"
         Enabled         =   0   'False
         Height          =   4935
         Left            =   -74850
         TabIndex        =   51
         Top             =   405
         Width           =   8120
         Begin VB.ComboBox cboDropdownListLinkView 
            Height          =   315
            Left            =   2100
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   300
            Width           =   4515
         End
         Begin VB.CommandButton cmdCopyDropdownListLink 
            Caption         =   "Cop&y ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   31
            Top             =   1750
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveAllDropdownListLinks 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   6720
            TabIndex        =   33
            Top             =   2745
            Width           =   1245
         End
         Begin VB.CommandButton cmdRemoveDropdownListLink 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   6720
            TabIndex        =   32
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CommandButton cmdEditDropdownListLink 
            Caption         =   "&Edit ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   30
            Top             =   1250
            Width           =   1245
         End
         Begin VB.CommandButton cmdAddDropdownListLink 
            Caption         =   "&Add ..."
            Height          =   400
            Left            =   6720
            TabIndex        =   29
            Top             =   750
            Width           =   1245
         End
         Begin SSDataWidgets_B.SSDBGrid grdDropdownListLinks 
            Height          =   3945
            Index           =   0
            Left            =   195
            TabIndex        =   28
            Top             =   750
            Visible         =   0   'False
            Width           =   6400
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   13
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   79
            Columns.Count   =   13
            Columns(0).Width=   10821
            Columns(0).Caption=   "Text"
            Columns(0).Name =   "Text"
            Columns(0).CaptionAlignment=   2
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "URL"
            Columns(1).Name =   "URL"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   3200
            Columns(2).Visible=   0   'False
            Columns(2).Caption=   "HRProScreenID"
            Columns(2).Name =   "HRProScreenID"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "PageTitle"
            Columns(3).Name =   "PageTitle"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "startMode"
            Columns(4).Name =   "startMode"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "UtilityType"
            Columns(5).Name =   "UtilityType"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   3200
            Columns(6).Visible=   0   'False
            Columns(6).Caption=   "UtilityID"
            Columns(6).Name =   "UtilityID"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "HiddenGroups"
            Columns(7).Name =   "HiddenGroups"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   32000
            Columns(8).Width=   3200
            Columns(8).Visible=   0   'False
            Columns(8).Caption=   "NewWindow"
            Columns(8).Name =   "NewWindow"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "EMailAddress"
            Columns(9).Name =   "EMailAddress"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   3200
            Columns(10).Visible=   0   'False
            Columns(10).Caption=   "EMailSubject"
            Columns(10).Name=   "EMailSubject"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "AppFilePath"
            Columns(11).Name=   "AppFilePath"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   3200
            Columns(12).Visible=   0   'False
            Columns(12).Caption=   "AppParameters"
            Columns(12).Name=   "AppParameters"
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   8
            Columns(12).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   11289
            _ExtentY        =   6959
            _StockProps     =   79
            Enabled         =   0   'False
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdMoveDropdownListLinkDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   6720
            TabIndex        =   35
            Top             =   3750
            Width           =   1245
         End
         Begin VB.CommandButton cmdMoveDropdownListLinkUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   6720
            TabIndex        =   34
            Top             =   3255
            Width           =   1245
         End
         Begin VB.Label lblDropdownListLinkView 
            Caption         =   "Table (View) :"
            Height          =   195
            Left            =   195
            TabIndex        =   52
            Top             =   360
            Width           =   1650
         End
      End
   End
End
Attribute VB_Name = "frmSSIntranetSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' CONSTANTS
'
' Page number constants.
Private Const giPAGE_GENERAL = 0
Private Const giPAGE_HYPERTEXTLINKS = 1
Private Const giPAGE_BUTTONLINKS = 2
Private Const giPAGE_DROPDOWNLISTLINKS = 3
Private Const giPAGE_DOCUMENTS = 4

Public Enum SSINTRANETLINKTYPES
  SSINTLINK_HYPERTEXT = 0
  SSINTLINK_BUTTON = 1
  SSINTLINK_DROPDOWNLIST = 2
  SSINTLINK_DOCUMENT = 3
End Enum

Private mblnReadOnly As Boolean
Private mfChanged As Boolean
Private mlngPersonnelTableID As Long
Private mfLoading As Boolean

Private Enum MoveDirection
  MOVEDIRECTION_UP = 0
  MOVEDIRECTION_DOWN = 1
End Enum

Private mcolSSITableViews As clsSSITableViews
Private mcolGroups As Collection

Private mfIsEditingSeparator As Boolean
Private mfIsCopyingSeparator As Boolean

Public Property Get SSITableViewsCollection() As clsSSITableViews
  SSITableViewsCollection = mcolSSITableViews
End Property

Private Function CurrentLinkGrid(piLinkType As SSINTRANETLINKTYPES) As SSDBGrid

  Dim ctlGridArray As Variant
  Dim ctlGrid As SSDBGrid
  Dim ctlCurrentGrid As SSDBGrid
  Dim ctlTableViewCombo As ComboBox
  
  Select Case piLinkType
    Case SSINTLINK_BUTTON
      Set ctlGridArray = grdButtonLinks
      Set ctlTableViewCombo = cboButtonLinkView
    Case SSINTLINK_DROPDOWNLIST
      Set ctlGridArray = grdDropdownListLinks
      Set ctlTableViewCombo = cboDropdownListLinkView
    Case SSINTLINK_HYPERTEXT
      Set ctlGridArray = grdHypertextLinks
      Set ctlTableViewCombo = cboHypertextLinkView
    Case SSINTLINK_DOCUMENT
      Set ctlGridArray = grdDocuments
      Set ctlTableViewCombo = cboDocumentView
  End Select
  
  If ctlTableViewCombo.ListIndex >= 0 Then
    For Each ctlGrid In ctlGridArray
      If ctlGrid.Tag = GetTagKeyFromCollection(mcolSSITableViews, ctlTableViewCombo.List(ctlTableViewCombo.ListIndex)) Then
        Set ctlCurrentGrid = ctlGrid
        Exit For
      End If
    Next ctlGrid
    Set ctlGrid = Nothing
  End If
  
  Set ctlGridArray = Nothing
  Set ctlTableViewCombo = Nothing
  
  Set CurrentLinkGrid = ctlCurrentGrid
  
End Function

Private Sub RefreshControls()

  Dim ctlGrid As SSDBGrid
  Dim iIndex As Integer
  
  Select Case ssTabStrip.Tab
    Case giPAGE_GENERAL
'      cboPersonnelTable.Enabled = (cboPersonnelTable.ListCount > 1) And _
'        (Not mblnReadOnly)
'      cboPersonnelTable.BackColor = IIf(cboPersonnelTable.Enabled, vbWindowBackground, vbButtonFace)
'      lblPersonnelTable.Enabled = cboPersonnelTable.Enabled

      cmdAddTableView.Enabled = (Not mblnReadOnly)

      Set ctlGrid = grdTableViews
      
      With ctlGrid
        If .Rows = 0 Then
          cmdEditTableView.Enabled = False
          cmdRemoveTableView.Enabled = False
          cmdRemoveAllTableViews.Enabled = False
        Else
          If .SelBookmarks.Count > 0 Then
            cmdEditTableView.Enabled = (Not mblnReadOnly) And _
              (.SelBookmarks.Count = 1)
            cmdRemoveTableView.Enabled = Not mblnReadOnly
          Else
            cmdEditTableView.Enabled = False
            cmdRemoveTableView.Enabled = False
          End If
          
          cmdRemoveAllTableViews.Enabled = Not mblnReadOnly
        End If
      
        If .SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveTableViewUp.Enabled = False
            cmdMoveTableViewDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveTableViewUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveTableViewDown.Enabled = False
          Else
            cmdMoveTableViewUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveTableViewDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          End If
        Else
          cmdMoveTableViewUp.Enabled = False
          cmdMoveTableViewDown.Enabled = False
        End If
      End With
      Set ctlGrid = Nothing

    Case giPAGE_HYPERTEXTLINKS
      iIndex = 0
      Set ctlGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
      If Not ctlGrid Is Nothing Then
        iIndex = ctlGrid.Index
      End If
      Set ctlGrid = Nothing
      
      cmdAddHyperTextLink.Enabled = fraHypertextLinks.Enabled _
        And (iIndex > 0) And (Not mblnReadOnly)
      
      Set ctlGrid = grdHypertextLinks(iIndex)
      
      With ctlGrid
        If .Rows = 0 Then
          cmdEditHypertextLink.Enabled = False
          cmdRemoveHypertextLink.Enabled = False
          cmdRemoveAllHypertextLinks.Enabled = False
        Else
          If .SelBookmarks.Count > 0 Then
            cmdEditHypertextLink.Enabled = (Not mblnReadOnly) And _
              (.SelBookmarks.Count = 1)
            cmdRemoveHypertextLink.Enabled = Not mblnReadOnly
          Else
            cmdEditHypertextLink.Enabled = False
            cmdRemoveHypertextLink.Enabled = False
          End If
          
          cmdRemoveAllHypertextLinks.Enabled = Not mblnReadOnly
        End If
    
        cmdCopyHypertextLink.Enabled = cmdEditHypertextLink.Enabled
        
        If .SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveHypertextLinkUp.Enabled = False
            cmdMoveHypertextLinkDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveHypertextLinkUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveHypertextLinkDown.Enabled = False
          Else
            cmdMoveHypertextLinkUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveHypertextLinkDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          End If
        Else
          cmdMoveHypertextLinkUp.Enabled = False
          cmdMoveHypertextLinkDown.Enabled = False
        End If
      End With
      Set ctlGrid = Nothing
      
      For Each ctlGrid In grdHypertextLinks
        ctlGrid.Visible = (ctlGrid.Index = iIndex)
      Next ctlGrid
      Set ctlGrid = Nothing
      
    Case giPAGE_BUTTONLINKS
      iIndex = 0
      Set ctlGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
      If Not ctlGrid Is Nothing Then
        iIndex = ctlGrid.Index
      End If
      Set ctlGrid = Nothing
      
      cmdAddButtonLink.Enabled = fraButtonLinks.Enabled _
        And (iIndex > 0) And (Not mblnReadOnly)
      
      Set ctlGrid = grdButtonLinks(iIndex)
      
      With ctlGrid
        If .Rows = 0 Then
          cmdEditButtonLink.Enabled = False
          cmdRemoveButtonLink.Enabled = False
          cmdRemoveAllButtonLinks.Enabled = False
        Else
          If .SelBookmarks.Count > 0 Then
            cmdEditButtonLink.Enabled = (Not mblnReadOnly) And _
              (.SelBookmarks.Count = 1)
            cmdRemoveButtonLink.Enabled = Not mblnReadOnly
          Else
            cmdEditButtonLink.Enabled = False
            cmdRemoveButtonLink.Enabled = False
          End If
          
          cmdRemoveAllButtonLinks.Enabled = Not mblnReadOnly
        End If
      
        cmdCopyButtonLink.Enabled = cmdEditButtonLink.Enabled
        
        If .SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveButtonLinkUp.Enabled = False
            cmdMoveButtonLinkDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveButtonLinkUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveButtonLinkDown.Enabled = False
          Else
            cmdMoveButtonLinkUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveButtonLinkDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          End If
        Else
          cmdMoveButtonLinkUp.Enabled = False
          cmdMoveButtonLinkDown.Enabled = False
        End If
      End With
      Set ctlGrid = Nothing

      For Each ctlGrid In grdButtonLinks
        ctlGrid.Visible = (ctlGrid.Index = iIndex)
      Next ctlGrid
      Set ctlGrid = Nothing
      
    Case giPAGE_DROPDOWNLISTLINKS
      iIndex = 0
      Set ctlGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
      If Not ctlGrid Is Nothing Then
        iIndex = ctlGrid.Index
      End If
      Set ctlGrid = Nothing
      
      cmdAddDropdownListLink.Enabled = fraDropdownListLinks.Enabled _
        And (iIndex > 0) And (Not mblnReadOnly)
      
      Set ctlGrid = grdDropdownListLinks(iIndex)
      
      With ctlGrid
        If .Rows = 0 Then
          cmdEditDropdownListLink.Enabled = False
          cmdRemoveDropdownListLink.Enabled = False
          cmdRemoveAllDropdownListLinks.Enabled = False
        Else
          If .SelBookmarks.Count > 0 Then
            cmdEditDropdownListLink.Enabled = (Not mblnReadOnly) And _
              (.SelBookmarks.Count = 1)
            cmdRemoveDropdownListLink.Enabled = Not mblnReadOnly
          Else
            cmdEditDropdownListLink.Enabled = False
            cmdRemoveDropdownListLink.Enabled = False
          End If
          
          cmdRemoveAllDropdownListLinks.Enabled = Not mblnReadOnly
        End If
    
        cmdCopyDropdownListLink.Enabled = cmdEditDropdownListLink.Enabled
    
        If .SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveDropdownListLinkUp.Enabled = False
            cmdMoveDropdownListLinkDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveDropdownListLinkUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveDropdownListLinkDown.Enabled = False
          Else
            cmdMoveDropdownListLinkUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveDropdownListLinkDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          End If
        Else
          cmdMoveDropdownListLinkUp.Enabled = False
          cmdMoveDropdownListLinkDown.Enabled = False
        End If
      End With
      Set ctlGrid = Nothing
    
      For Each ctlGrid In grdDropdownListLinks
        ctlGrid.Visible = (ctlGrid.Index = iIndex)
      Next ctlGrid
      Set ctlGrid = Nothing
    
    Case giPAGE_DOCUMENTS
      iIndex = 0
      Set ctlGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
      If Not ctlGrid Is Nothing Then
        iIndex = ctlGrid.Index
      End If
      Set ctlGrid = Nothing
      
      cmdAddDocument.Enabled = fraDocuments.Enabled _
        And (iIndex > 0) And (Not mblnReadOnly)
      
      Set ctlGrid = grdDocuments(iIndex)
      
      With ctlGrid
        If .Rows = 0 Then
          cmdEditDocument.Enabled = False
          cmdRemoveDocument.Enabled = False
          cmdRemoveAllDocuments.Enabled = False
        Else
          If .SelBookmarks.Count > 0 Then
            cmdEditDocument.Enabled = (Not mblnReadOnly) And _
              (.SelBookmarks.Count = 1)
            cmdRemoveDocument.Enabled = Not mblnReadOnly
          Else
            cmdEditDocument.Enabled = False
            cmdRemoveDocument.Enabled = False
          End If
          
          cmdRemoveAllDocuments.Enabled = Not mblnReadOnly
        End If
    
        cmdCopyDocument.Enabled = cmdEditDocument.Enabled
    
        If .SelBookmarks.Count = 1 Then
          If .AddItemRowIndex(.Bookmark) = 0 Then
            cmdMoveDocumentUp.Enabled = False
            cmdMoveDocumentDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
            cmdMoveDocumentUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveDocumentDown.Enabled = False
          Else
            cmdMoveDocumentUp.Enabled = (.Rows > 1) And (Not mblnReadOnly)
            cmdMoveDocumentDown.Enabled = (.Rows > 1) And (Not mblnReadOnly)
          End If
        Else
          cmdMoveDocumentUp.Enabled = False
          cmdMoveDocumentDown.Enabled = False
        End If
      End With
      Set ctlGrid = Nothing
    
      For Each ctlGrid In grdDocuments
        ctlGrid.Visible = (ctlGrid.Index = iIndex)
      Next ctlGrid
      Set ctlGrid = Nothing

  End Select

  cmdOk.Enabled = mfChanged
  cmdPreview.Enabled = (cboSecurityGroup.Text <> "(All Groups)")

End Sub

Private Sub RemoveLink(pctlGrid As SSDBGrid)

  Dim sRowsToDelete As String
  Dim iCount As Integer
  
  sRowsToDelete = ","

  With pctlGrid
    If .Rows = 1 Then
      .RemoveAll
    Else
      For iCount = 0 To .SelBookmarks.Count - 1
        sRowsToDelete = sRowsToDelete & CStr(.AddItemRowIndex(.SelBookmarks(iCount))) & ","
      Next iCount
      
      For iCount = (.Rows - 1) To 0 Step -1
        If InStr(sRowsToDelete, "," & CStr(iCount) & ",") > 0 Then
          .RemoveItem iCount
        End If
      Next iCount
    End If
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  Changed = True

End Sub

Private Sub SaveLinkParameters(piLinkType As SSINTRANETLINKTYPES)

  ' Save the link information to the database.
  Dim iLoop As Integer
  Dim sSQL As String
  Dim varBookMark As Variant
  Dim ctlGrid As SSDBGrid
  Dim sPrompt As String
  Dim sText As String
  Dim sScreenID As String
  Dim sPageTitle As String
  Dim sURL As String
  Dim sStartMode As String
  Dim sUtilityType As String
  Dim sUtilityID As String
  Dim sNewWindow As String
  Dim lngMaxID As Long
  Dim rsTemp As DAO.Recordset
  Dim sGroupName As String
  Dim sGroupNames As String
  Dim sTableID As String
  Dim sViewID As String
  Dim ctlGridArray As Variant
  'NPG20080211 Fault 12873
  Dim sEMailAddress As String
  Dim sEMailSubject As String
  Dim sAppFilePath As String
  Dim sAppParameters As String
  Dim sDocumentFilePath As String
  Dim fDisplayDocumentHyperlink As Boolean
  Dim sElement_Type As Integer
  'NPG20100111
  Dim sSeparatorOrientation As String
  Dim sPictureID As String
  Dim fChartShowLegend As Boolean
  Dim sChartType As Integer
  Dim fChartShowGrid As Boolean
  Dim fChartStackSeries As Boolean
  Dim sChartViewID As String
  Dim fChartShowPercentages As Boolean
  Dim sChartTableID As String
  Dim sChartColumnID As String
  Dim sChartFilterID As String
  Dim sChartAggregateType As Integer
  Dim fChartShowValues As Boolean
  Dim fUseFormatting As Boolean
  Dim iFormatting_DecimalPlaces As Integer
  Dim fFormatting_Use1000Separator As Boolean
  Dim sFormatting_Prefix As String
  Dim sFormatting_Suffix As String
  Dim fUseConditionalFormatting As Boolean
  Dim sConditionalFormatting_Operator_1 As String
  Dim sConditionalFormatting_Value_1 As String
  Dim sConditionalFormatting_Style_1 As String
  Dim sConditionalFormatting_Colour_1 As String
  Dim sConditionalFormatting_Operator_2 As String
  Dim sConditionalFormatting_Value_2 As String
  Dim sConditionalFormatting_Style_2 As String
  Dim sConditionalFormatting_Colour_2 As String
  Dim sConditionalFormatting_Operator_3 As String
  Dim sConditionalFormatting_Value_3 As String
  Dim sConditionalFormatting_Style_3 As String
  Dim sConditionalFormatting_Colour_3 As String
  Dim sSeparatorColour As String
  Dim iInitialDisplayMode As Integer
  Dim sChart_TableID_2 As String
  Dim sChart_ColumnID_2 As String
  Dim sChart_TableID_3 As String
  Dim sChart_ColumnID_3 As String
  Dim sChart_SortOrderID As String
  Dim sChart_SortDirection As String
  Dim sChart_ColourID As String
  Dim iRed As Integer
  Dim iGreen As Integer
  Dim iBlue As Integer
  Dim fOK As Boolean

  Select Case piLinkType
    Case SSINTLINK_HYPERTEXT
      Set ctlGridArray = grdHypertextLinks
    Case SSINTLINK_BUTTON
      Set ctlGridArray = grdButtonLinks
    Case SSINTLINK_DROPDOWNLIST
      Set ctlGridArray = grdDropdownListLinks
    Case SSINTLINK_DOCUMENT
      Set ctlGridArray = grdDocuments
  End Select
  
  For Each ctlGrid In ctlGridArray
    With ctlGrid
      If .Index > 0 Then
        For iLoop = 0 To (.Rows - 1)
          varBookMark = .AddItemBookmark(iLoop)
    
          Select Case piLinkType
            Case SSINTLINK_HYPERTEXT
              sPrompt = ""
              sText = .Columns("Text").CellText(varBookMark)
              sScreenID = ""
              sPageTitle = ""
              sURL = .Columns("URL").CellText(varBookMark)
              sStartMode = ""
              sUtilityType = .Columns("UtilityType").CellText(varBookMark)
              sUtilityID = .Columns("UtilityID").CellText(varBookMark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookMark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookMark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookMark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookMark)
              sAppParameters = .Columns("AppParameters").CellText(varBookMark)
              sElement_Type = .Columns("Element_Type").CellValue(varBookMark)
              sPictureID = .Columns("PictureID").CellText(varBookMark)

            Case SSINTLINK_BUTTON
              sPrompt = .Columns("Prompt").CellText(varBookMark)
              sText = .Columns("ButtonText").CellText(varBookMark)
              sScreenID = .Columns("HRProScreenID").CellText(varBookMark)
              sPageTitle = .Columns("PageTitle").CellText(varBookMark)
              sURL = .Columns("URL").CellText(varBookMark)
              sStartMode = .Columns("startMode").CellText(varBookMark)
              sUtilityType = .Columns("UtilityType").CellText(varBookMark)
              sUtilityID = .Columns("UtilityID").CellText(varBookMark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookMark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookMark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookMark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookMark)
              sAppParameters = .Columns("AppParameters").CellText(varBookMark)
              sElement_Type = .Columns("Element_Type").CellText(varBookMark)
              sSeparatorOrientation = .Columns("SeparatorOrientation").CellText(varBookMark)
              sPictureID = .Columns("PictureID").CellText(varBookMark)
              fChartShowLegend = .Columns("ChartShowLegend").CellText(varBookMark)
              sChartType = .Columns("ChartType").CellText(varBookMark)
              fChartShowGrid = .Columns("ChartShowGrid").CellText(varBookMark)
              fChartStackSeries = .Columns("ChartStackSeries").CellText(varBookMark)
              sChartViewID = .Columns("ChartViewID").CellText(varBookMark)
              sChartTableID = .Columns("ChartTableID").CellText(varBookMark)
              sChartColumnID = .Columns("ChartColumnID").CellText(varBookMark)
              sChartFilterID = .Columns("ChartFilterID").CellText(varBookMark)
              sChartAggregateType = val(.Columns("ChartAggregateType").CellText(varBookMark))
              fChartShowValues = val(.Columns("ChartShowValues").CellText(varBookMark))
              fUseFormatting = .Columns("UseFormatting").CellText(varBookMark)
              iFormatting_DecimalPlaces = val(.Columns("Formatting_DecimalPlaces").CellText(varBookMark))
              fFormatting_Use1000Separator = .Columns("Formatting_Use1000Separator").CellText(varBookMark)
              sFormatting_Prefix = .Columns("Formatting_Prefix").CellText(varBookMark)
              sFormatting_Suffix = .Columns("Formatting_Suffix").CellText(varBookMark)
              fUseConditionalFormatting = .Columns("UseConditionalFormatting").CellText(varBookMark)
              sConditionalFormatting_Operator_1 = .Columns("ConditionalFormatting_Operator_1").CellText(varBookMark)
              sConditionalFormatting_Value_1 = .Columns("ConditionalFormatting_Value_1").CellText(varBookMark)
              sConditionalFormatting_Style_1 = .Columns("ConditionalFormatting_Style_1").CellText(varBookMark)
              sConditionalFormatting_Colour_1 = .Columns("ConditionalFormatting_Colour_1").CellText(varBookMark)
              sConditionalFormatting_Operator_2 = .Columns("ConditionalFormatting_Operator_2").CellText(varBookMark)
              sConditionalFormatting_Value_2 = .Columns("ConditionalFormatting_Value_2").CellText(varBookMark)
              sConditionalFormatting_Style_2 = .Columns("ConditionalFormatting_Style_2").CellText(varBookMark)
              sConditionalFormatting_Colour_2 = .Columns("ConditionalFormatting_Colour_2").CellText(varBookMark)
              sConditionalFormatting_Operator_3 = .Columns("ConditionalFormatting_Operator_3").CellText(varBookMark)
              sConditionalFormatting_Value_3 = .Columns("ConditionalFormatting_Value_3").CellText(varBookMark)
              sConditionalFormatting_Style_3 = .Columns("ConditionalFormatting_Style_3").CellText(varBookMark)
              sConditionalFormatting_Colour_3 = .Columns("ConditionalFormatting_Colour_3").CellText(varBookMark)
              sSeparatorColour = .Columns("SeparatorColour").CellText(varBookMark)
              iInitialDisplayMode = val(.Columns("InitialDisplayMode").CellText(varBookMark))
              sChart_TableID_2 = .Columns("Chart_TableID_2").CellText(varBookMark)
              sChart_ColumnID_2 = .Columns("Chart_ColumnID_2").CellText(varBookMark)
              sChart_TableID_3 = .Columns("Chart_TableID_3").CellText(varBookMark)
              sChart_ColumnID_3 = .Columns("Chart_ColumnID_3").CellText(varBookMark)
              sChart_SortOrderID = .Columns("Chart_SortOrderID").CellText(varBookMark)
              sChart_SortDirection = .Columns("Chart_SortDirection").CellText(varBookMark)
              sChart_ColourID = .Columns("Chart_ColourID").CellText(varBookMark)
              fChartShowPercentages = .Columns("ChartShowPercentages").CellText(varBookMark)
                          
            Case SSINTLINK_DROPDOWNLIST
              sPrompt = ""
              sText = .Columns("Text").CellText(varBookMark)
              sScreenID = .Columns("HRProScreenID").CellText(varBookMark)
              sPageTitle = .Columns("PageTitle").CellText(varBookMark)
              sURL = .Columns("URL").CellText(varBookMark)
              sStartMode = .Columns("startMode").CellText(varBookMark)
              sUtilityType = .Columns("UtilityType").CellText(varBookMark)
              sUtilityID = .Columns("UtilityID").CellText(varBookMark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookMark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookMark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookMark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookMark)
              sAppParameters = .Columns("AppParameters").CellText(varBookMark)
            
            Case SSINTLINK_DOCUMENT
              sPrompt = ""
              sText = .Columns("Text").CellText(varBookMark)
              sScreenID = .Columns("HRProScreenID").CellText(varBookMark)
              sPageTitle = .Columns("PageTitle").CellText(varBookMark)
              sURL = .Columns("URL").CellText(varBookMark)
              sStartMode = .Columns("startMode").CellText(varBookMark)
              sUtilityType = .Columns("UtilityType").CellText(varBookMark)
              sUtilityID = .Columns("UtilityID").CellText(varBookMark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookMark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookMark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookMark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookMark)
              sAppParameters = .Columns("AppParameters").CellText(varBookMark)
              sDocumentFilePath = .Columns("DocumentFilePath").CellText(varBookMark)
              fDisplayDocumentHyperlink = .Columns("DisplayDocumentHyperlink").CellValue(varBookMark)
          End Select
    
          If Len(sScreenID) = 0 Then
            sScreenID = "0"
            sPageTitle = ""
            sStartMode = "0"
          End If
          
          If Len(sUtilityID) = 0 Then
            sUtilityID = "0"
            sUtilityType = "-1"
          End If
          
          If Len(sPictureID) = 0 Then sPictureID = "0"
          If Len(sSeparatorOrientation) = 0 Then sSeparatorOrientation = "0"
          
          sChartViewID = IIf(Len(sChartViewID) = 0, "0", sChartViewID)
          sChartTableID = IIf(Len(sChartTableID) = 0, "0", sChartTableID)
          sChartColumnID = IIf(Len(sChartColumnID) = 0, "0", sChartColumnID)
          sChartFilterID = IIf(Len(sChartFilterID) = 0, "0", sChartFilterID)
          
          sChart_TableID_2 = IIf(Len(sChart_TableID_2) = 0, "0", sChart_TableID_2)
          sChart_ColumnID_2 = IIf(Len(sChart_ColumnID_2) = 0, "0", sChart_ColumnID_2)
          sChart_TableID_3 = IIf(Len(sChart_TableID_3) = 0, "0", sChart_TableID_3)
          sChart_ColumnID_3 = IIf(Len(sChart_ColumnID_3) = 0, "0", sChart_ColumnID_3)
          sChart_SortOrderID = IIf(Len(sChart_SortOrderID) = 0, "0", sChart_SortOrderID)
          sChart_SortDirection = IIf(Len(sChart_SortDirection) = 0, "0", sChart_SortDirection)
          sChart_ColourID = IIf(Len(sChart_ColourID) = 0, "0", sChart_ColourID)
                   
          sFormatting_Prefix = IIf(sFormatting_Prefix = "''", "''", "'" & Replace(sFormatting_Prefix, "'", "''") & "'")
          sFormatting_Suffix = IIf(sFormatting_Suffix = "''", "''", "'" & Replace(sFormatting_Suffix, "'", "''") & "'")
      
          sConditionalFormatting_Operator_1 = IIf(sConditionalFormatting_Operator_1 = "''", "''", "'" & Replace(sConditionalFormatting_Operator_1, "'", "''") & "'")
          sConditionalFormatting_Value_1 = IIf(sConditionalFormatting_Value_1 = "''", "''", "'" & Replace(sConditionalFormatting_Value_1, "'", "''") & "'")
          sConditionalFormatting_Style_1 = IIf(sConditionalFormatting_Style_1 = "''", "''", "'" & Replace(sConditionalFormatting_Style_1, "'", "''") & "'")
          sConditionalFormatting_Colour_1 = IIf(sConditionalFormatting_Colour_1 = "''", "''", "'" & Replace(sConditionalFormatting_Colour_1, "'", "''") & "'")
          sConditionalFormatting_Operator_2 = IIf(sConditionalFormatting_Operator_2 = "''", "''", "'" & Replace(sConditionalFormatting_Operator_2, "'", "''") & "'")
          sConditionalFormatting_Value_2 = IIf(sConditionalFormatting_Value_2 = "''", "''", "'" & Replace(sConditionalFormatting_Value_2, "'", "''") & "'")
          sConditionalFormatting_Style_2 = IIf(sConditionalFormatting_Style_2 = "''", "''", "'" & Replace(sConditionalFormatting_Style_2, "'", "''") & "'")
          sConditionalFormatting_Colour_2 = IIf(sConditionalFormatting_Colour_2 = "''", "''", "'" & Replace(sConditionalFormatting_Colour_2, "'", "''") & "'")
          sConditionalFormatting_Operator_3 = IIf(sConditionalFormatting_Operator_3 = "''", "''", "'" & Replace(sConditionalFormatting_Operator_3, "'", "''") & "'")
          sConditionalFormatting_Value_3 = IIf(sConditionalFormatting_Value_3 = "''", "''", "'" & Replace(sConditionalFormatting_Value_3, "'", "''") & "'")
          sConditionalFormatting_Style_3 = IIf(sConditionalFormatting_Style_3 = "''", "''", "'" & Replace(sConditionalFormatting_Style_3, "'", "''") & "'")
          sConditionalFormatting_Colour_3 = IIf(sConditionalFormatting_Colour_3 = "''", "''", "'" & Replace(sConditionalFormatting_Colour_3, "'", "''") & "'")
          
          sSeparatorColour = IIf(sSeparatorColour = "''", "''", "'" & Replace(sSeparatorColour, "'", "''") & "'")
          
          iInitialDisplayMode = CStr(iInitialDisplayMode)
'          sChart_TableID_2 = .Columns("sChart_TableID_2").CellText(varBookmark)
'          sChart_ColumnID_2 = .Columns("sChart_ColumnID_2").CellText(varBookmark)
'          sChart_TableID_3 = .Columns("sChart_TableID_3").CellText(varBookmark)
'          sChart_ColumnID_3 = .Columns("sChart_ColumnID_3").CellText(varBookmark)
'          sChart_SortOrderID = .Columns("sChart_SortOrderID").CellText(varBookmark)
                        
          'NPG20080211 Fault 12873
           sSQL = "INSERT INTO tmpSSIntranetLinks" & _
            " ([linkType], [linkOrder], [prompt], [text], [screenID], [pageTitle], [url], [startMode], " & _
            "[utilityType], [utilityID], [viewID], [newWindow], [tableID], [EMailAddress], [EMailSubject], " & _
            "[AppFilePath], [AppParameters], [DocumentFilePath], [DisplayDocumentHyperlink], [Element_Type], " & _
            "[SeparatorOrientation], [PictureID], [Chart_ShowLegend], [Chart_Type], [Chart_ShowGrid], [Chart_StackSeries], " & _
            "[Chart_ViewID], [Chart_TableID], [Chart_ColumnID], [Chart_FilterID], [Chart_AggregateType], [Chart_ShowValues]," & _
            "[UseFormatting],[Formatting_DecimalPlaces],[Formatting_Use1000Separator],[Formatting_Prefix],[Formatting_Suffix]," & _
            "[UseConditionalFormatting],[ConditionalFormatting_Operator_1],[ConditionalFormatting_Value_1],[ConditionalFormatting_Style_1]," & _
            "[ConditionalFormatting_Colour_1],[ConditionalFormatting_Operator_2],[ConditionalFormatting_Value_2],[ConditionalFormatting_Style_2]," & _
            "[ConditionalFormatting_Colour_2],[ConditionalFormatting_Operator_3],[ConditionalFormatting_Value_3],[ConditionalFormatting_Style_3]," & _
            "[ConditionalFormatting_Colour_3],[SeparatorColour],[InitialDisplayMode],[Chart_TableID_2] ,[Chart_ColumnID_2],[Chart_TableID_3]," & _
            "[Chart_ColumnID_3],[Chart_SortOrderID],[Chart_SortDirection],[Chart_ColourID],[Chart_ShowPercentages])" & _
            " SELECT " & _
            CStr(piLinkType) & "," & CStr(iLoop) & "," & "'" & Replace(sPrompt, "'", "''") & "'," & _
            "'" & Replace(sText, "'", "''") & "'," & sScreenID & "," & "'" & Replace(sPageTitle, "'", "''") & "'," & _
            "'" & Replace(sURL, "'", "''") & "'," & sStartMode & "," & sUtilityType & "," & _
            sUtilityID & "," & sViewID & "," & sNewWindow & "," & sTableID & "," & _
            "'" & Replace(sEMailAddress, "'", "''") & "'," & "'" & Replace(sEMailSubject, "'", "''") & "'," & _
            "'" & Replace(sAppFilePath, "'", "''") & "'," & "'" & Replace(sAppParameters, "'", "''") & "'," & _
            "'" & Replace(sDocumentFilePath, "'", "''") & "',"

          sSQL = sSQL & _
            IIf(fDisplayDocumentHyperlink, "1", "0") & "," & _
            sElement_Type & "," & sSeparatorOrientation & "," & _
            sPictureID & "," & "" & IIf(fChartShowLegend, "1", "0") & "," & _
            sChartType & "," & "" & IIf(fChartShowGrid, "1", "0") & "," & _
            IIf(fChartStackSeries, "1", "0") & "," & _
            sChartViewID & "," & sChartTableID & "," & _
            sChartColumnID & "," & sChartFilterID & "," & _
            sChartAggregateType & "," & "" & IIf(fChartShowValues, "1", "0") & "," & _
            IIf(fUseFormatting, "1", "0") & "," & _
            iFormatting_DecimalPlaces & "," & _
            IIf(fFormatting_Use1000Separator, "1", "0") & "," & _
            sFormatting_Prefix & "," & _
            sFormatting_Suffix & "," & _
            IIf(fUseConditionalFormatting, "1", "0") & "," & _
            sConditionalFormatting_Operator_1 & "," & sConditionalFormatting_Value_1 & "," & sConditionalFormatting_Style_1 & "," & sConditionalFormatting_Colour_1 & "," & _
            sConditionalFormatting_Operator_2 & "," & sConditionalFormatting_Value_2 & "," & sConditionalFormatting_Style_2 & "," & sConditionalFormatting_Colour_2 & "," & _
            sConditionalFormatting_Operator_3 & "," & sConditionalFormatting_Value_3 & "," & sConditionalFormatting_Style_3 & "," & sConditionalFormatting_Colour_3 & "," & _
            sSeparatorColour & "," & CStr(iInitialDisplayMode) & "," & sChart_TableID_2 & "," & sChart_ColumnID_2 & "," & sChart_TableID_3 & "," & sChart_ColumnID_3 & "," & _
            sChart_SortOrderID & "," & sChart_SortDirection & "," & sChart_ColourID & "," & _
            IIf(fChartShowPercentages, "1", "0")

          daoDb.Execute sSQL, dbFailOnError
        
          ' Get the ID of the link just saved.
          sSQL = "SELECT MAX(id) AS [result] FROM tmpSSIntranetLinks"
          Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          lngMaxID = rsTemp!result
          rsTemp.Close
          Set rsTemp = Nothing
  
          sGroupNames = Mid(.Columns("HiddenGroups").CellText(varBookMark), 2)
          Do While InStr(sGroupNames, vbTab) > 0
            sGroupName = Left(sGroupNames, InStr(sGroupNames, vbTab) - 1)
            sGroupNames = Mid(sGroupNames, InStr(sGroupNames, vbTab) + 1)
  
            sSQL = "INSERT INTO tmpSSIHiddenGroups" & _
              " ([linkID], [groupName])" & _
              " VALUES (" & CStr(lngMaxID) & ", '" & Replace(sGroupName, "'", "''") & "')"
            daoDb.Execute sSQL, dbFailOnError
          Loop
      
        Next iLoop
      End If
    End With
  Next ctlGrid
  Set ctlGrid = Nothing
  Set ctlGridArray = Nothing

End Sub

Private Function SelectedTables() As String

  ' Return a string of the selected view IDs
  Dim iLoop As Integer
  Dim varBookMark As Variant
  Dim sSelectedTableIDs As String
  
  sSelectedTableIDs = "0"
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookMark = .AddItemBookmark(iLoop)

      sSelectedTableIDs = sSelectedTableIDs & "," & .Columns("TableID").CellText(varBookMark)
    Next iLoop
  End With

  SelectedTables = sSelectedTableIDs
  
End Function

Private Function SelectedViews() As String

  ' Return a string of the selected view IDs
  Dim iLoop As Integer
  Dim varBookMark As Variant
  Dim sSelectedViewIDs As String
  
  sSelectedViewIDs = "0"
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookMark = .AddItemBookmark(iLoop)

      sSelectedViewIDs = sSelectedViewIDs & "," & .Columns("ViewID").CellText(varBookMark)
    Next iLoop
  End With

  SelectedViews = sSelectedViewIDs
  
End Function

Private Function SingleRecordViewID() As Long

  ' Return the ID of the defined single record view.
  Dim iLoop As Integer
  Dim lngSingleRecordViewID As Long
  Dim varBookMark As Variant
  Dim fSingleRecord As Boolean
  Dim sViewID As String
  
  lngSingleRecordViewID = 0
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookMark = .AddItemBookmark(iLoop)

      fSingleRecord = .Columns("SingleRecord").CellValue(varBookMark)
      
      If fSingleRecord Then
        sViewID = .Columns("ViewID").CellText(varBookMark)
        lngSingleRecordViewID = CLng(sViewID)
        
        Exit For
      End If
    Next iLoop
  End With
  
  SingleRecordViewID = lngSingleRecordViewID

End Function

Private Sub cboButtonLinkView_Click()
  RefreshControls
End Sub

Private Sub cboDocumentView_Click()
  RefreshControls
End Sub

Private Sub cboDropdownListLinkView_Click()
  RefreshControls
End Sub

Private Sub cboHypertextLinkView_Click()
  RefreshControls
End Sub


Private Sub cboSecurityGroup_Click()
  Dim ctlSourceGrid As SSDBGrid
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  ctlSourceGrid.Refresh
  
  cmdPreview.Enabled = (cboSecurityGroup.Text <> "(All Groups)")
End Sub

Private Sub cmdAddButtonLink_Click()

  Dim sRow As String
  Dim frmLink As New frmSSIntranetLink
  Dim ctlCurrentGrid As SSDBGrid
  Dim ctlSourceGrid As SSDBGrid
  Dim iLoop As Integer
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  BuildUserGroupCollection
  
  PopulateWFAccessGroup ctlSourceGrid, -1
  
  With frmLink
    .InitialDisplayMode = 0
    .Chart_TableID_2 = 0
    .Chart_ColumnID_2 = 0
    .Chart_TableID_3 = 0
    .Chart_ColumnID_3 = 0
    .Chart_SortOrderID = 0
    .Chart_SortDirection = 0
    .Chart_ColourID = 0
    .Initialize SSINTLINK_BUTTON, _
      "", _
      "", _
      "", _
      "", _
      "", _
      DecodeTag(ctlSourceGrid.Tag, False), _
      "", _
      DecodeTag(ctlSourceGrid.Tag, True), _
      "", _
      "", _
      False, _
      "", _
      cboButtonLinkView.List(cboButtonLinkView.ListIndex), _
      False, _
      "", _
      "", _
      "", _
      "", _
      "", _
      False, False, _
      0, 0, _
      True, 1, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
      
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Prompt & vbTab & .Text & vbTab & .URL _
        & vbTab & .HRProScreenID & vbTab & .PageTitle & vbTab & .StartMode _
        & vbTab & .UtilityType & vbTab & .UtilityID & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress & vbTab & .EMailSubject & vbTab & .AppFilePath _
        & vbTab & .AppParameters & vbTab & .ElementType & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "") & vbTab & IIf(.chkShowLegend.value = 0, "0", "1") _
        & vbTab & .cboChartType.ItemData(.cboChartType.ListIndex) & vbTab & IIf(.chkDottedGridlines.value = 0, "0", "1") _
        & vbTab & IIf(.chkStackSeries.value = 0, "0", "1") & vbTab & "0" & vbTab & .ChartTableID & vbTab & .ChartColumnID _
        & vbTab & .ChartFilterID & vbTab & .ChartAggregateType & vbTab & IIf(.chkShowValues.value = 0, "0", "1") _
        & vbTab & .chkFormatting.value & vbTab & .spnDBValueDecimals.value & vbTab & .chkDBVaUseThousandSeparator.value _
        & vbTab & .txtDBValuePrefix.Text & vbTab & .txtDBValueSuffix.Text _
        & vbTab & .chkConditionalFormatting.value _
        & vbTab & .ConditionalFormatting_Operator_1 & vbTab & .ConditionalFormatting_Value_1 & vbTab & .ConditionalFormatting_Style_1 & vbTab & .ConditionalFormatting_Colour_1 _
        & vbTab & .ConditionalFormatting_Operator_2 & vbTab & .ConditionalFormatting_Value_2 & vbTab & .ConditionalFormatting_Style_2 & vbTab & .ConditionalFormatting_Colour_2 _
        & vbTab & .ConditionalFormatting_Operator_3 & vbTab & .ConditionalFormatting_Value_3 & vbTab & .ConditionalFormatting_Style_3 & vbTab & .ConditionalFormatting_Colour_3 _
        & vbTab & .SeparatorBorderColour & vbTab & .InitialDisplayMode & vbTab & .Chart_TableID_2 & vbTab & .Chart_ColumnID_2 & vbTab & .Chart_TableID_3 _
        & vbTab & .Chart_ColumnID_3 & vbTab & .Chart_SortOrderID & vbTab & .Chart_SortDirection & vbTab & .Chart_ColourID & vbTab & .ChartShowPercentages

      For iLoop = 0 To cboButtonLinkView.ListCount - 1
        If cboButtonLinkView.List(iLoop) = .TableViewName Then
          cboButtonLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
      
      If Not ctlCurrentGrid Is Nothing Then
        With ctlCurrentGrid
          .AddItem sRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
          
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End With

        Changed = True
      End If
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdAddDropdownListLink_Click()

  Dim sRow As String
  Dim frmLink As New frmSSIntranetLink
  Dim ctlCurrentGrid As SSDBGrid
  Dim ctlSourceGrid As SSDBGrid
  Dim iLoop As Integer

  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If

  With frmLink
    .Initialize SSINTLINK_DROPDOWNLIST, _
      "", _
      "", _
      "", _
      "", _
      "", _
      DecodeTag(ctlSourceGrid.Tag, False), _
      "", _
      DecodeTag(ctlSourceGrid.Tag, True), _
      "", _
      "", _
      False, _
      "", _
      cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex), _
      False, _
      "", _
      "", _
      "", _
      "", _
      "", _
      False, False, _
      0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
      
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .HRProScreenID _
        & vbTab & .PageTitle _
        & vbTab & .StartMode _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters
        
      For iLoop = 0 To cboDropdownListLinkView.ListCount - 1
        If cboDropdownListLinkView.List(iLoop) = .TableViewName Then
          cboDropdownListLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
      
      If Not ctlCurrentGrid Is Nothing Then
        With ctlCurrentGrid
          .AddItem sRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
          
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End With

        Changed = True
      End If
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdAddHyperTextLink_Click()

  Dim sRow As String
  Dim frmLink As New frmSSIntranetLink
  Dim ctlCurrentGrid As SSDBGrid
  Dim ctlSourceGrid As SSDBGrid
  Dim iLoop As Integer
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  With frmLink
    .Initialize SSINTLINK_HYPERTEXT, _
      "", _
      "", _
      "", _
      "", _
      "", _
      DecodeTag(ctlSourceGrid.Tag, False), _
      "", _
      DecodeTag(ctlSourceGrid.Tag, True), _
      "", _
      "", _
      False, _
      "", _
      cboHypertextLinkView.List(cboHypertextLinkView.ListIndex), _
      False, _
      "", _
      "", _
      "", _
      "", _
      "", _
      False, False, _
      0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
      
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters _
        & vbTab & .ElementType _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "")
        
      For iLoop = 0 To cboHypertextLinkView.ListCount - 1
        If cboHypertextLinkView.List(iLoop) = .TableViewName Then
          cboHypertextLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
      
      If Not ctlCurrentGrid Is Nothing Then
        With ctlCurrentGrid
          .AddItem sRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
          
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End With
        
        Changed = True
      End If
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property

Public Property Let Changed(ByVal pblnChanged As Boolean)
  
  mfChanged = pblnChanged
  RefreshControls
  
End Property

Private Sub cmdAddDocument_Click()

  Dim sRow As String
  Dim frmDocument As New frmSSIntranetLink
  Dim ctlCurrentGrid As SSDBGrid
  Dim ctlSourceGrid As SSDBGrid
  Dim iLoop As Integer

  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If

  With frmDocument
    .Initialize SSINTLINK_DOCUMENT, _
       "", _
      "", _
      "", _
      "", _
      "", _
      DecodeTag(ctlSourceGrid.Tag, False), _
      "", _
      DecodeTag(ctlSourceGrid.Tag, True), _
      "", _
      "", _
      False, _
      "", _
      cboDocumentView.List(cboDocumentView.ListIndex), _
      False, _
      "", _
      "", _
      "", _
      "", _
      "", _
      False, False, _
      0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
      
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .HRProScreenID _
        & vbTab & .PageTitle _
        & vbTab & .StartMode _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters _
        & vbTab & .DocumentFilePath _
        & vbTab & IIf(.DisplayDocumentHyperlink, "1", "0")

      For iLoop = 0 To cboDocumentView.ListCount - 1
        If cboDocumentView.List(iLoop) = .TableViewName Then
          cboDocumentView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
      
      If Not ctlCurrentGrid Is Nothing Then
        With ctlCurrentGrid
          .AddItem sRow
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .Columns("HiddenGroups").Text = frmDocument.HiddenGroups
          .Update
          
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End With

        Changed = True
      End If
    End If
  End With

  UnLoad frmDocument
  Set frmDocument = Nothing

End Sub

Private Sub cmdAddTableView_Click()
  
  Dim sRow As String
  Dim frmTableView As New frmSSIntranetView
  Dim iLoop As Integer
  Dim varBookMark As Variant

  With frmTableView
    .Initialize 0, _
      mlngPersonnelTableID, _
      SelectedViews, _
      SelectedTables, _
      False, _
      "", _
      "", _
      "", _
      "", _
      False, _
      False, _
      False, _
      "", _
      "", _
      mcolSSITableViews, _
      True
      
    If Not .Cancelled Then
      .Show vbModal
    End If
    
    If Not .Cancelled Then
      sRow = .TableViewName _
        & vbTab & CStr(.TableID) _
        & vbTab & CStr(.ViewID) _
        & vbTab & .ButtonLinkPromptText _
        & vbTab & .ButtonLinkButtonText _
        & vbTab & .HypertextLinkText _
        & vbTab & .DropdownListLinkText _
        & vbTab & IIf(.ButtonLink, "1", "0") _
        & vbTab & IIf(.HypertextLink, "1", "0") _
        & vbTab & IIf(.DropdownListLink, "1", "0") _
        & vbTab & .LinksLinkText _
        & vbTab & .PageTitle _
        & vbTab & .SingleRecordView _
        & vbTab & .WFOutOfOffice

      With grdTableViews
        .AddItem sRow
        
        If frmTableView.SingleRecordView Then
          ' Ensure only one 'single record view' exists.
          For iLoop = 0 To (.Rows - 1)
            varBookMark = .AddItemBookmark(iLoop)

            If .Columns("SingleRecord").CellValue(varBookMark) _
              And CLng(.Columns("ViewID").CellText(varBookMark)) <> frmTableView.ViewID Then
              ' Remove the other 'single record view' markers
              .Bookmark = varBookMark
              .Columns("SingleRecord").value = False
              .Update
            End If
          Next iLoop
        End If
        
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
        .Refresh
      End With

      AddTableViewGrids .TableID, .ViewID

      RefreshTableViewCombos
      
      RefreshTableViewsCollection
      
      Changed = True
    End If
  End With

  UnLoad frmTableView
  Set frmTableView = Nothing

End Sub

'NPG Dashboard
'Private Sub cmdButtonLinkSeparator_Click()
'
'  Dim sRow As String
'  Dim lngRow As Long
'  Dim frmLinkSeparator As New frmSSIntranetLinkSeparator
'  Dim ctlSourceGrid As SSDBGrid
'  Dim ctlDestinationGrid As SSDBGrid
'  Dim iLoop As Integer
'  Dim sViews As String
'  Dim fNew As Boolean
'
'  sViews = SelectedViews
'
'  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
'  If ctlSourceGrid Is Nothing Then
'    Exit Sub
'  End If
'
'  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
'  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)
'
'  fNew = (Not mfIsCopyingSeparator) And (Not mfIsEditingSeparator)
'
'  With frmLinkSeparator
'    .Initialize SSINTLINK_BUTTON, _
'      IIf(fNew, "", ctlSourceGrid.Columns("Prompt").Text), _
'      DecodeTag(ctlSourceGrid.Tag, False), _
'      DecodeTag(ctlSourceGrid.Tag, True), _
'      mfIsCopyingSeparator, _
'      IIf(fNew, 0, Val(ctlSourceGrid.Columns("SeparatorOrientation").Text)), _
'      IIf(fNew, 0, Val(ctlSourceGrid.Columns("PictureID").Text))
'    .Show vbModal
'
'    If Not .Cancelled Then
'      sRow = .Text _
'        & vbTab & "<SEPARATOR>" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & vbTab & "0" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "1" _
'        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
'        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "")
'
'
'      With ctlSourceGrid
'        If mfIsEditingSeparator Then
'          ctlSourceGrid.RemoveItem lngRow
'
'          .AddItem sRow, lngRow
'          .Bookmark = .AddItemBookmark(lngRow)
'          .Update
'          .SelBookmarks.RemoveAll
'          .SelBookmarks.Add .AddItemBookmark(lngRow)
'
'        ElseIf mfIsCopyingSeparator Then
'          .AddItem sRow, lngRow + 1
'          .Bookmark = .AddItemBookmark(lngRow + 1)
'          .Update
'          .SelBookmarks.RemoveAll
'          .SelBookmarks.Add .AddItemBookmark(lngRow + 1)
'
'        Else
'          ' Adding a new separator.
'          .AddItem sRow
'          .Bookmark = .AddItemBookmark(.Rows - 1)
'          .Update
'          .SelBookmarks.RemoveAll
'          .SelBookmarks.Add .Bookmark
'        End If
'      End With
'
'      mfIsEditingSeparator = False
'      mfIsCopyingSeparator = False
'      Changed = True
'    End If
'  End With
'
'  UnLoad frmLinkSeparator
'  Set frmLinkSeparator = Nothing
'
'End Sub

Private Sub cmdCancel_Click()
  'AE20071119 Fault #12607
'  Dim pintAnswer As Integer
'    If Changed = True And cmdOK.Enabled Then
'      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
'      If pintAnswer = vbYes Then
'        'AE20071108 Fault #12551
'        'Using Me.MousePointer = vbNormal forces the form to be reloaded
'        'after its been unloaded in cmdOK_Click, changed to Screen.MousePointer
'        'Me.MousePointer = vbHourglass
'        Screen.MousePointer = vbHourglass
'        cmdOK_Click 'This is just like saving
'        Screen.MousePointer = vbNormal
'        'Me.MousePointer = vbNormal
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:
  UnLoad Me
End Sub

Private Sub cmdCopyButtonLink_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmLink As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long

  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If

  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboButtonLinkView.List(cboButtonLinkView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboButtonLinkView.List(cboButtonLinkView.ListIndex))
  
  BuildUserGroupCollection
  PopulateWFAccessGroup ctlSourceGrid, -1
  
  With frmLink
    .InitialDisplayMode = ctlSourceGrid.Columns("InitialDisplayMode").Text
    .Chart_TableID_2 = ctlSourceGrid.Columns("Chart_TableID_2").Text
    .Chart_ColumnID_2 = ctlSourceGrid.Columns("Chart_ColumnID_2").Text
    .Chart_TableID_3 = ctlSourceGrid.Columns("Chart_TableID_3").Text
    .Chart_ColumnID_3 = ctlSourceGrid.Columns("Chart_ColumnID_3").Text
    .Chart_SortOrderID = ctlSourceGrid.Columns("Chart_SortOrderID").Text
    .Chart_SortDirection = ctlSourceGrid.Columns("Chart_SortDirection").Text
    .Chart_ColourID = ctlSourceGrid.Columns("Chart_ColourID").Text
    ' .ChartShowPercentages = ctlSourceGrid.Columns("ChartShowPercentages").Text
    
    .Initialize SSINTLINK_BUTTON, _
      ctlSourceGrid.Columns("Prompt").Text, ctlSourceGrid.Columns("ButtonText").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, ctlSourceGrid.Columns("UtilityID").Text, True, _
      ctlSourceGrid.Columns("HiddenGroups").Text, cboButtonLinkView.List(cboButtonLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, ctlSourceGrid.Columns("EMailAddress").Text, ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, ctlSourceGrid.Columns("AppParameters").Text, "", False, _
      ctlSourceGrid.Columns("Element_Type").value, val(ctlSourceGrid.Columns("SeparatorOrientation").Text), val(ctlSourceGrid.Columns("PictureID").Text), _
      ctlSourceGrid.Columns("ChartShowLegend").value, val(ctlSourceGrid.Columns("ChartType").Text), _
      ctlSourceGrid.Columns("ChartShowGrid").value, ctlSourceGrid.Columns("ChartStackSeries").value, _
      val(ctlSourceGrid.Columns("ChartViewID").Text), val(ctlSourceGrid.Columns("ChartTableID").Text), _
      val(ctlSourceGrid.Columns("ChartColumnID").Text), val(ctlSourceGrid.Columns("ChartFilterID").Text), _
      val(ctlSourceGrid.Columns("ChartAggregateType").Text), ctlSourceGrid.Columns("ChartShowValues").value, mcolGroups, _
      ctlSourceGrid.Columns("UseFormatting").Text, _
      ctlSourceGrid.Columns("Formatting_DecimalPlaces").Text, ctlSourceGrid.Columns("Formatting_Use1000Separator").Text, _
      ctlSourceGrid.Columns("Formatting_Prefix").Text, ctlSourceGrid.Columns("Formatting_Suffix").Text, _
      ctlSourceGrid.Columns("UseConditionalFormatting").Text, _
      ctlSourceGrid.Columns("ConditionalFormatting_Operator_1").Text, ctlSourceGrid.Columns("ConditionalFormatting_Value_1").Text, ctlSourceGrid.Columns("ConditionalFormatting_Style_1").Text, ctlSourceGrid.Columns("ConditionalFormatting_Colour_1").Text, _
      ctlSourceGrid.Columns("ConditionalFormatting_Operator_2").Text, ctlSourceGrid.Columns("ConditionalFormatting_Value_2").Text, ctlSourceGrid.Columns("ConditionalFormatting_Style_2").Text, ctlSourceGrid.Columns("ConditionalFormatting_Colour_2").Text, _
      ctlSourceGrid.Columns("ConditionalFormatting_Operator_3").Text, ctlSourceGrid.Columns("ConditionalFormatting_Value_3").Text, ctlSourceGrid.Columns("ConditionalFormatting_Style_3").Text, ctlSourceGrid.Columns("ConditionalFormatting_Colour_3").Text, _
      ctlSourceGrid.Columns("SeparatorColour").Text, ctlSourceGrid.Columns("ChartShowPercentages").Text, mcolSSITableViews
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Prompt _
        & vbTab & .Text & vbTab & .URL & vbTab & .HRProScreenID _
        & vbTab & .PageTitle & vbTab & .StartMode & vbTab & .UtilityType & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") & vbTab & .EMailAddress & vbTab & .EMailSubject _
        & vbTab & .AppFilePath & vbTab & .AppParameters & vbTab & IIf(.optLink(SSINTLINKSEPARATOR).value, 1, IIf(.optLink(SSINTLINKCHART).value, 2, IIf(.optLink(SSINTLINKPWFSTEPS).value, 3, IIf(.optLink(SSINTLINKDB_VALUE).value, 4, IIf(.optLink(SSINTLINKTODAYS_EVENTS).value, 5, 0))))) _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "") _
        & vbTab & IIf(.chkShowLegend.value = 0, "0", "1") & vbTab & .cboChartType.ItemData(.cboChartType.ListIndex) _
        & vbTab & IIf(.chkDottedGridlines.value = 0, "0", "1") & vbTab & IIf(.chkStackSeries.value = 0, "0", "1") _
        & vbTab & .ChartViewID & vbTab & .ChartTableID & vbTab & .ChartColumnID & vbTab & .ChartFilterID _
        & vbTab & .ChartAggregateType & vbTab & IIf(.chkShowValues.value = 0, "0", "1") _
        & vbTab & IIf(.UseFormatting = 0, "0", "1") & vbTab & .spnDBValueDecimals.value & vbTab & .chkDBVaUseThousandSeparator.value _
        & vbTab & .txtDBValuePrefix.Text & vbTab & .txtDBValueSuffix.Text _
        & vbTab & .chkConditionalFormatting.value _
        & vbTab & .ConditionalFormatting_Operator_1 & vbTab & .ConditionalFormatting_Value_1 & vbTab & .ConditionalFormatting_Style_1 & vbTab & .ConditionalFormatting_Colour_1 _
        & vbTab & .ConditionalFormatting_Operator_2 & vbTab & .ConditionalFormatting_Value_2 & vbTab & .ConditionalFormatting_Style_2 & vbTab & .ConditionalFormatting_Colour_2 _
        & vbTab & .ConditionalFormatting_Operator_3 & vbTab & .ConditionalFormatting_Value_3 & vbTab & .ConditionalFormatting_Style_3 & vbTab & .ConditionalFormatting_Colour_3 _
        & vbTab & .SeparatorBorderColour & vbTab & .InitialDisplayMode & vbTab & .Chart_TableID_2 & vbTab & .Chart_ColumnID_2 & vbTab & .Chart_TableID_3 _
        & vbTab & .Chart_ColumnID_3 & vbTab & .Chart_SortOrderID & vbTab & .Chart_SortDirection & vbTab & .Chart_ColourID & vbTab & .ChartShowPercentages

      For iLoop = 0 To cboButtonLinkView.ListCount - 1
        If cboButtonLinkView.List(iLoop) = .TableViewName Then
          cboButtonLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow + 1
          .Bookmark = .AddItemBookmark(lngRow + 1)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow + 1)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmLink.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
          
            .MoveLast
          End With
        End If
      End If
            
      Changed = True
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdCopyDropdownListLink_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmLink As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex))
  
  With frmLink
    .Initialize SSINTLINK_DROPDOWNLIST, _
      "", _
      ctlSourceGrid.Columns("Text").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, _
      ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, _
      True, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, _
      ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, _
      ctlSourceGrid.Columns("AppParameters").Text, _
      "", _
      False, False, _
      0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
    
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .HRProScreenID _
        & vbTab & .PageTitle _
        & vbTab & .StartMode _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters
        
      For iLoop = 0 To cboDropdownListLinkView.ListCount - 1
        If cboDropdownListLinkView.List(iLoop) = .TableViewName Then
          cboDropdownListLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow + 1
          .Bookmark = .AddItemBookmark(lngRow + 1)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow + 1)
        End With
        
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmLink.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
          
            .MoveLast
          End With
        End If
        
      End If
            
      Changed = True
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdCopyHypertextLink_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmLink As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboHypertextLinkView.List(cboHypertextLinkView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboHypertextLinkView.List(cboHypertextLinkView.ListIndex))
  
  With frmLink
    .Initialize SSINTLINK_HYPERTEXT, _
      "", _
      ctlSourceGrid.Columns("Text").Text, _
      "", _
      "", _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      "", _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, _
      True, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboHypertextLinkView.List(cboHypertextLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, _
      ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, _
      ctlSourceGrid.Columns("AppParameters").Text, _
      "", False, _
      ctlSourceGrid.Columns("Element_Type").value, val(ctlSourceGrid.Columns("SeparatorOrientation").Text), _
      val(ctlSourceGrid.Columns("PictureID").Text), _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
    
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters _
        & vbTab & IIf(.optLink(SSINTLINKSEPARATOR).value = True, 1, 0) _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "")
        
        
      For iLoop = 0 To cboHypertextLinkView.ListCount - 1
        If cboHypertextLinkView.List(iLoop) = .TableViewName Then
          cboHypertextLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow + 1
          .Bookmark = .AddItemBookmark(lngRow + 1)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow + 1)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmLink.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
          
            .MoveLast
          End With
        End If
      End If

      Changed = True
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdCopyDocument_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmDocument As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
 
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboDocumentView.List(cboDocumentView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboDocumentView.List(cboDocumentView.ListIndex))
  
  With frmDocument
    .Initialize SSINTLINK_DOCUMENT, _
      "", _
      ctlSourceGrid.Columns("Text").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, _
      ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, _
      True, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboButtonLinkView.List(cboDocumentView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, _
      ctlSourceGrid.Columns("EMailSubject").Text, _
      "", _
      "", _
      ctlSourceGrid.Columns("DocumentFilePath").Text, _
      ctlSourceGrid.Columns("DisplayDocumentHyperlink").value, _
      False, 0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
      
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .HRProScreenID _
        & vbTab & .PageTitle _
        & vbTab & .StartMode _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters _
        & vbTab & .DocumentFilePath _
        & vbTab & IIf(.DisplayDocumentHyperlink, "1", "0")

      For iLoop = 0 To cboDocumentView.ListCount - 1
        If cboDocumentView.List(iLoop) = .TableViewName Then
          cboDocumentView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow + 1
          .Bookmark = .AddItemBookmark(lngRow + 1)
          .Columns("HiddenGroups").Text = frmDocument.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow + 1)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmDocument.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
          
            .MoveLast
          End With
        End If
      End If
            
      Changed = True
    End If
  End With

  UnLoad frmDocument
  Set frmDocument = Nothing

End Sub

Private Sub cmdEditButtonLink_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmLink As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim sViews As String
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
      
  sViews = SelectedViews
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If

  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboButtonLinkView.List(cboButtonLinkView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboButtonLinkView.List(cboButtonLinkView.ListIndex))
  
  ' If pending workflow steps get the visibility details for all other wf steps...
  If ctlSourceGrid.Columns("Element_Type").value = 3 Or ctlSourceGrid.Columns("Element_Type").value = 5 Then
    BuildUserGroupCollection
    PopulateWFAccessGroup ctlSourceGrid, lngRow
  End If
  
  With frmLink
    .InitialDisplayMode = IIf(ctlSourceGrid.Columns("InitialDisplayMode").Text = vbNullString, 0, val(ctlSourceGrid.Columns("InitialDisplayMode").Text))
    .Chart_TableID_2 = IIf(ctlSourceGrid.Columns("Chart_TableID_2").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_TableID_2").Text))
    .Chart_ColumnID_2 = IIf(ctlSourceGrid.Columns("Chart_ColumnID_2").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_ColumnID_2").Text))
    .Chart_TableID_3 = IIf(ctlSourceGrid.Columns("Chart_TableID_3").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_TableID_3").Text))
    .Chart_ColumnID_3 = IIf(ctlSourceGrid.Columns("Chart_ColumnID_3").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_ColumnID_3").Text))
    .Chart_SortOrderID = IIf(ctlSourceGrid.Columns("Chart_SortOrderID").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_SortOrderID").Text))
    .Chart_SortDirection = IIf(ctlSourceGrid.Columns("Chart_SortDirection").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_SortDirection").Text))
    .Chart_ColourID = IIf(ctlSourceGrid.Columns("Chart_ColourID").Text = vbNullString, 0, val(ctlSourceGrid.Columns("Chart_ColourID").Text))
    ' .ChartShowPercentages = IIf(ctlSourceGrid.Columns("ChartShowPercentages").Text = vbNullString, 0, val(ctlSourceGrid.Columns("ChartShowPercentages").Text))
    
    .Initialize SSINTLINK_BUTTON, _
      ctlSourceGrid.Columns("Prompt").Text, ctlSourceGrid.Columns("ButtonText").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, ctlSourceGrid.Columns("UtilityID").Text, False, _
      ctlSourceGrid.Columns("HiddenGroups").Text, cboButtonLinkView.List(cboButtonLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, ctlSourceGrid.Columns("EMailAddress").Text, ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, ctlSourceGrid.Columns("AppParameters").Text, _
      "", False, ctlSourceGrid.Columns("Element_Type").value, val(ctlSourceGrid.Columns("SeparatorOrientation").Text), val(ctlSourceGrid.Columns("PictureID").Text), _
      ctlSourceGrid.Columns("ChartShowLegend").Text, val(ctlSourceGrid.Columns("ChartType").Text), ctlSourceGrid.Columns("ChartShowGrid").Text, _
      ctlSourceGrid.Columns("ChartStackSeries").Text, val(ctlSourceGrid.Columns("ChartviewID").Text), val(ctlSourceGrid.Columns("ChartTableID").Text), _
      val(ctlSourceGrid.Columns("ChartColumnID").Text), val(ctlSourceGrid.Columns("ChartFilterID").Text), val(ctlSourceGrid.Columns("ChartAggregateType").Text), _
      ctlSourceGrid.Columns("ChartShowValues").Text, mcolGroups, _
      ctlSourceGrid.Columns("UseFormatting").Text, _
      ctlSourceGrid.Columns("Formatting_DecimalPlaces").Text, ctlSourceGrid.Columns("Formatting_Use1000Separator").Text, _
      ctlSourceGrid.Columns("Formatting_Prefix").Text, ctlSourceGrid.Columns("Formatting_Suffix").Text, _
      ctlSourceGrid.Columns("UseConditionalFormatting").Text, _
      ctlSourceGrid.Columns("ConditionalFormatting_Operator_1").Text, ctlSourceGrid.Columns("ConditionalFormatting_Value_1").Text, ctlSourceGrid.Columns("ConditionalFormatting_Style_1").Text, ctlSourceGrid.Columns("ConditionalFormatting_Colour_1").Text, _
      ctlSourceGrid.Columns("ConditionalFormatting_Operator_2").Text, ctlSourceGrid.Columns("ConditionalFormatting_Value_2").Text, ctlSourceGrid.Columns("ConditionalFormatting_Style_2").Text, ctlSourceGrid.Columns("ConditionalFormatting_Colour_2").Text, _
      ctlSourceGrid.Columns("ConditionalFormatting_Operator_3").Text, ctlSourceGrid.Columns("ConditionalFormatting_Value_3").Text, ctlSourceGrid.Columns("ConditionalFormatting_Style_3").Text, ctlSourceGrid.Columns("ConditionalFormatting_Colour_3").Text, _
      ctlSourceGrid.Columns("SeparatorColour").Text, ctlSourceGrid.Columns("ChartShowPercentages").Text, mcolSSITableViews
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Prompt _
        & vbTab & .Text & vbTab & .URL & vbTab & .HRProScreenID _
        & vbTab & .PageTitle & vbTab & .StartMode & vbTab & .UtilityType _
        & vbTab & .UtilityID & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress & vbTab & .EMailSubject & vbTab & .AppFilePath _
        & vbTab & .AppParameters & vbTab & .ElementType & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "") & vbTab & IIf(.chkShowLegend = 0, "0", "1") _
        & vbTab & .cboChartType.ItemData(.cboChartType.ListIndex) _
        & vbTab & IIf(.chkDottedGridlines = 0, "0", "1") & vbTab & IIf(.chkStackSeries = 0, "0", "1") _
        & vbTab & 0 & vbTab & .ChartTableID & vbTab & .ChartColumnID & vbTab & .ChartFilterID _
        & vbTab & .ChartAggregateType & vbTab & IIf(.chkShowValues = 0, "0", "1") _
        & vbTab & IIf(.UseFormatting = 0, "0", "1") & vbTab & .Formatting_DecimalPlaces & vbTab & IIf(.Formatting_Use1000Separator = 0, "0", "1") _
        & vbTab & .Formatting_Prefix & vbTab & .Formatting_Suffix & vbTab & IIf(.UseConditionalFormatting = 0, "0", "1") _
        & vbTab & .ConditionalFormatting_Operator_1 & vbTab & .ConditionalFormatting_Value_1 & vbTab & .ConditionalFormatting_Style_1 & vbTab & .ConditionalFormatting_Colour_1 _
        & vbTab & .ConditionalFormatting_Operator_2 & vbTab & .ConditionalFormatting_Value_2 & vbTab & .ConditionalFormatting_Style_2 & vbTab & .ConditionalFormatting_Colour_2 _
        & vbTab & .ConditionalFormatting_Operator_3 & vbTab & .ConditionalFormatting_Value_3 & vbTab & .ConditionalFormatting_Style_3 & vbTab & .ConditionalFormatting_Colour_3 _
        & vbTab & .SeparatorBorderColour & vbTab & .InitialDisplayMode & vbTab & .Chart_TableID_2 & vbTab & .Chart_ColumnID_2 & vbTab & .Chart_TableID_3 _
        & vbTab & .Chart_ColumnID_3 & vbTab & .Chart_SortOrderID & vbTab & .Chart_SortDirection & vbTab & .Chart_ColourID & vbTab & IIf(.chkShowPercentages = 0, "0", "1")

      ctlSourceGrid.RemoveItem lngRow
      
      For iLoop = 0 To cboButtonLinkView.ListCount - 1
        If cboButtonLinkView.List(iLoop) = .TableViewName Then
          cboButtonLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow
          .Bookmark = .AddItemBookmark(lngRow)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmLink.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
  
            .MoveLast
          End With
        End If
      End If

      Changed = True
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdEditDropdownListLink_Click()
  
  Dim sRow As String
  Dim lngRow As Long
  Dim frmLink As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim sViews As String
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long

  sViews = SelectedViews
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If

  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex))
  
  With frmLink
    .Initialize SSINTLINK_DROPDOWNLIST, _
      "", _
      ctlSourceGrid.Columns("Text").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, _
      ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, _
      False, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboDropdownListLinkView.List(cboDropdownListLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, _
      ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, _
      ctlSourceGrid.Columns("AppParameters").Text, _
      "", _
      False, False, _
      0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
    
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .HRProScreenID _
        & vbTab & .PageTitle _
        & vbTab & .StartMode _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters

      ctlSourceGrid.RemoveItem lngRow
      
      For iLoop = 0 To cboDropdownListLinkView.ListCount - 1
        If cboDropdownListLinkView.List(iLoop) = .TableViewName Then
          cboDropdownListLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow
          .Bookmark = .AddItemBookmark(lngRow)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmLink.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
          
            .MoveLast
          End With
        End If
      End If
      
      Changed = True
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdEditHypertextLink_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmLink As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim sViews As String
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
  
  sViews = SelectedViews
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If

  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboHypertextLinkView.List(cboHypertextLinkView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboHypertextLinkView.List(cboHypertextLinkView.ListIndex))
  
  With frmLink
    .Initialize SSINTLINK_HYPERTEXT, _
      "", _
      ctlSourceGrid.Columns("Text").Text, _
      "", _
      "", _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      "", _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, _
      False, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboHypertextLinkView.List(cboHypertextLinkView.ListIndex), _
      (ctlSourceGrid.Columns("NewWindow").Text = "1"), _
      ctlSourceGrid.Columns("EMailAddress").Text, _
      ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, _
      ctlSourceGrid.Columns("AppParameters").Text, _
      "", False, _
      ctlSourceGrid.Columns("Element_Type").value, val(ctlSourceGrid.Columns("SeparatorOrientation").Text), _
      val(ctlSourceGrid.Columns("PictureID").Text), _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
    
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters _
        & vbTab & IIf(.optLink(SSINTLINKSEPARATOR).value = True, 1, 0) _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "")
        
        
      ctlSourceGrid.RemoveItem lngRow
      
      For iLoop = 0 To cboHypertextLinkView.ListCount - 1
        If cboHypertextLinkView.List(iLoop) = .TableViewName Then
          cboHypertextLinkView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow
          .Bookmark = .AddItemBookmark(lngRow)
          .Columns("HiddenGroups").Text = frmLink.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmLink.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
          
            .MoveLast
          End With
        End If
      End If

      Changed = True
    End If
  End With

  UnLoad frmLink
  Set frmLink = Nothing

End Sub

Private Sub cmdEditDocument_Click()

  Dim sRow As String
  Dim lngRow As Long
  Dim frmDocument As New frmSSIntranetLink
  Dim ctlSourceGrid As SSDBGrid
  Dim ctlDestinationGrid As SSDBGrid
  Dim iLoop As Integer
  Dim sViews As String
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
  
  sViews = SelectedViews
  
  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  If ctlSourceGrid Is Nothing Then
    Exit Sub
  End If
  
  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)

  lngOriginalTableID = GetTableIDFromCollection(mcolSSITableViews, cboDocumentView.List(cboDocumentView.ListIndex))
  lngOriginalViewID = GetViewIDFromCollection(mcolSSITableViews, cboDocumentView.List(cboDocumentView.ListIndex))
  
  With frmDocument
    .Initialize SSINTLINK_DOCUMENT, _
      "", _
      ctlSourceGrid.Columns("Text").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, _
      ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, _
      False, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboButtonLinkView.List(cboDocumentView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, _
      ctlSourceGrid.Columns("EMailSubject").Text, _
      "", _
      "", _
      ctlSourceGrid.Columns("DocumentFilePath").Text, _
      ctlSourceGrid.Columns("DisplayDocumentHyperlink").Text, _
      False, 0, 0, _
      False, 0, False, False, 0, 0, 0, 0, 0, 0, mcolGroups, _
      False, 0, False, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", False, _
      mcolSSITableViews
          
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Text _
        & vbTab & .URL _
        & vbTab & .HRProScreenID _
        & vbTab & .PageTitle _
        & vbTab & .StartMode _
        & vbTab & .UtilityType _
        & vbTab & .UtilityID _
        & vbTab & vbTab & IIf(.NewWindow, "1", "0") _
        & vbTab & .EMailAddress _
        & vbTab & .EMailSubject _
        & vbTab & .AppFilePath _
        & vbTab & .AppParameters _
        & vbTab & .DocumentFilePath _
        & vbTab & IIf(.DisplayDocumentHyperlink, "1", "0")

      ctlSourceGrid.RemoveItem lngRow
      
      For iLoop = 0 To cboDocumentView.ListCount - 1
        If cboDocumentView.List(iLoop) = .TableViewName Then
          cboDocumentView.ListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      If (lngOriginalTableID = .TableID) And _
        (lngOriginalViewID = .ViewID) Then
        ' Table and View(!) has NOT changed.
        With ctlSourceGrid
          .AddItem sRow, lngRow
          .Bookmark = .AddItemBookmark(lngRow)
          .Columns("HiddenGroups").Text = frmDocument.HiddenGroups
          .Update
      
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .AddItemBookmark(lngRow)
        End With
      Else
        ' Table or View(!) has changed.
        With ctlSourceGrid
          .SelBookmarks.RemoveAll
          If .Rows > 0 Then
            .SelBookmarks.Add .AddItemBookmark(0)
          End If
        End With
      
        Set ctlDestinationGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
        
        If Not ctlDestinationGrid Is Nothing Then
          With ctlDestinationGrid
            .AddItem sRow
            .Bookmark = .AddItemBookmark(.Rows - 1)
            .Columns("HiddenGroups").Text = frmDocument.HiddenGroups
            .Update
        
            .SelBookmarks.RemoveAll
            .SelBookmarks.Add .AddItemBookmark(.Rows - 1)
  
            .MoveLast
          End With
        End If
      End If

      Changed = True
    End If
  End With

  UnLoad frmDocument
  Set frmDocument = Nothing

End Sub

Private Sub cmdEditTableView_Click()
  Dim sRow As String
  Dim lngRow As Long
  Dim frmTableView As New frmSSIntranetView
  Dim iLoop As Integer
  Dim varBookMark As Variant
  Dim lngOriginalTableID As Long
  Dim lngOriginalViewID As Long
  Dim ctlGrid As SSDBGrid
  Dim sTableID As String
  Dim sViewID As String
  Dim sTag As String
  Dim sOriginalTag As String
  
  grdTableViews.Bookmark = grdTableViews.SelBookmarks(0)
  lngRow = grdTableViews.AddItemRowIndex(grdTableViews.Bookmark)

  lngOriginalTableID = CLng(grdTableViews.Columns("TableID").Text)
  lngOriginalViewID = CLng(grdTableViews.Columns("ViewID").Text)
  
  With frmTableView
    .Initialize CLng(grdTableViews.Columns("ViewID").Text), _
      CLng(grdTableViews.Columns("TableID").Text), _
      SelectedViews, _
      SelectedTables, _
      grdTableViews.Columns("SingleRecord").value, _
      grdTableViews.Columns("ButtonLinkPromptText").Text, _
      grdTableViews.Columns("ButtonLinkButtonText").Text, _
      grdTableViews.Columns("HypertextLinkText").Text, _
      grdTableViews.Columns("DropdownListLinkText").Text, _
      (grdTableViews.Columns("ButtonLink").Text = "1"), _
      (grdTableViews.Columns("HypertextLink").Text = "1"), _
      (grdTableViews.Columns("DropdownListLink").Text = "1"), _
      grdTableViews.Columns("LinksLinkText").Text, _
      grdTableViews.Columns("PageTitle").Text, _
      mcolSSITableViews, _
      grdTableViews.Columns("WFOutOfOffice").value

    .Show vbModal

    If Not .Cancelled Then
      sRow = .TableViewName _
        & vbTab & CStr(.TableID) _
        & vbTab & CStr(.ViewID) _
        & vbTab & .ButtonLinkPromptText _
        & vbTab & .ButtonLinkButtonText _
        & vbTab & .HypertextLinkText _
        & vbTab & .DropdownListLinkText _
        & vbTab & IIf(.ButtonLink, "1", "0") _
        & vbTab & IIf(.HypertextLink, "1", "0") _
        & vbTab & IIf(.DropdownListLink, "1", "0") _
        & vbTab & .LinksLinkText _
        & vbTab & .PageTitle _
        & vbTab & .SingleRecordView _
        & vbTab & .WFOutOfOffice

      grdTableViews.RemoveItem lngRow

      With grdTableViews
        .AddItem sRow, lngRow
        
        If frmTableView.SingleRecordView Then
          ' Ensure only one 'single record view' exists.
          For iLoop = 0 To (.Rows - 1)
            varBookMark = .AddItemBookmark(iLoop)

            If .Columns("SingleRecord").CellValue(varBookMark) _
              And CLng(.Columns("ViewID").CellText(varBookMark)) <> frmTableView.ViewID Then
              ' Remove the other 'single record view' markers
              .Bookmark = varBookMark
              .Columns("SingleRecord").value = False
              .Update
            End If
          Next iLoop
        End If
      
        .Bookmark = .AddItemBookmark(lngRow)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .AddItemBookmark(lngRow)
      End With

      If lngOriginalViewID <> .ViewID Then
          
          'NPG20080411 Fault 13060
          sTableID = grdTableViews.Columns("tableID").Text
          sViewID = grdTableViews.Columns("viewID").Text
          sTag = CreateTableViewTag(sTableID, sViewID)
          sOriginalTag = CreateTableViewTag(sTableID, CStr(lngOriginalViewID))
           
        For Each ctlGrid In grdHypertextLinks
          'NPG20080411 Fault 13060
          ' If ctlGrid.Tag = CStr(lngOriginalViewID) Then
            ' ctlGrid.Tag = CStr(.ViewID)
          If ctlGrid.Tag = sOriginalTag Then
            ctlGrid.Tag = sTag
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
        
        For Each ctlGrid In grdButtonLinks
          'NPG20080411 Fault 13060
          ' If ctlGrid.Tag = CStr(lngOriginalViewID) Then
            ' ctlGrid.Tag = CStr(.ViewID)
          If ctlGrid.Tag = sOriginalTag Then
            ctlGrid.Tag = sTag
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
        
        For Each ctlGrid In grdDropdownListLinks
         'NPG20080411 Fault 13060
         ' If ctlGrid.Tag = CStr(lngOriginalViewID) Then
            ' ctlGrid.Tag = CStr(.ViewID)
          If ctlGrid.Tag = sOriginalTag Then
            ctlGrid.Tag = sTag
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
      
        For Each ctlGrid In grdDocuments
          If ctlGrid.Tag = sOriginalTag Then
            ctlGrid.Tag = sTag
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
      
      End If
      
      RefreshTableViewCombos
      
      RefreshTableViewsCollection
      
      Changed = True
    End If
  End With

  UnLoad frmTableView
  Set frmTableView = Nothing

End Sub

'NPG Dashboard
'Private Sub cmdHypertextLinkSeparator_Click()
'
'  Dim sRow As String
'  Dim lngRow As Long
'  Dim frmLinkSeparator As New frmSSIntranetLinkSeparator
'  Dim ctlSourceGrid As SSDBGrid
'  Dim iLoop As Integer
'  Dim sViews As String
'  Dim fNew As Boolean
'
'  sViews = SelectedViews
'
'  Set ctlSourceGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
'  If ctlSourceGrid Is Nothing Then
'    Exit Sub
'  End If
'
'  ctlSourceGrid.Bookmark = ctlSourceGrid.SelBookmarks(0)
'  lngRow = ctlSourceGrid.AddItemRowIndex(ctlSourceGrid.Bookmark)
'
'  fNew = (Not mfIsCopyingSeparator) And (Not mfIsEditingSeparator)
'
'  With frmLinkSeparator
'    .Initialize SSINTLINK_HYPERTEXT, _
'      IIf(fNew, "", ctlSourceGrid.Columns("Text").CellValue(ctlSourceGrid.Bookmark)), _
'      DecodeTag(ctlSourceGrid.Tag, False), _
'      DecodeTag(ctlSourceGrid.Tag, True), _
'      mfIsCopyingSeparator, _
'      0, _
'      0
'
'    .Show vbModal
'
'    If Not .Cancelled Then
'      sRow = .Text _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & vbTab & "0" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "" _
'        & vbTab & "1"
'
'      With ctlSourceGrid
'        If mfIsEditingSeparator Then
'          ctlSourceGrid.RemoveItem lngRow
'
'          .AddItem sRow, lngRow
'          .Bookmark = .AddItemBookmark(lngRow)
'          .Update
'          .SelBookmarks.RemoveAll
'          .SelBookmarks.Add .AddItemBookmark(lngRow)
'
'        ElseIf mfIsCopyingSeparator Then
'          .AddItem sRow, lngRow + 1
'          .Bookmark = .AddItemBookmark(lngRow + 1)
'          .Update
'          .SelBookmarks.RemoveAll
'          .SelBookmarks.Add .AddItemBookmark(lngRow + 1)
'
'        Else
'          ' Adding a new separator.
'          .AddItem sRow
'          .Bookmark = .AddItemBookmark(.Rows - 1)
'          .Update
'          .SelBookmarks.RemoveAll
'          .SelBookmarks.Add .Bookmark
'        End If
'      End With
'
'      mfIsEditingSeparator = False
'      mfIsCopyingSeparator = False
'      Changed = True
'    End If
'  End With
'
'  UnLoad frmLinkSeparator
'  Set frmLinkSeparator = Nothing
'
'End Sub


Private Sub cmdMoveButtonLinkDown_Click()
  
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_DOWN
  End If
  Set ctlGrid = Nothing
  
End Sub


Private Sub cmdMoveButtonLinkDown_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveButtonLinkDown.Picture = cmdMoveButtonLinkDown.Picture

End Sub


Private Sub cmdMoveButtonLinkUp_Click()
  
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_UP
  End If
  Set ctlGrid = Nothing
  
End Sub


Private Sub cmdMoveButtonLinkUp_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveButtonLinkUp.Picture = cmdMoveButtonLinkUp.Picture

End Sub


Private Sub cmdMoveDropdownListLinkDown_Click()
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_DOWN
  End If
  Set ctlGrid = Nothing

End Sub


Private Sub cmdMoveDropdownListLinkDown_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveDropdownListLinkDown.Picture = cmdMoveDropdownListLinkDown.Picture

End Sub


Private Sub cmdMoveDropdownListLinkUp_Click()
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_UP
  End If
  Set ctlGrid = Nothing

End Sub


Private Sub cmdMoveDropdownListLinkUp_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveDropdownListLinkUp.Picture = cmdMoveDropdownListLinkUp.Picture

End Sub


Private Sub cmdMoveHypertextLinkDown_Click()
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_DOWN
  End If
  Set ctlGrid = Nothing
  
End Sub

Private Sub MoveView(piDirection As MoveDirection)

  Dim iLoop As Integer
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String

  intSourceRow = grdTableViews.AddItemRowIndex(grdTableViews.Bookmark)

  For iLoop = 0 To grdTableViews.Columns.Count - 1
    strSourceRow = strSourceRow & grdTableViews.Columns(iLoop).Text & _
      IIf(iLoop = grdTableViews.Columns.Count - 1, "", vbTab)
  Next iLoop

  If piDirection = MOVEDIRECTION_UP Then
    intDestinationRow = intSourceRow - 1
    grdTableViews.MovePrevious
  Else
    intDestinationRow = intSourceRow + 1
    grdTableViews.MoveNext
  End If

  For iLoop = 0 To grdTableViews.Columns.Count - 1
    strDestinationRow = strDestinationRow & grdTableViews.Columns(iLoop).Text & _
      IIf(iLoop = grdTableViews.Columns.Count - 1, "", vbTab)
  Next iLoop

  If piDirection = MOVEDIRECTION_UP Then
    grdTableViews.AddItem strSourceRow, intDestinationRow
    grdTableViews.RemoveItem intSourceRow + 1

    grdTableViews.SelBookmarks.RemoveAll
    grdTableViews.MovePrevious

  Else
    grdTableViews.RemoveItem intDestinationRow
    grdTableViews.RemoveItem intSourceRow

    grdTableViews.AddItem strDestinationRow, intSourceRow
    grdTableViews.AddItem strSourceRow, intDestinationRow

    grdTableViews.SelBookmarks.RemoveAll
    grdTableViews.MoveNext
  End If

  grdTableViews.Bookmark = grdTableViews.AddItemBookmark(intDestinationRow)
  grdTableViews.SelBookmarks.Add grdTableViews.AddItemBookmark(intDestinationRow)
  
  Changed = True

End Sub

Private Sub MoveLink(pctlGrid As SSDBGrid, _
  piDirection As MoveDirection)
  
  Dim iLoop As Integer
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim strSourceHiddenGroups As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  Dim strDestinationHiddenGroups As String
  
  intSourceRow = pctlGrid.AddItemRowIndex(pctlGrid.Bookmark)
  
  For iLoop = 0 To pctlGrid.Columns.Count - 1
    strSourceRow = strSourceRow & _
      IIf(pctlGrid.Columns(iLoop).Name = "HiddenGroups", "", pctlGrid.Columns(iLoop).Text) & _
      IIf(iLoop = pctlGrid.Columns.Count - 1, "", vbTab)
    strSourceHiddenGroups = pctlGrid.Columns("HiddenGroups").Text
  Next iLoop
  
  If piDirection = MOVEDIRECTION_UP Then
    intDestinationRow = intSourceRow - 1
    pctlGrid.MovePrevious
  Else
    intDestinationRow = intSourceRow + 1
    pctlGrid.MoveNext
  End If
  
  For iLoop = 0 To pctlGrid.Columns.Count - 1
    strDestinationRow = strDestinationRow & _
      IIf(pctlGrid.Columns(iLoop).Name = "HiddenGroups", "", pctlGrid.Columns(iLoop).Text) & _
      IIf(iLoop = pctlGrid.Columns.Count - 1, "", vbTab)
    strDestinationHiddenGroups = pctlGrid.Columns("HiddenGroups").Text
  Next iLoop
  
  If piDirection = MOVEDIRECTION_UP Then
    pctlGrid.AddItem strSourceRow, intDestinationRow
    pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intDestinationRow)
    pctlGrid.Columns("HiddenGroups").Text = strSourceHiddenGroups
    pctlGrid.Update
    
    pctlGrid.RemoveItem intSourceRow + 1
    
    pctlGrid.SelBookmarks.RemoveAll
    pctlGrid.MovePrevious
  
  Else
    pctlGrid.RemoveItem intDestinationRow
    pctlGrid.RemoveItem intSourceRow
    
    pctlGrid.AddItem strDestinationRow, intSourceRow
    pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intSourceRow)
    pctlGrid.Columns("HiddenGroups").Text = strDestinationHiddenGroups
    pctlGrid.Update
    
    pctlGrid.AddItem strSourceRow, intDestinationRow
    pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intDestinationRow)
    pctlGrid.Columns("HiddenGroups").Text = strSourceHiddenGroups
    pctlGrid.Update
    
    pctlGrid.SelBookmarks.RemoveAll
    pctlGrid.MoveNext
  End If
  
  pctlGrid.Bookmark = pctlGrid.AddItemBookmark(intDestinationRow)
  pctlGrid.SelBookmarks.Add pctlGrid.AddItemBookmark(intDestinationRow)
  
  Changed = True

End Sub

Private Sub cmdMoveHypertextLinkDown_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveHypertextLinkDown.Picture = cmdMoveHypertextLinkDown.Picture

End Sub

Private Sub cmdMoveHypertextLinkUp_Click()
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_UP
  End If
  Set ctlGrid = Nothing
  
End Sub

Private Sub cmdMoveHypertextLinkUp_LostFocus()
  'JPD 20031013 Fault 5827
  cmdMoveHypertextLinkUp.Picture = cmdMoveHypertextLinkUp.Picture
End Sub

Private Sub cmdMoveDocumentDown_Click()
  
  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_DOWN
  End If
  Set ctlGrid = Nothing

End Sub

Private Sub cmdMoveDocumentDown_LostFocus()
  cmdMoveDocumentDown.Picture = cmdMoveDocumentDown.Picture
End Sub

Private Sub cmdMoveDocumentUp_Click()

  Dim ctlGrid As SSDBGrid
  
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  If Not ctlGrid Is Nothing Then
    MoveLink ctlGrid, MOVEDIRECTION_UP
  End If
  Set ctlGrid = Nothing

End Sub

Private Sub cmdMoveDocumentUp_LostFocus()
  cmdMoveDocumentUp.Picture = cmdMoveDocumentUp.Picture
End Sub

Private Sub cmdMoveTableViewDown_Click()
  MoveView MOVEDIRECTION_DOWN
End Sub

Private Sub cmdMoveTableViewUp_Click()
  MoveView MOVEDIRECTION_UP
End Sub

Private Sub cmdOk_Click()
  'AE20071119 Fault #12607
  'If ValidateSetup Then
    'SaveChanges
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If
End Sub

Private Function ValidateSetup() As Boolean
  On Error GoTo ValidateError
  
  Dim sMsg As String
  Dim sMsgs As String
  Dim iLoop As Integer
  Dim varBookMark As Variant
  
  sMsg = ""
  
'  If (cboPersonnelTable.ListIndex <= 0) Then
'    sMsg = "The Personnel table is not defined."
'
'    ssTabStrip.Tab = giPAGE_GENERAL
'    cboPersonnelTable.SetFocus
'  End If
  
  ' Validate the view definitions
  If Len(sMsg) = 0 Then
    With grdTableViews
      For iLoop = 0 To (.Rows - 1)
        varBookMark = .AddItemBookmark(iLoop)
  
        If (Not .Columns("SingleRecord").CellValue(varBookMark)) _
          And (.Columns("HypertextLink").CellText(varBookMark) = "0") _
          And (.Columns("ButtonLink").CellText(varBookMark) = "0") _
          And (.Columns("DropdownListLink").CellText(varBookMark) = "0") Then
        
          sMsg = "No link type has been selected for the '" & .Columns("TableView").CellText(varBookMark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          Exit For
        End If
        
        If (Not .Columns("SingleRecord").CellValue(varBookMark)) _
          And (.Columns("HypertextLink").CellText(varBookMark) = "1") _
          And (Len(.Columns("HypertextLinkText").CellText(varBookMark)) = 0) Then
        
          sMsg = "No Hypertext Link text has been entered for the '" & .Columns("TableView").CellText(varBookMark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
          Exit For
        End If
        
        'JPD 20070710 Fault 12318
        'If (Not .Columns("SingleRecord").CellValue(varBookmark)) _
        '  And (.Columns("ButtonLink").CellText(varBookmark) = "1") _
        '  And (Len(.Columns("ButtonLinkPromptText").CellText(varBookmark)) = 0) Then
        '
        '  sMsg = "No Button Link prompt text has been entered for the '" & .Columns("TableView").CellText(varBookmark) & "' view."
        '
        '  ssTabStrip.Tab = giPAGE_GENERAL
        '  .SetFocus
        '  .Bookmark = .AddItemBookmark(iLoop)
        '  .SelBookmarks.RemoveAll
        '  .SelBookmarks.Add .Bookmark
        '
        '  Exit For
        'End If
        '
        If (Not .Columns("SingleRecord").CellValue(varBookMark)) _
          And (.Columns("ButtonLink").CellText(varBookMark) = "1") _
          And (Len(.Columns("ButtonLinkButtonText").CellText(varBookMark)) = 0) Then
        
          sMsg = "No Button Link button text has been entered for the '" & .Columns("TableView").CellText(varBookMark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
          Exit For
        End If
        
        If (Not .Columns("SingleRecord").CellValue(varBookMark)) _
          And (.Columns("DropdownListLink").CellText(varBookMark) = "1") _
          And (Len(.Columns("DropdownListLinkText").CellText(varBookMark)) = 0) Then
        
          sMsg = "No Dropdown List Link text has been entered for the '" & .Columns("TableView").CellText(varBookMark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
          Exit For
        End If
        
        If (Len(.Columns("LinksLinkText").CellText(varBookMark)) = 0) Then
        
          sMsg = "No Links Link text has been entered for the '" & .Columns("TableView").CellText(varBookMark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
          Exit For
        End If
        
      Next iLoop
    End With
  End If
    
  If Len(sMsg) > 0 Then
    MsgBox " The Self-service Intranet module is not configured correctly :" & vbCrLf & vbCrLf & _
      sMsg, vbExclamation + vbOKOnly, App.Title
  End If
  
  ValidateSetup = (Len(sMsg) = 0)
  Exit Function
  
ValidateError:
  MsgBox "Error validating the module setup." & vbCrLf & _
         Err.Description, vbExclamation + vbOKOnly, App.Title
  ValidateSetup = False

End Function

Private Function SaveChanges() As Boolean

  'AE20071119 Fault #12607
  SaveChanges = False
  
  If Not ValidateSetup Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass

  Dim iLoop As Integer
  Dim sSQL As String
  Dim sTableID As String
  Dim sViewID As String
  Dim sButtonLinkPromptText As String
  Dim sButtonLinkButtonText As String
  Dim sHypertextLinkText As String
  Dim sDropdownListLinkText As String
  Dim sButtonLink As String
  Dim sHypertextLink As String
  Dim sDropdownListLink As String
  Dim varBookMark As Variant
  Dim fSingleRecordView As Boolean
  Dim sLinksLinkText As String
  Dim sPageTitle As String
  Dim fWFOutOfOffice As Boolean
  
'  ' Save the configured Personnel table ID and Personnel table view ID.
'  With recModuleSetup
'    .Index = "idxModuleParameter"
'
'    ' Save the Self-service Intranet Personnel table ID.
'    .Seek "=", gsMODULEKEY_SSINTRANET, gsPARAMETERKEY_PERSONNELTABLE
'    If .NoMatch Then
'      .AddNew
'      !moduleKey = gsMODULEKEY_SSINTRANET
'      !parameterkey = gsPARAMETERKEY_PERSONNELTABLE
'    Else
'      .Edit
'    End If
'    !ParameterType = gsPARAMETERTYPE_TABLEID
'    !parametervalue = mlngPersonnelTableID
'    .Update
'  End With
  
  ' Clear the current database values.
  daoDb.Execute "DELETE FROM tmpSSIntranetLinks", dbFailOnError
  daoDb.Execute "DELETE FROM tmpSSIHiddenGroups", dbFailOnError
  daoDb.Execute "DELETE FROM tmpSSIViews", dbFailOnError
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookMark = .AddItemBookmark(iLoop)

      sTableID = .Columns("TableID").CellText(varBookMark)
      sViewID = .Columns("ViewID").CellText(varBookMark)
      fSingleRecordView = .Columns("SingleRecord").CellValue(varBookMark)
      fWFOutOfOffice = .Columns("WFOutOfOffice").CellValue(varBookMark)
      
      If fSingleRecordView Then
        sButtonLink = "0"
        sHypertextLink = "0"
        sDropdownListLink = "0"
        sPageTitle = ""
      Else
        sButtonLinkPromptText = .Columns("ButtonLinkPromptText").CellText(varBookMark)
        sButtonLinkButtonText = .Columns("ButtonLinkButtonText").CellText(varBookMark)
        sHypertextLinkText = .Columns("HypertextLinkText").CellText(varBookMark)
        sDropdownListLinkText = .Columns("DropdownListLinkText").CellText(varBookMark)
        sButtonLink = .Columns("ButtonLink").CellText(varBookMark)
        sHypertextLink = .Columns("HypertextLink").CellText(varBookMark)
        sDropdownListLink = .Columns("DropdownListLink").CellText(varBookMark)
        sPageTitle = .Columns("PageTitle").CellText(varBookMark)
      End If
      
      sLinksLinkText = .Columns("LinksLinkText").CellText(varBookMark)
      
      If sButtonLink = "0" Then
        sButtonLinkPromptText = ""
        sButtonLinkButtonText = ""
      End If

      If sHypertextLink = "0" Then
        sHypertextLinkText = ""
      End If

      If sDropdownListLink = "0" Then
        sDropdownListLinkText = ""
      End If

      sSQL = "INSERT INTO tmpSSIViews" & _
        " ([viewID], [tableID], [buttonLinkPromptText], [buttonLinkButtonText], [hypertextLinkText]," & _
        "  [dropdownListLinkText], [buttonLink], [hypertextLink], [dropdownListLink]," & _
        "  [singleRecordView], [sequence], [linksLinkText], [pageTitle], [WFOutOfOffice])" & _
        " VALUES" & _
        " (" & sViewID & "," & sTableID & ",'" & Replace(sButtonLinkPromptText, "'", "''") & "','" & Replace(sButtonLinkButtonText, "'", "''") & "','" & Replace(sHypertextLinkText, "'", "''") & "','" & _
        Replace(sDropdownListLinkText, "'", "''") & "'," & sButtonLink & "," & sHypertextLink & "," & sDropdownListLink & "," & _
        IIf(fSingleRecordView, "1", "0") & "," & CStr(iLoop) & ",'" & Replace(sLinksLinkText, "'", "''") & "','" & Replace(sPageTitle, "'", "''") & "'," & IIf(fWFOutOfOffice, "1", "0") & _
        ")"
      daoDb.Execute sSQL, dbFailOnError
    Next iLoop
  End With
  
  ' Write the parameter values to the local database.
  SaveLinkParameters SSINTLINK_HYPERTEXT
  SaveLinkParameters SSINTLINK_BUTTON
  SaveLinkParameters SSINTLINK_DROPDOWNLIST
  SaveLinkParameters SSINTLINK_DOCUMENT
  
  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbNormal
  
End Function

Private Sub ReadParameters()
  ' Read the parameter values from the database into local variables.
  ' Read the Self-service Intranet parameter values from the database into the grids.
  Dim sSQL As String
  Dim rsLinks As DAO.Recordset
  Dim rsHiddenGroups As DAO.Recordset
  Dim sAddString As String
  Dim ctlGrid As SSDBGrid
  Dim sHiddenGroups As String
  Dim iIndex As Integer
  Dim fWorkflowLicensed As Boolean
  
  ' Get the configured Personnel table ID and Personnel table view ID.
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Personnel module Personnel table ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If .NoMatch Then
      mlngPersonnelTableID = 0
    Else
      mlngPersonnelTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
  End With
  
  sSQL = "SELECT tmpSSIViews.*, tmpViews.viewName, tmpTables.TableName" & _
        " FROM tmpSSIViews, tmpViews, tmpTables " & _
        " WHERE (tmpSSIViews.ViewID <> -1) " & _
        "  AND (tmpSSIViews.viewID = tmpViews.viewID) " & _
        "  AND (tmpViews.viewTableID = tmpTables.tableID) " & _
        "UNION " & _
        "SELECT tmpSSIViews.*, '', tmpTables.TableName" & _
        " FROM tmpSSIViews, tmpTables " & _
        " WHERE (tmpSSIViews.ViewID = -1) " & _
        "  AND (tmpSSIViews.TableID = tmpTables.tableID) " & _
        " ORDER BY [sequence]"

  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  While Not rsLinks.EOF
    sAddString = CreateTableViewName(rsLinks!TableName, rsLinks!ViewName) & _
          vbTab & CStr(IIf(IsNull(rsLinks!TableID), 0, rsLinks!TableID)) & _
          vbTab & CStr(IIf(IsNull(rsLinks!ViewID), 0, rsLinks!ViewID)) & _
          vbTab & IIf(IsNull(rsLinks!ButtonLinkPromptText), "", rsLinks!ButtonLinkPromptText) & _
          vbTab & IIf(IsNull(rsLinks!ButtonLinkButtonText), "", rsLinks!ButtonLinkButtonText) & _
          vbTab & IIf(IsNull(rsLinks!HypertextLinkText), "", rsLinks!HypertextLinkText) & _
          vbTab & IIf(IsNull(rsLinks!DropdownListLinkText), "", rsLinks!DropdownListLinkText) & _
          vbTab & IIf(IIf(IsNull(rsLinks!ButtonLink), False, rsLinks!ButtonLink), "1", "0") & _
          vbTab & IIf(IIf(IsNull(rsLinks!HypertextLink), False, rsLinks!HypertextLink), "1", "0") & _
          vbTab & IIf(IIf(IsNull(rsLinks!DropdownListLink), False, rsLinks!DropdownListLink), "1", "0") & _
          vbTab & IIf(IsNull(rsLinks!LinksLinkText), "", rsLinks!LinksLinkText) & _
          vbTab & IIf(IsNull(rsLinks!PageTitle), "", rsLinks!PageTitle) & _
          vbTab & rsLinks!SingleRecordView & _
          vbTab & rsLinks!WFOutOfOffice

    grdTableViews.AddItem sAddString
    
    grdTableViews.MoveFirst
    grdTableViews.SelBookmarks.Add grdTableViews.AddItemBookmark(0)

'    If maryTableViewIndex(0) <> "" Then
'      ReDim Preserve maryTableViewIndex(UBound(maryTableViewIndex) + 1)
'      maryTableViewIndex(UBound(maryTableViewIndex)) = CreateTableViewTag(CStr(rsLinks!TableID), CStr(rsLinks!ViewID))
'    Else
'      maryTableViewIndex(0) = CreateTableViewTag(CStr(rsLinks!TableID), CStr(rsLinks!ViewID))
'    End If
    
    ' Create the link grids for the view.
    AddTableViewGrids rsLinks!TableID, rsLinks!ViewID
    
    rsLinks.MoveNext
  Wend
  rsLinks.Close
  Set rsLinks = Nothing
  
  ' Load the link definitions
  sSQL = "SELECT *" & _
    " FROM tmpSSIntranetLinks" & _
    " ORDER BY linkOrder"
    
  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  fWorkflowLicensed = IsModuleEnabled(modWorkflow)

  While Not rsLinks.EOF
    Set ctlGrid = Nothing
    
    ' NPG20100520 Fault HRPRO-938 - Don't display workflow items if not licensed for workflow.
    If Not (rsLinks!UtilityType = 25 And Not fWorkflowLicensed) And Not (rsLinks!Element_Type = 3 And Not fWorkflowLicensed) Then
    Select Case rsLinks!LinkType
      Case SSINTLINK_HYPERTEXT
        iIndex = -1
        For Each ctlGrid In grdHypertextLinks
          If ctlGrid.Tag = CreateTableViewTag(CStr(rsLinks!TableID), CStr(rsLinks!ViewID)) Then
            iIndex = ctlGrid.Index
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
        
        If iIndex > -1 Then
          Set ctlGrid = grdHypertextLinks(iIndex)
          sAddString = rsLinks!Text & _
            vbTab & rsLinks!URL & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & _
            vbTab & vbTab & IIf(IsNull(rsLinks!NewWindow), "0", IIf(rsLinks!NewWindow, "1", "0")) & _
            vbTab & rsLinks!EMailAddress & _
            vbTab & rsLinks!EMailSubject & _
            vbTab & rsLinks!AppFilePath & _
            vbTab & rsLinks!AppParameters & _
            vbTab & IIf(IsNull(rsLinks!Element_Type), "0", CStr(rsLinks!Element_Type)) & _
            vbTab & IIf(IsNull(rsLinks!SeparatorOrientation), "0", CStr(rsLinks!SeparatorOrientation)) & _
            vbTab & IIf(IsNull(rsLinks!PictureID), "0", CStr(rsLinks!PictureID))
        End If
        
      Case SSINTLINK_BUTTON
        iIndex = -1
        For Each ctlGrid In grdButtonLinks
          If ctlGrid.Tag = CreateTableViewTag(CStr(rsLinks!TableID), CStr(rsLinks!ViewID)) Then
            iIndex = ctlGrid.Index
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
        
        If iIndex > -1 Then
          Set ctlGrid = grdButtonLinks(iIndex)
          sAddString = rsLinks!Prompt & _
            vbTab & rsLinks!Text & vbTab & rsLinks!URL & _
            vbTab & CStr(rsLinks!ScreenID) & vbTab & rsLinks!PageTitle & _
            vbTab & CStr(rsLinks!StartMode) & vbTab & CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & _
            vbTab & vbTab & IIf(IsNull(rsLinks!NewWindow), "0", IIf(rsLinks!NewWindow, "1", "0")) & _
            vbTab & rsLinks!EMailAddress & vbTab & rsLinks!EMailSubject & vbTab & rsLinks!AppFilePath & vbTab & rsLinks!AppParameters & _
            vbTab & CStr(IIf(IsNull(rsLinks!Element_Type), "0", CStr(rsLinks!Element_Type))) & vbTab & IIf(IsNull(rsLinks!SeparatorOrientation), "0", CStr(rsLinks!SeparatorOrientation)) & _
            vbTab & IIf(IsNull(rsLinks!PictureID), "0", CStr(rsLinks!PictureID)) & vbTab & IIf(IsNull(rsLinks!Chart_ShowLegend), "0", IIf(rsLinks!Chart_ShowLegend, "1", "0")) & _
            vbTab & CStr(rsLinks!Chart_Type) & vbTab & IIf(IsNull(rsLinks!Chart_ShowGrid), "0", IIf(rsLinks!Chart_ShowGrid, "1", "0")) & _
            vbTab & IIf(IsNull(rsLinks!Chart_StackSeries), "0", IIf(rsLinks!Chart_StackSeries, "1", "0")) & _
            vbTab & CStr(rsLinks!Chart_ViewID) & vbTab & CStr(rsLinks!Chart_TableID) & vbTab & CStr(rsLinks!Chart_ColumnID) & vbTab & CStr(rsLinks!Chart_FilterID) & _
            vbTab & CStr(rsLinks!Chart_AggregateType) & vbTab & IIf(IsNull(rsLinks!Chart_ShowValues), "0", IIf(rsLinks!Chart_ShowValues, "1", "0")) & _
            vbTab & IIf(IsNull(rsLinks!UseFormatting), "0", IIf(rsLinks!UseFormatting, "1", "0")) & _
            vbTab & IIf(IsNull(rsLinks!Formatting_DecimalPlaces), "0", CStr(rsLinks!Formatting_DecimalPlaces)) & vbTab & IIf(IsNull(rsLinks!Formatting_Use1000Separator), "0", IIf(rsLinks!Formatting_Use1000Separator, "1", "0")) & _
            vbTab & rsLinks!Formatting_Prefix & vbTab & rsLinks!Formatting_Suffix & _
            vbTab & IIf(IsNull(rsLinks!UseConditionalFormatting), "0", IIf(rsLinks!UseConditionalFormatting, "1", "0")) & _
            vbTab & rsLinks!ConditionalFormatting_Operator_1 & vbTab & rsLinks!ConditionalFormatting_Value_1 & vbTab & rsLinks!ConditionalFormatting_Style_1 & vbTab & rsLinks!ConditionalFormatting_Colour_1 & _
            vbTab & rsLinks!ConditionalFormatting_Operator_2 & vbTab & rsLinks!ConditionalFormatting_Value_2 & vbTab & rsLinks!ConditionalFormatting_Style_2 & vbTab & rsLinks!ConditionalFormatting_Colour_2 & _
            vbTab & rsLinks!ConditionalFormatting_Operator_3 & vbTab & rsLinks!ConditionalFormatting_Value_3 & vbTab & rsLinks!ConditionalFormatting_Style_3 & vbTab & rsLinks!ConditionalFormatting_Colour_3 & _
            vbTab & rsLinks!SeparatorColour & vbTab & CStr(rsLinks!InitialDisplayMode) & vbTab & CStr(rsLinks!Chart_TableID_2) & vbTab & CStr(rsLinks!Chart_ColumnID_2) & _
            vbTab & CStr(rsLinks!Chart_TableID_3) & vbTab & CStr(rsLinks!Chart_ColumnID_3) & vbTab & CStr(rsLinks!Chart_SortOrderID & vbTab & CStr(rsLinks!Chart_SortDirection)) & vbTab & CStr(rsLinks!Chart_ColourID) & _
            vbTab & IIf(IsNull(rsLinks!Chart_ShowPercentages), "0", IIf(rsLinks!Chart_ShowPercentages, "1", "0"))
       End If
          
      Case SSINTLINK_DROPDOWNLIST
        iIndex = -1
        For Each ctlGrid In grdDropdownListLinks
          If ctlGrid.Tag = CreateTableViewTag(CStr(rsLinks!TableID), CStr(rsLinks!ViewID)) Then
            iIndex = ctlGrid.Index
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
    
        If iIndex > -1 Then
          Set ctlGrid = grdDropdownListLinks(iIndex)
          sAddString = rsLinks!Text & _
            vbTab & rsLinks!URL & _
            vbTab & CStr(rsLinks!ScreenID) & _
            vbTab & rsLinks!PageTitle & _
            vbTab & CStr(rsLinks!StartMode) & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & _
            vbTab & vbTab & IIf(IsNull(rsLinks!NewWindow), "0", IIf(rsLinks!NewWindow, "1", "0")) & _
            vbTab & rsLinks!EMailAddress & _
            vbTab & rsLinks!EMailSubject & _
            vbTab & rsLinks!AppFilePath & _
            vbTab & rsLinks!AppParameters
        End If
        
      Case SSINTLINK_DOCUMENT
        iIndex = -1
        For Each ctlGrid In grdDocuments
          If ctlGrid.Tag = CreateTableViewTag(CStr(rsLinks!TableID), CStr(rsLinks!ViewID)) Then
            iIndex = ctlGrid.Index
            Exit For
          End If
        Next ctlGrid
        Set ctlGrid = Nothing
    
        If iIndex > -1 Then
          Set ctlGrid = grdDocuments(iIndex)
          sAddString = rsLinks!Text & _
            vbTab & rsLinks!URL & _
            vbTab & CStr(rsLinks!ScreenID) & _
            vbTab & rsLinks!PageTitle & _
            vbTab & CStr(rsLinks!StartMode) & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & _
            vbTab & CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & _
            vbTab & vbTab & IIf(IsNull(rsLinks!NewWindow), "0", IIf(rsLinks!NewWindow, "1", "0")) & _
            vbTab & rsLinks!EMailAddress & _
            vbTab & rsLinks!EMailSubject & _
            vbTab & rsLinks!AppFilePath & _
            vbTab & rsLinks!AppParameters & _
            vbTab & rsLinks!DocumentFilePath & _
            vbTab & rsLinks!DisplayDocumentHyperlink
        End If

    End Select
        
    If Not ctlGrid Is Nothing Then
      ctlGrid.AddItem sAddString
          
      ' Add the hidden groups info.
      sHiddenGroups = ""
                
      sSQL = "SELECT *" & _
        " FROM tmpSSIHiddenGroups" & _
        " WHERE linkID = " & CStr(rsLinks!id)
      Set rsHiddenGroups = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      While Not rsHiddenGroups.EOF
        sHiddenGroups = sHiddenGroups & rsHiddenGroups!GroupName & vbTab
  
        rsHiddenGroups.MoveNext
      Wend
      rsHiddenGroups.Close
      Set rsHiddenGroups = Nothing
      
      If Len(sHiddenGroups) > 0 Then
        sHiddenGroups = vbTab & sHiddenGroups
      End If
      
      ctlGrid.Bookmark = ctlGrid.AddItemBookmark(ctlGrid.Rows - 1)
      'JPD 20050222 Fault 9762
      ' Don't ask me why, but sometimes assigning the 'sHiddenGroups' value to the
      ' 'text' property of the 'HiddenGroups' column didn't work. Assigning an
      ' empty string value first, and then the 'sHiddenGroups' value seemed to sort things out though.
      ctlGrid.Columns("HiddenGroups").Text = ""
      ctlGrid.Columns("HiddenGroups").Text = sHiddenGroups
      ctlGrid.Update
      
      'JPD 20050118 Fault 9722
      ctlGrid.MoveFirst
    End If
    End If
    rsLinks.MoveNext
  Wend
  rsLinks.Close
  Set rsLinks = Nothing
  
  mfChanged = False
  
End Sub

Private Sub PopulateAccessCombo()
  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset

  ' Get the recordset of user groups and their access on this definition.
  sSQL = "SELECT name FROM sysusers" & _
    " WHERE gid = uid AND gid > 0" & _
    "   AND not (name like 'ASRSys%') AND not (name like 'db[_]%')" & _
    " ORDER BY name"
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  ' Add the 'All Groups' item.
  With cboSecurityGroup
    .Clear
    .AddItem "(All Groups)"
  End With


  With rsGroups
    Do While Not .EOF
      ' Add the user groups and their access on this definition to the access grid.
'      If InStr(vbTab & UCase(psHiddenGroups) & vbTab, vbTab & UCase(Trim(!Name)) & vbTab) > 0 Then
'        sVisibility = "False"
'        fAllVisible = False
'      Else
'        sVisibility = "True"
'      End If
'
      cboSecurityGroup.AddItem !Name '
      .MoveNext
    Loop
      
    .Close
  End With
  Set rsGroups = Nothing
  
  cboSecurityGroup.ListIndex = 0

End Sub



Private Sub cmdPreview_Click()
  Dim strFileName As String
    
  strFileName = GeneratePreviewHTML()
  
  DisplayInBrowser
  
End Sub

Private Sub cmdRemoveAllButtonLinks_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  
  If Not ctlCurrentGrid Is Nothing Then
    ctlCurrentGrid.RemoveAll
    Changed = True
  End If
  Set ctlCurrentGrid = Nothing

End Sub

Private Sub cmdRemoveAllDropdownListLinks_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  
  If Not ctlCurrentGrid Is Nothing Then
    ctlCurrentGrid.RemoveAll
    Changed = True
  End If
  Set ctlCurrentGrid = Nothing
  
End Sub

Private Sub cmdRemoveAllHypertextLinks_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  
  If Not ctlCurrentGrid Is Nothing Then
    ctlCurrentGrid.RemoveAll
    Changed = True
  End If
  Set ctlCurrentGrid = Nothing
  
End Sub

Private Sub cmdRemoveAllDocuments_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  
  If Not ctlCurrentGrid Is Nothing Then
    ctlCurrentGrid.RemoveAll
    Changed = True
  End If
  Set ctlCurrentGrid = Nothing

End Sub

Private Sub cmdRemoveAllTableViews_Click()

  Dim ctlGrid As SSDBGrid
  
  'NPG20080409 Fault 13061
  If MsgBox("All existing links for all views will be deleted. Are you sure you wish to continue ?", vbYesNo + vbQuestion, Me.Caption) <> vbYes Then
    Exit Sub
  End If
  
  grdTableViews.RemoveAll
  
  With cboHypertextLinkView
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = -1
    .ListIndex = 0
  End With
  For Each ctlGrid In grdHypertextLinks
    If ctlGrid.Index > 0 Then
      UnLoad ctlGrid
    End If
  Next ctlGrid
  Set ctlGrid = Nothing
  
  With cboButtonLinkView
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = -1
    .ListIndex = 0
  End With
  For Each ctlGrid In grdButtonLinks
    If ctlGrid.Index > 0 Then
      UnLoad ctlGrid
    End If
  Next ctlGrid
  Set ctlGrid = Nothing
  
  With cboDropdownListLinkView
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = -1
    .ListIndex = 0
  End With
  For Each ctlGrid In grdDropdownListLinks
    If ctlGrid.Index > 0 Then
      UnLoad ctlGrid
    End If
  Next ctlGrid
  Set ctlGrid = Nothing
  
  With cboDocumentView
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = -1
    .ListIndex = 0
  End With
  For Each ctlGrid In grdDocuments
    If ctlGrid.Index > 0 Then
      UnLoad ctlGrid
    End If
  Next ctlGrid
  Set ctlGrid = Nothing

  Changed = True

  RefreshControls
  
End Sub

Private Sub cmdRemoveButtonLink_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  
  If Not ctlCurrentGrid Is Nothing Then
    RemoveLink ctlCurrentGrid
  End If
  Set ctlCurrentGrid = Nothing

End Sub

Private Sub cmdRemoveDropdownListLink_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_DROPDOWNLIST)
  
  If Not ctlCurrentGrid Is Nothing Then
    RemoveLink ctlCurrentGrid
  End If
  Set ctlCurrentGrid = Nothing

End Sub


Private Sub cmdRemoveHypertextLink_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_HYPERTEXT)
  
  If Not ctlCurrentGrid Is Nothing Then
    RemoveLink ctlCurrentGrid
  End If
  Set ctlCurrentGrid = Nothing
  
End Sub

Private Sub cmdRemoveDocument_Click()

  Dim ctlCurrentGrid As SSDBGrid
  
  Set ctlCurrentGrid = CurrentLinkGrid(SSINTLINK_DOCUMENT)
  
  If Not ctlCurrentGrid Is Nothing Then
    RemoveLink ctlCurrentGrid
  End If
  Set ctlCurrentGrid = Nothing

End Sub

Private Sub cmdRemoveTableView_Click()

  Dim sRowsToDelete As String
  Dim iCount As Integer
  Dim ctlGrid As SSDBGrid
  Dim iIndex As Integer
  Dim sTableID As String
  Dim sViewID As String
  Dim sTag As String
  
  sRowsToDelete = ","

  'NPG20080409 Fault 13061
  If MsgBox("All existing links for this view will be deleted. Are you sure you wish to continue ?", vbYesNo + vbQuestion, Me.Caption) <> vbYes Then
    Exit Sub
  End If

  With grdTableViews
    If .Rows = 1 Then
      cmdRemoveAllTableViews_Click
    Else
      For iCount = 0 To .SelBookmarks.Count - 1
        sRowsToDelete = sRowsToDelete & CStr(.AddItemRowIndex(.SelBookmarks(iCount))) & ","
      Next iCount
      
      For iCount = (.Rows - 1) To 0 Step -1
        If InStr(sRowsToDelete, "," & CStr(iCount) & ",") > 0 Then
          
          .Bookmark = .AddItemBookmark(iCount)
          sTableID = .Columns("viewID").Text
          sViewID = .Columns("viewID").Text
          sTag = CreateTableViewTag(sTableID, sViewID)
          
          For Each ctlGrid In grdHypertextLinks
            If ctlGrid.Tag = sTag Then
              UnLoad ctlGrid
              Exit For
            End If
          Next ctlGrid
          Set ctlGrid = Nothing

          For Each ctlGrid In grdButtonLinks
            If ctlGrid.Tag = sTag Then
              UnLoad ctlGrid
              Exit For
            End If
          Next ctlGrid
          Set ctlGrid = Nothing

          For Each ctlGrid In grdDropdownListLinks
            If ctlGrid.Tag = sTag Then
              UnLoad ctlGrid
              Exit For
            End If
          Next ctlGrid
          Set ctlGrid = Nothing
          
          .RemoveItem iCount
        End If
      Next iCount
    End If
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  RefreshTableViewCombos
  
  RefreshTableViewsCollection
  
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

Private Sub Form_Load()
  
  Const GRIDROWHEIGHT = 239
  
  Screen.MousePointer = vbHourglass
  
  Set mcolSSITableViews = New clsSSITableViews
  
  ReDim maryTableViewIndex(0)
  
  grdHypertextLinks(0).RowHeight = GRIDROWHEIGHT
  grdButtonLinks(0).RowHeight = GRIDROWHEIGHT
  grdDropdownListLinks(0).RowHeight = GRIDROWHEIGHT
  grdDocuments(0).RowHeight = GRIDROWHEIGHT
  grdTableViews.RowHeight = GRIDROWHEIGHT
  
  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If
  
  ' Read the current settings from the database.
  mfLoading = True
  
  ReadParameters
  
  RefreshTableViewCombos
  
  PopulateAccessCombo
  
  RefreshTableViewsCollection
  
  mfLoading = False
  
  ssTabStrip.Tab = giPAGE_GENERAL
  
  Changed = False
  
  RefreshControls
  
  Screen.MousePointer = vbNormal
  
End Sub

Private Sub RefreshTableViewsCollection()
  
  Dim iLoop As Integer
  Dim lngTableID As Long
  Dim lngViewID As Long
  Dim sTableViewName As String
  Dim varBookMark As Variant
 
  Set mcolSSITableViews = New clsSSITableViews

  With grdTableViews
  
    For iLoop = 0 To (.Rows - 1) Step 1
      
      varBookMark = .AddItemBookmark(iLoop)
      
      lngTableID = .Columns("TableID").CellValue(varBookMark)
      lngViewID = .Columns("ViewID").CellValue(varBookMark)
      sTableViewName = CreateTableViewName(GetTableName(lngTableID), GetViewName(lngViewID))
    
      mcolSSITableViews.Add lngTableID, lngViewID, sTableViewName
      
    Next iLoop
  
  End With
  
End Sub

Private Sub RefreshTableViewCombos()

  ' Populate the views combos.
  Dim sSQL As String
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim varBookMark As Variant
  Dim lngTableID As Long
  Dim lngViewID As Long
  Dim sTableViewName As String
  
  cboHypertextLinkView.Clear
  cboButtonLinkView.Clear
  cboDropdownListLinkView.Clear
  cboDocumentView.Clear
  
  With grdTableViews
  
    For iLoop = 0 To (.Rows - 1)
    
      varBookMark = .AddItemBookmark(iLoop)
      
      lngTableID = .Columns("TableID").CellValue(varBookMark)
      lngViewID = .Columns("ViewID").CellValue(varBookMark)
      sTableViewName = CreateTableViewName(GetTableName(lngTableID), GetViewName(lngViewID))
     
'      If maryTableViewIndex(0) <> "" Then
'        ReDim Preserve maryTableViewIndex(UBound(maryTableViewIndex) + 1)
'        maryTableViewIndex(UBound(maryTableViewIndex)) = CreateTableViewTag(CStr(lngTableID), CStr(lngViewID))
'      Else
'        maryTableViewIndex(0) = CreateTableViewTag(CStr(lngTableID), CStr(lngViewID))
'      End If
     
      ' Add the selected single record & multiple record views to the link view combos.
      cboHypertextLinkView.AddItem sTableViewName
      cboButtonLinkView.AddItem sTableViewName
      ' Add the viewID too!
      cboButtonLinkView.ItemData(iLoop) = lngViewID
      cboDropdownListLinkView.AddItem sTableViewName
      cboDocumentView.AddItem sTableViewName
    
    Next iLoop
  
  End With
  
  ' If the link view combos are empty, add a <None> item.
  ' Otherwise select the Single Record view if it exists.
  ' Otherwise select the first item.
  iIndex = 0
  
  If cboHypertextLinkView.ListCount = 0 Then
    cboHypertextLinkView.AddItem "<None>"
    cboHypertextLinkView.ItemData(cboHypertextLinkView.NewIndex) = 0
    iIndex = cboHypertextLinkView.NewIndex
  Else
    For iLoop = 0 To cboHypertextLinkView.ListCount - 1
      If cboHypertextLinkView.ItemData(iLoop) = SingleRecordViewID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
  If cboButtonLinkView.ListCount = 0 Then
    cboButtonLinkView.AddItem "<None>"
    cboButtonLinkView.ItemData(cboButtonLinkView.NewIndex) = 0
    iIndex = cboButtonLinkView.NewIndex
  Else
    For iLoop = 0 To cboButtonLinkView.ListCount - 1
      If cboButtonLinkView.ItemData(iLoop) = SingleRecordViewID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
  If cboDropdownListLinkView.ListCount = 0 Then
    cboDropdownListLinkView.AddItem "<None>"
    cboDropdownListLinkView.ItemData(cboDropdownListLinkView.NewIndex) = 0
    iIndex = cboDropdownListLinkView.NewIndex
  Else
    For iLoop = 0 To cboDropdownListLinkView.ListCount - 1
      If cboDropdownListLinkView.ItemData(iLoop) = SingleRecordViewID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
  If cboDocumentView.ListCount = 0 Then
    cboDocumentView.AddItem "<None>"
    cboDocumentView.ItemData(cboDocumentView.NewIndex) = 0
    iIndex = cboDocumentView.NewIndex
  Else
    For iLoop = 0 To cboDocumentView.ListCount - 1
      If cboDocumentView.ItemData(iLoop) = SingleRecordViewID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If


  cboHypertextLinkView.ListIndex = iIndex
  cboButtonLinkView.ListIndex = iIndex
  cboDropdownListLinkView.ListIndex = iIndex
  cboDocumentView.ListIndex = iIndex
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' If the user cancels or tries to close the form
  'AE20071119 Fault #12607
  'If UnloadMode <> vbFormCode And cmdOK.Enabled Then
  If Changed Then
    Select Case MsgBox("Apply module changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        'AE20071119 Fault #12607
'        If ValidateSetup Then
'          SaveChanges
'        Else
'          Cancel = True
'        End If
        Cancel = (Not SaveChanges)
    End Select
  End If

End Sub

Private Sub Form_Resize()
  
  On Error GoTo ErrorTrap
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
ErrorTrap:
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  ' Save the form position to registry.
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
  End If

End Sub

Private Sub grdButtonLinks_DblClick(Index As Integer)
  
  If Not mblnReadOnly Then
    If grdButtonLinks(Index).Rows > 0 Then
        cmdEditButtonLink_Click
    Else
      cmdAddButtonLink_Click
    End If
  End If

End Sub

Private Sub grdButtonLinks_InitColumnProps(Index As Integer)
        grdButtonLinks(Index).RowSelectionStyle = ssRowSelectionStyle3D

End Sub

Private Sub grdButtonLinks_RowColChange(Index As Integer, ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshControls
End Sub

Private Sub grdButtonLinks_RowLoaded(Index As Integer, ByVal Bookmark As Variant)
  
  If cboSecurityGroup.ListIndex >= 0 Then
      If (InStr(1, grdButtonLinks(Index).Columns("HiddenGroups").CellValue(Bookmark), cboSecurityGroup.List(cboSecurityGroup.ListIndex), vbTextCompare) > 0) _
            Or cboSecurityGroup.List(cboSecurityGroup.ListIndex) = "(All Groups)" Then
        grdButtonLinks(Index).Columns(0).CellStyleSet "ssDisabled"
        grdButtonLinks(Index).Columns(1).CellStyleSet "ssDisabled"
      Else
        grdButtonLinks(Index).Columns(0).CellStyleSet "ssEnabled"
        grdButtonLinks(Index).Columns(1).CellStyleSet "ssEnabled"

      End If
  End If

End Sub

Private Sub grdButtonLinks_SelChange(Index As Integer, ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  
  If Not cmdEditButtonLink.Enabled Then
    RefreshControls
  End If

End Sub

Private Sub grdDocuments_DblClick(Index As Integer)

  If Not mblnReadOnly Then
    If grdDocuments(Index).Rows > 0 Then
      cmdEditDocument_Click
    Else
      cmdAddDocument_Click
    End If
  End If

End Sub

Private Sub grdDocuments_InitColumnProps(Index As Integer)
        grdDocuments(Index).RowSelectionStyle = ssRowSelectionStyle3D

End Sub

Private Sub grdDocuments_RowColChange(Index As Integer, ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshControls
End Sub

Private Sub grdDocuments_SelChange(Index As Integer, ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

  If Not cmdEditDocument.Enabled Then
    RefreshControls
  End If

End Sub

Private Sub grdDropdownListLinks_DblClick(Index As Integer)
  
  If Not mblnReadOnly Then
    If grdDropdownListLinks(Index).Rows > 0 Then
      cmdEditDropdownListLink_Click
    Else
      cmdAddDropdownListLink_Click
    End If
  End If

End Sub

Private Sub grdDropdownListLinks_InitColumnProps(Index As Integer)
        grdDropdownListLinks(Index).RowSelectionStyle = ssRowSelectionStyle3D

End Sub

Private Sub grdDropdownListLinks_RowColChange(Index As Integer, ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshControls
End Sub

Private Sub grdDropdownListLinks_SelChange(Index As Integer, ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  
  If Not cmdEditDropdownListLink.Enabled Then
    RefreshControls
  End If
  
End Sub

Private Sub grdHypertextLinks_DblClick(Index As Integer)
  
  If Not mblnReadOnly Then
    If grdHypertextLinks(Index).Rows > 0 Then
        cmdEditHypertextLink_Click
    Else
      cmdAddHyperTextLink_Click
    End If
  End If
  
End Sub

Private Sub grdHypertextLinks_InitColumnProps(Index As Integer)
        grdHypertextLinks(Index).RowSelectionStyle = ssRowSelectionStyle3D

End Sub

Private Sub grdHypertextLinks_RowColChange(Index As Integer, ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshControls
End Sub

Private Sub grdHypertextLinks_SelChange(Index As Integer, ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  
  If Not cmdEditHypertextLink.Enabled Then
    RefreshControls
  End If
  
End Sub

Private Sub grdTableViews_DblClick()
  
  If Not mblnReadOnly Then
    If grdTableViews.Rows > 0 Then
      cmdEditTableView_Click
    Else
      cmdAddTableView_Click
    End If
  End If

End Sub

Private Sub grdTableViews_InitColumnProps()
        grdTableViews.RowSelectionStyle = ssRowSelectionStyle3D
End Sub

Private Sub grdTableViews_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshControls
End Sub

Private Sub ssTabStrip_Click(PreviousTab As Integer)
  
  If Not mblnReadOnly Then
    
    'NPG20071101 Fault 12566
    fraViews.Enabled = (ssTabStrip.Tab = giPAGE_GENERAL)
    
    fraHypertextLinks.Enabled = (ssTabStrip.Tab = giPAGE_HYPERTEXTLINKS)
    fraButtonLinks.Enabled = (ssTabStrip.Tab = giPAGE_BUTTONLINKS)
    fraDropdownListLinks.Enabled = (ssTabStrip.Tab = giPAGE_DROPDOWNLISTLINKS)
    fraDocuments.Enabled = (ssTabStrip.Tab = giPAGE_DOCUMENTS)
  End If

  RefreshControls
  
End Sub

Private Sub AddTableViewGrids(plngTableID As Long, plngViewID As Long)
  
  ' Create the link grids for the given view.
  Dim ctlGrid As SSDBGrid
  Dim iIndex As Integer
  Dim sTag As String
  
  iIndex = 0
  
  sTag = CreateTableViewTag(CStr(plngTableID), CStr(plngViewID))
  
  ' Check if the grid already exists.
  For Each ctlGrid In grdHypertextLinks
    If ctlGrid.Tag = sTag Then
      iIndex = ctlGrid.Index
    End If
  Next ctlGrid
  Set ctlGrid = Nothing
  
  If iIndex > 0 Then
    grdHypertextLinks(iIndex).RemoveAll
    grdButtonLinks(iIndex).RemoveAll
    grdDropdownListLinks(iIndex).RemoveAll
    grdDocuments(iIndex).RemoveAll
  Else
    Load grdHypertextLinks(grdHypertextLinks.UBound + 1)
    With grdHypertextLinks(grdHypertextLinks.UBound)
      .Tag = sTag
      .Enabled = True
    End With
    
    Load grdButtonLinks(grdButtonLinks.UBound + 1)
    With grdButtonLinks(grdButtonLinks.UBound)
      .Tag = sTag
      .Enabled = True
    End With
    
    Load grdDropdownListLinks(grdDropdownListLinks.UBound + 1)
    With grdDropdownListLinks(grdDropdownListLinks.UBound)
      .Tag = sTag
      .Enabled = True
    End With
    
    Load grdDocuments(grdDocuments.UBound + 1)
    With grdDocuments(grdDocuments.UBound)
      .Tag = sTag
      .Enabled = True
    End With

  End If
  
  RefreshTableViewsCollection
  
End Sub


Private Sub BuildUserGroupCollection()
  ' Populate the access grid.
  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset
  
  Dim objGroup As clsSecgroup

  Set mcolGroups = New Collection
      
  ' Get the recordset of user groups and their access on this definition.
  sSQL = "SELECT name FROM sysusers" & _
    " WHERE gid = uid AND gid > 0" & _
    "   AND not (name like 'ASRSys%') AND not (name like 'db[_]%')" & _
    " ORDER BY name"
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsGroups
    Do While Not .EOF
      Set objGroup = New clsSecgroup
                        
      ' Add element type 3 - PWF steps
      objGroup.GroupName = Trim(!Name) & "3"
      objGroup.Allow = True
      
      mcolGroups.Add objGroup, objGroup.GroupName
            
      Set objGroup = New clsSecgroup
            
      ' Repeat for Today's Events (element type 5)
      objGroup.GroupName = Trim(!Name) & "5"
      objGroup.Allow = True
      
      mcolGroups.Add objGroup, objGroup.GroupName
            
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsGroups = Nothing
      
End Sub

Public Function Exists(ByVal sColKey As String) As Boolean
  ' Return TRUE if the given key exists in the collection.
  Dim Item As Boolean
  
  On Error GoTo err_Exists
  
  Item = mcolGroups(sColKey).Allow
  Exists = True
  
  Exit Function
  
err_Exists:
  Exists = False
  
End Function


Private Sub PopulateWFAccessGroup(ctlSourceGrid As SSDBGrid, ExcludeRowNum As Long)
  Dim iLoop As Integer
  Dim jLoop As Integer
  Dim sHiddenGroups As String
  Dim aHiddenGroups As Variant
  Dim varBookMark As Variant
  Dim sCombinedHiddenGroups As String
  Dim fNoGroupsFound As Boolean
  
  ' this function will find all the hidden access (user) groups for Workflow Pending Steps only
  ' and concatenate the list into a string. This will be used in link validation to ensure only one
  ' workflow pending steps is on screen per access (user) group.
  
  fNoGroupsFound = True
  
  sCombinedHiddenGroups = vbTab & ""
  For iLoop = 0 To ctlSourceGrid.Rows - 1
    varBookMark = ctlSourceGrid.AddItemBookmark(iLoop)
    sHiddenGroups = ctlSourceGrid.Columns("HiddenGroups").CellText(varBookMark)
    
    If (ctlSourceGrid.Columns("Element_Type").CellText(varBookMark) = 3 Or ctlSourceGrid.Columns("Element_Type").CellText(varBookMark) = 5) And iLoop <> ExcludeRowNum Then
    
      fNoGroupsFound = False
      
      aHiddenGroups = Split(sHiddenGroups, vbTab)
    
      For jLoop = 1 To mcolGroups.Count
        ' if the hidden group list doesn't contain this security group it's visible (duh) so
        ' change the allow property to false
        
        If InStr(sHiddenGroups, vbTab & Left(mcolGroups(jLoop).GroupName, Len(mcolGroups(jLoop).GroupName) - 1) & vbTab) = 0 Then
          ' if workflow element, update workflow groupname
          If ctlSourceGrid.Columns("Element_Type").CellText(varBookMark) = 3 Then
            mcolGroups(Left(mcolGroups(jLoop).GroupName, Len(mcolGroups(jLoop).GroupName) - 1) & "3").Allow = False
          ElseIf ctlSourceGrid.Columns("Element_Type").CellText(varBookMark) = 5 Then
            mcolGroups(Left(mcolGroups(jLoop).GroupName, Len(mcolGroups(jLoop).GroupName) - 1) & "5").Allow = False
          End If
        End If
      Next
      
    End If
  Next
  
End Sub


Private Function GeneratePreviewHTML() As String

  Dim bOK As Boolean
  Dim iVersionsCount As Integer
  Dim iModulesCount As Integer
  Dim astrModules() As String
  Dim strHTML As String
  Dim intFileNo As Integer
  Dim strFileName As String
  
  Dim ctlGrid As SSDBGrid
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim varBookMark As Variant
  
  Dim sHexColour As String
  Dim fFirstRow As Boolean
  
  ' retrieve the list of links
  iIndex = 0
  Set ctlGrid = CurrentLinkGrid(SSINTLINK_BUTTON)
  If Not ctlGrid Is Nothing Then
    iIndex = ctlGrid.Index
  End If
  Set ctlGrid = Nothing
  
  Set ctlGrid = grdButtonLinks(iIndex)
  
  ' Generate the HTML...
  strFileName = GenerateUniqueName
  
  ' If filename specified already exists then delete it first.
  If Len(Dir(strFileName)) > 0 Then
    Kill strFileName
  End If

  intFileNo = FreeFile
  Open strFileName For Output As intFileNo
  
  ' Start document
  strHTML = "<html><head><title>HR Pro Self-service Intranet</title>" & _
            "<STYLE TYPE='text/css'>" & _
            "<!--" & _
            ".dashelement_width {width:100%;min-width:300px;}" & _
            ".dashelement_minwidth {border-left:300px solid #fff;}" & _
            ".dashelement_content {border:1px;padding:1px;}" & _
            ".dashelement_container {margin-left:-300px;}" & _
            "-->" & _
            "</STYLE>" & _
            "</head>" & vbCrLf & _
            "<body style='margin: 0px; padding: 0px; background-color: white;'>" & vbCrLf & _
            "<table width='100%' height='100%' border='0' cellspacing='0' cellpadding='0'>" & vbCrLf & _
            "  <tr bgcolor='#b0b2f5'><td colspan='3' height='6' style='text-align:right'></td></tr>" & vbCrLf & _
            "  <tr style='height: 39px;'><td width='40' valign='top' style='height: 39px;'>" & vbCrLf & _
            "<img src=data:image/png;base64,R0lGODlhJwAnAMQAAP///7Cy9ezt/bGz9e/v/ff3/ufo/Le49svM+La49uXm/Le59r2+9/39/+Pk+8/Q+d7f+83O+bW39tbX+sDC9/Ly/fT0/rK09eDh+8jJ+L/A97i69r6/99ra+snL+AAAACH5BAAAAAAALAAAAAAnACcAAAWgYCCOZGmeZEJFkOM0aCwHR6YAeI7PvCg9BZ1w14sxBMMkoHjSVJRKJumAhEalg4nVKt0Et1BmBLwtQsjl2cCATscuhLb7NIjLuTH2HX/C7OckCH+AIguDhAMWh3wjWothJIaPkCNVk0kkFJeUAXabQyMMn1cBlqM6PqeYAQ+qoAFfrjk0skIBHrWoN7k7vDkJvjgcwQCCwR3Eer5PwQ3EIQA7' alt='Corner Top' />" & vbCrLf & _
            "</td></tr>"
  Print #intFileNo, strHTML
  
  ' create the holding table
  strHTML = "  <tr height='100%'><td>" & vbCrLf & _
        "<table cellspacing='10' cellpadding='0' align='center' border='0' style='padding-bottom:20px!ie7'>" & vbCrLf & _
        "<td valign='top'>" & vbCrLf & _
        "<div class='dashelement_width'>" & _
        "<div class='dashelement_minwidth'>" & _
        "<div class='dashelement_content'>" & _
        "<div class='dashelement_container'>" & _
        "<table cellspacing='0' cellpadding='5' align='center' rules='none' frame='box' style='width:100%;vertical-align:top;border:3px solid #E9EEF7'>"

  Print #intFileNo, strHTML
  
  ' loop through links, create new table if required
  With ctlGrid
  
  fFirstRow = True
  
  For iLoop = 0 To (.Rows - 1)
    varBookMark = .AddItemBookmark(iLoop)
      If (InStr(1, .Columns("HiddenGroups").CellValue(varBookMark), cboSecurityGroup.List(cboSecurityGroup.ListIndex), vbTextCompare) > 0) _
            Or cboSecurityGroup.List(cboSecurityGroup.ListIndex) = "(All Groups)" Then
        ' not in list - ignore it
      Else
                
        Select Case .Columns("Element_Type").CellText(varBookMark)
          Case "0"  ' Button
            Print #intFileNo, "<tr height='24'>"
            Print #intFileNo, "  <td nowrap='nowrap'>"
            Print #intFileNo, "    <font face='Verdana' color='#333366' style='font-size: 10pt'>"
            Print #intFileNo, Trim(.Columns("Prompt").CellText(varBookMark))
            Print #intFileNo, "    </font>"
            Print #intFileNo, "  </td>"
'            Print #intFileNo, "  <td width='20'></td>"
          If Len(Trim(.Columns("Prompt").CellText(varBookMark))) > 0 Then
            Print #intFileNo, "  <td align='center'  style='background-position:center;background-repeat:no-repeat;width:200px' "
          Else
            Print #intFileNo, "  <td align='center'  style='background-position:center;background-repeat:no-repeat;' "
          End If
            Print #intFileNo, "background='data:image/gif;base64,R0lGODlhyAAYALMAACkpOSkxOTMzRj0/WkpKZ2Fkh4qMw6it7LW197W9972998HG98rO/9jb/+Ln/////ywAAAAAyAAYAAAE/vDJSZ+7LevNu/9gKI5kaZ5oqlVsO12YKs90bd8qoy+uCzsbnXBILBqPyKRyyWw6n9DocLFQ9CowjZDK7Xq/4LB4TC6bz+i0ev1VJBBXS4xRVbgT+Lx+z+/7/4CBgoOEhYaHhnaKbghwPUFVb42TlJWWl5iZmpucnZ6foKGil5IHB48ZdAqjrK2ur7CxspMHBgYtWlWzu7y9vr+Wpr"
            Print #intFileNo, "YtW6vAx8jJypi1tgUsOwuSy9TV1rDNBc8VXMbX3+DhmdkELN3i6Ongtdrl3HXq8fLI7AQDLIvz+vuxwgUDAvDZmcavoEFO/ggIAMDioMOHmvwNAHCvgqmLpiBqLCgRAMMWkLYMYNxIUt1FAwUIePzIImTIiyVjWjuJcqLHK9q0uRSJsafPn0CDCh1KtKjRo0iTKsXoUptNlj1ySp1KtarVq1izat3KtavXr2C32lsJIEAcCQTSql3Ltq3bt3Djyp1Lt67du3jzth0AkCzUsw8ACljot7Dhw4gTK17MuLHjx5AjPwbcQrLly5gza96c+WwEAAA7'>"
            Print #intFileNo, " <font face='Verdana' color='#333366' style='font-size: 10pt'>"
            Print #intFileNo, Trim(.Columns("ButtonText").CellText(varBookMark))
            Print #intFileNo, "    </font>"
            Print #intFileNo, "  </td>"
            Print #intFileNo, "</tr>"
            
          Case "1"  ' Separator
            sHexColour = IIf(.Columns("SeparatorColour").CellText(varBookMark) = "", "#E9EEF7", .Columns("SeparatorColour").CellText(varBookMark))
            
            If .Columns("SeparatorOrientation").CellText(varBookMark) = "1" And Not fFirstRow Then
              ' Column break
              strHTML = "<td height='5px' colspan='3' align='center'" & vbCrLf & _
              "</td>" & vbCrLf & _
              "</div></div></div></div>" & _
              "</table>" & vbCrLf & _
              "</td>" & vbCrLf & _
              "<td style='width:10'></td>" & vbCrLf & _
              "<td valign='top'>" & vbCrLf & _
              "<div class='dashelement_width'>" & _
              "<div class='dashelement_minwidth'>" & _
              "<div class='dashelement_content'>" & _
              "<div class='dashelement_container'>" & _
              "<table cellspacing='0' cellpadding='5' align='center' rules='none' frame='box' style='width:100%;vertical-align:top;border:3px solid " & sHexColour & "'>" & vbCrLf
              Print #intFileNo, strHTML
              Print #intFileNo, "<tr height='24'><td colspan='3' bgcolor='" & sHexColour & "' align='center'>"
              Print #intFileNo, "<font face='Verdana' Color = '#333366' style='font-size: 10pt; font-weight:bold;'>"
              Print #intFileNo, Trim(.Columns("ButtonText").CellText(varBookMark))
              Print #intFileNo, "</font></td></tr><tr height='5'><td colspan='3'></td></tr>"
            ElseIf .Columns("SeparatorOrientation").CellText(varBookMark) = "1" And fFirstRow Then
              ' Column break, first row
              Print #intFileNo, "<tr height='24'><td colspan='3' bgcolor='" & sHexColour & "' align='center'>"
              Print #intFileNo, "<font face='Verdana' Color = '#333366' style='font-size: 10pt; font-weight:bold;'>"
              Print #intFileNo, Trim(.Columns("ButtonText").CellText(varBookMark))
              Print #intFileNo, "</font></td></tr><tr height='5'><td colspan='3'></td></tr>"
            Else
              ' no column break
              If Not fFirstRow Then
                ' First row is not a separator so create the holding table...
                strHTML = "<tr><td height='5px' colspan='3' align='center'>" & vbCrLf & _
                "</td></tr>" & vbCrLf & _
                "</table>" & vbCrLf & _
                "&nbsp;&nbsp;" & _
                "<table cellspacing='0' cellpadding='5' align='center' rules='none' frame='box' style='width:100%;vertical-align:top;border:3px solid " & sHexColour & "'>" & vbCrLf
                Print #intFileNo, strHTML
              End If
              Print #intFileNo, "<tr height='24'><td colspan='3' bgcolor='" & sHexColour & "' align='center'>"
              Print #intFileNo, "<font face='Verdana' Color = '#333366' style='font-size: 10pt; font-weight:bold;'>"
              Print #intFileNo, Trim(.Columns("ButtonText").CellText(varBookMark))
              Print #intFileNo, "</font></td></tr><tr height='5'><td colspan='3'></td></tr>"
            End If
       
          Case "2"  ' Chart
                Print #intFileNo, "<tr><td colspan='3' align='center'>"
                Print #intFileNo, "<font face='Verdana' Color = '#333366' style='font-size: 10pt; font-weight:bold;'>"
                Print #intFileNo, Trim(.Columns("ButtonText").CellText(varBookMark))
                Print #intFileNo, "</font></td></tr>"
                Print #intFileNo, "<tr><td colspan='3' align='center'>"
            Select Case .Columns("ChartType").CellText(varBookMark)
              Case "1", "3", "5", "7", "9", "16"
                ' 2d bar chart image
                Print #intFileNo, "<img src='data:image/gif;base64,R0lGODlh3ACRALMAAAgGAzs3N8Gbg7Kzv8fOhc3Mwa7U8Nvr8/+qjv/mmPzq1f394ffv9/v5+ff//////ywAAAAA3ACRAAAE/vDJSau9OOvNu/9gKI5kaZ5oqq5s675wLM90bd94ru9875+KweBACQ4cvUZh+Gs6PQ3BQCBAPqLUI2+RfXq/FgZjaVUIDlwiTxEwgN9gQkD7MDvSkwN1z+/7/4AFgoMFBAOEg4CKi4yNjo+QjF54dVWUDwdCAVJCnZ6foKGdBKSlAZulqaKrrK2ur7CxrAFNZlNnUlFTdBRSLgupwcIEDXA+tD9KVEQFbmZnF74twMPVxcY8yNgZ0izU1cLX2znaYuZi290r3+Cp4uM3yAwG9PVu2Ooq7O2k7/A18uzVmxDlHocCamTkS7GPn79/MwIKNGiEwYEGB4wcuJhxgAEB/gESwliIomG7hxBjSBQooQGBj4aCUJkzQCaAKQZjkDxhEhzKlC9W2pPA5Z"
                Print #intFileNo, "SUJQI+ClmCVAqDGTtN9LQGFKCEeROvFnCgoADIQyA/zpEDVuSLqCWmDvtZlYXQgRkKzLGiAy0JteHaRtT7wO4IvMHY8kWhrapfEYDdDYZRGOjhEIlLCV5conHKxyAi96PswjJEzB80E+PcwvM/0B5ETyYdwnSOMQeeTuiKkO4E1B1Us17h+kaUU7xAzrWAm4Pu3Sl628DItcrtnBWKbziO/IRy3+qE88o0YBOTdfwCVyf8RIrZOiHzzNxUwFt4xeNNXKehq7aUBYe8npeugXr8EfPN/lDUKQe89Ft6xA3wy3uS/VcZX/xl4J+DIAT4RoQYTEihB8g0sMCHIC6QjoLTMLjZhq21NMyIC5q4GooVdLjiFcIZYJsGDRxy4woYXqAhjBrIKExLQ5jh1RkESEGFAUcu85UeSXkFHQk9WvAjkBgIGQyRaFABwE1KBgDmLkp+JMCXu5w3QpUVXImlBVqm0hInu4AlxFGGvGSIkp3oMSWVJLrn4psdxFlKS0dmtIyUSdaU0aNeDYCF"
                Print #intFileNo, "AQqouWagGSRAYGomjrYBWTt2YMZw8BhKCosdJMXpoKn+6UEQoW6ghCAvJocoIu3hg2kGsK7K4IsJ8CJCUrEG+SUAFloH4a4X/sj1nXGlCHfsKZsIcKIGzloqqxQG1EqBmF8mK9+yBy2zKrjHphuup9iaK4JMwnI4rQ3iesEmBcCgq6667MKBbr0kAOzEvRMsoO++6fb7xr/0OiFHGxRoimB0zKJwgL5UpHIKsgqDwbBVP+DHFR0GUxpvXxWfUEC61p6EiVfUGpVrHtQKoETMM8tllKQTdAccAz6fUsA7H18Bc80zK/tEr5XcAQ3FLqw8rUMNHJxuALL1dfW+AjBgNYL6fj3zx90hjGzWDzphBkWWPP2AztRiesBEdMNVgdTTorIWA2ZPi8DfVve9bwCACx6uGgyfKbjA3zohpQTNmJHRxLcxO3fd/nT7GDiy1UpWNeeB/43A5oYfK8DopSMrwb8HpC6A0j6wwblSB5680OWYs2SlHIYHsFkDFk1h9emoXy3IwdViXHy4Ai"
                Print #intFileNo, "xxMBH/Ir/x1bAbZnnuWe2eJOlf2vwAA4qbTbzyEgzAsujhA3D6vwr0km57/6b+pbcXMP7D7djr3qYw0qqL9cW9K9yx/ga584nufOyjAN7U94D4yU9bhSJXBXCXvz+pBSTpKlsAlwcAAr7NgH9D4LTalweWNXBa8qOcCOznA/xVkB4+4kf60oesPVxNgF/y4AKJF0LTLS8AJCxfBk8YLuPhShAQjKBeXPhCH2WsGunTl4L+FkUcdhABBfTh/gF9mMAJ6At+KHyfBZSQsg90SHRoRBUFKJi/ZuWtDweL4tFuyEEdghABIjSdIDA4rWJEz39H8dmXkrbCq6BRdC0JAkJmk0SoXO+Ff1pg7yRpNsLVEYsf1GIPu8fBvpGIdfIjZIq+d0gPYsFt3OoBEyvoxtT5Dnyls6S+7KhJPHKxdK9b3bw0KDhRVsiQpZzTHNgwBz4B5wDHa4Nw/NSCVbbRAhfjHseI4bUNzhKTO0RjHivpPV2GSwIA5Cb94ATMQ3LpAaMrpgESIIUgFAhm"
                Print #intFileNo, "U0AmwSTgTOzFkBTCidkTT9QkKiAgn8QDqCn5cEiC/rAPA0Db2/aQtEwgdGjkIWUw/q/ABGIKLwBLGJB2zrA2F9Qzd/f8lQVKSdI0tqmkouviF+RB0hS406OPbKL2RFoBlJbUSjb9YRC90CEF+PSnOy1BpYpVgo9iLqTveUhOJ4qvnKqUpxJcIyQt2CmlLtWkTbXpU5/AQlXGlJUzTepIr4rIk2p1hAuLah6mitTwWJWsmMwqWYPqhK6qoCu20c6OjFq3tjpkrHDFKVzp2gS7omB0wjrKAfb61Wfuj1UUgKsH5XpVwv7AsCggQy/cBbUJsjWsbgUsWQU717Q6TlhNoQvcgCNVSPrVZTWVLGkra9omaLZSGHnAJSoHzc8+Fhwx06dTJBvXgknWssdwgsSO/nCfuDG2t64FbTCkuQnZVuAWPLQpcnuAWROMYZEWqUNtotFYe0pXY/LLrk3vdseSbjcbagWnb/HVDmm296YKvG8p37uD7uKAr5k7rylYlojhLZW9tXRvba0HXZ"
                Print #intFileNo, "n+dhj66pdc6EjFfFZ3jVE0qB8KoNDCxhcT8y1YfRPmj3Dm8J9965o3K1nJMvb3w2w074OFEeGHzLCT/lux/IbI1YKF6ENqXGt0ZzxdEo/0hjTcFxj7xr0AjFMlRGmp0XhRqfuVF6QCJkWNR/qvCq8nx5m4miKmd6xGNky3UlZGQoJl5QaDlcjoPVbHfkhF0iEjm2jkI8LKGF6hEjVGUWZqFJZQ/qA96WcHAM4enAcsZ6XekLpZ5GRKDZdLC3R0BDcz8+oCbc45EZoUt3AVDRKtP/oC18jRYZl97wxC+1a6ArooAfg03UBOm5QNkhLOEuzAA1IPJcu8azScVO0/as0rk5K2Zd7IzEAMNMMEbuMAMhYgZQyQ0atudqypq7FlglSRwj+MtPqoCMIZYoCYtKaAS0glbVuXlVcdvoGv7bZtCCdMDHuMo7Jr+M8oituSN/4nxnq8xCsfFdjSq6QVK3mVHXON4E0AHqw3Eqp5wxDhKZx0727jcBPWVW1UoEvtnuvZIdebxrFUL6Ql"
                Print #intFileNo, "ILvBURrisQv5bELChv0YvK/ApmFwh1BS/gtvQkxiqjS1g1tdC6+HtR8/LV3scAnueEckMcbyoqvRYeKKQRzUHqw5JnCOrnsdHVb6sdh//BD/lkCzLG9bQqAks9Y6+OR/jW1gzVraCR7x7r5MgN73zve+673sTgAfNO6TFEOQ3O1vhjtsI2tdyi6VrsjEO64u4PfK8x3wTViuA16yUBXS8+YBnvpaRHvV2T4empKfvAUsz3rMt8XiVIXsBIhr+pxCPvWIoDzrK+96Bpf87SKWvQRoT3fa2h33g9D97vvee8eAXtGK9wnpD1z80x8f+XlfPvPJWfBsyzj6VJH7aKtve9RjP/va33vzL/P8Ugefpoyfu+PLf33k/is//QlY/2fa/2vR50X8pUd+2mV+2Hd/6ad/p8F/9PZ+YgWA1Dd/A1h/uGeA2ud6Xxdk8mVyDBhaDphTtReBa3R+6Id/mDcruacr3id14Dd6HbheAohSt3d+FLh8JXhEc1ITqjIwCnhx/iceLYhSHwiDBG"
                Print #intFileNo, "h/q4d/f8d9JkgIc9I8bRApCNEViLaDsQd/s9d4xjVYQziBRWiENah6OQISSXEmGJVKdSGFryV9P4hfECiEEph6M7h7XXiCOXIRYRhP5hGFKXhwPQgf8Td+a6hgbSh5b9h6SGiDV6AADNAAlRIpm3cyNQB7Zxh+fRiAf0hSMViAW0iChah6GRAF6cYC/pAIbP83iQ94hXUXgjKYiQeIhBvRihuBgSCmgURRVdPngS8IiKiIiRVghEcIaMnwMMFhFDaXhzi3hw2ShtVmisaXi0S4i7yIgDiQBjX3HNxghqLog6Roi5W4X1nohqpYgdwXOwHgNAlRIyK3R82DNlGnhys4ilUof8pofcyohc7IheHYA2xAjhUwjZCzc4inbRsYd9nogtt4SJfYjBHzjPe4BUlRc4aAH+gxjL+XeAG5eO/oh/FIf/PojfWoib74A8TkBoZwIGoSisZ4LRdJiRkJgnkggoNoedA4DibZjtiYkqXobhrZkqnYkav4kb73j99XkWg4kEB4i5bYjYL4"
                Print #intFileNo, "/o00uJDsR4yhR5N8aJPauJJsuJFJyZPg6JPO55TQJ5SSKJUESZW4qJO6mJD2qJVNOZEAOYvCh05WiJMsCU4uqZRwyJT7x5Xux5ZUOHxviWZYGIh495K8Z5cJiJf9B5XHSJRqKJZHCZh3J5h+F5OwuI7FiJgoyZfwCJdVSZYIOQG8mH+ECQ8z6ZUsqJjJqJljKZc7aZYe2ThXUHmTGWJ62YCmyVSMyY2OaYhYuZS+2ACwyXJCcB65xSuf6AGjOZscWJudVpBodJD0yJo96Zq+6XdLmIMRw24UI2pFZY0nOWd9mXWnyJnP6ZkK2Zu/yQVDoDMXJQQWtVpKIQXAYQgo/nCcukWLyGibqNmYVhmYdEmI5kmdE/AwwJEU7FQTwIEAA4CgSLFO2iGfJ0CfGkJ8zCk6zsmR0JmV0vmb4AMcAyo0yWMtCXoIFwVPDgpt3GmZ3pmZfhmeqlmW5HmWGQqgGGCdKDAqOQEzJXmipOmOmImR+Ymb+/mY/QmTSDh2GlBlKVAfJFN4tqOjyC"
                Print #intFileNo, "mQYFmUE/o3FXqVF8qbrimO+tg0u/V5hrmAT2qRPaqSP2qQSMmfu1mXaKkD+UgJTOc2q0V0cjqndFqndnqneJqnerqnfGqnAFALHMVRkzOoCUIopRJxwOgAuaBXhWqo22B2mLYRTxFeGLFY5OWoj/phmApVDHrxipsKB75kAxEAAAA7' />"
              Case "0", "2", "4", "6", "8"
                ' 3d Bar chart image
                Print #intFileNo, "<img src='data:image/gif;base64,R0lGODlh3ACZALMAAC42Kz5KTlZaRYRrS4yLiKbCqs25l9bX1PnAa/e2hvmcn/Xo1fD1+vv4+f//8////ywAAAAA3ACZAAAE/rC1JyWd+Ompu/9gKI5kaZ5oqq5sa1YZXFGXa994ru88GtObH6dHLBqPyNUMM2FsOsOkdEqttmJDxjBj7Xq/XkujcMAQMIdBGcxuu3OzAuBAIRTuDUFaSzgAzgIEfQ93AkxRbzeIIz9AjiQZTiIyFokXEmoNdmMEDAESBwVnog0AmndnBwIakkGLlpNiJ7I1MyFLlQwHu7y8AgDAAIZcbTKZm6QAFACkB2emDQGiD6oVzoLY2drb2QtgS7OUQVogC9zn2WUTC8OP3zN9p5yly6TT0Mypww2hBNJ3AAMKHFhAgMGDCBMqXMiwoTof4DAUDLCQIMBKHmA8WPDpUo5g/spYVGAgoIC5PoJK7SrZJ9muAM5"
                Print #intFileNo, "0Gaqhyl0KBgty6tzJs6fPnz7TgBxKFGQAo0UB5NSCaCQIWw1IAnm1AqRIDqoCaTJYII9BbwVLPhhgsAyZAWdg1KwBqwRGqhkpODG1LCldDrLgeuB46GMwRVBsKumwtm3gWE9ePDFl9y+jxxrYPXkVhQODVkOsugVTGNYWjB0cOACdgugHxzc8KZbbq5cHzUwvffbSGVbr260X6A0xVETI1AFU6ML9wAkDkKHoWNtVi/Yqww3ISkdHYICB3SB6myia3TFqDZo7IPcTbA0DOcECdZIwlJaV2omiHxygoL79+wSuVw0/wu7pxg9w/ideY1qEAhIZfSSV2HvPeSafQfTdJ2F+2P3HH2/+DdiYgBNsaIozRpVkF15dwOcGJhBGKKF9FO5nmm+/gYeadgtoF6BjpYxHHjDFEXDgjgAwMMAAVlXYg4knkiXAkCviZwALygBYQnggVVDjXxxoliMwvbR3JZdDnYQeMEbygGQxSjLZZH35QSmjgt29eCOZBVaZJY4"
                Print #intFileNo, "67mKnVXZaUKVzhlGQpopNtlnVa3C+KeCcIR1XpKIbgHTZpH9W6WGMVZzJxoNqrmmoCphq6J2UjFbQ3p08QspKqVsy5qFgSGj6zaBrsvnkof2Nqutvf7YKjaqMiiZqXTxeWuYOsn7BKaEr/ubnAK4wFvsdlcH4aSewIAkLaavcPrJbX7ZMBZVsXFTTwyEaJRYuFLTW+mlpoQ7LH7WpMhojvZhqxud3KVhQXUn8GEQOAQALmUlBaV1i7rGQSCDawxBHvMiytSrwLgoXQjrvjBxP2/G9HzPqW8N5kOHNcKvYIQpJK2lixxnkBBybDg5cUx111bWyQbue3grqhfoGHbKiINe7sdG7Fg1JHWQU90kp0ZVhBz7WzMTPHdLofINoWRmkskWdfECxuz7DuyGiGYpMNNq8kqq2vXLeMoEo8QQH9TFySABTV9VIIAqRg+jQQM0IfU3QAFqP3bObbsNtI75Ft532247zK/YE/nqYc0BHpuCNDwN9rIVVO4ITfpDhAyE"
                Print #intFileNo, "uNs+Flo2xqJbb+zbkbMtrOdBxQ6KGj7v0kfIdzAQgiDJ2OJNYNVrbwHXhFt2hekasN+s6ZAFSEW+01xuBF8EFUFCQJClRRkzfDI9gutfNF/A8uylWfHEsfhNgvdnZH6ERJeveLwkGoHdPmgvnC0T61rez9pEtHBoIQPKMUD8LNZAIUamEBIyDhZg1xxGkWOAVuibA5hFQUAZc3CwSSA4pPHBYVsAFVCQ4rvxNo3wj4CDqBPJBxbVOBdHQIBFOCCwq4O9+EhRXXryXFhiGQIYDTFz0JjQ9xCjQiKAKFBKiMA0eBOx0SVxd/ghvqAIFzk+KU6jiDgI4w4DUcIlOWoEXp8BDMN5AjDRDogeVuEXpqVGHPWijG10AxxyQMYvQqyMT7/jFPR6hjzj44xy1uCRmDbKLeOSBHg25AkRuTY4WOaMg0wjJQrqxEcSIC3aqCEUPKDKTdGyk+5pIgjViKFGKoYchsYC/DjAlXba8QEp4cMrDpbJTXEyBK+MES0ZEpXo"
                Print #intFileNo, "9fEO6huiOcIXrPEWMI/MWGUhVHrCT2JvcUywTnGSeiH+C8IYzusIJdThDOSbxnxhIwUtM+pKRwLQjNl8ZuXrdggaSm2QSmICNBVwDGnQ7wEkGwIn+LYGdY3Rn6n7pSE4KU4cZm5Mx/sWDQgcd5zL8kF8aBlG8VRzAoOyMCujWk9BpohKeDWURK0cwTAfSk23VUo2UFjUFzAXgKL0LGN6QJ7zuhdQfwJiJNLFITfZBaJWEXJA3Z5cUPVROWkV5mBQkoIdy6kMND+ioJv7g02gSQn7tNOk7q7kkpB4mCph5YhC0FVHtNIYOAGoMxGTjEY/IAFZQCNgpjEcRIu0VADhNQwCwGgewlpSoJyVrSm31FKmK7YmzoVztbuTSZJKJFUDihWjudzlnwgovfugK6OagBSeM5DIbuKUHLKk8hdIwlWW9JgYittnBiSYACxgNW5Yq2cn2dgMOWMCYmEGGh/FjF8NZQ1R4/hFJV0QEXVyoZSVYC0DXmhG2i7XYkzBggO5"
                Print #intFileNo, "697veFUBudZuYiEo2UeZNyh0OsFlnoCUr62mAOZakHLfclQm1ABcMVPvVHvRyoYyM7eJsO6QhLaTA/dxtej8mp4w1Zr2iMUdu5Yugj1KgZgWImV5IE91lOtMW1G3Bf18b4OwaSjRKmk+BB5CAFneXABP2bWXPZtkHFxfD8y3ATbuSWjLgt5QtCDELRnzdEq9yAihGiHVanAAEOBkBLR7AeGUMu7Ux1Z72Eu6BTBZcQYhCFOfcQCgSIeQVEBkgmjwq2ZDsgBQvyQBNfvKTEyBlbS1VaJGraLCCO9z1jvckobuaNZob/kbDDhV9RS2gNUU4ODcPQM6QpvOU9YzCo+WTbYPTcjAg/KwLWwZdliizCs7sPK0tIMUm1s8D2qxkSMtZ0p2m9JXz3EOriMZRwKhtRKCgrf95QdQpILX6tEaw+biPDhpg9XxcPec6V5Zo+QzJ4zxWLQcMZTScxYttsQ1kHQAbBcL+YLEF7KkhKBtCzHYyrB2oTRpX7qm5xra3gLt"
                Print #intFileNo, "ZvN7zCTrDBQ1CCWJD+9G6aM73uLNrMXMXGN3pljQxaQq7uHHH3asZzHMx8+H99tu/AC91RgZesYKH5uBLSreTpbxw864Nd3iOUuwgKAbo3neCL9fAt08QboEb230GN3DIRU7y/iM4dZYt33UQ9k3LDszcBDXf+M3dlfM0iRwBPTfCzz95CV0UCLnH/WhUNvGAX5C8sBgXK4ChwHGcf1znj+a5N3yObKqjASaa6GkpvNy/TXSCFNP1dyIzPmylq5npZ3e62pEw9U9OEBgVrkaQzjkMZtAB71sJareTzXcCbuXv5Q48wtMd9SIUHoyHL4szZPY0PQxD7qT0pyDWcOgOJnYDZQd8skGeds6vXeptB70ECDaWlawkQPxQwyo8YY9oziOsiB0r7Jee+dmj/emdJ8LnA9WX0N3UIHPPGja6kuFdyAChrS9jwP1O7kI1ffPMjn4Ppk99MWi9F1k3jo/1PUHI/ocfkMvHvPk1v3PbEz73bkd0AphXdTUZLjN5lCd"
                Print #intFileNo, "2JEZ2zLd/HUB70Hd7ngeAUpQ/gfEW+3ZPpHR/iXZ55dcsUbBi/edqCqdUOsB+lORtendJClhkDKh/IOgB0oF+riZl+AUHFoCCKYgDR1cCSfeCHzghIdhIQyJykoZLijA3Z7CDRdCDJPCD+ReE+GFwS9JIRmiDJugCZsBjTNgDTmg+lWdzMCiEH1eFIxhpX4eAHcB1XcgDXygCUFgHDRiDs9dIZwhpBuAA+yM4E8CGbUg9JvCGIRCHHkhwS1iHKxZnzJaHuwUHWdU9f+hhoFZLGwB+/9aC4weEhkiF06GIeMiI/kZgB5E4FRn4Q0NnAS80eYQYe823ageHFp4oZ90Va0Qgin84DhfUF80xQVh1WIj2enI4hlMYeENCALH4ZLN4BLbYhiTyXEF3GILCepeYfGMXhZtIjLCYbsmYbD2wjMxYHOJAS9VBTpWQCb7oesoXjFLIIk1XjCK3javWjZDIhOgiPOzBDGPQFQT1MmEjBubIga83OKzogK5IhATwjqB"
                Print #intFileNo, "YBN5IddTwDI94aszRdcKVYcogXvxADqGAMMiHPt11OK3hAMUWT3RYkESIkGr4AQtJPU0BcyVEdLoYF5uxDKfwVSGhB+RkChapBk5AN8Jgb8GGRAZAN9jgZn1lAKjWcYdI/pIFZpI+2QIpCQm69kMxhwVio5KC8IicYDd6sAlBAhO9pxG9yIHF5iEqhpRUWJLaqB8nqQFPaUxZ1wu6QYngIkqTcAlyIDVdIQh6sDnC5Ttj0ZfRpQeERnNyNG5EYYZ2uI5scpZLmZahtANtWZUXsAAGsCLqAArIhXUY4E+mZW9qQVyowD1eIQA7uQrOiAnSuHfTNG6NJAhokTqK6XF1iJaLqJYKOY8vIAEL0CSXmQaBAF9McWolYUGPMUHbxwmO8IyXQwP/OI2IxpoB0ZECIYzsWIa06Wrd9Zg6EJnbJF+8mWnxgBMVNjgFYADChWxQ1BQqEJbOGQgvhhDRGZTT/hmbSalsjVmb2pkD3JkRVsKbXIMgJUERGRYZv/Bznek"
                Print #intFileNo, "F7Kmap/OeBxGfA0Gdi2md94mdtlmLuGlf3mmZxokKYNYHWoBhA9qUR5CgLLigBgCd8imf83mNszmhn5ifOLCf6pKhEtKb2OB9ocAUqzeYSECirVU4DMoVKkoQECqbSlmEjimiKyCj4kKZEmIAlxmO2vYsPDqiFPijWHSi8DmkABFe9MmYSIqfSqoCTOpcUllXK0SAvkYFPlpdzKOlDWoRXnqNo2GfYUqhMHoDZfoE3FYBfUqKKrQpjtCmscRrhQmfctpdRbqEBJYitUehz1IZ23mhI3RvZbKWbqkzhDpC/qalSNBZntK5oh2nlvKFao/6iZ0Wczawp7coNw94pT7gCJ6KqKEqqrVyomgmDNeJnf34Y2uKoX7YqhARLptqX0Awq3GaPhBqAL8QDBThqAlnAAAwAOrEWavKhcIaDrhUrCSjFquJqMoam856fUTIYurGZC1GMIAFd2xhRE4QrNlKArsQl+UIqyeAE60RFs8Jrs1TpM0KDOS6JMboYitmYDd"
                Print #intFileNo, "1FMKgDi3JAnuAqYFCFmHDLqkJk1/FoXRzU0Pxrcnar4pplBg7rg1hEAB7sB8rilSxYdQwAGNKj02gCU5AWAU0Eq84H0RBsjZrs9QqCvxqEfp3osVIFiQLssIgDDdb/rQf+3NAlIWEobLxWpcVIF5fmUAIixTOCljrSrICcLNNpbMby7N9VbAKYbPjarRkG7TCgFFBpxdp0LR1eQFQ+4+YMLJaS7YG8bEAe7UHqxDNI59ZexBla7Rj+7eAy1UPgYQfsLZsu000IJSFKx9yK7jXRxFFew7gqqLfpWNWC7lFaxSau7nAMAiG6wGIm7j3lI/KMQTS8biCm7U3NUP5iqKherlH0blky7m0e7CZS1B5OrqkW5xbALFEgrd/y7rSsAtkcBsAGqex610Tobq3i7vC+7wIGw8xxLS9uwIQOzydW7fFixu98KkE0ZGsG7jSO7XRe7ufC2P4s7YOu4PZ/jsAswu53Hu83tsPz1oS6TMRdlu+Zuu8zxtUjGoNunu966k"
                Print #intFileNo, "eghC/w8u69Ou9okCuAjFwBtu//CvBEzy1AwoKA0zAKAC82IDAdKvA9csLQHW/CDG0GDu45/u//qvCTWMuGrzBBowWI1y28xvCvUPBuCu45Fu+O8zDv0lQLwzDrpkNHnyzNWzDPlLB0JvC6LvCtAsSQHyLoLEFliqDBqwNRYy1IIwbx9vATKzCSmy3wnspADspvMuy+VWA/JkX8oEOWUyuO2bDzvCxYTy1YVwUJIwQC9CZ6XLGOzgSHeZrT1vCV7wNb3zEIewPdHzHTqy1SYENf7B6pXiD4JCywso//pMyWkmBsLchPOgAv9GLyL0AEPYLvXVsx7NrFwvsGvLVERFEcWk6BH78x25Lskm5xu0aFZF8DsFrtCXBwMeryEvcxDU7u6wLMLksPqqRP0FnWuVivaMoROKjCWeVvdogw7/gy/RLBkJJxGK7wkkhw9qgN7ckBI0gNp5wV3RlrVgBzfHqHtEhC6OHEULSmtyQzZOrDcSrw49MEDrUUiygGko7b6I"
                Print #intFileNo, "bxebzFMlGiwSYV+acgTiYWoAFiTDn0G2MDvjMz1Z7t0VBmpk8Kb8KAgC9AgJtA0/jgwptAXa2BNw2GclZyVfQjC6XbZAcCOKcDWTxuHiMdkqSPC6sbUllAyMt/hLIVKmr5qe05aeVTH8rq7igFgQclBDZ8Ac1bdP/GtGhWTyBfETt0HJjGtLC0U1aSBdmxhMtlhNHPdFGrc42EBVLYb5WqwbeS3bCU8hR7Rol5KtqTBjDsLCN2EpVGgJBrQRQc6kRFosJYNb+VKATEV/RkTVZG6JaKMyYQS4XhJKefM1oAQ2xzEL56cJp/FA4ENg4NNj9UtiRlluX4aHieU4M4ADl2ZfcvEBCN6O+SsVcXQehg09DfDOuqdkO7bRa3W1e7QoUywp2c2/QGAQlHQ6m/WrjlaNgJkCSkAavqR6wDI7YDdPEHV14oAm905DF4yOm5QdDqQ36FEMNcgPD/n2aULDMFjDZtq0RymAkbLYAhr0TofMyK/OhQXl3gnDXZxoYH1U"
                Print #intFileNo, "G+3AKJHEHF7CP3YNRiCfLg1Xew2Mm6U3SeEQLYsAUy5xXe/iWxguwplkmhd1ic2bWhCMMeLAyIPRRFJF7IwG6zOE33YMSfmOc4bjO2O0RBYEWNX3eIpAsJ+DVZ7oE6bxf3FCwaEGJNBdcnljW4yWll6Bb/MaLurp6z2CyF0COTUDZ92WAXpG+US3hOjDc5SyVNaK3A5Fak7wuS6PkccZkZu0w6rxCQeQKYDbgOMEP8sXQk3gI8+pP/lS1gMUL07rjYJ4D6y1QOOFPfmC+MHEbLxnTA12V9b0T/hCD1M9YcdhRGcTAxzAuEFhLER5QWhcAX5A8GD1umjigVvk4n8dsNcXtmSytpFSc0N2JGOccq6c5UufgklmNBl4zJPMtNyo0xRNQE7Ot1mTnI1M9zh8dxP0i552+khZHSzFgzb8C3LawhxrQNwOIzkU7EIV7AQfq7GtddLtl7qBwXKDAMjHBCr8u1p1FLngNCsOAGxp9U1IJ3y7dvmw7QbwedHjHY/K"
                Print #intFileNo, "FEnPgMnhDAySxJNfeWVwNFVHtmgULVsy0xkE0AzPD79k67wgyBmYRyWNACKTpfT52HqN+usH3C4KcFwmfEASh1i+5nPlFDM1O7omeE9mtBXvsoU2QBjGuvieAFTO6QCJ6zsbqagE9ASTrihuS2tfOZesX2PDkzgg+0qzy4AwccQbmsO1nwKjeAs9Elzz5iHiFoBBqHvUpyB7z4H1jkBNYDl1mOueUjDPZwMbjbvZnDwppnLYYOHRyQcldl7kRrfTa9mPqYvd/HEFoKkTrfOEiXBSXWc6KC6ifhvGGfy7PODN0Je/+Xh3UOi7iIhhvUeO/Xfm6F8jPdYpAlLQVj1ezrl8zT/qwH/uyP/u0X/u2f/u4D/sRAAAAOw==' />"
                
              Case Else
                Print #intFileNo, "<img src='data:image/gif;base64,R0lGODlh3ACoALMAAF9iT5Kxxtalf9PCvbjjlcHnq87p7O337vmlhfucmP+koPzfzP/08fv99/f//////ywAAAAA3ACoAAAE/vDJSau9OOvNu/9gKI5kaZ5oqq5s675wLM90bd94ru987//AoHBILBqPyKRyyWw6n9CodPpgKK7XgWHLNVC/zgV2nCibB4F0uts9gN89K9ZMr5fR6ryeDe+/GoB2goMJeHqHhwQFgA1+jiMMVoSTdYaIlwEEmooHB41TDqGio48YkZKUqXeYrJuuBp1RDqwOpRMNC7mqu6usiAWuwQSxTrOYtQ+4A8tuEwsGz0+4vNSFvojC2cRLxpfIDAIAAAFezgEIAWEL1dSW12nZ8QvNSd2I3+HkDQMC5M/oTHSxa/cuD7B42QoUoGfE3iF84wIcCDcOHcAkuRQMrObuGsKP/goZDnGo5xsaCQv6VTyX7kjGjew6+"
                Print #intFileNo, "vpIU+GnkbQkMDj5YIG4lReLiIE5UGYrmjVvBiGZB1mDA+V2LoNGtYgcokULwkOKVCHOY1RQYY2pNRNXrgVGjlo7RexYslrPyk1ryweDt0SNXpI7V2ldG5Hw5o3Lt+/fG24Fw31XmO+iwzQSK+ZIuPHZhX4hr5A8mSBjy4Ufa2bBuTMvvXpANxY9GkVp07tQ51G9OnPrEK9hq5Kd5iBtx7dL5NadirfZ36GDixhOnJJx5JbpKu/AvPmk59BXT+fQQKN1mLx9Z5+7fcPd7+A9jtde/sJ59BstRUy9vnZ7Cu/hZ00j7pD4+obd/ledfoPI5x+A0dl2m3cE7sffgQg2tkB7AzZoh4H0RRiaSK0NZaGDPxmkYXQcarbOdwgIIABWqP034lklHuZhcykCsOJgmLj4Ylcx2jJjcxTdmB4mO1rWYyknWhfOkizmWORqR/rxI3EI/ARAk5fo+CRSCjqSJHpM4gjhloV12ceXH8aXJZmWyZjmWEZpyeZHZoJxgCZ"
                Print #intFileNo, "vDjnmnHLVhcsmQuZ52l58FhblFHcGo6KgsRFaKF+lNKDlooxe5+ijMD6SKEKUVnrhL5gC54ekXAXqaS8ZhpppHw3I1empMqkqKhxycnqqNYnIepkfqiHgaUe6zgrGb5UCG+yub2QnqLHH/qKV7HhvutMsecPW96G00zr7xabWEohttkl9MaKpzX0LLkjivviqbpacq60U3I5L7mTtulsTFbUCOK9g9dqL0EJT5KvvvnCqIbC9E0px8MCdGeIvUglHsXCEBDv4ME0RPzHxuHjhcTFShxax8YtvefzxRyFT0CpCKZsw8o5YnoxyCCvH03IJLxc5pMws0zwzBgoEgIzL52bFc88g1KzNBAx0coADC+jj9NAh5EwmZUfb7DPS4PwEjT9ej3CA1Wz6uttxWbtyczI/gxPAAlosA3XYIpA957qCoJG2MGsrzbdO/cC9xgPoBABAOVXLXHEhewvDWgd+B9PMTgbslIYb/j6JgzgIdheK9xmN//1B5GpbsI8BV"
                Print #intFileNo, "OOc9ecDhC550j/H0Hmo67r+OghO5+50nSjMLquQtpfuhO+6rhj8Jmv7QPzxzSbfw/LMB+s8D9BHL+v0O1RvfajY66D99o927z34oYs/Pvl7m5/D9+hvqT4O7IduZfgSk2nlT+fOX+j7N8QvzP33y5b++MQ/G/hvEwAEoADxt7/6vSiArmBgswY4pwLW4IAEoGAEFygO+kGBVCPSYDwS+D8GSlATGtRfCk+IQhGO53HDeyALs5HADkawhh0kIQJxCIAb2rCFCkRQxjQmwxkGg4c+5GEQgYjDHf6wiUIMWBFd2MIjShCCGRyg/ha3+EQTXtGIyLHgDXaERJpwsYdORCMT05hENM4Pi1n8YXakI4UildGKS6SgHr8oxzi6EX8r7CNy6BiFLUGRiXk84R672EcVAlKJapzjFxYwpyAqkY2YXKMmMflGSK4HYOmqpBdHqcZFRvKMePxjDj35wmeJcpVyRGUm/ThLWtKSivUhJLw0FMlU2nKTpkxiLR0Jy14"
                Print #intFileNo, "iCJTVQpAROxlLPp5Skc4U5hqpaEzV6HIKvGQhMW2IxWBm0pu2hCMYo9OHeI2HlYhMZDPXWco7/vKS0IEhFUiXHXf6UYe/zOcZ8anPQ/7mmlOg5zmX2MZ3QpOdbBRhIHFJTj+YM3oMLRRA/ueJwVBFlE8TpcJDj3fROdlibOTrKJkyuq2QjrNQfwFp+441RE2t9FiQAeFLuWeimaqKd2+QKUSd6EHN6DR4NlySIIskxiRUlE9opEg1iaqc+NlTmZpQKtqeVFQlOPWpcmHoKilSwe2wD5JLNeM4Q6TUrpbne6YcqljVqquqOkGZy2QrQkSK0fugBEAuTCs7LUnQLSGzBbjIhWAHi9MbUHI9eXVmN5vozyL9FbCTYIBQEBtXUvK1hveE447cmozISmAf/UldDjZqGWp60ZebrKI0NzuDBni2Cv3ox+Z08JR6ghGXsvwIXWnD2c++1nIpCQBFDEAR4fZnBqQtjGkb/onP5Wo2QjVwLSEkC9v+TKQ/VTrcc"
                Print #intFileNo, "GmQ3KzeVrGLFSRW61PYE0h3ENS13NPQQFyJXPcAwRWtC7rLFeemM7y9hKeGymve3/Ikc+MobnFtQN+1VhOCCI5mQfMJIP72d7pLOCxoEsvNg7YTobVcT28rcF5BUFcJC7AahVX5TEaW+MTH3DCHx8BiBXw4wqWdYW4NiuHUxlPFMmqMXtNI4gufuMefxDFkJHwWsKayuSTU7yCFXFPvNva+eu2hkmnzWLumAL5FxuxcmZnfBD9SrnLZhpVZ8CebMnk6ZV7pmbfzlKOGb83tAYSbyeSJMWePeXC28wTmrKE867kCCjlaAVr6VOceSIrPVCZpoXcwNkSHRtGL7kGjp6UQSEcaCJMOVaX9fGkSwJfITxr0PDptp04cMHekdsRTTD2eeYg51R/VndNArTZZd8LBsH7Eqm2d61772gQRAAA7' />"
                
            End Select
                
                Print #intFileNo, "</td></tr>"
            
          Case "3"  ' Pending workflow steps
            Print #intFileNo, "<tr style='height:151px'><td colspan='3' style='text-align:center'><font face='Verdana' color='#333366' style='font-size: 10pt; font-weight:bold;'>"
            Print #intFileNo, "Pending workflow steps</font></td></tr>"
        
          Case "4"  ' DB Value
            Print #intFileNo, "<tr height='24'><td align='left' nowrap='nowrap'><font face='Verdana' color='#333366' style='font-size: 10pt; font-weight:bold; '>"
            Print #intFileNo, Trim(.Columns("ButtonText").CellText(varBookMark))
            Print #intFileNo, "</font></td ><td align='center' width='200px' nowrap='nowrap'><font face='Verdana' color='gray' style='font-size: 10pt; font-weight:bold; font-style: italic'>"
            Print #intFileNo, "sample data"
            Print #intFileNo, "</font></td></tr>"
        
          Case "5"  ' Today's events
            Print #intFileNo, "<tr style='height:151px'><td colspan='3' style='text-align:center'><font face='Verdana' color='#333366' style='font-size: 10pt; font-weight:bold;'>"
            Print #intFileNo, "Today's Events</font></td></tr>"
        
          Case Else
            ' uh oh.
        
        End Select
        
        fFirstRow = False
      
      End If
  Next
  End With
  
  ' close table
  strHTML = "</td></table>" & vbCrLf & _
          "</div></div></div></div>" & _
          "</td></table></tr>" & vbCrLf & _
          "" & vbCrLf
  Print #intFileNo, strHTML
  
  ' Bottom of the document
  strHTML = "<tr ><td valign='bottom'>" & vbCrLf & _
      "<img src=data:image/png;base64,R0lGODlhJwAnAMQAAP///7Cy9ezt/bGz9e/v/ff3/ufo/Le49svM+La49uXm/Le59r2+9/39/+Pk+8/Q+d7f+83O+bW39tbX+sDC9/Ly/fT0/rK09eDh+8jJ+L/A97i69r6/99ra+snL+AAAACH5BAAAAAAALAAAAAAnACcAAAWlYAOMZGmeaDo6auum7Cu/0Gyr0a2b1O4Did8uINQFFEVbIJOcBQ5NWSBQiLqmD2trKtGqpgGBFwVmjE/gQOVcSmvYpHQYDpBD4fLABC8fVM95ARtsgQE5Y4UBNV6JAQaMiQMEWo0BF5NRlQEDj02aUxienwEISaNTCxZCp1MDez6sYAtiRrFgFJhOtmkMtC+7eRIPf1/AgQceSGTGjQkcCB0GFQ0hADs='/>" & _
      "</td></tr>" & _
      "<tr bgcolor='#b0b2f5'><td colspan='3' height='6'></td></tr>" & vbCrLf
  Print #intFileNo, strHTML
  
  strHTML = "    </table></body></html>"
  Print #intFileNo, strHTML

  ' Close the final output file
  Close #intFileNo

  GeneratePreviewHTML = strFileName

End Function


Private Function GenerateUniqueName() As String

  Dim strFileName As String

  strFileName = App.Path & "\" & Replace(gsUserName, " ", "") & "dat_PreviewSSIDash.htm"
  GenerateUniqueName = strFileName

End Function

Private Function DisplayInBrowser() As Boolean

  Dim IE As SHDocVw.InternetExplorer
  Dim dblWait As Double
  Dim dblWait2 As Double
  Dim blnOK As Boolean
  Dim strTempFileName As String
  Dim strFileName As String

  On Error GoTo LocalErr

  strFileName = GenerateUniqueName

  blnOK = True
  dblWait = Timer + 10

  Set IE = New SHDocVw.InternetExplorer
  
'  If mblnSave Then
    IE.Navigate strFileName
    Do While IE.Busy
      DoEvents
    Loop
  
  If blnOK Then
    IE.Visible = True
  End If
  Set IE = Nothing

  'MH20070301 Fault 12001
  If strFileName <> vbNullString Then
    If Dir(strFileName) <> vbNullString Then
      'Kill strFileName
    End If
  End If

  DisplayInBrowser = blnOK

Exit Function

LocalErr:
  dblWait2 = Timer + 2
  Do While dblWait2 > Timer
    DoEvents
  Loop

  If dblWait > Timer Then
    Err.Clear
    On Error GoTo LocalErr
'    GoTo RetryDisplay
  End If
  DisplayInBrowser = False
  Set IE = Nothing

End Function

'Private Function IsRowInSecurityGroup(ctlGrid As SSDBGrid, iRowNumber As Integer, sGroup As String) As Boolean
'
'  If cboSecurityGroup.ListIndex >= 0 Then
'      If (InStr(1, grdButtonLinks(Index).Columns("HiddenGroups").CellValue(Bookmark), cboSecurityGroup.List(cboSecurityGroup.ListIndex), vbTextCompare) > 0) _
'            Or cboSecurityGroup.List(cboSecurityGroup.ListIndex) = "(All Groups)" Then
'        grdButtonLinks(Index).Columns(0).CellStyleSet "ssDisabled"
'        grdButtonLinks(Index).Columns(1).CellStyleSet "ssDisabled"
'      Else
'        grdButtonLinks(Index).Columns(0).CellStyleSet "ssEnabled"
'        grdButtonLinks(Index).Columns(1).CellStyleSet "ssEnabled"
'
'      End If
'  End If
'
'End Function

