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
            Col.Count       =   27
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
            Columns.Count   =   27
            Columns(0).Width=   6324
            Columns(0).Caption=   "Prompt"
            Columns(0).Name =   "Prompt"
            Columns(0).CaptionAlignment=   0
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   11
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   4498
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
  Dim varBookmark As Variant
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
  Dim sChartViewID As Integer
  Dim sChartTableID As Integer
  Dim sChartColumnID As Integer
  Dim sChartFilterID As Integer
  Dim sChartAggregateType As Integer
  Dim fChartShowValues As Boolean

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
          varBookmark = .AddItemBookmark(iLoop)
    
          Select Case piLinkType
            Case SSINTLINK_HYPERTEXT
              sPrompt = ""
              sText = .Columns("Text").CellText(varBookmark)
              sScreenID = ""
              sPageTitle = ""
              sURL = .Columns("URL").CellText(varBookmark)
              sStartMode = ""
              sUtilityType = .Columns("UtilityType").CellText(varBookmark)
              sUtilityID = .Columns("UtilityID").CellText(varBookmark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookmark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookmark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookmark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookmark)
              sAppParameters = .Columns("AppParameters").CellText(varBookmark)
              sElement_Type = .Columns("Element_Type").CellValue(varBookmark)

            Case SSINTLINK_BUTTON
              sPrompt = .Columns("Prompt").CellText(varBookmark)
              sText = .Columns("ButtonText").CellText(varBookmark)
              sScreenID = .Columns("HRProScreenID").CellText(varBookmark)
              sPageTitle = .Columns("PageTitle").CellText(varBookmark)
              sURL = .Columns("URL").CellText(varBookmark)
              sStartMode = .Columns("startMode").CellText(varBookmark)
              sUtilityType = .Columns("UtilityType").CellText(varBookmark)
              sUtilityID = .Columns("UtilityID").CellText(varBookmark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookmark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookmark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookmark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookmark)
              sAppParameters = .Columns("AppParameters").CellText(varBookmark)
              sElement_Type = .Columns("Element_Type").CellText(varBookmark)
              sSeparatorOrientation = .Columns("SeparatorOrientation").CellText(varBookmark)
              sPictureID = .Columns("PictureID").CellText(varBookmark)
              fChartShowLegend = .Columns("ChartShowLegend").CellText(varBookmark)
              sChartType = .Columns("ChartType").CellText(varBookmark)
              fChartShowGrid = .Columns("ChartShowGrid").CellText(varBookmark)
              fChartStackSeries = .Columns("ChartStackSeries").CellText(varBookmark)
              sChartViewID = val(.Columns("ChartViewID").CellText(varBookmark))
              sChartTableID = val(.Columns("ChartTableID").CellText(varBookmark))
              sChartColumnID = val(.Columns("ChartColumnID").CellText(varBookmark))
              sChartFilterID = val(.Columns("ChartFilterID").CellText(varBookmark))
              sChartAggregateType = val(.Columns("ChartAggregateType").CellText(varBookmark))
              fChartShowValues = val(.Columns("ChartShowValues").CellText(varBookmark))
             
            Case SSINTLINK_DROPDOWNLIST
              sPrompt = ""
              sText = .Columns("Text").CellText(varBookmark)
              sScreenID = .Columns("HRProScreenID").CellText(varBookmark)
              sPageTitle = .Columns("PageTitle").CellText(varBookmark)
              sURL = .Columns("URL").CellText(varBookmark)
              sStartMode = .Columns("startMode").CellText(varBookmark)
              sUtilityType = .Columns("UtilityType").CellText(varBookmark)
              sUtilityID = .Columns("UtilityID").CellText(varBookmark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookmark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookmark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookmark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookmark)
              sAppParameters = .Columns("AppParameters").CellText(varBookmark)
            
            Case SSINTLINK_DOCUMENT
              sPrompt = ""
              sText = .Columns("Text").CellText(varBookmark)
              sScreenID = .Columns("HRProScreenID").CellText(varBookmark)
              sPageTitle = .Columns("PageTitle").CellText(varBookmark)
              sURL = .Columns("URL").CellText(varBookmark)
              sStartMode = .Columns("startMode").CellText(varBookmark)
              sUtilityType = .Columns("UtilityType").CellText(varBookmark)
              sUtilityID = .Columns("UtilityID").CellText(varBookmark)
              sTableID = DecodeTag(.Tag, False)
              sViewID = DecodeTag(.Tag, True)
              sNewWindow = .Columns("NewWindow").CellText(varBookmark)
              'NPG20080211 Fault 12873
              sEMailAddress = .Columns("EMailAddress").CellText(varBookmark)
              sEMailSubject = .Columns("EMailSubject").CellText(varBookmark)
              sAppFilePath = .Columns("AppFilePath").CellText(varBookmark)
              sAppParameters = .Columns("AppParameters").CellText(varBookmark)
              sDocumentFilePath = .Columns("DocumentFilePath").CellText(varBookmark)
              fDisplayDocumentHyperlink = .Columns("DisplayDocumentHyperlink").CellValue(varBookmark)
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
          
          'NPG20080211 Fault 12873
           sSQL = "INSERT INTO tmpSSIntranetLinks" & _
            " ([linkType], [linkOrder], [prompt], [text], [screenID], [pageTitle], [url], [startMode], " & _
            "[utilityType], [utilityID], [viewID], [newWindow], [tableID], [EMailAddress], [EMailSubject], " & _
            "[AppFilePath], [AppParameters], [DocumentFilePath], [DisplayDocumentHyperlink], [Element_Type], " & _
            "[SeparatorOrientation], [PictureID], [Chart_ShowLegend], [Chart_Type], [Chart_ShowGrid], [Chart_StackSeries], " & _
            "[Chart_ViewID], [Chart_TableID], [Chart_ColumnID], [Chart_FilterID], [Chart_AggregateType], [Chart_ShowValues])" & _
            " SELECT " & _
            CStr(piLinkType) & "," & _
            CStr(iLoop) & "," & _
            "'" & Replace(sPrompt, "'", "''") & "'," & _
            "'" & Replace(sText, "'", "''") & "'," & _
            sScreenID & "," & _
            "'" & Replace(sPageTitle, "'", "''") & "'," & _
            "'" & Replace(sURL, "'", "''") & "'," & _
            sStartMode & "," & _
            sUtilityType & "," & _
            sUtilityID & "," & _
            sViewID & "," & _
            sNewWindow & "," & _
            sTableID & "," & _
            "'" & Replace(sEMailAddress, "'", "''") & "'," & _
            "'" & Replace(sEMailSubject, "'", "''") & "'," & _
            "'" & Replace(sAppFilePath, "'", "''") & "'," & _
            "'" & Replace(sAppParameters, "'", "''") & "'," & _
            "'" & Replace(sDocumentFilePath, "'", "''") & "',"

          sSQL = sSQL & _
            "" & IIf(fDisplayDocumentHyperlink, "1", "0") & "," & _
            "" & sElement_Type & "," & _
            sSeparatorOrientation & "," & _
            sPictureID & "," & _
            "" & IIf(fChartShowLegend, "1", "0") & "," & _
            sChartType & "," & _
            "" & IIf(fChartShowGrid, "1", "0") & "," & _
            "" & IIf(fChartStackSeries, "1", "0") & "," & _
            sChartViewID & "," & _
            sChartTableID & "," & _
            sChartColumnID & "," & _
            sChartFilterID & "," & _
            sChartAggregateType & "," & _
            "" & IIf(fChartShowValues, "1", "0")

          daoDb.Execute sSQL, dbFailOnError
        
          ' Get the ID of the link just saved.
          sSQL = "SELECT MAX(id) AS [result] FROM tmpSSIntranetLinks"
          Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          lngMaxID = rsTemp!result
          rsTemp.Close
          Set rsTemp = Nothing
  
          sGroupNames = Mid(.Columns("HiddenGroups").CellText(varBookmark), 2)
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
  Dim varBookmark As Variant
  Dim sSelectedTableIDs As String
  
  sSelectedTableIDs = "0"
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)

      sSelectedTableIDs = sSelectedTableIDs & "," & .Columns("TableID").CellText(varBookmark)
    Next iLoop
  End With

  SelectedTables = sSelectedTableIDs
  
End Function

Private Function SelectedViews() As String

  ' Return a string of the selected view IDs
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim sSelectedViewIDs As String
  
  sSelectedViewIDs = "0"
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)

      sSelectedViewIDs = sSelectedViewIDs & "," & .Columns("ViewID").CellText(varBookmark)
    Next iLoop
  End With

  SelectedViews = sSelectedViewIDs
  
End Function

Private Function SingleRecordViewID() As Long

  ' Return the ID of the defined single record view.
  Dim iLoop As Integer
  Dim lngSingleRecordViewID As Long
  Dim varBookmark As Variant
  Dim fSingleRecord As Boolean
  Dim sViewID As String
  
  lngSingleRecordViewID = 0
  
  With grdTableViews
    For iLoop = 0 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)

      fSingleRecord = .Columns("SingleRecord").CellValue(varBookmark)
      
      If fSingleRecord Then
        sViewID = .Columns("ViewID").CellText(varBookmark)
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
      mcolSSITableViews
      
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Prompt _
        & vbTab & .Text _
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
        & vbTab & .ElementType _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "") _
        & vbTab & IIf(.chkShowLegend.value = 0, "0", "1") _
        & vbTab & .cboChartType.ItemData(.cboChartType.ListIndex) _
        & vbTab & IIf(.chkDottedGridlines.value = 0, "0", "1") _
        & vbTab & IIf(.chkStackSeries.value = 0, "0", "1") _
        & vbTab & "0" _
        & vbTab & .ChartTableID & vbTab & .ChartColumnID _
        & vbTab & .ChartFilterID & vbTab & .ChartAggregateType _
        & vbTab & IIf(.chkShowValues.value = 0, "0", "1")

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
  Dim varBookmark As Variant

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
            varBookmark = .AddItemBookmark(iLoop)

            If .Columns("SingleRecord").CellValue(varBookmark) _
              And CLng(.Columns("ViewID").CellText(varBookmark)) <> frmTableView.ViewID Then
              ' Remove the other 'single record view' markers
              .Bookmark = varBookmark
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
  
    .Initialize SSINTLINK_BUTTON, _
      ctlSourceGrid.Columns("Prompt").Text, _
      ctlSourceGrid.Columns("ButtonText").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, _
      ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, True, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboButtonLinkView.List(cboButtonLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, ctlSourceGrid.Columns("AppParameters").Text, "", False, _
      ctlSourceGrid.Columns("Element_Type").value, val(ctlSourceGrid.Columns("SeparatorOrientation").Text), val(ctlSourceGrid.Columns("PictureID").Text), _
      ctlSourceGrid.Columns("ChartShowLegend").value, val(ctlSourceGrid.Columns("ChartType").Text), _
      ctlSourceGrid.Columns("ChartShowGrid").value, ctlSourceGrid.Columns("ChartStackSeries").value, _
      val(ctlSourceGrid.Columns("ChartViewID").Text), _
      val(ctlSourceGrid.Columns("ChartTableID").Text), _
      val(ctlSourceGrid.Columns("ChartColumnID").Text), _
      val(ctlSourceGrid.Columns("ChartFilterID").Text), _
      val(ctlSourceGrid.Columns("ChartAggregateType").Text), _
      ctlSourceGrid.Columns("ChartShowValues").value, mcolGroups, mcolSSITableViews
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Prompt _
        & vbTab & .Text _
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
        & vbTab & IIf(.optLink(SSINTLINKSEPARATOR).value, 1, IIf(.optLink(SSINTLINKCHART).value, 2, IIf(.optLink(SSINTLINKPWFSTEPS).value, 3, IIf(.optLink(SSINTLINKDB_VALUE).value, 4, 0)))) _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "") _
        & vbTab & IIf(.chkShowLegend.value = 0, "0", "1") _
        & vbTab & .cboChartType.ItemData(.cboChartType.ListIndex) _
        & vbTab & IIf(.chkDottedGridlines.value = 0, "0", "1") _
        & vbTab & IIf(.chkStackSeries.value = 0, "0", "1") _
        & vbTab & .ChartViewID _
        & vbTab & .ChartTableID _
        & vbTab & .ChartColumnID _
        & vbTab & .ChartFilterID _
        & vbTab & .ChartAggregateType & vbTab & .ChartShowValues

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
  If ctlSourceGrid.Columns("Element_Type").value = 3 Then
    BuildUserGroupCollection
    PopulateWFAccessGroup ctlSourceGrid, lngRow
  End If
  
  With frmLink
    .Initialize SSINTLINK_BUTTON, _
      ctlSourceGrid.Columns("Prompt").Text, _
      ctlSourceGrid.Columns("ButtonText").Text, _
      ctlSourceGrid.Columns("HRProScreenID").Text, _
      ctlSourceGrid.Columns("PageTitle").Text, _
      ctlSourceGrid.Columns("URL").Text, _
      DecodeTag(ctlSourceGrid.Tag, False), _
      ctlSourceGrid.Columns("startMode").Text, _
      DecodeTag(ctlSourceGrid.Tag, True), _
      ctlSourceGrid.Columns("UtilityType").Text, _
      ctlSourceGrid.Columns("UtilityID").Text, False, _
      ctlSourceGrid.Columns("HiddenGroups").Text, _
      cboButtonLinkView.List(cboButtonLinkView.ListIndex), _
      ctlSourceGrid.Columns("NewWindow").Text, _
      ctlSourceGrid.Columns("EMailAddress").Text, ctlSourceGrid.Columns("EMailSubject").Text, _
      ctlSourceGrid.Columns("AppFilePath").Text, ctlSourceGrid.Columns("AppParameters").Text, _
      "", False, _
      ctlSourceGrid.Columns("Element_Type").value, val(ctlSourceGrid.Columns("SeparatorOrientation").Text), val(ctlSourceGrid.Columns("PictureID").Text), _
      ctlSourceGrid.Columns("ChartShowLegend").Text, val(ctlSourceGrid.Columns("ChartType").Text), ctlSourceGrid.Columns("ChartShowGrid").Text, _
      ctlSourceGrid.Columns("ChartStackSeries").Text, val(ctlSourceGrid.Columns("ChartviewID").Text), val(ctlSourceGrid.Columns("ChartTableID").Text), _
      val(ctlSourceGrid.Columns("ChartColumnID").Text), val(ctlSourceGrid.Columns("ChartFilterID").Text), val(ctlSourceGrid.Columns("ChartAggregateType").Text), _
      ctlSourceGrid.Columns("ChartShowValues").Text, mcolGroups, mcolSSITableViews
    .Show vbModal

    If Not .Cancelled Then
      sRow = .Prompt _
        & vbTab & .Text _
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
        & vbTab & .ElementType _
        & vbTab & IIf(.chkNewColumn.value = 0, "0", "1") _
        & vbTab & IIf(Len(.txtIcon.Text) > 0, CStr(.PictureID), "") _
        & vbTab & IIf(.chkShowLegend = 0, "0", "1") _
        & vbTab & .cboChartType.ItemData(.cboChartType.ListIndex) _
        & vbTab & IIf(.chkDottedGridlines = 0, "0", "1") _
        & vbTab & IIf(.chkStackSeries = 0, "0", "1") _
        & vbTab & 0 _
        & vbTab & .ChartTableID _
        & vbTab & .ChartColumnID _
        & vbTab & .ChartFilterID _
        & vbTab & .ChartAggregateType & vbTab & IIf(.chkShowValues = 0, "0", "1")
                
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
  Dim varBookmark As Variant
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
            varBookmark = .AddItemBookmark(iLoop)

            If .Columns("SingleRecord").CellValue(varBookmark) _
              And CLng(.Columns("ViewID").CellText(varBookmark)) <> frmTableView.ViewID Then
              ' Remove the other 'single record view' markers
              .Bookmark = varBookmark
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

Private Sub cmdOK_Click()
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
  Dim varBookmark As Variant
  
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
        varBookmark = .AddItemBookmark(iLoop)
  
        If (Not .Columns("SingleRecord").CellValue(varBookmark)) _
          And (.Columns("HypertextLink").CellText(varBookmark) = "0") _
          And (.Columns("ButtonLink").CellText(varBookmark) = "0") _
          And (.Columns("DropdownListLink").CellText(varBookmark) = "0") Then
        
          sMsg = "No link type has been selected for the '" & .Columns("TableView").CellText(varBookmark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          Exit For
        End If
        
        If (Not .Columns("SingleRecord").CellValue(varBookmark)) _
          And (.Columns("HypertextLink").CellText(varBookmark) = "1") _
          And (Len(.Columns("HypertextLinkText").CellText(varBookmark)) = 0) Then
        
          sMsg = "No Hypertext Link text has been entered for the '" & .Columns("TableView").CellText(varBookmark) & "' view."

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
        If (Not .Columns("SingleRecord").CellValue(varBookmark)) _
          And (.Columns("ButtonLink").CellText(varBookmark) = "1") _
          And (Len(.Columns("ButtonLinkButtonText").CellText(varBookmark)) = 0) Then
        
          sMsg = "No Button Link button text has been entered for the '" & .Columns("TableView").CellText(varBookmark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
          Exit For
        End If
        
        If (Not .Columns("SingleRecord").CellValue(varBookmark)) _
          And (.Columns("DropdownListLink").CellText(varBookmark) = "1") _
          And (Len(.Columns("DropdownListLinkText").CellText(varBookmark)) = 0) Then
        
          sMsg = "No Dropdown List Link text has been entered for the '" & .Columns("TableView").CellText(varBookmark) & "' view."

          ssTabStrip.Tab = giPAGE_GENERAL
          .SetFocus
          .Bookmark = .AddItemBookmark(iLoop)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
          
          Exit For
        End If
        
        If (Len(.Columns("LinksLinkText").CellText(varBookmark)) = 0) Then
        
          sMsg = "No Links Link text has been entered for the '" & .Columns("TableView").CellText(varBookmark) & "' view."

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
  Dim varBookmark As Variant
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
      varBookmark = .AddItemBookmark(iLoop)

      sTableID = .Columns("TableID").CellText(varBookmark)
      sViewID = .Columns("ViewID").CellText(varBookmark)
      fSingleRecordView = .Columns("SingleRecord").CellValue(varBookmark)
      fWFOutOfOffice = .Columns("WFOutOfOffice").CellValue(varBookmark)
      
      If fSingleRecordView Then
        sButtonLink = "0"
        sHypertextLink = "0"
        sDropdownListLink = "0"
        sPageTitle = ""
      Else
        sButtonLinkPromptText = .Columns("ButtonLinkPromptText").CellText(varBookmark)
        sButtonLinkButtonText = .Columns("ButtonLinkButtonText").CellText(varBookmark)
        sHypertextLinkText = .Columns("HypertextLinkText").CellText(varBookmark)
        sDropdownListLinkText = .Columns("DropdownListLinkText").CellText(varBookmark)
        sButtonLink = .Columns("ButtonLink").CellText(varBookmark)
        sHypertextLink = .Columns("HypertextLink").CellText(varBookmark)
        sDropdownListLink = .Columns("DropdownListLink").CellText(varBookmark)
        sPageTitle = .Columns("PageTitle").CellText(varBookmark)
      End If
      
      sLinksLinkText = .Columns("LinksLinkText").CellText(varBookmark)
      
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
            vbTab & rsLinks!Text & _
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
            vbTab & CStr(IIf(IsNull(rsLinks!Element_Type), "0", CStr(rsLinks!Element_Type))) & _
            vbTab & IIf(IsNull(rsLinks!SeparatorOrientation), "0", CStr(rsLinks!SeparatorOrientation)) & _
            vbTab & IIf(IsNull(rsLinks!PictureID), "0", CStr(rsLinks!PictureID)) & _
            vbTab & IIf(IsNull(rsLinks!Chart_ShowLegend), "0", IIf(rsLinks!Chart_ShowLegend, "1", "0")) & _
            vbTab & CStr(rsLinks!Chart_Type) & _
            vbTab & IIf(IsNull(rsLinks!Chart_ShowGrid), "0", IIf(rsLinks!Chart_ShowGrid, "1", "0")) & _
            vbTab & IIf(IsNull(rsLinks!Chart_StackSeries), "0", IIf(rsLinks!Chart_StackSeries, "1", "0")) & _
            vbTab & CStr(rsLinks!Chart_ViewID) & _
            vbTab & CStr(rsLinks!Chart_TableID) & _
            vbTab & CStr(rsLinks!Chart_ColumnID) & _
            vbTab & CStr(rsLinks!Chart_FilterID) & _
            vbTab & CStr(rsLinks!Chart_AggregateType) & vbTab & IIf(IsNull(rsLinks!Chart_ShowValues), "0", IIf(rsLinks!Chart_ShowValues, "1", "0"))
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
  Dim varBookmark As Variant
 
  Set mcolSSITableViews = New clsSSITableViews

  With grdTableViews
  
    For iLoop = 0 To (.Rows - 1) Step 1
      
      varBookmark = .AddItemBookmark(iLoop)
      
      lngTableID = .Columns("TableID").CellValue(varBookmark)
      lngViewID = .Columns("ViewID").CellValue(varBookmark)
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
  Dim varBookmark As Variant
  Dim lngTableID As Long
  Dim lngViewID As Long
  Dim sTableViewName As String
  
  cboHypertextLinkView.Clear
  cboButtonLinkView.Clear
  cboDropdownListLinkView.Clear
  cboDocumentView.Clear
  
  With grdTableViews
  
    For iLoop = 0 To (.Rows - 1)
    
      varBookmark = .AddItemBookmark(iLoop)
      
      lngTableID = .Columns("TableID").CellValue(varBookmark)
      lngViewID = .Columns("ViewID").CellValue(varBookmark)
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
                  
      objGroup.GroupName = Trim(!Name)
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
  Dim varBookmark As Variant
  Dim sCombinedHiddenGroups As String
  Dim fNoGroupsFound As Boolean
  
  ' this function will find all the hidden access (user) groups for Workflow Pending Steps only
  ' and concatenate the list into a string. This will be used in link validation to ensure only one
  ' workflow pending steps is on screen per access (user) group.
  
  fNoGroupsFound = True
  
  sCombinedHiddenGroups = vbTab & ""
  For iLoop = 0 To ctlSourceGrid.Rows - 1
    varBookmark = ctlSourceGrid.AddItemBookmark(iLoop)
    sHiddenGroups = ctlSourceGrid.Columns("HiddenGroups").CellText(varBookmark)
    
    If ctlSourceGrid.Columns("Element_Type").CellText(varBookmark) = 3 And iLoop <> ExcludeRowNum Then
    
      fNoGroupsFound = False
      
      aHiddenGroups = Split(sHiddenGroups, vbTab)
    
      For jLoop = 1 To mcolGroups.Count
        ' if the hidden group list doesn't contain this security group it's visible (duh) so
        ' change the allow property to false
        If InStr(sHiddenGroups, vbTab & mcolGroups(jLoop).GroupName) & vbTab = 0 Then
          mcolGroups(jLoop).Allow = False
        End If
      Next
      
    End If
  Next
  
End Sub
