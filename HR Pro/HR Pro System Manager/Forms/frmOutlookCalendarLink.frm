VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmOutlookCalendarLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outlook Calendar Link"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1061
   Icon            =   "frmOutlookCalendarLink.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1240
      Picture         =   "frmOutlookCalendarLink.frx":000C
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   56
      Top             =   5100
      Visible         =   0   'False
      Width           =   525
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
      Left            =   675
      Picture         =   "frmOutlookCalendarLink.frx":08D6
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   55
      Top             =   5100
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8385
      TabIndex        =   53
      Top             =   5175
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7080
      TabIndex        =   52
      Top             =   5175
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   45
      TabIndex        =   54
      Top             =   45
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   8837
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
      TabCaption(0)   =   "De&finition"
      TabPicture(0)   =   "frmOutlookCalendarLink.frx":11A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDateRange"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTimeRange"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraReminder"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDestCalendars"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraLinkDetails"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Colu&mns"
      TabPicture(1)   =   "frmOutlookCalendarLink.frx":11BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraColumns(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraColumns(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraColumns(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Co&ntent"
      TabPicture(2)   =   "frmOutlookCalendarLink.frx":11D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraContent"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraLinkDetails 
         Caption         =   "Link Details :"
         Height          =   1520
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4500
         Begin VB.ComboBox cboBusyStatus 
            Height          =   315
            ItemData        =   "frmOutlookCalendarLink.frx":11F4
            Left            =   1170
            List            =   "frmOutlookCalendarLink.frx":1204
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1040
            Width           =   3150
         End
         Begin VB.TextBox txtFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   640
            Width           =   2850
         End
         Begin VB.CommandButton cmdFilter 
            Height          =   315
            Left            =   4020
            Picture         =   "frmOutlookCalendarLink.frx":122E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   640
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.TextBox txtTitle 
            Height          =   315
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   2
            Top             =   240
            Width           =   3150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Status :"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   1100
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Filter :"
            Height          =   195
            Left            =   225
            TabIndex        =   3
            Top             =   700
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Title :"
            Height          =   195
            Left            =   225
            TabIndex        =   1
            Top             =   300
            Width           =   405
         End
      End
      Begin VB.Frame fraDestCalendars 
         Caption         =   "Destination Calendar(s) :"
         Height          =   2920
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   4500
         Begin VB.ListBox lstDestinations 
            Height          =   1860
            ItemData        =   "frmOutlookCalendarLink.frx":137C
            Left            =   200
            List            =   "frmOutlookCalendarLink.frx":137E
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   300
            Width           =   4140
         End
         Begin VB.CommandButton cmdAddresses 
            Caption         =   "Fol&ders..."
            Height          =   400
            Left            =   3135
            TabIndex        =   10
            Top             =   2360
            Width           =   1200
         End
      End
      Begin VB.Frame fraReminder 
         Caption         =   "Reminder :"
         Height          =   975
         Left            =   4760
         TabIndex        =   28
         Top             =   3865
         Width           =   4620
         Begin VB.ComboBox cboOffsetPeriod 
            Height          =   315
            ItemData        =   "frmOutlookCalendarLink.frx":1380
            Left            =   3200
            List            =   "frmOutlookCalendarLink.frx":1390
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   240
            Width           =   1275
         End
         Begin VB.CheckBox chkReminder 
            Caption         =   "Enable &Reminder"
            Height          =   195
            Left            =   225
            TabIndex        =   29
            Top             =   300
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin COASpinner.COA_Spinner spnOffset 
            Height          =   315
            Left            =   2430
            TabIndex        =   31
            Top             =   240
            Width           =   675
            _ExtentX        =   1191
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
         Begin VB.Label lblReminder 
            AutoSize        =   -1  'True
            Caption         =   "(Personal Mailboxes Only)"
            Height          =   195
            Left            =   165
            TabIndex        =   30
            Top             =   570
            Width           =   2220
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns Available :"
         Height          =   4515
         Index           =   0
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   3570
         Begin ComctlLib.ListView ListView1 
            Height          =   4065
            Left            =   180
            TabIndex        =   34
            Top             =   240
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   7170
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
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns Selected :"
         Height          =   4515
         Index           =   1
         Left            =   -69160
         TabIndex        =   42
         Top             =   360
         Width           =   3570
         Begin VB.TextBox txtHeading 
            Height          =   315
            Left            =   1185
            MaxLength       =   50
            TabIndex        =   45
            Top             =   3960
            Width           =   2220
         End
         Begin ComctlLib.ListView ListView2 
            Height          =   3615
            Left            =   180
            TabIndex        =   43
            Top             =   240
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   6376
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
         Begin VB.Label lblHeading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Heading :"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   4020
            Width           =   690
         End
      End
      Begin VB.Frame fraContent 
         Height          =   4500
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   9280
         Begin VB.TextBox txtSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   3580
         End
         Begin VB.CommandButton cmdSubject 
            Height          =   315
            Left            =   4755
            Picture         =   "frmOutlookCalendarLink.frx":13B2
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.TextBox txtBody 
            Height          =   3720
            Left            =   1170
            MaxLength       =   7000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   640
            Width           =   7955
         End
         Begin VB.Label lblSubject 
            AutoSize        =   -1  'True
            Caption         =   "Subject :"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Text :"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   700
            Width           =   435
         End
      End
      Begin VB.Frame fraTimeRange 
         Caption         =   "Time Range :"
         Height          =   2355
         Left            =   4760
         TabIndex        =   16
         Top             =   1465
         Width           =   4620
         Begin VB.ComboBox cboColumnStartTime 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1500
            Width           =   1935
         End
         Begin VB.ComboBox cboColumnEndTime 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1900
            Width           =   1935
         End
         Begin VB.OptionButton optAllDayEvent 
            Caption         =   "All Day E&vent"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   300
            Width           =   1650
         End
         Begin VB.OptionButton optFixed 
            Caption         =   "Fi&xed"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   660
            Width           =   855
         End
         Begin VB.OptionButton optColumns 
            Caption         =   "Co&lumns"
            Height          =   195
            Left            =   225
            TabIndex        =   23
            Top             =   1560
            Width           =   1110
         End
         Begin TDBMask6Ctl.TDBMask TDBFixedStartTime 
            Height          =   315
            Left            =   2520
            TabIndex        =   20
            Top             =   600
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            Caption         =   "frmOutlookCalendarLink.frx":1500
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmOutlookCalendarLink.frx":1565
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   -2147483633
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99"
            HighlightText   =   2
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "09:00"
            Value           =   "0900"
         End
         Begin TDBMask6Ctl.TDBMask TDBFixedEndTime 
            Height          =   315
            Left            =   2520
            TabIndex        =   22
            Top             =   1005
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            Caption         =   "frmOutlookCalendarLink.frx":15A7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmOutlookCalendarLink.frx":160C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   -2147483633
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "99:99"
            HighlightText   =   2
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "17:50"
            Value           =   "1750"
         End
         Begin VB.Label lblEndTime 
            Caption         =   "End Time :"
            Height          =   195
            Left            =   1440
            TabIndex        =   26
            Top             =   1960
            Width           =   975
         End
         Begin VB.Label lblStartTime 
            Caption         =   "Start Time :"
            Height          =   195
            Left            =   1440
            TabIndex        =   24
            Top             =   1560
            Width           =   1065
         End
         Begin VB.Label lblEndTimeFixed 
            Caption         =   "End Time :"
            Height          =   195
            Left            =   1440
            TabIndex        =   21
            Top             =   1060
            Width           =   975
         End
         Begin VB.Label lblStartTimeFixed 
            Caption         =   "Start Time :"
            Height          =   195
            Left            =   1440
            TabIndex        =   19
            Top             =   660
            Width           =   1065
         End
      End
      Begin VB.Frame fraDateRange 
         Caption         =   "Date Range :"
         Height          =   1060
         Left            =   4760
         TabIndex        =   11
         Top             =   360
         Width           =   4620
         Begin VB.ComboBox cboEndDate 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   640
            Width           =   3015
         End
         Begin VB.ComboBox cboStartDate 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "End Date :"
            Height          =   195
            Left            =   225
            TabIndex        =   14
            Top             =   705
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Start Date :"
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame fraColumns 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   2
         Left            =   -71060
         TabIndex        =   35
         Top             =   880
         Width           =   1695
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remo&ve All"
            Height          =   405
            Left            =   30
            TabIndex        =   39
            Top             =   1620
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add A&ll"
            Height          =   405
            Left            =   30
            TabIndex        =   37
            Top             =   615
            Width           =   1575
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "R&emove"
            Height          =   405
            Left            =   30
            TabIndex        =   38
            Top             =   1125
            Width           =   1575
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   405
            Left            =   30
            TabIndex        =   36
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Do&wn"
            Height          =   405
            Left            =   30
            TabIndex        =   41
            Top             =   2985
            Width           =   1575
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "&Up"
            Height          =   405
            Left            =   30
            TabIndex        =   40
            Top             =   2475
            Width           =   1575
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOutlookCalendarLink.frx":164E
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOutlookCalendarLink.frx":1A1A
            Key             =   "IMG_CALC"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOutlookCalendarLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjOutlookLink As clsOutlookLink
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mfColumnDrag As Boolean
Private mblnLoading As Boolean

Public Property Get OutlookLink() As clsOutlookLink
  Set OutlookLink = mobjOutlookLink
End Property

Public Property Let OutlookLink(ByVal objNewValue As clsOutlookLink)
  Set mobjOutlookLink = objNewValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Private Sub cboBusyStatus_Click()
  Changed = True
End Sub

Private Sub cboColumnEndTime_Click()
  Changed = True
End Sub

Private Sub cboColumnStartTime_Click()
  Changed = True
End Sub

Private Sub cboEndDate_Click()
  Changed = True
End Sub

Private Sub cboOffsetPeriod_Click()
  Changed = True
End Sub

Private Sub cboStartDate_Click()
  Changed = True
End Sub

Private Sub chkReminder_Click()
  ReminderClick True
End Sub

Private Function ReminderClick(blnSetValues As Boolean) As Boolean

  Dim blnReminder As Boolean

  blnReminder = (chkReminder.value = vbChecked)
  With spnOffset
    .Enabled = blnReminder And Not mblnReadOnly
    .BackColor = IIf(blnReminder, vbWindowBackground, vbButtonFace)
    If blnSetValues Then
      .Text = IIf(blnReminder, "0", vbNullString)
    End If
  End With

  With cboOffsetPeriod
    .Enabled = blnReminder And Not mblnReadOnly
    .BackColor = IIf(blnReminder, vbWindowBackground, vbButtonFace)
    If blnSetValues Then
      .ListIndex = IIf(blnReminder, 0, -1)
    End If
  End With

  Changed = True

End Function


Private Sub cmdAddresses_Click()

  Dim objOutlookFolder As clsOutlookFolder
  Dim lngStatus() As Long
  Dim strKey As String
  Dim intRow As Integer
  Dim lngCount As Long

  Set objOutlookFolder = New clsOutlookFolder

  mobjOutlookLink.ClearDestinations
  For lngCount = 0 To lstDestinations.ListCount - 1
    If lstDestinations.Selected(lngCount) Then
      mobjOutlookLink.AddDestination lstDestinations.ItemData(lngCount)
    End If
  Next

  ' Initialize the OutlookFolder object.
  With objOutlookFolder
    .FolderID = 0  'ssGrdRecipients.Columns(4).Value
    .TableID = mobjOutlookLink.TableID

    ' Instruct the OutlookFolder object to handle the selection.
    If .SelectOutlookFolder Then
      mobjOutlookLink.AddDestination .FolderID
      PopulateFolders .FolderID
      Changed = True
    Else
      PopulateFolders 0
    End If
  
  End With
  
  ' Disassociate object variables.
  Set objOutlookFolder = Nothing

End Sub

Private Sub cmdFilter_Click()

  ' Display the 'Where Clause' expression selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise mobjOutlookLink.TableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      Changed = True
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
      End If
    End If
  
  End With


TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub cmdSubject_Click()
  
  ' Display the Record Description selection form.
  Dim fOK As Boolean
  Dim objExpr As CExpression
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    fOK = .Initialise(mobjOutlookLink.TableID, Val(txtSubject.Tag), giEXPR_OUTLOOKSUBJECT, giEXPRVALUE_CHARACTER)
  
    If fOK Then
      ' Instruct the expression object to display the
      ' expression selection form.
      If .SelectExpression Then
        txtSubject.Tag = .ExpressionID
        If txtSubject.Tag < 0 Then
          txtSubject.Tag = 0
        End If
        txtSubject.Text = GetExpressionName(.ExpressionID)
        Changed = True
      Else
        ' Check in case the original expression has been deleted.
        With recExprEdit
          .Index = "idxExprID"
          .Seek "=", Val(txtSubject.Tag), False
  
          If .NoMatch Then
            txtSubject.Tag = 0
            txtSubject.Text = vbNullString
          End If
        End With
      End If
    End If
  End With
  
  ' Disassociate object variables.
  Set objExpr = Nothing

End Sub

Private Sub Form_Load()

  mblnCancelled = True
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  If SSTab1.Tab = 0 Then
    SSTab1_Click 0
  Else
    SSTab1.Tab = 0
  End If

End Sub

Private Sub cmdOK_Click()

  If ValidDefinition = False Then
    Exit Sub
  End If

  SaveDefinition
  mblnCancelled = False
  Me.Hide

  Application.ChangedOutlookLink = True

End Sub


Private Function ValidDefinition() As Boolean

  Dim lngStartTime As Long
  Dim lngEndTime As Long
  Dim blnFound As Boolean
  Dim lngCount As Long

  ValidDefinition = False

  If Trim(txtTitle.Text) = vbNullString Then
    SSTab1.Tab = 0
    MsgBox "You must give this link a title.", vbExclamation, Me.Caption
    txtTitle.SetFocus
    Exit Function
  End If

  
  If lstDestinations.SelCount = 0 Then
    SSTab1.Tab = 0
    MsgBox "You must select at least one destination calendar.", vbExclamation, Me.Caption
    txtTitle.SetFocus
    Exit Function
  End If


  If optFixed.value = True Then
    With TDBFixedStartTime
      If InStr(.Text, "_") > 0 Then
        SSTab1.Tab = 0
        MsgBox "You must enter a start time.", vbExclamation, Me.Caption
        .SetFocus
        Exit Function
      Else
        If Val(Left(.Text, 2)) > 23 Or Val(Right(.Text, 2)) > 59 Then
          SSTab1.Tab = 0
          MsgBox "Invalid start time.", vbExclamation, Me.Caption
          .SetFocus
          Exit Function
        End If
      End If
    End With

    With TDBFixedEndTime
      If InStr(.Text, "_") > 0 Then
        SSTab1.Tab = 0
        MsgBox "You must enter a end time.", vbExclamation, Me.Caption
        .SetFocus
        Exit Function
      Else
        If Val(Left(.Text, 2)) > 23 Or Val(Right(.Text, 2)) > 59 Then
          SSTab1.Tab = 0
          MsgBox "Invalid end time.", vbExclamation, Me.Caption
          .SetFocus
          Exit Function
        End If
      End If

    End With
    
    
    'If no end date then check that the end time is after the start time
    If cboEndDate.ListIndex = 0 Then
      lngStartTime = (Val(Left(TDBFixedStartTime.Text, 2)) * 60) + Val(Right(TDBFixedStartTime.Text, 2))
      lngEndTime = (Val(Left(TDBFixedEndTime.Text, 2)) * 60) + Val(Right(TDBFixedEndTime.Text, 2))
  
      If lngStartTime > lngEndTime Then
        SSTab1.Tab = 0
        MsgBox "The end time cannot be before the start time.", vbExclamation, Me.Caption
        TDBFixedEndTime.SetFocus
        Exit Function
      End If
    End If

'  ElseIf optColumns.Value = True Then
'
'    If cboColumnStartTime.ListIndex = -1 Then
'      SSTab1.Tab = 0
'      MsgBox "You must select a start time column.", vbExclamation, Me.Caption
'      Exit Function
'    End If
'
'    If cboColumnEndTime.ListIndex = -1 Then
'      SSTab1.Tab = 0
'      MsgBox "You must select an end time column.", vbExclamation, Me.Caption
'      Exit Function
'    End If


  End If

  ValidDefinition = True

End Function


Private Function SaveDefinition() As Boolean

  Dim lngColumnID As Long
  Dim lngCount As Long

  With mobjOutlookLink

    .Title = txtTitle.Text
    .FilterID = Val(txtFilter.Tag)
    .BusyStatus = SelectedComboItem(cboBusyStatus)

    .StartDate = SelectedComboItem(cboStartDate)
    .EndDate = SelectedComboItem(cboEndDate)

    .TimeRange = IIf(optFixed.value, 1, IIf(optColumns.value, 2, 0))
    .FixedStartTime = TDBFixedStartTime.Text
    .FixedEndTime = TDBFixedEndTime.Text
    .ColumnStartTime = SelectedComboItem(cboColumnStartTime)
    .ColumnEndTime = SelectedComboItem(cboColumnEndTime)

    .Reminder = (chkReminder.value = vbChecked)
    .ReminderOffset = spnOffset.value
    .ReminderPeriod = SelectedComboItem(cboOffsetPeriod)

    .Subject = Val(txtSubject.Tag)
    .content = txtBody.Text


    mobjOutlookLink.ClearDestinations
    For lngCount = 0 To lstDestinations.ListCount - 1
      If lstDestinations.Selected(lngCount) Then
        mobjOutlookLink.AddDestination lstDestinations.ItemData(lngCount)
      End If
    Next


    mobjOutlookLink.ClearColumns
    For lngCount = 1 To ListView2.ListItems.Count
      With ListView2.ListItems(lngCount)
        mobjOutlookLink.AddColumn Val(Mid(.key, 2)), .SubItems(2), lngCount
      End With
    Next

  End With

  SaveDefinition = True

End Function

Private Sub cmdCancel_Click()

  If Me.Changed Then
    Select Case MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
    Case vbYes
      cmdOK_Click
      Exit Sub
    Case vbCancel
      Exit Sub
    End Select
  End If

  Me.Hide

End Sub


Private Sub PopulateAvailable(blnFirstLoad As Boolean)
  
  ListView1.ListItems.Clear
  If blnFirstLoad Then
    cboStartDate.Clear
    cboEndDate.Clear
    cboEndDate.AddItem "<None>"
    cboColumnStartTime.Clear
    cboColumnEndTime.Clear
  End If


  With recColEdit
    .Index = "idxName"
    .Seek ">=", mobjOutlookLink.TableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mobjOutlookLink.TableID Then
          Exit Do
        End If

        
        'If (Not !Deleted) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) Then

        If (Not !Deleted) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
          (!ControlType <> giCTRL_OLE) And _
          (!ControlType <> giCTRL_PHOTO) And _
          (!ControlType <> giCTRL_LINK) Then

          If Not AlreadyUsed(!ColumnID) Then
            ListView1.ListItems.Add , "C" & CStr(!ColumnID), !ColumnName, "IMG_TABLE", "IMG_TABLE"
          End If
          
          If blnFirstLoad Then
            Select Case !DataType
            Case sqlDate
              cboStartDate.AddItem !ColumnName
              cboStartDate.ItemData(cboStartDate.NewIndex) = !ColumnID
              cboEndDate.AddItem !ColumnName
              cboEndDate.ItemData(cboEndDate.NewIndex) = !ColumnID
            
            Case sqlVarChar
  
              If (!Size = 2) Or (!Size = 5 And !Mask = "99:99") Then
                cboColumnStartTime.AddItem !ColumnName
                cboColumnStartTime.ItemData(cboColumnStartTime.NewIndex) = !ColumnID
                cboColumnEndTime.AddItem !ColumnName
                cboColumnEndTime.ItemData(cboColumnEndTime.NewIndex) = !ColumnID
              End If
  
            End Select

          End If

        End If

        .MoveNext
      Loop
    End If
  End With

  If blnFirstLoad Then
    If cboStartDate.ListIndex = -1 Then
      If cboStartDate.ListCount > 0 Then
        cboStartDate.ListIndex = 0
      End If
    End If
    If cboEndDate.ListIndex = -1 Then
      cboEndDate.ListIndex = 0
    End If
  
    optColumns.Enabled = (cboColumnStartTime.ListCount > 0 And Not mblnReadOnly)
    lblStartTime.Enabled = (cboColumnStartTime.ListCount > 0 And Not mblnReadOnly)
    lblEndTime.Enabled = (cboColumnStartTime.ListCount > 0 And Not mblnReadOnly)
  End If

End Sub


Private Function AlreadyUsed(lngColumnID As Long) As Boolean
  
  Dim objItem As ListItem
  
  For Each objItem In ListView2.ListItems
    If objItem.key = "C" & CStr(lngColumnID) Then
      AlreadyUsed = True
      Set objItem = Nothing
      Exit Function
    End If
  Next objItem
  
  Set objItem = Nothing

End Function

Private Function SelectedComboItem(cboTemp As ComboBox) As Long
  With cboTemp
    If .ListIndex >= 0 Then
      SelectedComboItem = .ItemData(.ListIndex)
    Else
      SelectedComboItem = 0
    End If
  End With
End Function


Public Function PopulateControls() As Boolean

  Dim objLinkColumn As clsOutlookLinkColumn
  Dim objListItem As ListItem
  Dim lngCount As Long
  Dim fOK As Boolean

  With mobjOutlookLink

    For lngCount = 1 To .LinkColumns.Count
      Set objLinkColumn = .LinkColumns.Item(lngCount)
      Set objListItem = ListView2.ListItems.Add(, "C" & CStr(objLinkColumn.ColumnID), GetColumnName(objLinkColumn.ColumnID, True), "IMG_TABLE", "IMG_TABLE")
      objListItem.SubItems(2) = objLinkColumn.Heading
    Next

    PopulateFolders 0
    PopulateAvailable True
    UpdateButtonStatus


    txtTitle.Text = .Title
    txtFilter.Tag = .FilterID
    txtFilter.Text = GetExpressionName(txtFilter.Tag)

    SetComboItem cboBusyStatus, .BusyStatus
    
    SetComboItem cboStartDate, .StartDate
    SetComboItem cboEndDate, .EndDate

    Select Case .TimeRange
    Case 0
      optAllDayEvent.value = True
    Case 1
      optFixed.value = True
      TDBFixedStartTime.Text = .FixedStartTime
      TDBFixedEndTime.Text = .FixedEndTime
    Case 2
      optColumns.value = True
      SetComboItem cboColumnStartTime, .ColumnStartTime
      SetComboItem cboColumnEndTime, .ColumnEndTime
    End Select

    chkReminder.value = IIf(.Reminder, vbChecked, vbUnchecked)
    spnOffset.value = .ReminderOffset
    SetComboItem cboOffsetPeriod, .ReminderPeriod

    txtSubject.Tag = .Subject
    txtSubject.Text = GetExpressionName(txtSubject.Tag)
    txtBody.Text = .content

  End With

  Set objLinkColumn = Nothing


  fOK = (cboStartDate.ListCount > 0)
  If Not fOK Then
    MsgBox "Unable to set up any outlook links on this table as it does not contain any date columns.", vbCritical
  End If

  Changed = Not fOK
  PopulateControls = fOK

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    Cancel = True
  End If

End Sub

Private Sub lstDestinations_ItemCheck(Item As Integer)

  If mblnReadOnly And Not mblnLoading Then
    lstDestinations.Selected(Item) = Not lstDestinations.Selected(Item)
  End If
  Changed = True

End Sub

Private Sub optAllDayEvent_Click()
  TimeRangeClick 0, True
  Changed = True
End Sub

Private Sub optFixed_Click()
  TimeRangeClick 1, True
  Changed = True
End Sub

Private Sub optColumns_Click()
  TimeRangeClick 2, True
  Changed = True
End Sub


Private Sub TimeRangeClick(intType As Integer, blnSetValues As Boolean)

  Dim blnFixed As Boolean
  Dim blnColumn As Boolean
  
  blnFixed = (intType = 1)
  With TDBFixedStartTime
    lblStartTimeFixed.Enabled = (blnFixed And Not mblnReadOnly)
    .Enabled = blnFixed And Not mblnReadOnly
    .BackColor = IIf(blnFixed, vbWindowBackground, vbButtonFace)
    .Format = IIf(blnFixed, "99:99", vbNullString)
    If blnSetValues Then
      .Text = IIf(blnFixed, "09:00", vbNullString)
    End If
  End With
  With TDBFixedEndTime
    lblEndTimeFixed.Enabled = (blnFixed And Not mblnReadOnly)
    .Enabled = blnFixed And Not mblnReadOnly
    .BackColor = IIf(blnFixed, vbWindowBackground, vbButtonFace)
    .Format = IIf(blnFixed, "99:99", vbNullString)
    If blnSetValues Then
      .Text = IIf(blnFixed, "17:00", vbNullString)
    End If
  End With

  blnColumn = (intType = 2 And cboColumnStartTime.ListCount > 0)
  With cboColumnStartTime
    lblStartTime.Enabled = (blnColumn And Not mblnReadOnly)
    .Enabled = blnColumn And Not mblnReadOnly
    .BackColor = IIf(blnColumn And Not mblnReadOnly, vbWindowBackground, vbButtonFace)
    If blnSetValues Then
      .ListIndex = IIf(blnColumn, 0, -1)
    End If
  End With
  With cboColumnEndTime
    lblEndTime.Enabled = (blnColumn And Not mblnReadOnly)
    .Enabled = blnColumn And Not mblnReadOnly
    .BackColor = IIf(blnColumn And Not mblnReadOnly, vbWindowBackground, vbButtonFace)
    If blnSetValues Then
      .ListIndex = IIf(blnColumn, 0, -1)
    End If
  End With

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

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If mblnReadOnly Then
    Exit Sub
  End If

  If mfColumnDrag Then
    ListView1.Drag vbCancel
    mfColumnDrag = False
  End If
  
End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)

  UpdateButtonStatus

End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  If mblnReadOnly Then
    Exit Sub
  End If

  If mfColumnDrag Then
    ListView2.Drag vbCancel
    mfColumnDrag = False
  End If

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
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

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
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
    'If ListView2.HitTest(x, y) Is Nothing Then
      CopyToSelected False
    'Else
    '  CopyToSelected False, ListView2.HitTest(x, y).Index
    'End If
    ListView1.Drag vbCancel
'  Else
'    If ListView2.HitTest(x, y) Is Nothing Then
'      ChangeSelectedOrder
'    Else
'      ChangeSelectedOrder ListView2.HitTest(x, y).Index
'    End If
'    ListView2.Drag vbCancel
  End If

End Sub


Private Sub fraColumns_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)

  ' Change pointer to the nodrop icon
  Source.DragIcon = picNoDrop.Picture
  
End Sub


Private Sub ListView2_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument(0).Picture
  End If

  ' Set DropHighlight to the mouse's coordinates.
  Set ListView2.DropHighlight = ListView2.HitTest(x, y)

End Sub

Private Sub ListView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  ' Change pointer to drop icon
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument(0).Picture
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
  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  
  Dim strText As String
  Dim strType As String
  Dim itmX As ListItem
  
  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass

  For Each itmX In ListView2.ListItems
    itmX.Selected = False
  Next
  ListView2.SelectedItem = Nothing
  ListView2.Refresh

  
  For iLoop = 1 To ListView1.ListItems.Count

    If bAll Or ListView1.ListItems(iLoop).Selected Then

      strType = Left(ListView1.ListItems(iLoop).key, 1)
      strText = ListView1.ListItems(iLoop).Text
      Set itmX = ListView2.ListItems.Add(, ListView1.ListItems(iLoop).key, strText, ListView1.ListItems(iLoop).Icon, ListView1.ListItems(iLoop).SmallIcon)
      itmX.Tag = ListView1.ListItems(iLoop).Tag
      itmX.SubItems(1) = strType & strText
      itmX.SubItems(2) = strText
      itmX.Selected = Not bAll
      
      Changed = True
  
    End If
  
  Next iLoop
  
  
  For iLoop = ListView1.ListItems.Count To 1 Step -1
    If bAll Or ListView1.ListItems(iLoop).Selected Then
      
      iTempItemIndex = iLoop
      ListView1.ListItems.Remove ListView1.ListItems(iLoop).key
    
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
  
  UpdateButtonStatus

  Me.Changed = True
  Screen.MousePointer = vbNormal

Exit Function

LocalErr:
  MsgBox "Error selecting columns", vbCritical

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
        strType = Left(.Item(iLoop).key, 1)
        lngID = Val(Mid(.Item(iLoop).key, 2))
        iTempItemIndex = iLoop
        .Remove .Item(iLoop).key
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
  
  PopulateAvailable False
  UpdateButtonStatus

  Me.Changed = True
  Screen.MousePointer = vbNormal

Exit Function

LocalErr:
  MsgBox "Error deselecting columns", vbCritical

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
    End If
  Next

  If intSelCount = 1 Then
    txtHeading.Enabled = Not mblnReadOnly
    txtHeading.BackColor = IIf(Not mblnReadOnly, vbWindowBackground, vbButtonFace)
    txtHeading.Text = ListView2.SelectedItem.SubItems(2)
  Else
    txtHeading.Enabled = False
    txtHeading.BackColor = vbButtonFace
    txtHeading.Text = ""
  End If


  If mblnReadOnly Then
    Exit Function
  End If
  
  cmdAddAll.Enabled = (ListView1.ListItems.Count > 0 And SSTab1.Tab = 1)
  cmdAdd.Enabled = (ListView1.ListItems.Count > 0 And SSTab1.Tab = 1)
  
  cmdRemoveAll.Enabled = (ListView2.ListItems.Count > 0 And SSTab1.Tab = 1)
  cmdRemove.Enabled = (ListView2.ListItems.Count > 0 And SSTab1.Tab = 1)
    
  cmdMoveUp.Enabled = ((ListView2.ListItems.Count > 0) And (intSelCount = 1) And iSelectedRow > 1 And SSTab1.Tab = 1)
  cmdMoveDown.Enabled = ((ListView2.ListItems.Count > 0) And (intSelCount = 1) And iSelectedRow < ListView2.ListItems.Count And SSTab1.Tab = 1)

End Function


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
  If MsgBox("Are you sure you wish to remove all columns from this definition ?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    CopyToAvailable True
  End If
  
End Sub

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


Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(ByVal blnNewValue As Boolean)
  If Not mblnLoading Then
    cmdOK.Enabled = blnNewValue And Not mblnReadOnly
  End If
End Property

Private Sub cmdAdd_LostFocus()
  cmdAdd.Picture = cmdAdd.Picture
End Sub

Private Sub cmdAddAll_LostFocus()
  cmdAddAll.Picture = cmdAddAll.Picture
End Sub

Private Sub cmdMoveDown_LostFocus()
  cmdMoveDown.Picture = cmdMoveDown.Picture
End Sub

Private Sub cmdMoveUp_LostFocus()
  cmdMoveUp.Picture = cmdMoveUp.Picture
End Sub

Private Sub cmdRemove_LostFocus()
  cmdRemove.Picture = cmdRemove.Picture
End Sub

Private Sub cmdRemoveAll_LostFocus()
  cmdRemoveAll.Picture = cmdRemoveAll.Picture
End Sub

Private Sub cmdMoveDown_Click()
  ChangeSelectedOrder ListView2.SelectedItem.Index + 2, True
End Sub

Private Sub cmdMoveUp_Click()
  ChangeSelectedOrder ListView2.SelectedItem.Index - 1
End Sub

Private Function ChangeSelectedOrder(Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)

  ' SUB COMPLETED 28/01/00
  ' This function changes the order of listitems in the selected listview.
  ' At the moment, different arrays are used depending on what information you
  ' need to store...change the array to a type if it would suit the purpose
  ' better
  
  ' Dimension arrays
  Dim iLoop As Integer, key() As String, Text() As String, Icon() As Variant, SmallIcon() As Variant
  
  Dim SubItem1() As Variant
  Dim SubItem2() As Variant
  
  ReDim key(0), Text(0), Icon(0), SmallIcon(0)
  ReDim SubItem1(0), SubItem2(0)
  
  Dim itmX As ListItem
  
  ' Clear the highlight
  Set ListView2.DropHighlight = Nothing
  
  ' If drop point is below all other items, then fix the intbeforeindex var
  If intBeforeIndex = 0 Then intBeforeIndex = ListView2.ListItems.Count + 1
  
  ' First get all the items that are above the drop point that arent selected
  For iLoop = 1 To (intBeforeIndex - 1)
    If ListView2.ListItems(iLoop).Selected = False Then
      ReDim Preserve key(UBound(key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)

      key(UBound(key) - 1) = ListView2.ListItems(iLoop).key
      Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
    
      ReDim Preserve SubItem1(UBound(SubItem1) + 1)
      ReDim Preserve SubItem2(UBound(SubItem2) + 1)
            
      SubItem1(UBound(SubItem1) - 1) = ListView2.ListItems(iLoop).SubItems(1)
      SubItem2(UBound(SubItem2) - 1) = ListView2.ListItems(iLoop).SubItems(2)
        
    End If
  Next iLoop
  
  ' Now get all the items that are selected
  For iLoop = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(iLoop).Selected = True Then
      ReDim Preserve key(UBound(key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
      key(UBound(key) - 1) = ListView2.ListItems(iLoop).key
      Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
    
      ReDim Preserve SubItem1(UBound(SubItem1) + 1)
      ReDim Preserve SubItem2(UBound(SubItem2) + 1)
            
      SubItem1(UBound(SubItem1) - 1) = ListView2.ListItems(iLoop).SubItems(1)
      SubItem2(UBound(SubItem2) - 1) = ListView2.ListItems(iLoop).SubItems(2)
    
    End If
  Next iLoop
  
  ' Now get all the items below the drop point that arent selected
  If intBeforeIndex <> 0 Then
    For iLoop = (intBeforeIndex) To ListView2.ListItems.Count
      If ListView2.ListItems(iLoop).Selected = False Then
        ReDim Preserve key(UBound(key) + 1)
        ReDim Preserve Text(UBound(Text) + 1)
        ReDim Preserve Icon(UBound(Icon) + 1)
        ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
        key(UBound(key) - 1) = ListView2.ListItems(iLoop).key
        Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
        Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
        SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
      
        ReDim Preserve SubItem1(UBound(SubItem1) + 1)
        ReDim Preserve SubItem2(UBound(SubItem2) + 1)

        SubItem1(UBound(SubItem1) - 1) = ListView2.ListItems(iLoop).SubItems(1)
        SubItem2(UBound(SubItem2) - 1) = ListView2.ListItems(iLoop).SubItems(2)
      
      End If
    Next iLoop
  End If
  
  ' Clear all items from the listview
  ListView2.ListItems.Clear
  
  ' Add items in the right order from the array
  For iLoop = LBound(key) To (UBound(key) - 1)
    
    Set itmX = ListView2.ListItems.Add(, key(iLoop), Text(iLoop), Icon(iLoop), SmallIcon(iLoop))
  
    itmX.SubItems(1) = SubItem1(iLoop)
    itmX.SubItems(2) = SubItem2(iLoop)
  
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

Private Sub spnOffset_Change()
  Changed = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  Dim blnEnabled As Boolean
  
  mblnLoading = True
  
  blnEnabled = (SSTab1.Tab = 0 And Not mblnReadOnly)
  
  ControlsDisableAll fraLinkDetails, blnEnabled
  'lstDestinations.Enabled = blnEnabled
  fraDestCalendars.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
  lstDestinations.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  lstDestinations.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
  cmdAddresses.Enabled = (SSTab1.Tab = 0)
  ControlsDisableAll fraDateRange, blnEnabled
  
  fraTimeRange.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
  optAllDayEvent.Enabled = blnEnabled
  optFixed.Enabled = blnEnabled
  optColumns.Enabled = (blnEnabled And cboColumnStartTime.ListCount > 0)

  fraReminder.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
  chkReminder.Enabled = blnEnabled
  lblReminder.Enabled = blnEnabled

  blnEnabled = (SSTab1.Tab = 1 And Not mblnReadOnly)

  ControlsDisableAll fraColumns(2), blnEnabled
  ListView1.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  ListView2.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  
  ControlsDisableAll fraContent, (SSTab1.Tab = 2 And Not mblnReadOnly)

  Select Case SSTab1.Tab
  Case 0
    If optAllDayEvent.value = True Then
      TimeRangeClick 0, False
    ElseIf optFixed.value = True Then
      TimeRangeClick 1, False
    Else
      TimeRangeClick 2, False
    End If
    ReminderClick False
    txtFilter.BackColor = vbButtonFace
    txtFilter.Enabled = False
    cmdFilter.Enabled = True

  Case 1
    UpdateButtonStatus

  Case 2
    txtSubject.BackColor = vbButtonFace
    txtSubject.Enabled = False
    cmdSubject.Enabled = True

  End Select

  mblnLoading = False

End Sub

Private Sub TDBFixedEndTime_Change()
  Changed = True
End Sub

Private Sub TDBFixedStartTime_Change()
  Changed = True
End Sub

Private Sub txtBody_Change()
  Changed = True
End Sub

Private Sub txtFilter_Change()
  Changed = True
End Sub

Private Sub txtHeading_Change()

  Dim lst As ListItem

  If txtHeading.Enabled Then
    For Each lst In ListView2.ListItems
      If lst.Selected Then
        If Not (lst.SubItems(2) = txtHeading.Text) Then
            lst.SubItems(2) = txtHeading.Text
            Me.Changed = True
        End If
      End If
    Next
  End If

End Sub


Private Sub PopulateFolders(lngSelected As Long)
  
  Dim rsRecipients As dao.Recordset
  Dim rsTemp As dao.Recordset
  Dim strSQL As String
  Dim lngCount As Long

  Dim lngMax As Long
  Dim lngLen As Long

  strSQL = "SELECT Name, FolderID " & _
           " FROM tmpOutlookFolders " & _
           " WHERE (tableID = 0 OR tableID = " & CStr(mobjOutlookLink.TableID) & ") " & _
           " AND deleted = false" & _
           " ORDER BY Name"
  Set rsRecipients = daoDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)

  With rsRecipients

    mblnLoading = True
    
    lstDestinations.Clear
    Do While Not .EOF
      lstDestinations.AddItem !Name
      lstDestinations.ItemData(lstDestinations.NewIndex) = !FolderID

      If lngSelected = !FolderID Then
        lstDestinations.ListIndex = lstDestinations.NewIndex
      End If

      For lngCount = 0 To UBound(mobjOutlookLink.Destinations)
        If mobjOutlookLink.Destinations(lngCount) = !FolderID Then
          lstDestinations.Selected(lstDestinations.NewIndex) = True
        End If
      Next

      .MoveNext
    Loop

    mblnLoading = False

  End With

  rsRecipients.Close
  Set rsRecipients = Nothing

  If lngSelected = 0 Then
    lstDestinations.ListIndex = -1
  End If

End Sub


Private Sub txtSubject_Change()
  Changed = True
End Sub

Private Sub txtTitle_Change()
  Changed = True
End Sub

Private Sub txtTitle_GotFocus()
  With txtTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub


Private Sub txtBody_GotFocus()
  With txtBody
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  cmdOK.Default = False
End Sub

Private Sub txtBody_LostFocus()
  cmdOK.Default = True
End Sub

