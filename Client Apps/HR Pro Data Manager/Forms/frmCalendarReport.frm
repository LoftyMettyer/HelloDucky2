VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmCalendarReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar Report Definition"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1066
   Icon            =   "frmCalendarReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleMode       =   0  'User
   ScaleWidth      =   9781.917
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   90
      TabIndex        =   79
      Top             =   90
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
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
      TabPicture(0)   =   "frmCalendarReport.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBase"
      Tab(0).Control(1)=   "fraInformation"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Eve&nt Details"
      TabPicture(1)   =   "frmCalendarReport.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraEvents"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Report Detai&ls"
      TabPicture(2)   =   "frmCalendarReport.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDisplayOptions"
      Tab(2).Control(1)=   "fraReportEnd"
      Tab(2).Control(2)=   "fraReportStart"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&Sort Order"
      TabPicture(3)   =   "frmCalendarReport.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraSort"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "O&utput"
      TabPicture(4)   =   "frmCalendarReport.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraOutputFormat"
      Tab(4).Control(1)=   "fraOutputDestination"
      Tab(4).ControlCount=   2
      Begin VB.Frame fraOutputDestination 
         Caption         =   "Output Destination(s) :"
         Height          =   3975
         Left            =   -72240
         TabIndex        =   57
         Top             =   400
         Width           =   6555
         Begin VB.CheckBox chkPreview 
            Caption         =   "P&review on screen"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   58
            Top             =   400
            Width           =   3495
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   69
            Top             =   2720
            Width           =   1515
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   63
            Top             =   1820
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   60
            Top             =   1300
            Width           =   1650
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   59
            Top             =   850
            Value           =   1  'Checked
            Width           =   3105
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6105
            Picture         =   "frmCalendarReport.frx":0098
            TabIndex        =   72
            Top             =   2660
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6105
            Picture         =   "frmCalendarReport.frx":0110
            TabIndex        =   66
            Top             =   1760
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   2660
            Width           =   2835
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3270
            TabIndex        =   74
            Top             =   3060
            Width           =   3135
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3270
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   2160
            Width           =   3135
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3270
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1240
            Width           =   3135
         End
         Begin VB.TextBox txtFilename 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1760
            Width           =   2835
         End
         Begin VB.TextBox txtEmailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3270
            TabIndex        =   76
            Tag             =   "0"
            Top             =   3460
            Width           =   3135
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   61
            Top             =   1305
            Width           =   1410
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   67
            Top             =   2220
            Width           =   1395
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   1800
            TabIndex        =   70
            Top             =   2715
            Width           =   1245
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   73
            Top             =   3120
            Width           =   1305
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   64
            Top             =   1815
            Width           =   1095
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   1800
            TabIndex        =   75
            Top             =   3525
            Width           =   1245
         End
      End
      Begin VB.Frame fraSort 
         Caption         =   "Sort Order :"
         Height          =   5730
         Left            =   -74880
         TabIndex        =   96
         Top             =   400
         Width           =   9180
         Begin VB.CommandButton cmdClearOrder 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   46
            Top             =   1860
            Width           =   1200
         End
         Begin VB.CommandButton cmdSortMoveDown 
            Caption         =   "Move Do&wn"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   48
            Top             =   5025
            Width           =   1200
         End
         Begin VB.CommandButton cmdSortMoveUp 
            Caption         =   "Move U&p"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   47
            Top             =   4575
            Width           =   1200
         End
         Begin VB.CommandButton cmdNewOrder 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7800
            TabIndex        =   43
            Top             =   315
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditOrder 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   44
            Top             =   765
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteOrder 
            Caption         =   "Re&move"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   45
            Top             =   1399
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdOrder 
            Height          =   5205
            Left            =   195
            TabIndex        =   49
            Top             =   315
            Width           =   7425
            ScrollBars      =   0
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   1
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
            Columns(0).Style=   3
            Columns(1).Width=   10848
            Columns(1).Caption=   "Column"
            Columns(1).Name =   "Column"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   2196
            Columns(2).Caption=   "Sort Order"
            Columns(2).Name =   "Order"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13097
            _ExtentY        =   9181
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
      Begin VB.Frame fraEvents 
         Caption         =   "Events :"
         Height          =   5760
         Left            =   120
         TabIndex        =   91
         Top             =   400
         Width           =   9180
         Begin VB.CommandButton cmdRemoveEvent 
            Caption         =   "Re&move"
            Height          =   400
            Left            =   7800
            TabIndex        =   18
            Top             =   1305
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddEvent 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   7800
            TabIndex        =   16
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveAllEvents 
            Caption         =   "Remo&ve All "
            Height          =   400
            Left            =   7800
            TabIndex        =   19
            Top             =   1755
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditEvent 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   7800
            TabIndex        =   17
            Top             =   750
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdEvents 
            Height          =   5235
            Left            =   195
            TabIndex        =   20
            Top             =   300
            Width           =   7440
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   28
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
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            RowNavigation   =   1
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   28
            Columns(0).Width=   1217
            Columns(0).Caption=   "Name"
            Columns(0).Name =   "Name"
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
            Columns(2).Width=   1138
            Columns(2).Caption=   "Table"
            Columns(2).Name =   "Table"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   3200
            Columns(3).Visible=   0   'False
            Columns(3).Caption=   "FilterID"
            Columns(3).Name =   "FilterID"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   1455
            Columns(4).Caption=   "Filter"
            Columns(4).Name =   "Filter"
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   8
            Columns(4).FieldLen=   256
            Columns(5).Width=   3200
            Columns(5).Visible=   0   'False
            Columns(5).Caption=   "StartDateID"
            Columns(5).Name =   "StartDateID"
            Columns(5).DataField=   "Column 5"
            Columns(5).DataType=   8
            Columns(5).FieldLen=   256
            Columns(6).Width=   1905
            Columns(6).Caption=   "Start Date"
            Columns(6).Name =   "Start Date"
            Columns(6).DataField=   "Column 6"
            Columns(6).DataType=   8
            Columns(6).FieldLen=   256
            Columns(7).Width=   3200
            Columns(7).Visible=   0   'False
            Columns(7).Caption=   "StartSessionID"
            Columns(7).Name =   "StartSessionID"
            Columns(7).DataField=   "Column 7"
            Columns(7).DataType=   8
            Columns(7).FieldLen=   256
            Columns(8).Width=   2302
            Columns(8).Caption=   "Start Session"
            Columns(8).Name =   "Start Session"
            Columns(8).DataField=   "Column 8"
            Columns(8).DataType=   8
            Columns(8).FieldLen=   256
            Columns(9).Width=   3200
            Columns(9).Visible=   0   'False
            Columns(9).Caption=   "EndDateID"
            Columns(9).Name =   "EndDateID"
            Columns(9).DataField=   "Column 9"
            Columns(9).DataType=   8
            Columns(9).FieldLen=   256
            Columns(10).Width=   1773
            Columns(10).Caption=   "End Date"
            Columns(10).Name=   "End Date"
            Columns(10).DataField=   "Column 10"
            Columns(10).DataType=   8
            Columns(10).FieldLen=   256
            Columns(11).Width=   3200
            Columns(11).Visible=   0   'False
            Columns(11).Caption=   "EndSessionID"
            Columns(11).Name=   "EndSessionID"
            Columns(11).DataField=   "Column 11"
            Columns(11).DataType=   8
            Columns(11).FieldLen=   256
            Columns(12).Width=   2196
            Columns(12).Caption=   "End Session"
            Columns(12).Name=   "End Session"
            Columns(12).DataField=   "Column 12"
            Columns(12).DataType=   8
            Columns(12).FieldLen=   256
            Columns(13).Width=   3200
            Columns(13).Visible=   0   'False
            Columns(13).Caption=   "DurationID"
            Columns(13).Name=   "DurationID"
            Columns(13).DataField=   "Column 13"
            Columns(13).DataType=   8
            Columns(13).FieldLen=   256
            Columns(14).Width=   1588
            Columns(14).Caption=   "Duration"
            Columns(14).Name=   "Duration"
            Columns(14).DataField=   "Column 14"
            Columns(14).DataType=   8
            Columns(14).FieldLen=   256
            Columns(15).Width=   3200
            Columns(15).Visible=   0   'False
            Columns(15).Caption=   "LegendType"
            Columns(15).Name=   "LegendType"
            Columns(15).DataField=   "Column 15"
            Columns(15).DataType=   8
            Columns(15).FieldLen=   256
            Columns(16).Width=   1402
            Columns(16).Caption=   "Key"
            Columns(16).Name=   "Legend"
            Columns(16).DataField=   "Column 16"
            Columns(16).DataType=   8
            Columns(16).FieldLen=   256
            Columns(17).Width=   3200
            Columns(17).Visible=   0   'False
            Columns(17).Caption=   "LegendTableID"
            Columns(17).Name=   "LegendTableID"
            Columns(17).DataField=   "Column 17"
            Columns(17).DataType=   8
            Columns(17).FieldLen=   256
            Columns(18).Width=   3200
            Columns(18).Visible=   0   'False
            Columns(18).Caption=   "LegendColumnID"
            Columns(18).Name=   "LegendColumnID"
            Columns(18).DataField=   "Column 18"
            Columns(18).DataType=   8
            Columns(18).FieldLen=   256
            Columns(19).Width=   3200
            Columns(19).Visible=   0   'False
            Columns(19).Caption=   "LegendCodeID"
            Columns(19).Name=   "LegendCodeID"
            Columns(19).DataField=   "Column 19"
            Columns(19).DataType=   8
            Columns(19).FieldLen=   256
            Columns(20).Width=   3200
            Columns(20).Visible=   0   'False
            Columns(20).Caption=   "LegendEventTypeID"
            Columns(20).Name=   "LegendEventTypeID"
            Columns(20).DataField=   "Column 20"
            Columns(20).DataType=   8
            Columns(20).FieldLen=   256
            Columns(21).Width=   3200
            Columns(21).Visible=   0   'False
            Columns(21).Caption=   "Desc1ID"
            Columns(21).Name=   "Desc1ID"
            Columns(21).DataField=   "Column 21"
            Columns(21).DataType=   8
            Columns(21).FieldLen=   256
            Columns(22).Width=   2249
            Columns(22).Caption=   "Description 1"
            Columns(22).Name=   "Description 1"
            Columns(22).DataField=   "Column 22"
            Columns(22).DataType=   8
            Columns(22).FieldLen=   256
            Columns(23).Width=   3200
            Columns(23).Visible=   0   'False
            Columns(23).Caption=   "Desc2ID"
            Columns(23).Name=   "Desc2ID"
            Columns(23).DataField=   "Column 23"
            Columns(23).DataType=   8
            Columns(23).FieldLen=   256
            Columns(24).Width=   2249
            Columns(24).Caption=   "Description 2"
            Columns(24).Name=   "Description 2"
            Columns(24).DataField=   "Column 24"
            Columns(24).DataType=   8
            Columns(24).FieldLen=   256
            Columns(25).Width=   3200
            Columns(25).Visible=   0   'False
            Columns(25).Caption=   "EventKey"
            Columns(25).Name=   "EventKey"
            Columns(25).DataField=   "Column 25"
            Columns(25).DataType=   8
            Columns(25).FieldLen=   256
            Columns(26).Width=   3200
            Columns(26).Caption=   "ColourName"
            Columns(26).Name=   "ColourName"
            Columns(26).DataField=   "Column 26"
            Columns(26).DataType=   8
            Columns(26).FieldLen=   256
            Columns(27).Width=   3200
            Columns(27).Visible=   0   'False
            Columns(27).Caption=   "ColourValue"
            Columns(27).Name=   "ColourValue"
            Columns(27).DataField=   "Column 27"
            Columns(27).DataType=   8
            Columns(27).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13123
            _ExtentY        =   9234
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
         Left            =   -74880
         TabIndex        =   100
         Top             =   400
         Width           =   2505
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Pivot Table"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   200
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   2800
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "Excel Char&t"
            Enabled         =   0   'False
            Height          =   195
            Index           =   5
            Left            =   200
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   54
            Top             =   2000
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   53
            Top             =   1600
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   52
            Top             =   1200
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "CS&V File"
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   51
            Top             =   800
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "D&ata Only"
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   50
            Top             =   400
            Value           =   -1  'True
            Width           =   1900
         End
      End
      Begin VB.Frame fraDisplayOptions 
         Caption         =   "Default Display Options :"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   99
         Top             =   2880
         Width           =   9140
         Begin VB.CheckBox chkStartOnCurrentMonth 
            Caption         =   "Start on Cu&rrent Month"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1650
            Value           =   1  'Checked
            Width           =   2820
         End
         Begin VB.CheckBox chkIncludeBHols 
            Caption         =   "Include &Bank Holidays"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   300
            Width           =   2655
         End
         Begin VB.CheckBox chkIncludeWorkingDaysOnly 
            Caption         =   "&Working Days Only"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   570
            Width           =   3000
         End
         Begin VB.CheckBox chkCaptions 
            Caption         =   "Show Calendar Ca&ptions"
            Height          =   240
            Left            =   240
            TabIndex        =   40
            Top             =   1110
            Value           =   1  'Checked
            Width           =   2835
         End
         Begin VB.CheckBox chkShadeWeekends 
            Caption         =   "Show Wee&kends"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1380
            Value           =   1  'Checked
            Width           =   2460
         End
         Begin VB.CheckBox chkShadeBHols 
            Caption         =   "Show Bank &Holidays"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   2595
         End
      End
      Begin VB.Frame fraReportEnd 
         Caption         =   "End Date : "
         Height          =   2310
         Left            =   -70200
         TabIndex        =   94
         Top             =   400
         Width           =   4450
         Begin VB.OptionButton optCustomEnd 
            Caption         =   "Cus&tom"
            Height          =   255
            Left            =   765
            TabIndex        =   35
            Top             =   1800
            Width           =   990
         End
         Begin VB.TextBox txtCustomEnd 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1800
            Width           =   1960
         End
         Begin VB.CommandButton cmdCustomEnd 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4005
            Picture         =   "frmCalendarReport.frx":0188
            TabIndex        =   36
            Top             =   1800
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.OptionButton optCurrentEnd 
            Caption         =   "Current Date"
            Height          =   255
            Left            =   765
            TabIndex        =   31
            Top             =   840
            Width           =   1470
         End
         Begin VB.ComboBox cboPeriodEnd 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":0200
            Left            =   2760
            List            =   "frmCalendarReport.frx":0210
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1260
            Width           =   1550
         End
         Begin VB.OptionButton optFixedEnd 
            Caption         =   "Fi&xed"
            Height          =   195
            Left            =   765
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOffsetEnd 
            Caption         =   "O&ffset"
            Height          =   255
            Left            =   765
            TabIndex        =   32
            Top             =   1320
            Width           =   855
         End
         Begin COASpinner.COA_Spinner spnFreqEnd 
            Height          =   315
            Left            =   2040
            TabIndex        =   33
            Top             =   1260
            Width           =   600
            _ExtentX        =   1058
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
            Enabled         =   0   'False
            MaximumValue    =   99
            MinimumValue    =   -99
            Text            =   "0"
         End
         Begin GTMaskDate.GTMaskDate GTMaskFixedEnd 
            Height          =   315
            Left            =   2040
            TabIndex        =   30
            Top             =   300
            Width           =   1305
            _Version        =   65537
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            NullText        =   "__/__/____"
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskCentury     =   2
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTips        =   0   'False
            BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblEnd 
            Caption         =   "End : "
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   365
            Width           =   450
         End
      End
      Begin VB.Frame fraReportStart 
         Caption         =   "Start Date :"
         Height          =   2310
         Left            =   -74880
         TabIndex        =   92
         Top             =   400
         Width           =   4450
         Begin VB.OptionButton optCustomStart 
            Caption         =   "Custo&m"
            Height          =   255
            Left            =   855
            TabIndex        =   27
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txtCustomStart 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1800
            Width           =   1960
         End
         Begin VB.CommandButton cmdCustomStart 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4005
            Picture         =   "frmCalendarReport.frx":0230
            TabIndex        =   28
            Top             =   1800
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.OptionButton optOffsetStart 
            Caption         =   "Offs&et"
            Height          =   255
            Left            =   855
            TabIndex        =   24
            Top             =   1320
            Width           =   855
         End
         Begin VB.ComboBox cboPeriodStart 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":02A8
            Left            =   2760
            List            =   "frmCalendarReport.frx":02B8
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1260
            Width           =   1550
         End
         Begin VB.OptionButton optFixedStart 
            Caption         =   "F&ixed"
            Height          =   255
            Left            =   855
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton optCurrentStart 
            Caption         =   "Current Date"
            Height          =   255
            Left            =   855
            TabIndex        =   23
            Top             =   840
            Width           =   1470
         End
         Begin GTMaskDate.GTMaskDate GTMaskFixedStart 
            Height          =   315
            Left            =   2040
            TabIndex        =   22
            Top             =   300
            Width           =   1305
            _Version        =   65537
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            NullText        =   "__/__/____"
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskCentury     =   2
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTips        =   0   'False
            BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin COASpinner.COA_Spinner spnFreqStart 
            Height          =   315
            Left            =   2040
            TabIndex        =   25
            Top             =   1260
            Width           =   600
            _ExtentX        =   1058
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
            Enabled         =   0   'False
            MaximumValue    =   99
            MinimumValue    =   -99
            Text            =   "0"
         End
         Begin VB.Label lblStart 
            Caption         =   "Start : "
            Height          =   255
            Left            =   200
            TabIndex        =   93
            Top             =   365
            Width           =   615
         End
      End
      Begin VB.Frame fraBase 
         Caption         =   "Data :"
         Height          =   3780
         Left            =   -74880
         TabIndex        =   85
         Top             =   2400
         Width           =   9180
         Begin VB.ComboBox cboDescriptionSeparator 
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":02D8
            Left            =   1575
            List            =   "frmCalendarReport.frx":0300
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3210
            Width           =   1425
         End
         Begin VB.CommandButton cmdDescExpr 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   4215
            Picture         =   "frmCalendarReport.frx":0346
            TabIndex        =   12
            Top             =   2385
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.TextBox txtDescExpr 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   104
            Tag             =   "0"
            Top             =   2385
            Width           =   2640
         End
         Begin VB.CheckBox chkPrintFilterHeader 
            Caption         =   "Display &title in the report header"
            Height          =   240
            Left            =   4815
            TabIndex        =   9
            Tag             =   "PrintFilterHeader"
            Top             =   1560
            Width           =   4110
         End
         Begin VB.ComboBox cboRegion 
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":03BE
            Left            =   5800
            List            =   "frmCalendarReport.frx":03C5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1980
            Width           =   3185
         End
         Begin VB.ComboBox cboDesc2 
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":03D1
            Left            =   1575
            List            =   "frmCalendarReport.frx":03D8
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1950
            Width           =   2955
         End
         Begin VB.ComboBox cboDesc1 
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":03E6
            Left            =   1575
            List            =   "frmCalendarReport.frx":03ED
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1530
            Width           =   2955
         End
         Begin VB.CheckBox chkGroupByDesc 
            Caption         =   "&Group by Description "
            Height          =   240
            Left            =   180
            TabIndex        =   13
            Top             =   2850
            Width           =   2145
         End
         Begin VB.TextBox txtBaseFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6705
            Locked          =   -1  'True
            TabIndex        =   87
            Tag             =   "0"
            Top             =   1100
            Width           =   1995
         End
         Begin VB.TextBox txtBasePicklist 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6705
            Locked          =   -1  'True
            TabIndex        =   86
            Tag             =   "0"
            Top             =   700
            Width           =   1995
         End
         Begin VB.OptionButton optBaseFilter 
            Caption         =   "&Filter"
            Height          =   195
            Left            =   5760
            TabIndex        =   7
            Top             =   1155
            Width           =   885
         End
         Begin VB.OptionButton optBasePicklist 
            Caption         =   "&Picklist"
            Height          =   195
            Left            =   5760
            TabIndex        =   5
            Top             =   765
            Width           =   975
         End
         Begin VB.OptionButton optBaseAllRecords 
            Caption         =   "&All"
            Height          =   195
            Left            =   5760
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.ComboBox cboBaseTable 
            Height          =   315
            ItemData        =   "frmCalendarReport.frx":03FA
            Left            =   1575
            List            =   "frmCalendarReport.frx":0401
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   300
            Width           =   2955
         End
         Begin VB.CommandButton cmdBasePicklist 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            Picture         =   "frmCalendarReport.frx":0410
            TabIndex        =   6
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.CommandButton cmdBaseFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   8700
            Picture         =   "frmCalendarReport.frx":0488
            TabIndex        =   8
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.Label lblDescSeparator 
            Caption         =   "Separator :"
            Height          =   285
            Left            =   195
            TabIndex        =   105
            Top             =   3270
            Width           =   1035
         End
         Begin VB.Label lblDescExpr 
            AutoSize        =   -1  'True
            Caption         =   "Description 3 : "
            Height          =   195
            Left            =   195
            TabIndex        =   103
            Top             =   2415
            Width           =   1260
         End
         Begin VB.Label lblRegion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region :"
            Height          =   195
            Left            =   4815
            TabIndex        =   98
            Top             =   2025
            Width           =   600
         End
         Begin VB.Label lblDesc2 
            AutoSize        =   -1  'True
            Caption         =   "Description 2 : "
            Height          =   195
            Left            =   195
            TabIndex        =   97
            Top             =   1995
            Width           =   1260
         End
         Begin VB.Label lblBaseTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Table :"
            Height          =   195
            Left            =   200
            TabIndex        =   90
            Top             =   360
            Width           =   885
         End
         Begin VB.Label lblBaseRecords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Records :"
            Height          =   195
            Left            =   4815
            TabIndex        =   89
            Top             =   360
            Width           =   825
         End
         Begin VB.Label lblDesc1 
            AutoSize        =   -1  'True
            Caption         =   "Description 1 : "
            Height          =   195
            Left            =   195
            TabIndex        =   88
            Top             =   1575
            Width           =   1260
         End
      End
      Begin VB.Frame fraInformation 
         Height          =   1950
         Left            =   -74880
         TabIndex        =   80
         Top             =   400
         Width           =   9180
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5670
            MaxLength       =   30
            TabIndex        =   2
            Top             =   300
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   0
            Top             =   300
            Width           =   2955
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1575
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   700
            Width           =   2955
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1080
            Left            =   5670
            TabIndex        =   106
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
            stylesets(0).Picture=   "frmCalendarReport.frx":0500
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
            stylesets(1).Picture=   "frmCalendarReport.frx":051C
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
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Left            =   4815
            TabIndex        =   84
            Top             =   315
            Width           =   765
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   200
            TabIndex        =   83
            Top             =   360
            Width           =   510
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   200
            TabIndex        =   82
            Top             =   750
            Width           =   900
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   4815
            TabIndex        =   81
            Top             =   750
            Width           =   780
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8357
      TabIndex        =   78
      Top             =   6495
      Width           =   1183
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7095
      TabIndex        =   77
      Top             =   6495
      Width           =   1183
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCalendarReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DataAccess Class
Private datData As HRProDataMgr.clsDataAccess
Private objOutputDef As clsOutputDef

'' Collection Class (Holds column details such as heading, size etc)
Private mcolEvents As clsCalendarEvents

' Long to hold current Calendar Report ID
Private mlngCalendarReportID As Long

Private mstrName As String
Private mstrDescription As String
Private mlngBaseTableID As Long
Private mblnAllRecords As Long
Private mlngPicklistID As Long
Private mlngFilterID As Long
Private mstrAccess As String
Private mstrUserName As String
Private mlngDesc1ID As Long
Private mlngDesc2ID As Long
Private mlngRegionID As Long
Private mblnGroupByDesc As Boolean
Private mintStartType As Integer
Private mstrFixedStart As String
Private mintStartFreq As Integer
Private mintStartPeriod As Integer
Private mlngCustomStart As Long
Private mintEndType As Integer
Private mstrFixedEnd As String
Private mintEndFreq As Integer
Private mintEndPeriod As Integer
Private mlngCustomEnd As Long
Private mblnShowBankHols As Boolean
Private mblnShowCaptions As Boolean
Private mblnShowWeekends As Boolean

' Variables to hold current (or previously) selected table details
Private mstrBaseTable As String
Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mblnFromPrint As Boolean
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mlngTimeStamp As Long
Private mblnForceHidden As Boolean
Private mblnRecordSelectionInvalid As Boolean
Private mblnDefinitionCreator As Boolean
Private mblnDeleted As Boolean
Private mbNeedsSave As Boolean
Private mblnNew As Boolean

Private Const lng_GRIDROWHEIGHT = 239

Private mcolColumnPivilages As CColumnPrivileges

Private Const mstrDateFormat = "mm/dd/yyyy"


Private Function EnableDisableTabControls()
  
  Dim i As Integer
 
  'enable/disable definition tab controls
  fraInformation.Enabled = ((SSTab1.Tab = 0))
  txtName.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  txtDesc.Enabled = (SSTab1.Tab = 0) 'And (Not mblnReadOnly))
  fraBase.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cboBaseTable.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  optBaseAllRecords.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  optBasePicklist.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cmdBasePicklist.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  optBaseFilter.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cmdBaseFilter.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  chkPrintFilterHeader.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cboDesc1.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cboDesc2.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cmdDescExpr.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  cboRegion.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
  chkGroupByDesc.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
    
  
  'enable/disable event details tab controls
  fraEvents.Enabled = ((SSTab1.Tab = 1))
'  grdEvents.Enabled = ((SSTab1.Tab = 1) And (Not mblnReadOnly))
  cmdAddEvent.Enabled = ((SSTab1.Tab = 1) And (Not mblnReadOnly))
  cmdEditEvent.Enabled = ((SSTab1.Tab = 1) And (Not mblnReadOnly))
  cmdRemoveEvent.Enabled = ((SSTab1.Tab = 1) And (Not mblnReadOnly))
  cmdRemoveAllEvents.Enabled = ((SSTab1.Tab = 1) And (Not mblnReadOnly))
  
  'enable/disable report details tab controls
  fraReportStart.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  optFixedStart.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  optCurrentStart.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  optOffsetStart.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  spnFreqStart.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  cboPeriodStart.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  fraReportEnd.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  optFixedEnd.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  optCurrentEnd.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  optOffsetEnd.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  spnFreqEnd.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  cboPeriodEnd.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  fraDisplayOptions.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  chkIncludeBHols.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  chkIncludeWorkingDaysOnly.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  chkShadeBHols.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  chkCaptions.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  chkShadeWeekends.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  chkStartOnCurrentMonth.Enabled = ((SSTab1.Tab = 2) And (Not mblnReadOnly))
  
  'enable/disable sort order tab controls
  fraSort.Enabled = ((SSTab1.Tab = 3))
'  grdOrder.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  cmdNewOrder.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  cmdEditOrder.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  cmdDeleteOrder.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  cmdClearOrder.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  cmdSortMoveUp.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  cmdSortMoveDown.Enabled = ((SSTab1.Tab = 3) And (Not mblnReadOnly))
  
  'enable/disable output options tab controls
  fraOutputFormat.Enabled = ((SSTab1.Tab = 4) And (Not mblnReadOnly))
  For i = 0 To optOutputFormat.UBound - 1 Step 1
    optOutputFormat(i).Enabled = ((SSTab1.Tab = 4) And (Not mblnReadOnly))
  Next i
  fraOutputDestination.Enabled = (SSTab1.Tab = 4) And (Not mblnReadOnly)
  'fraOutputFilename.Enabled = (SSTab1.Tab = 4) And (Not mblnReadOnly)
  
  
  Select Case SSTab1.Tab
    Case 0:
      If optBaseFilter.Value Then
        cmdBaseFilter.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
        cmdBasePicklist.Enabled = False
      ElseIf optBasePicklist.Value Then
        cmdBasePicklist.Enabled = ((SSTab1.Tab = 0) And (Not mblnReadOnly))
        cmdBaseFilter.Enabled = False
      Else
        cmdBasePicklist.Enabled = False
        cmdBaseFilter.Enabled = False
      End If
    
      If optBaseFilter.Value Or optBasePicklist.Value Then
        chkPrintFilterHeader.Enabled = (Not mblnReadOnly)
      Else
        chkPrintFilterHeader.Value = False
        chkPrintFilterHeader.Enabled = False
      End If
      
      Dim intDescCount As Integer
      
      If cboDesc1.ItemData(cboDesc1.ListIndex) > 0 Then intDescCount = intDescCount + 1
      If cboDesc2.ItemData(cboDesc2.ListIndex) > 0 Then intDescCount = intDescCount + 1
      If txtDescExpr.Tag > 0 Then intDescCount = intDescCount + 1
      If intDescCount < 2 Then cboDescriptionSeparator.ListIndex = 0
    
      cboDescriptionSeparator.Enabled = ((Not mblnReadOnly) And (intDescCount > 1))
      cboDescriptionSeparator.BackColor = IIf(cboDescriptionSeparator.Enabled, vbWindowBackground, vbButtonFace)
      lblDescSeparator.Enabled = cboDescriptionSeparator.Enabled
      
      UpdateReportDetailsTab
      
    Case 1:
      With grdEvents
        If (.Rows > 0) And (.SelBookmarks.Count <> 1) And (Not mblnReadOnly) Then
          .SelBookmarks.RemoveAll
          .MoveFirst
          .SelBookmarks.Add .Bookmark
        End If
        If mblnReadOnly Then .SelBookmarks.RemoveAll
      End With
      
      If cmdAddEvent.Enabled Then cmdAddEvent.SetFocus
      
      RefreshEventButtons
      
    Case 2:
      UpdateReportDetailsTab
      
    Case 3
      With grdOrder
      
        If (.Rows > 0) And (.SelBookmarks.Count <> 1) And (Not mblnReadOnly) Then
          .SelBookmarks.RemoveAll
          .MoveFirst
          .SelBookmarks.Add .Bookmark
        End If
        If mblnReadOnly Then .SelBookmarks.RemoveAll
      End With
      
      If cmdNewOrder.Enabled Then cmdNewOrder.SetFocus
      
      UpdateOrderButtonStatus
    
    Case 4:
       
  End Select

End Function

Private Function FormatGridColumnWidths() As Boolean

  Dim iColumn As Integer
  Dim iRow As Integer
  Dim lngTextWidth As Long
  Dim varBookmark As Variant
  Dim varOriginalPos As Variant

  lngTextWidth = 0
  With grdEvents
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

Private Sub GetCustomDate(ctlTarget As Control)
  
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  fOK = True
  
  Set objExpression = New clsExprExpression
  With objExpression
    
    fOK = .Initialise(0, Val(ctlTarget.Tag), giEXPR_RECORDINDEPENDANTCALC, giEXPRVALUE_DATE)
    
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

Private Function GetEventKey() As String

  'checks that the event name does not already exist in the current definition.
  Dim intCount As Integer
  Dim intCheck As Integer
  
  intCount = 1
  Do Until CheckUniqueEventName(CStr("EV_" & intCount))
    intCount = intCount + 1
  Loop
  
  GetEventKey = CStr("EV_" & intCount)

End Function
Private Function CheckUniqueEventName(pstrNewEventKey As String) As Boolean

  'checks that the event name does not already exist in the current definition.
  Dim i As Integer
  Dim objEvent As clsCalendarEvent
  
  CheckUniqueEventName = True
  
  For Each objEvent In mcolEvents.Collection
    If pstrNewEventKey = objEvent.Key Then
      CheckUniqueEventName = False
      Exit Function
    End If
  Next objEvent
  
  CheckUniqueEventName = True
  
End Function

Private Sub GetExpression(ctlSource As Control, ctlTarget As Control)
  
  ' Allow the user to select/create/modify an expression for the Calendar Report.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    If TypeOf ctlSource Is TextBox Then
      fOK = .Initialise(ctlSource.Tag, Val(ctlTarget.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
    ElseIf TypeOf ctlSource Is ComboBox Then
      fOK = .Initialise(ctlSource.ItemData(ctlSource.ListIndex), Val(ctlTarget.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
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

  EnableDisableTabControls
  
  ForceDefinitionToBeHiddenIfNeeded

End Sub

Public Sub PrintDef(plngCalendarReportID As Long)

  Dim objPrintDef As clsPrintDef
  Dim objEvent As clsCalendarEvent
  
  Dim sTemp As String
  Dim sTableName As String
  Dim sBaseTable As String
  Dim sPeriod As String
  Dim iLoop As Integer
  Dim fFirstLoop As Boolean
  Dim varBookmark As Variant
  
  Dim i As Integer
  
  mlngCalendarReportID = plngCalendarReportID

  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        ' First section --------------------------------------------------------
        .PrintHeader "Calendar Report : " & txtName.Text
    
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
        
        sBaseTable = cboBaseTable.List(cboBaseTable.ListIndex)
        .PrintNormal "Base Table : " & sBaseTable
        
        If optBaseAllRecords.Value Then
          .PrintNormal "Records : All Records"
        ElseIf optBasePicklist.Value Then
          .PrintNormal "Records : '" & datGeneral.GetPicklistName(txtBasePicklist.Tag) & "' picklist"
        ElseIf optBaseFilter.Value Then
          .PrintNormal "Records : '" & datGeneral.GetFilterName(txtBaseFilter.Tag) & "' filter"
        End If
        .PrintNormal
        .PrintNormal "Display filter or picklist title in the report header : " & IIf(chkPrintFilterHeader.Value = vbChecked, "Yes", "No")
        .PrintNormal
        
        .PrintNormal "Description 1 : " & cboDesc1.List(cboDesc1.ListIndex)
        .PrintNormal "Description 2 : " & cboDesc2.List(cboDesc2.ListIndex)
        .PrintNormal "Description 3 : " & txtDescExpr.Text
        .PrintNormal "Group By Description : " & IIf(chkGroupByDesc.Value = vbChecked, "Yes", "No")
        .PrintNormal "Separator : " & IIf(cboDescriptionSeparator.List(cboDescriptionSeparator.ListIndex) = vbNullString, "<None>", cboDescriptionSeparator.List(cboDescriptionSeparator.ListIndex))
        
        .PrintNormal "Region : " & cboRegion.List(cboRegion.ListIndex)
        
        .PrintNormal
        
        ' print the event details section
        .PrintTitle "Events Details"
        
        For Each objEvent In mcolEvents.Collection
          sTableName = objEvent.TableName
          .PrintNormal "Event Name : " & objEvent.Name
          .PrintNormal "Event Table : " & sTableName
          If objEvent.FilterID > 0 Then
            .PrintNormal "Event Filter : " & datGeneral.GetFilterName(objEvent.FilterID)
          Else
            .PrintNormal "Event Filter : <None>"
          End If
          
          .PrintNormal "Event Start Date : " & sTableName & "." & objEvent.StartDateName
          
          If objEvent.StartSessionID > 0 Then
            .PrintNormal "Event Start Session : " & sTableName & "." & objEvent.StartSessionName
          Else
            .PrintNormal "Event Start Session : <None>"
          End If
          
          If objEvent.EndDateID > 0 Then
            .PrintNormal "Event End Date : " & sTableName & "." & objEvent.EndDateName
            If objEvent.EndSessionID > 0 Then
              .PrintNormal "Event End Session : " & sTableName & "." & objEvent.EndSessionName
            Else
              .PrintNormal "Event End Session : <None>"
            End If
            
          ElseIf objEvent.DurationID > 0 Then
            .PrintNormal "Event Duration : " & sTableName & "." & objEvent.DurationName
          
          Else
            .PrintNormal "Event End : None"
          End If
          
          If objEvent.LegendType = 1 Then
            .PrintNormal "Key Event Type : " & objEvent.TableName & "." & objEvent.LegendEventTypeName
            .PrintNormal "Key Lookup Table : " & objEvent.LegendTableName
            .PrintNormal "Key Lookup Column : " & objEvent.LegendTableName & "." & objEvent.LegendColumnName
            .PrintNormal "Key Lookup Code : " & objEvent.LegendTableName & "." & objEvent.LegendCodeName
          Else
            .PrintNormal "Key Character : " & objEvent.LegendCharacter
          End If
          
          If objEvent.Description1ID > 0 Then
            .PrintNormal "Event Description 1 : " & datGeneral.GetColumnTableName(objEvent.Description1ID) & "." & objEvent.Description1Name
          Else
            .PrintNormal "Event Description 1 : <None>"
          End If
          
          If objEvent.Description2ID > 0 Then
            .PrintNormal "Event Description 2 : " & datGeneral.GetColumnTableName(objEvent.Description2ID) & "." & objEvent.Description2Name
          Else
            .PrintNormal "Event Description 2 : <None>"
          End If
          
          .PrintNormal ""
          
        Next objEvent

        .PrintTitle "Report Details"
        
        If optFixedStart.Value Then
          .PrintNormal "Report Fixed Start Date : " & IIf(IsNull(GTMaskFixedStart.DateValue), "<Null>", Format(GTMaskFixedStart.DateValue, DateFormat))
        ElseIf optCurrentStart.Value Then
          .PrintNormal "Report Start Date : " & "Current Date"
        ElseIf optOffsetStart.Value Then
          .PrintNormal "Report Start Date Offset: " & spnFreqStart.Value & " " & cboPeriodStart.List(cboPeriodStart.ListIndex)
        ElseIf optCustomEnd.Value Then
          .PrintNormal "Report Custom Start Date : " & txtCustomStart.Text
        End If
        
        If optFixedEnd.Value Then
          .PrintNormal "Report Fixed End Date : " & IIf(IsNull(GTMaskFixedEnd.DateValue), "<Null>", Format(GTMaskFixedEnd.DateValue, DateFormat))
        ElseIf optCurrentEnd.Value Then
          .PrintNormal "Report End Date : " & "Current Date"
        ElseIf optOffsetEnd.Value Then
          .PrintNormal "Report End Date Offset: " & spnFreqEnd.Value & " " & cboPeriodEnd.List(cboPeriodEnd.ListIndex)
        ElseIf optCustomEnd.Value Then
          .PrintNormal "Report Custom End Date : " & txtCustomEnd.Text
        End If
        
        .PrintNormal ""
        
        .PrintNormal "Include Bank Holidays: " & IIf(chkIncludeBHols.Value = vbChecked, "Yes", "No")
        .PrintNormal "Working Days Only : " & IIf(chkIncludeWorkingDaysOnly.Value = vbChecked, "Yes", "No")
        .PrintNormal "Show Bank Holidays : " & IIf(chkShadeBHols.Value = vbChecked, "Yes", "No")
        .PrintNormal "Show Captions : " & IIf(chkCaptions.Value = vbChecked, "Yes", "No")
        .PrintNormal "Show Weekends : " & IIf(chkShadeWeekends.Value = vbChecked, "Yes", "No")
        .PrintNormal "Start on Current Month : " & IIf(chkStartOnCurrentMonth.Value = vbChecked, "Yes", "No")
        
        .PrintNormal ""

        .PrintTitle "Sort Order"
          
        grdOrder.Redraw = False
        grdOrder.MoveFirst
        For i = 0 To grdOrder.Rows - 1 Step 1
          varBookmark = grdOrder.AddItemBookmark(i)
          .PrintNormal "Name : " & grdOrder.Columns("Column").CellText(varBookmark)
          .PrintNormal "Order : " & grdOrder.Columns("Order").CellText(varBookmark)
          .PrintNormal " "
        Next i
        grdOrder.Redraw = True
        
        .PrintNormal ""
        
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
          .PrintNormal "Output Destination : Display on screen after output"
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
          .PrintNormal "Email recipient : " & datGeneral.GetEmailGroupName(CLng(txtEmailGroup.Tag))
          .PrintNormal "Email Subject : " & txtEmailSubject.Text
          .PrintNormal "Email Attach As : " & txtEmailAttachAs.Text
        End If
        
        .PrintEnd
        .PrintConfirm "Calendar Report : " & txtName.Text, "Calendar Report Definition"
      End If
    
    End With
  
  End If

TidyUpAndExit:
  Set objPrintDef = Nothing
  Set objEvent = Nothing
  Exit Sub

LocalErr:
  MsgBox "Printing Calendar Report Definition Failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"
  GoTo TidyUpAndExit
  
End Sub
Private Sub ReportDateSetup()

  If Me.optFixedStart Then
    mintStartType = 0
    'JPD 20041117 Fault 8231
    mstrFixedStart = Replace(CStr(Format(GTMaskFixedStart.DateValue, mstrDateFormat)), UI.GetSystemDateSeparator, "/")
    mintStartFreq = 0
    mintStartPeriod = -1
    mlngCustomStart = 0
    
  ElseIf Me.optCurrentStart Then
    mintStartType = 1
    mstrFixedStart = ""
    mintStartFreq = 0
    mintStartPeriod = -1
    mlngCustomStart = 0
    
  ElseIf Me.optOffsetStart Then
    mintStartType = 2
    mstrFixedStart = ""
    mintStartFreq = Me.spnFreqStart.Value
    mintStartPeriod = Me.cboPeriodStart.ListIndex
    mlngCustomStart = 0
    
  ElseIf Me.optCustomStart Then
    mintStartType = 3
    mstrFixedStart = ""
    mintStartFreq = 0
    mintStartPeriod = -1
    mlngCustomStart = CLng(txtCustomStart.Tag)
  
  End If
  
  If Me.optFixedEnd Then
    mintEndType = 0
    'JPD 20041117 Fault 8231
    mstrFixedEnd = Replace(CStr(Format(GTMaskFixedEnd.DateValue, mstrDateFormat)), UI.GetSystemDateSeparator, "/")
    mintEndFreq = 0
    mintEndPeriod = -1
    mlngCustomEnd = 0
    
  ElseIf Me.optCurrentEnd Then
    mintEndType = 1
    mstrFixedEnd = ""
    mintEndFreq = 0
    mintEndPeriod = -1
    mlngCustomEnd = 0
    
  ElseIf Me.optOffsetEnd Then
    mintEndType = 2
    mstrFixedEnd = ""
    mintEndFreq = Me.spnFreqEnd.Value
    mintEndPeriod = Me.cboPeriodEnd.ListIndex
    mlngCustomEnd = 0
    
  ElseIf Me.optCustomEnd Then
    mintEndType = 3
    mstrFixedEnd = ""
    mintEndFreq = 0
    mintEndPeriod = -1
    mlngCustomEnd = CLng(txtCustomEnd.Tag)
      
  End If

End Sub
Private Sub RefreshEventButtons()

  cmdAddEvent.Enabled = Not mblnReadOnly

  cmdEditEvent.Enabled = (grdEvents.SelBookmarks.Count = 1) And (Not mblnReadOnly)
  
  cmdRemoveEvent.Enabled = (grdEvents.SelBookmarks.Count = 1) And (Not mblnReadOnly)
  
  cmdRemoveAllEvents.Enabled = (grdEvents.Rows > 0) And (Not mblnReadOnly)
  
End Sub
Private Function RetrieveCalendarReportDetails(plngCalendarReportID As Long) As Boolean

  Dim rsTemp As Recordset
  Dim iLoop As Integer
  Dim sText As String
  Dim fAlreadyNotified As Boolean
  Dim sMessage As String
  Dim rsChildren As ADODB.Recordset
  Dim sSQL As String
  Dim sAddLine As String
  
  Dim sTempTableName As String
  Dim sTempStartDateName As String
  Dim sTempStartSessionName As String
  Dim sTempEndDateName As String
  Dim sTempEndSessionName As String
  Dim sTempDurationName As String
  Dim sTempLegendTableName As String
  Dim sTempLegendColumnName As String
  Dim sTempLegendCodeName As String
  Dim sTempLegendEventTypeName As String
  Dim sTempDesc1Name As String
  Dim sTempDesc2Name As String
  
  On Error GoTo ErrorTrap
  
  'Load the basic guff first
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysCalendarReports.*, " & _
                                     "CONVERT(integer, ASRSysCalendarReports.TimeStamp) AS intTimeStamp " & _
                                     "FROM ASRSysCalendarReports WHERE ID = " & plngCalendarReportID)
  
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This Report definition has been deleted by another user.", vbExclamation + vbOKOnly, "Calendar Reports"
    Set rsTemp = Nothing
    RetrieveCalendarReportDetails = False
    mblnDeleted = True
    Exit Function
  End If
  
  ' Set Definition Name
  txtName.Text = rsTemp!Name
  
  ' Set Definition Description
  txtDesc.Text = IIf(IsNull(rsTemp!Description), "", rsTemp!Description)
    
  If FromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName.Text = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName.Text = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If
  
  ' Set Base Table
  
  LoadBaseCombo
  
  SetComboText cboBaseTable, datGeneral.GetTableName(rsTemp!BaseTable)
  mstrBaseTable = cboBaseTable.Text
  UpdateDependantFields
  
  
  ' Set Base Table Record Select Options
  If rsTemp!AllRecords Then optBaseAllRecords.Value = True
  If rsTemp!picklist > 0 Then
    optBasePicklist.Value = True
    txtBasePicklist.Tag = rsTemp!picklist
    txtBasePicklist.Text = datGeneral.GetPicklistName(rsTemp!picklist)

  End If
  
  If rsTemp!Filter > 0 Then
    optBaseFilter.Value = True
    txtBaseFilter.Tag = rsTemp!Filter
    txtBaseFilter.Text = datGeneral.GetFilterName(rsTemp!Filter)
  End If

  chkPrintFilterHeader.Value = IIf(rsTemp!PrintFilterHeader, vbChecked, vbUnchecked)

  mlngTimeStamp = rsTemp!intTimestamp
  
  ' =========================
  
  mblnReadOnly = Not datGeneral.SystemPermission("CALENDARREPORTS", "EDIT")
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(utlCalendarReport, mlngCalendarReportID) = ACCESS_READONLY)
  End If
  
  If rsTemp!Description1 > 0 Then
    SetComboText cboDesc1, datGeneral.GetColumnName(rsTemp!Description1)
  Else
    cboDesc1.ListIndex = 0
  End If
  If rsTemp!Description2 > 0 Then
    SetComboText cboDesc2, datGeneral.GetColumnName(rsTemp!Description2)
  Else
    cboDesc2.ListIndex = 0
  End If
  If rsTemp!DescriptionExpr > 0 Then
    txtDescExpr.Tag = rsTemp!DescriptionExpr
    txtDescExpr.Text = datGeneral.GetExpression(rsTemp!DescriptionExpr)
  Else
    txtDescExpr.Tag = 0
    txtDescExpr.Text = ""
  End If
  
  chkGroupByDesc.Value = IIf(rsTemp!GroupByDesc, vbChecked, vbUnchecked)
  
  If Not IsNull(rsTemp!DescriptionSeparator) Then
    If rsTemp!DescriptionSeparator = vbNullString Then
      SetComboText cboDescriptionSeparator, "<None>"
    ElseIf rsTemp!DescriptionSeparator = " " Then
      SetComboText cboDescriptionSeparator, "<Space>"
    Else
      SetComboText cboDescriptionSeparator, rsTemp!DescriptionSeparator
    End If
  Else
    SetComboText cboDescriptionSeparator, ", "
  End If
    
  If rsTemp!Region > 0 Then
    SetComboText cboRegion, datGeneral.GetColumnName(rsTemp!Region)
  Else
    cboRegion.ListIndex = 0
  End If
  
  'populate the report date options.
  Select Case rsTemp!StartType
  Case 0
    optFixedStart.Value = True
    GTMaskFixedStart.DateValue = CDate(rsTemp!FixedStart)
    
  Case 1
    optCurrentStart.Value = True
    
  Case 2
    optOffsetStart.Value = True
    spnFreqStart.Value = rsTemp!StartFrequency
    
    Select Case rsTemp!StartPeriod
    Case 0
      SetComboText cboPeriodStart, "Days"
    Case 1
      SetComboText cboPeriodStart, "Weeks"
    Case 2
      SetComboText cboPeriodStart, "Months"
    Case 3
      SetComboText cboPeriodStart, "Years"
    End Select
  
  Case 3
    optCustomStart.Value = True
    txtCustomStart.Text = datGeneral.GetExpression(rsTemp!StartDateExpr)
    txtCustomStart.Tag = rsTemp!StartDateExpr
    
  End Select
  
  Select Case rsTemp!EndType
  Case 0
    optFixedEnd.Value = True
    GTMaskFixedEnd.DateValue = CDate(rsTemp!FixedEnd)
    
  Case 1
    optCurrentEnd.Value = True
    
  Case 2
    optOffsetEnd.Value = True
    spnFreqEnd.Value = rsTemp!EndFrequency
    
    Select Case rsTemp!EndPeriod
    Case 0
      SetComboText cboPeriodEnd, "Days"
    Case 1
      SetComboText cboPeriodEnd, "Weeks"
    Case 2
      SetComboText cboPeriodEnd, "Months"
    Case 3
      SetComboText cboPeriodEnd, "Years"
    End Select
    
  Case 3
    optCustomEnd.Value = True
    txtCustomEnd.Text = datGeneral.GetExpression(rsTemp!EndDateExpr)
    txtCustomEnd.Tag = rsTemp!EndDateExpr
        
  End Select
 
  chkShadeBHols.Value = IIf(rsTemp!ShowBankHolidays, vbChecked, vbUnchecked)
  chkCaptions.Value = IIf(rsTemp!ShowCaptions, vbChecked, vbUnchecked)
  chkShadeWeekends.Value = IIf(rsTemp!ShowWeekends, vbChecked, vbUnchecked)
  chkIncludeWorkingDaysOnly.Value = IIf(rsTemp!IncludeWorkingDaysOnly, vbChecked, vbUnchecked)
  chkIncludeBHols.Value = IIf(rsTemp!IncludeBankHolidays, vbChecked, vbUnchecked)
  chkStartOnCurrentMonth.Value = IIf(rsTemp!StartOnCurrentMonth, vbChecked, vbUnchecked)
  
  'Output Options.
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
'    SetComboItem cboSaveExisting, rsTemp!OutputSaveExisting
'    txtFilename.Text = rsTemp!OutputFilename
'  End If
'
'  chkDestination(desEmail).Value = IIf(rsTemp!OutputEmail, vbChecked, vbUnchecked)
'
'  If rsTemp!OutputEmail Then
'    txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
'    txtEmailGroup.Tag = rsTemp!OutputEmailAddr
'    txtEmailSubject.Text = rsTemp!OutputEmailSubject
'    txtEmailAttachAs.Text = IIf(IsNull(rsTemp!OutputEmailAttachAs), vbNullString, rsTemp!OutputEmailAttachAs)
'  End If
  objOutputDef.PopulateOutputControls rsTemp

  If mblnReadOnly Then
    ControlsDisableAll Me
    grdEvents.Enabled = True
    grdOrder.Enabled = True
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  
  Set rsTemp = Nothing
  
  ' =========================
  
  sMessage = vbNullString
  
  ' Now load the events guff
  'Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysCalendarReportEvents WHERE CalendarReportID = " & plngCalendarReportID & " ORDER BY ID")


  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysCalendarReportEvents.*, ASRSysColours.ColDesc FROM ASRSysCalendarReportEvents " & _
                                     "JOIN ASRSysColours ON ASRSysColours.ColValue = ASRSysCalendarReportEvents.Colour " & _
                                     "WHERE CalendarReportID = " & plngCalendarReportID & " ORDER BY ID")




  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "Cannot load the event information for this Calendar Report", vbExclamation + vbOKOnly, "Calendar Reports"
    RetrieveCalendarReportDetails = False
    Set rsTemp = Nothing
    Exit Function
  End If

  Do Until rsTemp.EOF
  
    sTempTableName = datGeneral.GetTableName(rsTemp!TableID)

    If rsTemp!EventStartDateID > 0 Then
      sTempStartDateName = datGeneral.GetColumnName(rsTemp!EventStartDateID)
    Else
      MsgBox "Cannot load the event information for this Calendar Report", vbExclamation + vbOKOnly, "Calendar Reports"
      RetrieveCalendarReportDetails = False
      Set rsTemp = Nothing
      Exit Function
    End If
    
    If rsTemp!EventStartSessionID > 0 Then
      sTempStartSessionName = datGeneral.GetColumnName(rsTemp!EventStartSessionID)
    Else
      sTempStartSessionName = vbNullString
    End If
    
    If rsTemp!EventEndDateID > 0 Then
      sTempEndDateName = datGeneral.GetColumnName(rsTemp!EventEndDateID)
    Else
      sTempEndDateName = vbNullString
    End If
    
    If rsTemp!EventEndSessionID > 0 Then
      sTempEndSessionName = datGeneral.GetColumnName(rsTemp!EventEndSessionID)
    Else
      sTempEndSessionName = vbNullString
    End If
    
    If rsTemp!EventDurationID > 0 Then
      sTempDurationName = datGeneral.GetColumnName(rsTemp!EventDurationID)
    Else
      sTempDurationName = vbNullString
    End If
    
    If rsTemp!LegendLookupTableID > 0 Then
      sTempLegendTableName = datGeneral.GetTableName(rsTemp!LegendLookupTableID)
    Else
      sTempLegendTableName = vbNullString
    End If
    
    If rsTemp!LegendLookupColumnID > 0 Then
      sTempLegendColumnName = datGeneral.GetColumnName(rsTemp!LegendLookupColumnID)
    Else
      sTempLegendColumnName = vbNullString
    End If
    
    If rsTemp!LegendLookupCodeID > 0 Then
      sTempLegendCodeName = datGeneral.GetColumnName(rsTemp!LegendLookupCodeID)
    Else
      sTempLegendCodeName = vbNullString
    End If
    
    If rsTemp!LegendEventColumnID > 0 Then
      sTempLegendEventTypeName = datGeneral.GetColumnName(rsTemp!LegendEventColumnID)
    Else
      sTempLegendEventTypeName = vbNullString
    End If
    
    If rsTemp!EventDesc1ColumnID > 0 Then
      sTempDesc1Name = datGeneral.GetColumnName(rsTemp!EventDesc1ColumnID)
    Else
      sTempDesc1Name = vbNullString
    End If
    
    If rsTemp!EventDesc2ColumnID > 0 Then
      sTempDesc2Name = datGeneral.GetColumnName(rsTemp!EventDesc2ColumnID)
    Else
      sTempDesc2Name = vbNullString
    End If
    
    sAddLine = IIf(IsNull(rsTemp!Name), "", rsTemp!Name) & vbTab
    sAddLine = sAddLine & rsTemp!TableID & vbTab
    If rsTemp!TableID > 0 Then
      sAddLine = sAddLine & sTempTableName & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!FilterID & vbTab
    If rsTemp!FilterID > 0 Then
      sAddLine = sAddLine & datGeneral.GetFilterName(rsTemp!FilterID) & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!EventStartDateID & vbTab
    If rsTemp!EventStartDateID > 0 Then
      sAddLine = sAddLine & sTempStartDateName & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!EventStartSessionID & vbTab
    If rsTemp!EventStartSessionID > 0 Then
      sAddLine = sAddLine & sTempStartSessionName & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!EventEndDateID & vbTab
    If rsTemp!EventEndDateID > 0 Then
      sAddLine = sAddLine & sTempEndDateName & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
        
    sAddLine = sAddLine & rsTemp!EventEndSessionID & vbTab
    If rsTemp!EventEndSessionID > 0 Then
      sAddLine = sAddLine & sTempEndSessionName & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!EventDurationID & vbTab
    If rsTemp!EventDurationID > 0 Then
      sAddLine = sAddLine & sTempDurationName & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!LegendType & vbTab
    If rsTemp!LegendType = 1 Then
      sAddLine = sAddLine & sTempLegendTableName & "." & datGeneral.GetColumnName(rsTemp!LegendLookupColumnID) & vbTab
      sAddLine = sAddLine & rsTemp!LegendLookupTableID & vbTab
      sAddLine = sAddLine & rsTemp!LegendLookupColumnID & vbTab
      sAddLine = sAddLine & rsTemp!LegendLookupCodeID & vbTab
      sAddLine = sAddLine & rsTemp!LegendEventColumnID & vbTab
      
    Else
      sAddLine = sAddLine & rsTemp!LegendCharacter & vbTab
      sAddLine = sAddLine & 0 & vbTab
      sAddLine = sAddLine & 0 & vbTab
      sAddLine = sAddLine & 0 & vbTab
      sAddLine = sAddLine & 0 & vbTab
     
    End If
    
    sAddLine = sAddLine & rsTemp!EventDesc1ColumnID & vbTab
    If rsTemp!EventDesc1ColumnID > 0 Then
      sAddLine = sAddLine & sTempDesc1Name & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & rsTemp!EventDesc2ColumnID & vbTab
    If rsTemp!EventDesc2ColumnID > 0 Then
      sAddLine = sAddLine & sTempDesc2Name & vbTab
    Else
      sAddLine = sAddLine & vbNullString & vbTab
    End If
    
    sAddLine = sAddLine & Trim(rsTemp!EventKey) & vbTab
    sAddLine = sAddLine & Trim(rsTemp!ColDesc) & vbTab
    sAddLine = sAddLine & Trim(rsTemp!Colour)
    
    grdEvents.AddItem sAddLine
    
    mcolEvents.Add Trim(rsTemp!EventKey), rsTemp!Name, _
                rsTemp!TableID, sTempTableName, _
                rsTemp!FilterID, _
                rsTemp!EventStartDateID, sTempStartDateName, _
                rsTemp!EventStartSessionID, sTempStartSessionName, _
                rsTemp!EventEndDateID, sTempEndDateName, _
                rsTemp!EventEndSessionID, sTempEndSessionName, _
                rsTemp!EventDurationID, sTempDurationName, _
                rsTemp!LegendType, rsTemp!LegendCharacter, rsTemp!Colour, _
                rsTemp!LegendLookupTableID, sTempLegendTableName, _
                rsTemp!LegendLookupColumnID, sTempLegendColumnName, _
                rsTemp!LegendLookupCodeID, sTempLegendCodeName, _
                rsTemp!LegendEventColumnID, sTempLegendEventTypeName, _
                rsTemp!EventDesc1ColumnID, sTempDesc1Name, _
                rsTemp!EventDesc2ColumnID, sTempDesc2Name

    rsTemp.MoveNext

  Loop
  
  grdEvents.RowHeight = lng_GRIDROWHEIGHT
  
  FormatGridColumnWidths
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Cancelled = True
    RetrieveCalendarReportDetails = False
    Exit Function
  End If
 
 ' Now do the sort order guff
  Set rsTemp = datGeneral.GetRecords("SELECT * FROM ASRSysCalendarReportOrder WHERE CalendarReportID = " & plngCalendarReportID & " ORDER BY [OrderSequence]")
  
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "Cannot load the sort order for this Calendar Report", vbExclamation + vbOKOnly, "Calendar Reports"
    RetrieveCalendarReportDetails = False
    Set rsTemp = Nothing
    Exit Function
  End If
  
  ' Add to the sort order grid
  Do Until rsTemp.EOF
    Me.grdOrder.AddItem rsTemp!ColumnID & vbTab & _
                        datGeneral.GetTableName(rsTemp!TableID) & "." & datGeneral.GetColumnName(rsTemp!ColumnID) & vbTab & _
                        rsTemp!OrderType
    rsTemp.MoveNext
  Loop
  
  With Me.grdOrder
    .SelBookmarks.RemoveAll
    .MoveFirst
    .SelBookmarks.Add (.Bookmark)
    .RowHeight = lng_GRIDROWHEIGHT
  End With
 
  RetrieveCalendarReportDetails = True

TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function
  
ErrorTrap:
  MsgBox "Warning : Error whilst retrieving the Calendar Report definition." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Calendar Reports"
  RetrieveCalendarReportDetails = False
  GoTo TidyUpAndExit
  
End Function
Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOk.Enabled = pblnChanged
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngCalendarReportID
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
Public Property Get FromPrint() As Boolean
  FromPrint = mblnFromPrint
End Property

Public Property Let FromCopy(ByVal bCopy As Boolean)
  mblnFromCopy = bCopy
End Property
Public Property Let FromPrint(ByVal bPrint As Boolean)
  mblnFromPrint = bPrint
End Property
Public Property Get DefinitionOwner() As Boolean
  DefinitionOwner = mblnDefinitionCreator
End Property
Public Function Initialise(pbNew As Boolean, pbCopy As Boolean, Optional plngCalendarReportID As Long, Optional pbPrint) As Boolean

  ' This function is called from frmMain and prepares the form depending
  ' on whether the user is creating a new definition or editing an existing
  ' one.
  
  Screen.MousePointer = vbHourglass
  
  ' Set references to class modules
  Set datData = New HRProDataMgr.clsDataAccess
  
  mblnLoading = True
  mblnNew = pbNew
  
  grdEvents.RowHeight = lng_GRIDROWHEIGHT
  
  If pbNew Then
    mblnDefinitionCreator = True
    
    'Set ID to 0 to indicate new calendar report
    mlngCalendarReportID = 0
    
    chkGroupByDesc.Value = False
    SetComboText cboDescriptionSeparator, ", "
    
    'Report Details Tab
    optFixedStart.Value = True
    GTMaskFixedStart.DateValue = Null
    optFixedEnd.Value = True
    GTMaskFixedEnd.DateValue = Null
    
    chkCaptions.Value = vbChecked
    chkShadeBHols.Value = vbUnchecked
    chkShadeWeekends.Value = vbChecked
    chkIncludeWorkingDaysOnly.Value = vbUnchecked
    chkIncludeBHols.Value = vbUnchecked
    chkStartOnCurrentMonth.Value = vbChecked
    
    'Set controls to defaults
    ClearForNew
    
    'Load all possible Base Tables into combo
    LoadBaseCombo

    UpdateDependantFields
    
    optOutputFormat(0).Value = True
    objOutputDef.FormatClick 0, True
    
    PopulateAccessGrid
    
    Changed = False
    
  Else
    ' Make the CalendarReportID visible to the rest of the module
    mlngCalendarReportID = plngCalendarReportID
    
    ' Is is a copy of an existing one ?
    FromCopy = pbCopy
    
    ' We need to know if we are going to PRINT the definition.
    FromPrint = pbPrint
    
    PopulateAccessGrid
    
    If Not RetrieveCalendarReportDetails(plngCalendarReportID) Then
      If mblnDeleted Or Me.Cancelled Then
        Initialise = False
        Exit Function
      Else
        If MsgBox("HR Pro could not load all of the definition successfully. The recommendation is that" & vbCrLf & _
               "you delete the definition and create a new one, however, you may edit the existing" & vbCrLf & _
               "definition if you wish. Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, "Calendar Reports") = vbNo Then
          Me.Cancelled = True
          Initialise = False
          Exit Function
        End If
      End If
    End If
        
    If mblnReadOnly Then
      grdEvents.StyleSets.RemoveAll
      grdOrder.StyleSets.RemoveAll
    End If
    
    If pbCopy = True Then
      mlngCalendarReportID = 0
      Changed = True
    Else
      Changed = mblnRecordSelectionInvalid And (Not mblnReadOnly) ' False
    End If
    
  End If
  
  EnableDisableTabControls
  
  mblnLoading = False
  
  Cancelled = False
  Screen.MousePointer = vbNormal

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
  Set rsAccess = GetUtilityAccessRecords(utlCalendarReport, mlngCalendarReportID, mblnFromCopy)
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

Private Function SortOrderColumns() As String

  Dim i As Integer
  Dim s As String
  Dim varBookmark As Variant
  
  With Me.grdOrder
    If .Rows > 0 Then
      .Redraw = False
      '.MoveFirst
      For i = 0 To .Rows - 1 Step 1
        varBookmark = .AddItemBookmark(i)
        s = s & .Columns("ColumnID").CellValue(varBookmark) & ","
        '.MoveNext
      Next i
      s = Left(s, Len(s) - 1)
      .Redraw = True
    Else
      s = ""
    End If
  End With
  
  SortOrderColumns = s
  
End Function
Private Sub UpdateReportDetailsTab()

  mblnLoading = True
  
  ' update the start date frame
  If Me.optFixedStart Then
    Me.GTMaskFixedStart.Enabled = (Not mblnReadOnly)
    Me.spnFreqStart.Enabled = False
    Me.cboPeriodStart.Enabled = False
    Me.optFixedEnd.Enabled = (Not mblnReadOnly)
    Me.optOffsetEnd.Enabled = (Not mblnReadOnly)
    Me.optCurrentEnd.Enabled = (Not mblnReadOnly)
    Me.spnFreqStart.Enabled = False
    Me.spnFreqStart.Text = ""
    Me.cboPeriodStart.Enabled = False
    Me.cboPeriodStart.ListIndex = -1
    Me.cmdCustomStart.Enabled = False
    Me.txtCustomStart.Text = ""
    Me.txtCustomStart.Tag = 0
    
  ElseIf Me.optCurrentStart Then
    Me.GTMaskFixedStart.Enabled = False
    Me.GTMaskFixedStart.DateValue = Null
    Me.spnFreqStart.Enabled = False
    Me.spnFreqStart.Text = ""
    Me.cboPeriodStart.Enabled = False
    Me.cboPeriodStart.ListIndex = -1
    Me.optOffsetEnd.Enabled = (Not mblnReadOnly)
    Me.optCurrentEnd.Enabled = (Not mblnReadOnly)
    Me.optFixedEnd.Enabled = (Not mblnReadOnly)
    Me.cmdCustomStart.Enabled = False
    Me.txtCustomStart.Text = ""
    Me.txtCustomStart.Tag = 0
    
  ElseIf Me.optOffsetStart Then
    Me.GTMaskFixedStart.Enabled = False
    Me.GTMaskFixedStart.DateValue = Null

    If spnFreqStart.Value > 0 Then
      Me.optFixedEnd.Value = False
      Me.optFixedEnd.Enabled = False
      Me.GTMaskFixedEnd.Enabled = False
      Me.GTMaskFixedEnd.DateValue = Null
      Me.optCurrentEnd.Value = False
      Me.optCurrentEnd.Enabled = False
      If (Not Me.optCustomEnd.Value) Then
        Me.optOffsetEnd.Value = True
      End If
    Else
      Me.optFixedEnd.Enabled = (Not mblnReadOnly)
      Me.GTMaskFixedEnd.Enabled = ((Not mblnReadOnly) And (Me.optFixedEnd.Value))
      Me.optCurrentEnd.Enabled = (Not mblnReadOnly)
    End If
    
    Me.spnFreqStart.Enabled = (Not mblnReadOnly)
    Me.spnFreqStart.Text = CStr(Me.spnFreqStart.Value)
    Me.cboPeriodStart.Enabled = (Not mblnReadOnly)
    If Me.cboPeriodStart.ListIndex < 0 Then
      Me.cboPeriodStart.ListIndex = 0
    End If
    Me.optOffsetEnd.Enabled = (Not mblnReadOnly)
    Me.spnFreqEnd.Enabled = ((Not mblnReadOnly) And (Me.optOffsetEnd.Value))
    Me.cboPeriodEnd.Enabled = ((Not mblnReadOnly) And (Me.optOffsetEnd.Value))
    Me.cmdCustomStart.Enabled = False
    Me.txtCustomStart.Text = ""
    Me.txtCustomStart.Tag = 0
    
  ElseIf Me.optCustomStart Then
    Me.GTMaskFixedStart.Enabled = False
    Me.GTMaskFixedStart.DateValue = Null
    Me.spnFreqStart.Enabled = False
    Me.spnFreqStart.Text = ""
    Me.cboPeriodStart.Enabled = False
    Me.cboPeriodStart.ListIndex = -1
    Me.optOffsetEnd.Enabled = (Not mblnReadOnly)
    Me.optCurrentEnd.Enabled = (Not mblnReadOnly)
    Me.optFixedEnd.Enabled = (Not mblnReadOnly)
    Me.cmdCustomStart.Enabled = (Not mblnReadOnly)
  End If

  ' update the end date frame
  If Me.optFixedEnd Then
    Me.GTMaskFixedEnd.Enabled = (Not mblnReadOnly)
    Me.spnFreqEnd.Enabled = False
    Me.cboPeriodEnd.Enabled = False
    Me.optCurrentStart.Enabled = (Not mblnReadOnly)
    Me.optOffsetStart.Enabled = (Not mblnReadOnly)
    Me.spnFreqEnd.Enabled = False
    Me.spnFreqEnd.Text = ""
    Me.cboPeriodEnd.Enabled = False
    Me.cboPeriodEnd.ListIndex = -1
    Me.cmdCustomEnd.Enabled = False
    Me.txtCustomEnd.Text = ""
    Me.txtCustomEnd.Tag = 0
    
  ElseIf Me.optCurrentEnd Then
    Me.GTMaskFixedEnd.Enabled = False
    Me.GTMaskFixedEnd.DateValue = Null
    Me.spnFreqEnd.Enabled = False
    Me.cboPeriodEnd.Enabled = False
    Me.optOffsetStart.Enabled = (Not mblnReadOnly)
    Me.optCurrentStart.Enabled = (Not mblnReadOnly)
    Me.spnFreqEnd.Enabled = False
    Me.spnFreqEnd.Text = ""
    Me.cboPeriodEnd.Enabled = False
    Me.cboPeriodEnd.ListIndex = -1
    Me.cmdCustomEnd.Enabled = False
    Me.txtCustomEnd.Text = ""
    Me.txtCustomEnd.Tag = 0
    
  ElseIf Me.optOffsetEnd Then
    Me.GTMaskFixedEnd.Enabled = False
    Me.GTMaskFixedEnd.DateValue = Null
    Me.spnFreqEnd.Enabled = (Not mblnReadOnly)
    Me.spnFreqEnd.Text = CStr(Me.spnFreqEnd.Value)
    Me.cboPeriodEnd.Enabled = (Not mblnReadOnly)
    If Me.cboPeriodEnd.ListIndex < 0 Then
      Me.cboPeriodEnd.ListIndex = 0
    End If
    Me.optOffsetStart.Enabled = (Not mblnReadOnly)
    Me.spnFreqStart.Enabled = ((Not mblnReadOnly) And (Me.optOffsetStart.Value))
    Me.cboPeriodStart.Enabled = ((Not mblnReadOnly) And (Me.optOffsetStart.Value))
    Me.optCurrentStart.Enabled = (Not mblnReadOnly)
    Me.cmdCustomEnd.Enabled = False
    Me.txtCustomEnd.Text = ""
    Me.txtCustomEnd.Tag = 0
    
  ElseIf Me.optCustomEnd Then
    Me.GTMaskFixedEnd.Enabled = False
    Me.GTMaskFixedEnd.DateValue = Null
    Me.spnFreqEnd.Enabled = False
    Me.cboPeriodEnd.Enabled = False
    Me.optOffsetStart.Enabled = (Not mblnReadOnly)
    Me.optCurrentStart.Enabled = (Not mblnReadOnly)
    Me.spnFreqEnd.Enabled = False
    Me.spnFreqEnd.Text = ""
    Me.cboPeriodEnd.Enabled = False
    Me.cboPeriodEnd.ListIndex = -1
    Me.cmdCustomEnd.Enabled = (Not mblnReadOnly)
  End If

  Me.GTMaskFixedStart.BackColor = IIf(Me.GTMaskFixedStart.Enabled, vbWindowBackground, vbButtonFace)
  Me.GTMaskFixedEnd.BackColor = IIf(Me.GTMaskFixedEnd.Enabled, vbWindowBackground, vbButtonFace)
  Me.spnFreqStart.BackColor = IIf(Me.spnFreqStart.Enabled, vbWindowBackground, vbButtonFace)
  Me.cboPeriodStart.BackColor = IIf(Me.cboPeriodStart.Enabled, vbWindowBackground, vbButtonFace)
  Me.spnFreqEnd.BackColor = IIf(Me.spnFreqEnd.Enabled, vbWindowBackground, vbButtonFace)
  Me.cboPeriodEnd.BackColor = IIf(Me.cboPeriodEnd.Enabled, vbWindowBackground, vbButtonFace)
  
  chkIncludeBHols.Enabled = ((Not mblnReadOnly) _
                            And ((cboBaseTable.ItemData(cboBaseTable.ListIndex) = glngPersonnelTableID) _
                                Or (cboRegion.ItemData(cboRegion.ListIndex) > 0)) _
                            And (chkGroupByDesc.Value = False))
  
  chkIncludeWorkingDaysOnly.Enabled = ((Not mblnReadOnly) _
                                      And (cboBaseTable.ItemData(cboBaseTable.ListIndex) = glngPersonnelTableID) _
                                      And (chkGroupByDesc.Value = False))

  chkShadeBHols.Enabled = ((Not mblnReadOnly) _
                            And ((cboBaseTable.ItemData(cboBaseTable.ListIndex) = glngPersonnelTableID) _
                                Or (cboRegion.ItemData(cboRegion.ListIndex) > 0)) _
                            And (chkGroupByDesc.Value = False))

  chkGroupByDesc.Enabled = ((Not mblnReadOnly) _
                            And (chkIncludeBHols.Value = False) _
                            And (chkIncludeWorkingDaysOnly.Value = False) _
                            And (chkShadeBHols.Value = False) _
                            And (cboRegion.ItemData(cboRegion.ListIndex) < 1))
  
  cboRegion.Enabled = ((Not mblnReadOnly) _
                        And (chkGroupByDesc.Value = False))
  cboRegion.BackColor = IIf(cboRegion.Enabled, vbWindowBackground, vbButtonFace)
  lblRegion.Enabled = cboRegion.Enabled
  
  mblnLoading = False
  
End Sub
Public Function ValidateDefinition(lngCurrentID As Long) As Boolean

  Dim LLoop As Long
  Dim bm As Variant
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim strRecSelStatus As String
  Dim iCount As Integer
  
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
  Call UtilityAmended(utlCalendarReport, mlngCalendarReportID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName.Text = gsUserName
    
    'JPD 20030815 Fault 6698
    mblnDefinitionCreator = True
    
    mlngCalendarReportID = 0
    mblnReadOnly = False
    ForceAccess
  End If
  
  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngCalendarReportID) Then
    MsgBox "A Calendar Report definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    Exit Function
  End If
  
  ' BASE TABLE - If using a picklist, check one has been selected
  If optBasePicklist.Value Then
    If txtBasePicklist.Text = "" Or txtBasePicklist.Tag = "0" Or txtBasePicklist.Tag = "" Then
      MsgBox "You must select a picklist, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Calendar Reports"
      SSTab1.Tab = 0
      cmdBasePicklist.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' BASE TABLE - If using a filter, check one has been selected
  If optBaseFilter.Value Then
    If txtBaseFilter.Text = "" Or txtBaseFilter.Tag = "0" Or txtBaseFilter.Tag = "" Then
      MsgBox "You must select a filter, or change the record selection for your base table.", vbExclamation + vbOKOnly, "Calendar Reports"
      SSTab1.Tab = 0
      cmdBaseFilter.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  ' Check that a valid description column or a valid calculation has been selected
  If (cboDesc1.ItemData(cboDesc1.ListIndex) < 1) And (txtDescExpr.Tag < 1) And (cboDesc2.ItemData(cboDesc2.ListIndex) < 1) Then
    MsgBox "You must select at least one base description column or calculation for the report.", vbExclamation + vbOKOnly, "Calendar Reports"
    ValidateDefinition = False
    SSTab1.Tab = 0
    cboDesc1.SetFocus
    Exit Function
  End If
  
  
  ' Check that at least 1 column has been defined as the report order
  With grdEvents
    If .Rows = 0 Then
      MsgBox "You must select at least one event to report on.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 1
      cmdAddEvent.SetFocus
      Exit Function
    End If
  End With

  '******************************************************************************
  '                 Validate the Start & End Date Selections
  
  'check fixed start date is not empty (or null)
  If optFixedStart.Value Then
    If IsNull(GTMaskFixedStart.DateValue) Or IsEmpty(GTMaskFixedStart.DateValue) Then
      MsgBox "You must select a Fixed Start Date for the report.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      GTMaskFixedStart.SetFocus
      Exit Function
    End If
  End If
  
  'check fixed end date is not empty (or null)
  If optFixedEnd.Value Then
    If IsNull(GTMaskFixedEnd.DateValue) Or IsEmpty(GTMaskFixedEnd.DateValue) Then
      MsgBox "You must select a Fixed End Date for the report.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      GTMaskFixedEnd.SetFocus
      Exit Function
    End If
  End If

  'check fixed end date is greater than or equal to the fixed start date
  If optFixedStart.Value And optFixedEnd.Value Then
    If GTMaskFixedEnd.DateValue < GTMaskFixedStart.DateValue Then
      MsgBox "You must select a Fixed End Date later than or equal to the Fixed Start Date.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      GTMaskFixedEnd.SetFocus
      Exit Function
    End If
  End If
  
  'check the end date offset is >= zero when fixed start is selected
  If optFixedStart.Value And optOffsetEnd.Value Then
    If spnFreqEnd.Value < 0 Then
      MsgBox "You must select an End Date Offset greater than or equal to zero.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      spnFreqEnd.SetFocus
      Exit Function
    End If
  End If
  
  'check the end date offset is >= zero when current start is selected
  If optCurrentStart.Value And optOffsetEnd.Value Then
    If spnFreqEnd.Value < 0 Then
      MsgBox "You must select an End Date Offset greater than or equal to zero.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      spnFreqEnd.SetFocus
      Exit Function
    End If
  End If
  
  'check the start date offset is <= zero when fixed end or current end is selected
  If optOffsetStart.Value And (optFixedEnd.Value Or optCurrentEnd.Value) Then
    If spnFreqStart.Value > 0 Then
      MsgBox "You must select a Start Date Offset less than or equal to zero.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      spnFreqStart.SetFocus
      Exit Function
    End If
  End If
  
  'check if both start and end date offsets are selected and do the validation on them
  If optOffsetStart.Value And optOffsetEnd.Value Then
    
    'check the end date offset period (eg.days) is the same as the start date offset period
    If cboPeriodStart.ListIndex <> cboPeriodEnd.ListIndex Then
      MsgBox "You must select the same End Date Offset period as Start Date Offset period.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      cboPeriodEnd.SetFocus
      Exit Function
    End If
    
    'check the end date offset is >= start date offset
    If spnFreqEnd.Value < spnFreqStart.Value Then
      MsgBox "You must select an End Date Offset greater than or equal to the Start Date Offset.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      spnFreqEnd.SetFocus
      Exit Function
    End If
  
  End If
  
  If optCustomStart.Value Then
    If CLng(txtCustomStart.Tag) < 1 Then
      MsgBox "You must select a calculation for the Report Start Date.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      cmdCustomStart.SetFocus
      Exit Function
    End If
  End If
  
  If optCustomEnd.Value Then
    If CLng(txtCustomEnd.Tag) < 1 Then
      MsgBox "You must select a calculation for the Report End Date.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 2
      cmdCustomEnd.SetFocus
      Exit Function
    End If
  End If
  
  '******************************************************************************

  ' Check that at least 1 column has been defined as the report order
  With grdOrder
    If .Rows = 0 Then
      MsgBox "You must select at least one column to order the report by.", vbExclamation + vbOKOnly, "Calendar Reports"
      ValidateDefinition = False
      SSTab1.Tab = 3
      cmdNewOrder.SetFocus
      Exit Function
    End If
    
    
    'if group by description is checked that the sort order reflects the base description.
    'should still allow save if the user clicks ok to continue.
    If chkGroupByDesc.Value And txtDescExpr.Tag < 1 Then
      Dim lngDesc1ID As Long
      Dim lngDesc2ID As Long
      Dim strDesc As String
      Dim strTemp As String
      
      lngDesc1ID = cboDesc1.ItemData(cboDesc1.ListIndex)
      lngDesc2ID = cboDesc2.ItemData(cboDesc2.ListIndex)
      strDesc = lngDesc1ID
      If lngDesc2ID > 0 Then
        strDesc = strDesc & vbTab & lngDesc2ID
      End If
      strTemp = vbNullString
      .MoveFirst
      For iCount = 0 To (.Rows - 1) Step 1
        If (.Columns("ColumnID").Text > 0) Then
          If iCount > 0 Then
            strTemp = strTemp & vbTab
          End If
          strTemp = strTemp & .Columns("ColumnID").Value
        End If
        If iCount >= 1 Then
          Exit For
        End If
        .MoveNext
      Next iCount
      
      If strTemp <> strDesc Then
        If MsgBox("The sort order does not reflect the selected Group By Description columns. Do you wish to continue?", vbYesNo + vbInformation, Me.Caption) = vbNo Then
          ValidateDefinition = False
          SSTab1.Tab = 3
          cmdDeleteOrder.SetFocus
          Exit Function
        End If
      End If
    End If
  End With
 
  If Not objOutputDef.ValidDestination Then
    SSTab1.Tab = 4
    Exit Function
  End If

  If Not ForceDefinitionToBeHiddenIfNeeded(True) Then
    Exit Function
  End If
  
If mlngCalendarReportID > 0 Then
  sHiddenGroups = HiddenGroups
  If (Len(sHiddenGroups) > 0) And _
    (UCase(gsUserName) = UCase(txtUserName.Text)) Then
    
    CheckCanMakeHiddenInBatchJobs utlCalendarReport, _
      CStr(mlngCalendarReportID), _
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
               vbExclamation + vbOKOnly, "Calendar Reports"
      Else
        MsgBox "This definition cannot be made hidden as it is used in the following" & vbCrLf & _
               "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
               , "Calendar Reports"
      End If

      Screen.MousePointer = vbNormal
      SSTab1.Tab = 0
      Exit Function

    ElseIf (iCount_Owner > 0) Then
      If MsgBox("Making this definition hidden to user groups will automatically" & vbCrLf & _
                "make the following definition(s), of which you are the" & vbCrLf & _
                "owner, hidden to the same user groups:" & vbCrLf & vbCrLf & _
                sBatchJobDetails_Owner & vbCrLf & _
                "Do you wish to continue ?", vbQuestion + vbYesNo, "Calendar Reports") = vbNo Then
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

  'checks that the definition name does not already exist in the db.
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSysCalendarReports " & _
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
Private Function SaveDefinition() As Boolean

  On Error GoTo Save_ERROR
  
  Dim sSQL As String
  Dim iLoop As Integer
  Dim sKey As String
  Dim objEvent As clsCalendarEvent
  Dim strDescSeparator As String
  
  ReportDateSetup

  Select Case cboDescriptionSeparator.Text
    Case "<None>": strDescSeparator = ""
    Case "<Space>": strDescSeparator = " "
    Case Else: strDescSeparator = cboDescriptionSeparator.Text
  End Select

'*** Step 1 Of 3 - Save the report level information ***
  
  If mlngCalendarReportID > 0 Then
    ' Construct the SQL Update string (Editing an existing definition)
    sSQL = "UPDATE ASRSYSCalendarReports SET " & _
           "  Name = '" & Trim(Replace(txtName.Text, "'", "''")) & "'," & _
           "  Description = '" & Replace(txtDesc.Text, "'", "''") & "'," & _
           "  BaseTable = " & CStr(cboBaseTable.ItemData(cboBaseTable.ListIndex)) & "," & _
           "  AllRecords = " & IIf(Me.optBaseAllRecords.Value, 1, 0) & "," & _
           "  Picklist = " & IIf(Me.optBasePicklist.Value, Me.txtBasePicklist.Tag, 0) & "," & _
           "  Filter = " & IIf(Me.optBaseFilter.Value, Me.txtBaseFilter.Tag, 0) & "," & _
           "  Description1 = " & CStr(cboDesc1.ItemData(cboDesc1.ListIndex)) & ", " & _
           "  Description2 = " & CStr(cboDesc2.ItemData(cboDesc2.ListIndex)) & ", " & _
           "  DescriptionExpr = " & CStr(txtDescExpr.Tag) & ", " & _
           "  DescriptionSeparator = " & "'" & strDescSeparator & "', " & _
           "  Region = " & CStr(cboRegion.ItemData(cboRegion.ListIndex)) & ", " & _
           "  GroupByDesc = " & IIf(chkGroupByDesc.Value = vbChecked, 1, 0) & ", "
    
    sSQL = sSQL & "  StartType = " & CStr(mintStartType) & ", " & _
           "  FixedStart = '" & mstrFixedStart & "', " & _
           "  StartFrequency = " & CStr(mintStartFreq) & ", " & _
           "  StartPeriod = " & CStr(mintStartPeriod) & ", " & _
           "  StartDateExpr = " & CStr(mlngCustomStart) & ", " & _
           "  EndType = " & CStr(mintEndType) & ", " & _
           "  FixedEnd = '" & mstrFixedEnd & "', " & _
           "  EndFrequency = " & CStr(mintEndFreq) & ", " & _
           "  EndPeriod = " & CStr(mintEndPeriod) & ", " & _
           "  EndDateExpr = " & CStr(mlngCustomEnd) & ", " & _
           "  ShowBankHolidays = " & chkShadeBHols.Value & ", " & _
           "  ShowCaptions = " & chkCaptions.Value & ", " & _
           "  ShowWeekends = " & chkShadeWeekends.Value & ", " & _
           "  IncludeWorkingDaysOnly = " & chkIncludeWorkingDaysOnly.Value & ", " & _
           "  IncludeBankHolidays = " & chkIncludeBHols.Value & ", " & _
           "  StartOnCurrentMonth = " & chkStartOnCurrentMonth.Value & ", "
    sSQL = sSQL & " PrintFilterHeader = " & chkPrintFilterHeader.Value & ", "
    
    sSQL = sSQL & _
        " OutputPreview = " & IIf(chkPreview.Value = vbChecked, "1", "0") & ", " & _
        " OutputFormat = " & CStr(objOutputDef.GetSelectedFormatIndex) & ", " & _
        " OutputScreen = " & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", " & _
        " OutputPrinter = " & IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0") & ", " & _
        " OutputPrinterName = '" & Replace(cboPrinterName.Text, " '", "''") & "', "
      
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
        "OutputFilename = '" & Replace(txtFilename.Text, "'", "''") & "' "
            
    sSQL = sSQL & "WHERE ID = " & mlngCalendarReportID
          
    datData.ExecuteSql (sSQL)
    
    Call UtilUpdateLastSaved(utlCalendarReport, mlngCalendarReportID)
    
  Else
  
    ' Construct the SQL Insert string (Adding a new definition)
    
    sSQL = "INSERT ASRSYSCalendarReports (" & _
           "Name, Description, BaseTable, " & _
           "AllRecords, Picklist, Filter, " & _
           "UserName, Description1, " & _
           "Description2, DescriptionExpr, DescriptionSeparator, Region, GroupByDesc, " & _
           "StartType, FixedStart, StartFrequency, " & _
           "StartPeriod, StartDateExpr, EndType, FixedEnd, EndFrequency, " & _
           "EndPeriod, EndDateExpr, ShowBankHolidays, ShowCaptions, " & _
           "ShowWeekends, IncludeWorkingDaysOnly, IncludeBankHolidays, StartOnCurrentMonth, PrintFilterHeader, OutputPreview, OutputFormat, OutputScreen, OutputPrinter, " & _
           "OutputPrinterName, OutputSave, OutputSaveFormat, OutputSaveExisting, OutputEmail, " & _
           "OutputEmailAddr, OutputEmailSubject, OutputEmailAttachAs, OutputEmailFileFormat, OutputFileName " & _
           ") "
    
    sSQL = sSQL & "VALUES("
    sSQL = sSQL & "'" & Trim(Replace(txtName.Text, "'", "''")) & "',"       'Name
    sSQL = sSQL & "'" & Replace(txtDesc.Text, "'", "''") & "',"             'Description
    sSQL = sSQL & CStr(cboBaseTable.ItemData(cboBaseTable.ListIndex)) & "," 'BaseTableID
    
    sSQL = sSQL & IIf(Me.optBaseAllRecords.Value, 1, 0) & ","                     'AllRecords
    sSQL = sSQL & IIf(Me.optBasePicklist.Value, Me.txtBasePicklist.Tag, 0) & ","  'Picklist
    sSQL = sSQL & IIf(Me.optBaseFilter.Value, Me.txtBaseFilter.Tag, 0) & ","      'Filter
    
    sSQL = sSQL & "'" & datGeneral.UserNameForSQL & "',"                                   'Username
    sSQL = sSQL & CStr(cboDesc1.ItemData(cboDesc1.ListIndex)) & ","         'Description1
    sSQL = sSQL & CStr(cboDesc2.ItemData(cboDesc2.ListIndex)) & ","         'Description2
    sSQL = sSQL & CStr(txtDescExpr.Tag) & ","                               'DescriptionExpr
    sSQL = sSQL & "'" & strDescSeparator & "',"                             'Description Separator
    sSQL = sSQL & CStr(cboRegion.ItemData(cboRegion.ListIndex)) & ","       'Region
    sSQL = sSQL & IIf(chkGroupByDesc.Value = vbChecked, 1, 0) & ","         'GroupByDesc
    
    sSQL = sSQL & CStr(mintStartType) & ","                                 'StartType
    sSQL = sSQL & "'" & mstrFixedStart & "',"                               'FixedStart
    sSQL = sSQL & CStr(mintStartFreq) & ","                                 'StartFrequency
    sSQL = sSQL & CStr(mintStartPeriod) & ","                               'StartPeriod
    sSQL = sSQL & CStr(mlngCustomStart) & ","                               'StartDateExpr
    sSQL = sSQL & CStr(mintEndType) & ","                                   'EndType
    sSQL = sSQL & "'" & mstrFixedEnd & "',"                                 'FixedEnd
    sSQL = sSQL & CStr(mintEndFreq) & ","                                   'EndFrequency
    sSQL = sSQL & CStr(mintEndPeriod) & ","                                 'EndPeriod
    sSQL = sSQL & CStr(mlngCustomEnd) & ","                                 'EndDateExpr
    
    sSQL = sSQL & chkShadeBHols.Value & ","                                 'ShowBankHolidays
    sSQL = sSQL & chkCaptions.Value & ","                                   'ShowCaptions
    sSQL = sSQL & chkShadeWeekends.Value & ","                              'ShowWeekends
    sSQL = sSQL & chkIncludeWorkingDaysOnly.Value & ","                     'IncludeWorkingDaysOnly
    sSQL = sSQL & chkIncludeBHols.Value & ","                               'IncludeBankHolidays
    sSQL = sSQL & chkStartOnCurrentMonth.Value & ","                        'StartOnCurrentMonth
    sSQL = sSQL & chkPrintFilterHeader.Value & ","                          'PrintFilterHeader
    
    'Output Options
    sSQL = sSQL & CStr(IIf(chkPreview.Value = vbChecked, "1", "0")) & ","                 'OutputPreview
    sSQL = sSQL & CStr(objOutputDef.GetSelectedFormatIndex) & ","           'OutputFormat
    sSQL = sSQL & CStr(IIf(chkDestination(desScreen).Value = vbChecked, "1", "0")) & ","  'OutputScreen
    sSQL = sSQL & CStr(IIf(chkDestination(desPrinter).Value = vbChecked, "1", "0")) & "," 'OutputPrinter
    sSQL = sSQL & "'" & Replace(cboPrinterName.Text, "'", "''") & "',"              'OutputPrinterName

    If chkDestination(desSave).Value = vbChecked Then
      sSQL = sSQL & "1, " & CStr(Val(txtFilename.Tag)) & ", " & _
        cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "    'OutputSave, OutputSaveFormat, OutputSaveExisting
    Else
      sSQL = sSQL & "0, 0, 0, "    'OutputSave, OutputSaveFormat, OutputSaveExisting
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
        "'" & Replace(txtFilename.Text, "'", "''") & "'"  'OutputFilename
        
    sSQL = sSQL & ")"
    
    mlngCalendarReportID = InsertCalendarReport(sSQL)
    
    Call UtilCreated(utlCalendarReport, mlngCalendarReportID)
  
  End If
  
  SaveAccess

  '*** Step 2 Of 3 - Save the event level information
  ' First, remove any records from the event details table which has the
  ' current CalendarReportID
  ClearEventsTable mlngCalendarReportID
  Set objEvent = New clsCalendarEvent
  
  For Each objEvent In mcolEvents.Collection
    With objEvent
      sSQL = "INSERT ASRSysCalendarReportEvents (EventKey, " & _
             "CalendarReportID, Name, TableID, FilterID, " & _
             "EventStartDateID, EventStartSessionID, EventEndDateID, EventEndSessionID, " & _
             "EventDurationID, LegendType, LegendCharacter, LegendLookupTableID, LegendLookupColumnID, " & _
             "LegendLookupCodeID, LegendEventColumnID, Colour, " & _
             "EventDesc1ColumnID, EventDesc2ColumnID) "
  
  
        sSQL = sSQL & "VALUES('" & .Key & "',"
        sSQL = sSQL & CStr(mlngCalendarReportID) & ","
        sSQL = sSQL & "'" & Trim(Replace(.Name, "'", "''")) & "',"
        sSQL = sSQL & CStr(.TableID) & ","
        sSQL = sSQL & CStr(.FilterID) & ","
        sSQL = sSQL & CStr(.StartDateID) & ","
        sSQL = sSQL & CStr(.StartSessionID) & ","
        sSQL = sSQL & CStr(.EndDateID) & ","
        sSQL = sSQL & CStr(.EndSessionID) & ","
        sSQL = sSQL & CStr(.DurationID) & ","
        sSQL = sSQL & CStr(.LegendType) & ","
  
        If .LegendType = 1 Then
          sSQL = sSQL & "'',"
          sSQL = sSQL & CStr(.LegendTableID) & ","
          sSQL = sSQL & CStr(.LegendColumnID) & ","
          sSQL = sSQL & CStr(.LegendCodeID) & ","
          sSQL = sSQL & CStr(.LegendEventTypeID) & ","
          sSQL = sSQL & "0,"
        Else
          sSQL = sSQL & "'" & Replace(.LegendCharacter, "'", "''") & "',"
          sSQL = sSQL & "0,"
          sSQL = sSQL & "0,"
          sSQL = sSQL & "0,"
          sSQL = sSQL & "0,"
          sSQL = sSQL & "'" & CStr(.ColourValue) & "',"
  
        End If
        
        sSQL = sSQL & CStr(.Description1ID) & ","
        sSQL = sSQL & CStr(.Description2ID) & ")"
  
      datData.ExecuteSql (sSQL)
  
    End With
  Next objEvent
 
  
  '*** Step 3 Of 3 - Save the sort order information
  
  ' First, remove any records from the sort order table
  ClearOrderTable mlngCalendarReportID

  With grdOrder
    .Redraw = False
    .MoveFirst
    For iLoop = 0 To .Rows - 1 Step 1

      sSQL = "INSERT ASRSysCalendarReportOrder (" & _
             "CalendarReportID, TableID, " & _
             "ColumnID, OrderSequence, OrderType) "

      sSQL = sSQL & " VALUES ("
      sSQL = sSQL & CStr(mlngCalendarReportID) & ","
      sSQL = sSQL & CStr(cboBaseTable.ItemData(cboBaseTable.ListIndex)) & ","
      sSQL = sSQL & CStr(.Columns("ColumnID").Value) & ","
      sSQL = sSQL & CStr(iLoop) & ","
      sSQL = sSQL & "'" & .Columns("Order").Value & "')"

      datData.ExecuteSql (sSQL)

      .MoveNext
    Next iLoop
    .Redraw = True
  End With
  
  SaveDefinition = True
  Changed = False
  
  Exit Function

Save_ERROR:

  SaveDefinition = False
  MsgBox "Warning : An error has occurred whilst saving the Calendar Report. " & vbCrLf & Err.Description & vbCrLf & "Please cancel and try again. If this error continues, delete the definition.", vbCritical + vbOKOnly, "Calendar Reports"

End Function
Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysCalendarReportAccess WHERE ID = " & mlngCalendarReportID
  datData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysCalendarReportAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngCalendarReportID & ", sysusers.name," & _
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
      sSQL = "IF EXISTS (SELECT * FROM ASRSysCalendarReportAccess" & _
        " WHERE ID = " & CStr(mlngCalendarReportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysCalendarReportAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngCalendarReportID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      datData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub




Private Sub ClearEventsTable(plngCalendarReportID As Long)

  ' Delete all column information from the Calendar Events table.
  
  Dim sSQL As String
  
  sSQL = "Delete From ASRSysCalendarReportEvents Where CalendarReportID = " & plngCalendarReportID
  datData.ExecuteSql sSQL

End Sub
Private Sub ClearOrderTable(plngCalendarReportID As Long)

  ' Delete all column information from the Calendar Order table.
  
  Dim sSQL As String
  
  sSQL = "Delete From ASRSysCalendarReportOrder Where CalendarReportID = " & plngCalendarReportID
  datData.ExecuteSql sSQL

End Sub
Private Function InsertCalendarReport(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo ErrorTrap

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
    pmADO.Value = "ASRSysCalendarReports"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "ID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertCalendarReport = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertCalendarReport = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
ErrorTrap:
  
  fSavedOK = False
  Resume Next

End Function
Private Sub ClearForNew()
  
  'Clear out all fields required to be blank for a new report definition
  
  ' Def Tab
  optBaseAllRecords.Value = True
  txtBasePicklist.Text = ""
  txtBasePicklist.Tag = 0
  txtBaseFilter.Text = ""
  txtBaseFilter.Tag = 0
  txtDescExpr.Text = ""
  txtDescExpr.Tag = 0
  
  If mblnDefinitionCreator Then txtUserName.Text = gsUserName
    
  ' Event Tab
  grdEvents.RemoveAll
  cmdEditEvent.Enabled = False
  cmdRemoveEvent.Enabled = False
  cmdRemoveAllEvents.Enabled = False
  Set mcolEvents = Nothing
  Set mcolEvents = New clsCalendarEvents
  
  'Report Details Tab
  chkIncludeBHols.Value = vbUnchecked
  chkIncludeWorkingDaysOnly.Value = vbUnchecked
  chkShadeBHols.Value = vbUnchecked
  
  ' Order Tab
  grdOrder.RemoveAll
  cmdEditOrder.Enabled = False
  cmdDeleteOrder.Enabled = False
  cmdSortMoveUp.Enabled = False
  cmdSortMoveDown.Enabled = False
  
'  ' Output Tab
'  objOutputDef.FormatClick (0)
  
End Sub
Private Sub cboBaseTable_Click()
  
  Dim bBaseTableChanged As Boolean
  
  If mblnLoading = True Then Exit Sub
  If mstrBaseTable = Me.cboBaseTable.Text And (mblnLoading = False) Then Exit Sub

  If Me.grdEvents.Rows > 0 Or Me.grdOrder.Rows > 0 Then
    If MsgBox("Warning: Changing the base table will result in all table/column " & _
          "specific aspects of this report definition being cleared." & vbCrLf & _
          "Are you sure you wish to continue?", _
          vbQuestion + vbYesNo + vbDefaultButton2, "Calendar Reports") = vbYes Then
  
      bBaseTableChanged = True
    
    Else
      ' User opted to abort the base table change
      SetComboText cboBaseTable, mstrBaseTable
      Exit Sub
    
    End If
  Else
    bBaseTableChanged = True
    
  End If
  
  If bBaseTableChanged Then
    mblnLoading = True
    ClearForNew
    mblnLoading = False
    Changed = True
  End If
  
  mstrBaseTable = Me.cboBaseTable.Text
  mlngBaseTableID = Me.cboBaseTable.ItemData(Me.cboBaseTable.ListIndex)
  
  optBaseAllRecords.Value = True
  
  chkIncludeBHols.Value = vbUnchecked
  chkIncludeWorkingDaysOnly.Value = vbUnchecked
  chkShadeBHols.Value = vbUnchecked

  ForceDefinitionToBeHiddenIfNeeded
  
  UpdateDependantFields
  
  EnableDisableTabControls

End Sub
Private Sub cboDesc1_Click()
  If Not mblnLoading Then
    Changed = True
    EnableDisableTabControls
  End If
End Sub
Private Sub cboDesc2_Click()
  If Not mblnLoading Then
    Changed = True
    EnableDisableTabControls
  End If
End Sub

Private Sub cboDescriptionSeparator_Click()
  If Not mblnLoading Then
    Changed = True
    EnableDisableTabControls
  End If
End Sub

Private Sub chkPrintFilterHeader_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub


Private Sub chkStartOnCurrentMonth_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub cmdClearOrder_Click()
  If MsgBox("Are you sure you wish to clear the sort order?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    grdOrder.RemoveAll
    UpdateOrderButtonStatus
    Me.Changed = True
  End If
End Sub

Private Sub cmdCustomEnd_Click()
  GetCustomDate txtCustomEnd
End Sub

Private Sub cmdCustomStart_Click()
  GetCustomDate txtCustomStart
End Sub


Private Sub cmdDescExpr_Click()
  GetExpression Me.cboBaseTable, Me.txtDescExpr
End Sub

Private Sub cmdEmailGroup_Click()
  If Not mblnLoading Then
    Changed = True
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

Private Sub grdOrder_DblClick()
  If Not mblnReadOnly Then
    If grdOrder.Rows > 0 Then
      cmdEditOrder_Click
    Else
      cmdNewOrder_Click
    End If
  End If
End Sub

Private Sub GTMaskFixedEnd_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub GTMaskFixedEnd_LostFocus()
  ValidateGTMaskDate GTMaskFixedEnd
End Sub


Private Sub GTMaskFixedStart_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub GTMaskFixedStart_LostFocus()
  ValidateGTMaskDate GTMaskFixedStart
End Sub





Private Sub optCustomEnd_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub

Private Sub optCustomStart_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If Not mblnLoading Then
    EnableDisableTabControls
  End If
End Sub



Private Sub txtCustomEnd_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub txtCustomStart_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub




Private Sub txtEmailGroup_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub cboPeriodEnd_Click()
  If Not mblnLoading Then
    If (cboPeriodEnd.ListIndex <> cboPeriodStart.ListIndex) And (optOffsetStart.Value) And (optOffsetEnd.Value) Then
      MsgBox "The End Date Offset period must be the same as the Start Date Offset period", vbExclamation + vbOKOnly, "Calendar Reports"
      mblnLoading = True
      cboPeriodEnd.ListIndex = cboPeriodStart.ListIndex
      mblnLoading = False
    End If
    Changed = True
  End If
End Sub
Private Sub cboPeriodStart_Click()
  If Not mblnLoading Then
    cboPeriodEnd.ListIndex = cboPeriodStart.ListIndex
    UpdateReportDetailsTab
    Changed = True
  End If
End Sub

Private Sub cboPrinterName_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub cboRegion_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    
    Changed = True
  End If
End Sub

Private Sub cboSaveExisting_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub


Private Sub chkCaptions_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub chkGroupByDesc_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    
    Changed = True
  End If
End Sub

Private Sub chkIncludeBHols_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    
    Changed = True
  End If
End Sub

Private Sub chkIncludeWorkingDaysOnly_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    
    Changed = True
  End If
End Sub

Private Sub chkPreview_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub chkShadeBHols_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    
    Changed = True
  End If
End Sub

Private Sub chkShadeWeekends_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cmdAddEvent_Click()

  Dim plngRow As Long
  Dim pstrRow As String
  Dim pfrmEvents As New frmCalendarReportDates
  Dim strEventKey As String
  
  strEventKey = GetEventKey
  
  With pfrmEvents
    .Initialize True _
                , Me _
                , mcolEvents _
                , "" _
                , 0 _
                , 0 _
                , "" _
                , 0 _
                , 0 _
                , 0 _
                , 0 _
                , 0 _
                , "" _
                , 0 _
                , 0 _
                , 0 _
                , 0 _
                , 0 _
                , 0 _
                , 0 _
                , strEventKey
        
    If Not .Cancelled Then
      .Show vbModal
    End If
    
    If Not .Cancelled Then
      pstrRow = .EventName & vbTab & _
                .EventTableID & vbTab & _
                .EventTable & vbTab & _
                .EventFilterID & vbTab & _
                .EventFilterName & vbTab & _
                .EventStartDateID & vbTab & _
                .EventStartDateColumn & vbTab & _
                .EventStartSessionID & vbTab & _
                .EventStartSessionColumn & vbTab & _
                .EventEndDateID & vbTab & _
                .EventEndDateColumn & vbTab & _
                .EventEndSessionID & vbTab & _
                .EventEndSessionColumn & vbTab & _
                .EventDurationID & vbTab & _
                .EventDurationColumn & vbTab & _
                .EventLegendType & vbTab & _
                .EventLegendRef & vbTab & _
                .EventLegendTableID & vbTab & _
                .EventLegendColumnID & vbTab & _
                .EventLegendCodeID & vbTab & _
                .LegendEventTypeID & vbTab & _
                .EventDesc1ID & vbTab & _
                .EventDesc1Column & vbTab & _
                .EventDesc2ID & vbTab & _
                .EventDesc2Column & vbTab
      
      pstrRow = pstrRow & .Key & vbTab & _
                .EventColourName & vbTab & _
                .EventColour
      
      With Me.grdEvents
        .AddItem pstrRow
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With
      
      mcolEvents.Add strEventKey, .EventName, _
                  .EventTableID, .EventTable, _
                  .EventFilterID, _
                  .EventStartDateID, .EventStartDateColumn, _
                  .EventStartSessionID, .EventStartSessionColumn, _
                  .EventEndDateID, .EventEndDateColumn, _
                  .EventEndSessionID, .EventEndSessionColumn, _
                  .EventDurationID, .EventDurationColumn, _
                  .EventLegendType, .EventCharacter, .EventColour, _
                  .EventLegendTableID, .EventLegendTable, _
                  .EventLegendColumnID, .EventLegendColumn, _
                  .EventLegendCodeID, .EventLegendCode, _
                  .LegendEventTypeID, .LegendEventType, _
                  .EventDesc1ID, .EventDesc1Column, _
                  .EventDesc2ID, .EventDesc2Column
      
      Changed = True
    Else
      Exit Sub
    End If
    
  End With

  Unload pfrmEvents
  Set pfrmEvents = Nothing

  FormatGridColumnWidths
  
  RefreshEventButtons
  
  ForceDefinitionToBeHiddenIfNeeded
  
  Changed = True

End Sub
Public Sub LoadBaseCombo()

  ' Loads the Base combo with all tables (even lookups)
  
  Dim sSQL As String
  Dim objTable As CTablePrivilege
 
  With cboBaseTable
    .Clear
    For Each objTable In gcoTablePrivileges.Collection
      If objTable.IsTable Then
        .AddItem objTable.TableName
        .ItemData(.NewIndex) = objTable.TableID
      End If
    Next objTable
    If .ListCount > 0 Then
      If gsPersonnelTableName <> "" Then
        SetComboText cboBaseTable, gsPersonnelTableName
      Else
        .ListIndex = 0
      End If

      mstrBaseTable = .List(.ListIndex)
      mlngBaseTableID = .ItemData(.ListIndex)
    End If
  End With

End Sub
Private Sub UpdateDependantFields()

  Dim objColumn As CColumnPrivilege
  Dim i As Integer
  
  i = 0
  
  ' Clear Desc1 combo and add <None> entry
  With cboDesc1
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear Desc2 combo and add <None> entry
  With cboDesc2
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear Region combo and add <None> entry
  With cboRegion
    .Clear
    If mlngBaseTableID = glngPersonnelTableID Then
      .AddItem "<Default>"
    Else
      .AddItem "<None>"
    End If
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  Set mcolColumnPivilages = GetColumnPrivileges(mstrBaseTable)

  For Each objColumn In mcolColumnPivilages
  
    If (objColumn.ColumnType <> ColumnTypes.colSystem) _
      And (objColumn.ColumnType <> ColumnTypes.colLink) _
      And (objColumn.DataType <> SQLDataType.sqlOle) _
      And (objColumn.DataType <> SQLDataType.sqlVarBinary) Then
      
      'populate the Description 1 combo
      cboDesc1.AddItem objColumn.ColumnName
      cboDesc1.ItemData(cboDesc1.NewIndex) = objColumn.ColumnID
  
      'populate the Description 2 combo
      cboDesc2.AddItem objColumn.ColumnName
      cboDesc2.ItemData(cboDesc2.NewIndex) = objColumn.ColumnID
  
      If (objColumn.DataType = SQLDataType.sqlVarChar) Then
        'populate the Region combo
        cboRegion.AddItem objColumn.ColumnName
        cboRegion.ItemData(cboRegion.NewIndex) = objColumn.ColumnID
      End If
    End If
    
    i = i + 1
    
  Next objColumn

  If (cboDesc1.ListCount > 0) Then
    cboDesc1.ListIndex = 0
  End If
  
End Sub
Private Sub cmdBaseFilter_Click()
  GetFilter cboBaseTable, txtBaseFilter
End Sub
Private Sub cmdBasePicklist_Click()
  GetPicklist cboBaseTable, txtBasePicklist
End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdDeleteOrder_Click()
  
  Dim i As Integer
  
  With grdOrder
    If .SelBookmarks.Count < 1 Then
      Exit Sub
    End If
    
    For i = 0 To .SelBookmarks.Count - 1 Step 1
      .RemoveItem (.AddItemRowIndex(.SelBookmarks(i)))
    Next i
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      .MoveLast
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  Changed = True

  UpdateOrderButtonStatus
  
End Sub
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
      
        ctlTarget.Text = IIf(Len(.Name) = 0, "<None>", .Name)
        ctlTarget.Tag = .ExpressionID
        
        Changed = True
      End If
      
    End If
  End With
  
  Set objExpression = Nothing

  ForceDefinitionToBeHiddenIfNeeded

End Sub
Private Sub cmdEditEvent_Click()

  Dim plngRow As Long
  Dim pstrRow As String
  Dim pfrmEvents As New frmCalendarReportDates
  Dim iLegendType As Interior
  Dim sLegendChar As String
  Dim lColour As Long
  Dim strEventKey As String
  
  With Me.grdEvents
    plngRow = .AddItemRowIndex(.Bookmark)
    If .Columns("LegendType").Value = 0 Then
      sLegendChar = .Columns("Legend").Value
      lColour = .Columns("ColourValue").Value
    Else
      sLegendChar = ""
      lColour = 0   'NEEDS TO DEFAULT TO SOMETHING GOOD!
    End If
    
    strEventKey = Trim(.Columns("EventKey").Value)
    
    pfrmEvents.Initialize False _
                , Me _
                , mcolEvents _
                , .Columns("Name").Value _
                , CLng(.Columns("TableID").Value) _
                , CLng(.Columns("FilterID").Value) _
                , .Columns("Filter").Value _
                , CLng(.Columns("StartDateID").Value) _
                , CLng(.Columns("StartSessionID").Value) _
                , CLng(.Columns("EndDateID").Value) _
                , CLng(.Columns("EndSessionID").Value) _
                , CLng(.Columns("DurationID").Value) _
                , sLegendChar _
                , lColour _
                , CLng(.Columns("LegendTableID").Value) _
                , CLng(.Columns("LegendColumnID").Value) _
                , CLng(.Columns("LegendCodeID").Value) _
                , CLng(.Columns("LegendEventTypeID").Value) _
                , CLng(.Columns("Desc1ID").Value) _
                , CLng(.Columns("Desc2ID").Value) _
                , Trim(.Columns("EventKey").Value)
  End With
  
  With pfrmEvents
    If Not .Cancelled Then
      .Show vbModal
    End If
    
    If Not .Cancelled Then
      pstrRow = .EventName & vbTab & _
                .EventTableID & vbTab & _
                .EventTable & vbTab & _
                .EventFilterID & vbTab & _
                .EventFilterName & vbTab & _
                .EventStartDateID & vbTab & _
                .EventStartDateColumn & vbTab & _
                .EventStartSessionID & vbTab & _
                .EventStartSessionColumn & vbTab & _
                .EventEndDateID & vbTab & _
                .EventEndDateColumn & vbTab & _
                .EventEndSessionID & vbTab & _
                .EventEndSessionColumn & vbTab & _
                .EventDurationID & vbTab & _
                .EventDurationColumn & vbTab & _
                .EventLegendType & vbTab & _
                .EventLegendRef & vbTab & _
                .EventLegendTableID & vbTab & _
                .EventLegendColumnID & vbTab & _
                .EventLegendCodeID & vbTab & _
                .LegendEventTypeID & vbTab & _
                .EventDesc1ID & vbTab & _
                .EventDesc1Column & vbTab & _
                .EventDesc2ID & vbTab & _
                .EventDesc2Column & vbTab
                
      pstrRow = pstrRow & .Key & vbTab & _
                .EventColourName & vbTab & _
                .EventColour
      
      With Me.grdEvents
        .RemoveItem plngRow
        .AddItem pstrRow, plngRow
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With
      
      mcolEvents.Remove (strEventKey)
      mcolEvents.Add strEventKey, .EventName, _
                  .EventTableID, .EventTable, _
                  .EventFilterID, _
                  .EventStartDateID, .EventStartDateColumn, _
                  .EventStartSessionID, .EventStartSessionColumn, _
                  .EventEndDateID, .EventEndDateColumn, _
                  .EventEndSessionID, .EventEndSessionColumn, _
                  .EventDurationID, .EventDurationColumn, _
                  .EventLegendType, .EventCharacter, .EventColour, _
                  .EventLegendTableID, .EventLegendTable, _
                  .EventLegendColumnID, .EventLegendColumn, _
                  .EventLegendCodeID, .EventLegendCode, _
                  .LegendEventTypeID, .LegendEventType, _
                  .EventDesc1ID, .EventDesc1Column, _
                  .EventDesc2ID, .EventDesc2Column
                  
      Changed = True
    Else
      Exit Sub
    End If
    
  End With

  Unload pfrmEvents
  Set pfrmEvents = Nothing

  FormatGridColumnWidths
  
  RefreshEventButtons
    
  ForceDefinitionToBeHiddenIfNeeded
  
  Changed = True

End Sub
Private Sub UpdateOrderButtonStatus()

  If mblnReadOnly Then
    Exit Sub
  End If
 
  If grdOrder.Rows = 1 Then
    grdOrder.MoveFirst
  End If
  
  If grdOrder.Rows = 0 Then
    cmdEditOrder.Enabled = False
    cmdDeleteOrder.Enabled = False
    cmdClearOrder.Enabled = False
    cmdNewOrder.Enabled = (Not mblnReadOnly)
    cmdSortMoveDown.Enabled = False
    cmdSortMoveUp.Enabled = False
    grdOrder.ScrollBars = ssScrollBarsNone
  Else
    cmdEditOrder.Enabled = (Not mblnReadOnly)
    cmdDeleteOrder.Enabled = (Not mblnReadOnly)
    cmdClearOrder.Enabled = (Not mblnReadOnly)
    cmdNewOrder.Enabled = (Not mblnReadOnly)
    
    With grdOrder
      .Refresh
      
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
      
      If (CInt(.Rows) > CInt(.VisibleRows)) Then
        If .ScrollBars = ssScrollBarsNone Then
          .ScrollBars = ssScrollBarsVertical
          .Columns(1).Width = .Columns(1).Width - 235
        End If
      Else
        If .ScrollBars = ssScrollBarsVertical Then
          .ScrollBars = ssScrollBarsNone
          .Columns(1).Width = .Columns(1).Width + 235
          .FirstRow = 0
        End If
      End If

    End With
    
  End If

  
End Sub
Private Sub cmdEditOrder_Click()
  
  Dim pfrmOrderEdit As New frmCalendarReportOrder
  Dim lngColumnID As Long
  Dim strSortOrder As String
  Dim sSelectedOrderCols As String
  
  sSelectedOrderCols = SortOrderColumns()
  
  With grdOrder
  
    lngColumnID = .Columns("ColumnID").CellValue(.SelBookmarks(0))
    strSortOrder = .Columns("Order").CellText(.SelBookmarks(0))

    If lngColumnID > 0 Then
      If pfrmOrderEdit.Initialise(False, Me, mcolColumnPivilages, sSelectedOrderCols, lngColumnID, strSortOrder) = True Then
        pfrmOrderEdit.Show vbModal
      End If
    End If
  
  End With
  
  UpdateOrderButtonStatus
  
  'AE20071025 Fault #6797
  If Not pfrmOrderEdit.UserCancelled Then
    Changed = True
  End If
  
  Unload pfrmOrderEdit
  Set pfrmOrderEdit = Nothing
  
End Sub
Private Sub cmdNewOrder_Click()
  
  Dim pfrmOrder As New frmCalendarReportOrder
  Dim sSelectedOrderCols As String
  
  sSelectedOrderCols = SortOrderColumns()
  
  If pfrmOrder.Initialise(True, Me, mcolColumnPivilages, sSelectedOrderCols, 0, "") = True Then
    pfrmOrder.Show vbModal
  End If
  
  UpdateOrderButtonStatus
    
  If Not pfrmOrder.UserCancelled Then
    Changed = True
  End If
  
  'AE20071025 Fault #6797
  Unload pfrmOrder
  Set pfrmOrder = Nothing
 
End Sub
Private Sub cmdOK_Click()

  If Changed = True Then
    Screen.MousePointer = vbHourglass
    
    If Not ValidateDefinition(mlngCalendarReportID) Then
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
Private Sub cmdRemoveAllEvents_Click()
    
  Dim sMessage As String
  
  sMessage = "Are you sure you want to remove all the Events from this Calendar Report definition?"
  
  If MsgBox(sMessage, vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    grdEvents.RemoveAll
    mcolEvents.RemoveAll
    Changed = True
    
    FormatGridColumnWidths
  
    RefreshEventButtons
      
    ForceDefinitionToBeHiddenIfNeeded

  End If

End Sub
Private Sub cmdRemoveEvent_Click()
  
  Dim i As Integer
  Dim sEventKey As String
  
  With grdEvents
    For i = 0 To .SelBookmarks.Count - 1 Step 1
      sEventKey = .Columns("EventKey").CellText(.SelBookmarks(i))
      
      .RemoveItem (.AddItemRowIndex(.SelBookmarks(i)))
      
      mcolEvents.Remove (sEventKey)
      
      Changed = True
    Next i
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      .MoveLast
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  FormatGridColumnWidths

  RefreshEventButtons
    
  ForceDefinitionToBeHiddenIfNeeded

End Sub
Private Sub cmdSortMoveDown_Click()
  
  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdOrder.AddItemRowIndex(grdOrder.Bookmark)
  strSourceRow = grdOrder.Columns(0).Text & vbTab & grdOrder.Columns(1).Text & vbTab & grdOrder.Columns(2).Text
  
  intDestinationRow = intSourceRow + 1
  grdOrder.MoveNext
  strDestinationRow = grdOrder.Columns(0).Text & vbTab & grdOrder.Columns(1).Text & vbTab & grdOrder.Columns(2).Text
  
  grdOrder.RemoveItem intDestinationRow
  grdOrder.RemoveItem intSourceRow
  
  grdOrder.AddItem strDestinationRow, intSourceRow
  grdOrder.AddItem strSourceRow, intDestinationRow
  
  grdOrder.SelBookmarks.RemoveAll
  grdOrder.MoveNext
  grdOrder.Bookmark = grdOrder.AddItemBookmark(intDestinationRow)
  grdOrder.SelBookmarks.Add grdOrder.AddItemBookmark(intDestinationRow)
  
  UpdateOrderButtonStatus

  Changed = True
  
End Sub
Private Sub cmdSortMoveUp_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  intSourceRow = grdOrder.AddItemRowIndex(grdOrder.Bookmark)
  strSourceRow = grdOrder.Columns(0).Text & vbTab & grdOrder.Columns(1).Text & vbTab & grdOrder.Columns(2).Text
  
  intDestinationRow = intSourceRow - 1
  grdOrder.MovePrevious
  strDestinationRow = grdOrder.Columns(0).Text & vbTab & grdOrder.Columns(1).Text & vbTab & grdOrder.Columns(2).Text
  
  grdOrder.AddItem strSourceRow, intDestinationRow
  
  grdOrder.RemoveItem intSourceRow + 1
  
  grdOrder.SelBookmarks.RemoveAll
  grdOrder.MovePrevious
  grdOrder.Bookmark = grdOrder.AddItemBookmark(intDestinationRow)
  grdOrder.SelBookmarks.Add grdOrder.AddItemBookmark(intDestinationRow)

  UpdateOrderButtonStatus
  
  Changed = True

End Sub
Private Sub Form_Load()
  
  Set mcolEvents = New clsCalendarEvents
  
  'JPD 20041117 Fault 8231
  UI.FormatGTDateControl GTMaskFixedEnd
  UI.FormatGTDateControl GTMaskFixedStart
  
  SSTab1.Tab = 0
  grdAccess.RowHeight = 239
  
  Set objOutputDef = New clsOutputDef
  objOutputDef.ParentForm = Me
  objOutputDef.PopulateCombos True, True, True
  objOutputDef.ShowFormats True, False, True, True, True, False, False
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim pintAnswer As Integer
    
  If Changed = True And Not FromPrint Then
    
    pintAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, "Calendar Reports")
      
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

Private Sub grdEvents_Click()
  If Not mblnReadOnly Then
    RefreshEventButtons
  End If
End Sub

Private Sub grdEvents_DblClick()
  If Not mblnReadOnly Then
    If grdEvents.Rows > 0 Then
      cmdEditEvent_Click
    Else
      cmdAddEvent_Click
    End If
  End If
End Sub

Private Sub grdEvents_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  If Not mblnReadOnly Then
    RefreshEventButtons
  End If
End Sub

Private Sub grdEvents_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
  If Not mblnReadOnly Then
    RefreshEventButtons
  End If
End Sub

Private Sub grdOrder_Click()
  If Not mblnReadOnly Then
    UpdateOrderButtonStatus
  End If
End Sub
Private Sub grdOrder_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  
  If Not mblnReadOnly Then
  
    With grdOrder
      ' Set the styleSet of the rows to show which is selected.
      For iLoop = 0 To .Rows
        If iLoop = .Row Then
          .Columns(1).CellStyleSet "ssetActive", iLoop
        Else
          .Columns(1).CellStyleSet "ssetDormant", iLoop
        End If
      Next iLoop
      
      ' Activate the 'values' column.
      If .Col = 1 Then
        .Col = 0
      End If
  
      If .AddItemRowIndex(.Bookmark) = 0 Then
        Me.cmdSortMoveUp.Enabled = False
        Me.cmdSortMoveDown.Enabled = .Rows > 1
      ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
        Me.cmdSortMoveUp.Enabled = .Rows > 1
        Me.cmdSortMoveDown.Enabled = False
      Else
        Me.cmdSortMoveUp.Enabled = True
        Me.cmdSortMoveDown.Enabled = True
      End If
    End With
  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub GTMaskFixedEnd_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    GTMaskFixedEnd.DateValue = Date
  End If
End Sub

Private Sub GTMaskFixedStart_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    GTMaskFixedStart.DateValue = Date
  End If
End Sub



'*** OUTPUT OPTIONS ***
Private Sub optOutputFormat_Click(Index As Integer)
  objOutputDef.FormatClick Index
  
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub chkDestination_Click(Index As Integer)
  objOutputDef.DestinationClick Index
  
  If Not mblnLoading Then
    Changed = True
  End If

End Sub

Private Sub optBaseAllRecords_Click()
  If Not mblnLoading Then
    Changed = True
  
    With txtBasePicklist
      .Text = ""
      .Tag = 0
    End With
    
    cmdBasePicklist.Enabled = False
  
    With txtBaseFilter
      .Text = ""
      .Tag = 0
    End With
  
    EnableDisableTabControls

    cmdBaseFilter.Enabled = False
    ForceDefinitionToBeHiddenIfNeeded
  End If
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
  Dim sEventFilter As String
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
      txtBaseFilter.Tag = 0
      txtBaseFilter.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Check Report Description Calculation
  If (Len(txtDescExpr.Tag) > 0) And (Val(txtDescExpr.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateCalculation(CLng(txtDescExpr.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Calc hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly)
        If fRemove Then
          sBigMessage = "The report description calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "report description calculation"
        End If

      Case REC_SEL_VALID_DELETED
        ' Calc deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "report description calculation"

        fRemove = (Not mblnReadOnly)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Calc hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "report description calculation"
  
          fRemove = (Not mblnReadOnly)
        End If
      Case REC_SEL_VALID_INVALID
        ' Calc invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "report description calculation"

        fRemove = (Not mblnReadOnly)

    End Select

    If fRemove Then
      ' Calc invalid, deleted or hidden by another user. Remove it from this definition.
      txtDescExpr.Tag = 0
      txtDescExpr.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Check Start Date Calculation
  If (Len(txtCustomStart.Tag) > 0) And (Val(txtCustomStart.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateCalculation(CLng(txtCustomStart.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Calc hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly)
        If fRemove Then
          sBigMessage = "The report start date calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "report start date calculation"
        End If

      Case REC_SEL_VALID_DELETED
        ' Calc deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "report start date calculation"

        fRemove = (Not mblnReadOnly)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Calc hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "report start date calculation"
  
          fRemove = (Not mblnReadOnly)
        End If
      Case REC_SEL_VALID_INVALID
        ' Calc invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "report start date calculation"

        fRemove = (Not mblnReadOnly)

    End Select

    If fRemove Then
      ' Calc invalid, deleted or hidden by another user. Remove it from this definition.
      txtCustomStart.Tag = 0
      txtCustomStart.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Check End Date Calculation
  If (Len(txtCustomEnd.Tag) > 0) And (Val(txtCustomEnd.Tag) <> 0) Then
    fRemove = False
    iResult = ValidateCalculation(CLng(txtCustomEnd.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Calc hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly)
        If fRemove Then
          sBigMessage = "The report end date calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "report end date calculation"
        End If

      Case REC_SEL_VALID_DELETED
        ' Calc deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "report end date calculation"

        fRemove = (Not mblnReadOnly)

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Calc hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "report end date calculation"
  
          fRemove = (Not mblnReadOnly)
        End If
      Case REC_SEL_VALID_INVALID
        ' Calc invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "report end date calculation"

        fRemove = (Not mblnReadOnly)

    End Select

    If fRemove Then
      ' Calc invalid, deleted or hidden by another user. Remove it from this definition.
      txtCustomEnd.Tag = 0
      txtCustomEnd.Text = "<None>"
      mblnRecordSelectionInvalid = True
    End If
  End If

  ' Event Table Filters
  With grdEvents
    If .Rows > 0 Then
      For iLoop = .Rows - 1 To 0 Step -1
        varBookmark = .AddItemBookmark(iLoop)
        lngFilterID = .Columns("FilterID").CellValue(varBookmark)
        
        If lngFilterID > 0 Then
          fRemove = False
          iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
          sEventFilter = .Columns("Filter").CellValue(varBookmark)
          
          Select Case iResult
            Case REC_SEL_VALID_HIDDENBYUSER
              ' Calculation hidden by the current user.
              ' Only a problem if the current definition is NOT owned by the current user,
              ' or if the current definition is not already hidden.
              fRemove = (Not mblnDefinitionCreator) And _
                (Not mblnReadOnly)

              If fRemove Then
                sBigMessage = "The '" & sEventFilter & "' event table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
                MsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
              Else
                fNeedToForceHidden = True
  
                ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
                asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & sEventFilter & "' event table filter"
              End If

            Case REC_SEL_VALID_DELETED
              ' Calc deleted by another user.
              ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
              asDeletedParameters(UBound(asDeletedParameters)) = "'" & sEventFilter & "' event table filter"

              fRemove = (Not mblnReadOnly)

            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                ' Calc hidden by another user.
                ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
                asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & sEventFilter & "' event table filter"
  
                fRemove = (Not mblnReadOnly)
              End If
            Case REC_SEL_VALID_INVALID
              ' Calc invalid.
              ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
              asInvalidParameters(UBound(asInvalidParameters)) = "'" & sEventFilter & "' event table filter"

              fRemove = (Not mblnReadOnly)
          End Select

          If fRemove Then
            ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
            
            'JPD 20030731 Fault 6463
            mcolEvents.Item(Trim(.Columns("EventKey").Value)).FilterID = 0
            
            sRow = vbNullString
            sRow = sRow & .Columns("Name").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("TableID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Table").CellValue(varBookmark) & vbTab
            sRow = sRow & 0 & vbTab
            sRow = sRow & "" & vbTab
            sRow = sRow & .Columns("StartDateID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Start Date").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("StartSessionID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Start Session").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("EndDateID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("End Date").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("EndSessionID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("End Session").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("DurationID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Duration").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("LegendType").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Legend").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("LegendTableID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("LegendColumnID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("LegendCodeID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("LegendEventTypeID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Desc1ID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Description 1").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Desc2ID").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("Description 2").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("EventKey").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("ColourName").CellValue(varBookmark) & vbTab
            sRow = sRow & .Columns("ColourValue").CellValue(varBookmark) & vbTab
            
            If .Rows > 1 Then
              .RemoveItem iLoop
            Else
              .RemoveAll
            End If
            .AddItem sRow, iLoop

            SSTab1.Tab = 1
            .SetFocus

            mblnRecordSelectionInvalid = True
          End If
        End If
      Next iLoop
    End If
  End With

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



Private Sub optBaseFilter_Click()
  If Not mblnLoading Then
    Changed = True
  
    cmdBaseFilter.Enabled = True
  
    With txtBasePicklist
      .Text = ""
      .Tag = 0
    End With
  
    EnableDisableTabControls
    
    txtBaseFilter.Text = "<None>"
    cmdBasePicklist.Enabled = False
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optBasePicklist_Click()
  If Not mblnLoading Then
    Changed = True
  
    cmdBasePicklist.Enabled = True
  
    With txtBaseFilter
      .Text = ""
      .Tag = 0
    End With
  
    EnableDisableTabControls
  
    txtBasePicklist.Text = "<None>"
    cmdBaseFilter.Enabled = False
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optCurrentEnd_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optCurrentStart_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optFixedEnd_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optFixedStart_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optOffsetEnd_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub optOffsetStart_Click()
  If Not mblnLoading Then
    UpdateReportDetailsTab
    Changed = True
    
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub
Private Sub spnFreqEnd_Change()
  If Not mblnLoading Then
    If Me.spnFreqStart.Value > Me.spnFreqEnd.Value Then
      mblnLoading = True
      Me.spnFreqStart.Value = Me.spnFreqEnd.Value
      mblnLoading = False
    End If
    UpdateReportDetailsTab
    Changed = True
  End If
End Sub
Private Sub spnFreqStart_Change()
  If Not mblnLoading Then
    If Me.spnFreqStart.Value > Me.spnFreqEnd.Value Then
      mblnLoading = True
      Me.spnFreqEnd.Value = Me.spnFreqStart.Value
      mblnLoading = False
    End If
    UpdateReportDetailsTab
    Changed = True
  End If
End Sub
Private Sub txtBaseFilter_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub txtBasePicklist_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub txtDesc_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub txtEmailAttachAs_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub txtEmailSubject_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub


Private Sub txtFilename_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub


Private Sub txtName_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub



Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub






