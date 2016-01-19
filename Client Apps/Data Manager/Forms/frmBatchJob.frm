VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmBatchJob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Job Definition"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1008
   Icon            =   "frmBatchJob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5610
      Left            =   50
      TabIndex        =   80
      Top             =   120
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   9895
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Definition"
      TabPicture(0)   =   "frmBatchJob.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraScheduling"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraInfo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Jobs"
      TabPicture(1)   =   "frmBatchJob.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEMailNotify"
      Tab(1).Control(1)=   "fraJobs"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "O&utput"
      TabPicture(2)   =   "frmBatchJob.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOptions"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraDest"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraOutput"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame fraInfo 
         Height          =   2355
         Left            =   150
         TabIndex        =   0
         Top             =   450
         Width           =   9525
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5940
            MaxLength       =   30
            TabIndex        =   8
            Top             =   300
            Width           =   3405
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   2
            Top             =   300
            Width           =   3090
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1440
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1110
            Width           =   3090
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   3090
         End
         Begin SSDataWidgets_B.SSDBGrid grdAccess 
            Height          =   1485
            Left            =   5940
            TabIndex        =   10
            Top             =   720
            Width           =   3405
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   4
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
            stylesets(0).Picture=   "frmBatchJob.frx":0060
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
            stylesets(1).Picture=   "frmBatchJob.frx":007C
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
            Columns.Count   =   4
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
            Columns(3).Width=   3200
            Columns(3).Caption=   "ForcedAccess"
            Columns(3).Name =   "ForcedAccess"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
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
            Left            =   5085
            TabIndex        =   7
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   1155
            Width           =   1080
         End
         Begin VB.Label lblAccess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Left            =   5085
            TabIndex        =   9
            Top             =   765
            Width           =   825
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category :"
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   765
            Width           =   1005
         End
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Output Format :"
         Height          =   3000
         Left            =   -74850
         TabIndex        =   40
         Top             =   420
         Width           =   2265
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&HTML Document"
            Height          =   195
            Index           =   2
            Left            =   200
            TabIndex        =   41
            Top             =   480
            Width           =   1800
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "&Word Document"
            Height          =   195
            Index           =   3
            Left            =   200
            TabIndex        =   42
            Top             =   780
            Width           =   1900
         End
         Begin VB.OptionButton optOutputFormat 
            Caption         =   "E&xcel Worksheet"
            Height          =   195
            Index           =   4
            Left            =   200
            TabIndex        =   43
            Top             =   1080
            Value           =   -1  'True
            Width           =   1900
         End
      End
      Begin VB.Frame fraDest 
         Caption         =   "Output Destination(s) :"
         Height          =   3000
         Left            =   -72485
         TabIndex        =   44
         Top             =   420
         Width           =   7110
         Begin VB.CheckBox chkPreview 
            Caption         =   "Preview"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   3660
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   165
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Displa&y output on screen"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   135
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send as &email"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   56
            Top             =   1755
            Width           =   1515
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Save to &file"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   975
            Width           =   1455
         End
         Begin VB.CheckBox chkDestination 
            Caption         =   "Send to &printer"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   510
            Width           =   1620
         End
         Begin VB.CommandButton cmdEmailGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6550
            TabIndex        =   59
            Top             =   1710
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdFileName 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6550
            TabIndex        =   53
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtEmailGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1710
            Width           =   2900
         End
         Begin VB.TextBox txtEmailSubject 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3660
            TabIndex        =   61
            Top             =   2115
            Width           =   3240
         End
         Begin VB.ComboBox cboSaveExisting 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3660
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1320
            Width           =   3240
         End
         Begin VB.ComboBox cboPrinterName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3660
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   510
            Width           =   3240
         End
         Begin VB.TextBox txtFileName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   915
            Width           =   2900
         End
         Begin VB.TextBox txtEMailAttachAs 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3660
            TabIndex        =   63
            Tag             =   "0"
            Top             =   2520
            Width           =   3240
         End
         Begin VB.Label lblPrinter 
            AutoSize        =   -1  'True
            Caption         =   "Printer location :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2190
            TabIndex        =   48
            Top             =   510
            Width           =   1410
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            Caption         =   "If existing file :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2190
            TabIndex        =   54
            Top             =   1380
            Width           =   1350
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email group :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   2190
            TabIndex        =   57
            Top             =   1755
            Width           =   1200
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email subject :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   2190
            TabIndex        =   60
            Top             =   2160
            Width           =   1305
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2190
            TabIndex        =   51
            Top             =   975
            Width           =   1095
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Attach as :"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   2190
            TabIndex        =   62
            Top             =   2565
            Width           =   1065
         End
      End
      Begin VB.Frame fraEMailNotify 
         Caption         =   "Email Notifications :"
         Height          =   1400
         Left            =   -74850
         TabIndex        =   33
         Top             =   4080
         Width           =   9495
         Begin VB.TextBox txtEmailNotifyGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   4335
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   300
            Width           =   3900
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "&Send an email if the batch job fails"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   3390
         End
         Begin VB.CommandButton cmdEmailNotifyGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   8250
            TabIndex        =   36
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Send an email if the ba&tch job is successful"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   760
            Width           =   4035
         End
         Begin VB.TextBox txtEmailNotifyGroup 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   4335
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   700
            Width           =   3900
         End
         Begin VB.CommandButton cmdEmailNotifyGroup 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   8250
            TabIndex        =   39
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
      End
      Begin VB.Frame fraJobs 
         Height          =   3600
         Left            =   -74850
         TabIndex        =   25
         Top             =   420
         Width           =   9495
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "Re&move All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8130
            TabIndex        =   30
            Top             =   1860
            Width           =   1200
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   8130
            TabIndex        =   27
            Top             =   285
            Width           =   1200
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   8130
            TabIndex        =   28
            Top             =   795
            Width           =   1200
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8130
            TabIndex        =   29
            Top             =   1320
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveDown 
            Caption         =   "Mo&ve Down"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8130
            TabIndex        =   32
            Top             =   2925
            Width           =   1200
         End
         Begin VB.CommandButton cmdMoveUp 
            Caption         =   "Move &Up"
            Enabled         =   0   'False
            Height          =   400
            Left            =   8130
            TabIndex        =   31
            Top             =   2385
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdColumns 
            Height          =   3045
            Left            =   180
            TabIndex        =   26
            Top             =   285
            Width           =   7830
            ScrollBars      =   2
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   5
            stylesets.count =   5
            stylesets(0).Name=   "ssetHeaderDisabled"
            stylesets(0).ForeColor=   -2147483631
            stylesets(0).BackColor=   -2147483633
            stylesets(0).Picture=   "frmBatchJob.frx":0098
            stylesets(1).Name=   "ssetSelected"
            stylesets(1).ForeColor=   -2147483634
            stylesets(1).BackColor=   -2147483635
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
            stylesets(1).Picture=   "frmBatchJob.frx":00B4
            stylesets(2).Name=   "ssetEnabled"
            stylesets(2).ForeColor=   -2147483640
            stylesets(2).BackColor=   -2147483643
            stylesets(2).Picture=   "frmBatchJob.frx":00D0
            stylesets(3).Name=   "ssetHeaderEnabled"
            stylesets(3).ForeColor=   -2147483630
            stylesets(3).BackColor=   -2147483633
            stylesets(3).Picture=   "frmBatchJob.frx":00EC
            stylesets(4).Name=   "ssetDisabled"
            stylesets(4).ForeColor=   -2147483631
            stylesets(4).BackColor=   -2147483633
            stylesets(4).Picture=   "frmBatchJob.frx":0108
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
            RowNavigation   =   1
            MaxSelectedRows =   1
            StyleSet        =   "ssetDisabled"
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   5
            Columns(0).Width=   3281
            Columns(0).Caption=   "Job Type"
            Columns(0).Name =   "Job Type"
            Columns(0).CaptionAlignment=   2
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   250
            Columns(0).Locked=   -1  'True
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "Individual Job ID"
            Columns(1).Name =   "IndividualJobID"
            Columns(1).Alignment=   2
            Columns(1).CaptionAlignment=   2
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   3
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(2).Width=   5292
            Columns(2).Caption=   "Job Name"
            Columns(2).Name =   "Job Name"
            Columns(2).CaptionAlignment=   2
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(2).Locked=   -1  'True
            Columns(2).HasForeColor=   -1  'True
            Columns(3).Width=   4789
            Columns(3).Caption=   "Pause Parameter"
            Columns(3).Name =   "Parameter"
            Columns(3).CaptionAlignment=   2
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            Columns(4).Width=   3200
            Columns(4).Visible=   0   'False
            Columns(4).Caption=   "Hidden"
            Columns(4).Name =   "Hidden"
            Columns(4).Alignment=   2
            Columns(4).CaptionAlignment=   2
            Columns(4).DataField=   "Column 4"
            Columns(4).DataType=   11
            Columns(4).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13811
            _ExtentY        =   5371
            _StockProps     =   79
            Enabled         =   0   'False
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
      Begin VB.Frame fraScheduling 
         Caption         =   "Scheduling :"
         Height          =   2595
         Left            =   150
         TabIndex        =   11
         Top             =   2895
         Width           =   9525
         Begin GTMaskDate.GTMaskDate cboStartDate 
            Height          =   315
            Left            =   1815
            TabIndex        =   17
            Top             =   1050
            Width           =   1305
            _Version        =   65537
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   77
            Enabled         =   0   'False
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
            AutoSelect      =   -1  'True
            BackColor       =   -2147483633
            MaskCentury     =   2
            SpinButtonEnabled=   0   'False
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
         Begin VB.ComboBox cboRoleToPrompt 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmBatchJob.frx":0124
            Left            =   1815
            List            =   "frmBatchJob.frx":0126
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1845
            Width           =   2115
         End
         Begin VB.ComboBox cboPeriod 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmBatchJob.frx":0128
            Left            =   2595
            List            =   "frmBatchJob.frx":0138
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   640
            Width           =   1335
         End
         Begin VB.CheckBox chkRunOnce 
            Caption         =   "S&kip Missed Days"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5100
            TabIndex        =   24
            Top             =   1245
            Width           =   2160
         End
         Begin VB.CheckBox chkScheduled 
            Caption         =   "&Schedule"
            Height          =   315
            Left            =   200
            TabIndex        =   12
            Top             =   300
            Width           =   1100
         End
         Begin VB.CheckBox chkWeekEnds 
            Caption         =   "I&nclude Weekends"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5100
            TabIndex        =   23
            Top             =   930
            Width           =   2160
         End
         Begin VB.CheckBox chkIndefinitely 
            Caption         =   "Run Indefinitel&y"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5100
            TabIndex        =   22
            Top             =   630
            Width           =   2160
         End
         Begin COASpinner.COA_Spinner spnFrequency 
            Height          =   315
            Left            =   1815
            TabIndex        =   14
            Top             =   645
            Width           =   630
            _ExtentX        =   1111
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
            Text            =   "1"
         End
         Begin GTMaskDate.GTMaskDate cboEndDate 
            Height          =   315
            Left            =   1815
            TabIndex        =   19
            Top             =   1440
            Width           =   1305
            _Version        =   65537
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   77
            Enabled         =   0   'False
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
            AutoSelect      =   -1  'True
            BackColor       =   -2147483633
            MaskCentury     =   2
            SpinButtonEnabled=   0   'False
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
         Begin VB.Label lblRole 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Group :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   450
            TabIndex        =   20
            Top             =   1905
            Width           =   1110
         End
         Begin VB.Label lblRunEvery 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Period :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   450
            TabIndex        =   13
            Top             =   705
            Width           =   675
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   450
            TabIndex        =   18
            Top             =   1500
            Width           =   915
         End
         Begin VB.Label lblStart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   450
            TabIndex        =   16
            Top             =   1095
            Width           =   1020
         End
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Report Options :"
         Height          =   2020
         Left            =   -74850
         TabIndex        =   64
         Top             =   3420
         Width           =   9470
         Begin VB.CheckBox chkRetainPivot 
            Caption         =   "Retain pi&vot/chart"
            Height          =   255
            Left            =   2760
            TabIndex        =   77
            Top             =   1590
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CommandButton cmdTitlePageClear 
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
            Left            =   8820
            MaskColor       =   &H000000FF&
            TabIndex        =   68
            ToolTipText     =   "Clear Path"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdTitlePageTemplate 
            Caption         =   "..."
            Height          =   315
            Left            =   8490
            TabIndex        =   67
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdfilterClear 
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
            Left            =   8820
            MaskColor       =   &H000000FF&
            TabIndex        =   74
            ToolTipText     =   "Clear Path"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CheckBox chkForceCoverSheet 
            Caption         =   "Create c&over sheet(s)"
            Height          =   255
            Left            =   5400
            TabIndex        =   76
            Top             =   1320
            Width           =   3855
         End
         Begin VB.CheckBox chkTOC 
            Caption         =   "C&reate table of contents"
            Height          =   255
            Left            =   2760
            TabIndex        =   75
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtTitlePage 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2750
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   240
            Width           =   5750
         End
         Begin VB.TextBox txtReportPackTitle 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2750
            TabIndex        =   70
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   600
            Width           =   5750
         End
         Begin VB.CommandButton cmdOverrideFilter 
            Caption         =   "..."
            Height          =   315
            Left            =   8490
            TabIndex        =   73
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtOverrideFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2750
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Tag             =   "0"
            Text            =   "<None>"
            Top             =   960
            Width           =   5750
         End
         Begin VB.TextBox txtFilterSource 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2750
            TabIndex        =   81
            TabStop         =   0   'False
            Tag             =   "1"
            Text            =   "Personnel_Records"
            Top             =   960
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lblTitlePage 
            AutoSize        =   -1  'True
            Caption         =   "Report Pack Template :"
            Height          =   195
            Left            =   195
            TabIndex        =   65
            Top             =   300
            Width           =   2025
         End
         Begin VB.Label lblReportPackTitle 
            AutoSize        =   -1  'True
            Caption         =   "Report Pack Title :"
            Height          =   195
            Left            =   195
            TabIndex        =   69
            Top             =   660
            Width           =   1590
         End
         Begin VB.Label lblOverrideFilter 
            AutoSize        =   -1  'True
            Caption         =   "Personnel Override Filter :"
            Height          =   195
            Left            =   195
            TabIndex        =   71
            Top             =   1020
            Width           =   2265
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   400
      Left            =   7320
      TabIndex        =   78
      Top             =   5800
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8645
      TabIndex        =   79
      Top             =   5800
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBatchJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsData As DataMgr.clsDataAccess               'Data Access Class
Private mlngBatchJobID As Long                          'ID of current BatchJob
Private mblnFromCopy As Boolean                         'Is this definition a copy ?
Private mblnCancelled As Boolean                        'Has operation been cancelled ?
Private mblnReadOnly As Boolean                         'Is the definition read only ?
Private mlngTimeStamp As Long                           'Timestamp for def locking
Private mblnDefinitionCreator As Boolean
Private mblnGridFlag As Boolean
Private mblnForceChanged As Boolean
Private mblnForceHidden As Boolean
Private mblnLoading As Boolean
Private mblnAlreadyActivated As Boolean
Private mblnIsBatch As Boolean                          'Flag for whether this Batch Job or Report Pack
'Output Options
Private mobjOutputDef As clsOutputDef

Public IsReportPack As Boolean
Private mblnIsChildColumnSelected As Boolean

Private mlngWordFormat As Long
Private mlngExcelFormat As Long


Private Function BatchJobHiddenGroups() As String
  Dim sBatchJobHiddenGroups As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  sBatchJobHiddenGroups = vbTab
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        If .Columns("Access").CellText(varBookmark) = AccessDescription(ACCESS_HIDDEN) Then
          sBatchJobHiddenGroups = sBatchJobHiddenGroups & .Columns("GroupName").CellText(varBookmark) & vbTab
        End If
      End If
    Next iLoop

    .MoveFirst
  End With

  BatchJobHiddenGroups = sBatchJobHiddenGroups
  
End Function

Private Sub RefreshColumnsGrid()
  
  With grdColumns
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

  SetButtonState
  
End Sub

Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOk.Enabled = pblnChanged
End Property

Private Function JobUtilityType(psJobType As String) As UtilityType
  ' Return the utility type code fot the given job type code.
    
  Select Case UCase(psJobType)
    Case "ABSENCE BREAKDOWN"
      JobUtilityType = utlAbsenceBreakdown
    Case "BRADFORD FACTOR"
      JobUtilityType = utlBradfordFactor
    Case "CALENDAR REPORT"
      JobUtilityType = utlCalendarReport
    Case "CAREER PROGRESSION"
      JobUtilityType = utlCareer
    Case "CROSS TAB"
      JobUtilityType = utlCrossTab
    Case "CUSTOM REPORT"
      JobUtilityType = utlCustomReport
    Case "DATA TRANSFER"
      JobUtilityType = utlDataTransfer
    Case "EXPORT"
      JobUtilityType = utlExport
    Case "GLOBAL ADD"
      JobUtilityType = UtlGlobalAdd
    Case "GLOBAL DELETE"
      JobUtilityType = utlGlobalDelete
    Case "GLOBAL UPDATE"
      JobUtilityType = utlGlobalUpdate
    Case "IMPORT"
      JobUtilityType = utlImport
    Case "ENVELOPES & LABELS"
      JobUtilityType = utlLabel
    Case "MAIL MERGE"
      JobUtilityType = utlMailMerge
    Case "MATCH REPORT"
      JobUtilityType = utlMatchReport
    Case "RECORD PROFILE"
      JobUtilityType = utlRecordProfile
    Case "SUCESSION PLANNING"
      JobUtilityType = utlSuccession
  End Select

End Function

Private Sub RefreshRoleToPromptCombo(Optional psGivenRole As String)
  Dim iLoop As Integer
  Dim sCurrentRoleToPrompt As String
  Dim varBookmark As Variant
  Dim iGivenRoleIndex As Integer
  Dim iDfltRoleIndex As Integer
  Dim iCurrRoleIndex As Integer
  Dim sGroupName As String
  Dim fEnabled As Boolean
  
  iGivenRoleIndex = -1
  iDfltRoleIndex = -1
  iCurrRoleIndex = -1
  
  If cboRoleToPrompt.ListIndex >= 0 Then
    sCurrentRoleToPrompt = cboRoleToPrompt.List(cboRoleToPrompt.ListIndex)
  End If
  
  ' Use the update method to ensure that the CellText/CellValue methods
  ' return the correct values (they're not updated with the Text/Value values
  ' until the cell loses focus, or the Update method is called).
  grdAccess.Update
  
  With cboRoleToPrompt
    ' Remove all existing items.
    .Clear
  
    If chkScheduled.Value = vbChecked Then
      ' Loop through the items in the access grid, adding to
      ' the combo any roles that can see the Batch Job.
      For iLoop = 1 To (grdAccess.Rows - 1)
        varBookmark = grdAccess.AddItemBookmark(iLoop)
          
        If (grdAccess.Columns("SysSecMgr").CellText(varBookmark) = "1") Or _
          (grdAccess.Columns("Access").CellText(varBookmark) <> AccessDescription(ACCESS_HIDDEN)) Or _
          (UCase(grdAccess.Columns("GroupName").CellText(varBookmark)) = UCase(gsUserGroup)) Or _
          Not mblnDefinitionCreator Then
      
          sGroupName = grdAccess.Columns("GroupName").CellText(varBookmark)
          .AddItem sGroupName
          
          If UCase(sGroupName) = UCase(psGivenRole) Then
            iGivenRoleIndex = .ListCount - 1
          End If
          
          If UCase(sGroupName) = UCase(gsUserGroup) Then
            iDfltRoleIndex = .ListCount - 1
          End If
          
          If UCase(sGroupName) = UCase(sCurrentRoleToPrompt) Then
            iCurrRoleIndex = .ListCount - 1
          End If
        End If
      Next iLoop
      
      If .ListCount = 1 Then
        .ListIndex = 0
      ElseIf .ListCount > 1 Then
        If iGivenRoleIndex >= 0 Then
          .ListIndex = iGivenRoleIndex
        Else
          If iCurrRoleIndex >= 0 Then
            .ListIndex = iCurrRoleIndex
          Else
            If iDfltRoleIndex >= 0 Then
              .ListIndex = iDfltRoleIndex
            Else
              .ListIndex = 0
            End If
          End If
        End If
      End If
    End If
    
    If .ListCount <= 1 Then
      fEnabled = False
    Else
      fEnabled = Not mblnReadOnly
    End If

    lblRole.Enabled = fEnabled
    .Enabled = fEnabled
    If fEnabled Then
      .BackColor = vbWindowBackground
    Else
      .BackColor = vbButtonFace
    End If
  End With
  
End Sub

Public Property Get SelectedID() As Long
  SelectedID = mlngBatchJobID
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
Public Property Let FromCopy(ByVal bCopy As Boolean)
  mblnFromCopy = bCopy
End Property

Private Sub cboEndDate_Change()
  Changed = True
End Sub

Private Sub cboEndDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboEndDate.DateValue = Date
  End If

End Sub

Private Sub cboEndDate_LostFocus()

'  If IsNull(cboEndDate.DateValue) And Not _
'     IsDate(cboEndDate.DateValue) And _
'     cboEndDate.Text <> "  /  /" Then
'
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboEndDate.DateValue = Null
'     cboEndDate.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboEndDate

End Sub

Private Sub cboPeriod_Click()
  Changed = True
End Sub

Private Sub cboPrinterName_Click()
  Me.Changed = True
End Sub

Private Sub cboRoleToPrompt_Click()
  Changed = True
End Sub

Private Sub cboSaveExisting_Click()
  Me.Changed = True
End Sub

Private Sub cboStartDate_Change()
  Changed = True
End Sub

Private Sub cboStartDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboStartDate.DateValue = Date
  End If

End Sub

Private Sub cboStartDate_LostFocus()

'  If IsNull(cboStartDate.DateValue) And Not _
'     IsDate(cboStartDate.DateValue) And _
'     cboStartDate.Text <> "  /  /" Then
'
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboStartDate.DateValue = Null
'     cboStartDate.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboStartDate

End Sub

Private Sub chkDestination_Click(Index As Integer)
  mobjOutputDef.DestinationClick Index
  Changed = True
End Sub

Private Sub chkEmail_Click(Index As Integer)
  cmdEmailNotifyGroup(Index).Enabled = (chkEmail(Index).Value = vbChecked)
  txtEmailNotifyGroup(Index).Text = IIf(chkEmail(Index).Value = vbChecked, "<None>", vbNullString)
  txtEmailNotifyGroup(Index).Tag = 0
  Changed = True
End Sub

Private Sub chkForceCoverSheet_Click()
  Changed = True
End Sub
Private Sub chkRetainPivot_click()
  Changed = True
End Sub

Private Sub chkRunOnce_Click()
  Changed = True
End Sub

Private Sub chkTOC_Click()
  Changed = True
End Sub

Private Sub chkWeekEnds_Click()
  Changed = True
End Sub

Private Sub cmdEmailNotifyGroup_Click(Index As Integer)

  Dim frmDefinition As frmEmailDefGroup
  Dim frmSelection As frmDefSel
  Dim lForms As Long
  Dim blnExit As Boolean
  Dim blnOK As Boolean

  Set frmSelection = New frmDefSel
  blnExit = False

  Set frmDefinition = New frmEmailDefGroup

  With frmSelection

    .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
    .EnableRun = False
    .TableComboVisible = False
    .SelectedID = Val(txtEmailNotifyGroup(Index).Tag)

    Do While Not blnExit
      
      If .ShowList(utlEmailGroup) Then

        .Show vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise True, .FromCopy
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing

        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
        Case edtEdit
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtPrint
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.PrintDef .SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtSelect
          txtEmailNotifyGroup(Index).Text = .SelectedText
          txtEmailNotifyGroup(Index).Tag = .SelectedID
          Changed = True
          blnExit = True

        Case edtDeselect
          txtEmailNotifyGroup(Index).Text = "<None>"
          txtEmailNotifyGroup(Index).Tag = 0
          Changed = True
          blnExit = True

        Case 0
          blnExit = True  'cancel

        End Select

      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Private Sub cmdFilterClear_Click()
  txtOverrideFilter = ""
  txtOverrideFilter.Tag = 0
  ForceDefinitionToBeHiddenIfNeeded2
  cmdfilterClear.Enabled = False
  cmdOverrideFilter.SetFocus
  Changed = Not mblnLoading
End Sub

Private Sub cmdOverrideFilter_Click()
  GetFilter txtFilterSource, txtOverrideFilter
  cmdfilterClear.Enabled = txtOverrideFilter.Text <> ""
  Changed = Not mblnLoading
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

  If gblnReportPackMode Then
    ForceDefinitionToBeHiddenIfNeeded2
  Else
    ForceDefinitionToBeHiddenIfNeeded
  End If
End Sub

Private Sub cmdTitlePageClear_Click()
  txtTitlePage.Text = vbNullString
  txtReportPackTitle = vbNullString
  cmdTitlePageClear.Enabled = False
  cmdTitlePageTemplate.SetFocus
  
  Changed = Not mblnLoading
End Sub

Private Sub Form_Activate()
  'JPD 20031120 Fault 7512
  'If Not mblnDontDoActivateCheck Then
  If Not mblnAlreadyActivated Then
    ForceDefinitionToBeHiddenIfNeeded True
    Changed = (mblnFromCopy Or mblnForceChanged)  'MH20040422 Fault 8495
    mblnAlreadyActivated = True
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    ' NHRD we will need code here for the new helpcontextid whe known what the next one is.
    
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()

  SSTab1.Tab = 0
  grdAccess.RowHeight = 239
  grdColumns.RowHeight = 239
  
  If IsReportPack Then
    Me.HelpContextID = 5036
    'This will in effect remove the Pause Parameter column
    grdColumns.Columns(0).Width = (grdColumns.Width * 0.33) 'Job Type
    grdColumns.Columns(2).Width = (grdColumns.Width * 0.67) 'Job Name
    
    grdColumns.Columns(0).Caption = "Report Type"
    grdColumns.Columns(2).Caption = "Report Name"
    
    Set mobjOutputDef = New clsOutputDef
    mobjOutputDef.ParentForm = Me
    mobjOutputDef.PopulateCombos True, True, True
    
    SSTab1.TabCaption(1) = "Report &Items"
  End If

  SSTab1.TabVisible(2) = IsReportPack

  'JPD 20041117 Fault 8231
  UI.FormatGTDateControl cboEndDate
  UI.FormatGTDateControl cboStartDate
  
  txtFilterSource = gsPersonnelTableName 'Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
  txtFilterSource.Tag = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))

End Sub

Public Function Initialise(pblnNew As Boolean, pblnCopy As Boolean, Optional plngBatchJobID As Long) As Boolean
  
  Set mclsData = New DataMgr.clsDataAccess           'Instantiate class
  Dim iUtilityType As UtilityType
  Dim lngFormat As Long
  
  Screen.MousePointer = vbHourglass
  
  iUtilityType = IIf(IsReportPack, utlReportPack, utlBatchJob)
  
  mblnLoading = True
  
  If IsReportPack Then
    Me.Caption = "Report Pack Definition"
    chkEmail(0).Caption = "&Send email if the report pack fails"
    chkEmail(1).Caption = "Send email if the repor&t pack is successful"
    mobjOutputDef.FormatClick fmtExcelWorksheet
    optOutputFormat(fmtExcelWorksheet).Value = True
    chkRetainPivot = vbChecked: chkRetainPivot.Visible = False
  End If
  
  If pblnNew Then                                  'If this is a new definition
    mlngBatchJobID = 0                             'Set ID to 0 to indicate new record
    PopulateAccessGrid
    
    GetObjectCategories cboCategory, iUtilityType, 0, 0
    SetComboItem cboCategory, IIf(glngCurrentCategoryID = -1, 0, glngCurrentCategoryID)
    
    ClearForNew                                    'Clear fields ready for new entry
    Changed = False
    mblnDefinitionCreator = True
  Else                                             'Otherwise, editing an existing one
    mlngBatchJobID = plngBatchJobID                'Equate variables
    mblnFromCopy = pblnCopy
    
    PopulateAccessGrid
    
    If Not RetrieveBatchJobDetails Then
      If COAMsgBox("OpenHR could not load all of the definition successfully." & vbCrLf & vbCrLf & "The recommendation is that " & _
             "you delete the definition and create a new one," & vbCrLf & "however, you may edit the existing " & _
             "definition if you wish." & vbCrLf & vbCrLf & "Would you like to continue and edit this definition ?", vbQuestion + vbYesNo, IIf(gblnReportPackMode, "Report Pack", "Batch Job")) = vbNo Then
        Initialise = False
        Exit Function
      End If
    End If
        
    'Reset pointer so copy will be saved as a new one
    If mblnFromCopy Then
      mlngBatchJobID = 0
      Changed = True
    Else
      If Not mblnForceChanged Then Changed = False Else Changed = True
    End If
  End If

  Cancelled = False
  Screen.MousePointer = vbDefault
  Initialise = True
  mblnLoading = False
  Exit Function
  
Initialise_ERROR:
  
  COAMsgBox "Error whilst loading " & IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " Definition." & vbCrLf & "(" & Err.Description & ")"
  Initialise = False
  
End Function

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
    SetAllAccess AccessCode(grdAccess.Columns("Access").Text), False
  
    grdAccess.MoveFirst
    grdAccess.Col = 1
  End If

  RefreshRoleToPromptCombo

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
    
    If (Not mblnDefinitionCreator) Or _
      mblnReadOnly Or _
      (.Columns("ForcedAccess").CellText(varBkmk) = "1") Or _
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
    If (Not mblnDefinitionCreator) Or mblnReadOnly Or (.Columns("ForcedAccess").CellText(Bookmark) = "1") Then
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


Private Sub PopulateAccessGrid(Optional psAccessCode As String)
  ' Populate the access grid.
  Dim rsAccess As ADODB.Recordset
  
  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With
  
  ' Get the recordset of user groups and their access on this definition.
  Set rsAccess = GetUtilityAccessRecords(IIf(IsReportPack, utlReportPack, utlBatchJob), mlngBatchJobID, mblnFromCopy)
  If Not rsAccess Is Nothing Then
    ' Add the user groups and their access on this definition to the access grid.
    With rsAccess
      Do While Not .EOF
        grdAccess.AddItem !Name & _
          vbTab & AccessDescription(IIf(Len(psAccessCode) > 0, psAccessCode, !Access)) & _
          vbTab & !sysSecMgr & _
          vbTab & "0"
        
        .MoveNext
      Loop
    
      .Close
    End With
  End If
  Set rsAccess = Nothing

End Sub
Private Sub SetAllAccess(psAccess As String, pfForced As Boolean)
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
        If (.Columns("SysSecMgr").CellText(varBookmark) <> "1") And _
          (.Columns("ForcedAccess").CellText(varBookmark) <> "1") Then
          .Columns("Access").Text = AccessDescription(psAccess)
          
          If pfForced Then
            .Columns("ForcedAccess").Text = "1"
          End If
        End If
      End If
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
  
End Sub


Private Sub SaveAccess()
  Dim sSQL As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  
  ' Clear the access records first.
  sSQL = "DELETE FROM ASRSysBatchJobAccess WHERE ID = " & mlngBatchJobID
  mclsData.ExecuteSql sSQL
  
  ' Enter the new access records with dummy access values.
  sSQL = "INSERT INTO ASRSysBatchJobAccess" & _
    " (ID, groupName, access)" & _
    " (SELECT " & mlngBatchJobID & ", sysusers.name," & _
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
  mclsData.ExecuteSql (sSQL)

  ' Update the new access records with the real access values.
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      sSQL = "IF EXISTS (SELECT * FROM ASRSysBatchJobAccess" & _
        " WHERE ID = " & CStr(mlngBatchJobID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'" & _
        "  AND access <> '" & ACCESS_READWRITE & "')" & _
        " UPDATE ASRSysBatchJobAccess" & _
        "  SET access = '" & AccessCode(.Columns("Access").Text) & "'" & _
        "  WHERE ID = " & CStr(mlngBatchJobID) & _
        "  AND groupName = '" & .Columns("GroupName").Text & "'"
      mclsData.ExecuteSql (sSQL)
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow

End Sub




Private Sub grdColumns_GotFocus()

  mblnGridFlag = True

End Sub

Private Sub ClearForNew()
  
  'Clear all controls data for a new batch definition
  txtUserName = gsUserName
  txtName.Text = ""
  txtDesc.Text = ""
  spnFrequency.Value = 1
  SetComboText cboPeriod, "Day(s)"
  cboStartDate.Text = ""
  cboEndDate.Text = ""
  chkWeekEnds.Value = 0
  chkIndefinitely.Value = 0
  chkScheduled = False
  chkRunOnce.Value = 0
  grdColumns.RemoveAll
  
  If Not mblnLoading Then
    ForceDefinitionToBeHiddenIfNeeded False, gsUserGroup
  End If
  
End Sub

Private Function RetrieveBatchJobDetails() As Boolean
  Dim pstrAddString As String
  Dim prstTemp As Recordset
  Dim pblnDefinitionCreator As Boolean
  Dim sRoleToPrompt As String
  Dim fJobOK As Boolean
  Dim sMessage As String
  Dim iUtilityType As UtilityType
  
  On Error GoTo Load_ERROR
  
  iUtilityType = IIf(IsReportPack, utlReportPack, utlBatchJob)
  
  Set prstTemp = datGeneral.GetRecords( _
      "SELECT ASRSysBatchJobName.*, " & _
      "isnull(fail.name,'') as 'EmailFailedName', isnull(success.name,'') as 'EmailSuccessName', " & _
      "CONVERT(integer, ASRSysBatchJobName.TimeStamp) AS intTimeStamp " & _
      "FROM ASRSysBatchJobName " & _
      "LEFT OUTER JOIN ASRSysEmailGroupName fail ON fail.EmailGroupID = EmailFailed " & _
      "LEFT OUTER JOIN ASRSysEmailGroupName success ON success.EmailGroupID = EmailSuccess " & _
      "WHERE ID = " & mlngBatchJobID)
      
  If prstTemp.BOF And prstTemp.EOF Then
    COAMsgBox "Cannot load the definition for this " & IIf(gblnReportPackMode, "report pack", "batch job") & "." & vbCrLf & "(" & Err.Description & ")", vbCritical + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
    Set prstTemp = Nothing
    RetrieveBatchJobDetails = False
    Exit Function
  End If
  
  'Set definition description
  txtDesc.Text = IIf(IsNull(prstTemp!Description), "", prstTemp!Description)
  
  'Set definition name
  If FromCopy Then
    txtName.Text = "Copy of " & prstTemp!Name
    txtUserName.Text = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = prstTemp!Name
    txtUserName.Text = StrConv(prstTemp!userName, vbProperCase)
    mblnDefinitionCreator = (LCase(prstTemp!userName) = LCase(gsUserName))
  End If
  
  GetObjectCategories cboCategory, iUtilityType, mlngBatchJobID
  
  If IsReportPack Then
    mblnReadOnly = Not datGeneral.SystemPermission("REPORTPACKS", "EDIT")
  Else
    mblnReadOnly = Not datGeneral.SystemPermission("BATCHJOBS", "EDIT")
  End If
  
  If (Not mblnReadOnly) And (Not mblnDefinitionCreator) Then
    mblnReadOnly = (CurrentUserAccess(iUtilityType, mlngBatchJobID) = ACCESS_READONLY)
  End If
  
  chkScheduled.Value = IIf(prstTemp!scheduled = True, vbChecked, vbUnchecked)
  spnFrequency.Value = prstTemp!Frequency
  
  Select Case prstTemp!Period
    Case "W": SetComboText cboPeriod, "Week(s)"
    Case "M": SetComboText cboPeriod, "Month(s)"
    Case "Y": SetComboText cboPeriod, "Year(s)"
    Case Else: SetComboText cboPeriod, "Day(s)"
  End Select
  
  cboStartDate.Text = IIf(IsDate(prstTemp!StartDate) And Not IsNull(prstTemp!StartDate), Format(prstTemp!StartDate, DateFormat), "")
  cboEndDate.Text = IIf(IsDate(prstTemp!EndDate) And Not IsNull(prstTemp!EndDate), Format(prstTemp!EndDate, DateFormat), "")
  chkIndefinitely.Value = IIf(prstTemp!Indefinitely = True, vbChecked, vbUnchecked)
  chkWeekEnds.Value = IIf(prstTemp!Weekends = True, vbChecked, vbUnchecked)
  chkRunOnce.Value = IIf(prstTemp!RunOnce = True, vbChecked, vbUnchecked)
  sRoleToPrompt = IIf(IsNull(prstTemp!RoleToPrompt), "", prstTemp!RoleToPrompt)
  
  chkEmail(0).Value = IIf(Val(prstTemp!EmailFailed) > 0, vbChecked, vbUnchecked)
  txtEmailNotifyGroup(0).Tag = Val(prstTemp!EmailFailed)
  txtEmailNotifyGroup(0).Text = prstTemp!EmailFailedName
  
  chkEmail(1).Value = IIf(Val(prstTemp!EmailSuccess) > 0, vbChecked, vbUnchecked)
  txtEmailNotifyGroup(1).Tag = Val(prstTemp!EmailSuccess)
  txtEmailNotifyGroup(1).Text = prstTemp!EmailSuccessName

  If (chkScheduled.Value = vbChecked) And (sRoleToPrompt = "") Then
    COAMsgBox "The user group selected in this " & IIf(gblnReportPackMode, "report pack", "batch job") & " has been deleted.", vbInformation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
    mblnForceChanged = True
  End If

  If IsReportPack Then
    txtTitlePage.Text = prstTemp!OutputTitlePage
    txtReportPackTitle.Text = prstTemp!OutputReportPackTitle
    txtOverrideFilter.Text = prstTemp!OutputOverrideFilter
    txtOverrideFilter.Tag = prstTemp!OverrideFilterID
    
    cmdTitlePageClear.Enabled = Not txtTitlePage.Text = ""
    cmdfilterClear.Enabled = Not txtOverrideFilter.Text = ""
    
    optOutputFormat(prstTemp!OutputFormat).Value = True
    mobjOutputDef.PopulateOutputControls prstTemp
    
    chkForceCoverSheet.Value = IIf(prstTemp!OutputCoverSheet, vbChecked, vbUnchecked)
    chkTOC.Value = IIf(prstTemp!OutputTOC, vbChecked, vbUnchecked)
    chkRetainPivot.Value = IIf(prstTemp!OutputRetainPivotOrChart, vbChecked, vbUnchecked)
    'ChkRetainCharts.Value = IIf(prstTemp!OutputRetainCharts, vbChecked, vbUnchecked)
    chkForceCoverSheet.Value = IIf(prstTemp!OutputCoverSheet, vbChecked, vbUnchecked)

  End If
    
  If mblnReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  grdAccess.Enabled = True
  grdColumns.Enabled = True
  
  mlngTimeStamp = prstTemp!intTimestamp
  
  ' Now get the details...
  Set prstTemp = datGeneral.GetRecords("SELECT * " & _
                                      "FROM ASRSysBatchJobDetails WHERE BatchJobNameID = " & mlngBatchJobID & _
                                      "ORDER BY JobOrder")
  
  If prstTemp.BOF And prstTemp.EOF Then
    COAMsgBox "Cannot load the individual jobs for this " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition." & vbCrLf & "(" & Err.Description & ")", vbCritical + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
    Set prstTemp = Nothing
    RetrieveBatchJobDetails = False
    Exit Function
  End If
  
  sMessage = ""
  With grdColumns
    .RemoveAll
    
    Do While Not prstTemp.EOF
      fJobOK = True
      
      If (prstTemp!JobType = "Absence Breakdown") Or _
        (prstTemp!JobType = "Bradford Factor") Then
        fJobOK = gfAbsenceEnabled
      End If
      
      If (prstTemp!JobType = "Stability Index Report") Or _
        (prstTemp!JobType = "Turnover Report") Then
        fJobOK = gfPersonnelEnabled
      End If
      
      If fJobOK Then
        pstrAddString = prstTemp!JobType & vbTab
        pstrAddString = pstrAddString & prstTemp!JobID & vbTab
        
        'If prstTemp!jobtype <> "-- Pause --" Then
        If prstTemp!JobID > 0 Then
          pstrAddString = pstrAddString & GetJobName(prstTemp!JobType, prstTemp!JobID) & vbTab
          pstrAddString = pstrAddString & "N/A"
        Else
          pstrAddString = pstrAddString & vbTab
          
          If prstTemp!Parameter <> "" Then
            pstrAddString = pstrAddString & prstTemp!Parameter
          End If
        End If
        
        .AddItem pstrAddString
      Else
        sMessage = sMessage & vbCrLf & _
          vbTab & prstTemp!JobType
      End If
      
      prstTemp.MoveNext
    Loop
  End With
  
  If Not ForceDefinitionToBeHiddenIfNeeded(True, sRoleToPrompt) Then
    Set prstTemp = Nothing
    RetrieveBatchJobDetails = True
    Exit Function
  End If
  
  If Len(sMessage) > 0 Then
    If (Not mblnReadOnly) Then
      COAMsgBox "The following jobs have been removed from the " & IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " as the system is no longer configured to run them :" & vbCrLf & _
        sMessage, vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
      mblnForceChanged = True
    Else
      COAMsgBox "The system is no longer configured to run the following jobs, but you do not have permission to remove them from this " & IIf(gblnReportPackMode, "report pack", "batch job") & " :" & vbCrLf & _
        sMessage, vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
      RetrieveBatchJobDetails = True
      Exit Function
    End If
  End If
  
  Set prstTemp = Nothing
  
  RetrieveBatchJobDetails = True
  Exit Function
  
Load_ERROR:

  RetrieveBatchJobDetails = False
  
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim pintAnswer As Integer
  
  If Changed = True And Not mblnReadOnly Then
    
    pintAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, IIf(gblnReportPackMode, "Report Pack", "Batch Job"))
      
    If pintAnswer = vbYes Then
      Cancel = True
      If ValidateDefinition Then
        If IsReportPack Then
          If SaveDefinition2 Then
            Cancel = False
          End If
        Else
          If SaveDefinition Then
            Cancel = False
          End If
        End If
      End If
      Exit Sub
    ElseIf pintAnswer = vbCancel Then
      Cancel = True
      Me.Changed = True
      Set frmBatchJob = Nothing
      Exit Sub
    ElseIf pintAnswer = vbNo Then
      'Cancel = True
      'Changed = False
      Exit Sub
    End If
  
  End If
  
End Sub

Private Sub grdColumns_DblClick()

  If grdColumns.Rows > 0 And cmdEdit.Enabled Then
    cmdEdit_Click
  End If

End Sub

Private Sub cmdClearAll_Click()

  If COAMsgBox("Are you sure you wish to clear all jobs from this " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition ?", vbQuestion + vbYesNo, IIf(gblnReportPackMode, "Report Pack", "Batch Job")) = vbYes Then
    grdColumns.RemoveAll
    
    ForceDefinitionToBeHiddenIfNeeded
    Changed = True
  End If
  
End Sub

Private Sub chkScheduled_Click()
  
  If chkScheduled.Value = 1 Then                 'If Batch is to be scheduled
    SchedControls True               'Enable relevant ctls
    cboStartDate.DateValue = Now
  Else
    SchedControls False              'Else disable them
    cboStartDate.DateValue = Null
  End If

  Changed = True
  
End Sub

Private Sub SchedControls(Value As Boolean)

  'NHRD16092004 Fault 6823
  If Not Value Then
    spnFrequency.Text = ""
    'cboPeriod.re .Text = vbNullString  'this is a read only cbo
    'NHRD27092005 Fault 9353
    cboPeriod.ListIndex = 0
    cboStartDate.Text = vbNullString
    cboEndDate.Text = vbNullString
    chkIndefinitely.Value = vbUnchecked
    chkWeekEnds.Value = vbUnchecked
    chkRunOnce.Value = vbUnchecked
  End If
  
  ' Set enabled/disabled state of scheduling controls
  spnFrequency.Enabled = Value
  cboPeriod.Enabled = Value
  cboStartDate.Enabled = Value
  cboEndDate.Enabled = IIf(chkIndefinitely.Value = 1, False, Value)
  chkIndefinitely.Enabled = Value
  chkWeekEnds.Enabled = Value
  chkRunOnce.Enabled = Value
  
  'MH20010704
  'Skipped missed days should default to ticked
  chkRunOnce.Value = IIf(Value, vbChecked, vbUnchecked)
  
  lblStart.Enabled = Value
  lblEnd.Enabled = Value
  lblRunEvery.Enabled = Value
  
  cboRoleToPrompt.Enabled = Value
  lblRole.Enabled = IIf(Value, vbChecked, vbUnchecked)
  
  If cboEndDate.Enabled = False Then cboEndDate.BackColor = vbButtonFace Else cboEndDate.BackColor = vbWindowBackground
  
  If spnFrequency.Enabled = True Then
    spnFrequency.BackColor = vbWindowBackground
  Else
    spnFrequency.BackColor = vbButtonFace
  End If

  If cboPeriod.Enabled = True Then
    cboPeriod.BackColor = vbWindowBackground
  Else
    cboPeriod.BackColor = vbButtonFace
  End If

  If cboStartDate.Enabled = True Then
    cboStartDate.BackColor = vbWindowBackground
  Else
    cboStartDate.BackColor = vbButtonFace
  End If

  RefreshRoleToPromptCombo

End Sub


Private Sub cmdNew_Click()
  
  Dim pstrAddString As String
  Dim frmItem As frmBatchJobJobSelection
  
  'mblnDontDoActivateCheck = True
  
  Set frmItem = New frmBatchJobJobSelection
    
  With frmItem
    If gblnReportPackMode Then .Caption = "Report Pack Item Selection"
    .DefinitionCreator = mblnDefinitionCreator
    .BatchJobHiddenGroups = BatchJobHiddenGroups
    
    .Show vbModal
    
    If Not .Cancelled Then
      
      Changed = True
      
      'Compose the AddString to add to the grid
            
      pstrAddString = .cboJobType.Text & vbTab
      
      'If .cboJobType.Text = "-- Pause --" Then
      'If .cboJobName.Enabled Then
      If (.cboJobType.Text = "-- Pause --") Then
        pstrAddString = pstrAddString & "0" & vbTab
        pstrAddString = pstrAddString & "" & vbTab
        pstrAddString = pstrAddString & .txtParameter.Text
        
      ElseIf Not JobTypeRequiresDef(.cboJobType.Text) Then
        pstrAddString = pstrAddString & "0" & vbTab
        pstrAddString = pstrAddString & "" & vbTab
        pstrAddString = pstrAddString & "N/A"
      
      Else
        pstrAddString = pstrAddString & .cboJobName.ItemData(.cboJobName.ListIndex) & vbTab
        pstrAddString = pstrAddString & .cboJobName.Text & vbTab
        pstrAddString = pstrAddString & "N/A"
        
      End If

      grdColumns.AddItem pstrAddString

      DoEvents

      With grdColumns
        .MoveLast
        .SelBookmarks.Add .Bookmark
      End With
      
      DoEvents
      
      'If .cboJobType.Text <> "-- Pause --" Then
      ForceDefinitionToBeHiddenIfNeeded
    End If
  End With
  
  Unload frmItem
  Set frmItem = Nothing

  'mblnDontDoActivateCheck = False

End Sub

Private Sub SetButtonState()

  If mblnReadOnly Then Exit Sub
    
'#############################################

  If grdColumns.Rows = 1 Then
    grdColumns.MoveFirst
  End If
  
  If grdColumns.Rows = 0 Then
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdClearAll.Enabled = False
    cmdMoveUp.Enabled = False
    cmdMoveDown.Enabled = False
    Exit Sub
  Else
    'If grdColumns.AddItemRowIndex(grdColumns.Bookmark) > 0 Or mblnGridFlag = True Then
    If grdColumns.AddItemRowIndex(grdColumns.Bookmark) > 0 Or mblnGridFlag = True Or grdColumns.SelBookmarks.Count > 0 Then
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      cmdClearAll.Enabled = True
    End If
  End If

  'TM20030115 Fault 4797 - only enable any of the move buttons if more than one row.
  If grdColumns.AddItemRowIndex(grdColumns.Bookmark) < 1 Then
    cmdMoveUp.Enabled = False
    cmdMoveDown.Enabled = (grdColumns.Rows > 1)
  ElseIf grdColumns.AddItemRowIndex(grdColumns.Bookmark) = (grdColumns.Rows - 1) Then
    cmdMoveUp.Enabled = (grdColumns.Rows > 1)
    cmdMoveDown.Enabled = False
  Else
    cmdMoveUp.Enabled = (grdColumns.Rows > 1)
    cmdMoveDown.Enabled = (grdColumns.Rows > 1)
  End If

  If grdColumns.SelBookmarks.Count = 0 Then
    cmdMoveUp.Enabled = False
    cmdMoveDown.Enabled = False
  End If
  
End Sub

Private Sub grdColumns_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  
  With grdColumns

    ' Set the styleSet of the rows to show which is selected.
    For iLoop = 0 To .Rows - 1
      If (mblnReadOnly) Then
        .Columns(0).CellStyleSet "ssetDisabled", iLoop
        .Columns(1).CellStyleSet "ssetDisabled", iLoop
        .Columns(2).CellStyleSet "ssetDisabled", iLoop
        .Columns(3).CellStyleSet "ssetDisabled", iLoop
        .Columns(4).CellStyleSet "ssetDisabled", iLoop
      Else
        If iLoop = .Row Then
          .Columns(0).CellStyleSet "ssetSelected", iLoop
          .Columns(1).CellStyleSet "ssetSelected", iLoop
          .Columns(2).CellStyleSet "ssetSelected", iLoop
          .Columns(3).CellStyleSet "ssetSelected", iLoop
          .Columns(4).CellStyleSet "ssetSelected", iLoop
        Else
          .Columns(0).CellStyleSet "ssetEnabled", iLoop
          .Columns(1).CellStyleSet "ssetEnabled", iLoop
          .Columns(2).CellStyleSet "ssetEnabled", iLoop
          .Columns(3).CellStyleSet "ssetEnabled", iLoop
          .Columns(4).CellStyleSet "ssetEnabled", iLoop
        End If
      End If
    Next iLoop
    
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark

    If Not mblnReadOnly Then
      If .AddItemRowIndex(.Bookmark) = 0 Then
        Me.cmdMoveUp.Enabled = False
        Me.cmdMoveDown.Enabled = .Rows > 1
      ElseIf .AddItemRowIndex(.Bookmark) = (.Rows - 1) Then
        Me.cmdMoveUp.Enabled = .Rows > 1
        Me.cmdMoveDown.Enabled = False
      Else
        Me.cmdMoveUp.Enabled = .Rows > 1
        Me.cmdMoveDown.Enabled = .Rows > 1
      End If
    End If

  End With

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub grdColumns_RowLoaded(ByVal Bookmark As Variant)
  
  With grdColumns

    If (mblnReadOnly) Then
      .Columns(0).CellStyleSet "ssetDisabled"
      .Columns(1).CellStyleSet "ssetDisabled"
      .Columns(2).CellStyleSet "ssetDisabled"
      .Columns(3).CellStyleSet "ssetDisabled"
      .Columns(4).CellStyleSet "ssetDisabled"
    Else
      .Columns(0).CellStyleSet "ssetEnabled"
      .Columns(1).CellStyleSet "ssetEnabled"
      .Columns(2).CellStyleSet "ssetEnabled"
      .Columns(3).CellStyleSet "ssetEnabled"
      .Columns(4).CellStyleSet "ssetEnabled"
    End If
   
  End With

End Sub

Private Sub cmdDelete_Click()
  
  'Remove the selected row from the Batch Job Definition.
  
  Dim lRow As Long
    
  With grdColumns
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

  ForceDefinitionToBeHiddenIfNeeded

  Changed = True
    
End Sub

Private Sub cmdEdit_Click()
  
  ' Edit the row in the Batch Job Definition
  On Error GoTo EditERROR
  
  Dim pstrAddString As String
  Dim lRow As Long
  Dim frmItem As frmBatchJobJobSelection
  Set frmItem = New frmBatchJobJobSelection
  
  'mblnDontDoActivateCheck = True
  'frmItem.mblnLoading = True
  With grdColumns
    lRow = .AddItemRowIndex(.Bookmark)
    frmItem.Initialise .Columns("Job Type").Text, .Columns("Job Name").Text, .Columns("Parameter").Text
  End With
  
  frmItem.DefinitionCreator = mblnDefinitionCreator
  frmItem.BatchJobHiddenGroups = BatchJobHiddenGroups
  
  If frmBatchJob.SSTab1.TabVisible(2) Then frmItem.Caption = "Report Pack Item Selection"
  
  frmItem.Show vbModal
  
  'Me.Changed = frmItem.Changed
  
  If Not frmItem.Cancelled Then
    With frmItem
      pstrAddString = .cboJobType.Text & vbTab
      
      If (.cboJobType.Text = "-- Pause --") Then
        pstrAddString = pstrAddString & "0" & vbTab
        pstrAddString = pstrAddString & "" & vbTab
        pstrAddString = pstrAddString & .txtParameter.Text
        
      ElseIf Not JobTypeRequiresDef(.cboJobType.Text) Then
        pstrAddString = pstrAddString & "0" & vbTab
        pstrAddString = pstrAddString & "" & vbTab
        pstrAddString = pstrAddString & "N/A"
      
      Else
        pstrAddString = pstrAddString & .cboJobName.ItemData(.cboJobName.ListIndex) & vbTab
        pstrAddString = pstrAddString & .cboJobName.Text & vbTab
        pstrAddString = pstrAddString & "N/A"
        
      End If
      
    End With
    
    With grdColumns
      .RemoveItem lRow
      .AddItem pstrAddString, lRow
      .Bookmark = .AddItemBookmark(lRow)
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End With
    Me.Changed = True
    ForceDefinitionToBeHiddenIfNeeded
  End If
  
  Unload frmItem
  Set frmItem = Nothing
  
  'mblnDontDoActivateCheck = False
  
  Exit Sub
  
EditERROR:
  
  If Not frmItem Is Nothing Then
    Unload frmItem
    Set frmItem = Nothing
  End If
  
  'mblnDontDoActivateCheck = False
  
End Sub

Private Sub chkIndefinitely_Click()
  
  cboEndDate.Enabled = _
  IIf(chkIndefinitely.Value = 1, False, True)    'Disable EndDate if Indefinitely=1
  
  If cboEndDate.Enabled = True Then
    cboEndDate.BackColor = &H80000005
  Else
    cboEndDate.BackColor = &H8000000F
    cboEndDate.Text = ""
  End If
  
  Changed = True
  
End Sub

Private Sub cmdOK_Click()

  If Not ValidateDefinition Then Exit Sub
  If IsReportPack Then
    If Not SaveDefinition2 Then Exit Sub
  Else
    If Not SaveDefinition Then Exit Sub
  End If
  Unload Me
  
End Sub

Private Sub cmdCancel_Click()
  'Changed = False
  Unload Me
End Sub

Private Function SaveDefinition() As Boolean
  
  Dim sSQL As String
  Dim lCount As Long
  Dim lBatchJobID As Long
  Dim rsBatch As New Recordset
  Dim pstrPeriod As String
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim iUtilityType As UtilityType

  On Error GoTo Err_Trap
  
  Screen.MousePointer = vbHourglass              'We're gonna do something now

  If cboPeriod.Text = "Day(s)" Then                               'Period
    pstrPeriod = pstrPeriod & "'D'"
  ElseIf cboPeriod.Text = "Week(s)" Then
    pstrPeriod = pstrPeriod & "'W'"
  ElseIf cboPeriod.Text = "Month(s)" Then
    pstrPeriod = pstrPeriod & "'M'"
  ElseIf cboPeriod.Text = "Year(s)" Then
    pstrPeriod = pstrPeriod & "'Y'"
  End If

  iUtilityType = IIf(IsReportPack, utlReportPack, utlBatchJob)

  If mlngBatchJobID > 0 Then                       'Editing an existing batch job

    sSQL = "UPDATE ASRSysBatchJobName SET " & _
             "IsBatch = 1" & "," & _
             "Scheduled = " & IIf(chkScheduled.Value = 1, 1, 0) & "," & _
             "Name = '" & Trim(Replace(Me.txtName.Text, "'", "''")) & "'," & _
             "Description = '" & Replace(Me.txtDesc.Text, "'", "''") & "'," & _
             "Frequency = " & Me.spnFrequency.Value & "," & _
             "Period = " & pstrPeriod & ","

    If Not IsDate(cboStartDate.Text) Then
       sSQL = sSQL & "StartDate = Null,"
    Else
       sSQL = sSQL & "StartDate = '" & Replace(Format(CDate(cboStartDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & ","
    End If
    
    sSQL = sSQL & "Indefinitely = " & IIf(chkIndefinitely.Value = 1, 1, 0) & ","
    
    If Not IsDate(cboEndDate.Text) Then
      sSQL = sSQL & "EndDate = Null,"
    Else
      sSQL = sSQL & "EndDate = '" & Replace(Format(CDate(cboEndDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',"
    End If


    'MH20040415
    sSQL = sSQL & "EmailFailed = " & CStr(Val(txtEmailNotifyGroup(0).Tag)) & "," & _
                  "EmailSuccess = " & CStr(Val(txtEmailNotifyGroup(1).Tag)) & ","

'TM20010821 Fault 2513
'Removed code that set the 'Last Completed' date to Null.
'Therefore when editing the definition, the history of the batch job stays the same.
          
    sSQL = sSQL & "Weekends = " & IIf(chkWeekEnds.Value = 1, 1, 0) & "," & _
             "RunOnce = " & IIf(chkRunOnce.Value = 1, 1, 0) & "," & _
             "RoleToPrompt = '" & cboRoleToPrompt.Text & "'" & _
              " WHERE ID = " & mlngBatchJobID

    If ForceDefinitionToBeHiddenIfNeeded(True) = False Then
      SaveDefinition = False
      Screen.MousePointer = vbDefault
      Exit Function
    End If
          
     mclsData.ExecuteSql (sSQL)
  
    Call UtilUpdateLastSaved(iUtilityType, mlngBatchJobID)
   
  
  Else
                                                 'A NEW Batch Job Definition
    sSQL = "Insert ASRSysBatchJobName (" & _
           "IsBatch, Scheduled, Name, Description, Frequency, " & _
           "Period, StartDate, Indefinitely, " & _
           "EndDate, Weekends, " & _
           "UserName, " & _
           "RunOnce, RoleToPrompt, EmailFailed, EmailSuccess) "

    sSQL = sSQL & _
           "Values(1, " & _
           IIf(chkScheduled.Value = 1, 1, 0) & ",'" & _
           Trim(Replace(txtName.Text, "'", "''")) & "','" & _
           Replace(txtDesc.Text, "'", "''") & "'," & _
           Me.spnFrequency.Value & "," & _
           pstrPeriod & ","

'          If IsNull(cboStartDate.Text) Then
    If Not IsDate(cboStartDate.Text) Then
       sSQL = sSQL & "Null,"
    Else
       sSQL = sSQL & "'" & Replace(Format(CDate(cboStartDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & ","
    End If
           
    sSQL = sSQL & IIf(chkIndefinitely.Value = 1, 1, 0) & ","
    
'          If IsNull(cboEndDate.Text) Then
    If Not IsDate(cboEndDate.Text) Then
       sSQL = sSQL & "Null,"
    Else
       sSQL = sSQL & "'" & Replace(Format(CDate(cboEndDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & ","
    End If
    
    sSQL = sSQL & IIf(chkWeekEnds.Value = 1, 1, 0) & ",'" & _
           datGeneral.UserNameForSQL & "'," & _
           IIf(chkRunOnce.Value = 1, 1, 0) & ",'" & _
           cboRoleToPrompt.Text & "'," '"')"

    'MH20040415
    sSQL = sSQL & CStr(Val(txtEmailNotifyGroup(0).Tag)) & "," & _
                  CStr(Val(txtEmailNotifyGroup(1).Tag)) & ")"


    If ForceDefinitionToBeHiddenIfNeeded(True) = False Then
      SaveDefinition = False
      Screen.MousePointer = vbDefault
      Exit Function
    End If

    mlngBatchJobID = InsertBatchJob(sSQL)

    If mlngBatchJobID = 0 Then
      SaveDefinition = False
      Exit Function
    End If

    Call UtilCreated(iUtilityType, mlngBatchJobID)
    
  End If

  SaveAccess
  SaveObjectCategories cboCategory, iUtilityType, mlngBatchJobID

  ' Now save the column details

  ' First, remove any records from the detail table with the specified BatchJobID
  ClearBatchJobDetails

  ' Loop through the details grid
  With grdColumns

    .MoveFirst

    Do Until pintLoop = .Rows

      pvarbookmark = .GetBookmark(pintLoop)

      sSQL = "INSERT ASRSysBatchJobDetails (" & _
             "BatchJobNameID, " & _
             "JobType, " & _
             "JobID, " & _
             "Parameter, " & _
             "JobOrder)"

      sSQL = sSQL & " VALUES(" & mlngBatchJobID & ", "

      sSQL = sSQL & "'" & .Columns("Job Type").CellText(pvarbookmark) & "', "
      sSQL = sSQL & .Columns("IndividualJobID").CellText(pvarbookmark) & ", "
      sSQL = sSQL & "'" & Replace(.Columns("Parameter").CellText(pvarbookmark), "'", "''") & "', "
      sSQL = sSQL & .AddItemRowIndex(pvarbookmark) & ")"

      mclsData.ExecuteSql (sSQL)
      
      pintLoop = pintLoop + 1

    Loop

  End With

  SaveDefinition = True
  Changed = False

  Exit Function

Err_Trap:

  COAMsgBox "Error whilst saving " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
  SaveDefinition = False

End Function


Private Function SaveDefinition2() As Boolean
  
  Dim sSQL As String
  Dim lCount As Long
  Dim lBatchJobID As Long
  Dim rsBatch As New Recordset
  Dim pstrPeriod As String
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant

  On Error GoTo Err_Trap
  
  Screen.MousePointer = vbHourglass
  
  'Period is used in existing and NEW so need to set it up here somewhere
  Select Case cboPeriod.Text
      Case "Day(s)": pstrPeriod = pstrPeriod & "'D'"
      Case "Week(s)": pstrPeriod = pstrPeriod & "'W'"
      Case "Month(s)": pstrPeriod = pstrPeriod & "'M'"
      Case "Year(s)": pstrPeriod = pstrPeriod & "'Y'"
  End Select
  'sSQL = sSQL & IIf("", (T), (F))    - useful ~IIF template I've been using to build SQL can be deleted if too cluttered
  
  'Editing an existing batch job
  If mlngBatchJobID > 0 Then
  
    sSQL = "UPDATE ASRSysBatchJobName SET "
    
    'MUST set IsBatch variable to distinguish between Batch Job and Report Pack
    sSQL = sSQL & "IsBatch = 0,"
    'DEFINITION TAB
    sSQL = sSQL & "Name = '" & Trim(Replace(Me.txtName.Text, "'", "''")) & "',"
    sSQL = sSQL & "Description = '" & Replace(Me.txtDesc.Text, "'", "''") & "',"
    'SCHEDULING
    sSQL = sSQL & "Scheduled = " & IIf(chkScheduled.Value = 1, 1, 0) & ","
    sSQL = sSQL & "Frequency = " & Me.spnFrequency.Value & ","
    sSQL = sSQL & "Period = " & pstrPeriod & ","
    
    If Not IsDate(cboStartDate.Text) Then
       sSQL = sSQL & "StartDate = Null,"
    Else
       sSQL = sSQL & "StartDate = '" & Replace(Format(CDate(cboStartDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & ","
    End If
    
    If Not IsDate(cboEndDate.Text) Then
      sSQL = sSQL & "EndDate = Null,"
    Else
      sSQL = sSQL & "EndDate = '" & Replace(Format(CDate(cboEndDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',"
    End If
    
    sSQL = sSQL & "RoleToPrompt = '" & cboRoleToPrompt.Text & "', "
    sSQL = sSQL & "Indefinitely = " & IIf(chkIndefinitely.Value = 1, 1, 0) & ","
    sSQL = sSQL & "Weekends = " & IIf(chkWeekEnds.Value = 1, 1, 0) & ","
    sSQL = sSQL & "RunOnce = " & IIf(chkRunOnce.Value = 1, 1, 0) & ","
    
    'JOBS TAB
    'sSQL = sSQL & IIf(chkDestination(desSave), ("OutputSaveFormat = " & Val(txtFileName.Tag) & ", "), ("OutputSaveFormat = 0, "))
    'EMAIL NOTIFICATION FRAME
    sSQL = sSQL & "EmailFailed = " & CStr(Val(txtEmailNotifyGroup(0).Tag)) & ","
    sSQL = sSQL & "EmailSuccess = " & CStr(Val(txtEmailNotifyGroup(1).Tag)) & ","
    
    If IsReportPack Then
    
      'REPORT OPTIONS FRAME
      sSQL = sSQL & "OutputTitlePage = '" & Replace(txtTitlePage.Text, "'", "''") & "', "             'Title Pge Template
      SaveUserSetting "Output", "PackTemplate", txtTitlePage.Text                                     'Title Pge Template for usersetting in clOuputWord
      sSQL = sSQL & "OutputReportPackTitle = '" & Replace(txtReportPackTitle.Text, "'", "''") & "',"  'Report Pack Title
      sSQL = sSQL & "OutputOverrideFilter = '" & Replace(txtOverrideFilter.Text, "'", "''") & "',"    'Override Filter
      sSQL = sSQL & "OverrideFilterID = '" & Replace(txtOverrideFilter.Tag, "'", "''") & "',"         'Override FilterID
      sSQL = sSQL & "OutputTOC = " & IIf(chkTOC.Value = 1, 1, 0) & ","                                'Table of Contents
      sSQL = sSQL & "OutputCoverSheet = " & IIf(chkForceCoverSheet.Value = 1, 1, 0) & ","             'Force Cover sheet
      sSQL = sSQL & "OutputRetainPivotOrChart = " & IIf(chkRetainPivot.Value = 1, 1, 0) & ","         'Retain Pivot
      'sSQL = sSQL & "OutputRetainCharts = " & IIf(ChkRetainCharts.Value = 1, 1, 0) & ","         'Retain chart
      
      'OUTPUT FORMAT FRAME
      sSQL = sSQL & "OutputFormat = " & CStr(mobjOutputDef.GetSelectedFormatIndex) & ", "
      
      'OUTPUT DESTINATION FRAME
      sSQL = sSQL & "OutputPreview = 0" & ", " ' " & IIf(chkPreview.Value = vbChecked, "1", "0") & ", "
      sSQL = sSQL & "OutputScreen = " & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", "
      'Printer Options
      sSQL = sSQL & IIf(chkDestination(desPrinter), (" OutputPrinterName = '" & Replace(cboPrinterName.Text, " '", "''") & "',"), (" OutputPrinterName = '', "))
      sSQL = sSQL & "OutputFilename = '" & Replace(txtFileName.Text, "'", "''") & "',"
      'outputSaveExisting
      If chkDestination(desSave).Value = vbChecked Then
        sSQL = sSQL & "OutputSaveExisting = " & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
      Else
        sSQL = sSQL & "OutputSaveExisting = 0, "
      End If
      'Save Format
      sSQL = sSQL & IIf(chkDestination(desSave), (" OutputSave = 1, "), (" OutputSave = 0, "))
      sSQL = sSQL & IIf(chkDestination(desPrinter), (" OutputPrinter = 1, "), (" OutputPrinter = 0, "))
    End If
      
    'Email Options
    sSQL = sSQL & IIf(chkDestination(desEmail), ("OutputEmail = 1, "), ("OutputEmail = 0, "))
    sSQL = sSQL & IIf(chkDestination(desEmail), ("OutputEmailAddr = " & txtEmailGroup.Tag & ", "), ("OutputEmailAddr = 0, "))
    sSQL = sSQL & IIf(chkDestination(desEmail), ("OutputEmailSubject = '" & Replace(txtEmailSubject.Text, "'", "''") & "', "), ("OutputEmailSubject = '', "))
    sSQL = sSQL & IIf(chkDestination(desEmail), ("OutputEmailAttachAs = '" & Replace(txtEMailAttachAs.Text, "'", "''") & "'"), ("OutputEmailAttachAs = ''"))
    
    'FINAL WHERE CLAUSE
    sSQL = sSQL & " WHERE ID = " & mlngBatchJobID
    
    If ForceDefinitionToBeHiddenIfNeeded(True) = False Then
      SaveDefinition2 = False
      Screen.MousePointer = vbDefault
      Exit Function
    End If
          
     mclsData.ExecuteSql (sSQL)
  
    Call UtilUpdateLastSaved(utlReportPack, mlngBatchJobID)
  Else
    'A NEW Batch Job Definition
    sSQL = "Insert ASRSysBatchJobName (" & _
          "Scheduled, Name, Description, Frequency, " & _
          "Period, StartDate, Indefinitely, " & _
          "EndDate, Weekends, " & _
          "UserName, " & _
          "RunOnce, RoleToPrompt, EmailFailed, EmailSuccess," & _
          "IsBatch, OutputPreview, OutputFormat, OutputScreen," & _
          "OutputPrinter, OutputPrinterName, OutputSave, " & _
          "OutputSaveExisting, OutputEmail, OutputEmailAddr, " & _
          "OutputEmailSubject, OutputFilename, OutputEmailAttachAs, " & _
          "OutputTitlePage, OutputReportPackTitle, OutputOverrideFilter, " & _
          "OutputTOC, OutputCoverSheet, OverrideFilterID, OutputRetainPivotOrChart)"
          
    sSQL = sSQL & _
           "Values(" & _
           IIf(chkScheduled.Value = 1, 1, 0) & ",'" & _
           Trim(Replace(txtName.Text, "'", "''")) & "','" & _
           Replace(txtDesc.Text, "'", "''") & "'," & _
           Me.spnFrequency.Value & "," & _
           pstrPeriod & ","
           
          If Not IsDate(cboStartDate.Text) Then
             sSQL = sSQL & "Null,"
          Else
             sSQL = sSQL & "'" & Replace(Format(CDate(cboStartDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & ","
          End If
                 
          sSQL = sSQL & IIf(chkIndefinitely.Value = 1, 1, 0) & ","
          
          If Not IsDate(cboEndDate.Text) Then
             sSQL = sSQL & "Null,"
          Else
             sSQL = sSQL & "'" & Replace(Format(CDate(cboEndDate.Text), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'" & ","
          End If
    
    sSQL = sSQL & IIf(chkWeekEnds.Value = 1, 1, 0) & ",'" & _
           datGeneral.UserNameForSQL & "'," & _
           IIf(chkRunOnce.Value = 1, 1, 0) & ",'" & _
           cboRoleToPrompt.Text & "',"
          
    sSQL = sSQL & CStr(Val(txtEmailNotifyGroup(0).Tag)) & "," & _
                  CStr(Val(txtEmailNotifyGroup(1).Tag)) & ","
          'IsBatch
          sSQL = sSQL & "0,"
          'outputPreview
          sSQL = sSQL & "0,"
          'outputformat
          sSQL = sSQL & CStr(mobjOutputDef.GetSelectedFormatIndex) & ", "
          'outputScreen
          sSQL = sSQL & IIf(chkDestination(desScreen).Value = vbChecked, "1", "0") & ", "
          'outputPrinter
          sSQL = sSQL & IIf(chkDestination(desPrinter), (" 1, "), (" 0, "))
          'outputPrinterName
          sSQL = sSQL & IIf(chkDestination(desPrinter), ("'" & Replace(cboPrinterName.Text, " '", "''") & "',"), ("'', "))
          'outputSave
          sSQL = sSQL & IIf(chkDestination(desSave), (" 1, "), (" 0, "))
          'outputSaveExisting
          If chkDestination(desSave).Value = vbChecked Then
            sSQL = sSQL & cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ", "
          Else
            sSQL = sSQL & "0, "
          End If
          'sSQL = sSQL & IIf(chkDestination(desSave).Value = vbChecked, (cboSaveExisting.ItemData(cboSaveExisting.ListIndex) & ","), ("0,"))
          'outputEmail
          sSQL = sSQL & IIf(chkDestination(desEmail), ("1, "), ("0, "))
          'outputEmailAddr
          sSQL = sSQL & IIf(chkDestination(desEmail), (txtEmailGroup.Tag & ", "), ("0, "))
          'outputEmailSubject
          sSQL = sSQL & IIf(chkDestination(desEmail), ("'" & Replace(txtEmailSubject.Text, "'", "''") & "', "), ("'', "))
          'outputFilename
          sSQL = sSQL & "'" & Replace(txtFileName.Text, "'", "''") & "',"
          'outputEmailAttachAs
          sSQL = sSQL & IIf(chkDestination(desEmail), ("'" & Replace(txtEMailAttachAs.Text, "'", "''") & "',"), ("'',"))
          'outputTitlePage
          sSQL = sSQL & "'" & Replace(txtTitlePage.Text, "'", "''") & "', "
          'outputReportPackTitle
          sSQL = sSQL & "'" & Replace(txtReportPackTitle.Text, "'", "''") & "',"
          'outputOverrideFilter
          sSQL = sSQL & "'" & Replace(txtOverrideFilter.Text, "'", "''") & "',"
          'outputTOC
          sSQL = sSQL & IIf(chkTOC.Value = 1, 1, 0) & ","
          'outputCoverSheet
          sSQL = sSQL & IIf(chkForceCoverSheet.Value = 1, 1, 0) & ","
          'Override Filter ID
          sSQL = sSQL & Replace(txtOverrideFilter.Tag, "'", "''") & ","
          'Retain Pivot or Chart when excel selected
          sSQL = sSQL & IIf(chkRetainPivot.Value = 1, 1, 0) & ")"
          
    If ForceDefinitionToBeHiddenIfNeeded(True) = False Then
      SaveDefinition2 = False
      Screen.MousePointer = vbDefault
      Exit Function
    End If

    mlngBatchJobID = InsertBatchJob(sSQL)

    If mlngBatchJobID = 0 Then
      SaveDefinition2 = False
      Exit Function
    End If

    Call UtilCreated(utlReportPack, mlngBatchJobID)
    
  End If

  SaveAccess
  SaveObjectCategories cboCategory, utlReportPack, mlngBatchJobID

  ' Now save the column details

  ' First, remove any records from the detail table with the specified BatchJobID
  ClearBatchJobDetails

  ' Loop through the details grid
  With grdColumns

    .MoveFirst

    Do Until pintLoop = .Rows

      pvarbookmark = .GetBookmark(pintLoop)

      sSQL = "INSERT ASRSysBatchJobDetails (" & _
             "BatchJobNameID, " & _
             "JobType, " & _
             "JobID, " & _
             "Parameter, " & _
             "JobOrder)"

      sSQL = sSQL & " VALUES(" & mlngBatchJobID & ", "

      sSQL = sSQL & "'" & .Columns("Job Type").CellText(pvarbookmark) & "', "
      sSQL = sSQL & .Columns("IndividualJobID").CellText(pvarbookmark) & ", "
      sSQL = sSQL & "'" & Replace(.Columns("Parameter").CellText(pvarbookmark), "'", "''") & "', "
      sSQL = sSQL & .AddItemRowIndex(pvarbookmark) & ")"

      mclsData.ExecuteSql (sSQL)
      
      pintLoop = pintLoop + 1

    Loop

  End With

  SaveDefinition2 = True
  Changed = False

  Exit Function

Err_Trap:

  COAMsgBox "Error whilst saving " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
  SaveDefinition2 = False

End Function


Private Function InsertBatchJob(pstrSQL As String) As Long
  
' Save Definition and return new ID of definition
'  Dim sSQL As String
'  Dim rsBatch As Recordset
'
'  mclsData.ExecuteSql psSQL
'
'  sSQL = "SELECT MAX(id) FROM ASRSysBatchJobName"
'  Set rsBatch = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  InsertBatchJob = rsBatch(0)
'
'  rsBatch.Close
'  Set rsBatch = Nothing

  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fSavedOK As Boolean
  
  On Error GoTo InsertBatchJob_ERROR
  
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
    pmADO.Value = "AsrSysBatchJobName"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "ID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, app.ProductName
        InsertBatchJob = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertBatchJob = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function

InsertBatchJob_ERROR:

  fSavedOK = False
  Resume Next

End Function

Private Sub ClearBatchJobDetails()

  'Delete individual jobs contained within the current definition
  Dim sSQL As String

  sSQL = "DELETE FROM ASRSysBatchJobDetails WHERE BatchJobNameID = " & mlngBatchJobID
  mclsData.ExecuteSql sSQL

End Sub

Private Function ValidateDefinition() As Boolean
  
  'Purpose : Check all mandatory informaiton is entered and also check that
  '          If there is a problem with validation, the program will display
  '          the tab containing the problem to the user.
  'Input   : None
  'Output  : True/False
  
  On Error GoTo ValidateDefinition_ERROR
  
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim blnDestination As Boolean
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If ValidateGTMaskDate(cboStartDate) = False Or _
     ValidateGTMaskDate(cboEndDate) = False Then
        ValidateDefinition = False
        Exit Function
  End If


  ' Check a name has been entered
  If Trim(txtName.Text) = "" Then
    COAMsgBox "You must give this " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition a name.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
    SSTab1.Tab = 0
    txtName.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  'Check if this definition has been changed by another user
  Call UtilityAmended(utlBatchJob, mlngBatchJobID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    
    mblnDefinitionCreator = True
    mlngBatchJobID = 0
  
    UI.LockWindow grdAccess.hWnd
    
    With grdAccess
      For iLoop = 0 To (.Rows - 1)
        varBookmark = .AddItemBookmark(iLoop)
        .Bookmark = varBookmark
      Next iLoop
      
      .MoveFirst
    End With
  
    UI.UnlockWindow
  End If
  
  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngBatchJobID, gblnReportPackMode) Then
    COAMsgBox "A " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Check for a start date
  If IsNull(cboStartDate.DateValue) And chkScheduled.Value = vbChecked And Len(Trim(Replace(cboStartDate.Text, UI.GetSystemDateSeparator, ""))) = 0 Then
    COAMsgBox "You must select a start date if you wish the " & IIf(gblnReportPackMode, "report pack", "batch job") & " to be scheduled.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
    SSTab1.Tab = 0
    cboStartDate.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If ValidateGTMaskDate(cboStartDate) = False Then
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Check for an end date if not indefinetly
  If IsNull(cboEndDate.DateValue) And chkScheduled.Value = vbChecked And chkIndefinitely.Value = vbUnchecked Then
    COAMsgBox "You must select a valid end date or run indefinitely if you wish the " & IIf(gblnReportPackMode, "report pack", "batch job") & " to be scheduled.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
    SSTab1.Tab = 0
    cboEndDate.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If ValidateGTMaskDate(cboEndDate) = False Then
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Check end date is after start date
  If chkScheduled.Value = vbChecked And chkIndefinitely.Value = vbUnchecked And (cboStartDate.DateValue > cboEndDate.DateValue) Then
    COAMsgBox "Start date must be prior to end date.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
    SSTab1.Tab = 0
    cboEndDate.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Check that there are jobs defined in the batch job
  If grdColumns.Rows = 0 Or (OnlyPauseJobsDefined) Then
    COAMsgBox "You must select at least 1" & IIf(gblnReportPackMode, " ", " (non pause) ") & "job in your " & IIf(gblnReportPackMode, "report pack", "batch job") & ".", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
    SSTab1.Tab = 1
    Exit Function
  End If

  ' Ensure that a valid user group has been selected.
  ' (Needed incase a user group has been deleted)
  If chkScheduled.Value Then
    If cboRoleToPrompt.Text = "" Then
      COAMsgBox "You must select a valid user group for this " & IIf(gblnReportPackMode, "report pack", "batch job") & ".", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
      SSTab1.Tab = 0
      cboRoleToPrompt.SetFocus
      ValidateDefinition = False
      Exit Function
    End If
  End If
  
  
  If chkEmail(0).Value = vbChecked And Val(txtEmailNotifyGroup(0).Tag) = 0 Then
    SSTab1.Tab = 1
    COAMsgBox "You must select an email group for failed job notification.", vbExclamation, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
    ValidateDefinition = False
    Exit Function
  End If
  
  If chkEmail(1).Value = vbChecked And Val(txtEmailNotifyGroup(1).Tag) = 0 Then
    SSTab1.Tab = 1
    COAMsgBox "You must select an email group for successful job notification.", vbExclamation, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
    ValidateDefinition = False
    Exit Function
  End If
  
  If gblnReportPackMode Then
    If Not ValidDestination Then
      SSTab1.Tab = 2
      Exit Function
    End If
  End If
  
  ValidateDefinition = True

  Exit Function
  
ValidateDefinition_ERROR:
  
  COAMsgBox IIf(gblnReportPackMode, "Error whilst validating report pack definition.", "Error whilst validating " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition.") & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
  ValidateDefinition = False

End Function
Private Function ValidDestination() As Boolean

  Dim blnDestination As Boolean
  Dim blnPreview As Boolean

  ValidDestination = False

  If chkDestination(desSave).Value = vbChecked Then
    If txtFileName.Text = vbNullString Then
      COAMsgBox "You must enter a file name.", vbExclamation, Caption
      Exit Function
    End If
  End If

  If chkDestination(desEmail).Value = vbChecked Then
  
    If txtFileName.Text = vbNullString Then
      COAMsgBox "You must enter a file name.", vbExclamation, Caption
      Exit Function
    End If
    
    Select Case mobjOutputDef.Format
    Case fmtWordDoc, fmtExcelWorksheet, fmtExcelchart, fmtExcelPivotTable
      If txtEMailAttachAs.Text Like "*.html" Then
        COAMsgBox "You cannot email html output from word or excel.", vbExclamation, Caption
        Exit Function
      End If
    End Select
  
    If Val(txtEmailGroup.Tag) = 0 Then
      COAMsgBox "You must select an email group.", vbExclamation, Caption
      Exit Function
    End If
    
    

    If datGeneral.GetEmailGroupName(Val(txtEmailGroup.Tag)) = vbNullString Then
      COAMsgBox "The email group has been deleted by another user.", vbExclamation, Caption
      txtEmailGroup.Text = vbNullString
      txtEmailGroup.Tag = 0
      Exit Function
    End If

    If txtEMailAttachAs.Text = vbNullString Then
      COAMsgBox "You must enter an email attachment file name.", vbExclamation, Caption
      Exit Function
    End If
    
    If InStr(txtEMailAttachAs.Text, "/") Or _
       InStr(txtEMailAttachAs.Text, ":") Or _
       InStr(txtEMailAttachAs.Text, "?") Or _
       InStr(txtEMailAttachAs.Text, Chr(34)) Or _
       InStr(txtEMailAttachAs.Text, "<") Or _
       InStr(txtEMailAttachAs.Text, ">") Or _
       InStr(txtEMailAttachAs.Text, "|") Or _
       InStr(txtEMailAttachAs.Text, "\") Or _
       InStr(txtEMailAttachAs.Text, "*") Then
          COAMsgBox "The email attachment file name cannot contain any of the following characters:" & vbCrLf & _
                 "/  :  ?  " & Chr(34) & "  <  >  |  \  *", vbExclamation, Caption
          Exit Function
    End If
  
  
  End If
  
  On Local Error Resume Next
  'blnPreview = (chkPreview.Value = vbChecked)

  With chkDestination
    blnDestination = False
    'blnDestination = (blnDestination Or .Item(desScreen).Value = vbChecked)
    blnDestination = (blnDestination Or .item(desPrinter).Value = vbChecked)
    blnDestination = (blnDestination Or .item(desSave).Value = vbChecked)
    blnDestination = (blnDestination Or .item(desEmail).Value = vbChecked)
  End With

  If Not blnDestination Then
    COAMsgBox "You must select a destination", vbExclamation, Caption
    Exit Function
  End If

  ValidDestination = True

End Function


Private Function ForceDefinitionToBeHiddenIfNeeded2(Optional pvOnlyFatalMessages As Variant) As Boolean

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
  Dim vForceHidden As Variant
 
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

  ' Base Table Filter
  If Len(txtOverrideFilter.Tag) > 0 And Val(txtOverrideFilter.Tag) <> 0 Then
    fRemove = False
    iResult = ValidateRecordSelection(REC_SEL_FILTER, CLng(txtOverrideFilter.Tag))

    Select Case iResult
      Case REC_SEL_VALID_HIDDENBYUSER
        ' Filter hidden by the current user.
        ' Only a problem if the current definition is NOT owned by the current user,
        ' or if the current definition is not already hidden.
        fRemove = mblnDefinitionCreator And mblnReadOnly
        
        fRemove = (Not mblnDefinitionCreator) And _
          (Not mblnReadOnly)
          
        If fRemove Then
          sBigMessage = "This table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
          COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
        Else
          fNeedToForceHidden = True
  
          ReDim Preserve asHiddenBySelfParameters(UBound(asHiddenBySelfParameters) + 1)
          asHiddenBySelfParameters(UBound(asHiddenBySelfParameters)) = "'" & txtOverrideFilter.Text & "' table filter"
        End If
        
      Case REC_SEL_VALID_DELETED
        ' Deleted by another user.
        ReDim Preserve asDeletedParameters(UBound(asDeletedParameters) + 1)
        asDeletedParameters(UBound(asDeletedParameters)) = "'" & txtOverrideFilter.Text & "' table filter"

        fRemove = mblnReadOnly

      Case REC_SEL_VALID_HIDDENBYOTHER
        ' Hidden by another user.
        If Not gfCurrentUserIsSysSecMgr Then
          ReDim Preserve asHiddenByOtherParameters(UBound(asHiddenByOtherParameters) + 1)
          asHiddenByOtherParameters(UBound(asHiddenByOtherParameters)) = "'" & txtOverrideFilter.Text & "' table filter"
  
          fRemove = mblnReadOnly
        End If

      Case REC_SEL_VALID_INVALID
        ' Picklist invalid.
        ReDim Preserve asInvalidParameters(UBound(asInvalidParameters) + 1)
        asInvalidParameters(UBound(asInvalidParameters)) = "'" & txtOverrideFilter.Text & "' table filter"

        fRemove = mblnReadOnly
         ' (Not FormPrint)
    End Select

    If fRemove Then
      ' Filter invalid, deleted or hidden by another user. Remove it from this definition.
      txtOverrideFilter.Tag = 0
      txtOverrideFilter.Text = ""
      'mblnRecordSelectionInvalid = True
    End If
  End If
  
  ' Construct one big message with all of the required error messages.
  sBigMessage = ""

  If UBound(asHiddenBySelfParameters) = 1 Then
    If mblnReadOnly Then
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
    If mblnReadOnly Then
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

    For iLoop = 1 To UBound(asHiddenByOtherParameters)
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
  
  vForceHidden = IIf(fNeedToForceHidden, "HD", "RW")
  ForceAccess vForceHidden

  If Len(sBigMessage) > 0 Then
    COAMsgBox sBigMessage, vbExclamation + vbOKOnly, Me.Caption
  End If

  ForceDefinitionToBeHiddenIfNeeded2 = (Len(sBigMessage) = 0)
  
  'RefreshRepetitionGrid
  
End Function


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
Private Function ForceDefinitionToBeHiddenIfNeeded(Optional pvOnlyFatalMessages As Variant, _
  Optional psRoleToPrompt As String) As Boolean
  ' Check if the job selection requires this Batch Job definition to be made hidden
  ' from any user groups.
  ' Return false if some of the jobs need to be removed from the definition,
  ' or if the definition needs to be made hidden.
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim varBookmark As Variant
  Dim lngJobID As Long
  Dim rsAccess As ADODB.Recordset
  Dim sUtilityType As UtilityType
  Dim avJobs() As Variant
  Dim fFound As Boolean
  Dim sUserGroupName As String
  Dim sForcedAccess As String
  Dim sCurrentAccess As String
  Dim sJobOwner As String
  Dim iIndex As Integer
  Dim asGroups() As String
  Dim sUnhiddenGroups As String
  Dim iResult As RecordSelectionValidityCodes
  Dim sBigMessage As String
  Dim asDeletedParameters() As String
  Dim asHiddenBySelfParameters() As String
  Dim asHiddenByOtherParameters() As String
  Dim asInvalidParameters() As String
  Dim asHiddenBySelfToGroupParameters() As String
  Dim fChangesRequired As Boolean
  Dim fDefnAlreadyHidden As Boolean
  Dim fNeedToForceHidden As Boolean
  Dim fRemove As Boolean
  Dim fDone As Boolean
  Dim fDefnAlreadyForcedHidden As Boolean
  Dim fMakingAndForcingHidden As Boolean
  Dim fOnlyFatalMessages As Boolean
  Dim sSysSecGroups As String
  
  If IsMissing(pvOnlyFatalMessages) Then
    fOnlyFatalMessages = mblnLoading
  Else
    fOnlyFatalMessages = CBool(pvOnlyFatalMessages)
  End If
  
  fChangesRequired = False
  fNeedToForceHidden = False
  fMakingAndForcingHidden = False
  fDefnAlreadyHidden = True
  fDefnAlreadyForcedHidden = True
  
  ' Dimension arrays to hold details of the jobs that
  ' have been deleted, made hidden or are now invalid.
  ' Column 1 - job type
  ' Column 2 - job name
  ReDim asDeletedParameters(2, 0)
  ReDim asHiddenBySelfParameters(2, 0)
  ReDim asHiddenByOtherParameters(2, 0)
  ReDim asInvalidParameters(2, 0)
  ReDim asHiddenBySelfToGroupParameters(2, 0)
  
  ' Construct an array of the Batch Jobs and their status.
  ' Column 1 = job ID
  ' Column 2 = job type
  ' Column 3 = job name
  ' Column 4 = owned by Batch Job owner ?
  ' Column 5 = deleted ?
  ' Column 6 = hidden to all users ?
  ' Column 7 = invalid ?
  ' Column 8 = list of groups to whom this job is hidden
  ' Column 9 = hidden to current user ?
  ReDim avJobs(10, 0)
  
  ' Construct an array of the User Groups and their status.
  ' Column 1 = group name
  ' Column 2 = original access to the Batch Job
  ' Column 3 = required access to the Batch Job
  ' Column 4 = original forced access to the Batch Job
  ReDim asGroups(5, 0)
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        iIndex = UBound(asGroups, 2) + 1
        ReDim Preserve asGroups(5, iIndex)
        asGroups(1, iIndex) = .Columns("GroupName").CellText(varBookmark)
        asGroups(2, iIndex) = AccessCode(.Columns("Access").CellText(varBookmark))
        asGroups(3, iIndex) = ACCESS_UNKNOWN
        asGroups(4, iIndex) = .Columns("ForcedAccess").CellText(varBookmark)
      End If
      
      If .Columns("SysSecMgr").CellText(varBookmark) <> "1" Then
        If .Columns("Access").CellText(varBookmark) <> AccessDescription(ACCESS_HIDDEN) Then
          fDefnAlreadyHidden = False
        End If
      
        If .Columns("ForcedAccess").CellText(varBookmark) <> "1" Then
          fDefnAlreadyForcedHidden = False
        End If
      End If
    Next iLoop

    .MoveFirst
  End With
  
  'JPD 20040628 Fault 8390
  ' Get the list of Sys/Sec Mgr user groups
  sSysSecGroups = vbTab
  Set rsAccess = GetSysSecMgrUserGroups
  Do While Not rsAccess.EOF
    sSysSecGroups = sSysSecGroups & UCase(Trim(rsAccess!Name)) & vbTab
    rsAccess.MoveNext
  Loop
  rsAccess.Close
  Set rsAccess = Nothing
  
  ' Check if any of the selected jobs are hidden.
  With grdColumns
    For iLoop = 0 To (.Rows - 1)
      varBookmark = .AddItemBookmark(iLoop)
      sUtilityType = JobUtilityType(.Columns("Job Type").CellText(varBookmark))
      
      If sUtilityType > 0 Then
        lngJobID = CLng(.Columns("IndividualJobID").CellText(varBookmark))
      
        iIndex = UBound(avJobs, 2) + 1
        ReDim Preserve avJobs(10, iIndex)
        avJobs(1, iIndex) = lngJobID
        avJobs(2, iIndex) = .Columns("Job Type").CellText(varBookmark)
        avJobs(3, iIndex) = .Columns("Job Name").CellValue(varBookmark)
        avJobs(4, iIndex) = False
        avJobs(5, iIndex) = False
        avJobs(6, iIndex) = False
        avJobs(7, iIndex) = False
        avJobs(8, iIndex) = vbTab
        avJobs(9, iIndex) = False
        
        ' Utility with the new style access.
        If JobTypeRequiresDef(.Columns("Job Type").CellText(varBookmark)) Then
          sJobOwner = GetUtilityOwner(sUtilityType, lngJobID)
          If Len(sJobOwner) = 0 Then
            ' Job has been deleted
            avJobs(5, iIndex) = True
          Else
            avJobs(4, iIndex) = (UCase(sJobOwner) = UCase(txtUserName.Text))
            avJobs(9, iIndex) = (CurrentUserAccess(sUtilityType, lngJobID) = ACCESS_HIDDEN)

            'JPD 20040628 Fault 8390
            'Set rsAccess = GetUtilityAccessRecords(sUtilityType, lngJobID, False)
            Set rsAccess = GetUtilityAccessRecordsIgnoreSysSecUsers(sUtilityType, lngJobID, False)
            
            Do While Not rsAccess.EOF
              'JPD 20040628 Fault 8390
              'If rsAccess!Access = ACCESS_HIDDEN Then
              If (rsAccess!Access = ACCESS_HIDDEN) And _
                (InStr(sSysSecGroups, vbTab & UCase(Trim(rsAccess!Name)) & vbTab) <= 0) Then
                ' Job is hidden from the user group.
                sUserGroupName = UCase(rsAccess!Name)

                avJobs(8, iIndex) = avJobs(8, iIndex) & _
                  sUserGroupName & vbTab
              End If

              rsAccess.MoveNext
            Loop
            rsAccess.Close
            Set rsAccess = Nothing
          End If
        End If
      End If
    Next iLoop
  End With
        
  For iLoop = 1 To UBound(avJobs, 2)
    fRemove = False
  
    If avJobs(4, iLoop) And avJobs(6, iLoop) Then
      ' Job hidden by the current user.
      fNeedToForceHidden = True

      ReDim Preserve asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2) + 1)
      asHiddenBySelfParameters(1, UBound(asHiddenBySelfParameters, 2)) = avJobs(2, iLoop)
      asHiddenBySelfParameters(2, UBound(asHiddenBySelfParameters, 2)) = avJobs(3, iLoop)

      fRemove = (Not mblnDefinitionCreator) And _
        (Not mblnReadOnly)
    End If
    
    If avJobs(5, iLoop) Then
      ' Job deleted.
      ReDim Preserve asDeletedParameters(2, UBound(asDeletedParameters, 2) + 1)
      asDeletedParameters(1, UBound(asDeletedParameters, 2)) = avJobs(2, iLoop)
      asDeletedParameters(2, UBound(asDeletedParameters, 2)) = avJobs(3, iLoop)

      fRemove = (Not mblnReadOnly)
    End If
    
    If avJobs(6, iLoop) And avJobs(9, iLoop) Then
      ' Job hidden by another user.
      ReDim Preserve asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2) + 1)
      asHiddenByOtherParameters(1, UBound(asHiddenByOtherParameters, 2)) = avJobs(2, iLoop)
      asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2)) = avJobs(3, iLoop)

      fRemove = (Not mblnReadOnly)
    End If
    
    If avJobs(7, iLoop) Then
      ' Job invalid.
      ReDim Preserve asInvalidParameters(2, UBound(asInvalidParameters, 2) + 1)
      asInvalidParameters(1, UBound(asInvalidParameters, 2)) = avJobs(2, iLoop)
      asInvalidParameters(2, UBound(asInvalidParameters, 2)) = avJobs(3, iLoop)

      fRemove = (Not mblnReadOnly)
    End If
    
    If (avJobs(8, iLoop) <> vbTab) Then
      If avJobs(9, iLoop) Then
        ' Job hidden to current user.
        ReDim Preserve asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2) + 1)
        asHiddenByOtherParameters(1, UBound(asHiddenByOtherParameters, 2)) = avJobs(2, iLoop)
        asHiddenByOtherParameters(2, UBound(asHiddenByOtherParameters, 2)) = avJobs(3, iLoop)
    
        fRemove = (Not mblnReadOnly)
      Else
        ' Job hidden to some users but not the current one.
        fDone = False
        For iLoop2 = 1 To UBound(asGroups, 2)
          If InStr(avJobs(8, iLoop), vbTab & UCase(asGroups(1, iLoop2)) & vbTab) > 0 Then
            ' Batch Job does need to be hidden from the user group.
            If (asGroups(4, iLoop2) <> "1") Then
              'JPD 20040714 Fault 8476
              If (asGroups(2, iLoop2) <> ACCESS_HIDDEN) And _
                (asGroups(2, iLoop2) <> ACCESS_UNKNOWN) Then
                fMakingAndForcingHidden = True
              End If
              
              If Not fDone Then
                fDone = True
                
                ReDim Preserve asHiddenBySelfToGroupParameters(2, UBound(asHiddenBySelfToGroupParameters, 2) + 1)
                asHiddenBySelfToGroupParameters(1, UBound(asHiddenBySelfToGroupParameters, 2)) = avJobs(2, iLoop)
                asHiddenBySelfToGroupParameters(2, UBound(asHiddenBySelfToGroupParameters, 2)) = avJobs(3, iLoop)
              End If
              
              'JPD 20040628 Fault 8476
              If Not fRemove Then
                fRemove = (Not mblnDefinitionCreator) And _
                  (Not mblnReadOnly) And _
                  (fMakingAndForcingHidden)
              End If
            End If
            
            'JPD 20040628 Fault 8280
            If Not fRemove Then
              asGroups(3, iLoop2) = ACCESS_HIDDEN
            End If
          End If
        Next iLoop2
      End If
    End If
  
    If fRemove Then
      ' Job invalid, deleted or hidden by another user. Remove it from this definition.
      With grdColumns
        For iLoop2 = 0 To (.Rows - 1)
          varBookmark = .AddItemBookmark(iLoop2)
          
          If (avJobs(1, iLoop) = CLng(.Columns("IndividualJobID").CellText(varBookmark))) And _
            (avJobs(2, iLoop) = .Columns("Job Type").CellText(varBookmark)) Then
            
            grdColumns.RemoveItem iLoop2
            Exit For
          End If
        Next iLoop2
      End With
    End If
  Next iLoop
        
  If fNeedToForceHidden Then
    For iLoop = 1 To UBound(asGroups, 2)
      asGroups(3, iLoop) = ACCESS_HIDDEN
    Next iLoop
  End If
  
  ' Construct one big message with all of the required error messages.
  sBigMessage = ""

  If UBound(asHiddenBySelfParameters, 2) = 1 Then
    If mblnReadOnly Then
      sBigMessage = "This definition needs to be made hidden from all users as the selected " & asHiddenBySelfParameters(1, 1) & " '" & asHiddenBySelfParameters(2, 1) & "' is hidden."
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If Not fDefnAlreadyForcedHidden And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the selected " & asHiddenBySelfParameters(1, 1) & " '" & asHiddenBySelfParameters(2, 1) & "' is hidden."
        End If
      Else
        If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden from all users as the selected " & asHiddenBySelfParameters(1, 1) & " '" & asHiddenBySelfParameters(2, 1) & "' is hidden."
        End If
      End If
    Else
      sBigMessage = "The selected " & asHiddenBySelfParameters(1, 1) & " '" & asHiddenBySelfParameters(2, 1) & "' will be removed from this definition as it is hidden and you do not have permission to make this definition hidden."
    End If
  ElseIf UBound(asHiddenBySelfParameters, 2) > 1 Then
    If mblnReadOnly Then
      sBigMessage = "This definition needs to be made hidden from all users as the following jobs are hidden :" & vbCrLf
    ElseIf mblnDefinitionCreator Then
      If fDefnAlreadyHidden Then
        If Not fDefnAlreadyForcedHidden And (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed as the following parameters are hidden :" & vbCrLf
        End If
      Else
        If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
          sBigMessage = "This definition will now be made hidden from all users as the following jobs are hidden :" & vbCrLf
        End If
      End If
    Else
      sBigMessage = "The following jobs will be removed from this definition as they are hidden and you do not have permission to make this definition hidden :" & vbCrLf
    End If

    If Len(sBigMessage) > 0 Then
      For iLoop = 1 To UBound(asHiddenBySelfParameters, 2)
        sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenBySelfParameters(1, iLoop) & " '" & asHiddenBySelfParameters(2, iLoop) & "'"
      Next iLoop
    End If
  ElseIf UBound(asHiddenBySelfParameters, 2) = 0 Then
    ' Not hiding to all user groups. What if its hidden to some user groups.
    If UBound(asHiddenBySelfToGroupParameters, 2) = 1 Then
      If mblnReadOnly Then
        sBigMessage = "This definition needs to be made hidden from some user groups as the selected " & asHiddenBySelfToGroupParameters(1, 1) & " '" & asHiddenBySelfToGroupParameters(2, 1) & "' is hidden from some user groups."
      ElseIf mblnDefinitionCreator Then
        If fMakingAndForcingHidden Then
          If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
            sBigMessage = "This definition will now be made hidden from some user groups as the selected " & asHiddenBySelfToGroupParameters(1, 1) & " '" & asHiddenBySelfToGroupParameters(2, 1) & "' is hidden from some user groups."
          End If
        ElseIf (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed for some user groups as the selected " & asHiddenBySelfToGroupParameters(1, 1) & " '" & asHiddenBySelfToGroupParameters(2, 1) & "' is hidden from some user groups."
        End If
      ElseIf fMakingAndForcingHidden Then
        sBigMessage = "The selected " & asHiddenBySelfToGroupParameters(1, 1) & " '" & asHiddenBySelfToGroupParameters(2, 1) & "' will be removed from this definition as it is hidden to some user groups and you do not have permission to make this definition hidden."
      End If
    ElseIf UBound(asHiddenBySelfToGroupParameters, 2) > 1 Then
      If mblnReadOnly Then
        sBigMessage = "This definition needs to be made hidden from some user groups as the following jobs are hidden from some user groups :" & vbCrLf
      ElseIf mblnDefinitionCreator Then
        If fMakingAndForcingHidden Then
          If (Not mblnFromCopy) Or (Not fOnlyFatalMessages) Then
            sBigMessage = "This definition will now be made hidden from some user groups as the following jobs are hidden from some user groups :" & vbCrLf
          End If
        ElseIf (Not fOnlyFatalMessages) Then
          sBigMessage = "The definition access cannot be changed for some user groups as the following jobs are hidden from some user groups :" & vbCrLf
        End If
      ElseIf fMakingAndForcingHidden Then
        sBigMessage = "The following jobs will be removed from this definition as they are hidden to some user groups and you do not have permission to make this definition hidden :" & vbCrLf
      End If

      If Len(sBigMessage) > 0 Then
        For iLoop = 1 To UBound(asHiddenBySelfToGroupParameters, 2)
          sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenBySelfToGroupParameters(1, iLoop) & " '" & asHiddenBySelfToGroupParameters(2, iLoop) & "'"
        Next iLoop
      End If
    End If
  End If

  If UBound(asDeletedParameters, 2) = 1 Then
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected " & asDeletedParameters(1, 1) & " '" & asDeletedParameters(2, 1) & "' has been deleted."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected " & asDeletedParameters(1, 1) & " '" & asDeletedParameters(2, 1) & "' will be removed from this definition as it has been deleted."
    End If
  ElseIf UBound(asDeletedParameters, 2) > 1 Then
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been deleted :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been deleted :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asDeletedParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asDeletedParameters(1, iLoop) & " '" & asDeletedParameters(2, iLoop) & "'"
    Next iLoop
  End If

  If UBound(asHiddenByOtherParameters, 2) = 1 Then
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected " & asHiddenByOtherParameters(1, 1) & " '" & asHiddenByOtherParameters(2, 1) & "' has been made hidden by another user."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected " & asHiddenByOtherParameters(1, 1) & " '" & asHiddenByOtherParameters(2, 1) & "' will be removed from this definition as it has been made hidden by another user."
    End If
  ElseIf UBound(asHiddenByOtherParameters, 2) > 1 Then
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters have been made hidden by another user :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they have been made hidden by another user :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asHiddenByOtherParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asHiddenByOtherParameters(1, iLoop) & " '" & asHiddenByOtherParameters(2, iLoop) & "'"
    Next iLoop
  End If

  If UBound(asInvalidParameters, 2) = 1 Then
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected " & asInvalidParameters(1, 1) & " '" & asInvalidParameters(2, 1) & "' is invalid."
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The selected " & asInvalidParameters(1, 1) & " '" & asInvalidParameters(2, 1) & "' will be removed from this definition as it is invalid."
    End If
  ElseIf UBound(asInvalidParameters, 2) > 1 Then
    If mblnReadOnly Then
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition is currently invalid as the following parameters are invalid :" & vbCrLf
    Else
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "The following parameters will be removed from this definition as they are invalid :" & vbCrLf
    End If

    For iLoop = 1 To UBound(asInvalidParameters, 2)
      sBigMessage = sBigMessage & vbCrLf & vbTab & asInvalidParameters(1, iLoop) & " '" & asInvalidParameters(2, iLoop) & "'"
    Next iLoop
  End If

  If fNeedToForceHidden Then
    SetAllAccess ACCESS_HIDDEN, True
  Else
    sUnhiddenGroups = ""
    
    UI.LockWindow grdAccess.hWnd
    
    With grdAccess
      For iLoop = 1 To (.Rows - 1)
        varBookmark = .AddItemBookmark(iLoop)
        .Bookmark = varBookmark
        
        For iLoop2 = 1 To UBound(asGroups, 2)
          If UCase(asGroups(1, iLoop2)) = UCase(.Columns("GroupName").CellText(varBookmark)) Then
            If (.Columns("SysSecMgr").CellText(varBookmark) <> "1") Then
              
              If (asGroups(2, iLoop2) = ACCESS_HIDDEN) And _
                (asGroups(3, iLoop2) <> ACCESS_HIDDEN) And _
                (asGroups(4, iLoop2) = "1") Then
                ' Definition was HIDDEN from the user group, but no longer needs to be.
                sUnhiddenGroups = sUnhiddenGroups & _
                  IIf(Len(sUnhiddenGroups) > 0, vbCrLf, "") & _
                  vbTab & asGroups(1, iLoop2)
                
                .Columns("ForcedAccess").Text = "0"
              End If
              
              If (asGroups(3, iLoop2) = ACCESS_HIDDEN) Then
                .Columns("Access").Text = AccessDescription(ACCESS_HIDDEN)
                'JPD The next two lines look very dodgy. Why sert the Text property to
                ' an empty string and then assign it another value ? Well, I was finding
                ' that sometimes setting the Text to "1" didn't actually assign the value.
                ' Assigning "" and then "1" seemed to work fine.
                .Columns("ForcedAccess").Text = ""
                .Columns("ForcedAccess").Text = "1"
              End If
            End If
            
            Exit For
          End If
        Next iLoop2
      Next iLoop
      
      .MoveFirst
    End With
    
    UI.UnlockWindow
    
    If Len(sUnhiddenGroups) > 0 And (Not fOnlyFatalMessages) Then
      ' Inform the user if the definition no longer needs to be hidden.
      sBigMessage = sBigMessage & IIf(Len(sBigMessage) > 0, vbCrLf & vbCrLf, "") & _
        "This definition no longer has to be hidden from the following user groups :" & vbCrLf & vbCrLf & sUnhiddenGroups
    End If
  End If
  
  If Len(sBigMessage) > 0 Then
    COAMsgBox sBigMessage, vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
  End If
  ForceDefinitionToBeHiddenIfNeeded = (Len(sBigMessage) = 0)
     
  RefreshRoleToPromptCombo psRoleToPrompt
  
  SetButtonState
  CheckIfScrollBarRequired
      
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
Private Function OnlyPauseJobsDefined() As Boolean
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  ' Return true if only pause jobs exist
  
  grdColumns.MoveFirst
  For pintLoop = 0 To grdColumns.Rows - 1
    pvarbookmark = grdColumns.GetBookmark(pintLoop)
    If grdColumns.Columns("Job Type").CellText(pvarbookmark) <> "-- Pause --" Then
      OnlyPauseJobsDefined = False
      Exit Function
    End If
  Next pintLoop

  OnlyPauseJobsDefined = True

End Function

Private Sub optOutputFormat_Click(Index As Integer)
  If Not mblnLoading Then
    txtTitlePage = vbNullString
    cmdTitlePageClear.Enabled = False
  End If
  mobjOutputDef.FormatClick Index
  Changed = True
End Sub
Private Sub spnFrequency_Change()
  Changed = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  If mblnReadOnly Then Exit Sub

  If SSTab1.Tab = 1 Then
    RefreshColumnsGrid
  End If

  fraInfo.Enabled = (SSTab1.Tab = 0)
  fraScheduling.Enabled = (SSTab1.Tab = 0)
  fraJobs.Enabled = (SSTab1.Tab = 1)
  fraEMailNotify.Enabled = (SSTab1.Tab = 1)
  fraOutput.Enabled = (SSTab1.Tab = 2)
  fraDest.Enabled = (SSTab1.Tab = 2)
  fraOptions.Enabled = (SSTab1.Tab = 2)
  
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

Private Sub txtDesc_Change()
  Changed = True
End Sub

Private Sub txtDesc_GotFocus()
  With txtDesc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmailNotifyGroup_Change(Index As Integer)
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

Private Function CheckUniqueName(sName As String, lngCurrentID As Long, blnIsReport As Boolean) As Boolean
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSysBatchJobName " & _
         "WHERE UPPER(Name) = '" & UCase(Replace(sName, "'", "''")) & "' " & _
         "AND ID <> " & lngCurrentID & _
         " AND IsBatch = " & IIf(blnIsReport, "0", "1")
  
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckUniqueName = True
  Else
    CheckUniqueName = False
  End If
  
  Set rsTemp = Nothing

End Function


Private Sub cmdMoveDown_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  With grdColumns
  
    intSourceRow = .AddItemRowIndex(.Bookmark)
    strSourceRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text
    
    intDestinationRow = intSourceRow + 1
    .MoveNext
    strDestinationRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text
    
    .RemoveItem intDestinationRow
    .RemoveItem intSourceRow
    
    .AddItem strDestinationRow, intSourceRow
    .AddItem strSourceRow, intDestinationRow
    
    .SelBookmarks.RemoveAll
    .MoveNext
    '
    .Bookmark = .AddItemBookmark(intDestinationRow)
    .SelBookmarks.Add .AddItemBookmark(intDestinationRow)
  
  End With
  
  SetButtonState
  Changed = True
  
End Sub

Private Sub cmdMoveUp_Click()

  Dim intSourceRow As Integer
  Dim strSourceRow As String
  Dim intDestinationRow As Integer
  Dim strDestinationRow As String
  
  With grdColumns
  
    intSourceRow = .AddItemRowIndex(.Bookmark)
    strSourceRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text
    
    intDestinationRow = intSourceRow - 1
    .MovePrevious
    strDestinationRow = .Columns(0).Text & vbTab & .Columns(1).Text & vbTab & .Columns(2).Text & vbTab & .Columns(3).Text & vbTab & .Columns(4).Text
    
    .AddItem strSourceRow, intDestinationRow
    .RemoveItem intSourceRow + 1
    
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .AddItemBookmark(intDestinationRow)
    .MovePrevious
    .MovePrevious
  
  End With
  
  SetButtonState
  Changed = True
  
End Sub

Public Sub PrintDef(lBatchJobID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsColumns As Recordset
  Dim sSQL As String
  Dim sTemp As String
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim iUtilityType As UtilityType
  
  mlngBatchJobID = lBatchJobID
  
  Set rsTemp = datGeneral.GetRecords("SELECT ASRSysBatchJobName.*, " & _
                                      "CONVERT(integer, ASRSysBatchJobName.TimeStamp) AS intTimeStamp " & _
                                      "FROM ASRSysBatchJobName WHERE ID = " & mlngBatchJobID)
                                        
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Print Definition"
    Set rsTemp = Nothing
    Exit Sub
  End If

  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    iUtilityType = IIf(IsReportPack, utlReportPack, utlBatchJob)
  
    With objPrintDef
      If .PrintStart(False) Then
        .PrintHeader "" & IIf(gblnReportPackMode, "Report Pack : ", "Batch Job : ") & rsTemp!Name
        .PrintNormal "Category : " & GetObjectCategory(iUtilityType, mlngBatchJobID)
        .PrintNormal "Description : " & IIf(rsTemp!Description <> vbNullString, rsTemp!Description, "N/A")
        .PrintNormal
        .PrintNormal "Owner : " & rsTemp!userName
        .PrintNormal
        .PrintNormal "Scheduled : " & IIf(rsTemp!scheduled = True, "Yes", "No")
        
        If rsTemp!scheduled = True Then
        
          sTemp = "Run Every : " & rsTemp!Frequency
          
          Select Case rsTemp!Period
            Case "D": sTemp = sTemp & " Day(s)"
            Case "W": sTemp = sTemp & " Week(s)"
            Case "M": sTemp = sTemp & " Month(s)"
            Case "Y": sTemp = sTemp & " Year(s)"
          End Select
          
          .PrintNormal sTemp
          sTemp = vbNullString
          
          .PrintNormal
          .PrintNormal "Start Date : " & Format(rsTemp!StartDate, DateFormat)
          .PrintNormal "End Date : " & IIf(IsDate(rsTemp!EndDate) And rsTemp!EndDate <> "00:00:00", Format(rsTemp!EndDate, DateFormat), "")
          .PrintNormal
          .PrintNormal "Run Indefinitely : " & IIf(rsTemp!Indefinitely = True, "Yes", "No")
          .PrintNormal "Include Weekends : " & IIf(rsTemp!Weekends = True, "Yes", "No")
          .PrintNormal "Skip Missed Days : " & IIf(rsTemp!RunOnce = True, "Yes", "No")
          .PrintNormal
          .PrintNormal "User Group : " & rsTemp!RoleToPrompt
         End If
          
        ' Access section
        PopulateAccessGrid
        .PrintTitle "Access"
        For iLoop = 1 To (grdAccess.Rows - 1)
          varBookmark = grdAccess.AddItemBookmark(iLoop)
          .PrintNormal grdAccess.Columns("GroupName").CellValue(varBookmark) & " : " & grdAccess.Columns("Access").CellValue(varBookmark)
        Next iLoop
        
        If gblnReportPackMode Then
          ' Only include Output Options for Report Packs for the mo
          .PrintTitle "Output Options"
          .PrintNormal "Report Pack Template : " & rsTemp!OutputTitlePage
          .PrintNormal "Report Pack Title : " & rsTemp!OutputReportPackTitle
          .PrintNormal " "
          .PrintNormal "Override Filter : " & rsTemp!OutputOverrideFilter
          .PrintNormal "Create Table Of Contents : " & IIf(rsTemp!OutputTOC = True, "Yes", "No")
          .PrintNormal "Create Cover Sheet(s) : " & IIf(rsTemp!OutputCoverSheet = True, "Yes", "No")
          .PrintNormal "Retain Pivot/chart : " & IIf(rsTemp!OutputRetainPivotOrChart = True, "Yes", "No")
          '.PrintNormal "Retain Chart : " & IIf(rsTemp!OutputRetainCharts = True, "Yes", "No")
          
          .PrintNormal " "
          
          Select Case rsTemp!OutputFormat
            Case fmtExcelWorksheet
              .PrintNormal "Output Format : Excel Worksheet"
            Case fmtWordDoc
              .PrintNormal "Output Format : Word Document"
            Case fmtHTML
              .PrintNormal "Output Format : HTML"
          End Select
          
          .PrintNormal " "
          
          If rsTemp!OutputPrinter Then
            .PrintNormal "Output Destination : Send to printer"
            .PrintNormal "Printer Location : " & rsTemp!OutputPrinterName
            .PrintNormal "File Name : " & rsTemp!OutputFilename
            .PrintNormal " "
          End If
          
          If rsTemp!OutputSave Then
            .PrintNormal "Output Destination : Save to file"
            
            Select Case rsTemp!OutputSaveExisting
              Case 0: .PrintNormal "If Existing File : Overwrite"
              Case 1: .PrintNormal "If Existing File : Do not overwrite"
              Case 2: .PrintNormal "If Existing File : Add sequential number to name"
              Case 3: .PrintNormal "If Existing File : Append to file"
            End Select
            .PrintNormal " "
          End If
          
          If rsTemp!OutputEmail Then
            txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
            .PrintNormal "Output Destination : Send to email"
            .PrintNormal "Email Group : " & IIf(txtEmailGroup.Text = "", "N/A", txtEmailGroup.Text)
            .PrintNormal "Email Subject : " & rsTemp!OutputEmailSubject
            .PrintNormal "Email Attach As : " & rsTemp!OutputEmailAttachAs
          End If
        End If
      
      ' Now do the individual jobs
      If IsReportPack Then
        .PrintTitle "Reports"
      Else
        .PrintTitle "Jobs"
      End If
      
      Set rsColumns = datGeneral.GetRecords("SELECT * " & _
                                            "FROM ASRSysBatchJobDetails WHERE BatchJobNameID = " & mlngBatchJobID & _
                                            "ORDER BY JobOrder")
      
      ' Change Titles according to Batch or Reports
      If IsReportPack Then
        .PrintBold "Report Type" & vbTab & "Report Name"
      Else
        .PrintBold "Job Type" & vbTab & "Job Name"
      End If
    
        Do While Not rsColumns.EOF
          If Not JobTypeRequiresDef(rsColumns!JobType) Then
            .PrintNonBold rsColumns!JobType & vbTab & _
                          rsColumns!Parameter
          Else
            .PrintNonBold rsColumns!JobType & vbTab & _
                          GetJobName(rsColumns!JobType, rsColumns!JobID)
          End If
          rsColumns.MoveNext
        Loop
        
        .PrintEnd
        
        If gblnReportPackMode Then
          .PrintConfirm "Report Packs : " & rsTemp!Name, "Report Pack Definition"
        Else
          .PrintConfirm "Batch Job : " & rsTemp!Name, "Batch Job Definition"
        End If
      End If
    End With
  End If
  
  Set rsTemp = Nothing
  Set rsColumns = Nothing
  Changed = False

Exit Sub

LocalErr:
  COAMsgBox "Printing " & IIf(gblnReportPackMode, "Report Pack : ", "Batch Job : ") & "definition failed" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Print Definition"

End Sub

Private Function CheckIfScrollBarRequired()
  With grdColumns

    If .Rows > 11 Then
      .ScrollBars = ssScrollBarsVertical
      .Columns("Parameter").Width = 2710
    Else
      .ScrollBars = ssScrollBarsNone
      .Columns("Parameter").Width = 2935
    End If

  End With

End Function


Private Function JobTypeRequiresDef(strJobType As String) As Boolean

  JobTypeRequiresDef = _
      (strJobType <> "-- Pause --" And _
       strJobType <> "Absence Breakdown" And _
       strJobType <> "Bradford Factor" And _
       strJobType <> "Stability Index Report" And _
       strJobType <> "Turnover Report")

End Function


Private Sub txtOverrideFilter_Change()
 Changed = True
End Sub

Private Sub txtReportPackTitle_Change()
 Changed = True
End Sub

Private Sub txtTitlePage_Change()
 Changed = True
End Sub

Private Sub cboCategory_Click()
  Changed = True
End Sub



