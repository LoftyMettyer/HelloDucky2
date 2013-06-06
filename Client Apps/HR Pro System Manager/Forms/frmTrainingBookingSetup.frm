VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmTrainingBookingSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Training Booking Module"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
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
   HelpContextID   =   5089
   Icon            =   "frmTrainingBookingSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6195
      TabIndex        =   15
      Top             =   6375
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4935
      TabIndex        =   14
      Top             =   6375
      Width           =   1200
   End
   Begin TabDlg.SSTab ssTabStrip 
      Height          =   6075
      HelpContextID   =   5089
      Left            =   150
      TabIndex        =   16
      Top             =   150
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Cour&ses"
      TabPicture(0)   =   "frmTrainingBookingSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCourses"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Pre-requisites"
      TabPicture(1)   =   "frmTrainingBookingSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPreRequisites"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "De&legates"
      TabPicture(2)   =   "frmTrainingBookingSetup.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDelegates"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraUnavailability"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Waiting List / &Bookings"
      TabPicture(3)   =   "frmTrainingBookingSetup.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraWaitingList"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fraTrainingBookings"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&Related Columns"
      TabPicture(4)   =   "frmTrainingBookingSetup.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraRelatedColumns"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame fraRelatedColumns 
         Caption         =   "Related Columns :"
         Enabled         =   0   'False
         Height          =   5400
         Left            =   -74850
         TabIndex        =   77
         Top             =   500
         Width           =   6945
         Begin VB.CommandButton cmdAddRelatedColumn 
            Caption         =   "&Add"
            Height          =   400
            Left            =   4185
            TabIndex        =   42
            Top             =   4850
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteRelatedColumn 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   5595
            TabIndex        =   43
            Top             =   4850
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid ssgrdRelatedColumns 
            Height          =   4380
            Left            =   195
            TabIndex        =   41
            Top             =   300
            Width           =   6585
            ScrollBars      =   0
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   2
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
            Columns.Count   =   2
            Columns(0).Width=   4921
            Columns(0).Caption=   "Training Booking"
            Columns(0).Name =   "TrainingBooking"
            Columns(0).CaptionAlignment=   2
            Columns(0).AllowSizing=   0   'False
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(0).Locked=   -1  'True
            Columns(0).Style=   3
            Columns(1).Width=   4921
            Columns(1).Caption=   "Waiting List"
            Columns(1).Name =   "WaitingList"
            Columns(1).CaptionAlignment=   2
            Columns(1).AllowSizing=   0   'False
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(1).Locked=   -1  'True
            Columns(1).Style=   3
            TabNavigation   =   1
            _ExtentX        =   11615
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
      End
      Begin VB.Frame fraPreRequisites 
         Caption         =   "Pre-requisite Course History :"
         Enabled         =   0   'False
         Height          =   2600
         Left            =   -74850
         TabIndex        =   70
         Top             =   500
         Width           =   6900
         Begin VB.ComboBox cboPreRequisiteCourseTitle 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   700
            Width           =   2800
         End
         Begin VB.ComboBox cboPreRequisiteCourseGroup 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1100
            Width           =   2800
         End
         Begin VB.ComboBox cboPreRequisiteTable 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   300
            Width           =   2800
         End
         Begin VB.ComboBox cboPreRequisiteFailureNotification 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1500
            Width           =   2800
         End
         Begin VB.OptionButton optPreRequisiteDefaultFailureNotification 
            Caption         =   "&Error"
            Height          =   315
            Index           =   0
            Left            =   3855
            TabIndex        =   21
            Top             =   1915
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optPreRequisiteDefaultFailureNotification 
            Caption         =   "&Warning"
            Height          =   315
            Index           =   1
            Left            =   4950
            TabIndex        =   22
            Top             =   1915
            Width           =   1080
         End
         Begin VB.Label lblPreRequisiteCourseTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Course Title Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   75
            Top             =   750
            Width           =   1935
         End
         Begin VB.Label lblPreRequisiteFailureNotification 
            BackStyle       =   0  'Transparent
            Caption         =   "Failure Notification Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   74
            Top             =   1545
            Width           =   2550
         End
         Begin VB.Label lblPreRequisiteCourseGroup 
            BackStyle       =   0  'Transparent
            Caption         =   "Grouping Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   73
            Top             =   1155
            Width           =   1725
         End
         Begin VB.Label lblPreRequisiteTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre-requisite Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   72
            Top             =   360
            Width           =   1860
         End
         Begin VB.Label lblPreRequisiteDefaultFailureNotification 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default pre-requisite failure notification :"
            Height          =   390
            Left            =   195
            TabIndex        =   71
            Top             =   1950
            Width           =   3540
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraTrainingBookings 
         Caption         =   "Training Bookings History :"
         Enabled         =   0   'False
         Height          =   1980
         Left            =   -74850
         TabIndex        =   67
         Top             =   2280
         Width           =   6945
         Begin VB.ComboBox cboTrainingBookingsCancellationDate 
            Height          =   315
            Left            =   3765
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1100
            Width           =   2800
         End
         Begin VB.OptionButton optTrainingBookingsOverlappedNotification 
            Caption         =   "&Warning"
            Height          =   195
            Index           =   1
            Left            =   4860
            TabIndex        =   40
            Top             =   1560
            Width           =   1170
         End
         Begin VB.OptionButton optTrainingBookingsOverlappedNotification 
            Caption         =   "&Error"
            Height          =   195
            Index           =   0
            Left            =   3765
            TabIndex        =   39
            Top             =   1560
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.ComboBox cboTrainingBookingsTable 
            Height          =   315
            Left            =   3765
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   300
            Width           =   2800
         End
         Begin VB.ComboBox cboTrainingBookingsStatus 
            Height          =   315
            Left            =   3765
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   700
            Width           =   2800
         End
         Begin VB.Label lblTrainingBookingsCancellationDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancellation Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   78
            Top             =   1155
            Width           =   2715
         End
         Begin VB.Label lblTrainingBookingsOverlappedNotification 
            BackStyle       =   0  'Transparent
            Caption         =   "Overlapped Booking Notification :"
            Height          =   195
            Left            =   195
            TabIndex        =   76
            Top             =   1560
            Width           =   2985
         End
         Begin VB.Label lblTrainingBookingsTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Training Bookings Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   69
            Top             =   360
            Width           =   2505
         End
         Begin VB.Label lblTrainingBookingsStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "Status Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   68
            Top             =   765
            Width           =   1860
         End
      End
      Begin VB.Frame fraWaitingList 
         Caption         =   "Waiting List History :"
         Enabled         =   0   'False
         Height          =   1680
         Left            =   -74850
         TabIndex        =   64
         Top             =   500
         Width           =   6945
         Begin VB.ComboBox cboWaitingListOrderOveride 
            Height          =   315
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1125
            Width           =   2805
         End
         Begin VB.ComboBox cboWaitingListTable 
            Height          =   315
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   300
            Width           =   2805
         End
         Begin VB.ComboBox cboWaitingListCourseTitle 
            Height          =   315
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   712
            Width           =   2805
         End
         Begin VB.Label lblWaitingListOrderOveride 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Order Override Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   80
            Top             =   1185
            Width           =   2715
         End
         Begin VB.Label lblWaitingListTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Waiting List Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   66
            Top             =   360
            Width           =   1860
         End
         Begin VB.Label lblWaitingListCourseTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Course Title Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   65
            Top             =   765
            Width           =   2025
         End
      End
      Begin VB.Frame fraUnavailability 
         Caption         =   "Unavailability History :"
         Enabled         =   0   'False
         Height          =   2505
         Left            =   -74850
         TabIndex        =   58
         Top             =   3000
         Width           =   6945
         Begin VB.ComboBox cboUnavailabilityTable 
            Height          =   315
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   300
            Width           =   2800
         End
         Begin VB.ComboBox cboUnavailabilityFromDate 
            Height          =   315
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   700
            Width           =   2800
         End
         Begin VB.ComboBox cboUnavailabilityToDate 
            Height          =   315
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1100
            Width           =   2800
         End
         Begin VB.ComboBox cboUnavailabilityNotification 
            Height          =   315
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1500
            Width           =   2800
         End
         Begin VB.OptionButton optUnavailabilityDefaultNotification 
            Caption         =   "&Error"
            Height          =   315
            Index           =   0
            Left            =   3810
            TabIndex        =   31
            Top             =   1915
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optUnavailabilityDefaultNotification 
            Caption         =   "&Warning"
            Height          =   315
            Index           =   1
            Left            =   4905
            TabIndex        =   32
            Top             =   1915
            Width           =   1260
         End
         Begin VB.Label lblUnavailabilityTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Unavailability Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   63
            Top             =   360
            Width           =   2205
         End
         Begin VB.Label lblUnavailabilityFromDate 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   62
            Top             =   765
            Width           =   2145
         End
         Begin VB.Label lblUnavailabilityToDate 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   61
            Top             =   1155
            Width           =   1965
         End
         Begin VB.Label lblUnavailabilityNotification 
            BackStyle       =   0  'Transparent
            Caption         =   "Unavailablility Notification Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   60
            Top             =   1560
            Width           =   3225
         End
         Begin VB.Label lblUnavailabilityDefaultNotification 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Unavailability Notification :"
            Height          =   195
            Left            =   195
            TabIndex        =   59
            Top             =   1965
            Width           =   3195
         End
      End
      Begin VB.Frame fraDelegates 
         Caption         =   "Employee/Delegate Records :"
         Enabled         =   0   'False
         Height          =   2205
         Left            =   -74850
         TabIndex        =   55
         Top             =   500
         Width           =   6945
         Begin VB.ComboBox cboEmployeeTable 
            Height          =   315
            Left            =   3765
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   300
            Width           =   2800
         End
         Begin VB.ComboBox cboDefaultView 
            Height          =   315
            Left            =   3765
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1500
            Width           =   2800
         End
         Begin VB.CommandButton cmdEmployeeTableOrder 
            Caption         =   "..."
            Height          =   315
            Left            =   6255
            TabIndex        =   25
            Top             =   900
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txtEmployeeTableOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3765
            TabIndex        =   24
            Top             =   900
            Width           =   2485
         End
         Begin VB.Label Label1 
            Caption         =   "Personnel Default View :"
            Height          =   315
            Left            =   195
            TabIndex        =   79
            Top             =   1560
            Width           =   2760
         End
         Begin VB.Label lblEmployeeTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee/Delegate Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   57
            Top             =   360
            Width           =   2700
         End
         Begin VB.Label lblEmployeeOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order - used when creating bookings from the Waiting List when in Course Records :"
            Height          =   585
            Left            =   195
            TabIndex        =   56
            Top             =   705
            Width           =   3345
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraCourses 
         Caption         =   "Course Records :"
         Height          =   5400
         Left            =   150
         TabIndex        =   44
         Top             =   500
         Width           =   6945
         Begin VB.CommandButton cmdCourseTableOrder 
            Caption         =   "..."
            Height          =   315
            Left            =   6390
            TabIndex        =   13
            Top             =   4650
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.OptionButton optCourseOverbookingNotification 
            Caption         =   "&Warning"
            Height          =   315
            Index           =   1
            Left            =   4950
            TabIndex        =   11
            Top             =   4215
            Width           =   1215
         End
         Begin VB.OptionButton optCourseOverbookingNotification 
            Caption         =   "&Error"
            Height          =   315
            Index           =   0
            Left            =   3855
            TabIndex        =   10
            Top             =   4215
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.CheckBox chkCourseTransferProvisionalBookingsWhenCancelling 
            Caption         =   "&Transfer Provisional Bookings when cancelling courses"
            Height          =   315
            Left            =   200
            TabIndex        =   8
            Top             =   3500
            Width           =   5475
         End
         Begin VB.CheckBox chkCourseIncludeProvisionalBookingsInNumberBookedColumn 
            Caption         =   "&Include Provisional Bookings in Number Booked Column"
            Height          =   315
            Left            =   200
            TabIndex        =   9
            Top             =   3800
            Width           =   5475
         End
         Begin VB.TextBox txtCourseTableOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3900
            TabIndex        =   12
            Top             =   4650
            Width           =   2485
         End
         Begin VB.ComboBox cboCourseEndDate 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1500
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseTable 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   300
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseCancellationDate 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2700
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseNumberBooked 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1900
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseStartDate 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1100
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseTitle 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   700
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseCancelledBy 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3100
            Width           =   2800
         End
         Begin VB.ComboBox cboCourseMaxNumberOfDelegates 
            Height          =   315
            Left            =   3855
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2300
            Width           =   2800
         End
         Begin VB.Label lblOverbookingNotification 
            BackStyle       =   0  'Transparent
            Caption         =   "Overbooking Notification :"
            Height          =   195
            Left            =   195
            TabIndex        =   54
            Top             =   4260
            Width           =   2325
         End
         Begin VB.Label lblCourseTableOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order - used when creating bookings from the Waiting List when in Employee Records :"
            Height          =   585
            Left            =   195
            TabIndex        =   53
            Top             =   4605
            Width           =   3600
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblCourseEndDate 
            BackStyle       =   0  'Transparent
            Caption         =   "End Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   52
            Top             =   1560
            Width           =   1650
         End
         Begin VB.Label lblCourseTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Course Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   51
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblCourseMaxNumberOfDelegates 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Number Of Delegates Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   50
            Top             =   2355
            Width           =   3210
         End
         Begin VB.Label lblCourseCancelledBy 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelled By Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   49
            Top             =   3165
            Width           =   2085
         End
         Begin VB.Label lblCourseCancellationDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancellation Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   48
            Top             =   2760
            Width           =   2475
         End
         Begin VB.Label lblCourseNumberBooked 
            BackStyle       =   0  'Transparent
            Caption         =   "Number Booked Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   47
            Top             =   1965
            Width           =   2340
         End
         Begin VB.Label lblCourseStartDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   46
            Top             =   1155
            Width           =   1740
         End
         Begin VB.Label lblCourseTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Title Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   45
            Top             =   765
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frmTrainingBookingSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare events
Event UnLoad()

'
' CONSTANTS
'
' Page number constants.
Private Const giPAGE_COURSES = 0
Private Const giPAGE_PREREQUISITES = 1
Private Const giPAGE_DELEGATES = 2
Private Const giPAGE_WAITINGLISTBOOKINGS = 3
Private Const giPAGE_RELATEDCOLUMNS = 4

'
' VARIABLES
'
' Course Records variables.
Private mvar_lngCourseTableID As Long
Private mvar_lngCourseTitleID As Long
Private mvar_lngCourseStartDateID As Long
Private mvar_lngCourseEndDateID As Long
Private mvar_lngCourseNumberBookedID As Long
Private mvar_lngCourseMaxNumberID As Long
Private mvar_lngCourseCancelDateID As Long
Private mvar_lngCourseCancelledByID As Long
Private mvar_fCourseTransferProvisionals As Boolean
Private mvar_fCourseIncludeProvisionals As Boolean
Private mvar_iCourseOverbookingNotification As Integer
Private mvar_lngCourseOrderID As Long
' Pre-requisite variables.
Private mvar_lngPreReqTableID As Long
Private mvar_lngPreReqCourseTitleID As Long
Private mvar_lngPreReqGroupingID As Long
Private mvar_lngPreReqFailureNotificationID As Long
Private mvar_iPreReqDfltFailureNotification As Integer
' Employee Records variables.
Private mvar_lngEmployeeTableID  As Long
Private mvar_lngEmployeeOrderID As Long
Private mvar_lngBulkBookingDefaultViewID As Long 'NHRD01052003 Fault 4687
Private mlng_ListIndex As Long 'NHRD01052003 Fault 4687
' Unavailability variables.
Private mvar_lngUnavailTableID As Long
Private mvar_lngUnavailFromDateID As Long
Private mvar_lngUnavailToDateID As Long
Private mvar_lngUnavailFailureNotificationID As Long
Private mvar_iUnavailDfltFailureNotification As Integer
' Waiting List variables.
Private mvar_lngWaitListTableID As Long
Private mvar_lngWaitListCourseTitleID As Long
Private mvar_lngWaitListOrderOverideColumn As Long
' Training Booking variables.
Private mvar_lngTrainBookTableID As Long
Private mvar_lngTrainBookStatusID As Long
Private mvar_lngTrainBookCancelDateID As Long
Private mvar_iTrainBookOverlapNotification As Integer
' Related Column variables.
Private mvar_alngRelatedColumns() As Long

Private mblnReadOnly As Boolean

Private mstrViewName As String
Private mbLoading As Boolean
Private mfChanged As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOK.Enabled = True
End Property

Private Sub cboCourseCancellationDate_Click()
  ' Save the selected ID to a local variable.
  With cboCourseCancellationDate
    mvar_lngCourseCancelDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboCourseCancelledBy_Click()
  ' Save the selected ID to a local variable.
  With cboCourseCancelledBy
    mvar_lngCourseCancelledByID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboCourseEndDate_Click()
  ' Save the selected ID to a local variable.
  With cboCourseEndDate
    mvar_lngCourseEndDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboCourseMaxNumberOfDelegates_Click()
  ' Save the selected ID to a local variable.
  With cboCourseMaxNumberOfDelegates
    mvar_lngCourseMaxNumberID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboCourseNumberBooked_Click()
  ' Save the selected ID to a local variable.
  With cboCourseNumberBooked
    mvar_lngCourseNumberBookedID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboCourseStartDate_Click()
  ' Save the selected ID to a local variable.
  With cboCourseStartDate
    mvar_lngCourseStartDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboCourseTable_Click()
  ' Save the selected ID to a local variable.
  With cboCourseTable
    mvar_lngCourseTableID = .ItemData(.ListIndex)
  End With
  
  ' Refresh the Course column and Course History table controls.
  RefreshCourseColumnControls
  RefreshCoursesHistoryTableCombos

  RefreshAllControls fraCourses
  Changed = True
End Sub


Private Sub cboCourseTitle_Click()
  ' Save the selected ID to a local variable.
  With cboCourseTitle
    mvar_lngCourseTitleID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboDefaultView_Click()
  ' Save the selected ID to a local variable.
  With cboDefaultView
    mvar_lngBulkBookingDefaultViewID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub
Private Sub cboEmployeeTable_Click()
  ' Save the selected ID to a local variable.
  With cboEmployeeTable
    mvar_lngEmployeeTableID = .ItemData(.ListIndex)
  End With
 
  ' Refresh the Employee column and History table controls.
  RefreshEmployeeColumnControls
  RefreshEmployeesHistoryTableCombos
  RefreshAllControls fraDelegates
  Changed = True
End Sub

Private Sub cboPreRequisiteCourseGroup_Click()
  ' Save the selected ID to a local variable.
  With cboPreRequisiteCourseGroup
    mvar_lngPreReqGroupingID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboPreRequisiteCourseTitle_Click()
  ' Save the selected ID to a local variable.
  With cboPreRequisiteCourseTitle
    mvar_lngPreReqCourseTitleID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboPreRequisiteFailureNotification_Click()
  ' Save the selected ID to a local variable.
  With cboPreRequisiteFailureNotification
    mvar_lngPreReqFailureNotificationID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboPreRequisiteTable_Click()
  ' Save the selected ID to a local variable.
  With cboPreRequisiteTable
    mvar_lngPreReqTableID = .ItemData(.ListIndex)
  End With
 
  ' Refresh the form display.
  ' We need to do this here as VB screws up when populating combos (below)
  ' that are hidden by a dropped-down combo.
  DoEvents

  ' Refresh the Pre-requisite column combos.
  RefreshPreRequisiteColumnControls
  Changed = True
End Sub




Private Sub cboTrainingBookingsCancellationDate_Click()
  ' Save the selected ID to a local variable.
  With cboTrainingBookingsCancellationDate
    mvar_lngTrainBookCancelDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboTrainingBookingsStatus_Click()
  ' Save the selected ID to a local variable.
  With cboTrainingBookingsStatus
    mvar_lngTrainBookStatusID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboTrainingBookingsTable_Click()
  ' Save the selected ID to a local variable.
  With cboTrainingBookingsTable
    mvar_lngTrainBookTableID = .ItemData(.ListIndex)
  End With
  
  ' Refresh the Training Bookings column combos.
  RefreshTrainingBookingsColumnControls
  Changed = True
End Sub


Private Sub cboUnavailabilityFromDate_Click()
  ' Save the selected ID to a local variable.
  With cboUnavailabilityFromDate
    mvar_lngUnavailFromDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboUnavailabilityNotification_Click()
  ' Save the selected ID to a local variable.
  With cboUnavailabilityNotification
    mvar_lngUnavailFailureNotificationID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboUnavailabilityTable_Click()
  ' Save the selected ID to a local variable.
  With cboUnavailabilityTable
    mvar_lngUnavailTableID = .ItemData(.ListIndex)
  End With
  
  ' Refresh the form display.
  ' We need to do this here as VB screws up when populating combos (below)
  ' that are hidden by a dropped-down combo.
  DoEvents

  ' Refresh the Unavailability column combos.
  RefreshUnavailabilityColumnControls
  Changed = True
End Sub


Private Sub cboUnavailabilityToDate_Click()
  ' Save the selected ID to a local variable.
  With cboUnavailabilityToDate
    mvar_lngUnavailToDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWaitingListCourseTitle_Click()
  
  DoEvents
  
  ' Save the selected ID to a local variable.
  With cboWaitingListCourseTitle
    mvar_lngWaitListCourseTitleID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWaitingListOrderOveride_Click()
  
  DoEvents
  
  ' Save the selected ID to a local variable.
  With cboWaitingListOrderOveride
    mvar_lngWaitListOrderOverideColumn = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWaitingListTable_Click()
  
  DoEvents
  
  ' Save the selected ID to a local variable.
  With cboWaitingListTable
    mvar_lngWaitListTableID = .ItemData(.ListIndex)
  End With
  
  ' Refresh the Waiting List column combos.
  RefreshWaitingListColumnControls
  ' Refresh the Waiting List Override column combo.
  RefreshWaitingListOverrideColumnControls
  
  Changed = True
End Sub

Private Sub chkCourseIncludeProvisionalBookingsInNumberBookedColumn_Click()
  ' Save the selected ID to a local variable.
  mvar_fCourseIncludeProvisionals = (chkCourseIncludeProvisionalBookingsInNumberBookedColumn.value = vbChecked)
  Changed = True
End Sub

Private Sub chkCourseTransferProvisionalBookingsWhenCancelling_Click()
  ' Save the selected ID to a local variable.
  mvar_fCourseTransferProvisionals = (chkCourseTransferProvisionalBookingsWhenCancelling.value = vbChecked)
    Changed = True
End Sub

Private Sub cmdAddRelatedColumn_Click()
  
  ' Add a new column to the related columns grid.
  With ssgrdRelatedColumns
    .AddItem "" & ""
    .MoveLast
    .SetFocus
  
    ' Ensure the 'Delete' command button is enabled.
    'cmdDeleteRelatedColumn.Enabled = True
    cmdDeleteRelatedColumn.Enabled = Not mblnReadOnly
  End With
  
  'Refresh the Related columns grid
  RefreshRelatedColumns
  Changed = True
End Sub

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
'        Screen.MousePointer = vbDefault
'        'Me.MousePointer = vbNormal
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:
  UnLoad Me
End Sub

Private Sub cmdCourseTableOrder_Click()
  ' Display the Order selection form.
  Dim objOrder As Order
  
  ' Instantiate an order object.
  Set objOrder = New Order
  
  With objOrder
    ' Initialize the order object.
    .OrderID = mvar_lngCourseOrderID
    .TableID = mvar_lngCourseTableID
    .OrderType = giORDERTYPE_DYNAMIC

    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      mvar_lngCourseOrderID = .OrderID
    Else
      ' Check in case the original order has been deleted.
      With recOrdEdit
        .Index = "idxID"
        .Seek "=", mvar_lngCourseOrderID

        If .NoMatch Then
          mvar_lngCourseOrderID = 0
        Else
          If !Deleted Then
            mvar_lngCourseOrderID = 0
          End If
        End If
      End With
    End If
  End With
  
  Set objOrder = Nothing
  
  ' Update the controls properties.
  RefreshCourseOrderControls
  Changed = True
End Sub

Private Sub cmdDeleteRelatedColumn_Click()
  ' Remove the current Related Column pairing
  Dim lRow As Long

  If ssgrdRelatedColumns.Rows = 1 Then
    ssgrdRelatedColumns.RemoveAll
  Else
    With ssgrdRelatedColumns

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
 
  ' Enable the 'Delete' command button if we have at least one row.
  'cmdDeleteRelatedColumn.Enabled = (ssgrdRelatedColumns.Rows > 0)
  cmdDeleteRelatedColumn.Enabled = (ssgrdRelatedColumns.Rows > 0 And Not mblnReadOnly)

  'Refresh the Related columns grid
  RefreshRelatedColumns
  Changed = True
End Sub

Private Sub cmdEmployeeTableOrder_Click()
  ' Display the Order selection form.
  Dim objOrder As Order
  
  ' Instantiate an order object.
  Set objOrder = New Order
  
  With objOrder
    ' Initialize the order object.
    .OrderID = mvar_lngEmployeeOrderID
    .TableID = mvar_lngEmployeeTableID
    .OrderType = giORDERTYPE_DYNAMIC

    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      mvar_lngEmployeeOrderID = .OrderID
    Else
      ' Check in case the original order has been deleted.
      With recOrdEdit
        .Index = "idxID"
        .Seek "=", mvar_lngEmployeeOrderID

        If .NoMatch Then
          mvar_lngEmployeeOrderID = 0
        Else
          If !Deleted Then
            mvar_lngEmployeeOrderID = 0
          End If
        End If
      End With
    End If
  End With
  
  Set objOrder = Nothing
  
  ' Update the controls properties.
  RefreshEmployeeOrderControls
  Changed = True
End Sub


Private Sub cmdOK_Click()
  If SaveChanges Then
    'AE20071119 Fault #12607
    Changed = False
    UnLoad Me
  End If
  
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
  
  mbLoading = True
  cmdOK.Enabled = False
  
  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)
  
  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  ' Read the current settings from the database.
  ReadParameters
  
  ' Initialise all controls with the current settings, or defaults.
  InitialiseBaseTableCombos
  
  ssTabStrip.Tab = giPAGE_COURSES
  
  'AE20080107 Fault #12750
  Changed = False
  
  mbLoading = False
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'AE20071119 Fault #12607
  'If UnloadMode <> vbFormCode And cmdOK.Enabled Then
  If Changed Then
    Select Case MsgBox("Apply module changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
'        If Not SaveChanges Then
'          Cancel = True
'        End If
'        AE20071119 Fault #12607
        Cancel = (Not SaveChanges)
    End Select
  End If
  
End Sub

Private Function SaveChanges() As Boolean
  ' Write the parameter values to the local database.
  SaveChanges = False
  
  If ValidateRelatedColumns Then
  
    SaveCourseRecordParameters
    SavePreRequisiteParameters
    SaveEmployeeRecordParameters
    SaveUnavailabilityParameters
    SaveWaitingListParameters
    SaveTrainingBookingParameters
    SaveRelatedColumns
    
    Application.Changed = True
    
    SaveChanges = True
  End If
  
End Function

Private Sub ReadParameters()
  ' Read the parameter values from the database into local variables.
  ReadCourseRecordParameters
  ReadPreRequisiteParameters
  ReadEmployeeRecordParameters
  ReadUnavailabilityParameters
  ReadWaitingListParameters
  ReadTrainingBookingParameters
  ReadRelatedColumns
  
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Save the form position to registry.
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
  End If

End Sub



Private Sub RefreshCourseColumnControls()
  ' Refresh the Course column controls.
  Dim iCourseTitleListIndex As Integer
  Dim iCourseStartDateListIndex As Integer
  Dim iCourseEndDateListIndex As Integer
  Dim iCourseNumberBookedListIndex As Integer
  Dim iCourseMaxNumberListIndex As Integer
  Dim iCourseCancelDateListIndex As Integer
  Dim iCourseCancelledByListIndex As Integer
  
  iCourseTitleListIndex = 0
  iCourseStartDateListIndex = 0
  iCourseEndDateListIndex = 0
  iCourseNumberBookedListIndex = 0
  iCourseMaxNumberListIndex = 0
  iCourseCancelDateListIndex = 0
  iCourseCancelledByListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboCourseTitle
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboCourseStartDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboCourseEndDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboCourseNumberBooked
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboCourseMaxNumberOfDelegates
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboCourseCancellationDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboCourseCancelledBy
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With recColEdit
  
    .Index = "idxName"
    .Seek ">=", mvar_lngCourseTableID
    
    If Not .NoMatch Then
    
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        
        If !TableID <> mvar_lngCourseTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          If (!DataType = dtVARCHAR) Then
            cboCourseTitle.AddItem !ColumnName
            cboCourseTitle.ItemData(cboCourseTitle.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseTitleID Then
              iCourseTitleListIndex = cboCourseTitle.NewIndex
            End If
          
            cboCourseCancelledBy.AddItem !ColumnName
            cboCourseCancelledBy.ItemData(cboCourseCancelledBy.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseCancelledByID Then
              iCourseCancelledByListIndex = cboCourseCancelledBy.NewIndex
            End If
          End If
          
          If (!DataType = dtTIMESTAMP) Then
            cboCourseStartDate.AddItem !ColumnName
            cboCourseStartDate.ItemData(cboCourseStartDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseStartDateID Then
              iCourseStartDateListIndex = cboCourseStartDate.NewIndex
            End If
            
            cboCourseEndDate.AddItem !ColumnName
            cboCourseEndDate.ItemData(cboCourseEndDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseEndDateID Then
              iCourseEndDateListIndex = cboCourseEndDate.NewIndex
            End If
          
            cboCourseCancellationDate.AddItem !ColumnName
            cboCourseCancellationDate.ItemData(cboCourseCancellationDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseCancelDateID Then
              iCourseCancelDateListIndex = cboCourseCancellationDate.NewIndex
            End If
          End If
          
          If (!DataType = dtinteger) Then
            cboCourseNumberBooked.AddItem !ColumnName
            cboCourseNumberBooked.ItemData(cboCourseNumberBooked.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseNumberBookedID Then
              iCourseNumberBookedListIndex = cboCourseNumberBooked.NewIndex
            End If
          
            cboCourseMaxNumberOfDelegates.AddItem !ColumnName
            cboCourseMaxNumberOfDelegates.ItemData(cboCourseMaxNumberOfDelegates.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCourseMaxNumberID Then
              iCourseMaxNumberListIndex = cboCourseMaxNumberOfDelegates.NewIndex
            End If
          End If
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboCourseTitle.ListIndex = iCourseTitleListIndex
  cboCourseStartDate.ListIndex = iCourseStartDateListIndex
  cboCourseEndDate.ListIndex = iCourseEndDateListIndex
  cboCourseNumberBooked.ListIndex = iCourseNumberBookedListIndex
  cboCourseMaxNumberOfDelegates.ListIndex = iCourseMaxNumberListIndex
  cboCourseCancellationDate.ListIndex = iCourseCancelDateListIndex
  cboCourseCancelledBy.ListIndex = iCourseCancelledByListIndex

  ' Refresh the Course Order controls.
  RefreshCourseOrderControls
  
End Sub
Private Sub RefreshEmployeeColumnControls()
  ' Refresh the Employee column combos with the columns
  ' from the selected Employee table.

  ' Refresh the Employee Order controls.
  RefreshEmployeeOrderControls
  RefreshEmployeeDefaultView
  
End Sub

Private Sub RefreshTrainingBookingsColumnControls()
  ' Refresh the Training Bookings column controls.
  Dim iTrainBookStatusListIndex As Integer
  Dim iTrainBookCancelDateListIndex As Integer
  
  iTrainBookStatusListIndex = 0
  iTrainBookCancelDateListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboTrainingBookingsStatus
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboTrainingBookingsCancellationDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  ' Clear the current contents of the related columns grid combo.
  With ssgrdRelatedColumns.Columns(0)
    .RemoveAll
  End With
  
  ' Add items to the combos for each column that has not been deleted,
  ' or is a system or link column.
  With recColEdit
  
    .Index = "idxName"
    .Seek ">=", mvar_lngTrainBookTableID
    
    If Not .NoMatch Then
    
      Do While Not .EOF
        
        If !TableID <> mvar_lngTrainBookTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          ' Only have character columns for the status.
          If (!DataType = dtVARCHAR) Then
            cboTrainingBookingsStatus.AddItem !ColumnName
            cboTrainingBookingsStatus.ItemData(cboTrainingBookingsStatus.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngTrainBookStatusID Then
              iTrainBookStatusListIndex = cboTrainingBookingsStatus.NewIndex
            End If
          End If
          
          ' Only have date columns for the cancellation date.
          If (!DataType = dtTIMESTAMP) Then
            cboTrainingBookingsCancellationDate.AddItem !ColumnName
            cboTrainingBookingsCancellationDate.ItemData(cboTrainingBookingsCancellationDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngTrainBookCancelDateID Then
              iTrainBookCancelDateListIndex = cboTrainingBookingsCancellationDate.NewIndex
            End If
          End If
          
          ssgrdRelatedColumns.Columns(0).AddItem !ColumnName
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboTrainingBookingsStatus.ListIndex = iTrainBookStatusListIndex
  cboTrainingBookingsCancellationDate.ListIndex = iTrainBookCancelDateListIndex

  ' Refresh the related columns grid.
  RefreshRelatedColumnsGrid

End Sub

Private Sub RefreshPreRequisiteColumnControls()
  ' Refresh the Pre-requisite column controls.
  Dim iPreReqCourseTitleListIndex As Integer
  Dim iPreReqGroupingListIndex As Integer
  Dim iPreReqFailureNotificationListIndex As Integer
  
  iPreReqCourseTitleListIndex = 0
  iPreReqGroupingListIndex = 0
  iPreReqFailureNotificationListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboPreRequisiteCourseTitle
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboPreRequisiteCourseGroup
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboPreRequisiteFailureNotification
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngPreReqTableID
    
    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        
        If !TableID <> mvar_lngPreReqTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          If (!DataType = dtVARCHAR) Then
            cboPreRequisiteCourseTitle.AddItem !ColumnName
            cboPreRequisiteCourseTitle.ItemData(cboPreRequisiteCourseTitle.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngPreReqCourseTitleID Then
              iPreReqCourseTitleListIndex = cboPreRequisiteCourseTitle.NewIndex
            End If
          End If
          
          If (!DataType = dtVARCHAR) And (!Size = 1) Then
            cboPreRequisiteCourseGroup.AddItem !ColumnName
            cboPreRequisiteCourseGroup.ItemData(cboPreRequisiteCourseGroup.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngPreReqGroupingID Then
              iPreReqGroupingListIndex = cboPreRequisiteCourseGroup.NewIndex
            End If
          End If
          
          If (!DataType = dtVARCHAR) Then
            ' JPD 10/4/01 Ensure that the failure notification column
            ' is an option group or dropdown list.
            If ((!ControlType = giCTRL_OPTIONGROUP) Or (!ControlType = giCTRL_COMBOBOX)) And _
              (!columntype <> giCOLUMNTYPE_LOOKUP) Then
              cboPreRequisiteFailureNotification.AddItem !ColumnName
              cboPreRequisiteFailureNotification.ItemData(cboPreRequisiteFailureNotification.NewIndex) = !ColumnID
              If !ColumnID = mvar_lngPreReqFailureNotificationID Then
                iPreReqFailureNotificationListIndex = cboPreRequisiteFailureNotification.NewIndex
              End If
            End If
          End If
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  cboPreRequisiteCourseTitle.ListIndex = iPreReqCourseTitleListIndex
  cboPreRequisiteCourseGroup.ListIndex = iPreReqGroupingListIndex
  cboPreRequisiteFailureNotification.ListIndex = iPreReqFailureNotificationListIndex

End Sub

Private Sub RefreshUnavailabilityColumnControls()
  ' Refresh the Unavailability column controls.
  Dim iUnavailFromDateListIndex As Integer
  Dim iUnavailToDateListIndex As Integer
  Dim iUnavailNotificationListIndex As Integer
  
  iUnavailFromDateListIndex = 0
  iUnavailToDateListIndex = 0
  iUnavailNotificationListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboUnavailabilityFromDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboUnavailabilityToDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  With cboUnavailabilityNotification
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  ' Add items to the combos for each column that has not been deleted,
  ' or is a system or link column.
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngUnavailTableID
    
    If Not .NoMatch Then
      Do While Not .EOF
        If !TableID <> mvar_lngUnavailTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          If (!DataType = dtTIMESTAMP) Then
            cboUnavailabilityFromDate.AddItem !ColumnName
            cboUnavailabilityFromDate.ItemData(cboUnavailabilityFromDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngUnavailFromDateID Then
              iUnavailFromDateListIndex = cboUnavailabilityFromDate.NewIndex
            End If
            
            cboUnavailabilityToDate.AddItem !ColumnName
            cboUnavailabilityToDate.ItemData(cboUnavailabilityToDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngUnavailToDateID Then
              iUnavailToDateListIndex = cboUnavailabilityToDate.NewIndex
            End If
          End If
          
          If (!DataType = dtVARCHAR) Then
            ' JPD 10/4/01 Ensure that the failure notification column
            ' is an option group or dropdown list.
            If ((!ControlType = giCTRL_OPTIONGROUP) Or (!ControlType = giCTRL_COMBOBOX)) And _
              (!columntype <> giCOLUMNTYPE_LOOKUP) Then
              cboUnavailabilityNotification.AddItem !ColumnName
              cboUnavailabilityNotification.ItemData(cboUnavailabilityNotification.NewIndex) = !ColumnID
              If !ColumnID = mvar_lngUnavailFailureNotificationID Then
                iUnavailNotificationListIndex = cboUnavailabilityNotification.NewIndex
              End If
            End If
          End If
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboUnavailabilityFromDate.ListIndex = iUnavailFromDateListIndex
  cboUnavailabilityToDate.ListIndex = iUnavailToDateListIndex
  cboUnavailabilityNotification.ListIndex = iUnavailNotificationListIndex

End Sub
Private Sub RefreshWaitingListColumnControls()
  ' Refresh the Waiting List column controls.
  Dim iWaitListCourseTitleListIndex As Integer
  
  iWaitListCourseTitleListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboWaitingListCourseTitle
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  ' Clear the current contents of the related columns grid combo.
  With ssgrdRelatedColumns.Columns(1)
    .RemoveAll
  End With
  
  ' Add items to the combos for each column that has not been deleted,
  ' or is a system or link column.
  With recColEdit
  
    .Index = "idxName"
    .Seek ">=", mvar_lngWaitListTableID
    
    If Not .NoMatch Then
    
      Do While Not .EOF
        
        If !TableID <> mvar_lngWaitListTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          If (!DataType = dtVARCHAR) Then
            cboWaitingListCourseTitle.AddItem !ColumnName
            cboWaitingListCourseTitle.ItemData(cboWaitingListCourseTitle.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngWaitListCourseTitleID Then
              iWaitListCourseTitleListIndex = cboWaitingListCourseTitle.NewIndex
            End If
          End If
          
          ssgrdRelatedColumns.Columns(1).AddItem !ColumnName
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboWaitingListCourseTitle.ListIndex = iWaitListCourseTitleListIndex

  ' Refresh the related columns grid.
  RefreshRelatedColumnsGrid
  
End Sub
Private Sub RefreshWaitingListOverrideColumnControls()

  ' Refresh the override column controls.
  Dim iWaitListOverrideColumnIndex As Integer
  
  iWaitListOverrideColumnIndex = 0
  
  ' Clear the current contents of the combos.
  With cboWaitingListOrderOveride
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  ' Clear the current contents of the related columns grid combo.
  With ssgrdRelatedColumns.Columns(1)
    .RemoveAll
  End With
  
  ' Add items to the combos for each column that has not been deleted,
  ' or is a system or link column.
  With recColEdit
  
    .Index = "idxName"
    .Seek ">=", mvar_lngWaitListTableID
    
    If Not .NoMatch Then
    
      Do While Not .EOF
        
        If !TableID <> mvar_lngWaitListTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
                          
          If (!DataType = dtTIMESTAMP) Then
            cboWaitingListOrderOveride.AddItem !ColumnName
            cboWaitingListOrderOveride.ItemData(cboWaitingListOrderOveride.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngWaitListOrderOverideColumn Then
              iWaitListOverrideColumnIndex = cboWaitingListOrderOveride.NewIndex
            End If
          End If
          
          ssgrdRelatedColumns.Columns(1).AddItem !ColumnName
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboWaitingListOrderOveride.ListIndex = iWaitListOverrideColumnIndex
  ' Refresh the related columns grid.
  RefreshRelatedColumnsGrid
End Sub

Private Sub InitialiseBaseTableCombos()
  ' Initialise the Base Table combos (ie. Courses and Employees).
  ' NB. The columns and history table controls that are dependent on the selections in
  ' these combos, are automatically refreshed.
  Dim iCourseTableListIndex As Integer
  Dim iEmployeeTableListIndex As Integer
  
  
  iCourseTableListIndex = 0
  iEmployeeTableListIndex = 0

  
  ' Clear the combos, and add '<None>' items.
  With cboCourseTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboEmployeeTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  ' Add items to the combos for each table that has not been deleted.
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted Then
        cboCourseTable.AddItem !TableName
        cboCourseTable.ItemData(cboCourseTable.NewIndex) = !TableID
        If !TableID = mvar_lngCourseTableID Then
          iCourseTableListIndex = cboCourseTable.NewIndex
        End If
      
        cboEmployeeTable.AddItem !TableName
        cboEmployeeTable.ItemData(cboEmployeeTable.NewIndex) = !TableID
        If !TableID = mvar_lngEmployeeTableID Then
          iEmployeeTableListIndex = cboEmployeeTable.NewIndex
        End If
      End If
      
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
  With cboCourseTable
    '.Enabled = True
    .Enabled = Not mblnReadOnly
    .ListIndex = iCourseTableListIndex
  End With
  
  With cboEmployeeTable
    '.Enabled = True
    .Enabled = Not mblnReadOnly
    .ListIndex = iEmployeeTableListIndex
  End With

End Sub
Private Sub RefreshEmployeeDefaultView()
  'NHRD01052003 Fault 4687
  Dim iDefaultViewTableListIndex As Integer

  iDefaultViewTableListIndex = 0

  With cboDefaultView
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With

  ' Add items to the combos for each table that has not been deleted.
  With recViewEdit
    .Index = "idxViewName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If Not !Deleted And (!ViewTableID = mvar_lngEmployeeTableID) Then
        cboDefaultView.AddItem !ViewName
        cboDefaultView.ItemData(cboDefaultView.NewIndex) = !ViewID
        
        'NHRD01052003 Fault 4687
        If !ViewID = mvar_lngBulkBookingDefaultViewID Then
          iDefaultViewTableListIndex = cboDefaultView.NewIndex
        End If
       End If
       .MoveNext
    Loop
  End With

  With cboDefaultView
    .Enabled = Not mblnReadOnly
    .ListIndex = iDefaultViewTableListIndex
  End With

End Sub
Private Sub optCourseOverbookingNotification_Click(Index As Integer)
  ' Save the selected ID to a local variable.
  mvar_iCourseOverbookingNotification = Index
  Changed = True
End Sub

Private Sub optPreRequisiteDefaultFailureNotification_Click(Index As Integer)
  ' Save the selected ID to a local variable.
  mvar_iPreReqDfltFailureNotification = Index
  Changed = True
End Sub


Private Sub optTrainingBookingsOverlappedNotification_Click(Index As Integer)
  ' Save the selected ID to a local variable.
  mvar_iTrainBookOverlapNotification = Index
  Changed = True
End Sub

Private Sub optUnavailabilityDefaultNotification_Click(Index As Integer)
  ' Save the selected ID to a local variable.
  mvar_iUnavailDfltFailureNotification = Index
  Changed = True
End Sub


Private Sub ssgrdRelatedColumns_ComboCloseUp()
  ssgrdRelatedColumns.Redraw = False
  ssgrdRelatedColumns.Update
  ssgrdRelatedColumns.Redraw = True
  Me.Changed = True
End Sub

Private Sub ssgrdRelatedColumns_LostFocus()
  ' Save the grid columns into the array.
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim lngTrainBookColumnID As Long
  Dim lngWaitListColumnID As Long
  
  ReDim mvar_alngRelatedColumns(2, 0)

  With ssgrdRelatedColumns
    .MoveFirst
    
    For iLoop = 1 To .Rows

      recColEdit.Index = "idxName"

      recColEdit.Seek "=", mvar_lngTrainBookTableID, .Columns(0).Text, False
      If recColEdit.NoMatch Then
        lngTrainBookColumnID = 0
      Else
        lngTrainBookColumnID = recColEdit!ColumnID
      End If

      recColEdit.Seek "=", mvar_lngWaitListTableID, .Columns(1).Text, False
      If recColEdit.NoMatch Then
        lngWaitListColumnID = 0
      Else
        lngWaitListColumnID = recColEdit!ColumnID
      End If

      ' Add the column IDs to the array if they are both valid.
      If (lngTrainBookColumnID > 0) And _
        (lngWaitListColumnID > 0) Then

        iNextIndex = UBound(mvar_alngRelatedColumns, 2) + 1
        ReDim Preserve mvar_alngRelatedColumns(2, iNextIndex)
        mvar_alngRelatedColumns(1, iNextIndex) = lngTrainBookColumnID
        mvar_alngRelatedColumns(2, iNextIndex) = lngWaitListColumnID
      End If
      
      .MoveNext
    Next iLoop
  End With
    
End Sub


Private Sub ssTabStrip_Click(PreviousTab As Integer)
  ' Enable, and make visible the selected tab.
  
  If Not mblnReadOnly Then
    fraCourses.Enabled = (ssTabStrip.Tab = giPAGE_COURSES)
    fraPreRequisites.Enabled = (ssTabStrip.Tab = giPAGE_PREREQUISITES)
    fraDelegates.Enabled = (ssTabStrip.Tab = giPAGE_DELEGATES)
    fraUnavailability.Enabled = (ssTabStrip.Tab = giPAGE_DELEGATES)
    fraWaitingList.Enabled = (ssTabStrip.Tab = giPAGE_WAITINGLISTBOOKINGS)
    fraTrainingBookings.Enabled = (ssTabStrip.Tab = giPAGE_WAITINGLISTBOOKINGS)
    fraRelatedColumns.Enabled = (ssTabStrip.Tab = giPAGE_RELATEDCOLUMNS)
  End If

End Sub


Private Sub RefreshCoursesHistoryTableCombos()
  ' Refresh the Table combos that are dependent on the Course Table selection.
  Dim iPreRequisiteTableListIndex As Integer
  
  iPreRequisiteTableListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboPreRequisiteTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    ' Add  an item to the combos for each table that has not been deleted,
    ' and is a child of the selected Courses table.
    Do While Not .EOF

      If Not !Deleted Then
        With recRelEdit
          .Index = "idxParentID"
          .Seek "=", mvar_lngCourseTableID, recTabEdit!TableID
          If Not .NoMatch Then
            cboPreRequisiteTable.AddItem recTabEdit!TableName
            cboPreRequisiteTable.ItemData(cboPreRequisiteTable.NewIndex) = recTabEdit!TableID
          
            If recTabEdit!TableID = mvar_lngPreReqTableID Then
              iPreRequisiteTableListIndex = cboPreRequisiteTable.NewIndex
            End If
          End If
        End With
      End If

      .MoveNext
    Loop
  End With

  ' Select the appropriate combo items.
  cboPreRequisiteTable.ListIndex = iPreRequisiteTableListIndex

  ' Refresh the Table combos that are dependent on both the Courses and Employees Table selection.
  RefreshSharedHistoryTableCombos
  
End Sub
Private Sub RefreshSharedHistoryTableCombos()
  ' Refresh the Table combos that are dependent on the both the Courses and Employees Table selections.
  Dim iTrainBookTableListIndex As Integer
  
  iTrainBookTableListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboTrainingBookingsTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  ' Add an item to the combos for each table that has not been deleted,
  ' and is a child of both the selected Courses and Employees table.
  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF

      If Not !Deleted Then
      
        With recRelEdit
          .Index = "idxParentID"
          .Seek "=", mvar_lngCourseTableID, recTabEdit!TableID
          
          If Not .NoMatch Then
            .Seek "=", mvar_lngEmployeeTableID, recTabEdit!TableID
            If Not .NoMatch Then
              cboTrainingBookingsTable.AddItem recTabEdit!TableName
              cboTrainingBookingsTable.ItemData(cboTrainingBookingsTable.NewIndex) = recTabEdit!TableID
              If recTabEdit!TableID = mvar_lngTrainBookTableID Then
                iTrainBookTableListIndex = cboTrainingBookingsTable.NewIndex
              End If
            End If
          End If
          
        End With
      End If

      .MoveNext
    Loop
  End With

  ' Select the appropriate combo item.
  cboTrainingBookingsTable.ListIndex = iTrainBookTableListIndex

End Sub


Private Sub RefreshEmployeesHistoryTableCombos()
  ' Refresh the Table combos that are dependent on the Employees Table selection.
  Dim iUnavailabilityTableListIndex As Integer
  Dim iWaitingListTableListIndex As Integer
  Dim iPreRequisiteTableListIndex As Integer
  
  iUnavailabilityTableListIndex = 0
  iWaitingListTableListIndex = 0
  
  ' Clear the current contents of the combos.
  With cboUnavailabilityTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboWaitingListTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    ' Add  an item to the combos for each table that has not been deleted,
    ' and is a child of the selected Courses table.
    Do While Not .EOF

      If Not !Deleted Then
        With recRelEdit
          .Index = "idxParentID"
          .Seek "=", mvar_lngEmployeeTableID, recTabEdit!TableID
          If Not .NoMatch Then
            cboUnavailabilityTable.AddItem recTabEdit!TableName
            cboUnavailabilityTable.ItemData(cboUnavailabilityTable.NewIndex) = recTabEdit!TableID
            If recTabEdit!TableID = mvar_lngUnavailTableID Then
              iUnavailabilityTableListIndex = cboUnavailabilityTable.NewIndex
            End If
            
            cboWaitingListTable.AddItem recTabEdit!TableName
            cboWaitingListTable.ItemData(cboWaitingListTable.NewIndex) = recTabEdit!TableID
            If recTabEdit!TableID = mvar_lngWaitListTableID Then
              iWaitingListTableListIndex = cboWaitingListTable.NewIndex
            End If
          End If
        End With
      End If

      .MoveNext
    Loop
  End With

  cboUnavailabilityTable.ListIndex = iUnavailabilityTableListIndex
  cboWaitingListTable.ListIndex = iWaitingListTableListIndex

  ' Refresh the Table combos that are dependent on both the Courses and Employees Table selection.
  RefreshSharedHistoryTableCombos
  
End Sub


Private Sub ReadCourseRecordParameters()
  ' Read the Course Records parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Course table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETABLE
    If .NoMatch Then
      mvar_lngCourseTableID = 0
    Else
      mvar_lngCourseTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Title column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETITLE
    If .NoMatch Then
      mvar_lngCourseTitleID = 0
    Else
      mvar_lngCourseTitleID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Start date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSESTARTDATE
    If .NoMatch Then
      mvar_lngCourseStartDateID = 0
    Else
      mvar_lngCourseStartDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course End Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEENDDATE
    If .NoMatch Then
      mvar_lngCourseEndDateID = 0
    Else
      mvar_lngCourseEndDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Number Booked column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSENUMBERBOOKED
    If .NoMatch Then
      mvar_lngCourseNumberBookedID = 0
    Else
      mvar_lngCourseNumberBookedID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Max. Number of Delegates column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEMAXNUMBER
    If .NoMatch Then
      mvar_lngCourseMaxNumberID = 0
    Else
      mvar_lngCourseMaxNumberID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Cancellation Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLATIONDATE
    If .NoMatch Then
      mvar_lngCourseCancelDateID = 0
    Else
      mvar_lngCourseCancelDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Cancelled By column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLEDBY
    If .NoMatch Then
      mvar_lngCourseCancelledByID = 0
    Else
      mvar_lngCourseCancelledByID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Course Transfer Provisional Bookings flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETRANSFERPROVISIONALS
    If .NoMatch Then
      mvar_fCourseTransferProvisionals = False
    Else
      mvar_fCourseTransferProvisionals = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkCourseTransferProvisionalBookingsWhenCancelling.value = IIf(mvar_fCourseTransferProvisionals, vbChecked, vbUnchecked)
    
    ' Get the Course Include Provisional Bookings flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS
    If .NoMatch Then
      mvar_fCourseIncludeProvisionals = False
    Else
      mvar_fCourseIncludeProvisionals = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkCourseIncludeProvisionalBookingsInNumberBookedColumn.value = IIf(mvar_fCourseIncludeProvisionals, vbChecked, vbUnchecked)
    
    ' Get the Overbooking Notification flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION
    If .NoMatch Then
      mvar_iCourseOverbookingNotification = 0
    Else
      mvar_iCourseOverbookingNotification = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    optCourseOverbookingNotification(mvar_iCourseOverbookingNotification).value = True
    
    ' Get the Course Order ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEORDER
    If .NoMatch Then
      mvar_lngCourseOrderID = 0
    Else
      mvar_lngCourseOrderID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
  End With

End Sub
Private Sub ReadPreRequisiteParameters()
  ' Read the Pre-requisite parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"
        
    ' Get the Pre-requisite table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQTABLE
    If .NoMatch Then
      mvar_lngPreReqTableID = 0
    Else
      mvar_lngPreReqTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Pre-requisite Course Title column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQCOURSETITLE
    If .NoMatch Then
      mvar_lngPreReqCourseTitleID = 0
    Else
      mvar_lngPreReqCourseTitleID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Pre-requisite Grouping column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQGROUPING
    If .NoMatch Then
      mvar_lngPreReqGroupingID = 0
    Else
      mvar_lngPreReqGroupingID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Pre-requisite Failure Notification column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQFAILURE
    If .NoMatch Then
      mvar_lngPreReqFailureNotificationID = 0
    Else
      mvar_lngPreReqFailureNotificationID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Pre-requisite Default Failure Notification flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQDFLTFAILURE
    If .NoMatch Then
      mvar_iPreReqDfltFailureNotification = 0
    Else
      mvar_iPreReqDfltFailureNotification = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    optPreRequisiteDefaultFailureNotification(mvar_iPreReqDfltFailureNotification).value = True
    
  End With

End Sub

Private Sub ReadUnavailabilityParameters()
  ' Read the Unavailability parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Unavailability table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTABLE
    If .NoMatch Then
      mvar_lngUnavailTableID = 0
    Else
      mvar_lngUnavailTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Unavailability From Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFROMDATE
    If .NoMatch Then
      mvar_lngUnavailFromDateID = 0
    Else
      mvar_lngUnavailFromDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Unavailability To Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTODATE
    If .NoMatch Then
      mvar_lngUnavailToDateID = 0
    Else
      mvar_lngUnavailToDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Unavailability Failure Notification column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFAILURE
    If .NoMatch Then
      mvar_lngUnavailFailureNotificationID = 0
    Else
      mvar_lngUnavailFailureNotificationID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Unavailability Default Failure Notification flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILDFLTFAILURE
    If .NoMatch Then
      mvar_iUnavailDfltFailureNotification = 0
    Else
      mvar_iUnavailDfltFailureNotification = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    optUnavailabilityDefaultNotification(mvar_iUnavailDfltFailureNotification).value = True
    
  End With

End Sub


Private Sub ReadWaitingListParameters()
  ' Read the Waiting List parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Unavailability table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTTABLE
    If .NoMatch Then
      mvar_lngWaitListTableID = 0
    Else
      mvar_lngWaitListTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Unavailability From Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTCOURSETITLE
    If .NoMatch Then
      mvar_lngWaitListCourseTitleID = 0
    Else
      mvar_lngWaitListCourseTitleID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Override column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTOVERRIDECOLUMN
    If .NoMatch Then
      mvar_lngWaitListOrderOverideColumn = 0
    Else
      mvar_lngWaitListOrderOverideColumn = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
  End With

End Sub



Private Sub ReadTrainingBookingParameters()
  ' Read the Training Booking parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Training Booking table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKTABLE
    If .NoMatch Then
      mvar_lngTrainBookTableID = 0
    Else
      mvar_lngTrainBookTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Training Booking Status column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKSTATUS
    If .NoMatch Then
      mvar_lngTrainBookStatusID = 0
    Else
      mvar_lngTrainBookStatusID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Training Booking Cancellation Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKCANCELDATE
    If .NoMatch Then
      mvar_lngTrainBookCancelDateID = 0
    Else
      mvar_lngTrainBookCancelDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Training Booking Overlap Notification flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION
    If .NoMatch Then
      mvar_iTrainBookOverlapNotification = 0
    Else
      mvar_iTrainBookOverlapNotification = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    optTrainingBookingsOverlappedNotification(mvar_iTrainBookOverlapNotification).value = True
    
  End With

End Sub




Private Sub ReadEmployeeRecordParameters()
  ' Read the Employee Records parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Employee Table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEETABLE
    If .NoMatch Then
      mvar_lngEmployeeTableID = 0
    Else
      mvar_lngEmployeeTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Employee Order ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEEORDER
    If .NoMatch Then
      mvar_lngEmployeeOrderID = 0
    Else
      mvar_lngEmployeeOrderID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
  
    'NHRD01052003 Fault 4687 Get the Bulk Booking Default View
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW
    If .NoMatch Then
      mvar_lngBulkBookingDefaultViewID = 0
    Else
      mvar_lngBulkBookingDefaultViewID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
  End With

End Sub


Private Sub SaveCourseRecordParameters()
  ' Save the Course Records parameter values to the local database.

  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Save the Course table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSETABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngCourseTableID
    .Update
    
    ' Save the Course Title column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETITLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSETITLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseTitleID
    .Update
    
    ' Save the Course Start Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSESTARTDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSESTARTDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseStartDateID
    .Update
    
    ' Save the Course End Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEENDDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSEENDDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseEndDateID
    .Update
    
    ' Save the Course Number Booked column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSENUMBERBOOKED
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSENUMBERBOOKED
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseNumberBookedID
    .Update
        
    ' Save the Course Max. Number of Delegates column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEMAXNUMBER
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSEMAXNUMBER
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseMaxNumberID
    .Update
    
    ' Save the Course Cancelled Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLATIONDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSECANCELLATIONDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseCancelDateID
    .Update
    
    ' Save the Course Cancelled By column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLEDBY
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSECANCELLEDBY
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCourseCancelledByID
    .Update

    ' Save the Course Transfer Provisional Bookings flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETRANSFERPROVISIONALS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSETRANSFERPROVISIONALS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_fCourseTransferProvisionals, "TRUE", "FALSE")
    .Update
    
    ' Save the Course Include Provisional Bookings flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_fCourseIncludeProvisionals, "TRUE", "FALSE")
    .Update
    
    ' Save the Course Overbooking Notification Type flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_iCourseOverbookingNotification
    .Update
    
    ' Save the Course Order ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEORDER
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_COURSEORDER
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_ORDERID
    !parametervalue = mvar_lngCourseOrderID
    .Update
    
  End With

End Sub
Private Sub SavePreRequisiteParameters()
  ' Save the Pre-requisite Records parameter values to the local database.

  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Save the Pre-requisite table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_PREREQTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngPreReqTableID
    .Update
    
    ' Save the Pre-requisite Course Title column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQCOURSETITLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_PREREQCOURSETITLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngPreReqCourseTitleID
    .Update

    ' Save the Pre-requisite Grouping column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQGROUPING
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_PREREQGROUPING
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngPreReqGroupingID
    .Update

    ' Save the Pre-requisite Failure Notification column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQFAILURE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_PREREQFAILURE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngPreReqFailureNotificationID
    .Update
    
    ' Save the Pre-requisite Default Failure Notification Type flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQDFLTFAILURE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_PREREQDFLTFAILURE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_iPreReqDfltFailureNotification
    .Update

  End With

End Sub

Private Sub SaveUnavailabilityParameters()
  ' Save the Unavailability parameter values to the local database.
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Save the Unavailability table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_UNAVAILTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngUnavailTableID
    .Update
    
    ' Save the Unavailability From Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFROMDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_UNAVAILFROMDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngUnavailFromDateID
    .Update

    ' Save the Unavailability To Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTODATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_UNAVAILTODATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngUnavailToDateID
    .Update

    ' Save the Unavailability Failure Notification column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFAILURE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_UNAVAILFAILURE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngUnavailFailureNotificationID
    .Update
    
    ' Save the Unavailability Default Failure Notification Type flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILDFLTFAILURE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_UNAVAILDFLTFAILURE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_iUnavailDfltFailureNotification
    .Update

  End With

End Sub


Private Sub SaveWaitingListParameters()
  ' Save the Waiting List parameter values to the local database.
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Save the Waiting List table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_WAITLISTTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngWaitListTableID
    .Update
    
    ' Save the Waiting List Course Title column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTCOURSETITLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_WAITLISTCOURSETITLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngWaitListCourseTitleID
    .Update

    ' Save the Waiting List Override column.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTOVERRIDECOLUMN
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_WAITLISTOVERRIDECOLUMN
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngWaitListOrderOverideColumn
    .Update
    
  End With

End Sub


Private Sub SaveTrainingBookingParameters()
  ' Save the Training Booking parameter values to the local database.
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Save the Training Booking table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_TRAINBOOKTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngTrainBookTableID
    .Update
    
    ' Save the Training Booking Status column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKSTATUS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_TRAINBOOKSTATUS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngTrainBookStatusID
    .Update

    ' Save the Training Booking Cancellation Date column ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKCANCELDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_TRAINBOOKCANCELDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngTrainBookCancelDateID
    .Update

    ' Save the Training Booking Default Overbooking Notification Type flag.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_iTrainBookOverlapNotification
    .Update
  
  End With

End Sub



Private Sub SaveEmployeeRecordParameters()
  ' Save the Course Records parameter values to the local database.
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Employee table ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEETABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_EMPLOYEETABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngEmployeeTableID
    .Update
    
    ' Get the Employee Order ID.
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEEORDER
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_EMPLOYEEORDER
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_ORDERID
    !parametervalue = mvar_lngEmployeeOrderID
    .Update
    
    ' Get the Default View ID
    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_TRAININGBOOKING
      !parameterkey = gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_VIEWID
    !parametervalue = mvar_lngBulkBookingDefaultViewID 'NHRD01052003 Fault 4687
    .Update
  End With
End Sub


Private Sub RefreshCourseOrderControls()
  ' Refresh the Course Order controls.
  ' Validate the order selection at the same time.
  Dim sOrderName As String
  Dim objOrder As Order
  
  ' Check if the selected order is for the current Course table.
  With recOrdEdit
    .Index = "idxID"
    .Seek "=", mvar_lngCourseOrderID
    
    If Not .NoMatch Then
      If !TableID <> mvar_lngCourseTableID Then
        mvar_lngCourseOrderID = 0
      End If
    Else
      mvar_lngCourseOrderID = 0
    End If
  End With
  
  ' Initialise the default values.
  sOrderName = ""
    
  ' Instantiate a new Order object.
  Set objOrder = New Order
  With objOrder
    .OrderID = mvar_lngCourseOrderID
    
    ' Read the name of the current order.
    If .ConstructOrder Then
      sOrderName = .OrderName
    End If
  End With
  
  ' Disassociate object variables.
  Set objOrder = Nothing
  
  ' Update the control's properties.
  txtCourseTableOrder.Text = sOrderName

End Sub
Private Sub RefreshEmployeeOrderControls()
  ' Refresh the Employee Order controls.
  ' Validate the order selection at the same time.
  Dim sOrderName As String
  Dim objOrder As Order
  
  ' Check if the selected order is for the current Course table.
  With recOrdEdit
  
    .Index = "idxID"
    .Seek "=", mvar_lngEmployeeOrderID
    
    If Not .NoMatch Then
      If !TableID <> mvar_lngEmployeeTableID Then
        mvar_lngEmployeeOrderID = 0
      End If
    Else
      mvar_lngEmployeeOrderID = 0
    End If
  End With
  
  ' Initialise the default values.
  sOrderName = ""
    
  ' Instantiate a new Order object.
  Set objOrder = New Order
  With objOrder
    .OrderID = mvar_lngEmployeeOrderID
    
    ' Read the name of the current order.
    If .ConstructOrder Then
      sOrderName = .OrderName
    End If
  End With
  
  ' Disassociate object variables.
  Set objOrder = Nothing
  
  ' Update the control's properties.
  txtEmployeeTableOrder.Text = sOrderName

End Sub


Private Sub RefreshRelatedColumnsGrid()
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim sTBColumnName As String
  Dim sWLColumnName As String
  
  ' Refresh the array of related columns.
  ' Reset any columns are in not the current Training Booking or Waiting List tables.
  For iLoop = 1 To UBound(mvar_alngRelatedColumns, 2)
    
    With recColEdit
      .Index = "idxColumnID"
      
      ' Check that the given Training Booking column is in the Training Booking table.
      .Seek "=", mvar_alngRelatedColumns(1, iLoop)
      fOK = Not .NoMatch
      
      If fOK Then
        fOK = (!TableID = mvar_lngTrainBookTableID)
      End If
      
      If fOK Then
        ' Check that the given Waiting List column is in the Waiting List table.
        .Seek "=", mvar_alngRelatedColumns(2, iLoop)
        fOK = Not .NoMatch
      End If
      
      If fOK Then
        fOK = (!TableID = mvar_lngWaitListTableID)
      End If
      
      If Not fOK Then
        mvar_alngRelatedColumns(1, iLoop) = 0
        mvar_alngRelatedColumns(2, iLoop) = 0
      End If
      
    End With
  Next iLoop
  
  ' Populate the grid with the given columns.
  With ssgrdRelatedColumns
    .Update
    .RemoveAll

    For iLoop = 1 To UBound(mvar_alngRelatedColumns, 2)
      ' Get the column names.
      recColEdit.Index = "idxColumnID"

      recColEdit.Seek "=", mvar_alngRelatedColumns(1, iLoop)
      If recColEdit.NoMatch Then
        mvar_alngRelatedColumns(1, iLoop) = 0
      Else
        sTBColumnName = recColEdit!ColumnName
      End If

      recColEdit.Seek "=", mvar_alngRelatedColumns(2, iLoop)
      If recColEdit.NoMatch Then
        mvar_alngRelatedColumns(2, iLoop) = 0
      Else
        sWLColumnName = recColEdit!ColumnName
      End If
      
      If (mvar_alngRelatedColumns(1, iLoop) > 0) And _
        (mvar_alngRelatedColumns(2, iLoop) > 0) Then
        
        .AddItem sTBColumnName & vbTab & sWLColumnName
      End If
      
    Next iLoop
    
    ' Enable the 'Delete' command button if we have at least one row.
    If Not mblnReadOnly Then
      cmdDeleteRelatedColumn.Enabled = (.Rows > 0)
      cmdAddRelatedColumn.Enabled = (.Columns(0).ListCount > 0) And (.Columns(1).ListCount > 0)
    End If
  
  End With

End Sub

Private Sub ReadRelatedColumns()
  ' Read the Related Columns information into a local array.
  Dim iNextIndex As Integer
  Dim lngTrainBookColumnID As Long
  Dim lngWaitListColumnID As Long
  
  ReDim mvar_alngRelatedColumns(2, 0)
  
  With recModuleRelatedColumns
    .Index = "idxModuleParameter"

    .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TBWLRELATEDCOLUMNS
    If Not .NoMatch Then
      Do While Not .EOF
        If (!moduleKey <> gsMODULEKEY_TRAININGBOOKING) Or _
          (!parameterkey <> gsPARAMETERKEY_TBWLRELATEDCOLUMNS) Then
          
          Exit Do
        End If
        
        ' Read the related column IDs from the database.
        lngTrainBookColumnID = IIf(IsNull(!sourcecolumnid) Or Len(!sourcecolumnid) = 0, 0, !sourcecolumnid)
        lngWaitListColumnID = IIf(IsNull(!destcolumnid) Or Len(!destcolumnid) = 0, 0, !destcolumnid)
        
        ' Add the column IDs to the array if they are both valid.
        If (lngTrainBookColumnID > 0) And _
          (lngWaitListColumnID > 0) Then
          
          iNextIndex = UBound(mvar_alngRelatedColumns, 2) + 1
          ReDim Preserve mvar_alngRelatedColumns(2, iNextIndex)
          mvar_alngRelatedColumns(1, iNextIndex) = lngTrainBookColumnID
          mvar_alngRelatedColumns(2, iNextIndex) = lngWaitListColumnID
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
End Sub
Private Function ValidateRelatedColumns() As Boolean
  ' Save the Related Columns information to the database.
  Dim iLoop As Integer
  Dim sRelatedColumns As String
  Dim sSubRelatedColumns As String
  
  ssgrdRelatedColumns_LostFocus
  
  sRelatedColumns = ""
  
  ' Check for duplicates in the related columns array.
  For iLoop = 1 To UBound(mvar_alngRelatedColumns, 2)
    sSubRelatedColumns = CStr(mvar_alngRelatedColumns(1, iLoop)) & "-" & CStr(mvar_alngRelatedColumns(2, iLoop))
  
    If InStr(sRelatedColumns, sSubRelatedColumns) > 0 Then
      MsgBox "Unable to save changes." & vbCrLf & _
        "Duplicate entries in the Related Columns list.", vbExclamation + vbOKOnly, App.Title
      ValidateRelatedColumns = False
      Exit Function
    End If
  
    sRelatedColumns = sRelatedColumns & vbTab & sSubRelatedColumns
  Next iLoop
  
  ValidateRelatedColumns = True
  
End Function

Private Sub SaveRelatedColumns()
  ' Save the Related Columns information to the database.
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim lngTrainBookColumnID As Long
  Dim lngWaitListColumnID As Long
  
  ' Clear the current database values.
  daoDb.Execute "DELETE FROM tmpModuleRelatedColumns WHERE moduleKey = '" & gsMODULEKEY_TRAININGBOOKING & "'" & _
    " AND parameterKey = '" & gsPARAMETERKEY_TBWLRELATEDCOLUMNS & "'", dbFailOnError
  
  ' Refresh the related columns array.
  With ssgrdRelatedColumns
    .MoveFirst
    
    For iLoop = 1 To .Rows

      recColEdit.Index = "idxName"

      recColEdit.Seek "=", mvar_lngTrainBookTableID, .Columns(0).Text, False
      If recColEdit.NoMatch Then
        lngTrainBookColumnID = 0
      Else
        lngTrainBookColumnID = recColEdit!ColumnID
      End If

      recColEdit.Seek "=", mvar_lngWaitListTableID, .Columns(1).Text, False
      If recColEdit.NoMatch Then
        lngWaitListColumnID = 0
      Else
        lngWaitListColumnID = recColEdit!ColumnID
      End If

      ' Add the column IDs to the array if they are both valid.
      If (lngTrainBookColumnID > 0) And _
        (lngWaitListColumnID > 0) Then

        ' Write the new related column values to the database.
        With recModuleRelatedColumns
          .AddNew
          !moduleKey = gsMODULEKEY_TRAININGBOOKING
          !parameterkey = gsPARAMETERKEY_TBWLRELATEDCOLUMNS
          !sourcecolumnid = lngTrainBookColumnID
          !destcolumnid = lngWaitListColumnID
          .Update
        End With
      End If
      
      .MoveNext
      
    Next iLoop
  End With
  
End Sub


Private Sub RefreshRelatedColumns()

  DoEvents

  With ssgrdRelatedColumns
  
    .Columns("TrainingBooking").Width = (.Width / 2) - 5
    .Columns("WaitingList").Width = (.Width / 2) - 5
  
    If .VisibleRows < .Rows Then
      .ScrollBars = ssScrollBarsVertical
      .Columns("TrainingBooking").Width = .Columns("TrainingBooking").Width - 120
      .Columns("WaitingList").Width = .Columns("WaitingList").Width - 120
    Else
      .ScrollBars = ssScrollBarsNone
    End If

  End With

End Sub


