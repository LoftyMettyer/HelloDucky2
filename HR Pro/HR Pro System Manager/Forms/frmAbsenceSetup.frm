VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "coa_line.ocx"
Begin VB.Form frmAbsenceSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Absence Setup"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5004
   Icon            =   "frmAbsenceSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4755
      TabIndex        =   82
      Top             =   5535
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3510
      TabIndex        =   81
      Top             =   5535
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5310
      Left            =   105
      TabIndex        =   83
      Top             =   105
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   9366
      _Version        =   393216
      Style           =   1
      Tabs            =   5
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
      TabCaption(0)   =   "&Absence"
      TabPicture(0)   =   "frmAbsenceSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAbsence"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Absence &Type"
      TabPicture(1)   =   "frmAbsenceSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAbsenceType"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calen&dar"
      TabPicture(2)   =   "frmAbsenceSetup.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCalendarInclude"
      Tab(2).Control(1)=   "fraCalendarDef"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&SSP"
      TabPicture(3)   =   "frmAbsenceSetup.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraSSPColumns"
      Tab(3).Control(1)=   "fraWorkingDays"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&Parental Leave"
      TabPicture(4)   =   "frmAbsenceSetup.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(1)=   "Frame2"
      Tab(4).Control(2)=   "Frame1"
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Personnel Records :"
         Height          =   825
         Left            =   -74850
         TabIndex        =   62
         Top             =   390
         Width           =   5565
         Begin VB.ComboBox cboParentalLeaveRegion 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   300
            Width           =   2505
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parental Leave Region Column:"
            Height          =   195
            Left            =   150
            TabIndex        =   63
            Top             =   360
            Width           =   2715
         End
      End
      Begin VB.Frame fraCalendarDef 
         Caption         =   "Calendar Definition :"
         Enabled         =   0   'False
         Height          =   825
         Left            =   -74850
         TabIndex        =   5
         Top             =   390
         Width           =   5565
         Begin VB.ComboBox cboCalStartMonth 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   300
            Width           =   2500
         End
         Begin VB.Label lblCalStartMonth 
            BackStyle       =   0  'Transparent
            Caption         =   "Calendar Start Month :"
            Height          =   195
            Left            =   195
            TabIndex        =   85
            Top             =   360
            Width           =   2025
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dependants :"
         Height          =   2595
         Left            =   -74850
         TabIndex        =   70
         Top             =   2555
         Width           =   5565
         Begin VB.ComboBox cboDependantsTable 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   300
            Width           =   2505
         End
         Begin VB.ComboBox cboDepDisabled 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   2145
            Width           =   2505
         End
         Begin VB.ComboBox cboDepAdoptedDate 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1740
            Width           =   2505
         End
         Begin VB.ComboBox cboDepDateOfBirth 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1335
            Width           =   2505
         End
         Begin VB.ComboBox cboDepChildNo 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   945
            Width           =   2505
         End
         Begin COALine.COA_Line ASRDummyLine3 
            Height          =   30
            Left            =   120
            Top             =   760
            Width           =   5200
            _ExtentX        =   9181
            _ExtentY        =   53
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dependants Table :"
            Height          =   195
            Left            =   150
            TabIndex        =   71
            Top             =   360
            Width           =   2085
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disabled Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   79
            Top             =   2205
            Width           =   1950
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adopted Date Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   77
            Top             =   1800
            Width           =   2355
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   75
            Top             =   1395
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Child No. Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   73
            Top             =   1005
            Width           =   1995
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Absence :"
         Height          =   1260
         Left            =   -74850
         TabIndex        =   65
         Top             =   1250
         Width           =   5565
         Begin VB.ComboBox cboParentalLeaveAbsType 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   300
            Width           =   2505
         End
         Begin VB.ComboBox cboAbsChildNo 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   705
            Width           =   2505
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parental Leave Absence Type :"
            Height          =   195
            Left            =   150
            TabIndex        =   66
            Top             =   360
            Width           =   2700
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Child No. Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   68
            Top             =   760
            Width           =   1320
         End
      End
      Begin VB.Frame fraSSPColumns 
         Caption         =   "SSP Columns :"
         Enabled         =   0   'False
         Height          =   1980
         Left            =   -74850
         TabIndex        =   39
         Top             =   390
         Width           =   5565
         Begin VB.ComboBox cboSSPPaidDays 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1515
            Width           =   2505
         End
         Begin VB.ComboBox cboSSPWaitingDays 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1110
            Width           =   2505
         End
         Begin VB.ComboBox cboSSPQualifyingDays 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   705
            Width           =   2505
         End
         Begin VB.ComboBox cboSSPApplies 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   300
            Width           =   2505
         End
         Begin VB.Label lblSSPPaidDays 
            BackStyle       =   0  'Transparent
            Caption         =   "SSP Paid Days Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   46
            Top             =   1575
            Width           =   2340
         End
         Begin VB.Label lblSSPWaitingDays 
            BackStyle       =   0  'Transparent
            Caption         =   "SSP Waiting Days Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   44
            Top             =   1170
            Width           =   2535
         End
         Begin VB.Label lblSSPQualifyingDays 
            BackStyle       =   0  'Transparent
            Caption         =   "SSP Qualifying Days Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   42
            Top             =   765
            Width           =   2715
         End
         Begin VB.Label lblSSPApplies 
            BackStyle       =   0  'Transparent
            Caption         =   "SSP Applies Column :"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   360
            Width           =   2340
         End
      End
      Begin VB.Frame fraWorkingDays 
         Caption         =   "SSP Qualifying Days :"
         Enabled         =   0   'False
         Height          =   2700
         Left            =   -74850
         TabIndex        =   48
         Top             =   2430
         Width           =   5565
         Begin VB.Frame fraWorkingPattern 
            Height          =   750
            Left            =   2910
            TabIndex        =   54
            Top             =   1050
            Width           =   1800
            Begin VB.Frame Frame4 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "S"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   90
               TabIndex        =   86
               Top             =   135
               Width           =   1635
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "T"
                  Height          =   240
                  Index           =   1
                  Left            =   575
                  TabIndex        =   93
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "W"
                  Height          =   240
                  Index           =   6
                  Left            =   725
                  TabIndex        =   92
                  Top             =   0
                  Width           =   165
               End
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "S"
                  Height          =   240
                  Index           =   5
                  Left            =   1350
                  TabIndex        =   91
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "T"
                  Height          =   240
                  Index           =   4
                  Left            =   960
                  TabIndex        =   90
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "F"
                  Height          =   240
                  Index           =   3
                  Left            =   1160
                  TabIndex        =   89
                  Top             =   0
                  Width           =   105
               End
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "S"
                  Height          =   240
                  Index           =   2
                  Left            =   135
                  TabIndex        =   88
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label lblQualifyingDays 
                  Caption         =   "M"
                  Height          =   240
                  Index           =   0
                  Left            =   350
                  TabIndex        =   87
                  Top             =   0
                  Width           =   150
               End
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Height          =   195
               Index           =   6
               Left            =   1405
               TabIndex        =   61
               Top             =   385
               Width           =   195
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Height          =   195
               Index           =   5
               Left            =   1200
               TabIndex        =   60
               Top             =   385
               Width           =   195
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Height          =   195
               Index           =   4
               Left            =   1000
               TabIndex        =   59
               Top             =   385
               Width           =   195
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Height          =   195
               Index           =   3
               Left            =   800
               TabIndex        =   58
               Top             =   385
               Width           =   195
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Height          =   195
               Index           =   2
               Left            =   600
               TabIndex        =   57
               Top             =   385
               Width           =   195
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Height          =   195
               Index           =   1
               Left            =   400
               TabIndex        =   56
               Top             =   385
               Width           =   195
            End
            Begin VB.CheckBox chkWorkingPattern 
               Caption         =   "Check1"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   200
               TabIndex        =   55
               Top             =   385
               Width           =   195
            End
         End
         Begin VB.ComboBox cboWorkingDaysNumericValue 
            Height          =   315
            ItemData        =   "frmAbsenceSetup.frx":0098
            Left            =   2910
            List            =   "frmAbsenceSetup.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   700
            Width           =   2500
         End
         Begin VB.ComboBox cboWorkingDaysField 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   300
            Width           =   2505
         End
         Begin VB.OptionButton optWorkingDaysType 
            Caption         =   "Pa&rt-time Days"
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   53
            Top             =   1100
            Width           =   2355
         End
         Begin VB.OptionButton optWorkingDaysType 
            Caption         =   "F&ull-time Days"
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   51
            Top             =   700
            Width           =   1845
         End
         Begin VB.OptionButton optWorkingDaysType 
            Caption         =   "&Field"
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   49
            Top             =   300
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin VB.Frame fraCalendarInclude 
         Caption         =   "Default Options :"
         Enabled         =   0   'False
         Height          =   3840
         Left            =   -74850
         TabIndex        =   33
         Top             =   1290
         Width           =   5565
         Begin VB.CheckBox chkWeekendsShading 
            Caption         =   "Show W&eekends"
            Height          =   315
            Left            =   200
            TabIndex        =   38
            Top             =   1500
            Width           =   2160
         End
         Begin VB.CheckBox chkBankHolidaysShading 
            Caption         =   "Show Bank &Holidays"
            Height          =   315
            Left            =   200
            TabIndex        =   36
            Top             =   900
            Width           =   2460
         End
         Begin VB.CheckBox chkBankHolidaysInclude 
            Caption         =   "&Include Bank Holidays"
            Height          =   315
            Left            =   200
            TabIndex        =   34
            Top             =   300
            Width           =   2535
         End
         Begin VB.CheckBox chkIncludeWorkingDaysOnly 
            Caption         =   "&Working Days Only"
            Height          =   315
            Left            =   200
            TabIndex        =   35
            Top             =   600
            Width           =   2235
         End
         Begin VB.CheckBox chkShowCaptions 
            Caption         =   "Show Calendar Captio&ns"
            Height          =   315
            Left            =   200
            TabIndex        =   37
            Top             =   1200
            Width           =   3240
         End
      End
      Begin VB.Frame fraAbsence 
         Caption         =   "Absence :"
         Height          =   4740
         Left            =   150
         TabIndex        =   0
         Top             =   390
         Width           =   5565
         Begin VB.ComboBox cboContinuous 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   4155
            Width           =   2505
         End
         Begin VB.ComboBox cboAbsenceDuration 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3750
            Width           =   2505
         End
         Begin COALine.COA_Line ASRDummyLine1 
            Height          =   30
            Left            =   180
            Top             =   1155
            Width           =   5200
            _ExtentX        =   9181
            _ExtentY        =   53
         End
         Begin VB.ComboBox cboAbsenceScreen 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   705
            Width           =   2505
         End
         Begin VB.ComboBox cboAbsenceTable 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   300
            Width           =   2505
         End
         Begin VB.ComboBox cboStartDate 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   2505
         End
         Begin VB.ComboBox cboStartSession 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1725
            Width           =   2505
         End
         Begin VB.ComboBox cboEndDate 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2130
            Width           =   2505
         End
         Begin VB.ComboBox cboEndSession 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2535
            Width           =   2505
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   3345
            Width           =   2505
         End
         Begin VB.ComboBox cboReason 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2940
            Width           =   2505
         End
         Begin VB.Label lblContinuousAbsence 
            BackStyle       =   0  'Transparent
            Caption         =   "Continuous Absence Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   4215
            Width           =   2685
         End
         Begin VB.Label lblAbsenceDuration 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Duration Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   3810
            Width           =   2535
         End
         Begin VB.Label lblAbsenceScreen 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Screen :"
            Height          =   195
            Left            =   195
            TabIndex        =   3
            Top             =   765
            Width           =   1680
         End
         Begin VB.Label lblAbsenceTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   1
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label lblStartDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   1380
            Width           =   1830
         End
         Begin VB.Label lblStartSession 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Session Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   1785
            Width           =   2025
         End
         Begin VB.Label lblEndDate 
            BackStyle       =   0  'Transparent
            Caption         =   "End Date Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   2190
            Width           =   1740
         End
         Begin VB.Label lblEndColumn 
            BackStyle       =   0  'Transparent
            Caption         =   "End Session Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   2595
            Width           =   1935
         End
         Begin VB.Label lblType 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Type Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   3405
            Width           =   2100
         End
         Begin VB.Label lblReason 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Reason Column :"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   3000
            Width           =   2280
         End
      End
      Begin VB.Frame fraAbsenceType 
         Caption         =   "Absence Type Table :"
         Enabled         =   0   'False
         Height          =   4740
         Left            =   -74850
         TabIndex        =   84
         Top             =   390
         Width           =   5565
         Begin COALine.COA_Line ASRDummyLine2 
            Height          =   30
            Left            =   165
            Top             =   765
            Width           =   5200
            _ExtentX        =   9181
            _ExtentY        =   53
         End
         Begin VB.ComboBox cboAbsenceTypeTable 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   300
            Width           =   2500
         End
         Begin VB.ComboBox cboAbsenceTypeSSPApplicable 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1755
            Width           =   2500
         End
         Begin VB.ComboBox cboAbsenceTypeCalendarCode 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2160
            Width           =   2500
         End
         Begin VB.ComboBox cboAbsenceTypeType 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   945
            Width           =   2500
         End
         Begin VB.ComboBox cboAbsenceTypeCode 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1350
            Width           =   2500
         End
         Begin VB.Label lblAbsenceTypeTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Type Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblAbsenceTypeSSPApplicable 
            BackStyle       =   0  'Transparent
            Caption         =   "SSP Applicable Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   29
            Top             =   1815
            Width           =   2205
         End
         Begin VB.Label lblAbsenceTypeCalendarCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Calendar Code Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   31
            Top             =   2220
            Width           =   2235
         End
         Begin VB.Label lblAbsenceTypeType 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Type Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   25
            Top             =   1005
            Width           =   2190
         End
         Begin VB.Label lblAbsenceTypeCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Absence Code Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   27
            Top             =   1410
            Width           =   2205
         End
      End
   End
End
Attribute VB_Name = "frmAbsenceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##################################################
'
' DECLARE EVENTS
'
'  Event UnLoad()
'
'##################################################
'
' CONSTANTS
'
' Page number constants.
Private Const giPAGE_ABSENCE = 0
Private Const giPAGE_ABSENCETYPE = 1
Private Const giPAGE_CALENDAR = 2
Private Const giPAGE_SSP = 3

Private Const giWORKDAYSTYPE_FIELD = 0
Private Const giWORKDAYSTYPE_NUMERICVALUE = 1
Private Const giWORKDAYSTYPE_PATTERNVALUE = 2
'
' VARIABLES
'
' Personnel variables.
Private mvar_lngPersonnelTableID As Long
  
' Absence variables.
Private mvar_lngAbsenceTableID As Long
Private mvar_lngAbsenceTypeTableID As Long
Private mvar_lngAbsenceScreenID As Long
Private mvar_lngAbsenceStartDateID As Long
Private mvar_lngAbsenceStartSessionID As Long
Private mvar_lngAbsenceEndDateID As Long
Private mvar_lngAbsenceEndSessionID As Long
Private mvar_lngAbsenceTypeID As Long
Private mvar_lngAbsenceReasonID As Long
Private mvar_lngAbsenceDurationID As Long
Private mvar_lngAbsenceContinuousID As Long
'
Private mvar_lngAbsenceSSPAppliesID As Long
Private mvar_lngAbsenceSSPQualifyingDaysID As Long
Private mvar_lngAbsenceSSPWaitingDaysID As Long
Private mvar_lngAbsenceSSPPaidDaysID As Long
'
Private mvar_iAbsenceWorkingDaysType As Integer
Private mvar_iAbsenceWorkingDaysNumericValue As Integer
Private mvar_sAbsenceWorkingDaysPattern As String
Private mvar_lngAbsenceWorkingDaysID As Long
'
Private mvar_lngAbsenceTypeTypeID As Long
Private mvar_lngAbsenceTypeCodeID As Long
Private mvar_lngAbsenceTypeSSPID As Long
Private mvar_lngAbsenceTypeCalCodeID As Long
'Private mvar_lngAbsenceTypeIncludeID As Long
'Private mvar_lngAbsenceTypeBradfordID As Long
'
Private mvar_lngAbsenceCalStartMonthID As Long
Private mvar_lngAbsenceCalWeekendShading As Integer
Private mvar_lngAbsenceCalBHolShading As Integer
Private mvar_lngAbsenceCalIncludeWorkingDaysOnly As Integer
Private mvar_lngAbsenceCalBHolInclude As Integer
Private mvar_lngAbsenceCalShowCaptions As Integer

Private mvar_lngAbsenceParentRegion As Long
Private mvar_strAbsenceParentLeaveType As String
Private mvar_lngAbsenceChildNoID As Long

Private mvar_lngDependantsTableID As Long
Private mvar_lngDependantsChildNoID As Long
Private mvar_lngDependantsDateOfBirthID As Long
Private mvar_lngDependantsAdoptedDateID As Long
Private mvar_lngDependantsDisabledID As Long

' Regional Working Pattern Variables
Private mvar_lngWorkingPatternID As Long
Private mvar_lngHWorkingPatternTableID As Long
Private mvar_lngHWorkingPatternFieldID As Long
Private mvar_lngHWorkingPatternDateID As Long

Private mvar_lngOriginalDependantsTableID As Long
Private mvar_lngOriginalAbsenceTableID As Long
  
Private mfChanged As Boolean
Private mblnReadOnly As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  cmdOK.Enabled = mfChanged
End Property

Private Sub cboAbsenceDuration_Click()
  With cboAbsenceDuration
    mvar_lngAbsenceDurationID = .ItemData(.ListIndex)
  End With
    
  Changed = True
End Sub

'MH20030908 Fault 6888
'Private Sub cboAbsenceTypeBradford_Click()
'
'  With cboAbsenceTypeBradford
'    mvar_lngAbsenceTypeBradfordID = .ItemData(.ListIndex)
'  End With
'
'End Sub

Private Sub cboContinuous_Click()
  With cboContinuous
    mvar_lngAbsenceContinuousID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboDepAdoptedDate_Click()
  With cboDepAdoptedDate
    mvar_lngDependantsAdoptedDateID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboDepChildNo_Click()
  With cboDepChildNo
    mvar_lngDependantsChildNoID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboDepDateOfBirth_Click()
  With cboDepDateOfBirth
    mvar_lngDependantsDateOfBirthID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboDepDisabled_Click()
  With cboDepDisabled
    mvar_lngDependantsDisabledID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboDependantsTable_Click()
  With cboDependantsTable
    mvar_lngDependantsTableID = .ItemData(.ListIndex)
  End With

  RefreshDependantsControls
  
  Changed = True
End Sub

Private Sub cboParentalLeaveAbsType_Click()
  mvar_strAbsenceParentLeaveType = Trim(cboParentalLeaveAbsType.Text)
    
  Changed = True
End Sub

Private Sub cboParentalLeaveRegion_Click()
  With cboParentalLeaveRegion
    mvar_lngAbsenceParentRegion = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboSSPApplies_Click()
  With cboSSPApplies
    mvar_lngAbsenceSSPAppliesID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboSSPPaidDays_Click()
  With cboSSPPaidDays
    mvar_lngAbsenceSSPPaidDaysID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboSSPQualifyingDays_Click()
  With cboSSPQualifyingDays
    mvar_lngAbsenceSSPQualifyingDaysID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboSSPWaitingDays_Click()
  With cboSSPWaitingDays
    mvar_lngAbsenceSSPWaitingDaysID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboWorkingDaysField_Click()
  With cboWorkingDaysField
    mvar_lngAbsenceWorkingDaysID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboWorkingDaysNumericValue_Click()
  mvar_iAbsenceWorkingDaysNumericValue = val(cboWorkingDaysNumericValue.Text)
  
  Changed = True
End Sub


Private Sub chkWorkingPattern_Click(Index As Integer)
  Dim sPatternString As String
  
  If Len(mvar_sAbsenceWorkingDaysPattern) < 14 Then
    mvar_sAbsenceWorkingDaysPattern = mvar_sAbsenceWorkingDaysPattern & Space(14 - Len(mvar_sAbsenceWorkingDaysPattern))
  End If
  
  Select Case Index
    Case 0 ' Sunday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Ss", "  ")
    Case 1 ' Monday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Mm", "  ")
    Case 2 ' Tuesday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Tt", "  ")
    Case 3 ' Wednesday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Ww", "  ")
    Case 4 ' Thursday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Tt", "  ")
    Case 5 ' Friday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Ff", "  ")
    Case 6 ' Saturday
      sPatternString = IIf(chkWorkingPattern(Index).value = vbChecked, "Ss", "  ")
  End Select

  mvar_sAbsenceWorkingDaysPattern = Mid(mvar_sAbsenceWorkingDaysPattern, 1, (Index * 2)) & _
    sPatternString & Mid(mvar_sAbsenceWorkingDaysPattern, (2 * (Index + 1)) + 1)
  
  Changed = True
End Sub

Private Sub cboAbsChildNo_Click()

  With cboAbsChildNo
    mvar_lngAbsenceChildNoID = .ItemData(.ListIndex)
  End With
  
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

'
'##################################################


'##################################################
'
' FORM LOAD/UNLOAD SAVE CODE ETC
'
'##################################################

Private Sub Form_Load()
  
  Screen.MousePointer = vbHourglass

  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  ' Read the current settings from the database.
  ReadParameters
  
  ' Initialise all controls with the current settings, or defaults.
  InitialiseBaseTableCombos
  
  SSTab1.Tab = giPAGE_ABSENCE
  
  Changed = False
  
  Screen.MousePointer = vbDefault
 
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
        'SaveChanges
        Cancel = (Not SaveChanges)
    End Select
  End If
  
End Sub

'##################################################
'
' FORM INITIALISE CODE ETC
'
'##################################################

Private Sub InitialiseBaseTableCombos()
  
  ' Initialise the Base Table combo(s)
  Dim iAbsenceTableListIndex As Integer
  Dim iDependantsTableListIndex As Integer
  Dim iAbsenceTypeTableListIndex As Integer
  Dim intMonth As Integer
  Dim dtFirstOfDec As Date
  Dim lngNewIndex As Long
    
  iAbsenceTableListIndex = 0
  iDependantsTableListIndex = 0
  iAbsenceTypeTableListIndex = 0

  RefreshPersonnelControls

  ' Clear the combos, and add '<None>' items.
  cboAbsenceTable.Clear
  cboDependantsTable.Clear
  cboAbsenceTypeTable.Clear
  
  AddItemToComboBox cboAbsenceTable, "<None>", 0
  AddItemToComboBox cboDependantsTable, "<None>", 0
  AddItemToComboBox cboAbsenceTypeTable, "<None>", 0
    
  ' Add items to the combo for each table that has not been deleted,
  ' and is a child of the defined Personnel table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted Then
        recRelEdit.Index = "idxParentID"
        recRelEdit.Seek "=", mvar_lngPersonnelTableID, !TableID
        
        If Not recRelEdit.NoMatch Then
        
          lngNewIndex = AddItemToComboBox(cboAbsenceTable, !TableName, !TableID)
          If !TableID = mvar_lngAbsenceTableID Then
            iAbsenceTableListIndex = lngNewIndex
          End If
        
          lngNewIndex = AddItemToComboBox(cboDependantsTable, !TableName, !TableID)
          If !TableID = mvar_lngDependantsTableID Then
            iDependantsTableListIndex = lngNewIndex
          End If

        End If
       
        lngNewIndex = AddItemToComboBox(cboAbsenceTypeTable, !TableName, !TableID)
        If !TableID = mvar_lngAbsenceTypeTableID Then
          iAbsenceTypeTableListIndex = lngNewIndex
        End If
      
      End If
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
  With cboAbsenceTable
    If Not mblnReadOnly Then
      .Enabled = True
      cboDependantsTable.Enabled = True
      cboAbsenceTypeTable.Enabled = True
    End If
    .ListIndex = iAbsenceTableListIndex
    cboDependantsTable.ListIndex = iDependantsTableListIndex
    cboAbsenceTypeTable.ListIndex = iAbsenceTypeTableListIndex
  End With

  
  'MH20060613 Fault 10832
  ' Fill the CalStart Combo with Months
  dtFirstOfDec = DateAdd("m", Month(Now) * -1, DateAdd("d", (Day(Now) - 1) * -1, Now))
  
  For intMonth = 1 To 12
    AddItemToComboBox cboCalStartMonth, Format(DateAdd("m", intMonth, dtFirstOfDec), "mmmm"), intMonth
  Next intMonth

  ' Set January as default
  cboCalStartMonth.ListIndex = mvar_lngAbsenceCalStartMonthID - 1
  
End Sub



'##################################################
'
' CODE TO SET MODULE LEVEL VARIABLES WHEN USER CLICKS COMBO ETC
'
'##################################################


Private Sub cboAbsenceTable_Click()

  With cboAbsenceTable
    mvar_lngAbsenceTableID = .ItemData(.ListIndex)
  End With

  RefreshAbsenceControls
  
  Changed = True
End Sub

Private Sub cboAbsenceTypeTable_Click()

  With cboAbsenceTypeTable
    mvar_lngAbsenceTypeTableID = .ItemData(.ListIndex)
  End With

  RefreshAbsenceTypeControls
  RefreshAbsenceTypeValues
  
  Changed = True
End Sub


Private Sub cboAbsenceScreen_Click()

  With cboAbsenceScreen
    mvar_lngAbsenceScreenID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboStartDate_Click()

  With cboStartDate
    mvar_lngAbsenceStartDateID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboStartSession_Click()

  With cboStartSession
    mvar_lngAbsenceStartSessionID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboEndDate_Click()

  With cboEndDate
    mvar_lngAbsenceEndDateID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboEndSession_Click()

  With cboEndSession
    mvar_lngAbsenceEndSessionID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboType_Click()

  With cboType
    mvar_lngAbsenceTypeID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboReason_Click()

  With cboReason
    mvar_lngAbsenceReasonID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub


Private Sub cboAbsenceTypeType_Click()

  With cboAbsenceTypeType
    mvar_lngAbsenceTypeTypeID = .ItemData(.ListIndex)
  End With

  RefreshAbsenceTypeValues
  
  Changed = True
End Sub

Private Sub cboAbsenceTypeCode_Click()

  With cboAbsenceTypeCode
    mvar_lngAbsenceTypeCodeID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboAbsenceTypeSSPApplicable_Click()

  With cboAbsenceTypeSSPApplicable
    mvar_lngAbsenceTypeSSPID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboAbsenceTypeCalendarCode_Click()

  With cboAbsenceTypeCalendarCode
    mvar_lngAbsenceTypeCalCodeID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

'Private Sub cboAbsenceTypeInclude_Click()
'
'  With cboAbsenceTypeInclude
'    mvar_lngAbsenceTypeIncludeID = .ItemData(.ListIndex)
'  End With
'
'End Sub

Private Sub cboCalStartMonth_Click()

  With cboCalStartMonth
    mvar_lngAbsenceCalStartMonthID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub chkWeekendsShading_Click()

mvar_lngAbsenceCalWeekendShading = (chkWeekendsShading.value = vbChecked)
  
  Changed = True
End Sub

Private Sub chkShowCaptions_Click()

mvar_lngAbsenceCalShowCaptions = (chkShowCaptions.value = vbChecked)
  
  Changed = True
End Sub

Private Sub chkBankHolidaysShading_Click()

mvar_lngAbsenceCalBHolShading = (chkBankHolidaysShading.value = vbChecked)
  
  Changed = True
End Sub

Private Sub chkIncludeWorkingDaysOnly_Click()

mvar_lngAbsenceCalIncludeWorkingDaysOnly = (chkIncludeWorkingDaysOnly.value = vbChecked)
  
  Changed = True
End Sub

Private Sub chkBankHolidaysInclude_Click()
  
mvar_lngAbsenceCalBHolInclude = (chkBankHolidaysInclude.value = vbChecked)
    
  Changed = True
End Sub


'##################################################

Private Sub RefreshAbsenceControls()

  ' Refresh the Absence controls.
  Dim iAbsenceScreenListIndex As Integer
  Dim iAbsenceStartDateListIndex As Integer
  Dim iAbsenceStartSessionListIndex As Integer
  Dim iAbsenceEndDateListIndex As Integer
  Dim iAbsenceEndSessionListIndex As Integer
  Dim iAbsenceTypeListIndex As Integer
  Dim iAbsenceReasonListIndex As Integer
  Dim iAbsenceDurationListIndex As Integer
  Dim iAbsenceContinuousListIndex As Integer
  Dim iAbsenceSSPAppliesListIndex As Integer
  Dim iAbsenceSSPQualifyingDaysListIndex As Integer
  Dim iAbsenceSSPWaitingDaysListIndex As Integer
  Dim iAbsenceSSPPaidDaysListIndex As Integer
  Dim iAbsenceChildNoListIndex As Integer
  Dim objctl As Control
  Dim lngNewIndex As Long

  iAbsenceScreenListIndex = 0
  iAbsenceStartDateListIndex = 0
  iAbsenceStartSessionListIndex = 0
  iAbsenceEndDateListIndex = 0
  iAbsenceEndSessionListIndex = 0
  iAbsenceTypeListIndex = 0
  iAbsenceReasonListIndex = 0
  iAbsenceSSPAppliesListIndex = 0
  iAbsenceSSPQualifyingDaysListIndex = 0
  iAbsenceSSPWaitingDaysListIndex = 0
  iAbsenceSSPPaidDaysListIndex = 0
  iAbsenceDurationListIndex = 0
  iAbsenceChildNoListIndex = 0
  
  ' Clear the current contents of the combos.
  For Each objctl In Me
    If TypeOf objctl Is ComboBox And _
      (objctl.Name = "cboStartDate" Or _
      objctl.Name = "cboStartSession" Or _
      objctl.Name = "cboEndDate" Or _
      objctl.Name = "cboEndSession" Or _
      objctl.Name = "cboType" Or _
      objctl.Name = "cboReason" Or _
      objctl.Name = "cboAbsenceDuration" Or _
      objctl.Name = "cboContinuous" Or _
      objctl.Name = "cboSSPApplies" Or _
      objctl.Name = "cboSSPQualifyingDays" Or _
      objctl.Name = "cboSSPWaitingDays" Or _
      objctl.Name = "cboSSPPaidDays" Or _
      objctl.Name = "cboParentalLeaveAbsType" Or _
      objctl.Name = "cboAbsChildNo") Then

      objctl.Clear
      AddItemToComboBox objctl, "<None>", 0
      
    End If
  Next objctl
  Set objctl = Nothing
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngAbsenceTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngAbsenceTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          ' Only load date fields into the start/leaving date combos
          If !DataType = dtTIMESTAMP Then
          
            lngNewIndex = AddItemToComboBox(cboStartDate, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceStartDateID Then
              iAbsenceStartDateListIndex = lngNewIndex
            End If

            lngNewIndex = AddItemToComboBox(cboEndDate, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceEndDateID Then
              iAbsenceEndDateListIndex = lngNewIndex
            End If
          End If

          ' Only load numeric calc columns into the Duration Combo
          If !DataType = dtNUMERIC And !columntype = 2 Then
            
            lngNewIndex = AddItemToComboBox(cboAbsenceDuration, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceDurationID Then
              iAbsenceDurationListIndex = lngNewIndex
            End If

          End If

          ' Only load logic fields into the SSP Applies, Continuous Absence combo
          If !DataType = dtBIT Then
            
            lngNewIndex = AddItemToComboBox(cboSSPApplies, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceSSPAppliesID Then
              iAbsenceSSPAppliesListIndex = lngNewIndex
            End If
           
            lngNewIndex = AddItemToComboBox(cboContinuous, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceContinuousID Then
             iAbsenceContinuousListIndex = lngNewIndex
            End If

          End If

          ' Only load numeric fields into the SSP Qualifying Days, Waiting Days & Paid Days combos.
          If !DataType = dtNUMERIC Then
            lngNewIndex = AddItemToComboBox(cboSSPQualifyingDays, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceSSPQualifyingDaysID Then
              iAbsenceSSPQualifyingDaysListIndex = lngNewIndex
            End If
          
            lngNewIndex = AddItemToComboBox(cboSSPWaitingDays, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceSSPWaitingDaysID Then
              iAbsenceSSPWaitingDaysListIndex = lngNewIndex
            End If
                     
            lngNewIndex = AddItemToComboBox(cboSSPPaidDays, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceSSPPaidDaysID Then
              iAbsenceSSPPaidDaysListIndex = lngNewIndex
            End If
          
          End If
          
          
          If !DataType = dtinteger Then
            
            lngNewIndex = AddItemToComboBox(cboAbsChildNo, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceChildNoID Then
              iAbsenceChildNoListIndex = lngNewIndex
            End If

          End If

          ' Load varchar fields
          If !DataType = dtVARCHAR Then
            If !Size = 2 Then
             
              lngNewIndex = AddItemToComboBox(cboStartSession, !ColumnName, !ColumnID)
              If !ColumnID = mvar_lngAbsenceStartSessionID Then
                iAbsenceStartSessionListIndex = lngNewIndex
              End If
                        
              lngNewIndex = AddItemToComboBox(cboEndSession, !ColumnName, !ColumnID)
              If !ColumnID = mvar_lngAbsenceEndSessionID Then
                iAbsenceEndSessionListIndex = lngNewIndex
              End If
            
            End If
            
            If !Size = 6 Then
              lngNewIndex = AddItemToComboBox(cboType, !ColumnName, !ColumnID)
              If !ColumnID = mvar_lngAbsenceTypeID Then
                iAbsenceTypeListIndex = lngNewIndex
              End If
            End If
                       
            lngNewIndex = AddItemToComboBox(cboReason, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceReasonID Then
              iAbsenceReasonListIndex = lngNewIndex
            End If
          
          End If
          
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Now do the absence screen combo
  cboAbsenceScreen.Clear
  AddItemToComboBox cboAbsenceScreen, "<None>", 0
  
  With recScrEdit
    .Index = "idxName"
    .MoveFirst
    Do Until .EOF
      If (!TableID = mvar_lngAbsenceTableID) And (!Deleted = False) Then
        lngNewIndex = AddItemToComboBox(cboAbsenceScreen, !Name, !ScreenID)
        If !ScreenID = mvar_lngAbsenceScreenID Then
          iAbsenceScreenListIndex = lngNewIndex
        End If
      End If
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
  cboAbsenceScreen.ListIndex = iAbsenceScreenListIndex
  cboStartDate.ListIndex = iAbsenceStartDateListIndex
  cboStartSession.ListIndex = iAbsenceStartSessionListIndex
  cboEndDate.ListIndex = iAbsenceEndDateListIndex
  cboEndSession.ListIndex = iAbsenceEndSessionListIndex
  cboType.ListIndex = iAbsenceTypeListIndex
  cboReason.ListIndex = iAbsenceReasonListIndex
  cboAbsenceDuration.ListIndex = iAbsenceDurationListIndex
  cboContinuous.ListIndex = iAbsenceContinuousListIndex
  cboSSPApplies.ListIndex = iAbsenceSSPAppliesListIndex
  cboSSPQualifyingDays.ListIndex = iAbsenceSSPQualifyingDaysListIndex
  cboSSPWaitingDays.ListIndex = iAbsenceSSPWaitingDaysListIndex
  cboSSPPaidDays.ListIndex = iAbsenceSSPPaidDaysListIndex
  cboAbsChildNo.ListIndex = iAbsenceChildNoListIndex

End Sub

Private Sub RefreshAbsencePersonnelControls()
  
  ' Refresh the SSP Working Days Field Combo Box (Refresh the Absence/Personnel controls.)
  
  Dim iWorkingDaysListIndex As Integer
  Dim sAbsenceTableName As String
  Dim sPersonnelTableName As String
  Dim lngNewIndex As Long
  Dim objctl As Control
  
  iWorkingDaysListIndex = 0
  
  ' Clear the current contents of the combos.
  For Each objctl In Me
    If TypeOf objctl Is ComboBox And _
      (objctl.Name = "cboWorkingDaysField") Then

      With objctl
        .Clear
        AddItemToComboBox objctl, "<None>", 0
      End With
    End If
  Next objctl
  Set objctl = Nothing
  
  ' Read the historic wpattern stuff
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", mvar_lngAbsenceTableID 'mvar_lngHWorkingPatternTableID
    
    If Not .NoMatch Then
      sAbsenceTableName = !TableName
    End If
    
    .Seek "=", mvar_lngPersonnelTableID
    
    If Not .NoMatch Then
      sPersonnelTableName = !TableName
    End If
  End With
  
  
  With recColEdit
    
    ' Dont do the absence table...read which historic table is being used from the module setup and use that
    .Index = "idxName"
    .Seek ">=", mvar_lngAbsenceTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
'        If !TableID <> mvar_lngHWorkingPatternTableID Then
        If !TableID <> mvar_lngAbsenceTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          ' Only load integer and working pattern fields into the combo
          If (!DataType = dtinteger) Or (!DataType = dtlongvarchar) Then
            lngNewIndex = AddItemToComboBox(cboWorkingDaysField, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceWorkingDaysID Then
              iWorkingDaysListIndex = lngNewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  
    .Seek ">=", mvar_lngPersonnelTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngPersonnelTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          ' Only load integer and working pattern fields into the start/leaving date combos
          If (!DataType = dtinteger) Or (!DataType = dtlongvarchar) Then
            lngNewIndex = AddItemToComboBox(cboWorkingDaysField, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceWorkingDaysID Then
              iWorkingDaysListIndex = lngNewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboWorkingDaysField.ListIndex = iWorkingDaysListIndex

End Sub

Private Sub RefreshWorkingDaysControls(piWorkingDaysType As Integer)
  ' Enable/Disable the required working days controls.
  Dim fField As Boolean
  Dim fNumericValue As Boolean
  Dim fPatternValue As Boolean
  Dim ctlCheckBox As CheckBox
  
  fField = (piWorkingDaysType = 0)
  fNumericValue = (piWorkingDaysType = 1)
  fPatternValue = (piWorkingDaysType = 2)
  
  
  If Not mblnReadOnly Then
    cboWorkingDaysField.Enabled = fField
    cboWorkingDaysField.BackColor = IIf(fField, vbWindowBackground, vbButtonFace)
  
    cboWorkingDaysNumericValue.Enabled = fNumericValue
    cboWorkingDaysNumericValue.BackColor = IIf(fNumericValue, vbWindowBackground, vbButtonFace)
  
    fraWorkingPattern.Enabled = fPatternValue
  End If

  If Not fField Then
    If cboWorkingDaysField.Visible Then
      cboWorkingDaysField.ListIndex = 0
    End If
  End If
  
  If Not fNumericValue Then
    cboWorkingDaysNumericValue.ListIndex = 0
  End If
  
  For Each ctlCheckBox In chkWorkingPattern
    If Not fPatternValue Then
      ctlCheckBox.value = vbUnchecked
    End If
    'ctlCheckBox.Enabled = fPatternValue
    ctlCheckBox.Enabled = (fPatternValue And Not mblnReadOnly)
  Next ctlCheckBox
  Set ctlCheckBox = Nothing
  
End Sub


Private Sub RefreshAbsenceTypeControls()
  ' Refresh the Absence Type controls.
  Dim iAbsenceTypeTypeListIndex As Integer
  Dim iAbsenceTypeCodeListIndex As Integer
  Dim iAbsenceTypeSSPApplicableListIndex As Integer
  Dim iAbsenceTypeCalendarCodeListIndex As Integer
  Dim objctl As Control
  Dim lngNewIndex As Long

  iAbsenceTypeTypeListIndex = 0
  iAbsenceTypeCodeListIndex = 0
  iAbsenceTypeSSPApplicableListIndex = 0
  iAbsenceTypeCalendarCodeListIndex = 0

  ' Clear the current contents of the combos.
  For Each objctl In Me
    If TypeOf objctl Is ComboBox And _
      (objctl.Name = "cboAbsenceTypeType" Or _
      objctl.Name = "cboAbsenceTypeCode" Or _
      objctl.Name = "cboAbsenceTypeSSPApplicable" Or _
      objctl.Name = "cboAbsenceTypeCalendarCode" Or _
      objctl.Name = "cboParentalLeaveAbsType") Then
        
      With objctl
        .Clear
        AddItemToComboBox objctl, "<None>", 0
      End With
    End If
  Next objctl

  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngAbsenceTypeTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted, or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngAbsenceTypeTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          ' Load Logic fields
          If !DataType = dtBIT Then
            lngNewIndex = AddItemToComboBox(cboAbsenceTypeSSPApplicable, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceTypeSSPID Then
              iAbsenceTypeSSPApplicableListIndex = lngNewIndex
            End If
          End If
          
          ' Load varchar fields
          If !DataType = dtVARCHAR Then
            If !Size = 6 Then
              lngNewIndex = AddItemToComboBox(cboAbsenceTypeType, !ColumnName, !ColumnID)
              If !ColumnID = mvar_lngAbsenceTypeTypeID Then
                iAbsenceTypeTypeListIndex = lngNewIndex
              End If
            End If
            
            lngNewIndex = AddItemToComboBox(cboAbsenceTypeCode, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceTypeCodeID Then
              iAbsenceTypeCodeListIndex = lngNewIndex
            End If

            lngNewIndex = AddItemToComboBox(cboAbsenceTypeCalendarCode, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceTypeCalCodeID Then
              iAbsenceTypeCalendarCodeListIndex = lngNewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboAbsenceTypeType.ListIndex = iAbsenceTypeTypeListIndex
  cboAbsenceTypeCode.ListIndex = iAbsenceTypeCodeListIndex
  cboAbsenceTypeSSPApplicable.ListIndex = iAbsenceTypeSSPApplicableListIndex
  cboAbsenceTypeCalendarCode.ListIndex = iAbsenceTypeCalendarCodeListIndex

End Sub




'##################################################
'
' COMMAND BUTTONS AND ACTIONS
'
'##################################################

Private Sub cmdOK_Click()

  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If
  
End Sub

Private Function ValidateSetup() As Boolean
  'JPD 20040106 Fault 7894
  
  On Error GoTo ValidateError
  
  Dim fSpecialFunctionUsed As Boolean
  Dim sSQL As String
  Dim rsCheck As DAO.Recordset
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim objExpr As CExpression
  Dim sExprName As String
  Dim sExprType As String
  Dim sExprParentTable As String
  Dim lngExprBaseTableID As Long
  Dim sFunctionName As String
  Dim lngFunctionID As Long
  Dim objFunctionDef As clsFunctionDef
  
  ' Don't allow the Dependants table to change if the special
  ' functions are used anywhere.
  fSpecialFunctionUsed = False
  
  If (mvar_lngOriginalDependantsTableID <> mvar_lngDependantsTableID) Then
    ' Find any expression field components that use the special functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (62,63)"
    Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    If Not (rsCheck.BOF And rsCheck.EOF) Then
      fSpecialFunctionUsed = True

      Set objComp = New CExprComponent
      objComp.ComponentID = rsCheck!ComponentID
      lngExprID = objComp.RootExpressionID
      Set objComp = Nothing

      ' Get the expression name and type description.
      Set objExpr = New CExpression
      objExpr.ExpressionID = lngExprID

      If objExpr.ReadExpressionDetails Then
        sExprName = objExpr.Name
        sExprType = objExpr.ExpressionTypeName
        lngExprBaseTableID = objExpr.BaseTableID

        ' Get the expression's parent table name.
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", lngExprBaseTableID

        If Not recTabEdit.NoMatch Then
          sExprParentTable = recTabEdit!TableName
        End If
      Else
        sExprName = "<unknown>"
        sExprType = "<unknown>"
        sExprParentTable = "<unknown>"
      End If

      ' Disassociate object variables.
      Set objExpr = Nothing
    
      gobjFunctionDefs.Initialise
      sFunctionName = "<unknown>"
      If gobjFunctionDefs.IsValidID(rsCheck!FunctionID) Then
        Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim(Str(rsCheck!FunctionID)))
        sFunctionName = objFunctionDef.Name
        Set objFunctionDef = Nothing
      End If
    End If
    ' Close the recordset.
    rsCheck.Close

    If fSpecialFunctionUsed Then
      MsgBox "The 'Dependants' table cannot be changed." & vbCrLf & vbCrLf & _
        "It is used as the base table for the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If


   ' Don't allow the Absence table to change if the special
  ' functions are used anywhere.

  If (mvar_lngOriginalAbsenceTableID <> mvar_lngAbsenceTableID) Then
    ' Find any expression field components that use the special functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (30,47,73)"
    Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    If Not (rsCheck.BOF And rsCheck.EOF) Then
      fSpecialFunctionUsed = True

      Set objComp = New CExprComponent
      objComp.ComponentID = rsCheck!ComponentID
      lngExprID = objComp.RootExpressionID
      Set objComp = Nothing

      ' Get the expression name and type description.
      Set objExpr = New CExpression
      objExpr.ExpressionID = lngExprID

      If objExpr.ReadExpressionDetails Then
        sExprName = objExpr.Name
        sExprType = objExpr.ExpressionTypeName
        lngExprBaseTableID = objExpr.BaseTableID

        ' Get the expression's parent table name.
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", lngExprBaseTableID

        If Not recTabEdit.NoMatch Then
          sExprParentTable = recTabEdit!TableName
        End If
      Else
        sExprName = "<unknown>"
        sExprType = "<unknown>"
        sExprParentTable = "<unknown>"
      End If

      ' Disassociate object variables.
      Set objExpr = Nothing
    
      gobjFunctionDefs.Initialise
      sFunctionName = "<unknown>"
      If gobjFunctionDefs.IsValidID(rsCheck!FunctionID) Then
        Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim(Str(rsCheck!FunctionID)))
        sFunctionName = objFunctionDef.Name
        Set objFunctionDef = Nothing
      End If
    End If
    ' Close the recordset.
    rsCheck.Close

    If fSpecialFunctionUsed Then
      MsgBox "The 'Absence' table cannot be changed." & vbCrLf & vbCrLf & _
        "It is used as the base table for the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If

  ValidateSetup = True
  Exit Function
  
ValidateError:
  
  MsgBox "Error validating the module setup." & vbCrLf & _
         Err.Description, vbExclamation + vbOKOnly, App.Title
  ValidateSetup = False

End Function


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
'        'Me.MousePointer = vbNormal
'        Screen.MousePointer = vbDefault
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:

  UnLoad Me
End Sub

Private Function SaveChanges() As Boolean
  'AE20071119 Fault #12607
  SaveChanges = False
  
  'Check Start and End Dates and Sessions are not the same as this will cause an error at save time
  If cboEndSession.Text = cboStartSession.Text Then
    MsgBox "The Start Session Column and the End Session Column are the same value.", vbOKOnly + vbInformation, App.ProductName
    Exit Function
  End If
  If cboEndDate.Text = cboStartDate.Text Then
    MsgBox "The Start Date Column and the End Date Column are the same value.", vbOKOnly + vbInformation, App.ProductName
    Exit Function
  End If
  
  If Not ValidateSetup Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  
  ' Write the parameter values to the local database.

  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Save the Absence table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngAbsenceTableID
    .Update

    ' Save the Absence Type table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETYPETABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngAbsenceTypeTableID
    .Update

    ' Save the Absence Screen ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESCREEN
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESCREEN
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_SCREENID
    !parametervalue = mvar_lngAbsenceScreenID
    .Update

    ' Save the Absence Start Date column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESTARTDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceStartDateID
    .Update

    ' Save the Absence Start Session column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESTARTSESSION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceStartSessionID
    .Update

    ' Save the Absence End Date column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEENDDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceEndDateID
    .Update

    ' Save the Absence End Session column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEENDSESSION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceEndSessionID
    .Update

    ' Save the Absence Type column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETYPE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceTypeID
    .Update

    ' Save the Absence Reason column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEREASON
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceReasonID
    .Update

    ' Save the Absence Duration column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEDURATION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceDurationID
    .Update

    ' Save the Absence Continuous column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECONTINUOUS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceContinuousID
    .Update

    ' Save the Absence SSP Applies column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPAPPLIES
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESSPAPPLIES
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceSSPAppliesID
    .Update

    ' Save the Absence SSP Qualifying Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPQUALIFYINGDAYS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESSPQUALIFYINGDAYS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceSSPQualifyingDaysID
    .Update

    ' Save the Absence SSP Waiting Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPWAITINGDAYS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESSPWAITINGDAYS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceSSPWaitingDaysID
    .Update

    ' Save the Absence SSP Paid Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPPAIDDAYS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCESSPPAIDDAYS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceSSPPaidDaysID
    .Update
    
    ' Save the Absence Working Days type.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSTYPE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEWORKINGDAYSTYPE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    Select Case mvar_iAbsenceWorkingDaysType
      Case giWORKDAYSTYPE_NUMERICVALUE
        !parametervalue = 0
      Case giWORKDAYSTYPE_PATTERNVALUE
        !parametervalue = 1
      Case Else
        If mvar_lngAbsenceWorkingDaysID > 0 Then
          recColEdit.Index = "idxColumnID"
          recColEdit.Seek "=", mvar_lngAbsenceWorkingDaysID
  
          If recColEdit.NoMatch Then
            !parametervalue = 2
          Else
            If recColEdit!DataType = dtlongvarchar Then
              !parametervalue = 3
            Else
              !parametervalue = 2
            End If
          End If
        Else
          !parametervalue = 2
        End If
    End Select
    .Update
    
    ' Save the Absence Working Days numeric value.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSNUMERICVALUE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEWORKINGDAYSNUMERICVALUE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_iAbsenceWorkingDaysNumericValue
    .Update

    ' Save the Absence Working Days pattern value.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSPATTERNVALUE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEWORKINGDAYSPATTERNVALUE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_sAbsenceWorkingDaysPattern
    .Update

    ' Save the Absence Working Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSFIELD
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEWORKINGDAYSFIELD
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceWorkingDaysID
    .Update
    
    ' Save the AbsenceType Type column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETYPE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETYPETYPE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceTypeTypeID
    .Update

    ' Save the AbsenceType Code column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECODE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETYPECODE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceTypeCodeID
    .Update

    ' Save the AbsenceType SSP column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPESSP
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETYPESSP
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceTypeSSPID
    .Update

    ' Save the AbsenceType CalCode ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECALCODE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCETYPECALCODE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceTypeCalCodeID
    .Update

'MH20030908 Fault 6888
'    ' Save the AbsenceType Include column ID.
'    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEINCLUDE
'    If .NoMatch Then
'      .AddNew
'      !moduleKey = gsMODULEKEY_ABSENCE
'      !parameterkey = gsPARAMETERKEY_ABSENCETYPEINCLUDE
'    Else
'      .Edit
'    End If
'    !ParameterType = gsPARAMETERTYPE_COLUMNID
'    !parametervalue = mvar_lngAbsenceTypeIncludeID
'    .Update
'
'    ' Save the AbsenceType Include in Bradford column ID.
'    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX
'    If .NoMatch Then
'      .AddNew
'      !moduleKey = gsMODULEKEY_ABSENCE
'      !parameterkey = gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX
'    Else
'      .Edit
'    End If
'    !ParameterType = gsPARAMETERTYPE_COLUMNID
'    !parametervalue = mvar_lngAbsenceTypeBradfordID
'    .Update

    ' Save the Absence Cal Start Month ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALSTARTMONTH
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECALSTARTMONTH
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceCalStartMonthID
    .Update

    ' Save the Absence Cal Weekend Shading Default
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALWEEKENDSHADING
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECALWEEKENDSHADING
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_lngAbsenceCalWeekendShading, "TRUE", "FALSE")
    .Update

    ' Save the Absence Cal Show Captions Default
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_lngAbsenceCalShowCaptions, "TRUE", "FALSE")
    .Update

    ' Save the Absence Cal BHol Shading Default
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALBHOLSHADING
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECALBHOLSHADING
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_lngAbsenceCalBHolShading, "TRUE", "FALSE")
    .Update

    ' Save the Absence Cal Include Working Days Only Default
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_lngAbsenceCalIncludeWorkingDaysOnly, "TRUE", "FALSE")
    .Update

    ' Save the Absence Cal BHol Include Default
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALBHOLINCLUDE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECALBHOLINCLUDE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = IIf(mvar_lngAbsenceCalBHolInclude, "TRUE", "FALSE")
    .Update


    'Save Parental Leave Absence Type
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEPARENTALLEAVETYPE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEPARENTALLEAVETYPE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_OTHER
    !parametervalue = mvar_strAbsenceParentLeaveType
    .Update

    'Save Absence Child No Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECHILDNO
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCECHILDNO
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceChildNoID
    .Update
    
    
    'Save Parental Leave Region
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEPARENTALREGION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_ABSENCEPARENTALREGION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngAbsenceParentRegion
    .Update

    'Save Dependants table
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_DEPENDANTSTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngDependantsTableID
    .Update
    
    'Save Dependants Child No Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSCHILDNO
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_DEPENDANTSCHILDNO
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngDependantsChildNoID
    .Update
    
    'Save Dependants Date Of Birth Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSDATEOFBIRTH
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_DEPENDANTSDATEOFBIRTH
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngDependantsDateOfBirthID
    .Update
    
    'Save Adopted Date Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSADOPTEDDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_DEPENDANTSADOPTEDDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngDependantsAdoptedDateID
    .Update
    
    'Save Dependants Disabled Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSDISABLED
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_DEPENDANTSDISABLED
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngDependantsDisabledID
    .Update
  
  
  End With

  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault

End Function


Private Sub ReadParameters()
  ' Read the Absence parameter values from the database into local variables.
  Dim iTemp As Integer
  Dim iLoop As Integer
  
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Personnel table ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If .NoMatch Then
      mvar_lngPersonnelTableID = 0
    Else
      mvar_lngPersonnelTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    If mvar_lngPersonnelTableID = 0 Then
      MsgBox "The Personnel module has not been configured." & vbCrLf & _
        "The Absence module requires the Personnel table to be defined.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
    
    ' Get the Absence table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE
    If .NoMatch Then
      mvar_lngAbsenceTableID = 0
    Else
      mvar_lngAbsenceTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Type table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETABLE
    If .NoMatch Then
      mvar_lngAbsenceTypeTableID = 0
    Else
      mvar_lngAbsenceTypeTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Screen ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESCREEN
    If .NoMatch Then
      mvar_lngAbsenceScreenID = 0
    Else
      mvar_lngAbsenceScreenID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Start Date column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE
    If .NoMatch Then
      mvar_lngAbsenceStartDateID = 0
    Else
      mvar_lngAbsenceStartDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Start Session column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION
    If .NoMatch Then
      mvar_lngAbsenceStartSessionID = 0
    Else
      mvar_lngAbsenceStartSessionID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence End Date column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE
    If .NoMatch Then
      mvar_lngAbsenceEndDateID = 0
    Else
      mvar_lngAbsenceEndDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence End Session column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION
    If .NoMatch Then
      mvar_lngAbsenceEndSessionID = 0
    Else
      mvar_lngAbsenceEndSessionID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Type column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE
    If .NoMatch Then
      mvar_lngAbsenceTypeID = 0
    Else
      mvar_lngAbsenceTypeID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Reason column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON
    If .NoMatch Then
      mvar_lngAbsenceReasonID = 0
    Else
      mvar_lngAbsenceReasonID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Duration column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION
    If .NoMatch Then
      mvar_lngAbsenceDurationID = 0
    Else
      mvar_lngAbsenceDurationID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence Continuous column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS
    If .NoMatch Then
      mvar_lngAbsenceContinuousID = 0
    Else
      mvar_lngAbsenceContinuousID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    ' Get the Absence SSP Applies column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPAPPLIES
    If .NoMatch Then
      mvar_lngAbsenceSSPAppliesID = 0
    Else
      mvar_lngAbsenceSSPAppliesID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence SSP Qualifying Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPQUALIFYINGDAYS
    If .NoMatch Then
      mvar_lngAbsenceSSPQualifyingDaysID = 0
    Else
      mvar_lngAbsenceSSPQualifyingDaysID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence SSP Waiting Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPWAITINGDAYS
    If .NoMatch Then
      mvar_lngAbsenceSSPWaitingDaysID = 0
    Else
      mvar_lngAbsenceSSPWaitingDaysID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Absence SSP Paid Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPPAIDDAYS
    If .NoMatch Then
      mvar_lngAbsenceSSPPaidDaysID = 0
    Else
      mvar_lngAbsenceSSPPaidDaysID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Working Days type.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSTYPE
    If .NoMatch Then
      iTemp = 2
    Else
      iTemp = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 2, !parametervalue)
    End If
    Select Case iTemp
      Case 0
        mvar_iAbsenceWorkingDaysType = giWORKDAYSTYPE_NUMERICVALUE
      Case 1
        mvar_iAbsenceWorkingDaysType = giWORKDAYSTYPE_PATTERNVALUE
      Case Else
        mvar_iAbsenceWorkingDaysType = giWORKDAYSTYPE_FIELD
    End Select
    optWorkingDaysType(mvar_iAbsenceWorkingDaysType).value = True
    
    ' Get the Working Days numeric value.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSNUMERICVALUE
    If .NoMatch Then
      mvar_iAbsenceWorkingDaysNumericValue = 5
    Else
      mvar_iAbsenceWorkingDaysNumericValue = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 5, !parametervalue)
    End If
    cboWorkingDaysNumericValue.Text = mvar_iAbsenceWorkingDaysNumericValue

    ' Get the Working Days pattern value.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSPATTERNVALUE
    If .NoMatch Then
      mvar_sAbsenceWorkingDaysPattern = ""
    Else
      mvar_sAbsenceWorkingDaysPattern = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, "", !parametervalue)
    End If
    For iLoop = 0 To 6
      If Len(mvar_sAbsenceWorkingDaysPattern) >= (2 * (iLoop + 1)) Then
        If (Mid(mvar_sAbsenceWorkingDaysPattern, (2 * iLoop) + 1, 1) <> " ") And _
          (Mid(mvar_sAbsenceWorkingDaysPattern, 2 * (iLoop + 1), 1) <> " ") Then
      
          chkWorkingPattern(iLoop) = vbChecked
        End If
      End If
    Next iLoop
    
    ' Get the Working Days column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSFIELD
    If .NoMatch Then
      mvar_lngAbsenceWorkingDaysID = 0
    Else
      mvar_lngAbsenceWorkingDaysID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    RefreshAbsencePersonnelControls
    RefreshWorkingDaysControls mvar_iAbsenceWorkingDaysType

    ' Get the AbsenceType Type column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETYPE
    If .NoMatch Then
      mvar_lngAbsenceTypeTypeID = 0
    Else
      mvar_lngAbsenceTypeTypeID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the AbsenceType Code column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECODE
    If .NoMatch Then
      mvar_lngAbsenceTypeCodeID = 0
    Else
      mvar_lngAbsenceTypeCodeID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the AbsenceType SSP column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPESSP
    If .NoMatch Then
      mvar_lngAbsenceTypeSSPID = 0
    Else
      mvar_lngAbsenceTypeSSPID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the AbsenceType CalCode column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECALCODE
    If .NoMatch Then
      mvar_lngAbsenceTypeCalCodeID = 0
    Else
      mvar_lngAbsenceTypeCalCodeID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

'MH20030908 Fault 6888
'    ' Get the AbsenceType Include column ID.
'    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEINCLUDE
'    If .NoMatch Then
'      mvar_lngAbsenceTypeIncludeID = 0
'    Else
'      mvar_lngAbsenceTypeIncludeID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
'    End If
'
'    ' Get the AbsenceType Bradford Index column ID.
'    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX
'    If .NoMatch Then
'      mvar_lngAbsenceTypeBradfordID = 0
'    Else
'      mvar_lngAbsenceTypeBradfordID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
'    End If

    ' Get the AbsenceCal StartMonth ID (1-12)
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALSTARTMONTH
    If .NoMatch Then
      mvar_lngAbsenceCalStartMonthID = 0
    Else
      mvar_lngAbsenceCalStartMonthID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the AbsenceCal Weekend Shading
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALWEEKENDSHADING
    If .NoMatch Then
      mvar_lngAbsenceCalWeekendShading = False
    Else
      mvar_lngAbsenceCalWeekendShading = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkWeekendsShading.value = IIf(mvar_lngAbsenceCalWeekendShading, vbChecked, vbUnchecked)

    ' Get the AbsenceCal Show Captions
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS
    If .NoMatch Then
      mvar_lngAbsenceCalShowCaptions = False
    Else
      mvar_lngAbsenceCalShowCaptions = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkShowCaptions.value = IIf(mvar_lngAbsenceCalShowCaptions, vbChecked, vbUnchecked)

    ' Get the AbsenceCal BHol Shading
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALBHOLSHADING
    If .NoMatch Then
      mvar_lngAbsenceCalBHolShading = False
    Else
      mvar_lngAbsenceCalBHolShading = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkBankHolidaysShading.value = IIf(mvar_lngAbsenceCalBHolShading, vbChecked, vbUnchecked)
    
    ' Get the AbsenceCal Include Working Days Only
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY
    If .NoMatch Then
      mvar_lngAbsenceCalIncludeWorkingDaysOnly = False
    Else
      mvar_lngAbsenceCalIncludeWorkingDaysOnly = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkIncludeWorkingDaysOnly.value = IIf(mvar_lngAbsenceCalIncludeWorkingDaysOnly, vbChecked, vbUnchecked)

    ' Get the AbsenceCal BHol Include
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALBHOLINCLUDE
    If .NoMatch Then
      mvar_lngAbsenceCalBHolInclude = False
    Else
      mvar_lngAbsenceCalBHolInclude = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, False, IIf(!parametervalue = "TRUE", True, False))
    End If
    chkBankHolidaysInclude.value = IIf(mvar_lngAbsenceCalBHolInclude, vbChecked, vbUnchecked)
    
    
    
    ' Get the Absence Parental Leave Type
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEPARENTALLEAVETYPE
    If .NoMatch Then
      mvar_strAbsenceParentLeaveType = vbNullString
    Else
      mvar_strAbsenceParentLeaveType = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, vbNullString, !parametervalue)
    End If

    ' Get the Absence Child No Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECHILDNO
    If .NoMatch Then
      mvar_lngAbsenceChildNoID = 0
    Else
      mvar_lngAbsenceChildNoID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEPARENTALREGION
    If .NoMatch Then
      mvar_lngAbsenceParentRegion = 0
    Else
      mvar_lngAbsenceParentRegion = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    
    
    ' Get the Dependants Table
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE
    If .NoMatch Then
      mvar_lngDependantsTableID = 0
    Else
      mvar_lngDependantsTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Dependants Child No Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSCHILDNO
    If .NoMatch Then
      mvar_lngDependantsChildNoID = 0
    Else
      mvar_lngDependantsChildNoID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Dependants Date Of Birth Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSDATEOFBIRTH
    If .NoMatch Then
      mvar_lngDependantsDateOfBirthID = 0
    Else
      mvar_lngDependantsDateOfBirthID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Dependants Adopted Date Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSADOPTEDDATE
    If .NoMatch Then
      mvar_lngDependantsAdoptedDateID = 0
    Else
      mvar_lngDependantsAdoptedDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Dependants Disabled Column
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSDISABLED
    If .NoMatch Then
      mvar_lngDependantsDisabledID = 0
    Else
      mvar_lngDependantsDisabledID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
  
  End With

  mvar_lngOriginalDependantsTableID = mvar_lngDependantsTableID
  mvar_lngOriginalAbsenceTableID = mvar_lngAbsenceTableID
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub optWorkingDaysType_Click(Index As Integer)
  ' Save the selected ID to a local variable.
  mvar_iAbsenceWorkingDaysType = Index
  RefreshWorkingDaysControls mvar_iAbsenceWorkingDaysType
  
  Changed = True
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
  ' Enable, and make visible the selected tab.
  
  If mblnReadOnly Then
    Exit Sub
  End If
  
  fraAbsence.Enabled = (SSTab1.Tab = giPAGE_ABSENCE)
  fraAbsenceType.Enabled = (SSTab1.Tab = giPAGE_ABSENCETYPE)
  fraCalendarDef.Enabled = (SSTab1.Tab = giPAGE_CALENDAR)
  fraCalendarInclude.Enabled = (SSTab1.Tab = giPAGE_CALENDAR)
  fraSSPColumns.Enabled = (SSTab1.Tab = giPAGE_SSP)
  fraWorkingDays.Enabled = (SSTab1.Tab = giPAGE_SSP)
  
End Sub

Private Sub ReadHistoricWP()

  ' RH 19/02/01
  With recModuleSetup
    
    ' Get the Working Pattern column ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN
    If .NoMatch Then
      mvar_lngWorkingPatternID = 0
    Else
      mvar_lngWorkingPatternID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    ' Get the Historic Working Pattern Table ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE
    If .NoMatch Then
      mvar_lngHWorkingPatternTableID = 0
    Else
      mvar_lngHWorkingPatternTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    ' Get the Historic Working Pattern Field ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD
    If .NoMatch Then
      mvar_lngHWorkingPatternFieldID = 0
    Else
      mvar_lngHWorkingPatternFieldID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    ' Get the Historic Working Pattern Date Effective ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE
    If .NoMatch Then
      mvar_lngHWorkingPatternDateID = 0
    Else
      mvar_lngHWorkingPatternDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With
  
End Sub


Private Sub RefreshAbsenceTypeValues()

  Dim rsAbsenceTypes As New ADODB.Recordset
  Dim strSQL As String
  Dim strAbsenceType As String
  Dim iListIndex As Integer
  Dim lngNewIndex As Long

  On Error GoTo LocalErr
  
  cboParentalLeaveAbsType.Clear
  cboParentalLeaveAbsType.AddItem "<None>"
  
  
  If cboAbsenceTypeTable.ListIndex = 0 Or _
     cboAbsenceTypeType.ListIndex = 0 Then
    cboParentalLeaveAbsType.ListIndex = 0
    Exit Sub
  End If
  
  
  'Assume full access as this is a lookup table
  strSQL = "SELECT " & cboAbsenceTypeType.Text & " FROM " & cboAbsenceTypeTable.Text
  rsAbsenceTypes.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  Do While Not rsAbsenceTypes.EOF
    lngNewIndex = AddItemToComboBox(cboParentalLeaveAbsType, rsAbsenceTypes.Fields(0).value, 0)
    
    If Trim(rsAbsenceTypes.Fields(0).value) = Trim(mvar_strAbsenceParentLeaveType) Then
      iListIndex = lngNewIndex
    End If
    rsAbsenceTypes.MoveNext
  Loop

  cboParentalLeaveAbsType.ListIndex = iListIndex
  
  rsAbsenceTypes.Close
  Set rsAbsenceTypes = Nothing

Exit Sub

LocalErr:
  MsgBox Err.Description, vbCritical, Me.Caption

End Sub


Private Sub RefreshDependantsControls()

  ' Refresh the Dependants controls.
  Dim iDependantsChildNoListIndex As Integer
  Dim iDependantsDateOfBirthListIndex As Integer
  Dim iDependantsAdoptedDateListIndex As Integer
  Dim iDependantsDisabledListIndex As Integer
  Dim objctl As Control
  Dim lngNewIndex As Long

  iDependantsChildNoListIndex = 0
  iDependantsDateOfBirthListIndex = 0
  iDependantsAdoptedDateListIndex = 0
  iDependantsDisabledListIndex = 0
  
  ' Clear the current contents of the combos.
  For Each objctl In Me
    If TypeOf objctl Is ComboBox And _
      (objctl.Name = "cboDepChildNo" Or _
      objctl.Name = "cboDepDateOfBirth" Or _
      objctl.Name = "cboDepAdoptedDate" Or _
      objctl.Name = "cboDepDisabled") Then

      With objctl
        .Clear
        AddItemToComboBox objctl, "<None>", 0
      End With
    End If
  Next objctl
  Set objctl = Nothing
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngDependantsTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngDependantsTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          If !DataType = dtinteger Then
            'Dependants Child No
            lngNewIndex = AddItemToComboBox(cboDepChildNo, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngDependantsChildNoID Then
              iDependantsChildNoListIndex = lngNewIndex
            End If
          End If
          
          If !DataType = dtTIMESTAMP Then
            'Dependants Date Of Birth
            lngNewIndex = AddItemToComboBox(cboDepDateOfBirth, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngDependantsDateOfBirthID Then
              iDependantsDateOfBirthListIndex = lngNewIndex
            End If

            'Dependants Adopted Date
            lngNewIndex = AddItemToComboBox(cboDepAdoptedDate, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngDependantsAdoptedDateID Then
              iDependantsAdoptedDateListIndex = lngNewIndex
            End If
          End If

          If !DataType = dtBIT Then
            'Dependants Disabled Column
            lngNewIndex = AddItemToComboBox(cboDepDisabled, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngDependantsDisabledID Then
              iDependantsDisabledListIndex = lngNewIndex
            End If
          End If
        
        End If

        .MoveNext
      Loop
    End If
  End With
  
  
  ' Select the appropriate combo items.
  cboDepChildNo.ListIndex = iDependantsChildNoListIndex
  cboDepDateOfBirth.ListIndex = iDependantsDateOfBirthListIndex
  cboDepAdoptedDate.ListIndex = iDependantsAdoptedDateListIndex
  cboDepDisabled.ListIndex = iDependantsDisabledListIndex

End Sub


Private Sub RefreshPersonnelControls()

  ' Refresh the Personnel controls.
  Dim iRegionColumn As Integer
  Dim lngNewIndex As Long

  iRegionColumn = 0
  cboParentalLeaveRegion.Clear
  AddItemToComboBox cboParentalLeaveRegion, "<None>", 0
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngPersonnelTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngPersonnelTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          If !DataType = dtVARCHAR Then
            lngNewIndex = AddItemToComboBox(cboParentalLeaveRegion, !ColumnName, !ColumnID)
            If !ColumnID = mvar_lngAbsenceParentRegion Then
              iRegionColumn = lngNewIndex
            End If
          End If
        
        End If

        .MoveNext
      Loop
    End If
  End With


  ' Select the appropriate combo items.
  cboParentalLeaveRegion.ListIndex = iRegionColumn

End Sub


