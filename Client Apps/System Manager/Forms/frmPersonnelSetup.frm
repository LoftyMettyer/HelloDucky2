VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmPersonnelSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personnel"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5020
   Icon            =   "frmPersonnelSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   14843
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Personnel"
      TabPicture(0)   =   "frmPersonnelSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTableDefinition"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "C&areer Change"
      TabPicture(1)   =   "frmPersonnelSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraWorkingPattern"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraRegion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Hierarchy"
      TabPicture(2)   =   "frmPersonnelSetup.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPostAllocationTable"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraHierarchyTable"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraPostAllocationTable 
         Caption         =   "Post Allocation Table :"
         Enabled         =   0   'False
         Height          =   1860
         Left            =   -74850
         TabIndex        =   69
         Top             =   2400
         Width           =   5400
         Begin VB.ComboBox cboPostAllocationTable 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   300
            Width           =   2505
         End
         Begin VB.ComboBox cboPostAllocationEndDateColumn 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1350
            Width           =   2505
         End
         Begin VB.ComboBox cboPostAllocationStartDateColumn 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   945
            Width           =   2505
         End
         Begin COALine.COA_Line linPostAllocationTable1 
            Height          =   30
            Left            =   180
            Top             =   765
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   53
         End
         Begin VB.Label lblPostAllocationTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Allocation Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   39
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label lblPostAllocationEndDateColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   65
            Top             =   1410
            Width           =   1785
         End
         Begin VB.Label lblPostAllocationStartDateColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   70
            Top             =   1005
            Width           =   1875
         End
      End
      Begin VB.Frame fraHierarchyTable 
         Caption         =   "Hierarchy Table :"
         Enabled         =   0   'False
         Height          =   1860
         Left            =   -74850
         TabIndex        =   64
         Top             =   400
         Width           =   5400
         Begin VB.ComboBox cboHierarchyTable 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   300
            Width           =   2505
         End
         Begin VB.ComboBox cboHierarchyIdentifyingColumn 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   945
            Width           =   2505
         End
         Begin VB.ComboBox cboHierarchyReportsToColumn 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1350
            Width           =   2505
         End
         Begin COALine.COA_Line linHierarchyTable1 
            Height          =   30
            Left            =   180
            Top             =   765
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   53
         End
         Begin VB.Label lblHierarchyIdentifyingColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Identifying Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   68
            Top             =   1005
            Width           =   1905
         End
         Begin VB.Label lblHierarchyTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hierarchy Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   67
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label lblHierarchyReportsToColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reports To Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   66
            Top             =   1410
            Width           =   1920
         End
      End
      Begin VB.Frame fraWorkingPattern 
         Caption         =   "Working Pattern :"
         Enabled         =   0   'False
         Height          =   2850
         Left            =   -74850
         TabIndex        =   55
         Top             =   3345
         Width           =   5400
         Begin VB.ComboBox cboWorkingPattern 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   705
            Width           =   2500
         End
         Begin VB.OptionButton optWorkingPatternStatic 
            Caption         =   "Static Working Pattern"
            Height          =   255
            Left            =   150
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optWorkingPatternHistory 
            Caption         =   "Historical Working Pattern"
            Height          =   255
            Left            =   150
            TabIndex        =   26
            Top             =   1140
            Width           =   3045
         End
         Begin VB.ComboBox cboWorkingPatternTable 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1515
            Width           =   2500
         End
         Begin VB.ComboBox cboWorkingPatternField 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1920
            Width           =   2500
         End
         Begin VB.ComboBox cboWorkingPatternEffectiveDate 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2325
            Width           =   2500
         End
         Begin VB.Label lblHistorialWorkingPatternColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Working Pattern Column :"
            Height          =   195
            Left            =   435
            TabIndex        =   59
            Top             =   1980
            Width           =   2340
         End
         Begin VB.Label lblHistorialWorkingPatternEffectiveColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date Column :"
            Height          =   195
            Left            =   435
            TabIndex        =   58
            Top             =   2385
            Width           =   2205
         End
         Begin VB.Label lblWorkingPattern 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Working Pattern Column :"
            Height          =   195
            Left            =   435
            TabIndex        =   57
            Top             =   765
            Width           =   2250
         End
         Begin VB.Label lblHistoricalWorkingPatternTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   435
            TabIndex        =   56
            Top             =   1575
            Width           =   990
         End
      End
      Begin VB.Frame fraRegion 
         Caption         =   "Region :"
         Enabled         =   0   'False
         Height          =   2850
         Left            =   -74850
         TabIndex        =   50
         Top             =   400
         Width           =   5400
         Begin VB.ComboBox cboRegionEffectiveDate 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2325
            Width           =   2500
         End
         Begin VB.ComboBox cboRegionField 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1920
            Width           =   2500
         End
         Begin VB.ComboBox cboRegionTable 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1515
            Width           =   2500
         End
         Begin VB.OptionButton optRegionHistory 
            Caption         =   "Historical Region"
            Height          =   255
            Left            =   150
            TabIndex        =   20
            Top             =   1140
            Width           =   1935
         End
         Begin VB.OptionButton optRegionStatic 
            Caption         =   "Static Region"
            Height          =   255
            Left            =   150
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.ComboBox cboRegion 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   705
            Width           =   2500
         End
         Begin VB.Label lblHistoricalRegionTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   435
            TabIndex        =   54
            Top             =   1575
            Width           =   945
         End
         Begin VB.Label lblRegion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region Column :"
            Height          =   195
            Left            =   435
            TabIndex        =   53
            Top             =   765
            Width           =   1710
         End
         Begin VB.Label lblHistorialRegionEffectiveColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date Column :"
            Height          =   195
            Left            =   435
            TabIndex        =   52
            Top             =   2385
            Width           =   2160
         End
         Begin VB.Label lblHistorialRegionColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region Column :"
            Height          =   195
            Left            =   435
            TabIndex        =   51
            Top             =   1980
            Width           =   1620
         End
      End
      Begin VB.Frame fraTableDefinition 
         Caption         =   "Personnel Records :"
         Height          =   7890
         Left            =   120
         TabIndex        =   38
         Top             =   400
         Width           =   5415
         Begin VB.ComboBox cboSecurityGroup 
            Height          =   315
            Left            =   2715
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   7365
            Width           =   2505
         End
         Begin VB.ComboBox cboSSIPhotograph 
            Height          =   315
            ItemData        =   "frmPersonnelSetup.frx":0060
            Left            =   2730
            List            =   "frmPersonnelSetup.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   5325
            Width           =   2505
         End
         Begin VB.ComboBox cboSSIWelcome 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   4935
            Width           =   2505
         End
         Begin VB.ComboBox cboLoginName 
            Height          =   315
            Index           =   1
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   4540
            Width           =   2505
         End
         Begin VB.ComboBox cboJobTitle 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   6525
            Width           =   2505
         End
         Begin VB.ComboBox cboManagerStaffNo 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   6120
            Width           =   2505
         End
         Begin VB.ComboBox cboGrade 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   5715
            Width           =   2505
         End
         Begin VB.ComboBox cboLoginName 
            Height          =   315
            Index           =   0
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   4140
            Width           =   2505
         End
         Begin VB.ComboBox cboPersonnelTable 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   300
            Width           =   2505
         End
         Begin VB.ComboBox cboEmployeeNumber 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   940
            Width           =   2505
         End
         Begin VB.ComboBox cboForename 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1340
            Width           =   2505
         End
         Begin VB.ComboBox cboStartDate 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2140
            Width           =   2505
         End
         Begin VB.ComboBox cboSurname 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1740
            Width           =   2505
         End
         Begin VB.ComboBox cboLeavingDate 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2540
            Width           =   2505
         End
         Begin VB.ComboBox cboFullPartTime 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2940
            Width           =   2505
         End
         Begin VB.ComboBox cboDepartment 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3340
            Width           =   2505
         End
         Begin VB.ComboBox cboWorkEmail 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   6930
            Width           =   2505
         End
         Begin VB.ComboBox cboDateOfBirth 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3740
            Width           =   2505
         End
         Begin COALine.COA_Line ASRDummyLine1 
            Height          =   30
            Left            =   180
            Top             =   765
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   53
         End
         Begin VB.Label lblSecurityGroup 
            Caption         =   "User Group :"
            Height          =   300
            Left            =   210
            TabIndex        =   74
            Top             =   7395
            Width           =   1935
         End
         Begin VB.Label lblIntranetPhotograph 
            BackStyle       =   0  'Transparent
            Caption         =   "Self-service Photo :"
            Height          =   225
            Left            =   195
            TabIndex        =   72
            Top             =   5400
            Width           =   1905
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblIntranetWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   "Self-service Welcome Column :"
            Height          =   450
            Left            =   195
            TabIndex        =   71
            Top             =   4860
            Width           =   1905
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Job Title Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   63
            Top             =   6585
            Width           =   2355
         End
         Begin VB.Label lblManagerStaffNo 
            AutoSize        =   -1  'True
            Caption         =   "Manager Staff No. Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   62
            Top             =   6180
            Width           =   2370
         End
         Begin VB.Label lblGradeColumn 
            AutoSize        =   -1  'True
            Caption         =   "Grade Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   61
            Top             =   5775
            Width           =   2190
         End
         Begin VB.Label lblLoginNameColumn 
            AutoSize        =   -1  'True
            Caption         =   "Login Name Column(s) :"
            Height          =   195
            Left            =   195
            TabIndex        =   60
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label lblPersonnelTable 
            BackStyle       =   0  'Transparent
            Caption         =   "Personnel Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   49
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblStaffNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Staff No. Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   48
            Top             =   1005
            Width           =   2415
         End
         Begin VB.Label lblStartSession 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forename Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   47
            Top             =   1395
            Width           =   2475
         End
         Begin VB.Label lblEndSession 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   46
            Top             =   2205
            Width           =   2505
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Surname Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   45
            Top             =   1800
            Width           =   2385
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Leaving Date Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   44
            Top             =   2595
            Width           =   2700
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full / Part Time Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   43
            Top             =   3000
            Width           =   2820
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   42
            Top             =   3405
            Width           =   2610
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Work Email Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   41
            Top             =   6975
            Width           =   1830
         End
         Begin VB.Label lblDateOfBirth 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   40
            Top             =   3795
            Width           =   2670
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3315
      TabIndex        =   36
      Top             =   8655
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4575
      TabIndex        =   37
      Top             =   8655
      Width           =   1200
   End
End
Attribute VB_Name = "frmPersonnelSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Whats this for ? Comment it out n see if it still works ok!
'  Event UnLoad()

' Personnel Tab
Private mvar_lngPersonnelTableID As Long
Private mvar_lngEmployeeNumberID As Long
Private mvar_lngForenameID As Long
Private mvar_lngSurnameID As Long
Private mvar_lngStartDateID As Long
Private mvar_lngLeavingDateID As Long
Private mvar_lngFullPartTimeID As Long
Private mvar_lngWorkEmailID As Long
Private mvar_lngDepartmentID As Long
Private mvar_lngDateOfBirthID As Long
Private mvar_lngLoginNameID As Long
Private mvar_lngSecondLoginNameID As Long
Private mvar_lngGradeID As Long
Private mvar_lngManagerStaffNoID As Long
Private mvar_lngJobTitleID As Long
Private mvar_lngSecurityGroupID As Long

' Career Change Tab
Private mvar_lngRegionID As Long
Private mvar_lngHRegionTableID As Long
Private mvar_lngHRegionFieldID As Long
Private mvar_lngHRegionDateID As Long
Private mvar_lngWorkingPatternID As Long
Private mvar_lngHWorkingPatternTableID As Long
Private mvar_lngHWorkingPatternFieldID As Long
Private mvar_lngHWorkingPatternDateID As Long
Private mvar_lngSSIWelcomeID As Long
Private mvar_lngSSIPhotographID As Long

' Hierarchy Tab
Private mvar_lngHierarchyTableID As Long
Private mvar_lngHierarchyIdentifyingColumnID As Long
Private mvar_lngHierarchyReportsToColumnID As Long
Private mvar_lngPostAllocationTableID As Long
Private mvar_lngPostAllocationStartDateColumnID As Long
Private mvar_lngPostAllocationEndDateColumnID As Long

Private mvar_lngOriginalPersonnelTableID As Long
Private mvar_lngOriginalHierarchyTableID As Long
Private mvar_lngOriginalHierarchyIdentifyingColumnID As Long
Private mvar_lngOriginalHierarchyReportsToColumnID As Long

Private mblnReadOnly As Boolean
Private mbLoading As Boolean
Private mfChanged As Boolean

' Page number constants.
Private Const giPAGE_PERSONNEL = 0
Private Const giPAGE_CAREER_CHANGE = 1
Private Const giPAGE_HIERARCHY = 2


Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOk.Enabled = True
End Property

Private Sub cboGrade_Click()
  With cboGrade
    mvar_lngGradeID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboHierarchyIdentifyingColumn_Click()
  With cboHierarchyIdentifyingColumn
    mvar_lngHierarchyIdentifyingColumnID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboHierarchyReportsToColumn_Click()

  With cboHierarchyReportsToColumn
    mvar_lngHierarchyReportsToColumnID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboHierarchyTable_Click()
  With cboHierarchyTable
    mvar_lngHierarchyTableID = .ItemData(.ListIndex)
  End With

  RefreshHierarchyControls
  Changed = True
End Sub


Private Sub cboManagerStaffNo_Click()
  With cboManagerStaffNo
    mvar_lngManagerStaffNoID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboJobTitle_Click()
  With cboJobTitle
    mvar_lngJobTitleID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboLoginName_Click(Index As Integer)
  With cboLoginName(Index)
    If Index = 1 Then
      mvar_lngSecondLoginNameID = .ItemData(.ListIndex)
    Else
      mvar_lngLoginNameID = .ItemData(.ListIndex)
    End If
  End With
  
  If Not mbLoading Then
    mbLoading = True
    RefreshLoginColumnControls
    mbLoading = False
    Changed = True
  End If
  
End Sub

Private Sub cboPostAllocationEndDateColumn_Click()
  With cboPostAllocationEndDateColumn
    mvar_lngPostAllocationEndDateColumnID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboPostAllocationStartDateColumn_Click()

  With cboPostAllocationStartDateColumn
    mvar_lngPostAllocationStartDateColumnID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub cboPostAllocationTable_Click()
  With cboPostAllocationTable
    mvar_lngPostAllocationTableID = .ItemData(.ListIndex)
  End With

  RefreshPostAllocationControls
  Changed = True
End Sub

Private Sub cboSecurityGroup_Click()
  With cboSecurityGroup
    mvar_lngSecurityGroupID = .ItemData(.ListIndex)
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

Private Sub Form_Load()

  Screen.MousePointer = vbHourglass

  ' Ensure the first tab is displayed.
  SSTab1.Tab = 0
    
  mbLoading = True
  cmdOk.Enabled = False
  'Changed = False
  
  ' Only display the Hierarchy tab if UDFs are enabled.
  SSTab1.TabVisible(2) = gbEnableUDFFunctions

  mblnReadOnly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  ' Read the current settings from the database.
  ReadParameters

  ' Initialise all controls with the current settings, or defaults.
  InitialiseBaseTableCombos

  ' Set the correct option button (and therefore disable correct combos)
  If mvar_lngRegionID > 0 Then optRegionStatic.value = True Else optRegionHistory.value = True
  If mvar_lngWorkingPatternID > 0 Then optWorkingPatternStatic.value = True Else optWorkingPatternHistory.value = True
  
  Changed = False
  mbLoading = False
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


Private Sub InitialiseBaseTableCombos()
  
  ' Initialise the Base Table combo(s)
  Dim iPersonnelTableListIndex As Integer
  Dim iHierarchyTableListIndex As Integer
  
  iPersonnelTableListIndex = 0
  iHierarchyTableListIndex = 0
  
  ' Clear the combo, and add '<None>' items.
  With cboPersonnelTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboHierarchyTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  ' Add items to the combo for each table that has not been deleted.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted Then
        cboPersonnelTable.AddItem !TableName
        cboPersonnelTable.ItemData(cboPersonnelTable.NewIndex) = !TableID
        If !TableID = mvar_lngPersonnelTableID Then
          iPersonnelTableListIndex = cboPersonnelTable.NewIndex
        End If
      
        cboHierarchyTable.AddItem !TableName
        cboHierarchyTable.ItemData(cboHierarchyTable.NewIndex) = !TableID
        If !TableID = mvar_lngHierarchyTableID Then
          iHierarchyTableListIndex = cboHierarchyTable.NewIndex
        End If
      End If
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
  With cboHierarchyTable
    '.Enabled = True
    .Enabled = Not mblnReadOnly
    .ListIndex = iHierarchyTableListIndex
  End With
  
  With cboPersonnelTable
    '.Enabled = True
    .Enabled = Not mblnReadOnly
    .ListIndex = iPersonnelTableListIndex
  End With

End Sub

Private Sub cboPersonnelTable_Click()
  
  With cboPersonnelTable
    mvar_lngPersonnelTableID = .ItemData(.ListIndex)
  End With
  
  RefreshPersonnelColumnControls
  RefreshDependantColumnControls
  RefreshHierarchyControls
  Changed = True
End Sub

Private Sub cboEmployeeNumber_Click()
  
  With cboEmployeeNumber
    mvar_lngEmployeeNumberID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboForename_Click()

  With cboForename
    mvar_lngForenameID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboSurname_Click()

  With cboSurname
    mvar_lngSurnameID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboSSIWelcome_Click()

  With cboSSIWelcome
    mvar_lngSSIWelcomeID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboSSIPhotograph_Click()

  With cboSSIPhotograph
    mvar_lngSSIPhotographID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboStartDate_Click()

  With cboStartDate
    mvar_lngStartDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboLeavingDate_Click()

  With cboLeavingDate
    mvar_lngLeavingDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboFullPartTime_Click()

  With cboFullPartTime
    mvar_lngFullPartTimeID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWorkEmail_Click()

  With cboWorkEmail
    mvar_lngWorkEmailID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboDepartment_Click()

  With cboDepartment
    mvar_lngDepartmentID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboDateOfBirth_Click()
  
  With cboDateOfBirth
    mvar_lngDateOfBirthID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboRegion_Click()
  
  With cboRegion
    mvar_lngRegionID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboRegionTable_Click()

  With cboRegionTable
    mvar_lngHRegionTableID = .ItemData(.ListIndex)
  End With

  RefreshHistoricRegionColumnControls
  Changed = True
End Sub

Private Sub cboRegionField_Click()

  With cboRegionField
    mvar_lngHRegionFieldID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboRegionEffectiveDate_Click()

  With cboRegionEffectiveDate
    mvar_lngHRegionDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWorkingPattern_Click()
  
  With cboWorkingPattern
    mvar_lngWorkingPatternID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWorkingPatternTable_Click()

  With cboWorkingPatternTable
    mvar_lngHWorkingPatternTableID = .ItemData(.ListIndex)
  End With

  RefreshHistoricWorkingPatternColumnControls
  Changed = True
End Sub

Private Sub cboWorkingPatternField_Click()

  With cboWorkingPatternField
    mvar_lngHWorkingPatternFieldID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboWorkingPatternEffectiveDate_Click()

  With cboWorkingPatternEffectiveDate
    mvar_lngHWorkingPatternDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub


Private Sub RefreshPersonnelColumnControls()
  ' Refresh the Personnel column controls - including the static combos
  ' on the second tab of the form
  Dim iEmployeeNumberListIndex As Integer
  Dim iForenameListIndex As Integer
  Dim iSurnameListIndex As Integer
  Dim iStartDateListIndex As Integer
  Dim iLeavingDateListIndex As Integer
  Dim iFullPartTimeListIndex As Integer
  Dim iWorkEmailListIndex As Integer
  Dim iDepartmentListIndex As Integer
  Dim iWorkingPatternListIndex As Integer
  Dim iDateOfBirthListIndex As Integer
  Dim iRegionListIndex As Integer
  Dim iGradeListIndex As Integer
  Dim iManagerStaffNoListIndex As Integer
  Dim iJobTitleListIndex As Integer
  Dim iSSIWelcomeListIndex As Integer
  Dim iSSIPhotographListIndex As Integer
  Dim iSecurityGroupIndex As Integer
  Dim objctl As Control
  
  iEmployeeNumberListIndex = 0
  iForenameListIndex = 0
  iSurnameListIndex = 0
  iStartDateListIndex = 0
  iLeavingDateListIndex = 0
  iFullPartTimeListIndex = 0
  iWorkEmailListIndex = 0
  iDepartmentListIndex = 0
  iWorkingPatternListIndex = 0
  iDateOfBirthListIndex = 0
  iRegionListIndex = 0
  iGradeListIndex = 0
  iManagerStaffNoListIndex = 0
  iJobTitleListIndex = 0
  iSecurityGroupIndex = 0

  UI.LockWindow Me.hWnd
  
  ' Clear the current contents of the combos.
  For Each objctl In Me
    If (TypeOf objctl Is ComboBox) And _
      (objctl.Name <> "cboPersonnelTable") And _
      (objctl.Name <> "cboLoginName") And _
      (objctl.Container.Name <> "fraHierarchyTable" And _
      (objctl.Container.Name <> "fraPostAllocationTable")) Then
      With objctl
        .Clear
        .AddItem "<None>"
        .ItemData(.NewIndex) = 0
      End With
    End If
  Next objctl
  
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
          
          ' Only load date fields into the start/leaving date and dateOfBirth combos
          If !DataType = dtTIMESTAMP Then
            cboStartDate.AddItem !ColumnName
            cboStartDate.ItemData(cboStartDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngStartDateID Then
              iStartDateListIndex = cboStartDate.NewIndex
            End If
            
            cboLeavingDate.AddItem !ColumnName
            cboLeavingDate.ItemData(cboLeavingDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngLeavingDateID Then
              iLeavingDateListIndex = cboLeavingDate.NewIndex
            End If
            
            cboDateOfBirth.AddItem !ColumnName
            cboDateOfBirth.ItemData(cboDateOfBirth.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngDateOfBirthID Then
              iDateOfBirthListIndex = cboDateOfBirth.NewIndex
            End If
            
          End If
          
          'MH20030918 Fault 6995
          ''' Load unique fields for the employee number
          'If !uniqueCheck = True Then
          If !uniqueCheckType = -1 Then
            cboEmployeeNumber.AddItem !ColumnName
            cboEmployeeNumber.ItemData(cboEmployeeNumber.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngEmployeeNumberID Then
              iEmployeeNumberListIndex = cboEmployeeNumber.NewIndex
            End If
          End If
                  
          ' Load varchar fields
          If !DataType = dtVARCHAR Then
            cboForename.AddItem !ColumnName
            cboForename.ItemData(cboForename.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngForenameID Then
              iForenameListIndex = cboForename.NewIndex
            End If
            
            cboSurname.AddItem !ColumnName
            cboSurname.ItemData(cboSurname.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngSurnameID Then
              iSurnameListIndex = cboSurname.NewIndex
            End If
            
            cboFullPartTime.AddItem !ColumnName
            cboFullPartTime.ItemData(cboFullPartTime.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngFullPartTimeID Then
              iFullPartTimeListIndex = cboFullPartTime.NewIndex
            End If
            
            cboWorkEmail.AddItem !ColumnName
            cboWorkEmail.ItemData(cboWorkEmail.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngWorkEmailID Then
              iWorkEmailListIndex = cboWorkEmail.NewIndex
            End If
            
            cboJobTitle.AddItem !ColumnName
            cboJobTitle.ItemData(cboJobTitle.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngJobTitleID Then
              iJobTitleListIndex = cboJobTitle.NewIndex
            End If
            
            cboDepartment.AddItem !ColumnName
            cboDepartment.ItemData(cboDepartment.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngDepartmentID Then
              iDepartmentListIndex = cboDepartment.NewIndex
            End If
          
            cboRegion.AddItem !ColumnName
            cboRegion.ItemData(cboRegion.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngRegionID Then
              iRegionListIndex = cboRegion.NewIndex
            End If

            cboGrade.AddItem !ColumnName
            cboGrade.ItemData(cboGrade.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngGradeID Then
              iGradeListIndex = cboGrade.NewIndex
            End If

            cboSSIWelcome.AddItem !ColumnName
            cboSSIWelcome.ItemData(cboSSIWelcome.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngSSIWelcomeID Then
              iSSIWelcomeListIndex = cboSSIWelcome.NewIndex
            End If

          End If

          ' Load working pattern fields
          If !DataType = dtLONGVARCHAR Then
            cboWorkingPattern.AddItem !ColumnName
            cboWorkingPattern.ItemData(cboWorkingPattern.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngWorkingPatternID Then
              iWorkingPatternListIndex = cboWorkingPattern.NewIndex
            End If
          End If

          If !DataType = dtVARCHAR Or _
                  !DataType = dtNUMERIC Or _
                  !DataType = dtINTEGER Then
            cboManagerStaffNo.AddItem !ColumnName
            cboManagerStaffNo.ItemData(cboManagerStaffNo.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngManagerStaffNoID Then
              iManagerStaffNoListIndex = cboManagerStaffNo.NewIndex
            End If
          End If
                    
          ' Security Group
          If !DataType = dtVARCHAR Then
            cboSecurityGroup.AddItem !ColumnName
            cboSecurityGroup.ItemData(cboSecurityGroup.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngSecurityGroupID Then
              iSecurityGroupIndex = cboSecurityGroup.NewIndex
            End If
          End If
          
          ' Photograph column combo for SSI
          If !DataType = dtVARBINARY And !OLEType = 2 And !MaxOLESizeEnabled Then
            cboSSIPhotograph.AddItem !ColumnName
            cboSSIPhotograph.ItemData(cboSSIPhotograph.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngSSIPhotographID Then
              iSSIPhotographListIndex = cboSSIPhotograph.NewIndex
            End If
          End If
                
        End If

        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboEmployeeNumber.ListIndex = iEmployeeNumberListIndex
  cboForename.ListIndex = iForenameListIndex
  cboSurname.ListIndex = iSurnameListIndex
  cboStartDate.ListIndex = iStartDateListIndex
  cboLeavingDate.ListIndex = iLeavingDateListIndex
  cboFullPartTime.ListIndex = iFullPartTimeListIndex
  cboWorkEmail.ListIndex = iWorkEmailListIndex
  cboDepartment.ListIndex = iDepartmentListIndex
  cboWorkingPattern.ListIndex = iWorkingPatternListIndex
  cboDateOfBirth.ListIndex = iDateOfBirthListIndex
  cboRegion.ListIndex = iRegionListIndex
  cboGrade.ListIndex = iGradeListIndex
  cboManagerStaffNo.ListIndex = iManagerStaffNoListIndex
  cboJobTitle.ListIndex = iJobTitleListIndex
  cboSSIWelcome.ListIndex = iSSIWelcomeListIndex
  cboSSIPhotograph.ListIndex = iSSIPhotographListIndex
  cboSecurityGroup.ListIndex = iSecurityGroupIndex

  RefreshLoginColumnControls
  
  UI.UnlockWindow
  
End Sub


Private Sub RefreshLoginColumnControls()
  ' Refresh the Personnel Login column controls
  Dim iLoginNameListIndex As Integer
  Dim iSecondLoginNameListIndex As Integer
  Dim ctlCombo As ComboBox
  
  iLoginNameListIndex = 0
  iSecondLoginNameListIndex = 0

  UI.LockWindow Me.hWnd

  ' Clear the current contents of the combos.
  For Each ctlCombo In cboLoginName
    With ctlCombo
      .Clear
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End With
  Next ctlCombo
  
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
          (!columntype <> giCOLUMNTYPE_SYSTEM) And _
          (!DataType = dtVARCHAR) Then
          
          ' Load varchar fields
          If !ColumnID <> mvar_lngSecondLoginNameID Then
            cboLoginName(0).AddItem !ColumnName
            cboLoginName(0).ItemData(cboLoginName(0).NewIndex) = !ColumnID
            If !ColumnID = mvar_lngLoginNameID Then
              iLoginNameListIndex = cboLoginName(0).NewIndex
            End If
          End If

          If !ColumnID <> mvar_lngLoginNameID Then
            cboLoginName(1).AddItem !ColumnName
            cboLoginName(1).ItemData(cboLoginName(1).NewIndex) = !ColumnID
            If !ColumnID = mvar_lngSecondLoginNameID Then
              iSecondLoginNameListIndex = cboLoginName(1).NewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboLoginName(0).ListIndex = iLoginNameListIndex
  cboLoginName(1).ListIndex = iSecondLoginNameListIndex

  UI.UnlockWindow

End Sub



Private Sub RefreshDependantColumnControls()
  
  ' Refresh the Historic Region/WP column controls.
  Dim iHRegionTableListIndex As Integer
  Dim iHWorkingPatternTableListIndex As Integer
  
  iHRegionTableListIndex = 0
  iHWorkingPatternTableListIndex = 0
  
  ' Clear the current contents of the combo
  cboRegionTable.Clear
  cboRegionTable.AddItem "<None>"
  cboRegionTable.ItemData(cboRegionTable.NewIndex) = 0
  
  cboWorkingPatternTable.Clear
  cboWorkingPatternTable.AddItem "<None>"
  cboWorkingPatternTable.ItemData(cboWorkingPatternTable.NewIndex) = 0
  
  ' Add the children of the Personnel table
  
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
          
          cboRegionTable.AddItem !TableName
          cboRegionTable.ItemData(cboRegionTable.NewIndex) = !TableID
          If !TableID = mvar_lngHRegionTableID Then
            iHRegionTableListIndex = cboRegionTable.NewIndex
          End If
          
          cboWorkingPatternTable.AddItem !TableName
          cboWorkingPatternTable.ItemData(cboWorkingPatternTable.NewIndex) = !TableID
          If !TableID = mvar_lngHWorkingPatternTableID Then
            iHWorkingPatternTableListIndex = cboWorkingPatternTable.NewIndex
          End If
          
        End If
        
      End If
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
  cboRegionTable.ListIndex = iHRegionTableListIndex
  cboWorkingPatternTable.ListIndex = iHWorkingPatternTableListIndex

End Sub


Private Sub RefreshHistoricRegionColumnControls()
  
  ' Refresh the Historic Region column controls.
  Dim iRegionFieldListIndex As Integer
  Dim iRegionEffectiveDateListIndex As Integer
  
  iRegionFieldListIndex = 0
  iRegionEffectiveDateListIndex = 0
  
  cboRegionField.Clear
  cboRegionField.AddItem "<None>"
  cboRegionField.ItemData(cboRegionField.NewIndex) = 0
  cboRegionEffectiveDate.Clear
  cboRegionEffectiveDate.AddItem "<None>"
  cboRegionEffectiveDate.ItemData(cboRegionEffectiveDate.NewIndex) = 0
 
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngHRegionTableID
    
    If Not .NoMatch Then
      
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        
        If !TableID <> mvar_lngHRegionTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
          
          If !DataType = dtVARCHAR Then
            cboRegionField.AddItem !ColumnName
            cboRegionField.ItemData(cboRegionField.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngHRegionFieldID Then
              iRegionFieldListIndex = cboRegionField.NewIndex
            End If
          End If
          
          If !DataType = dtTIMESTAMP Then
            cboRegionEffectiveDate.AddItem !ColumnName
            cboRegionEffectiveDate.ItemData(cboRegionEffectiveDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngHRegionDateID Then
              iRegionEffectiveDateListIndex = cboRegionEffectiveDate.NewIndex
            End If
          End If
         End If
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboRegionField.ListIndex = iRegionFieldListIndex
  cboRegionEffectiveDate.ListIndex = iRegionEffectiveDateListIndex

End Sub

Private Sub RefreshHistoricWorkingPatternColumnControls()
  
  ' Refresh the Historic Region column controls.
  Dim iWorkingPatternFieldListIndex As Integer
  Dim iWorkingPatternEffectiveDateListIndex As Integer
  
  iWorkingPatternFieldListIndex = 0
  iWorkingPatternEffectiveDateListIndex = 0
  
  cboWorkingPatternField.Clear
  cboWorkingPatternField.AddItem "<None>"
  cboWorkingPatternField.ItemData(cboWorkingPatternField.NewIndex) = 0
  cboWorkingPatternEffectiveDate.Clear
  cboWorkingPatternEffectiveDate.AddItem "<None>"
  cboWorkingPatternEffectiveDate.ItemData(cboWorkingPatternEffectiveDate.NewIndex) = 0
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngHWorkingPatternTableID
    
    If Not .NoMatch Then
      
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        
        If !TableID <> mvar_lngHWorkingPatternTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then
          
          If !DataType = dtLONGVARCHAR Then
            cboWorkingPatternField.AddItem !ColumnName
            cboWorkingPatternField.ItemData(cboWorkingPatternField.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngHWorkingPatternFieldID Then
              iWorkingPatternFieldListIndex = cboWorkingPatternField.NewIndex
            End If
          End If
          
          If !DataType = dtTIMESTAMP Then
            cboWorkingPatternEffectiveDate.AddItem !ColumnName
            cboWorkingPatternEffectiveDate.ItemData(cboWorkingPatternEffectiveDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngHWorkingPatternDateID Then
              iWorkingPatternEffectiveDateListIndex = cboWorkingPatternEffectiveDate.NewIndex
            End If
          End If
        End If
        
        .MoveNext
      Loop
    End If
  End With
  
  ' Select the appropriate combo items.
  cboWorkingPatternField.ListIndex = iWorkingPatternFieldListIndex
  cboWorkingPatternEffectiveDate.ListIndex = iWorkingPatternEffectiveDateListIndex

End Sub


'##################################################
'
' COMMAND BUTTONS AND ACTIONS
'
'##################################################

Private Sub cmdOk_Click()

 'AE20071119 Fault #12607
  'If ValidateSetup Then
    'SaveChanges
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If
  
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

Private Function ValidateSetup() As Boolean

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
  
  If Me.optRegionHistory.value = True Then
    If Me.cboRegionTable.ListIndex > 0 Then
      If Me.cboRegionField.ListIndex = 0 Or Me.cboRegionEffectiveDate.ListIndex = 0 Then
        MsgBox "You have opted to use historic Regional data but have not defined" & vbCrLf & _
               "both the 'Region Column' and the 'Effective Date Column'.", vbExclamation + vbOKOnly, App.Title
        ValidateSetup = False
        Exit Function
      End If
    End If
  End If
  
  If Me.optWorkingPatternHistory.value = True Then
    If Me.cboWorkingPatternTable.ListIndex > 0 Then
      If Me.cboWorkingPatternField.ListIndex = 0 Or Me.cboWorkingPatternEffectiveDate.ListIndex = 0 Then
        MsgBox "You have opted to use historic Working Pattern data but have not defined" & vbCrLf & _
               "both the 'Working Pattern Column' and 'Effective Date Column'.", vbExclamation + vbOKOnly, App.Title
        ValidateSetup = False
        Exit Function
      End If
    End If
  End If


  'MH20030918 Fault 5553
  If mvar_lngEmployeeNumberID > 0 And mvar_lngManagerStaffNoID > 0 Then
    If GetColumnDataType(mvar_lngEmployeeNumberID) <> GetColumnDataType(mvar_lngManagerStaffNoID) Then
      MsgBox "The data types for the Staff No and Manager Staff No columns do not match.", vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If
  
  ' Don't allow the Personnel table to change if the special functions are used anywhere.
  fSpecialFunctionUsed = False
  
  If (mvar_lngOriginalPersonnelTableID <> mvar_lngPersonnelTableID) Then
    ' Find any expression field components that use the special functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (30,46,47,66,67,68,70,71,72,73)"
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
      MsgBox "The 'Personnel' table cannot be changed." & vbCrLf & vbCrLf & _
        "It is used as the base table for the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If
  
  ' Don't allow the Hierarchy table to change if the Hierarchy
  ' functions are used anywhere.
  fSpecialFunctionUsed = False
  
  If (mvar_lngOriginalHierarchyTableID <> mvar_lngHierarchyTableID) Or _
    (mvar_lngHierarchyIdentifyingColumnID = 0) Or _
    (mvar_lngHierarchyReportsToColumnID = 0) Then
    ' Find any expression field components that use the Hierarchy functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (65,66,67,68,69,70,71,72)"
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
      If (mvar_lngOriginalHierarchyTableID <> mvar_lngHierarchyTableID) Then
        MsgBox "The 'Hierarchy' table cannot be changed." & vbCrLf & vbCrLf & _
          "It is used as the base table for the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
          "which is owned by the '" & sExprParentTable & "' table.", _
          vbExclamation + vbOKOnly, App.Title
      ElseIf (mvar_lngHierarchyIdentifyingColumnID = 0) Then
        MsgBox "The 'Identifying' column cannot be set to '<None>'." & vbCrLf & vbCrLf & _
          "It is used by the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
          "which is owned by the '" & sExprParentTable & "' table.", _
          vbExclamation + vbOKOnly, App.Title
      Else
        MsgBox "The 'Reports To' column cannot be set to '<None>'." & vbCrLf & vbCrLf & _
          "It is used by the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
          "which is owned by the '" & sExprParentTable & "' table.", _
          vbExclamation + vbOKOnly, App.Title
      End If
      
      ValidateSetup = False
      Exit Function
    End If
  End If

  If (mvar_lngLoginNameID = 0) _
    And (mvar_lngSecondLoginNameID = 0) Then
    ' Find any expression field components that use the Hierarchy functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (66,68,70,72)"
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
      MsgBox "The 'Login Name' columns cannot both be set to '<None>'." & vbCrLf & vbCrLf & _
        "They are used by the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.Title
      
      ValidateSetup = False
      Exit Function
    End If
  End If

  If (mvar_lngHierarchyTableID <> mvar_lngPersonnelTableID) And _
    (mvar_lngPostAllocationTableID = 0) Then
    ' Find any expression field components that use the special functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (66,67,68,70,71,72)"
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
      MsgBox "The 'Post Allocation' table cannot be set to '<None>'." & vbCrLf & vbCrLf & _
        "It is used by the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If
  
  If (mvar_lngHierarchyIdentifyingColumnID > 0) And _
    (mvar_lngHierarchyReportsToColumnID > 0) And _
    (mvar_lngHierarchyIdentifyingColumnID <> mvar_lngHierarchyReportsToColumnID) Then
    If GetColumnDataType(mvar_lngHierarchyIdentifyingColumnID) <> GetColumnDataType(mvar_lngHierarchyReportsToColumnID) Then
      MsgBox "The data type of the 'Reports To' column must match that of the 'Identifying' column.", _
        vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If
  
  If (mvar_lngOriginalHierarchyIdentifyingColumnID > 0) And _
    (mvar_lngHierarchyIdentifyingColumnID > 0) And _
    (mvar_lngOriginalHierarchyIdentifyingColumnID <> mvar_lngHierarchyIdentifyingColumnID) Then
    
    If GetColumnDataType(mvar_lngOriginalHierarchyIdentifyingColumnID) <> GetColumnDataType(mvar_lngHierarchyIdentifyingColumnID) Then
      ' Identifying column has changed and is now a different data type.
      ' Check if its used in any expressions.
      sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
        " FROM tmpComponents" & _
        " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
        " AND tmpComponents.functionID IN (65,67,69,71)"
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
        MsgBox "The data type of the 'Identifying' column cannot change." & vbCrLf & vbCrLf & _
          "It is used by the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
          "which is owned by the '" & sExprParentTable & "' table.", _
          vbExclamation + vbOKOnly, App.Title
        ValidateSetup = False
        Exit Function
      End If
    End If
  End If
  
  If (mvar_lngOriginalHierarchyReportsToColumnID > 0) And _
    (mvar_lngHierarchyReportsToColumnID > 0) And _
    (mvar_lngOriginalHierarchyReportsToColumnID <> mvar_lngHierarchyReportsToColumnID) Then
      
    If GetColumnDataType(mvar_lngOriginalHierarchyReportsToColumnID) <> GetColumnDataType(mvar_lngOriginalHierarchyReportsToColumnID) Then
      ' ReportsTor column has changed and is now a different data type.
      ' Check if its used in any expressions.
      sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
        " FROM tmpComponents" & _
        " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
        " AND tmpComponents.functionID IN (65,67,69,71)"
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
        MsgBox "The data type of the 'Identifying' column cannot change." & vbCrLf & vbCrLf & _
          "It is used by the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
          "which is owned by the '" & sExprParentTable & "' table.", _
          vbExclamation + vbOKOnly, App.Title
        ValidateSetup = False
        Exit Function
      End If
    End If
  End If
  
  ValidateSetup = True
  Exit Function
  
ValidateError:
  
  MsgBox "Error validating the setup." & vbCrLf & _
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
  
  ' Write the parameter values to the local database.
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, gsPARAMETERTYPE_TABLEID, mvar_lngPersonnelTableID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER, gsPARAMETERTYPE_COLUMNID, mvar_lngEmployeeNumberID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME, gsPARAMETERTYPE_COLUMNID, mvar_lngForenameID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME, gsPARAMETERTYPE_COLUMNID, mvar_lngSurnameID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SSIWELCOME, gsPARAMETERTYPE_COLUMNID, mvar_lngSSIWelcomeID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SSIPHOTOGRAPH, gsPARAMETERTYPE_COLUMNID, mvar_lngSSIPhotographID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_STARTDATE, gsPARAMETERTYPE_COLUMNID, mvar_lngStartDateID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE, gsPARAMETERTYPE_COLUMNID, mvar_lngLeavingDateID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FULLPARTTIME, gsPARAMETERTYPE_COLUMNID, mvar_lngFullPartTimeID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKEMAIL, gsPARAMETERTYPE_COLUMNID, mvar_lngWorkEmailID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT, gsPARAMETERTYPE_COLUMNID, mvar_lngDepartmentID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DATEOFBIRTH, gsPARAMETERTYPE_COLUMNID, mvar_lngDateOfBirthID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME, gsPARAMETERTYPE_COLUMNID, mvar_lngLoginNameID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECONDLOGINNAME, gsPARAMETERTYPE_COLUMNID, mvar_lngSecondLoginNameID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_GRADE, gsPARAMETERTYPE_COLUMNID, mvar_lngGradeID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_MANAGERSTAFFNO, gsPARAMETERTYPE_COLUMNID, mvar_lngManagerStaffNoID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECURITYGROUP, gsPARAMETERTYPE_COLUMNID, mvar_lngSecurityGroupID
  
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_JOBTITLE, gsPARAMETERTYPE_COLUMNID, mvar_lngJobTitleID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_REGION, gsPARAMETERTYPE_COLUMNID, mvar_lngRegionID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE, gsPARAMETERTYPE_TABLEID, mvar_lngHRegionTableID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONFIELD, gsPARAMETERTYPE_COLUMNID, mvar_lngHRegionFieldID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONDATE, gsPARAMETERTYPE_COLUMNID, mvar_lngHRegionDateID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN, gsPARAMETERTYPE_COLUMNID, mvar_lngWorkingPatternID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE, gsPARAMETERTYPE_TABLEID, mvar_lngHWorkingPatternTableID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD, gsPARAMETERTYPE_COLUMNID, mvar_lngHWorkingPatternFieldID
  SaveModuleSetting gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE, gsPARAMETERTYPE_COLUMNID, mvar_lngHWorkingPatternDateID

  SaveModuleSetting gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE, gsPARAMETERTYPE_TABLEID, mvar_lngHierarchyTableID
  SaveModuleSetting gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_IDENTIFIER, gsPARAMETERTYPE_COLUMNID, mvar_lngHierarchyIdentifyingColumnID
  SaveModuleSetting gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_REPORTSTO, gsPARAMETERTYPE_COLUMNID, mvar_lngHierarchyReportsToColumnID
  SaveModuleSetting gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCATIONTABLE, gsPARAMETERTYPE_TABLEID, mvar_lngPostAllocationTableID
  SaveModuleSetting gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCSTARTDATE, gsPARAMETERTYPE_COLUMNID, mvar_lngPostAllocationStartDateColumnID
  SaveModuleSetting gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCENDDATE, gsPARAMETERTYPE_COLUMNID, mvar_lngPostAllocationEndDateColumnID

  'NPG20090225 Fault 13502
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_LOGINNAME, gsPARAMETERTYPE_COLUMNID, mvar_lngLoginNameID
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_SECONDLOGINNAME, gsPARAMETERTYPE_COLUMNID, mvar_lngSecondLoginNameID

  ' Update Mobile config
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_WORKEMAIL, gsPARAMETERTYPE_COLUMNID, mvar_lngWorkEmailID
    
  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault
End Function

Private Sub RefreshHierarchyControls()

  ' Refresh the Hierarchy controls
  Dim iHierarchyIdentifyingColumnIndex As Integer
  Dim iHierarchyReportsToColumnIndex As Integer
  Dim iPostAllocationTableListIndex As Integer

  iHierarchyIdentifyingColumnIndex = 0
  iHierarchyReportsToColumnIndex = 0
  iPostAllocationTableListIndex = 0
  
  ' Clear the current contents of the Identifying Column combo
  With cboHierarchyIdentifyingColumn
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboHierarchyReportsToColumn
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngHierarchyTableID

    If Not .NoMatch Then
      ' Add non system/link cols to the combo that have not been deleted
      Do While Not .EOF
        If !TableID <> mvar_lngHierarchyTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          ' Load varchar fields
          If (!DataType = dtVARCHAR) Or _
            (!DataType = dtNUMERIC) Or _
            (!DataType = dtINTEGER) Then
            cboHierarchyIdentifyingColumn.AddItem !ColumnName
            cboHierarchyIdentifyingColumn.ItemData(cboHierarchyIdentifyingColumn.NewIndex) = !ColumnID
            
            If !ColumnID = mvar_lngHierarchyIdentifyingColumnID Then
              iHierarchyIdentifyingColumnIndex = cboHierarchyIdentifyingColumn.NewIndex
            End If
          
            cboHierarchyReportsToColumn.AddItem !ColumnName
            cboHierarchyReportsToColumn.ItemData(cboHierarchyReportsToColumn.NewIndex) = !ColumnID
            
            If !ColumnID = mvar_lngHierarchyReportsToColumnID Then
              iHierarchyReportsToColumnIndex = cboHierarchyReportsToColumn.NewIndex
            End If
          End If
        End If
       
        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboHierarchyIdentifyingColumn.ListIndex = iHierarchyIdentifyingColumnIndex
  cboHierarchyReportsToColumn.ListIndex = iHierarchyReportsToColumnIndex

  ' Clear the current contents of the combo
  With cboPostAllocationTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With

  ' Add the tables that are children of the Hierarchy table AND
  ' the Personnel table
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If Not !Deleted Then
        recRelEdit.Index = "idxParentID"
        recRelEdit.Seek "=", mvar_lngHierarchyTableID, !TableID

        If Not recRelEdit.NoMatch Then
          recRelEdit.Seek "=", mvar_lngPersonnelTableID, !TableID
          
          If Not recRelEdit.NoMatch Then
            cboPostAllocationTable.AddItem !TableName
            cboPostAllocationTable.ItemData(cboPostAllocationTable.NewIndex) = !TableID

            If !TableID = mvar_lngPostAllocationTableID Then
              iPostAllocationTableListIndex = cboPostAllocationTable.NewIndex
            End If
          End If
        End If

      End If
      .MoveNext
    Loop
  End With

  ' Select the appropriate combo items.
  cboPostAllocationTable.ListIndex = iPostAllocationTableListIndex

End Sub



Private Sub RefreshPostAllocationControls()

  ' Refresh the Hierarchy controls
  Dim iPostAllocationStartDateColumnIndex As Integer
  Dim iPostAllocationEndDateColumnIndex As Integer

  iPostAllocationStartDateColumnIndex = 0
  iPostAllocationEndDateColumnIndex = 0
  
  ' Clear the current contents of the Identifying Column combo
  With cboPostAllocationStartDateColumn
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboPostAllocationEndDateColumn
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngPostAllocationTableID

    If Not .NoMatch Then
      ' Add non system/link cols to the combo that have not been deleted
      Do While Not .EOF
        If !TableID <> mvar_lngPostAllocationTableID Then
          Exit Do
        End If
        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          ' Load date fields
          If !DataType = dtTIMESTAMP Then
            cboPostAllocationStartDateColumn.AddItem !ColumnName
            cboPostAllocationStartDateColumn.ItemData(cboPostAllocationStartDateColumn.NewIndex) = !ColumnID
            
            If !ColumnID = mvar_lngPostAllocationStartDateColumnID Then
              iPostAllocationStartDateColumnIndex = cboPostAllocationStartDateColumn.NewIndex
            End If
          
            cboPostAllocationEndDateColumn.AddItem !ColumnName
            cboPostAllocationEndDateColumn.ItemData(cboPostAllocationEndDateColumn.NewIndex) = !ColumnID
            
            If !ColumnID = mvar_lngPostAllocationEndDateColumnID Then
              iPostAllocationEndDateColumnIndex = cboPostAllocationEndDateColumn.NewIndex
            End If
          End If
        End If
       
        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboPostAllocationStartDateColumn.ListIndex = iPostAllocationStartDateColumnIndex
  cboPostAllocationEndDateColumn.ListIndex = iPostAllocationEndDateColumnIndex

End Sub




Private Sub ReadParameters()
  ' Read the Personnel parameter values from the database into local variables.
  
  mvar_lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
  mvar_lngEmployeeNumberID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER, 0)
  mvar_lngForenameID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME, 0)
  mvar_lngSurnameID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME, 0)
  mvar_lngSSIWelcomeID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SSIWELCOME, 0)
  mvar_lngSSIPhotographID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SSIPHOTOGRAPH, 0)
  mvar_lngStartDateID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_STARTDATE, 0)
  mvar_lngLeavingDateID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE, 0)
  mvar_lngFullPartTimeID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FULLPARTTIME, 0)
  mvar_lngWorkEmailID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKEMAIL, 0)
  mvar_lngDepartmentID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT, 0)
  mvar_lngDateOfBirthID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DATEOFBIRTH, 0)
  
  mvar_lngLoginNameID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME, 0)
  mvar_lngSecondLoginNameID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECONDLOGINNAME, 0)
  If (mvar_lngLoginNameID = 0) _
    And (mvar_lngSecondLoginNameID > 0) Then
    
    mvar_lngLoginNameID = mvar_lngSecondLoginNameID
    mvar_lngSecondLoginNameID = 0
  End If
  
  mvar_lngGradeID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_GRADE, 0)
  mvar_lngManagerStaffNoID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_MANAGERSTAFFNO, 0)
  mvar_lngSecurityGroupID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECURITYGROUP, 0)
  mvar_lngJobTitleID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_JOBTITLE, 0)
  mvar_lngRegionID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_REGION, 0)
  mvar_lngHRegionTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE, 0)
  mvar_lngHRegionFieldID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONFIELD, 0)
  mvar_lngHRegionDateID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONDATE, 0)
  mvar_lngWorkingPatternID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN, 0)
  mvar_lngHWorkingPatternTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE, 0)
  mvar_lngHWorkingPatternFieldID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD, 0)
  mvar_lngHWorkingPatternDateID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE, 0)
  
  mvar_lngHierarchyTableID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE, 0)
  mvar_lngHierarchyIdentifyingColumnID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_IDENTIFIER, 0)
  mvar_lngHierarchyReportsToColumnID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_REPORTSTO, 0)
  mvar_lngPostAllocationTableID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCATIONTABLE, 0)
  mvar_lngPostAllocationStartDateColumnID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCSTARTDATE, 0)
  mvar_lngPostAllocationEndDateColumnID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_POSTALLOCENDDATE, 0)
  
  ' Remember some of the original values.
  mvar_lngOriginalPersonnelTableID = mvar_lngPersonnelTableID
  mvar_lngOriginalHierarchyTableID = mvar_lngHierarchyTableID
  mvar_lngOriginalHierarchyIdentifyingColumnID = mvar_lngHierarchyIdentifyingColumnID
  mvar_lngOriginalHierarchyReportsToColumnID = mvar_lngHierarchyReportsToColumnID
  
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optRegionStatic_Click()

  ' Unset all the historic region combos and disable them
  
  With cboRegionTable
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  With cboRegionField
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  With cboRegionEffectiveDate
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  lblHistorialRegionColumn.Enabled = False
  lblHistorialRegionEffectiveColumn.Enabled = False
  lblHistoricalRegionTable.Enabled = False
  'lblRegion.Enabled = True
  lblRegion.Enabled = Not mblnReadOnly
  
  ' Enable the static region combo
  With cboRegion
    '.Enabled = True
    .Enabled = Not mblnReadOnly
    .BackColor = vbWindowBackground
  End With

  mvar_lngHRegionTableID = 0
  mvar_lngHRegionFieldID = 0
  mvar_lngHRegionDateID = 0
  Changed = True
End Sub


Private Sub optRegionHistory_Click()
  
  'Unset the static region combo and disable it
  With cboRegion
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  lblRegion.Enabled = False
  
  If Not mblnReadOnly Then
  
    ' Enable the historic region combos
    With cboRegionTable
      .Enabled = True
      .BackColor = vbWindowBackground
    End With
    
    With cboRegionField
      .Enabled = True
      .BackColor = vbWindowBackground
    End With
    
    With cboRegionEffectiveDate
      .Enabled = True
      .BackColor = vbWindowBackground
    End With
    
    lblHistorialRegionColumn.Enabled = True
    lblHistorialRegionEffectiveColumn.Enabled = True
    lblHistoricalRegionTable.Enabled = True
  
  End If
  
  mvar_lngRegionID = 0
  Changed = True
End Sub


Private Sub optWorkingPatternStatic_Click()

  ' Unset all the historic wpattern combos and disable them
  With cboWorkingPatternTable
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  With cboWorkingPatternField
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  With cboWorkingPatternEffectiveDate
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  lblHistorialWorkingPatternColumn.Enabled = False
  lblHistorialWorkingPatternEffectiveColumn.Enabled = False
  lblHistoricalWorkingPatternTable.Enabled = False
  
  ' Enable the static wpattern combo
  With cboWorkingPattern
    '.Enabled = True
    .Enabled = Not mblnReadOnly
    .BackColor = vbWindowBackground
  End With

  'lblWorkingPattern.Enabled = True
  lblWorkingPattern.Enabled = Not mblnReadOnly

  mvar_lngHWorkingPatternTableID = 0
  mvar_lngHWorkingPatternFieldID = 0
  mvar_lngHWorkingPatternDateID = 0
  Changed = True
End Sub


Private Sub optWorkingPatternHistory_Click()

  'Unset the static wpattern combo and disable it
  With cboWorkingPattern
    .Enabled = False
    .BackColor = vbButtonFace
    .ListIndex = 0
  End With
  
  lblWorkingPattern.Enabled = False
  
  If Not mblnReadOnly Then
  
    ' Enable the historic wpattern combos
    With cboWorkingPatternTable
      .Enabled = True
      .BackColor = vbWindowBackground
    End With
    
    With cboWorkingPatternField
      .Enabled = True
      .BackColor = vbWindowBackground
    End With
    
    With cboWorkingPatternEffectiveDate
      .Enabled = True
      .BackColor = vbWindowBackground
    End With
    
    lblHistorialWorkingPatternColumn.Enabled = True
    lblHistorialWorkingPatternEffectiveColumn.Enabled = True
    lblHistoricalWorkingPatternTable.Enabled = True
  
  End If
  
  mvar_lngWorkingPatternID = 0
  Changed = True
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
  If Not mblnReadOnly Then
    
    fraTableDefinition.Enabled = (SSTab1.Tab = giPAGE_PERSONNEL)
    fraRegion.Enabled = (SSTab1.Tab = giPAGE_CAREER_CHANGE)
    fraWorkingPattern.Enabled = (SSTab1.Tab = giPAGE_CAREER_CHANGE)
    fraHierarchyTable.Enabled = (SSTab1.Tab = giPAGE_HIERARCHY)
    fraPostAllocationTable.Enabled = (SSTab1.Tab = giPAGE_HIERARCHY)
  End If
  
End Sub

Private Sub SSTab1_DblClick()

  If Not mblnReadOnly Then
    Me.fraTableDefinition.Enabled = (SSTab1.Tab = 0)
    Me.fraRegion.Enabled = (SSTab1.Tab = 1)
    Me.fraWorkingPattern.Enabled = (SSTab1.Tab = 1)
  End If

End Sub
