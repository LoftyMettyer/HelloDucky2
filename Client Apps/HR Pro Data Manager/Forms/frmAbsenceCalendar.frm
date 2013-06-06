VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{96E404DC-B217-4A2D-A891-C73A92A628CC}#1.0#0"; "COA_WorkingPattern.ocx"
Begin VB.Form frmAbsenceCalendar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Absence Calendar "
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   HelpContextID   =   1005
   Icon            =   "frmAbsenceCalendar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10200
      TabIndex        =   34
      Top             =   7080
      Width           =   1200
   End
   Begin VB.CommandButton cmdYearAdd 
      DisabledPicture =   "frmAbsenceCalendar.frx":000C
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11340
      Picture         =   "frmAbsenceCalendar.frx":03BF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   285
   End
   Begin VB.CommandButton cmdYearSubtract 
      DisabledPicture =   "frmAbsenceCalendar.frx":077A
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      Picture         =   "frmAbsenceCalendar.frx":0B29
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   285
   End
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   11295
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   435
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   165
      TabIndex        =   0
      Top             =   7125
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Height          =   1425
      Left            =   240
      TabIndex        =   27
      Top             =   7440
      Visible         =   0   'False
      Width           =   8520
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   8
      BalloonHelp     =   0   'False
      RowHeight       =   423
      Columns.Count   =   8
      Columns(0).Width=   2090
      Columns(0).Caption=   "S Date"
      Columns(0).Name =   "ug"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1164
      Columns(1).Caption=   "S Sess"
      Columns(1).Name =   "rt"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2117
      Columns(2).Caption=   "E Date"
      Columns(2).Name =   "rtgh"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1111
      Columns(3).Caption=   "E Sess"
      Columns(3).Name =   "rtewt"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1931
      Columns(4).Caption=   "Type"
      Columns(4).Name =   "ewtwetg"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1296
      Columns(5).Caption=   "CalCode"
      Columns(5).Name =   "fdhdfhd"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1561
      Columns(6).Caption=   "TypeCode"
      Columns(6).Name =   "TypeCode"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "Reason"
      Columns(7).Name =   "Reason"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   15028
      _ExtentY        =   2514
      _StockProps     =   79
      Caption         =   "SSDBGrid1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Frame fraEmployeeInformation 
      Caption         =   "Employee Information :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   9000
      TabIndex        =   23
      Top             =   480
      Width           =   2700
      Begin COAWorkingPattern.COA_WorkingPattern ASRWorkingPattern1 
         Height          =   765
         Left            =   330
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   840
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   1349
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaving Date :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   645
         Width           =   1365
      End
      Begin VB.Label lblEndDate 
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label lblStartDateLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   32
         Top             =   450
         Width           =   1170
      End
      Begin VB.Label lblRegionLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblRegion 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   450
         Width           =   1245
      End
   End
   Begin VB.Frame fraColourKey 
      Caption         =   "Key : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   9000
      TabIndex        =   15
      Top             =   4200
      Width           =   2700
      Begin VB.Label lblColourKey_Type 
         BackStyle       =   0  'Transparent
         Caption         =   "WWWWWW"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9999
         Left            =   645
         TabIndex        =   17
         Top             =   135
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblColourKey_Colour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   9999
         Left            =   435
         TabIndex        =   16
         Top             =   150
         Visible         =   0   'False
         Width           =   220
      End
   End
   Begin VB.Frame fraOptionsShade 
      Caption         =   "Options :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   9000
      TabIndex        =   14
      Top             =   2145
      Width           =   2700
      Begin VB.ComboBox cboStartMonth 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmAbsenceCalendar.frx":0EDD
         Left            =   1290
         List            =   "frmAbsenceCalendar.frx":0EDF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   1320
      End
      Begin VB.CheckBox chkCaptions 
         Caption         =   "Show Calendar Cap&tions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   1410
         Width           =   2445
      End
      Begin VB.CheckBox chkIncludeWorkingDaysOnly 
         Caption         =   "&Working Days Only"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   870
         Width           =   2235
      End
      Begin VB.CheckBox chkIncludeBHols 
         Caption         =   "Include &Bank Holidays"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   2250
      End
      Begin VB.CheckBox chkShadeBHols 
         Caption         =   "Show Bank &Holidays"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   1140
         Width           =   2235
      End
      Begin VB.CheckBox chkShadeWeekends 
         Caption         =   "Show Wee&kends"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   1670
         Width           =   2100
      End
      Begin VB.Label lblStartMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Month :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.Frame fraAbsenceCalendar 
      BorderStyle     =   0  'None
      Caption         =   "Absence Calendar :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   0
      TabIndex        =   9
      Top             =   -135
      Width           =   8880
      Begin VB.Line linHorizontal 
         Index           =   0
         Visible         =   0   'False
         X1              =   6030
         X2              =   7710
         Y1              =   7065
         Y2              =   7065
      End
      Begin VB.Line linVertical 
         Index           =   0
         Visible         =   0   'False
         X1              =   7845
         X2              =   7845
         Y1              =   4965
         Y2              =   6570
      End
      Begin VB.Label lblCalDates 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   7725
         TabIndex        =   13
         Top             =   6660
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblCal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9999
         Left            =   7710
         TabIndex        =   12
         Top             =   6885
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BB3C3C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1130
         TabIndex        =   11
         Tag             =   "2"
         Top             =   270
         Width           =   210
      End
      Begin VB.Label lblMonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BB3C3C&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   555
         Index           =   1
         Left            =   130
         TabIndex        =   10
         Top             =   465
         Width           =   1005
      End
   End
   Begin VB.Label lblCurrentYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   9600
      TabIndex        =   18
      Top             =   165
      Width           =   570
   End
   Begin VB.Label lblDash 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   10290
      TabIndex        =   22
      Top             =   165
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblCurrentYear2 
      BackStyle       =   0  'Transparent
      Caption         =   "1999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   10575
      TabIndex        =   21
      Top             =   165
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblFirstDate 
      BackStyle       =   0  'Transparent
      Caption         =   "First Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblLastDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1665
      TabIndex        =   20
      Top             =   7050
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "frmAbsenceCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form printing stuff
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2

'AE20071108 Fault #12547
Private Const KEY_FONTSIZE_SMALL = 5
Private Const KEY_FONTSIZE_NORMAL = 6.75

' Flag to prevent redraw of calendar when checkbox options are being set
Private mblnLoading As Boolean

' Holds the absence information for the current employee
Private mrstAbsenceRecords As Recordset

' Indicates if user has permission to see (and therefore use) bank holidays
'Private mblnBankHolsPermission As Boolean
'Private mblnBankHolidaysEnabled As Boolean

Public mlngPersonnelID As Long
Private mdtLeavingDate As Date

Private mfCanReadAbsenceDuration As Boolean

Private mblnDisableRegions As Boolean
Private mblnDisableWPs As Boolean
Private mblnFailReport As Boolean

Private mstrSQLSelect_RegInfoRegion As String
Private mstrSQLSelect_BankHolDate As String
Private mstrSQLSelect_BankHolDesc As String

Private mstrSQLSelect_PersonnelStaticRegion As String
Private mstrSQLSelect_PersonnelStaticWP As String
Private mstrSQLSelect_PersonnelHRegion As String
Private mstrSQLSelect_PersonnelHDate As String


Private mstrSQLSelect_AbsenceStartDate As String
Private mstrSQLSelect_AbsenceStartSession As String
Private mstrSQLSelect_AbsenceEndDate As String
Private mstrSQLSelect_AbsenceEndSession As String
Private mstrSQLSelect_AbsenceType As String
Private mstrSQLSelect_AbsenceReason As String
Private mstrSQLSelect_AbsenceDuration As String

Private mstrSQLSelect_AbsenceTypeType As String
Private mstrSQLSelect_AbsenceTypeCode As String
Private mstrSQLSelect_AbsenceTypeCalCode As String

Private mstrSQLSelect_PersonnelStartDate As String
Private mstrSQLSelect_PersonnelLeavingDate As String

Private mstrBaseTableName As String

Private mvarTableViews() As Variant
Private mobjTableView As CTablePrivilege
Private mobjColumnPrivileges As CColumnPrivileges
Private mstrTempRealSource As String
Private mstrRealSource As String
Private mstrViews() As String

Private mblnShowBankHols As Boolean
Private mblnShowCaptions As Boolean
Private mblnShowWeekends As Boolean

Private mblnRegions As Boolean
Private mblnWorkingPatterns As Boolean

Private mstrAbsenceTableRealSource As String
Private mintMultipleKeyControlIndex As Integer

Public Function AbsCal_IsDayABankHoliday(intIndex As Integer) As Boolean

  ' This function returns true if the date of the index passed to it is defined
  ' as a bank holiday for the current employee.
  Dim dtmCurrentDate As Date
  Dim strRegionAtCurrentDate As String
  Dim rstBankHolRegion As Recordset
  Dim strSQL As String
  Dim lngCount As Long
  
  On Error GoTo AbsCal_IsDayABankHolidayERROR
  
  
  '31/07/2001 MH Fault 2477
  If Not Me.RegionsEnabled Then
    AbsCal_IsDayABankHoliday = False
    Exit Function
  End If
  
  
  dtmCurrentDate = GetCalDay(intIndex)
  
  If Me.RegionsEnabled Then
    ' If we are using historic region, get the employees region on this day
    If grtRegionType = rtHistoricRegion Then
        
      ' NB : have to format the date to mm/dd/yy here otherwise sql doesnt like it
      strSQL = "SELECT " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                                          "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                                         "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & frmAbsenceCalendar.PersonnelRecordID & " " & _
                                                         "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " <= '" & Replace(Format(dtmCurrentDate, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                         "ORDER BY " & gsPersonnelHRegionDateColumnName & " DESC"
      Set rstBankHolRegion = datGeneral.GetRecords(strSQL)
                
      If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then
        AbsCal_IsDayABankHoliday = False
        Set rstBankHolRegion = Nothing
        Exit Function
      Else
        strRegionAtCurrentDate = rstBankHolRegion.Fields("Region").Value
      End If
  
    Else
    
      strRegionAtCurrentDate = frmAbsenceCalendar.lblRegion.Caption
    
    End If
  
    strSQL = vbNullString
    strSQL = strSQL & "SELECT " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " AS 'Date' " & vbCrLf
    strSQL = strSQL & "FROM " & gsBHolTableRealSource & " " & vbCrLf
    strSQL = strSQL & "WHERE " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & " = " & vbCrLf
    strSQL = strSQL & "        (SELECT " & gsBHolRegionTableName & ".ID " & vbCrLf
    strSQL = strSQL & "         FROM " & gsBHolRegionTableName & vbCrLf
    For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
      '<REGIONAL CODE>
      If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
        strSQL = strSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbCrLf
        strSQL = strSQL & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbCrLf
      End If
    Next lngCount
    strSQL = strSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbCrLf
    strSQL = strSQL & " AND " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " = '" & Replace(Format(dtmCurrentDate, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & vbCrLf
    strSQL = strSQL & "ORDER BY " & gsBHolDateColumnName & " ASC" & vbCrLf
    Set rstBankHolRegion = datGeneral.GetRecords(strSQL)
    If Not rstBankHolRegion.EOF Then
      AbsCal_IsDayABankHoliday = True
      Set rstBankHolRegion = Nothing
      Exit Function
    End If
  
  End If 'Me.RegionsEnabled
  
  Exit Function
  
AbsCal_IsDayABankHolidayERROR:
  
  COAMsgBox "Error whilst checking for bank holidays." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  AbsCal_IsDayABankHoliday = False
  Set rstBankHolRegion = Nothing

End Function

Public Function DisableUnused() As Boolean

  ' This function shades the caldate and cal labels that arent used for the current year.
  
  On Error Resume Next    ' THIS IS ESSENTIAL - DO NOT REMOVE
  
  Dim objCtl As Control
  Dim pintTemp As Integer
  Dim pintTempHundred As Integer
  Dim strRegionAtCurrentDate As String
  Dim dtmCurrentDate As Date
  Dim dtmNextChangeDate As Date
  Dim rstBankHolRegion As Recordset
  Dim pintStartIndex As Integer
  Dim pdtmStartDate As Date
  Dim pintEndIndex As Integer
  Dim pdtmEndDate As Date
  Dim pintCount As Integer
  Dim dtCurrentDate As Date
  Dim blnOutsideEmployment As Boolean
  Dim bNewRegionFound As Boolean
  Dim iBankHolidayCell As Integer
  Dim sSQL As String
  Dim lngCount As Long
  
  ' Loop through the controls
  For Each objCtl In Me.Controls

    If objCtl.Name = "lblCalDates" Then
      If objCtl.Caption = "" Then
        If objCtl.Index > 0 Then
          ' CalDates
          Me.lblCalDates(objCtl.Index).BackColor = Me.lblCalDates(objCtl.Index).BackColor - &H333333
          ' Cal
          pintTempHundred = CInt(objCtl.Index / 100)
          pintTemp = objCtl.Index
          Me.lblCal((pintTemp * 2) - (pintTempHundred * 100)).BackColor = 13389133   'Me.lblCal((pintTemp * 2) - (pintTempHundred * 100)).BackColor - &H333333 'vbBlack
          Me.lblCal(((pintTemp * 2) - (pintTempHundred * 100)) - 1).BackColor = 13389133  'Me.lblCal(((pintTemp * 2) - (pintTempHundred * 100)) - 1).BackColor - &H333333 'vbblack
        End If
      End If
    End If
  
    If objCtl.Name = "lblCal" Then
      If objCtl.Index <> 9999 Then

        dtCurrentDate = GetCalDay(objCtl.Index)
        blnOutsideEmployment = False
        
        With Me
          If .lblStartDate.Caption <> "" Then
            blnOutsideEmployment = blnOutsideEmployment Or _
              (DateDiff("d", .lblStartDate.Caption, dtCurrentDate) < 0)
          End If
          If .lblEndDate.Caption <> "" Then
            blnOutsideEmployment = blnOutsideEmployment Or _
              (DateDiff("d", .lblEndDate.Caption, dtCurrentDate) > 0)
          End If
        End With
        
        
        '09/08/2001 MH Fault 2608
        If blnOutsideEmployment Then
            'Shade days outside before start date and after end date
            objCtl.BackColor = objCtl.BackColor - &H333333
        
        Else
          ' Shade the weekends if the option has been set
          If Me.ShowWeekends Then
            
            If Weekday(dtCurrentDate) = vbSaturday Or Weekday(dtCurrentDate) = vbSunday Then
              objCtl.BackColor = glngColour_Weekend
            End If
          
          End If
        End If
      
      End If
    End If
    
  Next objCtl
  
  
  If Me.RegionsEnabled Then
 
    ' Show bhols if option is set
    If Me.ShowBankHolidays Then

      bNewRegionFound = False
      
      ' If we are using historic region, find the region change dates
      If grtRegionType = rtHistoricRegion Then
        ' DONE
        ' Get the first region for this employee within this calendar year
        Set rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                                     "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                                     "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & Me.PersonnelRecordID & " " & _
                                                     "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " <= '" & Replace(Format(Me.lblFirstDate.Caption, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                     "ORDER BY " & gsPersonnelHRegionDateColumnName & " DESC")
        
        ' Was there a region at the start of the calendar
        If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then
          strRegionAtCurrentDate = ""
        Else
          strRegionAtCurrentDate = rstBankHolRegion.Fields("Region").Value
          bNewRegionFound = True
        End If
          
        ' DONE
        ' Get the second region for this employee within this calendar year
        Set rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                                     "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                                     "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & Me.PersonnelRecordID & " " & _
                                                     "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " > '" & Replace(Format(Me.lblFirstDate.Caption, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                     "ORDER BY " & gsPersonnelHRegionDateColumnName & " ASC")
        
        ' Was there a region at the start of the calendar
        If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then
          'JPD 20030828 Fault 2012
          'dtmNextChangeDate = CDate("31/12/9999")
          dtmNextChangeDate = CDate(ConvertSQLDateToLocale("12/31/9999"))
        Else
          dtmNextChangeDate = rstBankHolRegion.Fields("Date").Value
        End If
  
        ' Get the start and end date/indexes for the currently displayed year
        pdtmStartDate = Me.lblFirstDate.Caption
        pintStartIndex = GetCalIndex(pdtmStartDate, False)
        pdtmEndDate = Me.lblLastDate.Caption
        pintEndIndex = GetCalIndex(pdtmEndDate, True)
        
        ' Because we are working with whole days, we can go through the indexes
        ' in twos, but to do this, we need to ensure that pintEndIndex is even.
        If pintEndIndex Mod 2 = 1 Then
          pintEndIndex = pintEndIndex + 1
        End If
        
        For pintCount = pintStartIndex To pintEndIndex Step 2
        
          ' Get the date of the current index
          dtmCurrentDate = GetCalDay(pintCount)
          
          ' Only refer to the region table if the current date is a region change date
          If (dtmCurrentDate >= dtmNextChangeDate) And (dtmCurrentDate <> ConvertSQLDateToLocale("12/31/9999")) Then
          
            
            'JDM - 11/09/01 - Fault 2820 - Bank hols not showing for year starting with working pattern.
            ' Find the employees region for this date
            ' DONE
            Set rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                                         "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                                         "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & Me.PersonnelRecordID & " " & _
                                                         "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " >= '" & Replace(Format(dtmNextChangeDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                         "ORDER BY " & gsPersonnelHRegionDateColumnName & " ASC")
                                                         
            If rstBankHolRegion.BOF And rstBankHolRegion.EOF Then
            
              ' No regions found for this user
              'JPD 20030828 Fault 2012
              'dtmNextChangeDate = CDate("31/12/9999")
              dtmNextChangeDate = CDate(ConvertSQLDateToLocale("12/31/9999"))
              
            Else
              
              strRegionAtCurrentDate = rstBankHolRegion.Fields("Region").Value
              bNewRegionFound = True
            
              ' Now get the next change date
              ' DONE
              Set rstBankHolRegion = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " AS 'Date', " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                                           "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                                           "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & Me.PersonnelRecordID & " " & _
                                                           "AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " > '" & Replace(Format(rstBankHolRegion.Fields("Date").Value, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                           "ORDER BY " & gsPersonnelHRegionDateColumnName & " ASC")
              If rstBankHolRegion.EOF Then
                dtmNextChangeDate = ConvertSQLDateToLocale("12/31/9999")
              Else
                dtmNextChangeDate = rstBankHolRegion.Fields("Date").Value
              End If
              
            End If
  
          End If '(dtmCurrentDate >= dtmNextChangeDate) And (dtmCurrentDate <> ConvertSQLDateToLocale("12/31/9999"))
          
          ' If current region has changed
          If bNewRegionFound Then
            
            ' Get bank holidays for this region
            ' DONE
            sSQL = vbNullString
            sSQL = sSQL & "SELECT " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " AS 'Date' " & vbCrLf
            sSQL = sSQL & "FROM " & gsBHolTableRealSource & " " & vbCrLf
            
            sSQL = sSQL & "WHERE " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & " = " & vbCrLf
            sSQL = sSQL & "        (SELECT " & gsBHolRegionTableName & ".ID " & vbCrLf
            sSQL = sSQL & "         FROM " & gsBHolRegionTableName & vbCrLf
            For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
              '<REGIONAL CODE>
              If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
                sSQL = sSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbCrLf
                sSQL = sSQL & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbCrLf
              End If
            Next lngCount
            sSQL = sSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbCrLf
            
            sSQL = sSQL & " AND " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & Replace(Format(dtmCurrentDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "' " & vbCrLf
            sSQL = sSQL & " AND " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & Replace(Format(dtmNextChangeDate - 1, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "' " & vbCrLf
            sSQL = sSQL & "ORDER BY " & gsBHolDateColumnName & " ASC"
            Set rstBankHolRegion = datGeneral.GetRecords(sSQL)
          
            ' Cycle through the recordset checking for the current day
            If Not (rstBankHolRegion.BOF And rstBankHolRegion.EOF) Then
              rstBankHolRegion.MoveFirst
              Do Until rstBankHolRegion.EOF
                iBankHolidayCell = GetCalIndex(rstBankHolRegion.Fields("Date").Value, False)
                'JPD 20030828 Fault 6493
                If iBankHolidayCell <> 9999 Then
                  Me.lblCal(iBankHolidayCell).BackColor = glngColour_BankHoliday
                  Me.lblCal(iBankHolidayCell + 1).BackColor = glngColour_BankHoliday
                End If
                rstBankHolRegion.MoveNext
              Loop
              
              ' Flag this region has had it's bank holidays drawn
              bNewRegionFound = False
            End If 'Not (rstBankHolRegion.BOF And rstBankHolRegion.EOF)
          
          End If 'bNewRegionFound
            
        Next pintCount
        
      Else 'grtRegionType = rtHistoricRegion
        
        ' We are using a static region so just use the employees current region
        strRegionAtCurrentDate = Me.lblRegion.Caption
        ' DONE
        sSQL = vbNullString
        sSQL = sSQL & "SELECT " & gsBHolTableRealSource & "." & gsBHolDateColumnName & " AS 'Date' " & vbCrLf
        sSQL = sSQL & "FROM " & gsBHolTableRealSource & " " & vbCrLf
        sSQL = sSQL & "WHERE " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & " = " & vbCrLf
        sSQL = sSQL & "        (SELECT " & gsBHolRegionTableName & ".ID " & vbCrLf
        sSQL = sSQL & "         FROM " & gsBHolRegionTableName & vbCrLf
        For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
          '<REGIONAL CODE>
          If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
            sSQL = sSQL & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbCrLf
            sSQL = sSQL & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbCrLf
          End If
        Next lngCount
        sSQL = sSQL & "         WHERE " & mstrSQLSelect_RegInfoRegion & " = '" & strRegionAtCurrentDate & "') " & vbCrLf
        sSQL = sSQL & "ORDER BY " & gsBHolDateColumnName & " ASC" & vbCrLf
        
        Set rstBankHolRegion = datGeneral.GetRecords(sSQL)
        
        ' Cycle through the recordset checking for the current day
        If Not (rstBankHolRegion.BOF And rstBankHolRegion.EOF) Then
          rstBankHolRegion.MoveFirst
          Do Until rstBankHolRegion.EOF
            pintTemp = GetCalIndex(CDate(rstBankHolRegion.Fields("Date").Value), False)
            If pintTemp <> 9999 Then
'              Me.lblCal(pintTemp).BackColor = Me.lblCal(9999).BackColor - &H333333
'              Me.lblCal(pintTemp + 1).BackColor = Me.lblCal(9999).BackColor - &H333333
              Me.lblCal(pintTemp).BackColor = glngColour_BankHoliday
              Me.lblCal(pintTemp + 1).BackColor = glngColour_BankHoliday
            End If
            rstBankHolRegion.MoveNext
          Loop
        End If
        
      End If 'grtRegionType = rtHistoricRegion
        
    End If 'If Me.ShowBankHolidays Then
    
  End If 'Me.RegionsEnabled
  
  DisableUnused = True
  
End Function

Public Property Get AbsenceTableRealSource() As String
  AbsenceTableRealSource = mstrAbsenceTableRealSource
End Property

Public Property Get PersonnelRecordID() As String
  PersonnelRecordID = mlngPersonnelID
End Property

Public Property Get RegionsEnabled() As Boolean
  RegionsEnabled = Not mblnDisableRegions
End Property

Public Property Get WPsEnabled() As Boolean
  WPsEnabled = Not mblnDisableWPs
End Property

Private Function CheckPermission_AbsCalSpecifics() As Boolean

  Dim strTableColumn As String
  Dim strModulePermErrorMSG As String
  
  strModulePermErrorMSG = vbNullString
  
  'Check Module Setup on each of the module columns.
  '
  '                       II
  '                       II
  '                       II
  '                       II
  '                    \  II  /
  '                     \ II /
  '                      \II/
  '                       \/
  
  'Check the Absence Table
  '          Absence Table - Start Date Column
  '          Absence Table - Start Session Column
  '          Absence Table - End Date Column
  '          Absence Table - End Session Column
  '          Absence Table - Absence Type Column
  '          Absence Table - Absence Reason Column
  '          Absence Table - Absence Duration Column
  '...Absence module setup information.
  'If any are blank then we need to fail the Absence Calendar report.
  If gsAbsenceTableName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Table' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceStartDateColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Start Date Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceStartSessionColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Start Session Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceEndDateColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'End Date Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceEndSessionColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'End Session Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceTypeColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Type Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceReasonColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Reason Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceDurationColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Duration Column' in the Absence module setup must be defined." & vbCrLf
  End If
  
 
  'Check the Absence Type Table
  '          Absence Type Table - Absence Type Column
  '          Absence Type Table - Absence Code Column
  '          Absence Type Table - Calendar Code Column
  '...Absence module setup information.
  'If any are blank then we need to fail the Absence Calendar report.
  If gsAbsenceTypeTableName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Type Table' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceTypeTypeColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Type Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceTypeCodeColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Absence Code Column' in the Absence module setup must be defined." & vbCrLf
  End If
  If gsAbsenceTypeCalCodeColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Calendar Code Column' in the Absence module setup must be defined." & vbCrLf
  End If
  
  
  'Check the Personnel Table
  '          Personnel Table - Start Date Column
  '          Personnel Table - Leaving Date Column
  '...Personnel module setup information.
  'If any are blank then we need to fail the Absence Calendar report.
  If gsPersonnelTableName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Personnel Table' in the Personnel module setup must be defined." & vbCrLf
  End If
  If gsPersonnelStartDateColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Start Date Column' in the Personnel module setup must be defined." & vbCrLf
  End If
  If gsPersonnelLeavingDateColumnName = "" Then
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "The 'Leaving Date Column' in the Personnel module setup must be defined." & vbCrLf
  End If
  
  If Len(strModulePermErrorMSG) > 0 Then
    strModulePermErrorMSG = strModulePermErrorMSG & vbCrLf
  End If
  
  If Len(strModulePermErrorMSG) > 0 Then
    strModulePermErrorMSG = "The Absence Calendar failed for the following reasons: " & _
      vbCrLf & vbCrLf & strModulePermErrorMSG
    COAMsgBox strModulePermErrorMSG, vbOKOnly + vbExclamation, "Absence Calendar"
    GoTo FailReport
  End If
  
  'Check Permissions on each of these columns and set the select string for each.
  '
  '                                     II
  '                                     II
  '                                     II
  '                                     II
  '                                  \  II  /
  '                                   \ II /
  '                                    \II/
  '                                     \/

  'Absence Specifics
  'Absence Table - Start Date Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceStartDateColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceStartDate = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - Start Date Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Table - Start Session Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceStartSessionColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceStartSession = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - Start Session Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Table - End Date Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceEndDateColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceEndDate = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - End Date Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Table - End Session Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceEndSessionColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceEndSession = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - End Session Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Table - Absence Type Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceTypeColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceType = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - Absence Type Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Table - Absence Reason Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceReasonColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceReason = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - Absence Reason Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Table - Absence Duration Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTableID, gsAbsenceTableName, _
                            gsAbsenceDurationColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceDuration = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Table - Absence Duration Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  
  'Absence Type Specifics
  'Absence Type Table - Absence Type Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTypeTableID, gsAbsenceTypeTableName, _
                            gsAbsenceTypeTypeColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceTypeType = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Type Table - Absence Type Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Type Table - Absence Code Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTypeTableID, gsAbsenceTypeTableName, _
                            gsAbsenceTypeCodeColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceTypeCode = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Type Table - Absence Code Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Absence Type Table - Calendar Code Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngAbsenceTypeTableID, gsAbsenceTypeTableName, _
                            gsAbsenceTypeCalCodeColumnName, strTableColumn) Then
    mstrSQLSelect_AbsenceTypeCalCode = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Absence Type Table - Calendar Code Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  
  'Personnel Specifics
  'Personnel Table - Start Date Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, _
                            gsPersonnelStartDateColumnName, strTableColumn) Then
    mstrSQLSelect_PersonnelStartDate = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Personnel Table - Start Date Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  'Personnel Table - Leaving Date Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, _
                            gsPersonnelLeavingDateColumnName, strTableColumn) Then
    mstrSQLSelect_PersonnelLeavingDate = strTableColumn
    strTableColumn = vbNullString
  Else
    strModulePermErrorMSG = strModulePermErrorMSG & _
      "Permission Denied on 'Personnel Table - Leaving Date Column'" & vbCrLf
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  If Len(strModulePermErrorMSG) > 0 Then
    strModulePermErrorMSG = "The Absence Calendar failed for the following reasons: " & _
      vbCrLf & vbCrLf & strModulePermErrorMSG
    COAMsgBox strModulePermErrorMSG, vbOKOnly + vbExclamation, "Absence Calendar"
    GoTo FailReport
  End If
  
  CheckPermission_AbsCalSpecifics = True
  
TidyUpAndExit:
  Exit Function

FailReport:
  mblnFailReport = True
  CheckPermission_AbsCalSpecifics = False
  GoTo TidyUpAndExit
  
End Function

Private Function GetInfoFromDB() As Boolean
  
  Dim fOK As Boolean
  Dim blnRegionEnabled As Boolean
  Dim blnWorkingPatternEnabled As Boolean
  
  On Error GoTo ErrorTrap
  
  fOK = True
  
  ' Retrieve the calendars default display options
  If fOK Then fOK = LoadDefaultDisplayOptions
  
  ' Check the Module Setup and Data Permissions for the Absence Calendar Specific columns
  If fOK Then
    fOK = CheckPermission_AbsCalSpecifics
    If Not fOK Then
      GoTo ErrorTrap
      Exit Function
    End If
  End If
  
  ' Check the Module Setup and Data Permissions for the Regional/Bank Holiday columns
  blnRegionEnabled = CheckPermission_RegionInfo
  
  ' Check the Module Setup and Data Permissions for the Working Pattern columns
  blnWorkingPatternEnabled = CheckPermission_WPInfo
  
  ' Get the personnel information
  If fOK Then fOK = GetPersonnelRecordSet

  ' Get the absence recordset
  If fOK Then fOK = GetAbsenceRecordSet
  
  If fOK = False Then Screen.MousePointer = vbDefault

  GetInfoFromDB = fOK
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Screen.MousePointer = vbDefault
  GetInfoFromDB = False
  GoTo TidyUpAndExit
    
End Function

Private Function GetAbsenceRecordSet() As Boolean

  Dim sSQL As String

  On Error GoTo GetAbsenceRecordSet_ERROR
  
  ' Get Recordset Containing Absence info for the current employee
  sSQL = "SELECT " & mstrSQLSelect_AbsenceStartDate & " as 'StartDate', " & vbCrLf & _
    mstrSQLSelect_AbsenceStartSession & " as 'StartSession', " & vbCrLf
    
  If mdtLeavingDate <> CDate(vbNull) Then
    sSQL = sSQL & _
    "isnull(" & mstrSQLSelect_AbsenceEndDate & ",'" & _
    Replace(Format(mdtLeavingDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') as 'EndDate', " & vbCrLf
  Else
    sSQL = sSQL & _
    "isnull(" & mstrSQLSelect_AbsenceEndDate & ",'" & _
    Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') as 'EndDate', " & vbCrLf
  End If
  
  sSQL = sSQL & _
    mstrSQLSelect_AbsenceEndSession & " as 'EndSession', " & vbCrLf & _
    mstrSQLSelect_AbsenceType & " as 'Type', " & vbCrLf & _
    mstrSQLSelect_AbsenceTypeCalCode & " as 'CalendarCode', " & vbCrLf & _
    mstrSQLSelect_AbsenceTypeCode & " as 'Code', " & vbCrLf & _
    mstrSQLSelect_AbsenceReason & " as 'Reason', " & vbCrLf & _
    mstrSQLSelect_AbsenceDuration & " as 'Duration' " & vbCrLf
  
  sSQL = sSQL & "FROM " & mstrAbsenceTableRealSource & vbCrLf
  sSQL = sSQL & "           INNER JOIN " & gsAbsenceTypeTableName & vbCrLf
  sSQL = sSQL & "           ON " & mstrAbsenceTableRealSource & "." & gsAbsenceTypeColumnName & " = " & gsAbsenceTypeTableName & "." & gsAbsenceTypeTypeColumnName & vbCrLf
  
  sSQL = sSQL & "WHERE " & mstrAbsenceTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelID & vbCrLf
  sSQL = sSQL & " AND (" & mstrSQLSelect_AbsenceStartDate & " IS NOT NULL) " & vbCrLf
  sSQL = sSQL & "ORDER BY 'StartDate' ASC" & vbCrLf

  Set mrstAbsenceRecords = datGeneral.GetRecords(sSQL)
  GetAbsenceRecordSet = True
  Exit Function
  
GetAbsenceRecordSet_ERROR:
  
  COAMsgBox "Error retrieving the Absence recordset." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Set mrstAbsenceRecords = Nothing
  GetAbsenceRecordSet = False

End Function

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

End Sub

Private Sub cboStartMonth_Click()

  If mblnLoading Then Exit Sub
  
  Dim sDateFormat
  
  sDateFormat = LCase(DateFormat)
  
  sDateFormat = Replace(sDateFormat, "dd", "01")
  sDateFormat = Replace(sDateFormat, "mm", Me.cboStartMonth.ListIndex + 1)
  
  If InStr(sDateFormat, "yyyy") Then
    sDateFormat = Replace(sDateFormat, "yyyy", Me.lblCurrentYear.Caption)
  Else
    sDateFormat = Replace(sDateFormat, "yy", Me.lblCurrentYear.Caption)
  End If
  
  gdtmStartMonth = sDateFormat
  
  AbsCal_GetFirstAndLastViewedDates
  DrawMonths
  RefreshCal
  
  lblDash.Visible = (Month(gdtmStartMonth) <> 1)
  lblCurrentYear2.Visible = lblDash.Visible
  
  If Month(gdtmStartMonth) <> 1 Then
    lblCurrentYear2.Caption = lblCurrentYear.Caption + 1
    lblCurrentYear.Left = 9700
  Else
    lblCurrentYear.Left = 9970
  End If
  
End Sub

Private Sub chkCaptions_Click()

  RefreshLegend ((chkCaptions.Value = vbChecked))
  
  If Not mblnLoading Then
    Screen.MousePointer = vbHourglass
    RefreshCal
    Screen.MousePointer = vbDefault
  End If

End Sub

Private Sub chkShadeWeekends_Click()

  If Not mblnLoading Then
    Screen.MousePointer = vbHourglass
    RefreshCal
    Screen.MousePointer = vbDefault
  End If
  
End Sub

Private Sub chkShadeBHols_Click()

  If Not mblnLoading Then
    Screen.MousePointer = vbHourglass
    RefreshCal
    Screen.MousePointer = vbDefault
  End If
  
End Sub

Private Sub chkIncludeWorkingDaysOnly_Click()

  If Not mblnLoading Then
    Screen.MousePointer = vbHourglass
    RefreshCal
    Screen.MousePointer = vbDefault
  End If
  
End Sub

Private Sub chkIncludeBHols_Click()

  If Not mblnLoading Then
    Screen.MousePointer = vbHourglass
    RefreshCal
    Screen.MousePointer = vbDefault
  End If
  
End Sub

Private Sub RefreshCal()

  Screen.MousePointer = vbHourglass
  
  'Show the current year
  GetYearLayout

  ' Shade the cells that arent used
  DisableUnused

  'Complete the grid according to options set by the user
  FillGridWithData

  Screen.MousePointer = vbDefault
  
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdYearAdd_Click()

  Screen.MousePointer = vbHourglass
    
  ' Add one year to the current year label and refresh the data
  lblCurrentYear.Caption = lblCurrentYear.Caption + 1
  lblCurrentYear2.Caption = lblCurrentYear.Caption + 1
      
  ' Get the first and last displayed dates (essential)
  AbsCal_GetFirstAndLastViewedDates

  ' Do not allow user to view before start date or after leaving date
  If IsDate(lblStartDate.Caption) Then
    If CDate(lblFirstDate.Caption) < CDate(lblStartDate.Caption) Then
      cmdYearSubtract.Enabled = False
    Else
      cmdYearSubtract.Enabled = True
    End If
  End If
  
  If IsDate(lblEndDate.Caption) Then
    If CDate(lblEndDate.Caption) <= CDate(lblLastDate.Caption) Then
      cmdYearAdd.Enabled = False
    Else
      cmdYearAdd.Enabled = True
    End If
  End If

  RefreshCal
  
  Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdYearSubtract_Click()

  Screen.MousePointer = vbHourglass
    
  ' Subtract one year from the current year label and refresh the data
  lblCurrentYear.Caption = lblCurrentYear.Caption - 1
  lblCurrentYear2.Caption = lblCurrentYear.Caption + 1
  
  ' Get the first and last displayed dates (essential)
  AbsCal_GetFirstAndLastViewedDates
  
  ' Do not allow user to view before start date or after leaving date
  If IsDate(lblStartDate.Caption) Then
    If CDate(lblFirstDate.Caption) <= CDate(lblStartDate.Caption) Then
      cmdYearSubtract.Enabled = False
    Else
      cmdYearSubtract.Enabled = True
    End If
  End If
  If IsDate(lblEndDate.Caption) Then
    If CDate(lblEndDate.Caption) <= CDate(lblLastDate.Caption) Then
      cmdYearAdd.Enabled = False
    Else
      cmdYearAdd.Enabled = True
    End If
  End If
  
  RefreshCal
  
  Screen.MousePointer = vbDefault
    
End Sub

Public Sub Initialise()

  Screen.MousePointer = vbHourglass
  
  ' Set the loading flag
  mblnLoading = True
  
  ' Store the current personnel ID
  mlngPersonnelID = frmMain.ActiveForm.Recordset.Fields("ID")
  
  ' Load the setup information from the Absence table and the Personnel table
  If Not GetInfoFromDB Then Exit Sub
    
  If mrstAbsenceRecords.BOF And mrstAbsenceRecords.EOF Then
    If COAMsgBox("The current employee does not have any absence records." & vbCrLf & vbCrLf & _
              "Are you sure you want to view the absence calendar ?", vbYesNo + vbQuestion, App.Title) = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
    End If
  End If
  
  ' Setup the form
  If Not DrawMonths Then Exit Sub
  If Not DrawDays Then Exit Sub
  If Not DrawCalDates Then Exit Sub
  If Not DrawCal Then Exit Sub
  If Not DrawLines Then Exit Sub

  ' Load the colour key grid
  If LoadColourKey = False Then
    Unload frmAbsenceCalendar
    Exit Sub
  End If

  If IsDate(lblStartDate.Caption) Then
    If CDate(lblFirstDate.Caption) < CDate(lblStartDate.Caption) Then
      cmdYearSubtract.Enabled = False
    Else
      cmdYearSubtract.Enabled = True
    End If
  End If
  If IsDate(lblEndDate.Caption) Then
    If CDate(lblEndDate.Caption) < CDate(lblLastDate.Caption) Then
      cmdYearAdd.Enabled = False
    Else
      cmdYearAdd.Enabled = True
    End If
  End If
  
  RefreshCal
  
  ' Unset the loading flag
  mblnLoading = False
    
  Screen.MousePointer = vbDefault
  
  ' Show the form
  Me.Show vbModal
  
End Sub


Private Sub FillGridWithData()
  
  On Error Resume Next
  
  Dim counter As Integer, intStart As Integer, intEnd As Integer, booOK As Boolean
  Dim sSQL As String

  booOK = True
  
  ' If there are no absence records for the current employee then skip
  ' this bit (but still show the form)
  If mrstAbsenceRecords.BOF And mrstAbsenceRecords.EOF Then
    Exit Sub
  End If
    
  mrstAbsenceRecords.MoveFirst
  ' Loop through the absence recordset
  Do Until mrstAbsenceRecords.EOF
    
    ' Load each absence record data into variables
    ' (has to be done because start/end dates may be modified by code to fill grid correctly)
    
    dtmAbsStartDate = Format(mrstAbsenceRecords.Fields("StartDate"), DateFormat)
    'dtmAbsStartDate = IIf(IsNull(mrstAbsenceRecords.Fields("StartDate")), Format(mrstAbsenceRecords.Fields("StartDate"), DateFormat), Format(mrstAbsenceRecords.Fields("EndDate"), DateFormat))
    
    strAbsStartSession = UCase(mrstAbsenceRecords.Fields("StartSession"))
    
    dtmAbsEndDate = IIf(IsNull(mrstAbsenceRecords.Fields("EndDate")), Format(Now, DateFormat), Format(mrstAbsenceRecords.Fields("EndDate"), DateFormat))
    
    strAbsEndSession = UCase(mrstAbsenceRecords.Fields("EndSession"))
    strAbsType = Replace(mrstAbsenceRecords.Fields("Type"), "&", "&&")
    strAbsCalendarCode = IIf(IsNull(mrstAbsenceRecords.Fields("CalendarCode")), "", mrstAbsenceRecords.Fields("CalendarCode"))
    strAbsCode = mrstAbsenceRecords.Fields("Code")
      
    ' If the start date is after the end date, ignore the record
    If (dtmAbsStartDate > dtmAbsEndDate) Then
    
    ' if the absence record is totally before the currently viewed timespan then do nothing
    ElseIf (dtmAbsStartDate < lblFirstDate.Caption) And (dtmAbsEndDate < lblFirstDate.Caption) Then
    
    ' if the absence record is totally after the currently viewed timespan then do nothing
    ElseIf (dtmAbsStartDate > lblLastDate.Caption) And (dtmAbsEndDate > lblLastDate.Caption) Then
    
    ' if the absence record starts before currently viewed timespan, but ends in the timspan then
    ElseIf (dtmAbsStartDate < lblFirstDate.Caption) And (dtmAbsEndDate < lblLastDate.Caption) Then
      dtmAbsStartDate = lblFirstDate.Caption
      strAbsStartSession = "AM"
      If Month(dtmAbsStartDate) = Month(dtmAbsEndDate) Then
        booOK = FillSameMonths
      Else
        booOK = FillDifferentMonths
      End If
    
    ' if the absence record starts in the currently viweed timespan, but ends after it then
    ElseIf (dtmAbsStartDate >= lblFirstDate.Caption) And (dtmAbsEndDate > lblLastDate.Caption) Then
      dtmAbsEndDate = lblLastDate.Caption
      strAbsEndSession = "PM"
      If Month(dtmAbsStartDate) = Month(dtmAbsEndDate) Then
        booOK = FillSameMonths
      Else
        booOK = FillDifferentMonths
      End If
      
    ' if the absence record is enclosed within viewed timespan, and months are equal then
    ElseIf (dtmAbsStartDate >= lblFirstDate.Caption) And (dtmAbsEndDate <= lblLastDate.Caption) And (Month(dtmAbsStartDate) = Month(dtmAbsEndDate)) Then
      booOK = FillSameMonths
    
    ' if the absence record is enclosed within viewed timespan, and months are different then
    ElseIf (dtmAbsStartDate >= lblFirstDate.Caption) And (dtmAbsEndDate <= lblLastDate.Caption) And (Month(dtmAbsStartDate) <> Month(dtmAbsEndDate)) Then
      booOK = FillDifferentMonths
      
    ElseIf (dtmAbsStartDate < CDate(lblFirstDate.Caption)) And (dtmAbsEndDate > CDate(lblLastDate.Caption)) Then
      dtmAbsStartDate = lblFirstDate.Caption
      dtmAbsEndDate = lblLastDate.Caption
      strAbsStartSession = "AM"
      strAbsEndSession = "PM"
      
      booOK = FillDifferentMonths
    
    End If
    
    If booOK = False Then Exit Do
  
    mrstAbsenceRecords.MoveNext
  
  Loop

  If booOK = False Then
    COAMsgBox "An Error Has Occurred Whilst Filling The Cal Labels:" & vbCrLf & Err.Number & " - " & Err.Description
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
  
  Dim iCount As Integer
  
  ReDim mvarTableViews(3, 0)
  ReDim mstrViews(0)
  
  For iCount = 1 To 12
    cboStartMonth.AddItem StrConv(MonthName(iCount), vbProperCase)
'    cboStartMonth.ItemData(cboStartMonth.NewIndex) = iCount
  Next iCount

 ' lblDay(1).Caption = StrConv(Left(WeekdayName(1), 1), vbProperCase)
  'JPD 20030828 Fault 2012
  'JPD 20030828 Fault 1630
  'lblDay(1).Caption = StrConv(Left(WeekdayName(Weekday("01/01/2001") - 1, , vbSunday), 1), vbProperCase)
  lblDay(1).Caption = StrConv(Left(WeekdayName(Weekday(ConvertSQLDateToLocale("01/01/2001")) - 1, , vbSunday), 1), vbProperCase)
  
  'Set caption of form to show the name of the person whose absence
  'we are viewing
  Me.Caption = Me.Caption & " - " & Right(frmMain.ActiveForm.Caption, (Len(frmMain.ActiveForm.Caption)) - (InStr(frmMain.ActiveForm.Caption, "-")) - 1)

  'Sort out the year labelleling depending on whether Jan-Dec is displayed or not
  If Month(gdtmStartMonth) <> 1 Then
    lblDash.Visible = True
    lblCurrentYear2.Visible = True
    lblCurrentYear2.Caption = lblCurrentYear.Caption + 1
    lblCurrentYear.Left = 9700
  End If

  cboStartMonth.ListIndex = (Month(gdtmStartMonth) - 1)
  If Month(gdtmStartMonth) > Month(Now) Then
    frmAbsenceCalendar.lblCurrentYear.Caption = Year(Now) - 1
  Else
    frmAbsenceCalendar.lblCurrentYear.Caption = Year(Now)
  End If

  frmAbsenceCalendar.lblCurrentYear2.Caption = frmAbsenceCalendar.lblCurrentYear.Caption + 1

  'Obtain the currently viewed timespan of the calendar
  AbsCal_GetFirstAndLastViewedDates

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Release references and unload the form
  Set mrstAbsenceRecords = Nothing
End Sub

Private Sub lblCal_Click(Index As Integer)
  
  Dim dtmDateToFind As Date, strSession As String

  ' If there is nothing there, or its a weekend, then dont bother checking the recordset !
  If (lblCal(Index).BackColor = lblCal(9999).BackColor) Or _
     (lblCal(Index).BackColor = 13405581) Then Exit Sub

  dtmDateToFind = GetCalDay(Index)

  If Index Mod 2 = 0 Then
    strSession = "PM"
  Else
    strSession = "AM"
  End If
  
  If dtmDateToFind = ConvertSQLDateToLocale("12/12/9999") Then Exit Sub
  
  mrstAbsenceRecords.MoveFirst
  
  If mrstAbsenceRecords.BOF And mrstAbsenceRecords.EOF Then
    COAMsgBox "Sorry, there are no absences for this employee", vbExclamation + vbOKCancel, "Absence Calendar"
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  If frmAbsenceCalendarDetails.Initialise(mrstAbsenceRecords, dtmDateToFind, strSession, (Not mblnDisableRegions), (Not mblnDisableWPs)) Then
    Screen.MousePointer = vbDefault
    frmAbsenceCalendarDetails.Show vbModal
  End If

  Screen.MousePointer = vbDefault
    
End Sub

Private Function LoadColourKey() As Boolean

  On Error GoTo errLoadColourKey
  
  Dim rstColourKey As Recordset, strColourKeySQL As String, intCounter As Integer
  
  strColourKeySQL = "SELECT DISTINCT " & gsAbsenceTypeTypeColumnName & ", " & gsAbsenceTypeCalCodeColumnName & " FROM " & gsAbsenceTypeTableName & " ORDER BY " & gsAbsenceTypeTypeColumnName
  
  Set rstColourKey = datGeneral.GetRecords(strColourKeySQL)

  If rstColourKey.BOF And rstColourKey.EOF Then
    COAMsgBox "You have no absence types defined in your Absence Type table", vbExclamation + vbOKOnly, "Absence Calendar"
    LoadColourKey = False
    Exit Function
  End If
  
  Do Until rstColourKey.EOF
    
    intCounter = intCounter + 1
    
    If intCounter <= 18 Then
      
      ' Load another instance of the 2 controls
      Load lblColourKey_Colour(intCounter)
      Load lblColourKey_Type(intCounter)
      
      ' Position the colour box control depending on its index
      If intCounter <= 10 Then
        lblColourKey_Colour(intCounter).Top = 250 * (intCounter)
      Else
        lblColourKey_Colour(intCounter).Top = 250 * (intCounter - 10)
      End If
      
      ' Position the colour box control depending on its index
      If intCounter <= 10 Then
        lblColourKey_Colour(intCounter).Left = 150
      Else
        lblColourKey_Colour(intCounter).Left = 1250
      End If
      
      ' Set the colour box caption and show the label
      If IsNull(rstColourKey.Fields(1).Value) Or (chkCaptions.Value = vbUnchecked) Then
        lblColourKey_Colour(intCounter).Caption = ""
      Else
        lblColourKey_Colour(intCounter).Caption = Replace(rstColourKey.Fields(1).Value, "&", "&&")
      End If
      
      'AE20071108 Fault #12547
      If Len(lblColourKey_Colour(intCounter).Caption) > 1 Then
        lblColourKey_Colour(intCounter).FontSize = KEY_FONTSIZE_SMALL
      Else
        lblColourKey_Colour(intCounter).FontSize = KEY_FONTSIZE_NORMAL
      End If
      
      lblColourKey_Colour(intCounter).Visible = True
      
      ' Set the colour box background colour depending on the index
      Select Case intCounter
      
         Case 1, 6, 11, 16: lblColourKey_Colour(intCounter).BackColor = &HC0C0FF
         Case 2, 7, 12, 17: lblColourKey_Colour(intCounter).BackColor = &HC0FFC0
         Case 3, 8, 13, 18: lblColourKey_Colour(intCounter).BackColor = &HC0FFFF
         Case 4, 9, 14: lblColourKey_Colour(intCounter).BackColor = &HC0E0FF
         Case 5, 10, 15: lblColourKey_Colour(intCounter).BackColor = &HFFFFC0
         
      End Select
    
      ' Position the colour key control depending on its index
      If intCounter <= 10 Then
        lblColourKey_Type(intCounter).Top = 250 * (intCounter)
      Else
        lblColourKey_Type(intCounter).Top = 250 * (intCounter - 10)
      End If
      
      ' Position the colour key control depending on its index
      If intCounter <= 10 Then
        lblColourKey_Type(intCounter).Left = 400
      Else
        lblColourKey_Type(intCounter).Left = 1500
      End If
    
      ' Set the colour box caption and show the control
      lblColourKey_Type(intCounter).Caption = IIf(IsNull(rstColourKey.Fields(0)), "", Replace(rstColourKey.Fields(0), "&", "&&"))
      lblColourKey_Type(intCounter).Visible = True
    
    End If
    
    rstColourKey.MoveNext
    
  Loop
  
  ' Now add the 'Other' box (if needed)
  If intCounter > 18 Then
    intCounter = 19
    Load lblColourKey_Colour(intCounter)
    Load lblColourKey_Type(intCounter)
    'Position the colour box
    lblColourKey_Colour(intCounter).Top = 250 * (intCounter - 10)
    'Position the colour box
    lblColourKey_Colour(intCounter).Left = 1250
    ' Set Caption, Colour and Show it
    lblColourKey_Colour(intCounter).BackColor = vbBlack
    lblColourKey_Colour(intCounter).Caption = "-"
    lblColourKey_Colour(intCounter).Visible = True
    'Position the colour key box
    lblColourKey_Type(intCounter).Top = 250 * (intCounter - 10)
    'Position the colour key box
    lblColourKey_Type(intCounter).Left = 1500
    ' Set the colour key box caption and show the control
    lblColourKey_Type(intCounter).Caption = "Other"
    lblColourKey_Type(intCounter).Visible = True
  End If
  
  ' Now add the multiple box
  intCounter = intCounter + 1
  Load lblColourKey_Colour(intCounter)
  Load lblColourKey_Type(intCounter)
  ' Position the colour box control depending on its index
  If intCounter < 11 Then
    lblColourKey_Colour(intCounter).Top = 250 * (intCounter)
  Else
    lblColourKey_Colour(intCounter).Top = 250 * (intCounter - 10)
  End If
  
  ' Position the colour box control depending on its index
  If intCounter < 11 Then
    lblColourKey_Colour(intCounter).Left = 150
  Else
    lblColourKey_Colour(intCounter).Left = 1250
  End If
  
  mintMultipleKeyControlIndex = intCounter
  lblColourKey_Colour(intCounter).Visible = True
  lblColourKey_Colour(intCounter).Caption = IIf(mblnShowBankHols, ".", "")
  lblColourKey_Colour(intCounter).BackColor = vbWhite
  
  ' Position the colour key control depending on its index
  If intCounter < 11 Then
    lblColourKey_Type(intCounter).Top = 250 * (intCounter)
  Else
    lblColourKey_Type(intCounter).Top = 250 * (intCounter - 10)
  End If
  
  ' Position the colour key control depending on its index
  If intCounter <= 10 Then
    lblColourKey_Type(intCounter).Left = 400
  Else
    lblColourKey_Type(intCounter).Left = 1500
  End If
  
  lblColourKey_Type(intCounter).Visible = True
  lblColourKey_Type(intCounter).Caption = "Multiple"
  'NHRD08092006 Fault 11461 Dot wasn't being inserted on load up
  lblColourKey_Colour(mintMultipleKeyControlIndex).Caption = "."
  
  ' If we are here, then notify calling procedure of success and exit
  LoadColourKey = True
  Exit Function
  
errLoadColourKey:

  COAMsgBox "An error has occurred - LoadColourKey." & vbCrLf & "Please check your absence module setup.", vbCritical + vbOKOnly, "Absence Calendar"
  LoadColourKey = False
  
End Function

Private Sub cmdPrint_Click()
    
  Dim aspect As Long
  Dim wID As Long
  Dim hgt As Long
  Dim xmin As Long
  Dim ymin As Long
  Dim oPrinter As clsPrintDef

  On Error GoTo Print_ERR
  
  Screen.MousePointer = vbHourglass
  DoEvents
  
  Set oPrinter = New clsPrintDef
  
  'NHRD16072004 Fault 8739
  If oPrinter.IsOK = False Then GoTo TidyUpAndExit
  
  If UI.GetOSName = "Windows NT" Then
  
    ' Win NT
    
      ' Press Alt.
      keybd_event VK_MENU, 0, 0, 0
      DoEvents
  
      cmdYearAdd.Visible = False
      cmdYearSubtract.Visible = False
      cmdPrint.Visible = False
      DoEvents
  
      ' Press Print Scrn.
      keybd_event VK_SNAPSHOT, 1, 0, 0
  
      DoEvents
  
      cmdYearAdd.Visible = True
      cmdYearSubtract.Visible = True
      cmdPrint.Visible = True
  
      ' Release Alt.
      keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
      DoEvents
    
  Else
  
    ' Win 95/8
    
    cmdYearAdd.Visible = False
    cmdYearSubtract.Visible = False
    cmdPrint.Visible = False
    DoEvents
  
    ' Press Print Scrn.
    keybd_event vbKeySnapshot, 0, 0, 0
  
    DoEvents
  
    cmdYearAdd.Visible = True
    cmdYearSubtract.Visible = True
    cmdPrint.Visible = True
    
  End If
 
  ' Display the printing options (if switch is on)
  If oPrinter.PrintStart_AbsenceCalendar Then
  
    ' Copy the image into the hidden PictureBox.
    HiddenPict.Picture = Clipboard.GetData(vbCFBitmap)
  
    ' Print the image.
    wID = Printer.ScaleX(HiddenPict.ScaleWidth, ScaleMode, Printer.ScaleMode)
    hgt = Printer.ScaleY(HiddenPict.ScaleHeight, ScaleMode, Printer.ScaleMode)
    xmin = (Printer.ScaleWidth - wID) / 2
    ymin = (Printer.ScaleHeight - hgt) / 2
  
    ' Make the image as large as possible
    Printer.PaintPicture HiddenPict.Picture, xmin, ymin, wID, hgt
    Printer.EndDoc
    
    ' Display a printing complete prompt
    oPrinter.PrintConfirm "Absence Calendar", App.ProductName
    
    Dim objDefPrinter As cSetDfltPrinter
    Set objDefPrinter = New cSetDfltPrinter
    Do
      objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
    Loop While Printer.DeviceName <> gstrDefaultPrinterName
    Set objDefPrinter = Nothing

  End If
TidyUpAndExit:
  Screen.MousePointer = vbDefault
  DoEvents
    
  Exit Sub
  
Print_ERR:
  
  Screen.MousePointer = vbDefault

  Select Case Err.Number
    Case 482: COAMsgBox "Printer Error : Please check printer connection.", vbExclamation + vbInformation, "HR Pro"
    Case Else: COAMsgBox Err.Description, vbExclamation + vbInformation, "HR Pro"
  End Select
  
End Sub


Private Function LoadDefaultDisplayOptions() As Boolean

  On Error GoTo Load_ERROR
  
'TM20060619 - Fault 10831
'  ' Load the month that is at the top of the calendar
'  gdtmStartMonth = "01/" & Str(giAbsenceCalStartMonth) & "/" & Str(Year(Now))
  
  Dim sDateFormat
  
  sDateFormat = LCase(DateFormat)
  
  sDateFormat = Replace(sDateFormat, "dd", "01")
  sDateFormat = Replace(sDateFormat, "mm", giAbsenceCalStartMonth)
  
  If InStr(sDateFormat, "yyyy") Then
    sDateFormat = Replace(sDateFormat, "yyyy", Year(Now))
  Else
    sDateFormat = Replace(sDateFormat, "yy", Year(Now))
  End If
  
  gdtmStartMonth = sDateFormat
  
  
  ' Should the Weekends be shaded ?
  chkShadeWeekends.Value = IIf(gfAbsenceCalWeekendShading, vbChecked, vbUnchecked)
  ' Should the BHols be shaded ?
  chkShadeBHols.Value = IIf(gfAbsenceCalBHolShading, vbChecked, vbUnchecked)
  ' Should the Weekends be included ?
  chkIncludeWorkingDaysOnly.Value = IIf(gfAbsenceCalIncludeWorkingDaysOnly, vbChecked, vbUnchecked)
  ' Should the BHols be included ?
  chkIncludeBHols.Value = IIf(gfAbsenceCalBHolInclude, vbChecked, vbUnchecked)
  ' Should we Show Captions ?
  chkCaptions.Value = IIf(gfAbsenceCalShowCaptions, vbChecked, vbUnchecked)
  ' Set the current viewed year
  'frmAbsenceCalendar.lblCurrentYear.Caption = Year(Now)

  LoadDefaultDisplayOptions = True
  Exit Function
  
Load_ERROR:
  
  COAMsgBox "Error whilst loading default display options." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Absence Calendar"
  LoadDefaultDisplayOptions = False
  
End Function

Public Property Get IncludeBankHolidays() As Boolean
  IncludeBankHolidays = (chkIncludeBHols.Value = vbChecked)
End Property
Public Property Get WorkingDaysOnly() As Boolean
  WorkingDaysOnly = (chkIncludeWorkingDaysOnly.Value = vbChecked)
End Property
Public Property Get ShowBankHolidays() As Boolean
  ShowBankHolidays = (chkShadeBHols.Value = vbChecked)
End Property
Public Property Get ShowCaptions() As Boolean
  ShowCaptions = (chkCaptions.Value = vbChecked)
End Property
Public Property Get ShowWeekends() As Boolean
  ShowWeekends = (chkShadeWeekends.Value = vbChecked)
End Property

Public Property Let ShowBankHolidays(pblnShowBankHols As Boolean)
  If (Not mblnDisableRegions) Then
    chkShadeBHols.Enabled = True
    If pblnShowBankHols Then
      chkShadeBHols.Value = vbChecked
    Else
      chkShadeBHols.Value = vbUnchecked
    End If
  Else
    chkShadeBHols.Value = vbUnchecked
    chkShadeBHols.Enabled = False
  End If
End Property

Public Property Let WorkingDaysOnly(pblnIncludeWorkingDaysOnly As Boolean)
  If (Not mblnDisableWPs) Then
    chkIncludeWorkingDaysOnly.Enabled = True
    If pblnIncludeWorkingDaysOnly Then
      chkIncludeWorkingDaysOnly.Value = vbChecked
    Else
      chkIncludeWorkingDaysOnly.Value = vbUnchecked
    End If
  Else
    chkIncludeWorkingDaysOnly.Value = vbUnchecked
    chkIncludeWorkingDaysOnly.Enabled = False
  End If
End Property

Public Property Let IncludeBankHolidays(pblnIncludeBankHolidays As Boolean)
  If (Not mblnDisableRegions) Then
    chkIncludeBHols.Enabled = True
    If pblnIncludeBankHolidays Then
      chkIncludeBHols.Value = vbChecked
    Else
      chkIncludeBHols.Value = vbUnchecked
    End If
  Else
    chkIncludeBHols.Value = vbUnchecked
    chkIncludeBHols.Enabled = False
  End If
End Property

Public Property Let ShowCaptions(pblnShowCaptions As Boolean)
  chkCaptions.Value = IIf(pblnShowCaptions, vbChecked, vbUnchecked)
End Property

Public Property Let ShowWeekends(pblnShowWeekends As Boolean)
  chkShadeWeekends.Value = IIf(pblnShowWeekends, vbChecked, vbUnchecked)
End Property

Private Function CheckPermission_Columns(plngTableID As Long, pstrTableName As String, _
                                        pstrColumnName As String, strSQLRef As String) As Boolean

  'This function checks if the current user has read(select) permissions
  'on this column. If the user only has access through views then the
  'relevent views are added to the mvarTableViews() array which in turn
  'are used to create the join part of the query.

  Dim lngTempTableID As Long
  Dim strTempTableName As String
  Dim strTempColumnName As String
  Dim blnColumnOK As Boolean
  Dim blnFound As Boolean
  Dim blnNoSelect As Boolean
  Dim iLoop1 As Integer
  Dim intLoop As Integer
  Dim strColumnCode As String
  Dim strSource As String
  Dim intNextIndex As Integer
  Dim blnOK As Boolean
  Dim strTable As String
  Dim strColumn As String
  
  Dim pintNextIndex  As Integer
  
  ' Set flags with their starting values
  blnOK = True
  blnNoSelect = False

  strTable = vbNullString
  strColumn = vbNullString
 
  ' Load the temp variables
  lngTempTableID = plngTableID
  strTempTableName = pstrTableName
  strTempColumnName = pstrColumnName

  ' Check permission on that column
  Set mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
  mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

  blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

  If blnColumnOK Then
    blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
  End If

  If blnColumnOK Then
    ' this column can be read direct from the tbl/view or from a parent table
    strTable = mstrRealSource
    strColumn = strTempColumnName
    
    If (plngTableID = glngAbsenceTableID) And (mstrAbsenceTableRealSource = vbNullString) Then
      mstrAbsenceTableRealSource = strTable
    End If
    
'    ' If the table isnt the base table (or its realsource) then
'    ' Check if it has already been added to the array. If not, add it.
'    If lngTempTableID <> mlngCalendarReportsBaseTable Then
      blnFound = False
      For intNextIndex = 1 To UBound(mvarTableViews, 2)
        If mvarTableViews(1, intNextIndex) = 0 And _
        mvarTableViews(2, intNextIndex) = lngTempTableID Then
        blnFound = True
          Exit For
        End If
      Next intNextIndex

      If Not blnFound Then
        intNextIndex = UBound(mvarTableViews, 2) + 1
        ReDim Preserve mvarTableViews(3, intNextIndex)
        mvarTableViews(1, intNextIndex) = 0
        mvarTableViews(2, intNextIndex) = lngTempTableID
      End If
'    End If
  
    strSQLRef = strTable & "." & strColumn
  Else

    ' this column cannot be read direct. If its from a parent, try parent views
    ' Loop thru the views on the table, seeing if any have read permis for the column

    ReDim mstrViews(0)
    For Each mobjTableView In gcoTablePrivileges.Collection
      If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = lngTempTableID) And _
          (mobjTableView.AllowSelect) Then

        strSource = mobjTableView.ViewName
        mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

        ' Get the column permission for the view
        Set mobjColumnPrivileges = GetColumnPrivileges(strSource)

        ' If we can see the column from this view
        If mobjColumnPrivileges.IsValid(strTempColumnName) Then
          If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then
            
            ReDim Preserve mstrViews(UBound(mstrViews) + 1)
            mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

            ' Check if view has already been added to the array
            blnFound = False
            For intNextIndex = 0 To UBound(mvarTableViews, 2)
              If mvarTableViews(1, intNextIndex) = 1 And _
              mvarTableViews(2, intNextIndex) = mobjTableView.ViewID Then
                blnFound = True
                Exit For
              End If
            Next intNextIndex

            If Not blnFound Then
              ' View hasnt yet been added, so add it !
              intNextIndex = UBound(mvarTableViews, 2) + 1
              ReDim Preserve mvarTableViews(3, intNextIndex)
              mvarTableViews(0, intNextIndex) = mobjTableView.TableID
              mvarTableViews(1, intNextIndex) = 1
              mvarTableViews(2, intNextIndex) = mobjTableView.ViewID
              mvarTableViews(3, intNextIndex) = mobjTableView.ViewName
            End If
            
          End If
        End If
      End If

    Next mobjTableView
    Set mobjTableView = Nothing

    ' Does the user have select permission thru ANY views ?
    If UBound(mstrViews()) = 0 Then
      blnNoSelect = True
    Else
      strSQLRef = ""
      For pintNextIndex = 1 To UBound(mstrViews)
        If pintNextIndex = 1 Then
          strSQLRef = "CASE"
        End If
        
        strSQLRef = strSQLRef & _
        " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & strTempColumnName
      Next pintNextIndex
      
      If Len(strSQLRef) > 0 Then
        strSQLRef = strSQLRef & _
        " ELSE NULL" & _
        " END "
      End If

    End If

    ' If we cant see a column, then get outta here
    If blnNoSelect Then
      strSQLRef = vbNullString
      CheckPermission_Columns = False
      Exit Function
    End If

    If Not blnOK Then
      strSQLRef = vbNullString
      CheckPermission_Columns = False
      Exit Function
    End If

  End If

  CheckPermission_Columns = True
  
End Function

Private Function CheckPermission_RegionInfo() As Boolean

  Dim strTableColumn As String
  
  'Check the  Bank Holiday Region Table - Region Table
  '           Bank Holiday Region Table - Region Column
  '           Bank Holidays Table - Bank Holiday Table
  '           Bank Holidays Table - Date Column
  '           Bank Holidays Table - Descripiton Column
  '...Bank Holiday module setup information.
  'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
  If gsBHolRegionTableName = "" Or _
     gsBHolRegionColumnName = "" Or _
     gsBHolTableName = "" Or _
     gsBHolDateColumnName = "" Or _
     gsBHolDescriptionColumnName = "" Then
     
    GoTo DisableRegions
  End If
   
  'Check the  Career Change Region - Static Region Column
  '           Career Change Region - Historic Region Table
  '           Career Change Region - Historic Region Column
  '           Career Change Region - Historic Region Effective Date Column
  '...Personnel - Career Change module setup information.
  'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
  If gsPersonnelRegionColumnName = "" Then
    If gsPersonnelHRegionTableName = "" Or _
       gsPersonnelHRegionColumnName = "" Or _
       gsPersonnelHRegionDateColumnName = "" Then
       
      GoTo DisableRegions
    End If
  End If




  '*******************************************************************
  ' All Region module information is setup                           *
  ' Now check the permissions on the Region module setup information *
  '*******************************************************************
  'Bank Holiday Region Table - Region Table (Regional Information)
  'Bank Holiday Region Table - Region Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngBHolRegionTableID, gsBHolRegionTableName, _
                            gsBHolRegionColumnName, strTableColumn) Then
    mstrSQLSelect_RegInfoRegion = strTableColumn
    strTableColumn = vbNullString
  Else
    GoTo DisableRegions
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
 
  'Bank Holidays Table - Bank Holiday Table (Region History)
  'Bank Holidays Table - Date Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, _
                            gsBHolDateColumnName, strTableColumn) Then
    mstrSQLSelect_BankHolDate = strTableColumn
    strTableColumn = vbNullString
  Else
    GoTo DisableRegions
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  
  'Bank Holidays Table - Bank Holiday Table (Region History)
  'Bank Holidays Table - Descripiton Column
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, _
                            gsBHolDescriptionColumnName, strTableColumn) Then
    mstrSQLSelect_BankHolDesc = strTableColumn
    strTableColumn = vbNullString
  Else
    GoTo DisableRegions
  End If
  '///////////////////////////////////////////////
  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\




  '*******************************************************************
  ' Permission granted on all Region module information.             *
  ' Now check the permissions on the                                 *
  ' Personnel Career Change Region module setup information          *
  '*******************************************************************
  'Check Career Change Region access
  If gsPersonnelRegionColumnName <> "" Then
    'Personnel Table
    'Career Change Region - Static Region Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, _
                              gsPersonnelRegionColumnName, strTableColumn) Then
      mstrSQLSelect_PersonnelStaticRegion = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
  Else
    'Career Change Region - Historic Region Table
    'Career Change Region - Historic Region Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, _
                              gsPersonnelHRegionColumnName, strTableColumn) Then
      mstrSQLSelect_PersonnelHRegion = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    'Career Change Region - Historic Region Table
    'Career Change Region - Historic Region Effective Date Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, _
                              gsPersonnelHRegionDateColumnName, strTableColumn) Then
      mstrSQLSelect_PersonnelHDate = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

  End If
  
  CheckPermission_RegionInfo = True

TidyUpAndExit:
  Exit Function

DisableRegions:
  mblnDisableRegions = True
  ShowBankHolidays = False
  IncludeBankHolidays = False
  mblnShowBankHols = False
  mblnRegions = False
  CheckPermission_RegionInfo = False
  GoTo TidyUpAndExit
  
End Function

Private Function CheckPermission_WPInfo() As Boolean
 
  Dim objTable As CTablePrivilege
  Dim objColumn As CColumnPrivileges
  Dim pblnColumnOK As Boolean
  Dim strTableColumn As String
  
  'Check the  Career Change Working Pattern - Static Working Pattern Column
  '           Career Change Working Pattern - Historic Working Pattern Table
  '           Career Change Working Pattern - Historic Working Pattern Column
  '           Career Change Working Pattern - Historic Working Pattern Effective Date Column
  '...Personnel - Career Change module setup information.
  'If any are blank then we need to allow the report to run, but disable the Working Dys Display Option.
  If gsPersonnelWorkingPatternColumnName = "" Then
    If gsPersonnelHWorkingPatternTableName = "" Or _
       gsPersonnelHWorkingPatternColumnName = "" Or _
       gsPersonnelHWorkingPatternDateColumnName = "" Then
       
      GoTo DisableWPs
    End If
  End If
   
  '****************************************************************************
  ' All Working Pattern module information is setup                           *
  ' Now check the permissions on the Working Pattern module setup information *
  '****************************************************************************
  'Check Career Change Working Pattern access
  If gsPersonnelWorkingPatternColumnName <> "" Then
    'Career Change Working Pattern - Static Working Pattern Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, _
                              gsPersonnelWorkingPatternColumnName, strTableColumn) Then
      mstrSQLSelect_PersonnelStaticWP = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableWPs
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
  Else
    'Career Change Working Pattern - Historic Working Pattern Table
    Set objColumn = GetColumnPrivileges(gsPersonnelHWorkingPatternTableName)
    
    'Career Change Working Pattern - Historic Working Pattern Column
    pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternColumnName)
    If pblnColumnOK Then
      pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternColumnName).AllowSelect
    End If
    If pblnColumnOK = False Then
      GoTo DisableWPs
    End If

    'Career Change Working Pattern - Historic Working Pattern Effective Date Column
    pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternDateColumnName)
    If pblnColumnOK Then
      pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternDateColumnName).AllowSelect
    End If
    If pblnColumnOK = False Then
      GoTo DisableWPs
    End If

  End If

  CheckPermission_WPInfo = True
  
TidyUpAndExit:
  Set objTable = Nothing
  Set objColumn = Nothing
  Exit Function

DisableWPs:
  mblnDisableWPs = True
  WorkingDaysOnly = False
  mblnWorkingPatterns = False
  CheckPermission_WPInfo = False
  GoTo TidyUpAndExit

End Function

Private Function GetPersonnelRecordSet() As Boolean

  On Error GoTo PersonnelERROR
  
  Dim lngCount As Long
  Dim sSQL As String
  Dim prstPersonnelData As Recordset
  
  If Not mblnFailReport Then
    sSQL = vbNullString
    sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStartDate & " AS 'StartDate', " & vbCrLf
    sSQL = sSQL & "      " & mstrSQLSelect_PersonnelLeavingDate & " AS 'LeavingDate' " & vbCrLf
    sSQL = sSQL & "FROM " & gsPersonnelTableName & vbCrLf
    For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
      '<Personnel CODE>
      If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
        sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbCrLf
        sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbCrLf
      End If
    Next lngCount
    sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelID
  
    ' Get the start and leaving date
    Set prstPersonnelData = datGeneral.GetRecords(sSQL)
  
    If Not prstPersonnelData.BOF And Not prstPersonnelData.EOF Then
      lblStartDate.Caption = IIf(IsNull(prstPersonnelData.Fields("StartDate").Value), "<None>", Format(prstPersonnelData.Fields("StartDate").Value, DateFormat))
      lblEndDate.Caption = IIf(IsNull(prstPersonnelData.Fields("LeavingDate").Value), "<None>", Format(prstPersonnelData.Fields("LeavingDate").Value, DateFormat))
      
      mdtLeavingDate = IIf(IsNull(prstPersonnelData.Fields("LeavingDate").Value), CDate(vbNull), prstPersonnelData.Fields("LeavingDate").Value)
    End If
    prstPersonnelData.Close
    Set prstPersonnelData = Nothing
  Else
    GoTo PersonnelERROR
  End If
  
  If Me.RegionsEnabled Then
    ' Get the employees current region
    If grtRegionType = rtStaticRegion Then
      ' Its a static region, get it from personnel
      sSQL = vbNullString
      sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStaticRegion & "  AS 'Region'  " & vbCrLf
      sSQL = sSQL & "FROM " & gsPersonnelTableName & vbCrLf
      For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
        '<Personnel CODE>
        If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
          sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbCrLf
          sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbCrLf
        End If
      Next lngCount
      sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelID
      Set prstPersonnelData = datGeneral.GetRecords(sSQL)
    Else
     ' Its a historic region, so get topmost from the history
      Set prstPersonnelData = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & " AS 'Region' " & _
                                                      "FROM " & gsPersonnelHRegionTableRealSource & " " & _
                                                     "WHERE " & gsPersonnelHRegionTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelID & _
                                                     " ORDER BY " & gsPersonnelHRegionDateColumnName & " DESC")
    End If
    
    If Not prstPersonnelData.BOF And Not prstPersonnelData.EOF Then
      lblRegion.Caption = Replace(IIf(IsNull(prstPersonnelData.Fields("Region").Value), "", IIf(prstPersonnelData.Fields("Region").Value = "", "", prstPersonnelData.Fields("Region").Value)), "&", "&&")
    Else
      lblRegion.Caption = "<None>"
    End If
    prstPersonnelData.Close
    Set prstPersonnelData = Nothing
  Else
    'Regions DISABLED
    lblRegion.Caption = vbNullString
    lblRegion.Visible = False
    lblRegionLabel.Visible = False
  End If
  
  If Me.WPsEnabled Then
    ' Get the employees current working pattern
    If gwptWorkingPatternType = wptStaticWPattern Then
      ' Its a static working pattern, get it from personnel
      sSQL = vbNullString
      sSQL = sSQL & "SELECT " & mstrSQLSelect_PersonnelStaticWP & "  AS 'WP'  " & vbCrLf
      sSQL = sSQL & "FROM " & gsPersonnelTableName & vbCrLf
      For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
        '<Personnel CODE>
        If mvarTableViews(0, lngCount) = glngPersonnelTableID Then
          sSQL = sSQL & "     LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbCrLf
          sSQL = sSQL & "     ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbCrLf
        End If
      Next lngCount
      sSQL = sSQL & "WHERE " & gsPersonnelTableName & "." & "ID = " & mlngPersonnelID
      Set prstPersonnelData = datGeneral.GetRecords(sSQL)
    
    Else
     ' Its a historic working pattern, so get topmost from the history
      Set prstPersonnelData = datGeneral.GetRecords("SELECT TOP 1 " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & " AS 'WP' " & _
                                                      "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & _
                                                     "WHERE " & gsPersonnelHWorkingPatternTableRealSource & "." & "ID_" & glngPersonnelTableID & " = " & mlngPersonnelID & _
                                                     "AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " <= '" & Replace(Format(Now, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "' " & _
                                                     "ORDER BY " & gsPersonnelHWorkingPatternDateColumnName & " DESC")
    End If
    
    If Not prstPersonnelData.BOF And Not prstPersonnelData.EOF Then
      If IsNull(prstPersonnelData.Fields("WP").Value) Then
        ASRWorkingPattern1.Value = Space(14)
      Else
        ASRWorkingPattern1.Value = prstPersonnelData.Fields("WP").Value
      End If
      ASRWorkingPattern1.Enabled = False
      ASRWorkingPattern1.BorderStyle = 0
      ASRWorkingPattern1.ForeColor = vbHighlight
    Else
      ' JDM - 14/09/01 - Fault 2830 - Set blank working pattern if nothing found
      ASRWorkingPattern1.Value = Space(14)
      ASRWorkingPattern1.Enabled = False
      ASRWorkingPattern1.BorderStyle = 0
      ASRWorkingPattern1.ForeColor = vbHighlight
    End If
    
    ' JDM - 14/09/01 - Fault 2830 - Set Default working pattern
    strAbsWPattern = ASRWorkingPattern1.Value
  
    prstPersonnelData.Close
    Set prstPersonnelData = Nothing
    
  Else
    'WPs DISABLED
    strAbsWPattern = "SSMMTTWWTTFFSS"
    ASRWorkingPattern1.Value = strAbsWPattern
    ASRWorkingPattern1.Visible = False
    
  End If
  
  Set prstPersonnelData = Nothing
  GetPersonnelRecordSet = True
  Exit Function
  
PersonnelERROR:
  Set prstPersonnelData = Nothing
  GetPersonnelRecordSet = False
  COAMsgBox "Error whilst retrieving the personnel information." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Exit Function

End Function

Private Function RefreshLegend(pblnShowCaptions As Boolean) As Boolean
  
  Dim iCount As Integer
  
  For iCount = 1 To lblColourKey_Colour.Count - 1
    If pblnShowCaptions Then
      LoadKeyCaption
      lblColourKey_Colour(mintMultipleKeyControlIndex).Caption = "."
      Exit Function
    Else
      lblColourKey_Colour(iCount).Caption = ""
    End If
  Next iCount
  
End Function

Private Function LoadKeyCaption() As Boolean

  On Error GoTo errLoadColourKey
  
  Dim rstColourKey As Recordset, strColourKeySQL As String, intCounter As Integer
  
  strColourKeySQL = "SELECT DISTINCT " & gsAbsenceTypeTypeColumnName & ", " & gsAbsenceTypeCalCodeColumnName & " FROM " & gsAbsenceTypeTableName & " ORDER BY " & gsAbsenceTypeTypeColumnName
  Set rstColourKey = datGeneral.GetRecords(strColourKeySQL)

  If rstColourKey.BOF And rstColourKey.EOF Then
    COAMsgBox "You have no absence types defined in your Absence Type table", vbExclamation + vbOKOnly, "Absence Calendar"
    LoadKeyCaption = False
    Exit Function
  End If
  
  Do Until rstColourKey.EOF
    intCounter = intCounter + 1
    If intCounter <= 18 Then
      ' Position the colour box control depending on its index
      If intCounter <= 10 Then
        lblColourKey_Colour(intCounter).Top = 250 * (intCounter)
      Else
        lblColourKey_Colour(intCounter).Top = 250 * (intCounter - 10)
      End If
      
      ' Set the colour box caption and show the label
      If IsNull(rstColourKey.Fields(1).Value) Then
        lblColourKey_Colour(intCounter).Caption = ""
      Else
        lblColourKey_Colour(intCounter).Caption = Replace(rstColourKey.Fields(1).Value, "&", "&&")
      End If
      lblColourKey_Colour(intCounter).Visible = True
    End If
    rstColourKey.MoveNext
  Loop
  ' If we are here, then notify calling procedure of success and exit
  LoadKeyCaption = True
  Set rstColourKey = Nothing
  
  Exit Function
  
errLoadColourKey:
  COAMsgBox "An error has occurred - LoadKeyCaption." & vbCrLf & "Please check your absence module setup.", vbCritical + vbOKOnly, "Absence Calendar"
  LoadKeyCaption = False
End Function
