VERSION 5.00
Begin VB.Form frmAuditOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audit Log Order"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14610
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8009
   Icon            =   "frmAuditOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAccess 
      Caption         =   "Audit Log Columns :"
      Height          =   3210
      Left            =   165
      TabIndex        =   69
      Top             =   4455
      Width           =   4800
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   24
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2700
         Width           =   1395
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   24
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   2700
         Width           =   1095
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   19
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   705
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   20
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1095
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   21
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1500
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   22
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   1905
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   23
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2295
         Width           =   1395
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   19
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   705
         Width           =   1095
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   20
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1095
         Width           =   1095
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   21
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1500
         Width           =   1095
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   22
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   1905
         Width           =   1095
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   23
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   2295
         Width           =   1095
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action :"
         Height          =   195
         Index           =   24
         Left            =   195
         TabIndex        =   89
         Tag             =   "Action"
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date / Time :"
         Height          =   195
         Index           =   19
         Left            =   195
         TabIndex        =   88
         Tag             =   "DateTimeStamp"
         Top             =   765
         Width           =   930
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Group :"
         Height          =   195
         Index           =   20
         Left            =   195
         TabIndex        =   87
         Tag             =   "UserGroup"
         Top             =   1155
         Width           =   915
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         Height          =   195
         Index           =   21
         Left            =   195
         TabIndex        =   86
         Tag             =   "UserName"
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name :"
         Height          =   195
         Index           =   22
         Left            =   195
         TabIndex        =   85
         Tag             =   "ComputerName"
         Top             =   1965
         Width           =   1260
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HR Pro Module :"
         Height          =   195
         Index           =   23
         Left            =   195
         TabIndex        =   84
         Tag             =   "HRProModule"
         Top             =   2355
         Width           =   1155
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asc/Desc :"
         Height          =   195
         Index           =   3
         Left            =   2000
         TabIndex        =   83
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Index           =   0
         Left            =   3500
         TabIndex        =   82
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame fraGroups 
      Caption         =   "Audit Log Columns :"
      Height          =   2800
      Left            =   9850
      TabIndex        =   57
      Top             =   100
      Width           =   4800
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   18
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2300
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   17
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1900
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   16
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1500
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   15
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1100
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   14
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   700
         Width           =   1100
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   18
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2300
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   17
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1900
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   16
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1500
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   15
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1100
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   14
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   700
         Width           =   1400
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Index           =   2
         Left            =   3500
         TabIndex        =   64
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asc/Desc :"
         Height          =   195
         Index           =   1
         Left            =   2000
         TabIndex        =   63
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action :"
         Height          =   195
         Index           =   18
         Left            =   195
         TabIndex        =   62
         Tag             =   "Action"
         Top             =   2355
         Width           =   555
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Login :"
         Height          =   195
         Index           =   17
         Left            =   195
         TabIndex        =   61
         Tag             =   "UserLogin"
         Top             =   1965
         Width           =   855
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Group :"
         Height          =   195
         Index           =   16
         Left            =   195
         TabIndex        =   60
         Tag             =   "GroupName"
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date / Time :"
         Height          =   195
         Index           =   15
         Left            =   195
         TabIndex        =   59
         Tag             =   "DateTimeStamp"
         Top             =   1155
         Width           =   1200
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Index           =   14
         Left            =   195
         TabIndex        =   58
         Tag             =   "UserName"
         Top             =   765
         Width           =   435
      End
   End
   Begin VB.Frame fraPermissions 
      Caption         =   "Audit Log Columns :"
      Height          =   3600
      Left            =   5050
      TabIndex        =   48
      Top             =   100
      Width           =   4800
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   11
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2300
         Width           =   1100
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   11
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2300
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   7
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   705
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   8
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1100
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   9
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1500
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   10
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1900
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   12
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2700
         Width           =   1400
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   13
         Left            =   2000
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3100
         Width           =   1400
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   7
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   700
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   8
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1100
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   9
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1500
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   10
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1900
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   12
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2700
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   13
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3100
         Width           =   1100
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Index           =   11
         Left            =   195
         TabIndex        =   65
         Tag             =   "ColumnName"
         Top             =   2355
         Width           =   630
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Index           =   7
         Left            =   195
         TabIndex        =   56
         Tag             =   "UserName"
         Top             =   765
         Width           =   435
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date / Time :"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   55
         Tag             =   "DateTimeStamp"
         Top             =   1155
         Width           =   930
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Group :"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   54
         Tag             =   "GroupName"
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Index           =   10
         Left            =   195
         TabIndex        =   53
         Tag             =   "ViewTableName"
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action :"
         Height          =   195
         Index           =   12
         Left            =   195
         TabIndex        =   52
         Tag             =   "Action"
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permission :"
         Height          =   195
         Index           =   13
         Left            =   195
         TabIndex        =   51
         Tag             =   "Permission"
         Top             =   3165
         Width           =   855
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asc/Desc :"
         Height          =   195
         Index           =   10
         Left            =   2000
         TabIndex        =   50
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Index           =   9
         Left            =   3500
         TabIndex        =   49
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   400
      Left            =   150
      TabIndex        =   38
      Top             =   3850
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2400
      TabIndex        =   39
      Top             =   3850
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3750
      TabIndex        =   40
      Top             =   3850
      Width           =   1200
   End
   Begin VB.Frame fraRecords 
      Caption         =   "Audit Log Columns :"
      Height          =   3600
      Left            =   150
      TabIndex        =   41
      Top             =   100
      Width           =   4800
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   6
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3100
         Width           =   1100
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   6
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3100
         Width           =   1395
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   5
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2700
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   4
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2300
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   3
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1900
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   2
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1500
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   1
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1100
         Width           =   1100
      End
      Begin VB.ComboBox cboOrder 
         Height          =   315
         Index           =   0
         Left            =   3530
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   1100
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   5
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2700
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   4
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2300
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   3
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1900
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   2
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   1
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1100
         Width           =   1395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   0
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   700
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asc/Desc :"
         Height          =   195
         Index           =   7
         Left            =   1980
         TabIndex        =   68
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Index           =   8
         Left            =   3530
         TabIndex        =   67
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Description :"
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   66
         Tag             =   "RecordDesc"
         Top             =   3165
         Width           =   1725
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Value :"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   47
         Tag             =   "NewValue"
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Value :"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   46
         Tag             =   "OldValue"
         Top             =   2355
         Width           =   1050
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   45
         Tag             =   "ColumnName"
         Top             =   1965
         Width           =   900
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   44
         Tag             =   "TableName"
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date / Time :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   43
         Tag             =   "DateTimeStamp"
         Top             =   1155
         Width           =   1200
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   42
         Tag             =   "UserName"
         Top             =   765
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAuditOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfCancelled As Boolean
Private msSortOrder As String
Private mfLoading As Boolean
Private mfChanged As Boolean

Private mlMin As Long
Private mlMax As Long

Public Sub Initialise(AuditType As audType)
  ' Display the appropriate order controls.
  Dim dblCurrentHeight As Double
  
  Const iFRAMETOP = 100
  Const iFRAMELEFT = 150
  Const iYGAP = 200
  Const iFORMWIDTH = 5200
    
  With fraRecords
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audRecords)
  
    If (AuditType = audRecords) Then
      mlMin = 0
      mlMax = 6
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  With fraPermissions
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audPermissions)
  
    If (AuditType = audPermissions) Then
      mlMin = 7
      mlMax = 13
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  With fraGroups
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audGroups)
  
    If (AuditType = audGroups) Then
      mlMin = 14
      mlMax = 18
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  With fraAccess
    .Top = iFRAMETOP
    .Left = iFRAMELEFT
    .Visible = (AuditType = audAccess)
  
    If (AuditType = audAccess) Then
      mlMin = 19
      mlMax = 24
      dblCurrentHeight = .Top + .Height
    End If
  End With
  
  cmdCancel.Top = dblCurrentHeight + iYGAP
  cmdOk.Top = cmdCancel.Top
  cmdClear.Top = cmdCancel.Top
  
  Me.Height = cmdCancel.Top + cmdCancel.Height + iYGAP + UI.CaptionHeight + (2 * UI.YFrame)
  Me.Width = iFORMWIDTH
    
  
  
End Sub

Private Sub cboOrder_Click(Index As Integer)
  Dim lCount As Long
  Dim lValue As Long
  Dim bGot As Boolean
  Dim lNew As Long
  Dim lOrig As Long

  If mfLoading Then
    Exit Sub
  End If
  
  mfChanged = True

  With cboOrder(Index)
  
    bGot = Len(.Tag) > 0
    If bGot Then
      lOrig = .Tag
    End If
    
    If .Text = "Unsorted" Then
      mfLoading = True
      .Tag = ""
      .Clear
      cboType(Index).Clear
      For lCount = mlMin To mlMax
        If Val(cboOrder(lCount).Text) > lOrig Then
          lValue = Val(cboOrder(lCount).Text) - 1
          cboOrder(lCount).Clear
          cboOrder(lCount).AddItem lValue
          cboOrder(lCount).Text = lValue
          cboOrder(lCount).Tag = lValue
        End If
      Next
      mfLoading = False
    Else
      lValue = Val(.Text)
      
      If lValue = 0 Then
        If bGot Then
          mfLoading = True
          .ListIndex = 0
          mfLoading = False
          Exit Sub
        End If
      Else
        .Clear
        .Tag = lValue
        .AddItem lValue
        mfLoading = True
        .Text = lValue
        
        If bGot Then
          For lCount = mlMin To mlMax
            If lCount <> Index Then
              If Val(cboOrder(lCount).Text) > 0 Then
                If lValue = 1 Then
                  If Val(cboOrder(lCount).Text) < lOrig Then
                    lNew = Val(cboOrder(lCount).Text) + 1
                    cboOrder(lCount).Clear
                    cboOrder(lCount).Tag = lNew
                    cboOrder(lCount).AddItem lNew
                    cboOrder(lCount).Text = lNew
                  End If
                Else
                  If lValue > lOrig Then
                    If Val(cboOrder(lCount).Text) > lOrig And Val(cboOrder(lCount).Text) <= lValue Then
                      lNew = Val(cboOrder(lCount).Text) - 1
                      cboOrder(lCount).Clear
                      cboOrder(lCount).Tag = lNew
                      cboOrder(lCount).AddItem lNew
                      cboOrder(lCount).Text = lNew
                    End If
                  Else
                    If Val(cboOrder(lCount).Text) < lOrig And Val(cboOrder(lCount).Text) >= lValue Then
                        lNew = Val(cboOrder(lCount).Text) + 1
                      cboOrder(lCount).Clear
                      cboOrder(lCount).Tag = lNew
                      cboOrder(lCount).AddItem lNew
                      cboOrder(lCount).Text = lNew
                    End If
                  End If
                End If
              End If
            End If
          Next
        Else
          cboType(Index).AddItem "Ascending"
          cboType(Index).ListIndex = 0
          
          For lCount = mlMin To mlMax
            If lCount <> Index Then
              If Val(cboOrder(lCount).Text) > 0 Then
                If lValue = 1 Then
                  lNew = Val(cboOrder(lCount).Text) + 1
                  cboOrder(lCount).Clear
                  cboOrder(lCount).Tag = lNew
                  cboOrder(lCount).AddItem lNew
                  cboOrder(lCount).Text = lNew
                Else
                  If Val(cboOrder(lCount).Text) > 1 Then
                    If Val(cboOrder(lCount).Text) >= lValue Then
                      lNew = Val(cboOrder(lCount).Text) + 1
                      cboOrder(lCount).Clear
                      cboOrder(lCount).Tag = lNew
                      cboOrder(lCount).AddItem lNew
                      cboOrder(lCount).Text = lNew
                    End If
                  End If
                End If
              End If
            End If
          Next
        End If
        
        mfLoading = False
      End If
    End If
  End With
  
  RefreshButton
  
End Sub

Private Sub cboOrder_DropDown(Index As Integer)

  Dim lCount As Long
  Dim lMax As Long
  Dim bGot As Boolean
  Dim lValue As Long
  
  lValue = Val(cboOrder(Index).Text)
  bGot = (lValue > 0)
  cboOrder(Index).Clear
  
  For lCount = mlMin To mlMax
    If Val(cboOrder(lCount).Text) > lMax Then
      lMax = Val(cboOrder(lCount).Text)
    End If
  Next
  
  If lValue > lMax Then
    lMax = lValue
  End If
  
  If lMax > 0 Then
    If Not bGot Then
      lMax = lMax + 1
    End If
      
    For lCount = 1 To lMax
      cboOrder(Index).AddItem lCount
    Next
      
    If bGot Then
      cboOrder(Index).AddItem "Unsorted"
    End If
  Else
    cboOrder(Index).AddItem "1"
  End If
  
  If bGot Then
    mfLoading = True
    cboOrder(Index).Text = lValue
    mfLoading = False
  End If

End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property

Private Sub cboType_Click(Index As Integer)
  Dim lCount As Long
  Dim lMax As Long
  Dim lOrig As Long
  Dim lValue As Long
  
  If mfLoading Then
    Exit Sub
  End If
      
  mfChanged = True

  If cboType(Index).ListIndex >= 0 Then
  
    If Len(cboType(Index).Text) = 0 Then
      cboType(Index).Text = cboType(Index).Tag
      Exit Sub
    End If
    
    If cboType(Index).ItemData(cboType(Index).ListIndex) = 2 Then
      mfLoading = True
      cboOrder(Index).Tag = ""
      lOrig = Val(cboOrder(Index).Text)
      cboOrder(Index).Clear
      cboType(Index).Clear
      For lCount = mlMin To mlMax
        If Val(cboOrder(lCount).Text) > lOrig Then
          lValue = Val(cboOrder(lCount).Text) - 1
          cboOrder(lCount).Clear
          cboOrder(lCount).AddItem lValue
          cboOrder(lCount).Text = lValue
          cboOrder(lCount).Tag = lValue
        End If
      Next
      mfLoading = False
      RefreshButton
      Exit Sub
    End If
    
    If Val(cboOrder(Index).Text) = 0 Then
      For lCount = mlMin To mlMax
        If Val(cboOrder(lCount).Text) > lMax Then
          lMax = Val(cboOrder(lCount).Text)
        End If
      Next
        
      lMax = IIf(lMax = 0, 1, lMax + 1)
      mfLoading = True
      cboOrder(Index).AddItem lMax
      cboOrder(Index).Tag = lMax
      cboOrder(Index).Text = lMax
      mfLoading = False
    End If
  
  End If

  RefreshButton
  
End Sub

Private Sub cboType_DropDown(Index As Integer)
  Dim sText As String
  
  sText = cboType(Index).Text
  
  With cboType(Index)
    .Tag = .Text
    .Clear
    .AddItem "Ascending"
    .ItemData(.NewIndex) = 0
    .AddItem "Descending"
    .ItemData(.NewIndex) = 1
    If Val(cboOrder(Index).Text) > 0 Then
      .AddItem "Unsorted"
      .ItemData(.NewIndex) = 2
    End If
  End With
  
  If Len(sText) > 0 Then
    mfLoading = True
    cboType(Index).Text = sText
    mfLoading = False
  End If

End Sub

Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide

End Sub

Private Sub cmdClear_Click()
  Dim lCount As Long
  
  mfLoading = True
  For lCount = mlMin To mlMax
    cboType(lCount).Tag = ""
    cboType(lCount).Clear
    cboOrder(lCount).Tag = ""
    cboOrder(lCount).Clear
  Next
  mfLoading = False

  mfChanged = True
  RefreshButton
  
End Sub

Private Sub cmdOK_Click()
  SaveOrder
  mfCancelled = False
  Me.Hide

End Sub

Private Sub Form_Activate()

mfChanged = False
RefreshButton

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
    Cancel = True
    Me.Hide
  End If

End Sub

Public Property Get SortOrder() As String
  SortOrder = msSortOrder

End Property

Private Sub SaveOrder()
  Dim lCount As Long
  Dim lMax As Long
  Dim lRunning As Long
  
  For lCount = mlMin To mlMax
    If Val(cboOrder(lCount).Text) > lMax Then
      lMax = Val(cboOrder(lCount).Text)
    End If
  Next
  
  If lMax = 0 Then
    msSortOrder = ""
    Exit Sub
  End If
  
  msSortOrder = "Order By "
  For lRunning = 1 To lMax
    For lCount = mlMin To mlMax
      If Val(cboOrder(lCount).Text) = lRunning Then
        msSortOrder = msSortOrder & lblColumn(lCount).Tag
        msSortOrder = msSortOrder & IIf(cboType(lCount).Text = "Descending", " DESC", "") & ", "
        Exit For
      End If
    Next
  Next
  
  msSortOrder = Mid$(msSortOrder, 1, Len(msSortOrder) - 2)

End Sub

Private Sub RefreshButton()

  cmdOk.Enabled = mfChanged
  
  If fraRecords.Visible Then
  
    Me.cmdClear.Enabled = IIf((cboType(0).Text = "") And _
                          (cboType(1).Text = "") And _
                          (cboType(2).Text = "") And _
                          (cboType(3).Text = "") And _
                          (cboType(4).Text = "") And _
                          (cboType(5).Text = "") And _
                          (cboType(6).Text = ""), False, True)
    
  ElseIf fraPermissions.Visible Then
  
    Me.cmdClear.Enabled = IIf((cboType(7).Text = "") And _
                          (cboType(8).Text = "") And _
                          (cboType(9).Text = "") And _
                          (cboType(10).Text = "") And _
                          (cboType(11).Text = "") And _
                          (cboType(12).Text = "") And _
                          (cboType(13).Text = ""), False, True)
  
  
  ElseIf fraGroups.Visible Then
  
    Me.cmdClear.Enabled = IIf((cboType(14).Text = "") And _
                          (cboType(15).Text = "") And _
                          (cboType(16).Text = "") And _
                          (cboType(17).Text = "") And _
                          (cboType(18).Text = ""), False, True)

  ElseIf fraAccess.Visible Then
  
    Me.cmdClear.Enabled = IIf((cboType(19).Text = "") And _
                          (cboType(20).Text = "") And _
                          (cboType(21).Text = "") And _
                          (cboType(22).Text = "") And _
                          (cboType(23).Text = "") And _
                          (cboType(24).Text = ""), False, True)

  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


