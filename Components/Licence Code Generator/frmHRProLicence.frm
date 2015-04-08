VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmHRProLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OpenHR Licence Generator"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRProLicence.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7590
      Left            =   75
      TabIndex        =   46
      Top             =   60
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   13388
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Ge&nerate Licence"
      TabPicture(0)   =   "frmHRProLicence.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCustomerDetails(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraLicenceGenerate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraLicenceType"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Support"
      TabPicture(1)   =   "frmHRProLicence.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraLicenceType 
         Caption         =   "Licence Details :"
         Height          =   1665
         Left            =   150
         TabIndex        =   50
         Top             =   420
         Width           =   6585
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "frmHRProLicence.frx":0044
            Left            =   1485
            List            =   "frmHRProLicence.frx":0057
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   705
            Width           =   2850
         End
         Begin GTMaskDate.GTMaskDate txtExpiryDate 
            Height          =   300
            Left            =   1485
            TabIndex        =   3
            Top             =   1170
            Width           =   2850
            _Version        =   65537
            _ExtentX        =   5027
            _ExtentY        =   529
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtCustomerNo 
            Height          =   315
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   1
            Top             =   285
            Width           =   2850
         End
         Begin VB.Label lblModel 
            Caption         =   "Model :"
            Height          =   240
            Left            =   150
            TabIndex        =   54
            Top             =   765
            Width           =   795
         End
         Begin VB.Label lblExpiryDate 
            Caption         =   "Expiry Date :"
            Height          =   225
            Left            =   135
            TabIndex        =   52
            Top             =   1215
            Width           =   1185
         End
         Begin VB.Label lblCustomerNo 
            AutoSize        =   -1  'True
            Caption         =   "Customer No. :"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   330
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   4485
         Begin VB.CommandButton cmdSuppGenerate 
            Caption         =   "&Generate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2580
            TabIndex        =   35
            Top             =   960
            Width           =   1680
         End
         Begin VB.CommandButton cmdSuppClipboard 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2040
            Picture         =   "frmHRProLicence.frx":00C9
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Copy to clipboard"
            Top             =   960
            Width           =   360
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   215
            TabIndex        =   36
            Top             =   1560
            Width           =   4095
            Begin VB.TextBox txtSupportOutput 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   37
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportOutput 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   840
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   39
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportOutput 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   41
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportOutput 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   2520
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   43
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportOutput 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   3365
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   45
               Top             =   0
               Width           =   700
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   11
               Left            =   735
               TabIndex        =   38
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   10
               Left            =   1575
               TabIndex        =   40
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   9
               Left            =   2415
               TabIndex        =   42
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   8
               Left            =   3260
               TabIndex        =   44
               Top             =   45
               Width           =   90
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   215
            TabIndex        =   24
            Top             =   360
            Width           =   4095
            Begin VB.TextBox txtSupportInput 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   3365
               MaxLength       =   4
               TabIndex        =   33
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportInput 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   31
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportInput 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   29
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportInput 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   840
               MaxLength       =   4
               TabIndex        =   27
               Top             =   0
               Width           =   700
            End
            Begin VB.TextBox txtSupportInput 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   0
               MaxLength       =   4
               TabIndex        =   25
               Top             =   0
               Width           =   700
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   3260
               TabIndex        =   32
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   2415
               TabIndex        =   30
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1575
               TabIndex        =   28
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   735
               TabIndex        =   26
               Top             =   45
               Width           =   90
            End
         End
      End
      Begin VB.Frame fraLicenceGenerate 
         Caption         =   "Licence Key :"
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   6060
         Width           =   6585
         Begin VB.CommandButton cmdClipboard 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3120
            Picture         =   "frmHRProLicence.frx":094B
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Copy to clipboard"
            Top             =   765
            Width           =   360
         End
         Begin VB.CommandButton cmdGenerate 
            Caption         =   "&Generate"
            Height          =   400
            Left            =   3660
            TabIndex        =   11
            Top             =   765
            Width           =   1680
         End
         Begin VB.CommandButton cmdRead 
            Caption         =   "&Read"
            Height          =   400
            Left            =   1260
            TabIndex        =   9
            Top             =   765
            Width           =   1680
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   6255
            Begin VB.TextBox txtLicence 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   5235
               MaxLength       =   6
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   0
               Width           =   825
            End
            Begin VB.TextBox txtLicence 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   4260
               MaxLength       =   6
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   0
               Width           =   825
            End
            Begin VB.TextBox txtLicence 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   45
               MaxLength       =   6
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   0
               Width           =   825
            End
            Begin VB.TextBox txtLicence 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1080
               MaxLength       =   6
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   0
               Width           =   825
            End
            Begin VB.TextBox txtLicence 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   2200
               MaxLength       =   6
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   0
               Width           =   825
            End
            Begin VB.TextBox txtLicence 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   3240
               MaxLength       =   6
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   12
               Left            =   5115
               TabIndex        =   55
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   4
               Left            =   4110
               TabIndex        =   49
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   7
               Left            =   917
               TabIndex        =   17
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   6
               Left            =   2007
               TabIndex        =   19
               Top             =   45
               Width           =   90
            End
            Begin VB.Label lblLicence 
               AutoSize        =   -1  'True
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   3087
               TabIndex        =   21
               Top             =   45
               Width           =   90
            End
         End
      End
      Begin VB.Frame fraCustomerDetails 
         Caption         =   "Details :"
         Height          =   3720
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   2265
         Width           =   6585
         Begin VB.TextBox txtHeadcount 
            Height          =   285
            Left            =   1485
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "0"
            Top             =   2325
            Width           =   810
         End
         Begin VB.TextBox txtSSIUsers 
            Height          =   315
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   6
            Text            =   "0"
            Top             =   1110
            Width           =   810
         End
         Begin VB.TextBox txtDatUsers 
            Height          =   315
            Left            =   1500
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "0"
            Top             =   315
            Width           =   810
         End
         Begin VB.ListBox lstModules 
            Height          =   3210
            ItemData        =   "frmHRProLicence.frx":11CD
            Left            =   2580
            List            =   "frmHRProLicence.frx":11CF
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   315
            Width           =   3735
         End
         Begin VB.TextBox txtIntUsers 
            Height          =   315
            Left            =   1500
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "0"
            Top             =   705
            Width           =   810
         End
         Begin VB.Label lblHeadcount 
            Caption         =   "Headcount :"
            Height          =   420
            Left            =   105
            TabIndex        =   53
            Top             =   2385
            Width           =   1065
         End
         Begin VB.Label lblNoUsers 
            AutoSize        =   -1  'True
            Caption         =   "SSI Users :"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   1155
            Width           =   1065
         End
         Begin VB.Label lblNoUsers 
            AutoSize        =   -1  'True
            Caption         =   "Dat Users :"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblNoUsers 
            AutoSize        =   -1  'True
            Caption         =   "DMI Users :"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   750
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frmHRProLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Private mstrAllowedInputCharacters As String

Private Sub PopulateModules(lstTemp As ListBox)

  Dim lngBit As Long
  
  lngBit = 1
  With lstTemp
    .AddItem "Personnel  ": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Recruitment": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Absence    ": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Training   ": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Intranet   ": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "AFD        ": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Full SysMgr": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "CMG        ": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Quick Address": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Payroll (Shared Table)": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Workflow": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "V1 Integration": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Mobile Interface": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Fusion Integration": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "XML Exports": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "OpenLMS Integration": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "9-Box Grid Reports": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Editable Grids": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Power Customisation Pack": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    ' Deselect all rows in Module box for TFS 14733
    .ListIndex = -1
  End With

End Sub

Private Sub cboType_Click()

  If cboType.ListIndex = 0 Then
    txtHeadcount.Text = 0
  Else
    txtSSIUsers.Text = 0
  End If

  txtSSIUsers.Enabled = IIf(cboType.ListIndex = 0, True, False)
  txtHeadcount.Enabled = Not txtSSIUsers.Enabled

End Sub

Private Sub cmdClipboard_Click()
  Clipboard.Clear
  Clipboard.SetText Me.LicenceKey
End Sub

Private Sub cmdRead_Click()

  Dim objLicence As New clsLicence
  Dim lngCustNo As Long
  Dim lngUsers As Long
  Dim lngModules As Long
  Dim lngCount As Long

  With objLicence
    .ValidateCreationDate = False
    .LicenceKey = Me.LicenceKey
    
    If Not .IsValid Then
      MsgBox ("Invalid Key")
      Exit Sub
    End If

    If (.CustomerNo < 1000 Or .CustomerNo > 9999) And vbCompiled Then
      MsgBox "Invalid Licence Key", vbExclamation
    Else
      txtCustomerNo.Text = CStr(.CustomerNo)
      txtDatUsers.Text = CStr(.DATUsers)
      txtIntUsers.Text = CStr(.DMIMUsers)
      txtSSIUsers.Text = CStr(.SSIUsers)
      txtHeadcount.Text = CStr(.Headcount)
      
      If IsDate(.ExpiryDate) And Year(.ExpiryDate) > 1900 Then
        txtExpiryDate.Text = .ExpiryDate
      Else
        txtExpiryDate.Text = ""
      End If
      cboType.ListIndex = .LicenceType

      For lngCount = 0 To lstModules.ListCount - 1
        lstModules.Selected(lngCount) = (.Modules And lstModules.ItemData(lngCount))
      Next
    
    End If
  End With

  Set objLicence = Nothing

End Sub

Private Sub cmdSuppClipboard_Click()
  Clipboard.Clear
  Clipboard.SetText _
      txtSupportOutput(0).Text & "-" & _
      txtSupportOutput(1).Text & "-" & _
      txtSupportOutput(2).Text & "-" & _
      txtSupportOutput(3).Text & "-" & _
      txtSupportOutput(4).Text
End Sub

Private Sub Form_Load()

  SSTab1.Tab = 0
  Frame2.BackColor = Me.BackColor
  Frame3.BackColor = Me.BackColor
  Frame4.BackColor = Me.BackColor
  PopulateModules lstModules
  'PopulateModules lstModules

  mstrAllowedInputCharacters = GenerateAlphaString

  'Only show the read licence tab if in development!
  'On Local Error Resume Next
  'Err.Clear
  'Debug.Print 1 / 0
  'Me.SSTab1.TabVisible(1) = (Err.Number > 0)
  
  cboType.ListIndex = 0
    
  Set gSysTray = New clsSysTray
  Set gSysTray.SourceWindow = Me
  gSysTray.ChangeIcon Me.Icon

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  gSysTray.RemoveFromSysTray
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then
    gSysTray.MinToSysTray
  End If
End Sub

Private Sub gSysTray_LButtonDblClk()
  If Me.WindowState = vbMinimized Then
    gSysTray.RemoveFromSysTray
    Me.WindowState = vbNormal
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  fraCustomerDetails(0).Enabled = (SSTab1.Tab = 0)
  fraLicenceGenerate.Enabled = (SSTab1.Tab = 0)

  'fraLicenceRead.Enabled = (SSTab1.Tab = 1)
  'fraCustomerDetails(1).Enabled = (SSTab1.Tab = 1)

End Sub


Private Sub cmdGenerate_Click()

  Dim objLicence As clsLicenceWrite2
  Dim lngCount As Long
  Dim lngModules As Long

  'Validate customer number...
  With txtCustomerNo
    If Len(.Text) <> 4 Or Val(.Text) < 1000 Then
      MsgBox "Invalid Customer Number", vbExclamation
      .SetFocus
      Exit Sub
    End If
  End With


  'Check with modules have been selected...
  With lstModules
    lngModules = 0
    For lngCount = 0 To .ListCount - 1
      If .Selected(lngCount) Then
        lngModules = lngModules + .ItemData(lngCount)
      End If
    Next

    If lngModules = 0 Then
      MsgBox "No Modules selected", vbExclamation
      .SetFocus
      Exit Sub
    End If
  End With

  
  Set objLicence = New clsLicenceWrite2

  With objLicence

    .CustomerNo = Val(txtCustomerNo.Text)
    .DATUsers = Val(txtDatUsers.Text)
    .DMIMUsers = Val(txtIntUsers.Text)
    .SSIUsers = Val(txtSSIUsers.Text)
    .Headcount = Val(txtHeadcount.Text)
    
    If IsDate(txtExpiryDate.DateValue) Then
      .ExpiryDate = txtExpiryDate.DateValue
    End If
    
    .LicenceType = cboType.ListIndex
   
    .Modules = lngModules

    Me.LicenceKey = .LicenceKey2

  End With

  Set objLicence = Nothing

End Sub


Private Function GenerateAlphaString() As String

  Dim strOutput As String
  Dim lngCount As Long
  Dim lngLoop As Long

  'Only allow these characters...
  strOutput = vbNullString

  For lngCount = Asc("A") + lngLoop To Asc("Z")
    strOutput = strOutput & Chr(lngCount)
  Next

  For lngCount = Asc("0") + lngLoop To Asc("9")
    strOutput = strOutput & Chr(lngCount)
  Next

  GenerateAlphaString = strOutput

End Function

Private Sub txtCustomerNo_GotFocus()
  With txtCustomerNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDatUsers_GotFocus()
  With txtDatUsers
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtIntUsers_GotFocus()
  With txtIntUsers
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLicence_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSSIUsers_GotFocus()
  With txtSSIUsers
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLicence_Change(Index As Integer)

  If Len(txtLicence(Index).Text) >= 3 And txtLicence(Index).SelStart = 4 Then
    If Index < txtLicence.UBound Then
      txtLicence(Index + 1).SetFocus
    End If
  End If

End Sub

Private Sub txtLicence_GotFocus(Index As Integer)
  With txtLicence(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLicence_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyV And (Shift And vbCtrlMask) Then
    If Clipboard.GetText Like "??????-??????-??????-??????-??????-??????" Then
      LicenceKey = Clipboard.GetText
      KeyCode = 0
      Shift = 0
    End If
  End If

End Sub

Private Sub txtSupportInput_Change(Index As Integer)

  If Len(txtSupportInput(Index).Text) >= 4 And txtSupportInput(Index).SelStart = 4 Then
    If Index < txtSupportInput.UBound Then
      txtSupportInput(Index + 1).SetFocus
    End If
  End If

End Sub

Private Sub txtSupportInput_GotFocus(Index As Integer)
  With txtSupportInput(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSupportInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  'Check if a user is trying to paste in a whole licence key
  'If they are, then separate it into each text box.
  If KeyCode = vbKeyV And (Shift And vbCtrlMask) Then
    If Clipboard.GetText Like "????-????-????-????-????" Then
      txtSupportInput(0).Text = Mid(Clipboard.GetText, 1, 4)
      txtSupportInput(1).Text = Mid(Clipboard.GetText, 6, 4)
      txtSupportInput(2).Text = Mid(Clipboard.GetText, 11, 4)
      txtSupportInput(3).Text = Mid(Clipboard.GetText, 16, 4)
      txtSupportInput(4).Text = Mid(Clipboard.GetText, 21, 4)
      KeyCode = 0
      Shift = 0
    End If
  End If

End Sub

Private Sub txtSupportInput_KeyPress(Index As Integer, KeyAscii As Integer)

  Dim strChar As String
  
  'Allow control characters...
  If KeyAscii > 31 Then
  
    strChar = UCase(Chr(KeyAscii))
    If InStr(mstrAllowedInputCharacters, strChar) > 0 Then
      KeyAscii = Asc(strChar)
    Else
      KeyAscii = 0
    End If
  
  End If

End Sub

Public Property Get LicenceKey() As String
  
  Dim lngCount As Long
  
  LicenceKey = vbNullString
  For lngCount = txtLicence.LBound To txtLicence.UBound
    LicenceKey = LicenceKey & _
      IIf(LicenceKey <> vbNullString, "-", "") & _
      txtLicence(lngCount).Text
  Next

End Property

Public Property Let LicenceKey(ByVal strNewValue As String)

  Dim lngCount As Long
 
  If strNewValue Like "??????-??????-??????-??????-??????-??????" Then
    txtLicence(0).Text = Mid(strNewValue, 1, 6)
    txtLicence(1).Text = Mid(strNewValue, 8, 6)
    txtLicence(2).Text = Mid(strNewValue, 15, 6)
    txtLicence(3).Text = Mid(strNewValue, 22, 6)
    txtLicence(4).Text = Mid(strNewValue, 29, 6)
    txtLicence(5).Text = Mid(strNewValue, 36, 6)
  End If

End Property

Private Sub txtSupportOutput_Change(Index As Integer)

  If Len(txtSupportOutput(Index).Text) >= 4 And txtSupportOutput(Index).SelStart = 4 Then
    If Index < txtSupportOutput.UBound Then
      txtSupportOutput(Index + 1).SetFocus
    End If
  End If

End Sub

Private Function vbCompiled() As Boolean

  On Local Error Resume Next
  Err.Clear
  Debug.Print 1 / 0
  vbCompiled = (Err.Number = 0)

End Function
'
'Public Function ConvertStringToNumber2(strInput As String) As Long
'
'  Dim lngRandomDigit As Long
'  Dim strAlphaString As String
'  Dim lngOutput As Long
'  Dim lngFactor As Double
'  Dim lngCount As Long
'
'  On Error GoTo exitf
'
'  lngRandomDigit = Asc(Mid(strInput, 1, 1)) - 64
'  strAlphaString = GenerateAlphaString(lngRandomDigit)
'
'  lngOutput = (InStr(strAlphaString, Mid(strInput, Len(strInput), 1)) - 1)
'
'  lngFactor = 32
'  For lngCount = Len(strInput) - 1 To 2 Step -1
'    lngOutput = lngOutput + _
'      ((InStr(strAlphaString, Mid(strInput, lngCount, 1)) - 1) * lngFactor)
'    lngFactor = lngFactor * 32
'  Next
'
'  ConvertStringToNumber2 = lngOutput
'
'exitf:
'
'End Function
