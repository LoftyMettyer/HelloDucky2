VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLabelTypeDefinition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envelope & Label Template Definition"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1082
   Icon            =   "frmLabelTypeDefinition.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6345
      Left            =   90
      TabIndex        =   33
      Top             =   90
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   11192
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dimensions"
      TabPicture(0)   =   "frmLabelTypeDefinition.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEnvelope"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraLabel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDefinition"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraType"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraMeasurements"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Format"
      TabPicture(1)   =   "frmLabelTypeDefinition.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ImageList1"
      Tab(1).Control(1)=   "picColour"
      Tab(1).Control(2)=   "fraFont(0)"
      Tab(1).Control(3)=   "fraFont(1)"
      Tab(1).Control(4)=   "fraPreview"
      Tab(1).ControlCount=   5
      Begin VB.Frame fraPreview 
         Caption         =   "Preview :"
         Height          =   3690
         Left            =   -69960
         TabIndex        =   95
         Top             =   405
         Width           =   4230
         Begin VB.PictureBox picPreview 
            BackColor       =   &H80000005&
            Height          =   2535
            Left            =   195
            ScaleHeight     =   2475
            ScaleWidth      =   3750
            TabIndex        =   96
            Top             =   645
            Width           =   3810
            Begin VB.Label lblPreviewStandard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Standard Text"
               Height          =   195
               Index           =   2
               Left            =   105
               TabIndex        =   100
               Top             =   1125
               Width           =   1035
            End
            Begin VB.Label lblPreviewStandard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Standard Text"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   99
               Top             =   915
               Width           =   1035
            End
            Begin VB.Label lblPreviewStandard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Standard Text"
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   98
               Top             =   705
               Width           =   1035
            End
            Begin VB.Label lblPreviewHeading 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Heading Text"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   90
               TabIndex        =   97
               Top             =   30
               Width           =   1410
            End
         End
      End
      Begin VB.Frame fraFont 
         Caption         =   "Standard Text : "
         Height          =   1680
         Index           =   1
         Left            =   -74850
         TabIndex        =   91
         Top             =   2415
         Width           =   4740
         Begin VB.ComboBox cboFontName 
            Height          =   315
            Index           =   1
            Left            =   1020
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   315
            Width           =   2025
         End
         Begin VB.CheckBox chkFontUnderLine 
            Caption         =   "U&nderline"
            Height          =   195
            Index           =   1
            Left            =   3345
            TabIndex        =   86
            Top             =   645
            Width           =   1215
         End
         Begin VB.CheckBox chkFontBold 
            Caption         =   "Bo&ld"
            Height          =   195
            Index           =   1
            Left            =   3345
            TabIndex        =   85
            Top             =   330
            Width           =   1215
         End
         Begin VB.CheckBox chkFontItalic 
            Caption         =   "It&alic"
            Height          =   195
            Index           =   1
            Left            =   3345
            TabIndex        =   87
            Top             =   960
            Width           =   1170
         End
         Begin VB.ComboBox cboFontSize 
            Height          =   315
            Index           =   1
            ItemData        =   "frmLabelTypeDefinition.frx":0044
            Left            =   1020
            List            =   "frmLabelTypeDefinition.frx":0046
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   720
            Width           =   2025
         End
         Begin MSComctlLib.ImageCombo cboFontColour 
            Height          =   330
            Index           =   1
            Left            =   1020
            TabIndex        =   84
            Top             =   1110
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin VB.Label lblFontName 
            Caption         =   "Font :"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   94
            Top             =   375
            Width           =   1290
         End
         Begin VB.Label lblFontColour 
            Caption         =   "Colour :"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   93
            Top             =   1185
            Width           =   1005
         End
         Begin VB.Label lblFontSize 
            Caption         =   "Size :"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   92
            Top             =   780
            Width           =   600
         End
      End
      Begin VB.Frame fraFont 
         Caption         =   "Heading Text : "
         Height          =   1680
         Index           =   0
         Left            =   -74850
         TabIndex        =   75
         Top             =   400
         Width           =   4740
         Begin VB.ComboBox cboFontSize 
            Height          =   315
            Index           =   0
            ItemData        =   "frmLabelTypeDefinition.frx":0048
            Left            =   1020
            List            =   "frmLabelTypeDefinition.frx":004A
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   720
            Width           =   2025
         End
         Begin VB.CheckBox chkFontItalic 
            Caption         =   "I&talic"
            Height          =   195
            Index           =   0
            Left            =   3345
            TabIndex        =   81
            Top             =   960
            Width           =   1110
         End
         Begin VB.CheckBox chkFontBold 
            Caption         =   "&Bold"
            Height          =   195
            Index           =   0
            Left            =   3345
            TabIndex        =   79
            Top             =   330
            Width           =   1215
         End
         Begin VB.CheckBox chkFontUnderLine 
            Caption         =   "&Underline"
            Height          =   195
            Index           =   0
            Left            =   3345
            TabIndex        =   80
            Top             =   645
            Width           =   1215
         End
         Begin VB.ComboBox cboFontName 
            Height          =   315
            Index           =   0
            Left            =   1020
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   315
            Width           =   2025
         End
         Begin MSComctlLib.ImageCombo cboFontColour 
            Height          =   330
            Index           =   0
            Left            =   1020
            TabIndex        =   78
            Top             =   1110
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin VB.Label lblFontSize 
            Caption         =   "Size :"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   90
            Top             =   780
            Width           =   600
         End
         Begin VB.Label lblFontColour 
            Caption         =   "Colour :"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   89
            Top             =   1185
            Width           =   1005
         End
         Begin VB.Label lblFontName 
            Caption         =   "Font :"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   88
            Top             =   375
            Width           =   1290
         End
      End
      Begin VB.PictureBox picColour 
         Height          =   255
         Left            =   -66780
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   74
         Top             =   5625
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame fraMeasurements 
         Caption         =   "Measurement : "
         Height          =   2265
         Left            =   150
         TabIndex        =   63
         Top             =   3945
         Width           =   2820
         Begin VB.OptionButton optMeasurement 
            Caption         =   "&Points"
            Height          =   345
            Index           =   3
            Left            =   315
            TabIndex        =   9
            Top             =   1500
            Width           =   1665
         End
         Begin VB.OptionButton optMeasurement 
            Caption         =   "I&nches"
            Height          =   345
            Index           =   2
            Left            =   315
            TabIndex        =   8
            Top             =   1110
            Width           =   1665
         End
         Begin VB.OptionButton optMeasurement 
            Caption         =   "&Millimetres"
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   7
            Top             =   735
            Width           =   1665
         End
         Begin VB.OptionButton optMeasurement 
            Caption         =   "Cen&timetres"
            Height          =   345
            Index           =   0
            Left            =   315
            TabIndex        =   6
            Top             =   345
            Value           =   -1  'True
            Width           =   1665
         End
      End
      Begin VB.Frame fraType 
         Caption         =   "Type :"
         Height          =   1335
         Left            =   150
         TabIndex        =   62
         Top             =   2415
         Width           =   2820
         Begin VB.OptionButton optLabelEnvelope 
            Caption         =   "&Label"
            Height          =   255
            Index           =   0
            Left            =   315
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton optLabelEnvelope 
            Caption         =   "&Envelope"
            Height          =   240
            Index           =   1
            Left            =   315
            TabIndex        =   5
            Top             =   750
            Width           =   1620
         End
      End
      Begin VB.Frame fraDefinition 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   150
         TabIndex        =   34
         Top             =   400
         Width           =   9165
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1620
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   705
            Width           =   3000
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   0
            Top             =   315
            Width           =   3000
         End
         Begin VB.OptionButton optReadOnly 
            Caption         =   "&Read Only"
            Height          =   195
            Left            =   6000
            TabIndex        =   3
            Top             =   1200
            Width           =   1470
         End
         Begin VB.OptionButton optReadWrite 
            Caption         =   "Read / &Write"
            Height          =   195
            Left            =   6000
            TabIndex        =   2
            Top             =   810
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6000
            MaxLength       =   30
            TabIndex        =   35
            Top             =   315
            Width           =   3000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Index           =   3
            Left            =   5100
            TabIndex        =   39
            Top             =   810
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   38
            Top             =   750
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   37
            Top             =   365
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Index           =   2
            Left            =   5100
            TabIndex        =   36
            Top             =   365
            Width           =   585
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -66465
         Top             =   5625
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
      End
      Begin VB.Frame fraLabel 
         Caption         =   "Label :"
         Height          =   3795
         Left            =   3105
         TabIndex        =   40
         Top             =   2415
         Width           =   6210
         Begin VB.ComboBox cboLabelPageSize 
            Height          =   315
            ItemData        =   "frmLabelTypeDefinition.frx":004C
            Left            =   1665
            List            =   "frmLabelTypeDefinition.frx":004E
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   315
            Width           =   4290
         End
         Begin VB.CheckBox chkLabelPageLandscape 
            Caption         =   "L&andscape"
            Height          =   240
            Left            =   165
            TabIndex        =   13
            Top             =   1260
            Width           =   2430
         End
         Begin MSComCtl2.UpDown upnLabelPageHeight 
            Height          =   240
            Left            =   5685
            TabIndex        =   41
            Top             =   810
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            Value           =   100
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelPageWidth 
            Height          =   240
            Left            =   2735
            TabIndex        =   42
            Top             =   810
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            Value           =   100
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelNumberDown 
            Height          =   240
            Left            =   5685
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   3225
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   100
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelNumberAcross 
            Height          =   255
            Left            =   5685
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   2760
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   100
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelWidth 
            Height          =   255
            Left            =   5685
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   2295
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelHeight 
            Height          =   255
            Left            =   5685
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1845
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelHorizontalPitch 
            Height          =   240
            Left            =   2735
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2760
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelVerticalPitch 
            Height          =   255
            Left            =   2735
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   3225
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelSideMargin 
            Height          =   255
            Left            =   2735
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   2295
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnLabelTopMargin 
            Height          =   255
            Left            =   2735
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1845
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtLabelPageHeight 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   300
            Left            =   4620
            MaxLength       =   35
            TabIndex        =   12
            Top             =   780
            Width           =   1350
         End
         Begin VB.TextBox txtLabelPageWidth 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   300
            Left            =   1665
            MaxLength       =   35
            TabIndex        =   11
            Top             =   780
            Width           =   1350
         End
         Begin VB.TextBox txtLabelHeight 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   4620
            MaxLength       =   35
            TabIndex        =   18
            Text            =   "1.00 cm"
            Top             =   1815
            Width           =   1350
         End
         Begin VB.TextBox txtLabelWidth 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   4620
            MaxLength       =   35
            TabIndex        =   19
            Text            =   "1.00 cm"
            Top             =   2265
            Width           =   1350
         End
         Begin VB.TextBox txtLabelNumberAcross 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   4620
            MaxLength       =   35
            TabIndex        =   20
            Text            =   "1"
            Top             =   2730
            Width           =   1350
         End
         Begin VB.TextBox txtLabelNumberDown 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   4620
            MaxLength       =   35
            TabIndex        =   21
            Text            =   "1"
            Top             =   3195
            Width           =   1350
         End
         Begin VB.TextBox txtLabelTopMargin 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1665
            MaxLength       =   35
            TabIndex        =   14
            Text            =   "1.00 cm"
            Top             =   1815
            Width           =   1350
         End
         Begin VB.TextBox txtLabelSideMargin 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1665
            MaxLength       =   35
            TabIndex        =   15
            Text            =   "1.00 cm"
            Top             =   2265
            Width           =   1350
         End
         Begin VB.TextBox txtLabelVerticalPitch 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1665
            MaxLength       =   35
            TabIndex        =   17
            Text            =   "1.00 cm"
            Top             =   3195
            Width           =   1350
         End
         Begin VB.TextBox txtLabelHorizontalPitch 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1665
            MaxLength       =   35
            TabIndex        =   16
            Text            =   "1.00 cm"
            Top             =   2730
            Width           =   1350
         End
         Begin VB.Label lblPageSize 
            Caption         =   "Page Size :"
            Height          =   300
            Left            =   165
            TabIndex        =   61
            Top             =   375
            Width           =   1485
         End
         Begin VB.Label lblTopMargin 
            Caption         =   "Top Margin :"
            Height          =   255
            Left            =   165
            TabIndex        =   60
            Top             =   1845
            Width           =   1290
         End
         Begin VB.Label lblSideMargin 
            Caption         =   "Side Margin :"
            Height          =   255
            Left            =   165
            TabIndex        =   59
            Top             =   2325
            Width           =   1290
         End
         Begin VB.Label lblVerticalPitch 
            Caption         =   "Vertical Pitch :"
            Height          =   255
            Left            =   165
            TabIndex        =   58
            Top             =   3270
            Width           =   1290
         End
         Begin VB.Label lblHorizontalPitch 
            Caption         =   "Horizontal Pitch :"
            Height          =   255
            Left            =   165
            TabIndex        =   57
            Top             =   2790
            Width           =   1515
         End
         Begin VB.Label lblLabelHeight 
            Caption         =   "Label Height :"
            Height          =   255
            Left            =   3100
            TabIndex        =   56
            Top             =   1845
            Width           =   1290
         End
         Begin VB.Label lblLabelWidth 
            Caption         =   "Label Width :"
            Height          =   255
            Left            =   3100
            TabIndex        =   55
            Top             =   2325
            Width           =   1290
         End
         Begin VB.Label lblNumberAcross 
            Caption         =   "Number Across :"
            Height          =   255
            Left            =   3100
            TabIndex        =   54
            Top             =   2790
            Width           =   1470
         End
         Begin VB.Label lblNumberDown 
            Caption         =   "Number Down :"
            Height          =   255
            Left            =   3100
            TabIndex        =   53
            Top             =   3270
            Width           =   1425
         End
         Begin VB.Label lblPageWidth 
            Caption         =   "Width :"
            Height          =   255
            Left            =   165
            TabIndex        =   52
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label lblPageHeight 
            Caption         =   "Height :"
            Height          =   255
            Left            =   3100
            TabIndex        =   51
            Top             =   840
            Width           =   990
         End
      End
      Begin VB.Frame fraEnvelope 
         Caption         =   "Envelope :"
         Height          =   3795
         Left            =   3105
         TabIndex        =   64
         Top             =   2415
         Visible         =   0   'False
         Width           =   6210
         Begin VB.ComboBox cboEnvelopePageSize 
            Height          =   315
            ItemData        =   "frmLabelTypeDefinition.frx":0050
            Left            =   1620
            List            =   "frmLabelTypeDefinition.frx":0052
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   315
            Width           =   3980
         End
         Begin VB.CheckBox chkEnvelopeAlign 
            Caption         =   "Align te&xt with longest edge"
            Height          =   240
            Left            =   225
            TabIndex        =   25
            Top             =   1260
            Value           =   1  'Checked
            Width           =   3270
         End
         Begin VB.CheckBox chkEnvelopeAutoTop 
            Caption         =   "A&utomatic"
            Height          =   195
            Left            =   3240
            TabIndex        =   27
            Top             =   1845
            Width           =   1305
         End
         Begin VB.CheckBox chkEnvelopeAutoLeft 
            Caption         =   "&Automatic"
            Height          =   255
            Left            =   3240
            TabIndex        =   29
            Top             =   2280
            Width           =   1425
         End
         Begin MSComCtl2.UpDown upnEnvelopeDefinition 
            Height          =   240
            Index           =   2
            Left            =   2690
            TabIndex        =   65
            Top             =   1815
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            Value           =   100
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnEnvelopeDefinition 
            Height          =   240
            Index           =   3
            Left            =   2690
            TabIndex        =   66
            Top             =   2280
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            Value           =   100
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnEnvelopeDefinition 
            Height          =   240
            Index           =   1
            Left            =   5315
            TabIndex        =   67
            Top             =   810
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            Value           =   100
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upnEnvelopeDefinition 
            Height          =   240
            Index           =   0
            Left            =   2690
            TabIndex        =   68
            Top             =   810
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   423
            _Version        =   393216
            Value           =   100
            OrigLeft        =   2805
            OrigTop         =   3360
            OrigRight       =   3060
            OrigBottom      =   3840
            Increment       =   5
            Max             =   7000
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtEnvelopeDimension 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1620
            TabIndex        =   28
            Text            =   "1.00 cm"
            Top             =   2250
            Width           =   1350
         End
         Begin VB.TextBox txtEnvelopeDimension 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1620
            TabIndex        =   26
            Text            =   "1.00 cm"
            Top             =   1785
            Width           =   1350
         End
         Begin VB.TextBox txtEnvelopeDimension 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   4245
            TabIndex        =   24
            Top             =   780
            Width           =   1350
         End
         Begin VB.TextBox txtEnvelopeDimension 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1620
            TabIndex        =   23
            Top             =   780
            Width           =   1350
         End
         Begin VB.Label Label4 
            Caption         =   "Height :"
            Height          =   255
            Left            =   3220
            TabIndex        =   73
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Width :"
            Height          =   255
            Left            =   225
            TabIndex        =   72
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblEnvelopeSize 
            Caption         =   "Envelope Size : "
            Height          =   330
            Left            =   225
            TabIndex        =   71
            Top             =   375
            Width           =   1410
         End
         Begin VB.Label lblEnvelopeFromTop 
            Caption         =   "From Top :"
            Height          =   270
            Left            =   225
            TabIndex        =   70
            Top             =   1845
            Width           =   1065
         End
         Begin VB.Label lblEnvelopeFromLeft 
            Caption         =   "From Left :"
            Height          =   210
            Left            =   225
            TabIndex        =   69
            Top             =   2325
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton cmdImportFromWord 
      Caption         =   "&Import Template..."
      Height          =   400
      Left            =   75
      TabIndex        =   30
      Top             =   6510
      Width           =   1830
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8370
      TabIndex        =   32
      Top             =   6495
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   400
      Left            =   7110
      TabIndex        =   31
      Top             =   6495
      Width           =   1200
   End
End
Attribute VB_Name = "frmLabelTypeDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrSQLTableDef As String = "ASRSysLabelTypes"
Private Const msngPixelsToCMs As Single = 28.346 '28.35 '28.34375

Private Const msngMinLabelPageWidth As Single = 0.3
Private Const msngMaxLabelPageWidth As Single = 21.59

Private Const msngMinLabelPageHeight As Single = 0.3
Private Const msngMaxLabelPageHeight As Single = 29.7 '55.88  'Points = 1584

Private Const msngMinLabelWidth As Single = 0.3
Private Const msngMaxLabelWidth As Single = 20.99028

Private Const msngMinLabelHeight As Single = 0.3
Private Const msngMaxLabelHeight As Single = 28.575     'Points = 810

Private Const msngLabelMinPageMargin As Single = 0   '  0.63

Private Const msngMaxEnvelopePageWidth As Single = 45.8
Private Const msngMaxEnvelopePageHeight As Single = 32.4

Private Const msngMinEnvelopePageWidth As Single = 0.3
Private Const msngMinEnvelopePageHeight As Single = 0.3

Private Const msngEnvelopeFromTopMin As Single = 3.85
Private Const msngEnvelopeFromLeftMin As Single = 2.55


Private Const mstrCentimetreText = "cm"
Private Const mstrMillimetreText = "mm"
Private Const mstrInchText = "in"
Private Const mstrPixelText = "pt"

Private Const msngCentimetreConversion = 100
Private Const msngMillimetreConversion = 10
Private Const msngInchConversion = 254
'Private Const msngPixelConversion = 3.25945241199478
Private Const msngPixelConversion = 3.52733686067019    ' Fault 6674

Private mstrMeasurementText As String
Private msngConversion As Single

Private mlngLabelDefinitionID As Long
Private mblnCancelled As Boolean
Private malngLabelDimensions(10) As Long
Private malngEnvelopeDimensions(6) As Long

Private datData As HRProDataMgr.clsDataAccess          'DataAccess Class

Private mbIsEnvelope As Boolean                        'Is this definition an envelope

Private mlngTimeStamp As Long
Private mblnReadOnly As Boolean
Private mblnFromCopy As Boolean
Private mblnFormPrint As Boolean
Private mblnDefinitionCreator As Boolean
Private mblnHiddenCalculation As Boolean
Private mblnHiddenPicklistOrFilter As Boolean

Private mavPageLabelDimensions() As Variant
Private mavPageEnvelopeDimensions() As Variant

Private msngRememberVPitch As Single
Private msngRememberHPitch As Single

Private mbIsLoading As Boolean
Private mbPreviouslyWarned As Boolean
Private mbLandscaping As Boolean

Private mobjFormatDetails(1) As HRProDataMgr.clsOutputStyle

Private Sub cboEnvelopePageSize_Change()
  Me.Changed = True
End Sub

Private Sub cboEnvelopePageSize_Click()

  Dim bIsCustom As Boolean
  Dim sngWidth As Single
  Dim sngHeight As Single

  bIsCustom = (cboEnvelopePageSize.ListIndex = cboEnvelopePageSize.ListCount - 1)

  ' Enable / disable
  txtEnvelopeDimension(0).Enabled = bIsCustom And Not mblnReadOnly
  txtEnvelopeDimension(0).BackColor = IIf(bIsCustom And Not mblnReadOnly, vbWhite, vbButtonFace)
  upnEnvelopeDefinition(0).Enabled = bIsCustom And Not mblnReadOnly

  txtEnvelopeDimension(1).Enabled = bIsCustom And Not mblnReadOnly
  txtEnvelopeDimension(1).BackColor = IIf(bIsCustom And Not mblnReadOnly, vbWhite, vbButtonFace)
  upnEnvelopeDefinition(1).Enabled = bIsCustom And Not mblnReadOnly

  sngWidth = mavPageEnvelopeDimensions(cboEnvelopePageSize.ListIndex, 1) * 100
  sngHeight = mavPageEnvelopeDimensions(cboEnvelopePageSize.ListIndex, 2) * 100
   
  ' Put values into display field
  If bIsCustom Then
    upnEnvelopeDefinition(0).Value = upnEnvelopeDefinition(0).Min
    upnEnvelopeDefinition(1).Value = upnEnvelopeDefinition(1).Min
  Else
    upnEnvelopeDefinition(0).Value = sngWidth
    upnEnvelopeDefinition(1).Value = sngHeight
  End If

End Sub

Public Property Get FormPrint() As Boolean
  FormPrint = mblnFormPrint
End Property

Public Property Let FormPrint(ByVal bPrint As Boolean)
  mblnFormPrint = bPrint
End Property

Private Sub cboLabelPageSize_Click()

  Dim bIsCustom As Boolean

  bIsCustom = (cboLabelPageSize.ListIndex = cboLabelPageSize.ListCount - 1)

  ' Enable / disable
  txtLabelPageWidth.Enabled = bIsCustom And Not mblnReadOnly
  txtLabelPageWidth.BackColor = IIf(bIsCustom And Not mblnReadOnly, vbWhite, vbButtonFace)
  upnLabelPageWidth.Enabled = bIsCustom And Not mblnReadOnly

  txtLabelPageHeight.Enabled = bIsCustom And Not mblnReadOnly
  txtLabelPageHeight.BackColor = IIf(bIsCustom And Not mblnReadOnly, vbWhite, vbButtonFace)
  upnLabelPageHeight.Enabled = bIsCustom And Not mblnReadOnly

  ' Put values into display field
  If Not bIsCustom Then
    If chkLabelPageLandscape.Value = vbChecked Then
      upnLabelPageWidth.Value = mavPageLabelDimensions(cboLabelPageSize.ListIndex, 2) * 100
      upnLabelPageHeight.Value = mavPageLabelDimensions(cboLabelPageSize.ListIndex, 1) * 100
    Else
      If mavPageLabelDimensions(cboLabelPageSize.ListIndex, 1) * 100 > upnLabelPageWidth.Max Then
        upnLabelPageWidth.Value = upnLabelPageWidth.Max
      Else
        upnLabelPageWidth.Value = mavPageLabelDimensions(cboLabelPageSize.ListIndex, 1) * 100
      End If
      
      If mavPageLabelDimensions(cboLabelPageSize.ListIndex, 1) * 100 > upnLabelPageHeight.Max Then
        upnLabelPageHeight.Value = upnLabelPageHeight.Max
      Else
        upnLabelPageHeight.Value = mavPageLabelDimensions(cboLabelPageSize.ListIndex, 2) * 100
      End If
    End If
  Else
    'upnLabelPageWidth.Value = upnLabelPageWidth.Min
    'upnLabelPageHeight.Value = upnLabelPageHeight.Min
  End If

  Me.Changed = True

End Sub

Private Sub chkEnvelopeAlign_Click()

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

Private Sub chkEnvelopeAutoLeft_Click()

  txtEnvelopeDimension(3).Enabled = (chkEnvelopeAutoLeft.Value = vbUnchecked)
  txtEnvelopeDimension(3).BackColor = IIf(txtEnvelopeDimension(3).Enabled, vbWhite, vbButtonFace)
  upnEnvelopeDefinition(3).Enabled = txtEnvelopeDimension(3).Enabled

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

Private Sub chkEnvelopeAutoTop_Click()

  txtEnvelopeDimension(2).Enabled = (chkEnvelopeAutoTop.Value = vbUnchecked)
  txtEnvelopeDimension(2).BackColor = IIf(txtEnvelopeDimension(2).Enabled, vbWhite, vbButtonFace)
  upnEnvelopeDefinition(2).Enabled = txtEnvelopeDimension(2).Enabled

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

' Swap the height and width around
Private Sub chkLabelPageLandscape_Click()

  Dim strTemp As String
  Dim iCount As Integer
  Dim sngPitchTemp As Single
     
  ' Temporarily disable the minimum and maximum sizes
  upnLabelPageHeight.Max = 10000
  upnLabelPageWidth.Max = 10000
  upnLabelTopMargin.Max = 10000
  upnLabelSideMargin.Max = 10000
  upnLabelHorizontalPitch.Max = 10000
  upnLabelVerticalPitch.Max = 10000
  upnLabelWidth.Max = 10000
  upnLabelHeight.Max = 10000
  upnLabelNumberAcross.Max = 10000
  upnLabelNumberDown.Max = 10000

  ' JDM - 11/09/03 - Fault 6641 - Problems when swapping landscape/portrait
  upnLabelPageHeight.Min = 0
  upnLabelPageWidth.Min = 0
  upnLabelTopMargin.Min = 0
  upnLabelSideMargin.Min = 0
  upnLabelHorizontalPitch.Min = 0
  upnLabelVerticalPitch.Min = 0
  upnLabelWidth.Min = 0
  upnLabelHeight.Min = 0
  upnLabelNumberAcross.Min = 0
  upnLabelNumberDown.Min = 0
  
  ' Swap page sizes around
  strTemp = upnLabelPageWidth.Value
  upnLabelPageWidth.Value = upnLabelPageHeight.Value
  upnLabelPageHeight.Value = strTemp
  
  If Not mbIsLoading Then
  
    ' Flag to stop controls resetting min/max value during swap over
    mbLandscaping = True
  
    ' Swap the remembered pitch settings
    sngPitchTemp = msngRememberVPitch
    msngRememberVPitch = msngRememberHPitch
    msngRememberHPitch = sngPitchTemp
    
    ' Swap the margins around
    strTemp = upnLabelTopMargin.Value
    upnLabelTopMargin.Value = upnLabelSideMargin.Value
    upnLabelSideMargin.Value = strTemp
    
    ' Swap the pitches around
    strTemp = upnLabelHorizontalPitch.Value
    upnLabelHorizontalPitch.Value = upnLabelVerticalPitch.Value
    upnLabelVerticalPitch.Value = strTemp
    
    ' Swap the width and height of the label
    strTemp = upnLabelWidth.Value
    upnLabelWidth.Value = upnLabelHeight.Value
    upnLabelHeight.Value = strTemp
  
    ' Swap the number across/down
    strTemp = upnLabelNumberAcross.Value
    upnLabelNumberAcross.Value = upnLabelNumberDown.Value
    upnLabelNumberDown.Value = strTemp
  
    ' Reset the maximum values
    mbLandscaping = False
    Me.Changed = True
  End If

  SetLabelMaxAndMinSizes

End Sub

Private Sub cmdCancel_Click()
  Unload Me
'
'  Dim strSQL As String
'  Dim strMBText As String
'  Dim intMBButtons As Integer
'  Dim strMBTitle As String
'  Dim intMBResponse As Integer
'
'  If Me.Changed And Not mblnReadOnly Then
'
'    strMBText = "You have changed the current definition. Save changes ?"
'    intMBButtons = vbQuestion + vbYesNoCancel + vbDefaultButton1
'    strMBTitle = Me.Caption
'    intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)
'
'    Select Case intMBResponse
'    Case vbYes
'      Call cmdOK_Click
'      Exit Sub
'    Case vbCancel
'      Exit Sub
'    End Select
'  End If
'
'
'  Me.Hide
'  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdImportFromWord_Click()

  On Error GoTo ImportFromWord_Error

'  Dim wrdApp As Word.Application
'  Dim wrdDoc As Word.Document
'  Dim wrdLabel As MailingLabel
  Dim wrdApp As Object
  Dim wrdDoc As Object
  Dim wrdLabel As Object
  
  Dim strTmpPath As String
  Dim iErrorCount As Integer
  Dim lngCount As Long
  Dim bOK As Boolean
  Dim iLoop As Integer
  Dim iCount As Integer
  Dim strFileName As String
  Dim sngTempSize1 As Single
  Dim sngTempSize2 As Single
  Dim sngWidth As Single
  Dim sngHeight As Single

  bOK = True
  iErrorCount = 0

  strTmpPath = Space(1024)
  Call GetTempPath(1024, strTmpPath)
  
  strTmpPath = Left(Trim(strTmpPath), Len(Trim(strTmpPath)) - 1)
  strFileName = strTmpPath & "mailmerge.tmp"

  ' Create word application
  Set wrdApp = CreateObject("Word.Application")
  Set wrdDoc = wrdApp.Documents.Add

  With wrdDoc.MailMerge

    ' Clear minimum and maximum sizes for import to run smoothly
    SetLabelMaxAndMinSizesToZero

    ' Create a temporary file to fool Word with it's mail merging
    .CreateDataSource Name:=strFileName

    .OpenDataSource Name:=strFileName, AddToRecentFiles:=False
    .Destination = wdSendToNewDocument

    wrdApp.Visible = True
    wrdApp.Activate
  
    If optLabelEnvelope(0).Value = True Then
  
      ' Labels
      .MainDocumentType = wdMailingLabels
      wrdApp.MailingLabel.LabelOptions
     
      wrdApp.Visible = False
      wrdApp.ActiveDocument.Tables(1).AllowAutoFit = False
  
      'Clear existing definition
      upnLabelTopMargin.Value = upnLabelTopMargin.Min
      upnLabelSideMargin.Value = upnLabelSideMargin.Min
      upnLabelHorizontalPitch.Value = upnLabelHorizontalPitch.Min
      upnLabelVerticalPitch.Value = upnLabelVerticalPitch.Min
      upnLabelWidth.Value = upnLabelWidth.Min
      upnLabelHeight.Value = upnLabelHeight.Min
      upnLabelNumberAcross.Value = upnLabelNumberAcross.Min
      upnLabelNumberDown.Value = upnLabelNumberDown.Min
  
      ' Special error handler to trap for error Microsoft error 5992
      On Error GoTo ResumeToNextLine
  
      ' Vertical Pitch
      If wrdApp.ActiveDocument.Tables(1).Rows.Count > 1 Then
        If wrdApp.ActiveDocument.Tables(1).Rows(1).Height <> wrdApp.ActiveDocument.Tables(1).Rows(2).Height Then
          upnLabelVerticalPitch.Value = (wrdApp.PointsToCentimeters(wrdApp.ActiveDocument.Tables(1).Rows(1).Height) _
                                 + wrdApp.PointsToCentimeters(wrdApp.ActiveDocument.Tables(1).Rows(2).Height)) * 100
        Else
          upnLabelVerticalPitch.Value = wrdApp.PointsToCentimeters(wrdApp.ActiveDocument.Tables(1).Rows(1).Height) * 100
        End If
      Else
        sngTempSize1 = wrdApp.ActiveDocument.Tables(1).Rows(1).Height
        upnLabelVerticalPitch.Value = wrdApp.PointsToCentimeters(sngTempSize1) * 100
      End If
  
      ' Horizontal Pitch
      If wrdApp.ActiveDocument.Tables(1).Columns.Count > 1 Then
        If wrdApp.ActiveDocument.Tables(1).Columns(1).Width <> wrdApp.ActiveDocument.Tables(1).Columns(2).Width Then
          upnLabelHorizontalPitch.Value = (wrdApp.PointsToCentimeters(wrdApp.ActiveDocument.Tables(1).Columns(1).Width) _
                                 + wrdApp.PointsToCentimeters(wrdApp.ActiveDocument.Tables(1).Columns(2).Width)) * 100
        Else
          'upnLabelHorizontalPitch.Value = wrdapp.PointsToCentimeters  (wrdApp.ActiveDocument.Tables(1).Columns(1).Width / msngPixelsToCMs) * 100
        End If
      Else
        upnLabelHorizontalPitch.Value = wrdApp.PointsToCentimeters(wrdApp.ActiveDocument.Tables(1).Columns(1).Width) * 100
      End If
  
      ' Remember entered settings
      msngRememberVPitch = upnLabelVerticalPitch.Value
      msngRememberHPitch = upnLabelHorizontalPitch.Value
  
      
      ' Fudge to get names to work
'      wrdDoc.PageSetup.SetAsTemplateDefault
      
      
      
  
      ' Label name
'      txtName.Text = wrdApp.MailingLabel.DefaultLabelName
  
      ' Label description
'      txtDesc.Text = "Imported " & wrdApp.MailingLabel.DefaultLabelName & " from " & wrdApp.Application.Name
  
      With wrdApp.ActiveDocument
  
        ' Turn off autofit
        .Tables(1).AllowAutoFit = False
  
        ' Top Margin
        upnLabelTopMargin.Value = wrdApp.PointsToCentimeters(.PageSetup.TopMargin) * 100
  
        ' Side Margin
        upnLabelSideMargin.Value = wrdApp.PointsToCentimeters(.PageSetup.LeftMargin) * 100
  
        ' Label Height
        upnLabelHeight.Value = wrdApp.PointsToCentimeters(.Tables(1).Rows(1).Height) * 100
  
        ' Label Width
        'DoEvents
        sngTempSize1 = .Tables(1).Columns(1).Width
        upnLabelWidth.Value = wrdApp.PointsToCentimeters(sngTempSize1) * 100
  
        ' Number Down
        iCount = 0
        For iLoop = 1 To .Tables(1).Rows.Count
          If .Tables(1).Cell(iLoop, 1).Height = .Tables(1).Rows(1).Height Then
            iCount = iCount + 1
          End If
        Next iLoop
        upnLabelNumberDown.Value = iCount * 100
  
        ' Number Across
        iCount = 0
        For iLoop = 1 To .Tables(1).Columns.Count
          
          sngTempSize1 = .Tables(1).Cell(1, iLoop).Width
          sngTempSize2 = .Tables(1).Columns(1).Width
          
          If sngTempSize1 = sngTempSize2 Then
            iCount = iCount + 1
          End If
        Next iLoop
        upnLabelNumberAcross.Value = iCount * 100
  
        ' Set page size to known type / custom
        If Not ConvertLabelToKnownType(Round(wrdApp.PointsToCentimeters(.PageSetup.PageWidth), 2), Round(wrdApp.PointsToCentimeters(.PageSetup.PageHeight), 2), .PageSetup.Orientation) Then
          cboLabelPageSize.ListIndex = cboLabelPageSize.ListCount - 1
          upnLabelPageWidth.Value = (.PageSetup.PageWidth / msngPixelsToCMs) * 100
          upnLabelPageHeight.Value = (.PageSetup.PageHeight / msngPixelsToCMs) * 100
        End If
  
        On Error GoTo ImportFromWord_Error
  
        DoEvents
        wrdApp.Visible = False
  
      End With
    Else
    
      ' Envelopes
      .MainDocumentType = wdEnvelopes
      wrdDoc.Envelope.Options
      wrdApp.Visible = False
        
      ' Set page size to known type / custom
      sngWidth = Round((wrdDoc.Envelope.DefaultWidth / msngPixelsToCMs), 2)
      sngHeight = Round((wrdDoc.Envelope.DefaultHeight / msngPixelsToCMs), 2)
      If Not ConvertEnvelopeToKnownType(sngWidth, sngHeight) Then
        cboEnvelopePageSize.ListIndex = cboEnvelopePageSize.ListCount - 1
        upnEnvelopeDefinition(0).Value = sngWidth * 100
        upnEnvelopeDefinition(1).Value = sngHeight * 100
      End If
    
    End If

  End With

  ' Reset the minimum and maximum sizes
  SetLabelMaxAndMinSizes

TidyUpAndExit:

  'wrdDoc.Close False
  wrdApp.Documents.Close False
  wrdApp.Quit

  Set wrdDoc = Nothing
  Set wrdApp = Nothing

  ' JDM - 24/02/2004 - Fault 8120 - Tidy up temp data source
  Kill strFileName

  Exit Sub

' User might have clicked cancel
ImportFromWord_Error:
  bOK = False
  GoTo TidyUpAndExit

' A bug in some versions of Microsoft Word which causes the first attempt to read a column width fails,
' but which works successfully if retried
ResumeToNextLine:

  'Have a fallout just in case a REAL error trips this piece of code
  If Err.Number = 5992 Then
    iErrorCount = iErrorCount + 1
  
    If iErrorCount > 100 Then
      GoTo TidyUpAndExit
    Else
      Resume
    End If
  Else
    GoTo ImportFromWord_Error
  End If

End Sub

Private Sub cmdOK_Click()

  Dim fOK As Boolean
  
  If Me.Changed = True Then
    If ValidateDefinition = False Then
      Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
      
    If Not SaveDefinition Then
      Screen.MousePointer = vbNormal
      Exit Sub
    End If
    
    Screen.MousePointer = vbNormal
  End If
  
  Me.Hide

End Sub

Public Property Get SelectedID() As Long
  SelectedID = mlngLabelDefinitionID
End Property

Private Function SaveDefinition() As Boolean
  
  Dim rsTemp As Recordset
  Dim iCount As Integer
  Dim strSQL As String
  Dim strName As String
  Dim strDesc As String
  
  Dim strAccess As String
  
  Dim strPageHeight As String
  Dim strPageWidth As String
  
  Dim strTopMargin As String
  Dim strSideMargin As String
  Dim strVerticalPitch As String
  Dim strHorizontalPitch As String
  Dim strNumberDown As String
  Dim strNumberAcross As String
  Dim strLabelType As String
  Dim strMeasurementMethod As String
  Dim strLabelHeight As String
  Dim strLabelWidth As String
  Dim strPageTypeID As String
  
  Dim strPageOrientation As String
  Dim strFromTop As String
  Dim strFromLeft As String
  Dim strFromTopAuto As String
  Dim strFromLeftAuto As String

  On Error GoTo LocalErr

  strName = "'" & Replace(Trim(txtName.Text), "'", "''") & "'"
  strDesc = "'" & Replace(txtDesc.Text, "'", "''") & "'"

  strAccess = "'RW'"
  If optReadOnly = True Then
    strAccess = "'RO'"
  End If
 
  strTopMargin = Str(upnLabelTopMargin.Value / 100)
  strSideMargin = Str(upnLabelSideMargin.Value / 100)
  strVerticalPitch = Str(upnLabelVerticalPitch.Value / 100)
  strHorizontalPitch = Str(upnLabelHorizontalPitch.Value / 100)
  strLabelHeight = Str(upnLabelHeight.Value / 100)
  strLabelWidth = Str(upnLabelWidth.Value / 100)
  strNumberDown = Str(upnLabelNumberDown.Value / 100)
  strNumberAcross = Str(upnLabelNumberAcross.Value / 100)
  strLabelType = "'" & "" & "'"
  
  For iCount = 0 To optMeasurement.Count - 1
    If optMeasurement(iCount).Value = True Then
      strMeasurementMethod = "'" & Str(iCount) & "'"
    End If
  Next iCount
  
  If mbIsEnvelope Then
    ' Landscape is a way of turning the default behaviour of the envelope
    strPageHeight = Str(upnEnvelopeDefinition(1).Value / 100)
    strPageWidth = Str(upnEnvelopeDefinition(0).Value / 100)
    strPageTypeID = mavPageEnvelopeDimensions(cboEnvelopePageSize.ListIndex, 4)
    strPageOrientation = IIf(chkEnvelopeAlign.Value = Checked, "1", "0")
    strFromTop = Str(upnEnvelopeDefinition(2).Value / 100)
    strFromLeft = Str(upnEnvelopeDefinition(3).Value / 100)
    strFromTopAuto = IIf(chkEnvelopeAutoTop.Value = Checked, "1", "0")
    strFromLeftAuto = IIf(chkEnvelopeAutoLeft.Value = Checked, "1", "0")
  
  Else
    ' Always store the page dimensions in portrait mode
    strPageWidth = Str(IIf(chkLabelPageLandscape.Value = Checked, upnLabelPageHeight.Value, upnLabelPageWidth) / 100)
    strPageHeight = Str(IIf(chkLabelPageLandscape.Value = Checked, upnLabelPageWidth.Value, upnLabelPageHeight) / 100)
    strPageTypeID = mavPageLabelDimensions(cboLabelPageSize.ListIndex, 4)
    strPageOrientation = IIf(chkLabelPageLandscape.Value = Checked, "1", "0")
    strFromTop = "0"
    strFromLeft = "0"
    strFromTopAuto = "0"
    strFromLeftAuto = "0"
  End If

  If mlngLabelDefinitionID > 0 Then
    strSQL = "UPDATE " & mstrSQLTableDef & " SET " _
             & "Name = " & strName & ", " _
             & "Description = " & strDesc & ", " _
             & "Access = " & strAccess & ", " _
             & "ASRDefined = 0, " _
             & "NumberDown = " & strNumberDown & ", " _
             & "NumberAcross = " & strNumberAcross & ", " _
             & "LabelHeight = " & strLabelHeight & ", " _
             & "LabelWidth =" & strLabelWidth & ", " _
             & "TopMargin = " & strTopMargin & ", " _
             & "SideMargin =" & strSideMargin & ", " _
             & "Type = " & strLabelType & ", " _
             & "VerticalPitch = " & strVerticalPitch & ", " _
             & "HorizontalPitch = " & strHorizontalPitch & ", " _
             & "PageOrientation = " & strPageOrientation & ", " _
             & "PageWidth = " & strPageWidth & ", " _
             & "PageHeight = " & strPageHeight & ", " _
             & "PageTypeID = " & strPageTypeID & ", " _
             & "MeasurementMethod = " & strMeasurementMethod & ", " _
             & "IsEnvelope = " & IIf(mbIsEnvelope, "1", "0") & ", " _
             & "FromTop = " & strFromTop & ", " & "FromTopAuto = " & strFromTopAuto & ", " _
             & "FromLeft = " & strFromLeft & ", " & "FromLeftAuto = " & strFromLeftAuto
  strSQL = strSQL _
             & ", HeadingFontName = '" & mobjFormatDetails(0).FontName & "', " _
             & "HeadingFontSize = " & mobjFormatDetails(0).FontSize & ", " _
             & "HeadingFontColour = " & mobjFormatDetails(0).ForeCol & ", " _
             & "HeadingFontBold = " & IIf(mobjFormatDetails(0).Bold = True, "1", "0") & ", " _
             & "HeadingFontItalic = " & IIf(mobjFormatDetails(0).Italic, "1", "0") & ", " _
             & "HeadingFontUnderline = " & IIf(mobjFormatDetails(0).Underline, "1", "0") & ", " _
             & "StandardFontName = '" & mobjFormatDetails(1).FontName & "', " _
             & "StandardFontSize = " & mobjFormatDetails(1).FontSize & ", " _
             & "StandardFontColour = " & mobjFormatDetails(1).ForeCol & ", " _
             & "StandardFontBold = " & IIf(mobjFormatDetails(1).Bold = True, "1", "0") & ", " _
             & "StandardFontItalic = " & IIf(mobjFormatDetails(1).Italic, "1", "0") & ", " _
             & "StandardFontUnderline = " & IIf(mobjFormatDetails(1).Underline, "1", "0") _
             & " WHERE LabelTypeID = " & CStr(mlngLabelDefinitionID)
  
    gADOCon.Execute strSQL, , adCmdText

    Call UtilUpdateLastSaved(utlLabelType, mlngLabelDefinitionID)

  Else
    strSQL = "INSERT " & mstrSQLTableDef & " (" _
              & " Name, Description," _
              & " UserName, Access, ASRDefined," _
              & " NumberDown, NumberAcross," _
              & " LabelHeight, LabelWidth," _
              & " TopMargin, SideMargin," _
              & " IsEnvelope, MeasurementMethod, Type," _
              & " VerticalPitch, HorizontalPitch," _
              & " FromTop, FromTopAuto, FromLeft, FromLeftAuto, " _
              & " PageOrientation, PageWidth, PageHeight, PageTypeID," _
              & " HeadingFontName, HeadingFontSize, HeadingFontColour, " _
              & " HeadingFontBold, HeadingFontItalic, HeadingFontUnderline," _
              & " StandardFontName, StandardFontSize, StandardFontColour," _
              & " StandardFontBold, StandardFontItalic, StandardFontUnderline)"

    strSQL = strSQL & " VALUES( " _
              & strName & ", " & strDesc _
              & ", '" & datGeneral.UserNameForSQL & "', " & strAccess & ",0" _
              & ", " & strNumberDown & ", " & strNumberAcross _
              & ", " & strLabelHeight & ", " & strLabelWidth _
              & ", " & strTopMargin & ", " & strSideMargin _
              & ", " & IIf(mbIsEnvelope, "1", "0") & ", " & strMeasurementMethod & ", " & strLabelType _
              & ", " & strVerticalPitch & ", " & strHorizontalPitch _
              & ", " & strFromTop & ", " & strFromTopAuto & ", " & strFromLeft & ", " & strFromLeftAuto _
              & ", " & strPageOrientation & ", " & strPageWidth & ", " & strPageHeight & ", " & strPageTypeID _
              & ", '" & mobjFormatDetails(0).FontName & "'" _
              & ", " & mobjFormatDetails(0).FontSize _
              & ", " & mobjFormatDetails(0).ForeCol _
              & ", " & IIf(mobjFormatDetails(0).Bold = True, "1", "0") _
              & ", " & IIf(mobjFormatDetails(0).Italic, "1", "0") _
              & ", " & IIf(mobjFormatDetails(0).Underline, "1", "0") _
              & ", '" & mobjFormatDetails(1).FontName & "'" _
              & ", " & mobjFormatDetails(1).FontSize _
              & ", " & mobjFormatDetails(1).ForeCol _
              & ", " & IIf(mobjFormatDetails(1).Bold = True, "1", "0") _
              & ", " & IIf(mobjFormatDetails(1).Italic, "1", "0") _
              & ", " & IIf(mobjFormatDetails(1).Underline, "1", "0") & ")"

    ' Insert the new label/envelope type
    mlngLabelDefinitionID = InsertLabelType(strSQL)
    Call UtilCreated(utlLabelType, mlngLabelDefinitionID)
  End If

  Me.Changed = False
  SaveDefinition = True
Exit Function

LocalErr:
  'ErrorMsgbox "Error saving Label Type definition"
  SaveDefinition = False

End Function

Private Function ValidateDefinition() As Boolean
  
  ValidateDefinition = True
  
  On Error GoTo ValidateDefinition_ERROR
  
  Dim pblnContinueSave As Boolean
  Dim pblnSaveAsNew As Boolean
  Dim iCount As Integer
  Dim bDimensionsNotComplete As Boolean
  Dim sngLabelWidth As Single
  Dim sngLabelHeight As Single
  Dim strValidationMessage As String
  Dim bOK As Boolean
  
  bOK = True
  
  ' Check a name has been entered
  If Trim(txtName.Text) = "" Then
    MsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    txtName.SetFocus
    ValidateDefinition = False
    Exit Function
  End If
  
  'Check if this definition has been changed by another user
  Call UtilityDefAmended("ASRSysLabelTypes", "LabelTypeID", mlngLabelDefinitionID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Exit Function
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    optReadWrite.Enabled = True
    optReadOnly.Enabled = True
    mlngLabelDefinitionID = 0
  End If
  
  ' Check the name is unique
  If Not CheckUniqueName(Trim(txtName.Text), mlngLabelDefinitionID) Then
    MsgBox "An Envelope and Label Template called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    ValidateDefinition = False
    Exit Function
  End If
  
  ' Validate that all of the lable sizes are set
  If Not mbIsEnvelope Then
    bOK = (txtLabelTopMargin.Text <> "")
    bOK = bOK And (txtLabelSideMargin.Text <> "")
    bOK = bOK And (txtLabelHorizontalPitch.Text <> "")
    bOK = bOK And (txtLabelVerticalPitch.Text <> "")
    bOK = bOK And (txtLabelHeight.Text <> "")
    bOK = bOK And (txtLabelWidth.Text <> "")
    bOK = bOK And (txtLabelNumberAcross.Text <> "")
    bOK = bOK And (txtLabelNumberDown.Text <> "")
  
    If Not bOK Then
      MsgBox "Label dimensions are not complete.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
  
    ' Validate that this label fits on the page
    sngLabelWidth = (((upnLabelNumberAcross.Value - 100) / 100) * upnLabelHorizontalPitch.Value) + upnLabelWidth.Value + upnLabelSideMargin.Value
    If sngLabelWidth > upnLabelPageWidth.Value Then
      MsgBox "Labels are too wide to fit on the page.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
    
    sngLabelHeight = (((upnLabelNumberDown.Value - 100) / 100) * upnLabelVerticalPitch.Value) + upnLabelHeight.Value + upnLabelTopMargin.Value
    If sngLabelHeight > upnLabelPageHeight.Value Then
      MsgBox "Labels are too tall to fit on the page.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
 
    ' Check that the number of labels > 0
    If upnLabelNumberAcross.Value = 0 Then
      MsgBox "Invalid number of labels across.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
     
    If upnLabelNumberDown.Value = 0 Then
      MsgBox "Invalid number of labels down.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If

    ' Check that labels dimensions are greater than 0
    If upnLabelHeight.Value = 0 Then
      MsgBox "Invalid label height.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
     
    ' Check that labels dimensions are greater than 0
    If upnLabelWidth.Value = 0 Then
      MsgBox "Invalid label width.", vbExclamation, Me.Caption
      ValidateDefinition = False
      Exit Function
    End If
 
  End If

  Exit Function
  
ValidateDefinition_ERROR:
  
  MsgBox "Error whilst validating " & IIf(mbIsEnvelope, "an envelope", "a label") & " definition." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Export"
  ValidateDefinition = False
  
End Function

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)

  If Not mbIsLoading Then
    cmdOK.Enabled = blnChanged
  End If

End Property

Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional lLabelDefinitionID As Long, Optional bPrint As Boolean) As Boolean

  Dim sngWidth As Single
  Dim sngHeight As Single
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim fOK As Boolean
  Dim bFound As Boolean
  Dim strLoadingWarning As String
  Dim sAccess As String

  On Error GoTo LocalErr

  strLoadingWarning = ""

  ' Only allow import button if Word is XP or greater
  cmdImportFromWord.Enabled = IIf(GetWordVersion >= 10, True, False)

  Set datData = New HRProDataMgr.clsDataAccess

  Screen.MousePointer = vbHourglass

  mbIsLoading = True
  fOK = True
   
  ' Load the different page types
  LoadPageSizes
  cboLabelPageSize.ListIndex = 0
  cboEnvelopePageSize.ListIndex = 0

  ' Format Tab
  PopulateFormatTab
  Set mobjFormatDetails(0) = New HRProDataMgr.clsOutputStyle
  Set mobjFormatDetails(1) = New HRProDataMgr.clsOutputStyle

  ' Remember pitch settings
  msngRememberVPitch = msngMinLabelHeight * 100
  msngRememberHPitch = msngMinLabelWidth * 100

  sAccess = GetUserSetting("utils&reports", "dfltaccess labeldefinition", ACCESS_READWRITE)
  
  If bNew Then
    
    ' Initialise default values
    NewDefinition
    
    ' this MUST be before optallrecordsclick !!! RH 14/06/00
    mblnDefinitionCreator = True
    
    'Set ID to 0 to indicate new record
    mlngLabelDefinitionID = 0
    txtUserName.Text = gsUserName
    mblnDefinitionCreator = True
    optLabelEnvelope(0).Value = True

    ' Default sizes to 1cm
    upnLabelTopMargin.Value = 100
    upnLabelSideMargin.Value = 100
    upnLabelHeight.Value = 100
    upnLabelWidth.Value = 100

    Select Case sAccess
      Case ACCESS_READWRITE
        optReadWrite.Value = True
      Case Else
        optReadOnly.Value = True
    End Select

    Me.Changed = False

  Else
    mlngLabelDefinitionID = lLabelDefinitionID
    mblnFromCopy = bCopy
  
    ' We need to know if we are going to PRINT the definition.
    FormPrint = bPrint
    
    Call RetreiveDefinition

    If fOK And Not Me.Cancelled Then
      
      If mblnFromCopy Then
        mlngLabelDefinitionID = 0
        
        Select Case sAccess
          Case ACCESS_READWRITE
            optReadWrite.Value = True
          Case Else
            optReadOnly.Value = True
        End Select
        
        Me.Changed = True
      Else
        Me.Changed = (Not mblnReadOnly)
      End If
      
    End If

  End If

  ' Auto format the definition controls
  If Not mbIsEnvelope Then
    txtLabelTopMargin.Text = Format(Str(upnLabelTopMargin.Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtLabelSideMargin.Text = Format(Str(upnLabelSideMargin.Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtLabelVerticalPitch.Text = Format(Str(upnLabelVerticalPitch.Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtLabelHorizontalPitch.Text = Format(Str(upnLabelHorizontalPitch.Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtLabelWidth.Text = Format(Str(upnLabelWidth.Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtLabelHeight.Text = Format(Str(upnLabelHeight.Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtLabelNumberAcross.Text = Trim(Str(upnLabelNumberAcross.Value / 100))
    txtLabelNumberDown.Text = Trim(Str(upnLabelNumberDown.Value / 100))
        
'    ' Set page size to known type / custom
'    sngWidth = Round((upnLabelPageWidth.Value / msngConversion), 2)
'    sngHeight = Round((upnLabelPageHeight.Value / msngConversion), 2)
'    If Not ConvertLabelToKnownType(sngWidth, sngHeight, chkLabelPageLandscape.Value) Then
'      cboLabelPageSize.ListIndex = cboLabelPageSize.ListCount - 1
'      upnLabelPageWidth.Value = sngWidth * 100
'      upnLabelPageHeight.Value = sngHeight * 100
'    End If
    
  Else
    txtEnvelopeDimension(0).Text = Format(Str(upnEnvelopeDefinition(0).Value / msngConversion), "0.00 " & mstrMeasurementText)
    txtEnvelopeDimension(1).Text = Format(Str(upnEnvelopeDefinition(1).Value / msngConversion), "0.00 " & mstrMeasurementText)
  End If

  ' Set max / in values
  upnEnvelopeDefinition(0).Min = msngMinEnvelopePageWidth
  upnEnvelopeDefinition(1).Min = msngMinEnvelopePageHeight
  
  upnEnvelopeDefinition(2).Min = msngEnvelopeFromTopMin * 100
  upnEnvelopeDefinition(3).Min = msngEnvelopeFromLeftMin * 100

  ' Formatting controls
  For iCount = 0 To 1
    
    bFound = False
    For iCount2 = 1 To cboFontName(iCount).ListCount
      If mobjFormatDetails(iCount).FontName = cboFontName(iCount).List(iCount2) Then
        bFound = True
        Exit For
      End If
    Next iCount2
    
    If Not bFound Then
      strLoadingWarning = strLoadingWarning & IIf(Len(strLoadingWarning) > 0, ", ", "") & mobjFormatDetails(iCount).FontName
      mobjFormatDetails(iCount).FontName = "Verdana"
    End If
        
    cboFontName(iCount).Text = mobjFormatDetails(iCount).FontName
    cboFontSize(iCount).Text = mobjFormatDetails(iCount).FontSize
    cboFontColour(iCount).Text = mobjFormatDetails(iCount).ForeCol
    chkFontBold(iCount).Value = IIf(mobjFormatDetails(iCount).Bold = True, vbChecked, vbUnchecked)
    chkFontItalic(iCount).Value = IIf(mobjFormatDetails(iCount).Italic = True, vbChecked, vbUnchecked)
    chkFontUnderLine(iCount).Value = IIf(mobjFormatDetails(iCount).Underline = True, vbChecked, vbUnchecked)
  
    For iCount2 = 1 To cboFontColour(iCount).ComboItems.Count
      If cboFontColour(iCount).ComboItems.Item(iCount2).Tag = mobjFormatDetails(iCount).ForeCol Then
        cboFontColour(iCount).SelectedItem = cboFontColour(iCount).ComboItems.Item(iCount2)
        Exit For
      End If
    Next iCount2
  Next iCount

  If Len(strLoadingWarning) > 0 Then
    MsgBox strLoadingWarning & " fonts are not installed on this machine." _
        & vbCrLf & "Font(s) have been set to Verdana." _
        & vbCrLf & vbCrLf & "Please contact your system administrator", vbOKOnly + vbExclamation
    Me.Changed = True
  End If

  UpdatePreview
  EnableDisableTabControls
  
  Screen.MousePointer = vbNormal
  mbIsLoading = False
  
  ' Fault 6184 - Not enabling ok button
  If mblnFromCopy Then
    Me.Changed = True
  End If
  
  
  Initialise = fOK

  Screen.MousePointer = vbDefault

Exit Function

LocalErr:
  MsgBox "Error with Label type definition"

End Function

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Private Sub Form_Load()

  On Error GoTo InvalidPageSize

  ' Default measurement style
  mbIsLoading = True
  mstrMeasurementText = mstrCentimetreText
  msngConversion = msngCentimetreConversion

  ' Set the maxmimum and minimum sizes for labels
  SetLabelMaxAndMinSizes

  ' Maximum values (Envelopes)
  upnEnvelopeDefinition(0).Max = msngMaxEnvelopePageWidth * 100
  upnEnvelopeDefinition(1).Max = msngMaxEnvelopePageHeight * 100

  ' Minimum values (Labels)
'  upnEnvelopeDefinition(0).Min = msngMinEnvelopePageWidth * 100
'  upnEnvelopeDefinition(1).Min = msngMinEnvelopePageHeight * 100


  Exit Sub

InvalidPageSize:

  ' Handle if previously defined page size that is out of Word range (should only ever effect development builds (i.e QA)
  If Not mbPreviouslyWarned Then
    MsgBox "This label type is outside the range of Microsoft Word. Please resave.", vbExclamation, Me.Caption
    mbPreviouslyWarned = True
    Resume Next
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer
    
    If Changed = True And Not FormPrint Then
      
      pintAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, Me.Caption)
        
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
'
'  'NHRD27102004 Fault 8305
'  Dim strSQL As String
'  Dim strMBText As String
'  Dim intMBButtons As Integer
'  Dim strMBTitle As String
'  Dim intMBResponse As Integer
'
'  If Me.Changed And Not mblnReadOnly Then
'
'    strMBText = "You have changed the current definition. Save changes ?"
'    intMBButtons = vbQuestion + vbYesNoCancel + vbDefaultButton1
'    strMBTitle = Me.Caption
'    intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)
'
'    Select Case intMBResponse
'    Case vbYes
'      Call cmdOK_Click
'      Exit Sub
'    Case vbCancel
'      Exit Sub
'    End Select
'  End If
'
'
'  Me.Hide
'  Screen.MousePointer = vbDefault
'End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optMeasurement_Click(Index As Integer)

  Dim lngOldConversion As Long
  Dim iCount As Integer
  lngOldConversion = msngConversion

  ' Set the text for this template definition
  Select Case Index
    Case 0
      mstrMeasurementText = mstrCentimetreText
      msngConversion = msngCentimetreConversion
    Case 1
      mstrMeasurementText = mstrMillimetreText
      msngConversion = msngMillimetreConversion
    Case 2
      mstrMeasurementText = mstrInchText
      msngConversion = msngInchConversion
    Case 3
      mstrMeasurementText = mstrPixelText
      msngConversion = msngPixelConversion
  End Select

  ' Force definitions to refresh themselves
  ' Labels
  upnLabelTopMargin.Value = upnLabelTopMargin.Value
  upnLabelSideMargin.Value = upnLabelSideMargin.Value
  upnLabelHorizontalPitch.Value = upnLabelHorizontalPitch.Value
  upnLabelVerticalPitch.Value = upnLabelVerticalPitch.Value
  upnLabelWidth.Value = upnLabelWidth.Value
  upnLabelHeight.Value = upnLabelHeight.Value
  upnLabelPageWidth.Value = upnLabelPageWidth.Value
  upnLabelPageHeight.Value = upnLabelPageHeight.Value
  
  ' Envelopes
  For iCount = 0 To upnEnvelopeDefinition.UBound
    upnEnvelopeDefinition(iCount).Value = upnEnvelopeDefinition(iCount).Value
  Next iCount

End Sub

Private Sub optLabelEnvelope_Click(Index As Integer)

  mbIsEnvelope = (optLabelEnvelope(1).Value = True)
  cmdImportFromWord.Enabled = Not mbIsEnvelope And (GetWordVersion >= 10)

  fraLabel.Visible = (optLabelEnvelope(0).Value = True)
  fraEnvelope.Visible = Not fraLabel.Visible

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

Private Sub optReadOnly_Click()

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

Private Sub optReadWrite_Click()

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  EnableDisableTabControls
End Sub

Private Sub txtDesc_Change()

  If Not mbIsLoading Then
    Me.Changed = True
  End If

End Sub

Private Sub txtDesc_GotFocus()
  cmdOK.Default = False
End Sub

Private Sub txtDesc_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub txtEnvelopeDimension_Change(Index As Integer)
  Me.Changed = True
End Sub

Private Sub txtEnvelopeDimension_KeyPress(Index As Integer, KeyAscii As Integer)

  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If

End Sub

Private Sub txtEnvelopeDimension_LostFocus(Index As Integer)

  Dim strSize As String
  Dim intSizeEnd As Integer
  Dim lngStartValue As Long

  lngStartValue = upnEnvelopeDefinition(Index).Value
  strSize = txtEnvelopeDimension(Index).Text
  intSizeEnd = InStr(1, txtEnvelopeDimension(Index).Text, mstrMeasurementText, vbTextCompare)

  If intSizeEnd > 0 Then
    strSize = Trim(Left(txtEnvelopeDimension(Index).Text, intSizeEnd - 1))
  End If

  If IsNumeric(strSize) Then
    If strSize <= upnEnvelopeDefinition(Index).Max / msngConversion Then
      malngEnvelopeDimensions(Index) = Val(txtEnvelopeDimension(Index).Text) * msngConversion
      
      If strSize > upnEnvelopeDefinition(Index).Min / msngConversion Then
        upnEnvelopeDefinition(Index).Value = malngEnvelopeDimensions(Index)
      Else
        upnEnvelopeDefinition(Index).Value = upnEnvelopeDefinition(Index).Min
      End If
    Else
      malngEnvelopeDimensions(Index) = upnEnvelopeDefinition(Index).Max
      upnEnvelopeDefinition(Index).Value = malngEnvelopeDimensions(Index)
    End If
  End If

  ' Has value changed
  If upnEnvelopeDefinition(Index).Value <> lngStartValue Then
    Me.Changed = True
  End If

End Sub

Private Sub txtLabelHeight_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelHorizontalPitch_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelNumberAcross_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelNumberDown_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelPageHeight_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelPageWidth_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelSideMargin_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelTopMargin_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelTopMargin_KeyPress(KeyAscii As Integer)
  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub txtLabelVerticalPitch_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtLabelWidth_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub txtName_Change()
  If Not mbIsLoading Then
    Me.Changed = True
  End If
End Sub

Private Sub upnEnvelopeDefinition_Change(Index As Integer)

  txtEnvelopeDimension(Index).Text = Format(Str(upnEnvelopeDefinition(Index).Value / msngConversion), "0.00 " & mstrMeasurementText)

End Sub

Private Function InsertLabelType(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo InsertLabelType_ERROR

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
    pmADO.Value = mstrSQLTableDef
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "LabelTypeID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertLabelType = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertLabelType = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
InsertLabelType_ERROR:
  
  fSavedOK = False
  Resume Next
  
End Function
Private Function GetDefinition() As Recordset

  Dim strSQL As String

  strSQL = "SELECT " & _
           mstrSQLTableDef & ".*, " & _
           "CONVERT(integer," & mstrSQLTableDef & ".TimeStamp) AS intTimeStamp " & _
           "FROM " & mstrSQLTableDef & " " & _
           "WHERE " & mstrSQLTableDef & ".LabelTypeID = " & CStr(mlngLabelDefinitionID)
  Set GetDefinition = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

End Function
Private Sub RetreiveDefinition()

  Dim rsTemp As Recordset
  Dim sMessage As String
  Dim iCount As Integer
  Dim bPreviouslyWarned As Boolean
  Dim fOK As Boolean

  On Error GoTo LocalErr

  bPreviouslyWarned = False
  mbIsLoading = True

  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Label Definition"
    fOK = False
    Exit Sub
  End If

  txtDesc.Text = IIf(rsTemp!Description <> vbNullString, rsTemp!Description, vbNullString)
  
  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If

  mblnReadOnly = Not datGeneral.SystemPermission("LABELDEFINITION", "EDIT")
  
  ' Is this an envelope definition
  optLabelEnvelope(0).Value = IIf(IsNull(rsTemp!IsEnvelope), True, Not rsTemp!IsEnvelope)
  optLabelEnvelope(1).Value = IIf(IsNull(rsTemp!IsEnvelope), False, rsTemp!IsEnvelope)
  
  ' Set the access type
  Select Case rsTemp!Access
  Case "RW"
    optReadWrite = True
  Case "RO"
    optReadOnly = True
    'MH20040127 Fault 7889
    'mblnReadOnly = (mblnReadOnly Or Not mblnDefinitionCreator)
    mblnReadOnly = ((mblnReadOnly Or Not mblnDefinitionCreator) And Not gfCurrentUserIsSysSecMgr)
  End Select
  
  CheckAccess

  If mblnReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If
  
  mlngTimeStamp = rsTemp!intTimestamp

  ' Load the values
  If mbIsEnvelope Then
    
    SetCombo cboEnvelopePageSize, rsTemp.Fields("PageTypeID").Value
    upnEnvelopeDefinition(0).Value = rsTemp.Fields("PageWidth").Value * 100
    upnEnvelopeDefinition(1).Value = rsTemp.Fields("PageHeight").Value * 100
    chkEnvelopeAlign.Value = IIf(rsTemp.Fields("PageOrientation").Value, 1, 0)
  
    upnEnvelopeDefinition(2).Value = rsTemp.Fields("FromTop").Value * 100
    upnEnvelopeDefinition(3).Value = rsTemp.Fields("FromLeft").Value * 100
    chkEnvelopeAutoTop.Value = IIf(rsTemp.Fields("FromTopAuto").Value = True, vbChecked, vbUnchecked)
    chkEnvelopeAutoLeft.Value = IIf(rsTemp.Fields("FromLeftAuto").Value = True, vbChecked, vbUnchecked)

  Else
    upnLabelNumberAcross.Value = IIf(rsTemp.Fields("NumberAcross").Value = 0, 100, rsTemp.Fields("NumberAcross").Value * 100)
    upnLabelNumberDown.Value = IIf(rsTemp.Fields("NumberDown").Value = 0, 100, rsTemp.Fields("NumberDown").Value * 100)
    upnLabelTopMargin.Value = rsTemp.Fields("TopMargin").Value * 100
    upnLabelSideMargin.Value = rsTemp.Fields("SideMargin").Value * 100
    upnLabelHeight.Value = rsTemp.Fields("LabelHeight").Value * 100
    upnLabelWidth.Value = rsTemp.Fields("LabelWidth").Value * 100
    upnLabelVerticalPitch.Value = rsTemp.Fields("VerticalPitch").Value * 100
    upnLabelHorizontalPitch.Value = rsTemp.Fields("HorizontalPitch").Value * 100
    upnLabelPageWidth.Value = rsTemp.Fields("PageWidth").Value * 100
    upnLabelPageHeight.Value = rsTemp.Fields("PageHeight").Value * 100
    
    ' Remember pitch settings
    msngRememberVPitch = upnLabelVerticalPitch.Value
    msngRememberHPitch = upnLabelHorizontalPitch.Value
    
    ' Set page size
    SetCombo cboLabelPageSize, rsTemp.Fields("PageTypeID").Value
    chkLabelPageLandscape.Value = IIf(rsTemp.Fields("PageOrientation").Value, 1, 0)
      
  End If
    
  ' Set the measurement method
  For iCount = 0 To optMeasurement.UBound
    If Not IsNull(rsTemp!MeasurementMethod) Then
      optMeasurement(iCount).Value = (rsTemp!MeasurementMethod = iCount)
    End If
  Next iCount
    
  ' Heading Format
  mobjFormatDetails(0).FontName = rsTemp.Fields("HeadingFontName")
  mobjFormatDetails(0).FontSize = rsTemp.Fields("HeadingFontSize")
  mobjFormatDetails(0).ForeCol = rsTemp.Fields("HeadingFontColour")
  mobjFormatDetails(0).Bold = rsTemp.Fields("HeadingFontBold")
  mobjFormatDetails(0).Italic = rsTemp.Fields("HeadingFontItalic")
  mobjFormatDetails(0).Underline = rsTemp.Fields("HeadingFontUnderline")
        
  ' Standard Format
  mobjFormatDetails(1).FontName = rsTemp.Fields("StandardFontName")
  mobjFormatDetails(1).FontSize = rsTemp.Fields("StandardFontSize")
  mobjFormatDetails(1).ForeCol = rsTemp.Fields("StandardFontColour")
  mobjFormatDetails(1).Bold = rsTemp.Fields("StandardFontBold")
  mobjFormatDetails(1).Italic = rsTemp.Fields("StandardFontItalic")
  mobjFormatDetails(1).Underline = rsTemp.Fields("StandardFontUnderline")

Exit Sub

LocalErr:
  
  ' Is the dimensions of the label/envelope outside of the required parameters
  If Err.Number = 380 Then
    If Not bPreviouslyWarned Then
      MsgBox "This template definition is not supported by Microsoft Word. Please redefine.", vbExclamation + vbOKOnly, "Label Definition"
      bPreviouslyWarned = True
    End If
    Resume Next
  End If
  
  MsgBox "Error retrieving type definition"
    
End Sub

Private Function CheckAccess(Optional strType As String) As Boolean

  'strType should be 'picklist', 'filter' or 'calculation' or null

  Dim strMessage As String
  Dim blnAccessEnabled As String
  Dim blnHiddenComponent As Boolean

  CheckAccess = True

  blnHiddenComponent = (mblnHiddenCalculation And strType = "calculation") Or _
                       (mblnHiddenPicklistOrFilter And strType = "picklist") Or _
                       (mblnHiddenPicklistOrFilter And strType = "filter")
  
  strMessage = vbNullString
  
  'Check if status has changed then alert the user
  If strType <> vbNullString Then
  
    'Somebody else selects hidden component
    If Not mblnDefinitionCreator Then
      If blnHiddenComponent Then
        MsgBox "Unable to select this " & strType & _
               " as it is a hidden " & strType & _
               " and you are not the owner of this definition", vbExclamation
        CheckAccess = False
        
        'TM20020715 Fault 4034
'        If Not IsMissing(strType) Then
'          If strType = "picklist" Then
'            txtPicklist.Text = "<None>"
'            txtPicklist.Tag = 0
'          ElseIf strType = "filter" Then
'            txtFilter.Text = "<None>"
'            txtFilter.Tag = 0
'          End If
'        End If
        
        Exit Function
      End If
    End If
  End If
  
  
  'Update controls
  blnAccessEnabled = (mblnDefinitionCreator And mblnHiddenCalculation = False And mblnHiddenPicklistOrFilter = False)
  
  optReadWrite.Enabled = blnAccessEnabled
  optReadOnly.Enabled = blnAccessEnabled

  'Alert user of new status
  If strMessage <> vbNullString Then
    MsgBox strMessage, vbInformation, "Mail Merge"
  End If

End Function

Private Sub LoadPageSizes()

  ' Loads the Base combo with all tables (even lookups)
  
  Dim sSQL As String
  Dim rsPages As New Recordset
  Dim iCount As Integer

  ' Labels
  sSQL = "Select * From ASRSysPageSizes"
  sSQL = sSQL & "  WHERE IsEnvelope <> 1 ORDER BY DisplayOrder, Name"
  Set rsPages = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)

  If Not rsPages.EOF And Not rsPages.BOF Then

    rsPages.MoveFirst

    ReDim mavPageLabelDimensions(rsPages.RecordCount, 4)

    With cboLabelPageSize
    
      .Clear
      Do While Not rsPages.EOF
        .AddItem rsPages!Name
        .ItemData(.NewIndex) = rsPages!PageSizeID
        
        mavPageLabelDimensions(.NewIndex, 1) = Format(rsPages!Width, "##0.00")
        mavPageLabelDimensions(.NewIndex, 2) = Format(rsPages!Height, "##0.00")
        mavPageLabelDimensions(.NewIndex, 3) = Format(rsPages!WordTemplateID, "###")
        mavPageLabelDimensions(.NewIndex, 4) = rsPages!PageSizeID
        rsPages.MoveNext

      Loop
    End With

    rsPages.Close

  End If

  ' Envelopes
  sSQL = "Select * From ASRSysPageSizes"
  sSQL = sSQL & " WHERE IsEnvelope = 1 ORDER BY DisplayOrder, Name"
  Set rsPages = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)

  If Not rsPages.EOF And Not rsPages.BOF Then

    rsPages.MoveFirst

    ReDim mavPageEnvelopeDimensions(rsPages.RecordCount, 4)

    With cboEnvelopePageSize
    
      .Clear
      Do While Not rsPages.EOF
        .AddItem rsPages!Name
        .ItemData(.NewIndex) = rsPages!PageSizeID
        
        mavPageEnvelopeDimensions(.NewIndex, 1) = Format(rsPages!Width, "##0.00")
        mavPageEnvelopeDimensions(.NewIndex, 2) = Format(rsPages!Height, "##0.00")
        mavPageEnvelopeDimensions(.NewIndex, 3) = Format(rsPages!WordTemplateID, "###")
        mavPageEnvelopeDimensions(.NewIndex, 4) = rsPages!PageSizeID
        rsPages.MoveNext

      Loop
    End With

    rsPages.Close

  End If

  Set rsPages = Nothing

End Sub

Private Function CheckUniqueName(sName As String, lngCurrentID As Long) As Boolean
  ' Is there already a definition with the same name (that isnt the
  ' definition we are editing ?)
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSysLabelTypes " & _
         "WHERE UPPER(Name) = '" & UCase(Replace(sName, "'", "''")) & "' " & _
         "AND LabelTypeID <> " & lngCurrentID
  
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckUniqueName = True
  Else
    CheckUniqueName = False
  End If
  
  Set rsTemp = Nothing

End Function

Public Sub PrintDefinition(plngLabelTypeID)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim rsPages As Recordset
  Dim strSQL As String
  Dim bIsEnvelope As Boolean
  Dim strTemplateTypeName As String
  Dim strAlignment As String
  Dim strFontStyle As String
  Dim sngWidth As Single
  Dim sngHeight As Single
  
  Set datData = New HRProDataMgr.clsDataAccess
  
  mlngLabelDefinitionID = plngLabelTypeID
  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation, "Template Definition"
    Exit Sub
  End If
  
  bIsEnvelope = IIf(IsNull(rsTemp.Fields("IsEnvelope")), False, rsTemp.Fields("IsEnvelope"))
  strTemplateTypeName = IIf(bIsEnvelope, "Envelope", "Label")
  
  ' Measurement type
  If Not IsNull(rsTemp.Fields("MeasurementMethod")) Then
    Select Case rsTemp.Fields("MeasurementMethod")
      Case 0
        mstrMeasurementText = mstrCentimetreText
        msngConversion = msngCentimetreConversion
      Case 1
        mstrMeasurementText = mstrMillimetreText
        msngConversion = msngMillimetreConversion
      Case 2
        mstrMeasurementText = mstrInchText
        msngConversion = msngInchConversion
      Case 3
        mstrMeasurementText = mstrPixelText
        msngConversion = msngPixelConversion
    End Select
  Else
    mstrMeasurementText = mstrCentimetreText
    msngConversion = msngCentimetreConversion
  End If
  
  If Not IsNull(rsTemp.Fields("PageOrientation").Value) Then
    If bIsEnvelope Then
      strAlignment = IIf(rsTemp.Fields("PageOrientation").Value = True, "Longest Edge", "Shortest Edge")
    Else
      strAlignment = IIf(rsTemp.Fields("PageOrientation").Value = True, "Landscape", "Portrait")
    End If
  Else
    strAlignment = IIf(bIsEnvelope, "Portrait", "Shortest Edge")
  End If
  
  ' get the page size type
  strSQL = "Select * From ASRSysPageSizes"
  strSQL = strSQL & " WHERE PageSizeID = " & rsTemp.Fields("PageTypeID").Value
  Set rsPages = datData.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)
  
  If rsPages.BOF And rsPages.EOF Then
    MsgBox "This label type has been deleted by another user.", vbExclamation, "Template Definition"
    Exit Sub
  End If
   
  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
      
        .PrintHeader strTemplateTypeName & " Template Definition: " & rsTemp!Name
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal
    
        .PrintNormal "Owner : " & rsTemp!UserName
        Select Case rsTemp!Access
          Case "RW": .PrintNormal "Access : Read / Write"
          Case "RO": .PrintNormal "Access : Read only"
        End Select
        
        .PrintNormal
        
        If bIsEnvelope Then
          
          sngHeight = IIf(strAlignment = "Portrait", rsPages!Height, rsTemp!PageHeight)
          sngWidth = IIf(strAlignment = "Portrait", rsPages!Width, rsTemp!PageWidth)
          
          .PrintTitle "Page Size"
          .PrintNormal "Type : " & rsPages!Name
          .PrintNormal "Width : " & Round(((100 / msngConversion) * sngWidth), 2) & " " & mstrMeasurementText
          .PrintNormal "Height : " & Round(((100 / msngConversion) * sngHeight), 2) & " " & mstrMeasurementText
          .PrintNormal "Orientation : " & strAlignment

          .PrintTitle "Text Position"
          .PrintNormal "From Top : " & IIf(rsTemp!FromTopAuto, "Automatic", Round(((100 / msngConversion) * rsTemp!fromtop), 2) & " " & mstrMeasurementText)
          .PrintNormal "From Left : " & IIf(rsTemp!FromLeftAuto, "Automatic", Round(((100 / msngConversion) * rsTemp!fromLeft), 2) & " " & mstrMeasurementText)

        Else
          .PrintTitle "Dimensions"
          .PrintNormal "Top Margin : " & Round(((100 / msngConversion) * rsTemp!TopMargin), 2) & " " & mstrMeasurementText
          .PrintNormal "Side Margin : " & Round(((100 / msngConversion) * rsTemp!SideMargin), 2) & " " & mstrMeasurementText
          .PrintNormal "Horizontal Pitch : " & Round(((100 / msngConversion) * rsTemp!HorizontalPitch), 2) & " " & mstrMeasurementText
          .PrintNormal "Vertical Pitch : " & Round(((100 / msngConversion) * rsTemp!VerticalPitch), 2) & " " & mstrMeasurementText
          .PrintNormal "Label Height : " & Round(((100 / msngConversion) * rsTemp!LabelHeight), 2) & " " & mstrMeasurementText
          .PrintNormal "Label Width : " & Round(((100 / msngConversion) * rsTemp!LabelWidth), 2) & " " & mstrMeasurementText
          .PrintNormal "Number Across : " & rsTemp!NumberAcross
          .PrintNormal "Number Down : " & rsTemp!NumberDown
          .PrintNormal

          .PrintTitle "Page Size"
          If rsPages!PageSizeID = 1 Then
            sngHeight = IIf(strAlignment = "Portrait", rsPages!Height, rsTemp!PageHeight)
            sngWidth = IIf(strAlignment = "Portrait", rsPages!Width, rsTemp!PageWidth)
            
            .PrintNormal "Type : Custom"
            .PrintNormal "Width : " & Round(((100 / msngConversion) * sngWidth), 2) & " " & mstrMeasurementText
            .PrintNormal "Height : " & Round(((100 / msngConversion) * sngHeight), 2) & " " & mstrMeasurementText
          Else
            sngHeight = IIf(strAlignment = "Portrait", rsPages!Height, rsPages!Width)
            sngWidth = IIf(strAlignment = "Portrait", rsPages!Width, rsPages!Height)
            
            .PrintNormal "Page Size : " & rsPages!Name
            .PrintNormal "Width : " & Round(((100 / msngConversion) * sngWidth), 2) & " " & mstrMeasurementText
            .PrintNormal "Height : " & Round(((100 / msngConversion) * sngHeight), 2) & " " & mstrMeasurementText
          End If
          .PrintNormal "Orientation : " & strAlignment
          
        End If
        '--------
  
        .PrintTitle "Format"
        
        ' Heading text
        .Font = rsTemp.Fields("HeadingFontName")
        .FontSize = rsTemp.Fields("HeadingFontSize")
        .FontColour = rsTemp.Fields("HeadingFontColour")
        .PrintNormal "Heading Text Format"
        .ResetFontToDefault
               
        strFontStyle = ""
        If rsTemp.Fields("HeadingFontBold") = True Then
          strFontStyle = ", Bold"
        End If
        
        If rsTemp.Fields("HeadingFontUnderline") = True Then
          strFontStyle = strFontStyle & ", Underline"
        End If
        
        If rsTemp.Fields("HeadingFontItalic") = True Then
          strFontStyle = strFontStyle & ", Italic"
        End If
 
        .PrintNormal "Font : " & rsTemp.Fields("HeadingFontName") & " (" & rsTemp.Fields("HeadingFontSize") & "pt" _
           & ", " & GetColourName(rsTemp.Fields("HeadingFontColour")) & " " & strFontStyle & ")"

        ' Standard text
        .PrintNormal
        .Font = rsTemp.Fields("StandardFontName")
        .FontSize = rsTemp.Fields("StandardFontSize")
        .FontColour = rsTemp.Fields("StandardFontColour")
        .PrintNormal "Standard Text Format"
        .ResetFontToDefault
        
        strFontStyle = ""
        If rsTemp.Fields("StandardFontBold") = True Then
          strFontStyle = ", Bold"
        End If
        
        If rsTemp.Fields("StandardFontUnderline") = True Then
          strFontStyle = strFontStyle & ", Underline"
        End If
        
        If rsTemp.Fields("StandardFontItalic") = True Then
          strFontStyle = strFontStyle & ", Italic"
        End If
               
        .PrintNormal "Font : " & rsTemp.Fields("StandardFontName") _
            & " (" & rsTemp.Fields("StandardFontSize") & "pt" _
            & ", " & GetColourName(rsTemp.Fields("StandardFontColour")) & " " & strFontStyle & ")"
    
        .PrintEnd
        
      End If

  
    End With
    
  End If
  
  Set datData = Nothing

Exit Sub

LocalErr:
  MsgBox "Printing Microsoft Word Template Definition Failed"

End Sub


Private Function GetWordVersion() As Integer
  
  Dim wrdApp As Object
  
  On Error GoTo LocalErr

  Set wrdApp = CreateObject("Word.Application")

  GetWordVersion = Val(wrdApp.Version)
  wrdApp.Quit
  Set wrdApp = Nothing
Exit Function

LocalErr:
  GetWordVersion = 0

End Function

Private Sub SetCombo(objControl As ComboBox, lItemData As Long)

  Dim lCount As Long
  
  For lCount = 0 To objControl.ListCount
    If objControl.ItemData(lCount) = lItemData Then
      objControl.ListIndex = lCount
      Exit Sub
    End If
  Next lCount

End Sub

' Sets the page combo to the correct word template
Private Function ConvertEnvelopeToKnownType(plngWidth As Single, plngHeight As Single) As Boolean

Dim iCount As Integer
Dim bFound As Boolean

bFound = True

'For iCount = LBound(mavPageEnvelopeDimensions, 1) To UBound(mavPageEnvelopeDimensions, 1)
'  If plngWidth = Val(mavPageEnvelopeDimensions(iCount, 1)) And plngHeight = Val(mavPageEnvelopeDimensions(iCount, 2)) Then
'    cboEnvelopePageSize.ListIndex = iCount
'    bFound = True
'  End If
'Next iCount

ConvertEnvelopeToKnownType = bFound

End Function

' Sets the page combo to the correct word template
Private Function ConvertLabelToKnownType(plngWidth As Single, plngHeight As Single, pbLandscape As Boolean) As Boolean

Dim iCount As Integer
Dim bFound As Boolean

bFound = True

For iCount = LBound(mavPageLabelDimensions, 1) To UBound(mavPageLabelDimensions, 1)
  If (plngWidth = Val(mavPageLabelDimensions(iCount, 1)) And plngHeight = Val(mavPageLabelDimensions(iCount, 2)) And Not pbLandscape) _
    Or (plngHeight = Val(mavPageLabelDimensions(iCount, 1)) And plngWidth = Val(mavPageLabelDimensions(iCount, 2)) And pbLandscape) Then
    cboLabelPageSize.ListIndex = iCount
    bFound = True
  End If
Next iCount

ConvertLabelToKnownType = bFound

End Function

' VB doesn't have a maximum function of its own???? (unless I'm being exceptionally dosile today...)
Private Function MaxValue(psngValue1 As Single, psngValue2 As Single) As Single
  MaxValue = IIf(psngValue1 > psngValue2, psngValue1, psngValue2)
End Function

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEnvelopeDimension_GotFocus(Index As Integer)
  With txtEnvelopeDimension(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub PopulateColourCombos()
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strColour As String
  Dim lngColour As Long
  
  On Local Error GoTo LocalErr

  strSQL = "SELECT ColValue, ColDesc " & _
           "FROM ASRSysColours " & _
           "ORDER BY ColOrder"
           
  
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  Do While Not rsTemp.EOF
 
    With picColour
      .AutoRedraw = True
      .Width = picColour.ScaleX(12, vbPixels, vbTwips)
      .Height = picColour.ScaleY(12, vbPixels, vbTwips)
      .BorderStyle = 0
      .Appearance = 0
      .BackColor = rsTemp!ColValue
      .ForeColor = vbBlack
      picColour.Line (0, 0)-(picColour.Width - picColour.ScaleX(1, vbPixels, vbTwips) _
                 , picColour.Height - picColour.ScaleY(1, vbPixels, vbTwips)), , B
      .Picture = picColour.Image
      ImageList1.ListImages.Add , rsTemp!ColDesc, .Picture
      .Cls
      .Picture = Nothing
    End With

    rsTemp.MoveNext
  Loop

  If Not rsTemp.BOF Or Not rsTemp.EOF Then
    rsTemp.MoveFirst
  
    cboFontColour(0).ComboItems.Clear
    cboFontColour(1).ComboItems.Clear

    cboFontColour(0).ImageList = ImageList1
    cboFontColour(1).ImageList = ImageList1
    
    Do While Not rsTemp.EOF

      strColour = rsTemp!ColDesc
      lngColour = rsTemp!ColValue

      cboFontColour(0).Refresh
      cboFontColour(0).ComboItems.Add , "C" & CStr(lngColour), strColour, strColour
      cboFontColour(0).ComboItems.Item(cboFontColour(0).ComboItems.Count).Tag = lngColour
      cboFontColour(1).Refresh
      cboFontColour(1).ComboItems.Add , "C" & CStr(lngColour), strColour, strColour
      cboFontColour(1).ComboItems.Item(cboFontColour(1).ComboItems.Count).Tag = lngColour

      rsTemp.MoveNext
    Loop
  End If
  
  rsTemp.Close

  Set rsTemp = Nothing

Exit Sub

LocalErr:
  MsgBox Err.Description, vbExclamation, Me.Caption

End Sub

Private Sub FillComboWithFontSizes(objControl As ComboBox)

  With objControl
    .Clear
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "14"
    .AddItem "16"
    .AddItem "18"
    .AddItem "20"
    .AddItem "22"
    .AddItem "24"
    .AddItem "26"
    .AddItem "28"
    .AddItem "36"
    .AddItem "48"
    .AddItem "72"
  End With

End Sub

Private Sub UpdatePreview()

  Dim iCount As Integer
  Dim iXPos As Integer

  lblPreviewHeading.Font = cboFontName(0).Text
  lblPreviewHeading.FontSize = cboFontSize(0).Text
  lblPreviewHeading.ForeColor = cboFontColour(0).ComboItems.Item(cboFontColour(0).SelectedItem.Index).Tag
  lblPreviewHeading.FontBold = IIf(chkFontBold(0).Value = vbChecked, True, False)
  lblPreviewHeading.FontItalic = IIf(chkFontItalic(0).Value = vbChecked, True, False)
  lblPreviewHeading.FontUnderline = IIf(chkFontUnderLine(0).Value = vbChecked, True, False)
  lblPreviewHeading.Width = picPreview.Width
  iXPos = lblPreviewHeading.Height + 150

  For iCount = 0 To lblPreviewStandard.Count - 1
    lblPreviewStandard(iCount).Font = cboFontName(1).Text
    lblPreviewStandard(iCount).FontSize = cboFontSize(1).Text
    lblPreviewStandard(iCount).ForeColor = cboFontColour(1).ComboItems.Item(cboFontColour(1).SelectedItem.Index).Tag
    lblPreviewStandard(iCount).FontBold = IIf(chkFontBold(1).Value = vbChecked, True, False)
    lblPreviewStandard(iCount).FontItalic = IIf(chkFontItalic(1).Value = vbChecked, True, False)
    lblPreviewStandard(iCount).FontUnderline = IIf(chkFontUnderLine(1).Value = vbChecked, True, False)
    lblPreviewStandard(iCount).Top = iXPos
    lblPreviewStandard(iCount).Width = picPreview.Width
    iXPos = iXPos + lblPreviewStandard(iCount).Height
  Next iCount

End Sub

Private Sub PopulateFormatTab()

  ' Fonts
  FillComboWithFonts cboFontName(0)
  FillComboWithFonts cboFontName(1)

  ' Sizes
  FillComboWithFontSizes cboFontSize(0)
  FillComboWithFontSizes cboFontSize(1)

  ' Colours
  PopulateColourCombos

End Sub

Private Sub cboFontName_Click(Index As Integer)
  If Not mbIsLoading Then
    mobjFormatDetails(Index).FontName = cboFontName(Index).Text
    UpdatePreview
    Me.Changed = True
  End If
End Sub

Private Sub cboFontSize_Click(Index As Integer)
  If Not mbIsLoading Then
    mobjFormatDetails(Index).FontSize = cboFontSize(Index).Text
    UpdatePreview
    Me.Changed = True
  End If
End Sub

Private Sub cboFontColour_Click(Index As Integer)
  If Not mbIsLoading Then
    mobjFormatDetails(Index).ForeCol = cboFontColour(Index).ComboItems.Item(cboFontColour(Index).SelectedItem.Index).Tag
    UpdatePreview
    Me.Changed = True
  End If
End Sub

Private Sub chkFontBold_Click(Index As Integer)
  If Not mbIsLoading Then
    mobjFormatDetails(Index).Bold = IIf(chkFontBold(Index).Value = vbChecked, True, False)
    UpdatePreview
    Me.Changed = True
  End If
End Sub

Private Sub chkFontItalic_Click(Index As Integer)
  If Not mbIsLoading Then
    mobjFormatDetails(Index).Italic = IIf(chkFontItalic(Index).Value = vbChecked, True, False)
    UpdatePreview
    Me.Changed = True
  End If
End Sub

Private Sub chkFontUnderLine_Click(Index As Integer)
  If Not mbIsLoading Then
    mobjFormatDetails(Index).Underline = IIf(chkFontUnderLine(Index).Value = vbChecked, True, False)
    UpdatePreview
    Me.Changed = True
  End If
End Sub

' Initialses a new definition
Private Sub NewDefinition()

  ' Heading Format
  mobjFormatDetails(0).FontName = "Verdana"
  mobjFormatDetails(0).FontSize = 12
  mobjFormatDetails(0).ForeCol = 0
  mobjFormatDetails(0).Bold = True
  mobjFormatDetails(0).Italic = False
  mobjFormatDetails(0).Underline = False

  ' Standard Format
  mobjFormatDetails(1).FontName = "Verdana"
  mobjFormatDetails(1).FontSize = 10
  mobjFormatDetails(1).ForeCol = 0
  mobjFormatDetails(1).Bold = False
  mobjFormatDetails(1).Italic = False
  mobjFormatDetails(1).Underline = False

End Sub

Private Sub EnableDisableTabControls()

  ' Definition tab page controls
  fraDefinition.Enabled = (SSTab1.Tab = 0)
  fraType.Enabled = (SSTab1.Tab = 0)
  fraMeasurements.Enabled = (SSTab1.Tab = 0)
  fraEnvelope.Enabled = (SSTab1.Tab = 0)
  fraLabel.Enabled = (SSTab1.Tab = 0)
  
  ' Format controls
  fraFont(0).Enabled = (SSTab1.Tab = 1)
  fraFont(1).Enabled = (SSTab1.Tab = 1)
  fraPreview.Enabled = (SSTab1.Tab = 1)

End Sub

Private Function GetColourName(lngColour As Long) As String

  Dim strSQL As String
  Dim rsColours As Recordset

  ' Get the colour types
  strSQL = "SELECT ColValue, ColDesc " & _
           "FROM ASRSysColours " & _
           "WHERE ColValue = " & Str(lngColour)
  Set rsColours = datGeneral.GetReadOnlyRecords(strSQL)
  
  If Not (rsColours.EOF And rsColours.BOF) Then
    rsColours.MoveFirst
    GetColourName = rsColours.Fields("ColDesc").Value
  Else
    GetColourName = "Unknown"
  End If

End Function



Private Sub txtLabelSideMargin_KeyPress(KeyAscii As Integer)
  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub txtLabelHorizontalPitch_KeyPress(KeyAscii As Integer)
  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub txtLabelVerticalPitch_KeyPress(KeyAscii As Integer)
  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub txtLabelWidth_KeyPress(KeyAscii As Integer)
  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub txtLabelHeight_KeyPress(KeyAscii As Integer)
  ' Turned if statement around to make it easier to read
  If KeyAscii = vbKeyBack Or IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "." Then
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub UpDownChange(pobjUpDownControl As UpDown)

  Dim strDisplayText As String
  Dim objTextBox As TextBox
  Dim bHasChanged As Boolean

  Select Case pobjUpDownControl.Name
    Case "upnLabelTopMargin"
      Set objTextBox = txtLabelTopMargin
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)

    Case "upnLabelSideMargin"
      Set objTextBox = txtLabelSideMargin
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
    
    Case "upnLabelHorizontalPitch"
      Set objTextBox = txtLabelHorizontalPitch
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
    
    Case "upnLabelVerticalPitch"
      Set objTextBox = txtLabelVerticalPitch
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
      
    Case "upnLabelWidth"
      Set objTextBox = txtLabelWidth
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
      
    Case "upnLabelHeight"
      Set objTextBox = txtLabelHeight
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
      
    Case "upnLabelNumberDown"
      Set objTextBox = txtLabelNumberDown
      strDisplayText = Trim(Str(pobjUpDownControl.Value / 100))
      
    Case "upnLabelNumberAcross"
      Set objTextBox = txtLabelNumberAcross
      strDisplayText = Trim(Str(pobjUpDownControl.Value / 100))
  
    Case "upnLabelPageHeight"
      Set objTextBox = txtLabelPageHeight
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
    
    Case "upnLabelPageWidth"
      Set objTextBox = txtLabelPageWidth
      strDisplayText = Format(Str(pobjUpDownControl.Value / msngConversion), "0.00 " & mstrMeasurementText)
  
  End Select

  objTextBox.Text = strDisplayText

  If Not mbIsLoading And Not mbLandscaping Then
    SetLabelMaxAndMinSizes
    'Me.Changed = True
  End If

End Sub
Private Sub upnLabelPageWidth_Change()
  UpDownChange upnLabelPageWidth
End Sub
Private Sub upnLabelPageHeight_Change()
  UpDownChange upnLabelPageHeight
End Sub
Private Sub upnLabelHorizontalPitch_Change()
  UpDownChange upnLabelHorizontalPitch
End Sub
Private Sub upnLabelNumberAcross_Change()
  UpDownChange upnLabelNumberAcross
  upnLabelHorizontalPitch.Enabled = (upnLabelNumberAcross.Value > 100) And Not mblnReadOnly
  txtLabelHorizontalPitch.Enabled = upnLabelHorizontalPitch.Enabled And Not mblnReadOnly
  txtLabelHorizontalPitch.BackColor = IIf(upnLabelHorizontalPitch.Enabled And Not mblnReadOnly, vbWhite, vbButtonFace)
End Sub
Private Sub upnLabelNumberDown_Change()
  UpDownChange upnLabelNumberDown
  upnLabelVerticalPitch.Enabled = (upnLabelNumberDown.Value > 100) And Not mblnReadOnly
  txtLabelVerticalPitch.Enabled = upnLabelVerticalPitch.Enabled And Not mblnReadOnly
  txtLabelVerticalPitch.BackColor = IIf(upnLabelVerticalPitch.Enabled And Not mblnReadOnly, vbWhite, vbButtonFace)
End Sub
Private Sub upnLabelSideMargin_Change()
  UpDownChange upnLabelSideMargin
End Sub
Private Sub upnLabelTopMargin_Change()
  UpDownChange upnLabelTopMargin
End Sub
Private Sub upnLabelVerticalPitch_Change()
  UpDownChange upnLabelVerticalPitch
End Sub
Private Sub upnLabelWidth_Change()
  UpDownChange upnLabelWidth
  
  ' Reset to remembered last setting
  If upnLabelWidth.Value < upnLabelHorizontalPitch.Value Then
    upnLabelHorizontalPitch.Value = IIf(msngRememberHPitch < upnLabelHorizontalPitch.Min, upnLabelHorizontalPitch.Min, msngRememberHPitch)
  End If
  
  ' Force pitch to be equal or greater than width
  If upnLabelWidth.Value > upnLabelHorizontalPitch.Value Then
    upnLabelHorizontalPitch.Value = upnLabelWidth.Value
  End If
  
End Sub

Private Sub upnLabelHeight_Change()
  UpDownChange upnLabelHeight
  
  ' Reset to remembered last setting
  If upnLabelHeight.Value < upnLabelVerticalPitch.Value Then
    upnLabelVerticalPitch.Value = IIf(msngRememberVPitch < upnLabelVerticalPitch.Min, upnLabelVerticalPitch.Min, msngRememberVPitch)
  End If
  
  ' Force pitch to be equal or greater than width
  If upnLabelHeight.Value > upnLabelVerticalPitch.Value Then
    upnLabelVerticalPitch.Value = upnLabelHeight.Value
  End If
End Sub

Private Sub HighlightTextBox(pobjTextBox As TextBox)
  With pobjTextBox
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLabelTopMargin_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelSideMargin_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelHorizontalPitch_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelVerticalPitch_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelWidth_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelHeight_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelNumberAcross_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelNumberDown_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelPageHeight_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub txtLabelPageWidth_GotFocus()
  HighlightTextBox Me.ActiveControl
End Sub
Private Sub TextBoxLostFocus(pobjTextBox As TextBox)

  Dim strSize As String
  Dim intSizeEnd As Integer
  Dim sngEnteredValue As Single
  Dim upnTemp As UpDown
  Dim strMeasurementText As String
  Dim strFormat As String

  ' Get the numeric part of the entered value
  intSizeEnd = InStr(1, pobjTextBox.Text, mstrMeasurementText, vbTextCompare)
  If intSizeEnd > 0 Then
    strSize = Trim(Left(pobjTextBox.Text, intSizeEnd - 1))
  Else
    strSize = pobjTextBox.Text
  End If

  If IsNumeric(strSize) Then
  
    sngEnteredValue = Val(strSize) * msngConversion
  
    Select Case pobjTextBox.Name
      Case "txtLabelTopMargin"
        Set upnTemp = upnLabelTopMargin
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
      
      Case "txtLabelSideMargin"
        Set upnTemp = upnLabelSideMargin
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
      
      Case "txtLabelHorizontalPitch"
        Set upnTemp = upnLabelHorizontalPitch
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
      
      Case "txtLabelVerticalPitch"
        Set upnTemp = upnLabelVerticalPitch
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
        
      Case "txtLabelWidth"
        Set upnTemp = upnLabelWidth
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
        
      Case "txtLabelHeight"
        Set upnTemp = upnLabelHeight
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
        
      Case "txtLabelNumberDown"
        Set upnTemp = upnLabelNumberDown
        strMeasurementText = ""
        strFormat = "0"
        
      Case "txtLabelNumberAcross"
        Set upnTemp = upnLabelNumberAcross
        strMeasurementText = ""
        strFormat = "0"
        
      Case "txtLabelPageHeight"
        Set upnTemp = upnLabelPageHeight
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
      
      Case "txtLabelPageWidth"
        Set upnTemp = upnLabelPageWidth
        strMeasurementText = " " & mstrMeasurementText
        strFormat = "0.00"
        
    End Select

    ' Make sure value is within boundary
    sngEnteredValue = IIf(sngEnteredValue > upnTemp.Max, upnTemp.Max, sngEnteredValue)
    sngEnteredValue = IIf(sngEnteredValue < upnTemp.Min, upnTemp.Min, sngEnteredValue)

    ' Set spinner control value
    upnTemp.Value = sngEnteredValue

    ' Format the entered text
    pobjTextBox.Text = Format(Str(sngEnteredValue / msngConversion), strFormat) & strMeasurementText
  Else
    'NHRD26082004 Fault 7997
    MsgBox "Measurement type should be '" + mstrMeasurementText + "'" + vbCrLf + vbCrLf + "Re-enter numeric part of measurement.", vbExclamation, Me.Caption
    SSTab1.Tab = 0
    pobjTextBox.SetFocus
  End If

End Sub
Private Sub txtLabelTopMargin_LostFocus()
  TextBoxLostFocus txtLabelTopMargin
End Sub
Private Sub txtLabelSideMargin_LostFocus()
  TextBoxLostFocus txtLabelSideMargin
End Sub
Private Sub txtLabelHorizontalPitch_LostFocus()
  TextBoxLostFocus txtLabelHorizontalPitch
  msngRememberHPitch = upnLabelHorizontalPitch.Value
End Sub
Private Sub txtLabelVerticalPitch_LostFocus()
  TextBoxLostFocus txtLabelVerticalPitch
  msngRememberVPitch = upnLabelVerticalPitch.Value
End Sub
Private Sub txtLabelWidth_LostFocus()
  TextBoxLostFocus txtLabelWidth
End Sub
Private Sub txtLabelHeight_LostFocus()
  TextBoxLostFocus txtLabelHeight
End Sub
Private Sub txtLabelNumberAcross_LostFocus()
  TextBoxLostFocus txtLabelNumberAcross
End Sub
Private Sub txtLabelNumberDown_LostFocus()
  TextBoxLostFocus txtLabelNumberDown
End Sub
Private Sub txtLabelPageHeight_LostFocus()
  TextBoxLostFocus txtLabelPageHeight
End Sub
Private Sub txtLabelPageWidth_LostFocus()
  TextBoxLostFocus txtLabelPageWidth
End Sub

Private Sub SetLabelMaxAndMinSizes()

  ' Maximum values (Labels)
  upnLabelVerticalPitch.Max = IIf(chkLabelPageLandscape.Value = vbChecked, msngMaxLabelWidth, msngMaxLabelHeight) * 100
  upnLabelHorizontalPitch.Max = IIf(chkLabelPageLandscape.Value = vbChecked, msngMaxLabelHeight, msngMaxLabelWidth) * 100
  upnLabelWidth.Max = IIf(chkLabelPageLandscape.Value = vbChecked, msngMaxLabelHeight, msngMaxLabelWidth) * 100
  upnLabelHeight.Max = IIf(chkLabelPageLandscape.Value = vbChecked, msngMaxLabelWidth, msngMaxLabelHeight) * 100
  upnLabelPageWidth.Max = IIf(chkLabelPageLandscape.Value = vbChecked, msngMaxLabelPageHeight, msngMaxLabelPageWidth) * 100
  upnLabelPageHeight.Max = IIf(chkLabelPageLandscape.Value = vbChecked, msngMaxLabelPageWidth, msngMaxLabelPageHeight) * 100

  ' Minimum values (Labels)
  upnLabelTopMargin.Min = msngLabelMinPageMargin * 100
  upnLabelSideMargin.Min = msngLabelMinPageMargin * 100
  upnLabelVerticalPitch.Min = IIf(chkLabelPageLandscape.Value = vbChecked, upnLabelWidth.Value, upnLabelHeight.Value)
  upnLabelHorizontalPitch.Min = IIf(chkLabelPageLandscape.Value = vbChecked, upnLabelHeight.Value, upnLabelWidth.Value)
'  upnLabelVerticalPitch.Min = IIf(chkLabelPageLandscape.Value = vbChecked, msngMinLabelWidth, msngMinLabelHeight) * 100
'  upnLabelHorizontalPitch.Min = IIf(chkLabelPageLandscape.Value = vbChecked, msngMinLabelHeight, msngMinLabelWidth) * 100
  upnLabelWidth.Min = IIf(chkLabelPageLandscape.Value = vbChecked, msngMinLabelWidth, msngMinLabelHeight) * 100
  upnLabelHeight.Min = IIf(chkLabelPageLandscape.Value = vbChecked, msngMinLabelHeight, msngMinLabelWidth) * 100
  upnLabelPageWidth.Min = msngMinLabelPageWidth * 100
  upnLabelPageHeight.Min = msngMinLabelPageWidth * 100
  upnLabelNumberAcross.Min = 100
  upnLabelNumberDown.Min = 100

End Sub

Private Sub upnLabelVerticalPitch_DownClick()
  msngRememberVPitch = upnLabelVerticalPitch.Value
End Sub
Private Sub upnLabelVerticalPitch_UpClick()
  msngRememberVPitch = upnLabelVerticalPitch.Value
End Sub
Private Sub upnLabelHorizontalPitch_DownClick()
  msngRememberHPitch = upnLabelHorizontalPitch.Value
End Sub
Private Sub upnLabelHorizontalPitch_UpClick()
  msngRememberHPitch = upnLabelHorizontalPitch.Value
End Sub

Private Sub SetLabelMaxAndMinSizesToZero()

  ' Maximum values (Labels)
  upnLabelVerticalPitch.Max = 99999
  upnLabelHorizontalPitch.Max = 99999
  upnLabelWidth.Max = 99999
  upnLabelHeight.Max = 99999
  upnLabelPageWidth.Max = 99999
  upnLabelPageHeight.Max = 99999

  ' Minimum values (Labels)
  upnLabelTopMargin.Min = 0
  upnLabelSideMargin.Min = 0
  upnLabelVerticalPitch.Min = 0
  upnLabelHorizontalPitch.Min = 0
  upnLabelWidth.Min = 0
  upnLabelHeight.Min = 0
  upnLabelPageWidth.Min = 0
  upnLabelPageHeight.Min = 0
  upnLabelNumberAcross.Min = 0
  upnLabelNumberDown.Min = 0

End Sub


