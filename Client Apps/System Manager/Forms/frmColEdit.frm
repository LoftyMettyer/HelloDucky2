VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "TIMASK6.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "COA_ColourPicker.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Object = "{96E404DC-B217-4A2D-A891-C73A92A628CC}#1.0#0"; "COA_WorkingPattern.ocx"
Object = "{19400013-2704-42FE-AAA4-45D1A725A895}#1.0#0"; "COA_ColourSelector.ocx"
Begin VB.Form frmColEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Column Properties"
   ClientHeight    =   6525
   ClientLeft      =   300
   ClientTop       =   1920
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5007
   Icon            =   "frmColEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7140
      TabIndex        =   89
      Top             =   6030
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   88
      Top             =   6030
      Width           =   1200
   End
   Begin TabDlg.SSTab tabColProps 
      Height          =   5940
      Left            =   60
      TabIndex        =   91
      Tag             =   "QA"
      Top             =   30
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
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
      TabPicture(0)   =   "frmColEdit.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDefinitionPage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Screen Control "
      TabPicture(1)   =   "frmColEdit.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraControlPage"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Opt&ions"
      TabPicture(2)   =   "frmColEdit.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOptionsPage"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Valida&tion"
      TabPicture(3)   =   "frmColEdit.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraValidationPage"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Diar&y Links"
      TabPicture(4)   =   "frmColEdit.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDiaryLinkPage"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Email Lin&ks"
      TabPicture(5)   =   "frmColEdit.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraEmailLinkPage"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "A&fd PostCode"
      TabPicture(6)   =   "frmColEdit.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraAfdPage"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&Quick Address"
      TabPicture(7)   =   "frmColEdit.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraQAPage"
      Tab(7).ControlCount=   1
      Begin VB.Frame fraQAPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74955
         TabIndex        =   148
         Top             =   320
         Width           =   8205
         Begin VB.Frame fraFieldMapping 
            Caption         =   "Column Mapping :"
            Height          =   4375
            Index           =   1
            Left            =   200
            TabIndex        =   150
            Tag             =   "QA"
            Top             =   720
            Width           =   7815
            Begin VB.ComboBox cboQACounty 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   158
               Tag             =   "QA"
               Top             =   2115
               Width           =   2310
            End
            Begin VB.ComboBox cboQATown 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   157
               Tag             =   "QA"
               Top             =   1770
               Width           =   2310
            End
            Begin VB.ComboBox cboQALocality 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   156
               Tag             =   "QA"
               Top             =   1425
               Width           =   2310
            End
            Begin VB.ComboBox cboQAStreet 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   155
               Tag             =   "QA"
               Top             =   1095
               Width           =   2310
            End
            Begin VB.ComboBox cboQAProperty 
               Height          =   315
               ItemData        =   "frmColEdit.frx":00EC
               Left            =   1485
               List            =   "frmColEdit.frx":00EE
               Style           =   2  'Dropdown List
               TabIndex        =   154
               Tag             =   "QA"
               Top             =   750
               Width           =   2310
            End
            Begin VB.ComboBox cboQAAddress 
               Height          =   315
               Left            =   5205
               Style           =   2  'Dropdown List
               TabIndex        =   153
               Tag             =   "QA"
               Top             =   750
               Width           =   2310
            End
            Begin VB.OptionButton optQAAddressType 
               Caption         =   "In&dividual Address Columns :"
               Height          =   270
               Index           =   0
               Left            =   200
               TabIndex        =   152
               Tag             =   "QA"
               Top             =   345
               Value           =   -1  'True
               Width           =   3030
            End
            Begin VB.OptionButton optQAAddressType 
               Caption         =   "Si&ngle Address Column :"
               Height          =   270
               Index           =   1
               Left            =   4020
               TabIndex        =   151
               Tag             =   "QA"
               Top             =   345
               Width           =   2730
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "County :"
               Height          =   255
               Index           =   28
               Left            =   495
               TabIndex        =   164
               Tag             =   "QA"
               Top             =   2175
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Town :"
               Height          =   255
               Index           =   27
               Left            =   495
               TabIndex        =   163
               Tag             =   "QA"
               Top             =   1830
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Locality :"
               Height          =   255
               Index           =   26
               Left            =   495
               TabIndex        =   162
               Tag             =   "QA"
               Top             =   1485
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Street :"
               Height          =   255
               Index           =   25
               Left            =   495
               TabIndex        =   161
               Tag             =   "QA"
               Top             =   1155
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Property :"
               Height          =   255
               Index           =   24
               Left            =   495
               TabIndex        =   160
               Tag             =   "QA"
               Top             =   810
               Width           =   870
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Address :"
               Height          =   255
               Index           =   23
               Left            =   4320
               TabIndex        =   159
               Tag             =   "QA"
               Top             =   810
               Width           =   855
            End
         End
         Begin VB.CheckBox chkQAPostCodeColumn 
            Caption         =   "Is this a postcode column which you wish to be 'Quick Address &enabled' ?"
            Height          =   255
            Left            =   200
            TabIndex        =   149
            Top             =   300
            Width           =   6945
         End
      End
      Begin VB.Frame fraEmailLinkPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74955
         TabIndex        =   131
         Top             =   320
         Visible         =   0   'False
         Width           =   8205
         Begin VB.Frame fraEmail 
            Caption         =   "Email Links :"
            Height          =   5025
            Left            =   200
            TabIndex        =   132
            Top             =   200
            Width           =   7815
            Begin VB.CommandButton cmdAddEmailLink 
               Caption         =   "&New"
               Enabled         =   0   'False
               Height          =   400
               Left            =   420
               TabIndex        =   70
               Top             =   4440
               Width           =   1200
            End
            Begin VB.CommandButton cmdRemoveEmailLink 
               Caption         =   "&Delete"
               Height          =   400
               Left            =   4210
               TabIndex        =   72
               Top             =   4440
               Width           =   1200
            End
            Begin VB.CommandButton cmdRemoveAllEmailLinks 
               Caption         =   "Delete &All"
               Height          =   400
               Left            =   6105
               TabIndex        =   73
               Top             =   4440
               Width           =   1200
            End
            Begin VB.CommandButton cmdEmailLinkProperties 
               Caption         =   "&Edit"
               Height          =   400
               Left            =   2315
               TabIndex        =   71
               Top             =   4440
               Width           =   1200
            End
            Begin SSDataWidgets_B.SSDBGrid ssGrdEmailLinks 
               Height          =   3900
               Left            =   195
               TabIndex        =   74
               Top             =   315
               Width           =   7425
               ScrollBars      =   2
               _Version        =   196617
               DataMode        =   2
               RecordSelectors =   0   'False
               Col.Count       =   4
               DefColWidth     =   26458
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
               RowNavigation   =   3
               MaxSelectedRows =   1
               ForeColorEven   =   0
               BackColorEven   =   -2147483643
               BackColorOdd    =   -2147483643
               RowHeight       =   423
               Columns.Count   =   4
               Columns(0).Width=   5001
               Columns(0).Caption=   "Title"
               Columns(0).Name =   "colTitle"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               Columns(1).Width=   1799
               Columns(1).Caption=   "Offset"
               Columns(1).Name =   "colOffset"
               Columns(1).DataField=   "Column 1"
               Columns(1).DataType=   8
               Columns(1).FieldLen=   256
               Columns(2).Width=   5821
               Columns(2).Caption=   "Subject"
               Columns(2).Name =   "colSubject"
               Columns(2).DataField=   "Column 2"
               Columns(2).DataType=   8
               Columns(2).FieldLen=   256
               Columns(3).Width=   26458
               Columns(3).Visible=   0   'False
               Columns(3).Caption=   "colLinkID"
               Columns(3).Name =   "colLinkID"
               Columns(3).DataField=   "Column 3"
               Columns(3).DataType=   8
               Columns(3).FieldLen=   256
               UseDefaults     =   0   'False
               TabNavigation   =   1
               _ExtentX        =   13097
               _ExtentY        =   6879
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
      End
      Begin VB.Frame fraAfdPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74955
         TabIndex        =   119
         Top             =   320
         Width           =   8205
         Begin VB.CheckBox chkAFDPostCodeColumn 
            Caption         =   "Is this a postcode column which you wish to be 'Afd &enabled' ?"
            Height          =   255
            Left            =   200
            TabIndex        =   75
            Top             =   300
            Width           =   5790
         End
         Begin VB.Frame fraFieldMapping 
            Caption         =   "Column Mapping :"
            Height          =   4375
            Index           =   0
            Left            =   200
            TabIndex        =   120
            Tag             =   "AFD"
            Top             =   720
            Width           =   7815
            Begin VB.ComboBox cboAFDForename 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Tag             =   "AFD"
               Top             =   405
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDInitial 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Tag             =   "AFD"
               Top             =   750
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDSurname 
               Height          =   315
               Left            =   5235
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Tag             =   "AFD"
               Top             =   405
               Width           =   2310
            End
            Begin VB.OptionButton optAFDAddressType 
               Caption         =   "Si&ngle Address Column :"
               Height          =   270
               Index           =   1
               Left            =   4050
               TabIndex        =   81
               Tag             =   "AFD"
               Top             =   1395
               Width           =   2685
            End
            Begin VB.OptionButton optAFDAddressType 
               Caption         =   "In&dividual Address Columns :"
               Height          =   270
               Index           =   0
               Left            =   200
               TabIndex        =   80
               Tag             =   "AFD"
               Top             =   1395
               Value           =   -1  'True
               Width           =   2985
            End
            Begin VB.ComboBox cboAFDAddress 
               Height          =   315
               Left            =   5235
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Tag             =   "AFD"
               Top             =   1800
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDProperty 
               Height          =   315
               ItemData        =   "frmColEdit.frx":00F0
               Left            =   1485
               List            =   "frmColEdit.frx":00F2
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Tag             =   "AFD"
               Top             =   1800
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDStreet 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Tag             =   "AFD"
               Top             =   2145
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDLocality 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Tag             =   "AFD"
               Top             =   2475
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDTown 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   85
               Tag             =   "AFD"
               Top             =   2820
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDCounty 
               Height          =   315
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Tag             =   "AFD"
               Top             =   3165
               Width           =   2310
            End
            Begin VB.ComboBox cboAFDTelephone 
               Height          =   315
               Left            =   5235
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Tag             =   "AFD"
               Top             =   750
               Width           =   2310
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Forename :"
               Height          =   255
               Index           =   0
               Left            =   195
               TabIndex        =   130
               Tag             =   "AFD"
               Top             =   465
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Initial(s) :"
               Height          =   255
               Index           =   1
               Left            =   195
               TabIndex        =   129
               Tag             =   "AFD"
               Top             =   810
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Surname :"
               Height          =   255
               Index           =   2
               Left            =   4050
               TabIndex        =   128
               Tag             =   "AFD"
               Top             =   465
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Address :"
               Height          =   255
               Index           =   3
               Left            =   4350
               TabIndex        =   127
               Tag             =   "AFD"
               Top             =   1860
               Width           =   855
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Property :"
               Height          =   255
               Index           =   4
               Left            =   495
               TabIndex        =   126
               Tag             =   "AFD"
               Top             =   1860
               Width           =   870
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Street :"
               Height          =   255
               Index           =   5
               Left            =   495
               TabIndex        =   125
               Tag             =   "AFD"
               Top             =   2205
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Locality :"
               Height          =   255
               Index           =   6
               Left            =   495
               TabIndex        =   124
               Tag             =   "AFD"
               Top             =   2535
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Town :"
               Height          =   255
               Index           =   7
               Left            =   495
               TabIndex        =   123
               Tag             =   "AFD"
               Top             =   2880
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "County :"
               Height          =   255
               Index           =   8
               Left            =   495
               TabIndex        =   122
               Tag             =   "AFD"
               Top             =   3225
               Width           =   1005
            End
            Begin VB.Label lblFieldMapping 
               BackStyle       =   0  'Transparent
               Caption         =   "Telephone :"
               Height          =   255
               Index           =   10
               Left            =   4050
               TabIndex        =   121
               Tag             =   "AFD"
               Top             =   810
               Width           =   1095
            End
         End
      End
      Begin VB.Frame fraOptionsPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74950
         TabIndex        =   90
         Top             =   320
         Visible         =   0   'False
         Width           =   8205
         Begin VB.Frame fraDefault 
            Caption         =   "Default Value :"
            Height          =   2235
            Left            =   210
            TabIndex        =   136
            Top             =   3075
            Width           =   7815
            Begin VB.ComboBox cboDefault 
               Height          =   315
               Left            =   2745
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   1590
               Width           =   1400
            End
            Begin VB.CommandButton cmdDfltValueExpression 
               Caption         =   "..."
               Height          =   315
               Left            =   7260
               TabIndex        =   54
               Top             =   700
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtDfltValueExpression 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1830
               TabIndex        =   53
               Top             =   700
               Width           =   5430
            End
            Begin VB.OptionButton optDfltType 
               Caption         =   "&Value"
               Height          =   195
               Index           =   0
               Left            =   200
               TabIndex        =   43
               Top             =   360
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton optDfltType 
               Caption         =   "Ca&lculation"
               Height          =   195
               Index           =   1
               Left            =   200
               TabIndex        =   44
               Top             =   760
               Width           =   1500
            End
            Begin VB.Frame fraLogicDefaults 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   195
               Left            =   330
               TabIndex        =   137
               Top             =   1305
               Width           =   2000
               Begin VB.OptionButton optDefault 
                  Caption         =   "F&alse"
                  Height          =   195
                  Index           =   1
                  Left            =   1130
                  TabIndex        =   47
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton optDefault 
                  Caption         =   "T&rue"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   46
                  Top             =   0
                  Width           =   900
               End
            End
            Begin COAWorkingPattern.COA_WorkingPattern ASRDefaultWorkingPattern 
               Height          =   765
               Left            =   4020
               TabIndex        =   51
               Top             =   1260
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   1349
            End
            Begin COASpinner.COA_Spinner asrDefault 
               Height          =   315
               Left            =   1695
               TabIndex        =   49
               Top             =   1605
               Width           =   1005
               _ExtentX        =   1773
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
               MaximumValue    =   250
               Text            =   "0"
            End
            Begin TDBNumberCtrl.TDBNumber TDBDefaultNumber 
               Height          =   315
               Left            =   5820
               TabIndex        =   52
               Top             =   1875
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   556
               _Version        =   65537
               AlignHorizontal =   1
               ClipMode        =   0
               ErrorBeep       =   0   'False
               ReadOnly        =   0   'False
               HighlightText   =   -1  'True
               ZeroAllowed     =   -1  'True
               MinusColor      =   0
               MaxValue        =   99999999
               MinValue        =   -99999999
               Value           =   0
               SelStart        =   0
               SelLength       =   0
               KeyClear        =   "{F2}"
               KeyNext         =   ""
               KeyPopup        =   "{SPACE}"
               KeyPrevious     =   ""
               KeyThreeZero    =   ""
               SepDecimal      =   "."
               SepThousand     =   ","
               Text            =   ""
               Format          =   "##,###,###; -##,###,###"
               DisplayFormat   =   "##,###,###; -##,###,###"
               Appearance      =   1
               BackColor       =   -2147483643
               Enabled         =   -1  'True
               ForeColor       =   -2147483640
               BorderStyle     =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               DropdownButton  =   0   'False
               SpinButton      =   0   'False
               Caption         =   "&Caption"
               CaptionAlignment=   3
               CaptionColor    =   0
               CaptionWidth    =   0
               CaptionPosition =   0
               CaptionSpacing  =   3
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SpinAutowrap    =   0   'False
               _StockProps     =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "frmColEdit.frx":00F4
               MousePointer    =   0
            End
            Begin GTMaskDate.GTMaskDate ASRDate1 
               Height          =   315
               Left            =   240
               TabIndex        =   48
               Top             =   1620
               Width           =   1395
               _Version        =   65537
               _ExtentX        =   2469
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
               AutoSelect      =   -1  'True
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
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.TextBox txtDefault 
               Height          =   315
               Left            =   1845
               TabIndex        =   45
               Text            =   "txtDefault"
               Top             =   300
               Width           =   5730
            End
            Begin COAColourSelector.COA_ColourSelector selDefaultColour 
               Height          =   315
               Left            =   1755
               TabIndex        =   173
               Top             =   375
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
            End
         End
         Begin VB.Frame fraStorage 
            Caption         =   "Storage Type : "
            Height          =   1905
            Left            =   210
            TabIndex        =   165
            Top             =   3075
            Width           =   7815
            Begin VB.CheckBox chkEnableOLEMaxSize 
               Caption         =   "Enable document e&mbedding"
               Enabled         =   0   'False
               Height          =   285
               Left            =   540
               TabIndex        =   170
               Top             =   1290
               Width           =   2850
            End
            Begin VB.OptionButton optOLEStorageType 
               Caption         =   "Lin&ked / Embedded in database"
               Height          =   270
               Index           =   2
               Left            =   180
               TabIndex        =   168
               Top             =   945
               Width           =   3270
            End
            Begin VB.OptionButton optOLEStorageType 
               Caption         =   "Copi&ed to local OLE directory"
               Height          =   270
               Index           =   0
               Left            =   180
               TabIndex        =   167
               Top             =   615
               Width           =   3630
            End
            Begin VB.OptionButton optOLEStorageType 
               Caption         =   "Co&pied to server OLE directory"
               Height          =   270
               Index           =   1
               Left            =   180
               TabIndex        =   166
               Top             =   285
               Value           =   -1  'True
               Width           =   3630
            End
            Begin COASpinner.COA_Spinner asrMaxOLESize 
               Height          =   315
               Left            =   5700
               TabIndex        =   171
               Top             =   1275
               Width           =   735
               _ExtentX        =   1296
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
               Increment       =   100
               MaximumValue    =   8000
               Text            =   "100"
            End
            Begin VB.Label lblMb 
               Caption         =   "Kb"
               Height          =   195
               Left            =   6540
               TabIndex        =   172
               Top             =   1320
               Width           =   285
            End
            Begin VB.Label lblMaximumOLESize 
               Caption         =   "Maximum size :"
               Height          =   300
               Left            =   4215
               TabIndex        =   169
               Top             =   1320
               Width           =   1395
            End
         End
         Begin VB.Frame fraOptions 
            Caption         =   "Options :"
            Height          =   735
            Left            =   210
            TabIndex        =   110
            Top             =   200
            Width           =   7815
            Begin VB.CheckBox chkReadOnly 
               Caption         =   "Rea&d only"
               Height          =   195
               Left            =   200
               TabIndex        =   32
               Top             =   300
               Width           =   1365
            End
            Begin VB.CheckBox chkAudit 
               Caption         =   "&Audit"
               Height          =   195
               Left            =   4215
               TabIndex        =   33
               Top             =   300
               Width           =   975
            End
         End
         Begin VB.Frame fraFormat 
            Caption         =   "Format :"
            Height          =   1845
            Left            =   210
            TabIndex        =   111
            Top             =   1095
            Width           =   7815
            Begin VB.CheckBox chkUse1000Separator 
               Caption         =   "&Use 1000 separator"
               Height          =   255
               Left            =   4215
               TabIndex        =   42
               Top             =   585
               Width           =   2325
            End
            Begin VB.ComboBox cboTrimming 
               Height          =   315
               ItemData        =   "frmColEdit.frx":0110
               Left            =   1770
               List            =   "frmColEdit.frx":0120
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   1410
               Width           =   1755
            End
            Begin VB.ComboBox cboTextAlignment 
               Height          =   315
               ItemData        =   "frmColEdit.frx":014F
               Left            =   1770
               List            =   "frmColEdit.frx":015C
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   1005
               Width           =   1755
            End
            Begin VB.ComboBox cboCase 
               Height          =   315
               ItemData        =   "frmColEdit.frx":0185
               Left            =   1770
               List            =   "frmColEdit.frx":0195
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   600
               Width           =   1755
            End
            Begin VB.CheckBox chkZeroBlank 
               Caption         =   "Blank if &zero"
               Height          =   255
               Left            =   4215
               TabIndex        =   41
               Top             =   240
               Width           =   1560
            End
            Begin VB.CheckBox chkMultiLine 
               Caption         =   "&Multi-line"
               Height          =   255
               Left            =   180
               TabIndex        =   34
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label lblTrimming 
               Caption         =   "Trimming :"
               Height          =   270
               Left            =   210
               TabIndex        =   39
               Top             =   1470
               Width           =   1230
            End
            Begin VB.Label lblTextAlignment 
               Caption         =   "Text Alignment :"
               Height          =   195
               Left            =   210
               TabIndex        =   37
               Top             =   1065
               Width           =   1500
            End
            Begin VB.Label lblCase 
               Caption         =   "Case :"
               Height          =   195
               Left            =   210
               TabIndex        =   35
               Top             =   660
               Width           =   690
            End
         End
      End
      Begin VB.Frame fraValidationPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74950
         TabIndex        =   108
         Top             =   320
         Visible         =   0   'False
         Width           =   8205
         Begin VB.Frame fraMask 
            Caption         =   "Mask Validation :"
            Height          =   1600
            Left            =   200
            TabIndex        =   140
            Top             =   2355
            Width           =   7815
            Begin VB.Frame fraMaskKey 
               Caption         =   "Key : "
               Height          =   795
               Left            =   200
               TabIndex        =   141
               Top             =   650
               Width           =   7230
               Begin VB.Label lblMaskKey6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "# - Numbers, Symbols"
                  Height          =   195
                  Left            =   5100
                  TabIndex        =   147
                  Top             =   495
                  Width           =   1950
               End
               Begin VB.Label lblMaskKey5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "9 - Numbers (0-9)"
                  Height          =   195
                  Left            =   5100
                  TabIndex        =   146
                  Top             =   240
                  Width           =   1560
               End
               Begin VB.Label lblMaskKey2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "a - Lowercase"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   145
                  Top             =   495
                  Width           =   1200
               End
               Begin VB.Label lblMaskKey4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "s - Lowercase or space"
                  Height          =   195
                  Left            =   2295
                  TabIndex        =   144
                  Top             =   495
                  Width           =   1980
               End
               Begin VB.Label lblMaskKey3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "S - Uppercase or space"
                  Height          =   195
                  Left            =   2295
                  TabIndex        =   143
                  Top             =   240
                  Width           =   2010
               End
               Begin VB.Label lblMaskKey1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "A - Uppercase"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   142
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.TextBox txtMask 
               Height          =   315
               Left            =   200
               TabIndex        =   60
               Top             =   300
               Width           =   7230
            End
            Begin TDBMask6Ctl.TDBMask txtMaskTest 
               Height          =   180
               Left            =   0
               TabIndex        =   61
               Top             =   700
               Visible         =   0   'False
               Width           =   375
               _Version        =   65536
               _ExtentX        =   661
               _ExtentY        =   317
               Caption         =   "frmColEdit.frx":01C6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmColEdit.frx":022B
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&"
               HighlightText   =   0
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
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "_"
               Value           =   ""
            End
         End
         Begin VB.Frame fraStandardValidation 
            Caption         =   "Standard Validation :"
            Height          =   2055
            Left            =   200
            TabIndex        =   138
            Top             =   200
            Width           =   7815
            Begin VB.ListBox lstUniqueParents 
               Height          =   735
               Left            =   3120
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   59
               Top             =   1155
               Width           =   3825
            End
            Begin VB.CheckBox chkUnique 
               Caption         =   "&Unique (within the entire table)"
               Height          =   195
               Left            =   200
               TabIndex        =   57
               Top             =   560
               Width           =   3540
            End
            Begin VB.CheckBox chkMandatory 
               Caption         =   "&Mandatory"
               Height          =   195
               Left            =   4215
               TabIndex        =   56
               Top             =   250
               Width           =   1500
            End
            Begin VB.CheckBox chkDuplicate 
               Caption         =   "Duplicate C&heck"
               Height          =   195
               Left            =   200
               TabIndex        =   55
               Top             =   250
               Width           =   2445
            End
            Begin VB.CheckBox chkChildUnique 
               Caption         =   "Uni&que (within sibling records)"
               Height          =   195
               Left            =   200
               TabIndex        =   58
               Top             =   870
               Width           =   3450
            End
            Begin VB.Label lblSiblingParentList 
               AutoSize        =   -1  'True
               Caption         =   "Check for uniqueness within sibling records related to parents :"
               Height          =   390
               Left            =   195
               TabIndex        =   139
               Top             =   1215
               Width           =   2535
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraLostFocusClause 
            Caption         =   "Custom Validation :"
            Height          =   1175
            Left            =   200
            TabIndex        =   114
            Top             =   4065
            Width           =   7815
            Begin VB.TextBox txtErrorMessage 
               Height          =   315
               Left            =   1665
               MaxLength       =   255
               TabIndex        =   64
               Top             =   700
               Width           =   5715
            End
            Begin VB.TextBox txtLostFocusClause 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   200
               TabIndex        =   62
               Top             =   300
               Width           =   6870
            End
            Begin VB.CommandButton cmdLostFocusClause 
               Caption         =   "..."
               Height          =   315
               Left            =   7065
               TabIndex        =   63
               Top             =   300
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.Label lblErrorMessage 
               Caption         =   "Error Message :"
               Height          =   195
               Left            =   195
               TabIndex        =   135
               Top             =   765
               Width           =   1365
            End
         End
      End
      Begin VB.Frame fraDiaryLinkPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74955
         TabIndex        =   109
         Top             =   320
         Visible         =   0   'False
         Width           =   8205
         Begin VB.Frame fraDiary 
            Caption         =   "Diary Links :"
            Height          =   5025
            Left            =   200
            TabIndex        =   112
            Top             =   200
            Width           =   7815
            Begin VB.CommandButton cmdDiaryLinkProperties 
               Caption         =   "&Edit"
               Height          =   400
               Left            =   2315
               TabIndex        =   66
               Top             =   4440
               Width           =   1200
            End
            Begin VB.CommandButton cmdRemoveAllDiaryLinks 
               Caption         =   "Delete &All"
               Height          =   400
               Left            =   6105
               TabIndex        =   68
               Top             =   4440
               Width           =   1200
            End
            Begin VB.CommandButton cmdRemoveDiaryLink 
               Caption         =   "&Delete"
               Height          =   400
               Left            =   4210
               TabIndex        =   67
               Top             =   4440
               Width           =   1200
            End
            Begin VB.CommandButton cmdAddDiaryLink 
               Caption         =   "&New"
               Height          =   400
               Left            =   400
               TabIndex        =   65
               Top             =   4440
               Width           =   1200
            End
            Begin SSDataWidgets_B.SSDBGrid ssGrdDiaryLinks 
               Height          =   3900
               Left            =   200
               TabIndex        =   69
               Top             =   320
               Width           =   7425
               ScrollBars      =   2
               _Version        =   196617
               DataMode        =   2
               RecordSelectors =   0   'False
               Col.Count       =   8
               DefColWidth     =   26458
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
               RowNavigation   =   3
               MaxSelectedRows =   1
               ForeColorEven   =   0
               BackColorEven   =   -2147483643
               BackColorOdd    =   -2147483643
               RowHeight       =   423
               Columns.Count   =   8
               Columns(0).Width=   1773
               Columns(0).Caption=   "Comment"
               Columns(0).Name =   "colComment"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               Columns(0).Locked=   -1  'True
               Columns(1).Width=   1773
               Columns(1).Caption=   "Offset"
               Columns(1).Name =   "colOffset"
               Columns(1).DataField=   "Column 1"
               Columns(1).DataType=   8
               Columns(1).FieldLen=   256
               Columns(1).Locked=   -1  'True
               Columns(2).Width=   2302
               Columns(2).Caption=   "Alarmed Events"
               Columns(2).Name =   "colReminder"
               Columns(2).DataField=   "Column 2"
               Columns(2).DataType=   11
               Columns(2).FieldLen=   256
               Columns(2).Locked=   -1  'True
               Columns(2).Style=   2
               Columns(3).Width=   26458
               Columns(3).Visible=   0   'False
               Columns(3).Caption=   "OffsetValue"
               Columns(3).Name =   "colOffsetValue"
               Columns(3).DataField=   "Column 3"
               Columns(3).DataType=   8
               Columns(3).FieldLen=   256
               Columns(3).Locked=   -1  'True
               Columns(4).Width=   26458
               Columns(4).Visible=   0   'False
               Columns(4).Caption=   "PeriodValue"
               Columns(4).Name =   "colPeriodValue"
               Columns(4).DataField=   "Column 4"
               Columns(4).DataType=   8
               Columns(4).FieldLen=   256
               Columns(4).Locked=   -1  'True
               Columns(5).Width=   26458
               Columns(5).Visible=   0   'False
               Columns(5).Caption=   "FilterID"
               Columns(5).Name =   "FilterID"
               Columns(5).DataField=   "Column 5"
               Columns(5).DataType=   8
               Columns(5).FieldLen=   256
               Columns(6).Width=   26458
               Columns(6).Visible=   0   'False
               Columns(6).Caption=   "Effective Date"
               Columns(6).Name =   "Effective Date"
               Columns(6).DataField=   "Column 6"
               Columns(6).DataType=   8
               Columns(6).FieldLen=   256
               Columns(7).Width=   26458
               Columns(7).Visible=   0   'False
               Columns(7).Caption=   "CheckLeavingDate"
               Columns(7).Name =   "CheckLeavingDate"
               Columns(7).DataField=   "Column 7"
               Columns(7).DataType=   8
               Columns(7).FieldLen=   256
               UseDefaults     =   0   'False
               TabNavigation   =   1
               _ExtentX        =   13097
               _ExtentY        =   6879
               _StockProps     =   79
               BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
      Begin VB.Frame fraControlPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5350
         Left            =   -74950
         TabIndex        =   92
         Top             =   320
         Visible         =   0   'False
         Width           =   8205
         Begin VB.Frame fraListValues 
            Caption         =   "Control Values :"
            Height          =   1850
            Left            =   195
            TabIndex        =   97
            Top             =   585
            Width           =   7815
            Begin VB.TextBox txtListValues 
               Height          =   1335
               Left            =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   27
               Top             =   300
               Width           =   7365
            End
         End
         Begin VB.Frame fraStatusBarMessage 
            Caption         =   "Status Bar Message :"
            Height          =   800
            Left            =   195
            TabIndex        =   115
            Top             =   4300
            Width           =   7815
            Begin VB.TextBox txtStatusBarMessage 
               Height          =   285
               Left            =   180
               TabIndex        =   31
               Top             =   300
               Width           =   7395
            End
         End
         Begin VB.ComboBox cboControl 
            Height          =   315
            ItemData        =   "frmColEdit.frx":026D
            Left            =   1575
            List            =   "frmColEdit.frx":0286
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   200
            Width           =   3000
         End
         Begin VB.Frame fraSpinnerProperties 
            Caption         =   "Spinner Properties :"
            Height          =   1600
            Left            =   195
            TabIndex        =   93
            Top             =   2700
            Width           =   7815
            Begin COASpinner.COA_Spinner asrMinVal 
               Height          =   315
               Left            =   1890
               TabIndex        =   28
               Top             =   300
               Width           =   810
               _ExtentX        =   1429
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
               MaximumValue    =   32767
               MinimumValue    =   -32767
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner asrMaxVal 
               Height          =   315
               Left            =   1890
               TabIndex        =   29
               Top             =   705
               Width           =   810
               _ExtentX        =   1429
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
               MaximumValue    =   32767
               MinimumValue    =   -32767
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner asrIncVal 
               Height          =   315
               Left            =   1890
               TabIndex        =   30
               Top             =   1095
               Width           =   810
               _ExtentX        =   1429
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
               MaximumValue    =   32767
               MinimumValue    =   1
               Text            =   "1"
            End
            Begin VB.Label lblMinimumValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Minimum Value :"
               Height          =   195
               Left            =   195
               TabIndex        =   96
               Top             =   360
               Width           =   1140
            End
            Begin VB.Label lblMaximumValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Maximum Value :"
               Height          =   195
               Left            =   200
               TabIndex        =   95
               Top             =   760
               Width           =   1200
            End
            Begin VB.Label lblIncrement 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Increment :"
               Height          =   195
               Left            =   195
               TabIndex        =   94
               Top             =   1155
               Width           =   840
            End
         End
         Begin VB.Label lblControlType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Control Type :"
            Height          =   195
            Left            =   195
            TabIndex        =   98
            Top             =   255
            Width           =   1260
         End
      End
      Begin VB.Frame fraDefinitionPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   5550
         Left            =   45
         TabIndex        =   99
         Top             =   315
         Width           =   8205
         Begin VB.TextBox txtColumnName 
            Height          =   315
            Left            =   900
            MaxLength       =   30
            TabIndex        =   0
            Text            =   "txtColName"
            Top             =   200
            Width           =   7110
         End
         Begin VB.Frame fraColumnType 
            Caption         =   "Column Type :"
            Height          =   750
            Left            =   200
            TabIndex        =   106
            Top             =   600
            Width           =   7815
            Begin VB.OptionButton optColumnType 
               Caption         =   "&Link"
               Height          =   255
               Index           =   3
               Left            =   4545
               TabIndex        =   4
               Tag             =   "4"
               Top             =   300
               Width           =   1005
            End
            Begin VB.OptionButton optColumnType 
               Caption         =   "C&alculated"
               Height          =   255
               Index           =   2
               Left            =   2896
               TabIndex        =   3
               Tag             =   "2"
               Top             =   300
               Width           =   1500
            End
            Begin VB.OptionButton optColumnType 
               Caption         =   "Look&up"
               Height          =   255
               Index           =   1
               Left            =   1448
               TabIndex        =   2
               Tag             =   "1"
               Top             =   300
               Width           =   1305
            End
            Begin VB.OptionButton optColumnType 
               Caption         =   "&Data"
               Height          =   255
               Index           =   0
               Left            =   200
               TabIndex        =   1
               Tag             =   "0"
               Top             =   300
               Value           =   -1  'True
               Width           =   1110
            End
         End
         Begin VB.Frame fraDataType 
            Caption         =   "Data Type :"
            Height          =   870
            Left            =   200
            TabIndex        =   103
            Top             =   1440
            Width           =   7815
            Begin VB.ComboBox cboDataType 
               Height          =   315
               ItemData        =   "frmColEdit.frx":02CD
               Left            =   195
               List            =   "frmColEdit.frx":02CF
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   300
               Width           =   1815
            End
            Begin COASpinner.COA_Spinner asrSize 
               Height          =   315
               Left            =   2625
               TabIndex        =   6
               Top             =   300
               Width           =   870
               _ExtentX        =   1535
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
               MaximumValue    =   8000
               Text            =   "1"
            End
            Begin COASpinner.COA_Spinner asrDecimals 
               Height          =   315
               Left            =   4590
               TabIndex        =   7
               Top             =   300
               Width           =   735
               _ExtentX        =   1296
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
               MaximumValue    =   6
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnDefaultDisplayWidth 
               Height          =   315
               Left            =   6465
               TabIndex        =   8
               Top             =   300
               Width           =   735
               _ExtentX        =   1296
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
               MaximumValue    =   8000
               MinimumValue    =   1
               Text            =   "1"
            End
            Begin VB.Label lblDefaultDisplayWidth 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Display :"
               Height          =   195
               Left            =   5595
               TabIndex        =   134
               Top             =   360
               Width           =   795
            End
            Begin VB.Label lblSize 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Size :"
               Height          =   195
               Left            =   2085
               TabIndex        =   105
               Top             =   360
               Width           =   480
            End
            Begin VB.Label lblDecimals 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Decimals :"
               Height          =   195
               Left            =   3645
               TabIndex        =   104
               Top             =   360
               Width           =   900
            End
         End
         Begin VB.Frame fraLookup 
            Caption         =   "Lookup :"
            Height          =   3000
            Left            =   200
            TabIndex        =   100
            Top             =   2445
            Width           =   7815
            Begin VB.CheckBox chkAutoUpdateRecords 
               Caption         =   "Auto U&pdate Records"
               Height          =   195
               Left            =   5100
               TabIndex        =   11
               Top             =   1305
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.ComboBox cboLookupFilterOperator 
               Height          =   315
               Left            =   1740
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   2125
               Width           =   5520
            End
            Begin VB.ComboBox cboLookupFilterValue 
               Height          =   315
               Left            =   1740
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   2530
               Width           =   5520
            End
            Begin VB.CheckBox chkLookupFilter 
               Caption         =   "Filter &Lookup Values"
               Height          =   195
               Left            =   180
               TabIndex        =   12
               Top             =   1305
               Width           =   3555
            End
            Begin VB.ComboBox cboLookupFilterColumn 
               Height          =   315
               Left            =   1740
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1720
               Width           =   5520
            End
            Begin VB.ComboBox cboLookupTables 
               Height          =   315
               Left            =   1035
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   300
               Width           =   6225
            End
            Begin VB.ComboBox cboLookupColumns 
               Height          =   315
               Left            =   1035
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   700
               Width           =   6225
            End
            Begin VB.Label Label1 
               Caption         =   "Filter Operator :"
               Height          =   270
               Left            =   180
               TabIndex        =   15
               Top             =   2190
               Width           =   1425
            End
            Begin VB.Label txtLookupFilterValue 
               Caption         =   "Filter Value :"
               Height          =   270
               Left            =   180
               TabIndex        =   17
               Top             =   2580
               Width           =   1260
            End
            Begin VB.Label txtLookupFilterField 
               Caption         =   "Filter Column :"
               Height          =   285
               Left            =   180
               TabIndex        =   13
               Top             =   1785
               Width           =   1395
            End
            Begin VB.Label lblTable 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Table :"
               Height          =   195
               Left            =   200
               TabIndex        =   102
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblColumn 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Column :"
               Height          =   195
               Left            =   195
               TabIndex        =   101
               Top             =   765
               Width           =   630
            End
         End
         Begin VB.Frame fraLink 
            Caption         =   "Link to parent table :"
            Height          =   1650
            Left            =   180
            TabIndex        =   116
            Top             =   2445
            Width           =   7815
            Begin VB.ComboBox cboLinkViews 
               Height          =   315
               Left            =   1740
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   700
               Width           =   5550
            End
            Begin VB.TextBox txtLinkOrder 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1740
               Locked          =   -1  'True
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   1100
               Width           =   5235
            End
            Begin VB.CommandButton cmdLinkOrder 
               Caption         =   "..."
               Height          =   315
               Left            =   6960
               TabIndex        =   25
               Top             =   1100
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.ComboBox cboLinkTables 
               Height          =   315
               Left            =   1740
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   300
               Width           =   5550
            End
            Begin VB.Label lblLinkView 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Default View :"
               Height          =   195
               Left            =   195
               TabIndex        =   133
               Top             =   765
               Width           =   1365
            End
            Begin VB.Label lblLinkOrder 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Default Order :"
               Height          =   195
               Left            =   195
               TabIndex        =   118
               Top             =   1155
               Width           =   1455
            End
            Begin VB.Label lblLinkTable 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Table :"
               Height          =   195
               Left            =   195
               TabIndex        =   117
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame fraCalculation 
            Caption         =   "Calculation :"
            Height          =   1155
            Left            =   180
            TabIndex        =   113
            Top             =   2445
            Width           =   7815
            Begin VB.CheckBox chkCalculateIfEmpty 
               Caption         =   "Calculate only if e&mpty"
               Height          =   375
               Left            =   200
               TabIndex        =   21
               Top             =   720
               Width           =   3195
            End
            Begin VB.TextBox txtCalculation 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   200
               TabIndex        =   19
               Text            =   "txtCalculation"
               Top             =   300
               Width           =   6765
            End
            Begin VB.CommandButton cmdCalculation 
               Caption         =   "..."
               Height          =   315
               Left            =   6960
               TabIndex        =   20
               Top             =   300
               UseMaskColor    =   -1  'True
               Width           =   315
            End
         End
         Begin VB.Label lblColumnName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   195
            TabIndex        =   107
            Top             =   255
            Width           =   510
         End
      End
   End
   Begin COAColourPicker.COA_ColourPicker COA_ColourPicker1 
      Left            =   165
      Top             =   5955
      _ExtentX        =   820
      _ExtentY        =   820
      ShowSysColorButton=   0   'False
   End
End
Attribute VB_Name = "frmColEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Column definition variables.
Private mobjColumn As Column
Private miColumnType As ColumnTypes
Private miDataType As DataTypes
Private mlngLinkTableID As Long
Private mlngLinkViewID As Long
Private mlngLookupTableID As Long
Private mlngLookupColumnID As Long
Private mlngLookupFilterColumnID As Long
Private mlngLookupFilterColumnType As DataTypes
Private miLookupFilterOperator As FilterOperators
Private mlngLookupFilterValueID As Long
Private mlngCalcExprID As Long
Private miControlType As ControlTypes
Private miConvertCase As Integer
Private miAlignment As Integer
Private mlngValidationExprID As Long
Private msDefault As String
Private mlngDfltValueExprID As Long
Private mlngLinkOrderID As Long
Private miParentCount As Integer
Private miTrimming As Integer
Private mvarEmailLinks As Collection
Private mfClearDefault As Boolean

' Flag to see if any changes have been made by the user
Private mblnChanged As Boolean

' Form handling variables.
Private mfCancelled As Boolean
Private mfIsSaved As Boolean
Private mfReading As Boolean
Private mfLoading As Boolean

' Page number constants.
Const iPAGE_DEFINITION = 0
Const iPAGE_CONTROL = 1
Const iPAGE_OPTIONS = 2
Const iPAGE_VALIDATION = 3
Const iPAGE_DIARY = 4
Const iPAGE_EMAIL = 5
Const iPAGE_AFD = 6
Const iPAGE_QADDRESS = 7

Private mblnReadOnly As Boolean

Private Enum FilterOperators
  giFILTEROP_UNDEFINED = 0
  giFILTEROP_EQUALS = 1
  giFILTEROP_NOTEQUALTO = 2
  giFILTEROP_ISATMOST = 3
  giFILTEROP_ISATLEAST = 4
  giFILTEROP_ISMORETHAN = 5
  giFILTEROP_ISLESSTHAN = 6
  giFILTEROP_ON = 7
  giFILTEROP_NOTON = 8
  giFILTEROP_AFTER = 9
  giFILTEROP_BEFORE = 10
  giFILTEROP_ONORAFTER = 11
  giFILTEROP_ONORBEFORE = 12
  giFILTEROP_CONTAINS = 13
  giFILTEROP_IS = 14
  giFILTEROP_DOESNOTCONTAIN = 15
  giFILTEROP_ISNOT = 16
End Enum

Public Property Get Cancelled() As Boolean
  ' Return the cancelled property.
  Cancelled = mfCancelled
  
End Property

Public Property Let Changed(pblnNewValue As Boolean)
  mblnChanged = pblnNewValue
  cmdOK.Enabled = mblnChanged
End Property

Public Property Get Changed() As Boolean
  ' RH 22/08/00 - Dont write changes to the DB if no changes have been made!
  ' Return the (data) changed property
  Changed = mblnChanged
End Property

Public Property Set Column(pcColumnObject As Column)
  ' Set the column object.
  Set mobjColumn = pcColumnObject
  
  'Read the column properties and initialize the form controls.
  ReadColumnProperties
  
  'if mobjcolumn.TableID
  
  
  
End Property


Private Sub ASRDate1_Change()
  If Not mfLoading Then Changed = True

End Sub

Private Sub ASRDate1_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    ASRDate1.DateValue = Date
  End If

End Sub

Private Sub ASRDate1_LostFocus()

  'MH20020424 Fault 3760
  'Dunno what Hokey Kokey thing Tim is going on about below.
  '"In, Out, In, Out, Shake it all about..."
  '(Anyway, avoid date automatically changing 01/13/2002 to 13/01/2002)

'  If IsNull(ASRDate1.DateValue) And Not _
'    IsDate(ASRDate1.DateValue) And _
'    ASRDate1.Text <> "  /  /" Then
'
'    MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'    ASRDate1.DateValue = Null
'    tabColProps.Tab = iPAGE_OPTIONS
'
'    'TM20011030 Fault 3040 - Commented out the SetFocus.
'    'TM20011121 Fault 3151 - Uncommented the SetFocus. Sorry, split personality kicking in???
'    ASRDate1.SetFocus
'
'    Exit Sub
'  End If
  ValidateGTMaskDate ASRDate1

End Sub

Private Sub asrDecimals_Change()

  If Not mfLoading Then Changed = True

End Sub

Private Sub asrDefault_Change()
  If Not mfLoading Then Changed = True
End Sub

Private Sub ASRDefaultWorkingPattern_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub asrIncVal_Change()
  If Not mfLoading Then Changed = True
 
End Sub

Private Sub asrMaxOLESize_Change()
  If Not mfLoading Then Changed = True
End Sub

Private Sub asrMaxVal_Change()
  
  If asrMaxVal.value < asrDefault.value Then
    asrDefault.value = asrMaxVal.value
  End If
  
  If asrMaxVal.value < asrMinVal.value Then
    asrMinVal.value = asrMaxVal.value
  End If
  
  ' JDM - Fault 3549 - 17/06/02 - Set here not in incval_change as it was causing a loop
  asrIncVal.MinimumValue = IIf(asrMaxVal.value = asrMinVal.value, 0, 1)
  asrIncVal.MaximumValue = asrMaxVal.value - asrMinVal.value
  asrIncVal.value = IIf(asrIncVal.value > asrIncVal.MaximumValue, asrIncVal.MaximumValue, IIf(asrIncVal.value < asrIncVal.MinimumValue, asrIncVal.MinimumValue, asrIncVal.value))
  
  If Not mfLoading Then
    Changed = True
  End If
  
End Sub

Private Sub asrMinVal_Change()
  
  If asrMinVal.value > asrDefault.value Then
    asrDefault.value = asrMinVal.value
  End If
  
  If asrMinVal.value > asrMaxVal.value Then
    asrMaxVal.value = asrMinVal.value
  End If

  ' JDM - Fault 3549 - 17/06/02 - Set here not in incval_change as it was causing a loop
  asrIncVal.MinimumValue = IIf(asrMaxVal.value = asrMinVal.value, 0, 1)
  asrIncVal.MaximumValue = asrMaxVal.value - asrMinVal.value
  asrIncVal.value = IIf(asrIncVal.value > asrIncVal.MaximumValue, asrIncVal.MaximumValue, IIf(asrIncVal.value < asrIncVal.MinimumValue, asrIncVal.MinimumValue, asrIncVal.value))
  
  If Not mfLoading Then
    Changed = True
  End If
  
End Sub

Private Sub asrSize_Change()

  If Not mfLoading Then Changed = True

  'MH20010202 fault 1787
  'Do not allow decimals to exceed size
  If miDataType = dtNUMERIC Then
    'Increase the maximum decimal value to 6 as part of the
    'implementation of the Currency Conversion module.
    asrDecimals.MaximumValue = IIf(asrSize.value < 6, asrSize.value, 6)
    If asrDecimals.value > asrDecimals.MaximumValue Then
      asrDecimals.value = asrDecimals.MaximumValue
    End If
  End If

  ' Update the default display width
  If (miColumnType = giCOLUMNTYPE_CALCULATED) Or (miColumnType = giCOLUMNTYPE_DATA) Then
    Select Case miDataType
      Case dtNUMERIC, dtVARCHAR
        spnDefaultDisplayWidth.value = asrSize.value
      Case dtINTEGER, dtTIMESTAMP
        spnDefaultDisplayWidth.value = 10
      Case dtBIT
        spnDefaultDisplayWidth.value = 1
      Case dtLONGVARCHAR
        spnDefaultDisplayWidth.value = 14
      Case dtVARBINARY
        spnDefaultDisplayWidth.value = 255
      Case dtLONGVARBINARY
        spnDefaultDisplayWidth.value = 255
    End Select
  End If
  
End Sub

Private Sub cboAFDAddress_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboCase_Click()
  'If Not mfReading Then
    ' Update the Convert Case global variable.
    miConvertCase = cboCase.ItemData(cboCase.ListIndex)
  'End If
  
  If Not mfLoading Then Changed = True

End Sub


Private Sub cboAFDCounty_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboDefault_Click()
  If Not mfReading Then
    ' Update the global variable.
    
    'JPD 20041115 Fault 8970
    If miDataType = dtTIMESTAMP Then
      If cboDefault.Text <> "<None>" Then
        msDefault = UI.ConvertDateLocaleToSQL(cboDefault.Text)
      Else
        msDefault = ""
      End If
    Else
      msDefault = cboDefault.Text
    End If
  End If
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDForename_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDInitial_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboLinkTables_Click()
  
  If Not mfReading Then
    If mlngLinkTableID <> cboLinkTables.ItemData(cboLinkTables.ListIndex) Then
      mlngLinkOrderID = 0
      txtLinkOrder.Text = ""
    End If

    ' Set the linktable.
    mlngLinkTableID = cboLinkTables.ItemData(cboLinkTables.ListIndex)
    
    miDataType = dtBINARY
    cboDataType_Refresh
  End If
  
  cboLinkViews_Initialize
  If Not mfLoading Then Changed = True
  
End Sub


Private Sub cboLinkViews_Click()
  
  With cboLinkViews
    If .ListIndex <> -1 Then
      mlngLinkViewID = .ItemData(.ListIndex)
    End If
  End With

  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDLocality_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboLookupFilterColumn_Click()

  If Not mfReading Then
    ' Set the lookup column.
    mlngLookupFilterColumnID = cboLookupFilterColumn.ItemData(cboLookupFilterColumn.ListIndex)
      
    ' Get the column data type, etc.
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", mlngLookupFilterColumnID
  
      If .NoMatch Then
        mlngLookupFilterColumnType = sqlUnknown
        cboLookupFilterValue_Refresh
      Else
        Do While Not .EOF
          If .Fields("tableID") <> mlngLookupTableID Then
            Exit Do
          End If
          
          If (.Fields("columnid") = mlngLookupFilterColumnID) Then
            If .Fields("DataType") <> mlngLookupFilterColumnType Then
              mlngLookupFilterColumnType = .Fields("DataType")
              cboLookupFilterValue_Refresh
            End If
            
            Exit Do
          End If
          .MoveNext
        Loop
      End If
    End With
  
  End If
  
  If Not mfLoading Then Changed = True
  
  cboLookupFilterOperator_Refresh
  
End Sub

Private Sub cboLookupFilterOperator_Refresh()
  ' Populate the Filter Operator combo.
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  With cboLookupFilterOperator
    ' Clear the combo.
    .Clear
      
    Select Case mlngLookupFilterColumnType
      Case sqlOle  ' Not required as OLEs are not permitted in the Lookup Filter Column selection.
      
      Case sqlBoolean ' Logic columns.
        .AddItem OperatorDescription(giFILTEROP_EQUALS)
        .ItemData(.NewIndex) = giFILTEROP_EQUALS

      Case sqlNumeric, sqlInteger ' Numeric and Integer columns.
        .AddItem OperatorDescription(giFILTEROP_EQUALS)
        .ItemData(.NewIndex) = giFILTEROP_EQUALS
        .AddItem OperatorDescription(giFILTEROP_NOTEQUALTO)
        .ItemData(.NewIndex) = giFILTEROP_NOTEQUALTO
        .AddItem OperatorDescription(giFILTEROP_ISMORETHAN)
        .ItemData(.NewIndex) = giFILTEROP_ISMORETHAN
        .AddItem OperatorDescription(giFILTEROP_ISATLEAST)
        .ItemData(.NewIndex) = giFILTEROP_ISATLEAST
        .AddItem OperatorDescription(giFILTEROP_ISLESSTHAN)
        .ItemData(.NewIndex) = giFILTEROP_ISLESSTHAN
        .AddItem OperatorDescription(giFILTEROP_ISATMOST)
        .ItemData(.NewIndex) = giFILTEROP_ISATMOST

      Case sqlDate ' Date columns.
        .AddItem OperatorDescription(giFILTEROP_ON)
        .ItemData(.NewIndex) = giFILTEROP_ON
        .AddItem OperatorDescription(giFILTEROP_NOTON)
        .ItemData(.NewIndex) = giFILTEROP_NOTON
        .AddItem OperatorDescription(giFILTEROP_AFTER)
        .ItemData(.NewIndex) = giFILTEROP_AFTER
        .AddItem OperatorDescription(giFILTEROP_BEFORE)
        .ItemData(.NewIndex) = giFILTEROP_BEFORE
        .AddItem OperatorDescription(giFILTEROP_ONORAFTER)
        .ItemData(.NewIndex) = giFILTEROP_ONORAFTER
        .AddItem OperatorDescription(giFILTEROP_ONORBEFORE)
        .ItemData(.NewIndex) = giFILTEROP_ONORBEFORE

      Case sqlVarChar, sqlLongVarChar, sqlVarBinary  ' Character and Photo columns (photo columns are really character columns).
        .AddItem OperatorDescription(giFILTEROP_IS)
        .ItemData(.NewIndex) = giFILTEROP_IS
        .AddItem OperatorDescription(giFILTEROP_ISNOT)
        .ItemData(.NewIndex) = giFILTEROP_ISNOT
        .AddItem OperatorDescription(giFILTEROP_CONTAINS)
        .ItemData(.NewIndex) = giFILTEROP_CONTAINS
        .AddItem OperatorDescription(giFILTEROP_DOESNOTCONTAIN)
        .ItemData(.NewIndex) = giFILTEROP_DOESNOTCONTAIN
    End Select

    .Enabled = (.ListCount > 0)
    
    If .ListCount > 0 Then
      iIndex = 0
      
      For iLoop = 0 To .ListCount - 1
        If .ItemData(iLoop) = miLookupFilterOperator Then
          iIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      .ListIndex = iIndex
    End If
  End With

End Sub


Private Function OperatorDescription(piOperatorCode As Integer) As String
  ' Return the textual description og the given operator.
  Dim sDesc As String
  
  Select Case piOperatorCode
    Case giFILTEROP_EQUALS
      sDesc = "is equal to"
    Case giFILTEROP_NOTEQUALTO
      sDesc = "is NOT equal to"
    Case giFILTEROP_ISATMOST
      sDesc = "is less than or equal to"
    Case giFILTEROP_ISATLEAST
      sDesc = "is greater than or equal to"
    Case giFILTEROP_ISMORETHAN
      sDesc = "is greater than"
    Case giFILTEROP_ISLESSTHAN
      sDesc = "is less than"
    Case giFILTEROP_ON
      sDesc = "is equal to"
    Case giFILTEROP_NOTON
      sDesc = "is NOT equal to"
    Case giFILTEROP_AFTER
      sDesc = "after"
    Case giFILTEROP_BEFORE
      sDesc = "before"
    Case giFILTEROP_ONORAFTER
      sDesc = "is equal to or after"
    Case giFILTEROP_ONORBEFORE
      sDesc = "is equal to or before"
    Case giFILTEROP_CONTAINS
      sDesc = "contains"
    Case giFILTEROP_IS
      sDesc = "is equal to"
    Case giFILTEROP_DOESNOTCONTAIN
      sDesc = "does not contain"
    Case giFILTEROP_ISNOT
      sDesc = "is NOT equal to"
    Case Else
      sDesc = ""
  End Select
  
  OperatorDescription = sDesc
  
End Function


Private Sub cboLookupFilterOperator_Click()
  If Not mfReading Then
    ' Set the lookup column.
    miLookupFilterOperator = cboLookupFilterOperator.ItemData(cboLookupFilterOperator.ListIndex)
    Changed = True
  End If

End Sub


Private Sub cboLookupFilterValue_Click()
  
  If Not mfReading Then
    ' Set the lookup column.
    mlngLookupFilterValueID = cboLookupFilterValue.ItemData(cboLookupFilterValue.ListIndex)
    Changed = True
  End If

End Sub

Private Sub cboLookupTables_Click()
  Dim lngOldTableID As Long
  
  If Not mfReading Then
    lngOldTableID = mlngLookupTableID
    
    ' Set the lookup table global variable.
    mlngLookupTableID = cboLookupTables.ItemData(cboLookupTables.ListIndex)
      
    ' If the table has changed then reset the lookup column id.
    If lngOldTableID <> mlngLookupTableID Then
      mlngLookupColumnID = 0
    End If
    
    ' Refresh the lookup columns combo.
    cboLookupColumns_Refresh
    cboLookupFilterColumn_Refresh
  End If
  
  If Not mfLoading Then Changed = True
  
End Sub

Private Sub cboAFDProperty_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDStreet_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDSurname_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDTelephone_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboQAAddress_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboQACounty_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboQALocality_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboQAProperty_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboQAStreet_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboQATown_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboTextAlignment_Click()
  'If Not mfReading Then
    ' Update the Alignment global variable.
    miAlignment = cboTextAlignment.ItemData(cboTextAlignment.ListIndex)
  'End If
  
  If Not mfLoading Then Changed = True

End Sub

Private Sub cboAFDTown_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cboTrimming_Click()

'  If Not mfReading Then
    ' Update the Digit grouping global variable.
    miTrimming = cboTrimming.ListIndex
'  End If
  If Not mfLoading Then Changed = True

End Sub



Private Sub chkAudit_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub chkAutoUpdateRecords_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub chkCalculateIfEmpty_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub chkChildUnique_Click()
  If Not mfReading Then
    ' Force the column to be mandatory also.
    If chkChildUnique.value Then
      'TM05052004 - Unique Not Mandatory
      'chkMandatory.Value = vbChecked
      chkUnique.value = vbUnchecked
    End If
  
    RefreshValidationTab
  End If
  If Not mfLoading Then Changed = True

End Sub

Private Sub chkDuplicate_Click()
  If Not mfLoading Then Changed = True

End Sub

Private Sub chkEnableOLEMaxSize_Click()
  
  Dim bAllowEnableEmbed As Boolean
  
  If Not mfLoading Then
    
    If mfReading Then
      bAllowEnableEmbed = True
    Else
      If chkEnableOLEMaxSize.value = vbChecked Then
        bAllowEnableEmbed = MsgBox("Enabling document embedding can affect database performance." & vbCrLf & "Are you sure you want to enable this option?", vbQuestion & vbYesNo, "Embedding") = vbYes
        chkEnableOLEMaxSize.value = IIf(bAllowEnableEmbed = True, vbChecked, vbUnchecked)
      Else
        bAllowEnableEmbed = True
      End If
    End If
    
    If bAllowEnableEmbed Then
      lblMaximumOLESize.Enabled = chkEnableOLEMaxSize.value = vbChecked
      asrMaxOLESize.Enabled = chkEnableOLEMaxSize.value = vbChecked
      asrMaxOLESize.value = IIf(chkEnableOLEMaxSize.value = vbChecked, asrMaxOLESize.value, 100)
      asrMaxOLESize.BackColor = IIf(asrMaxOLESize.Enabled, vbWindowBackground, vbButtonFace)
      lblMb.Enabled = chkEnableOLEMaxSize.value = vbChecked
      Changed = True
    End If
  End If

End Sub

Private Sub chkMandatory_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub chkMultiLine_Click()
'  ' Reset the Mask checkbox as Multiline/Mask are mutually exclusive.
'  If Not mfReading Then
'    If chkMultiLine.Value Then
'      txtMask.Text = ""
'      txtMask.Enabled = False
'    Else
'      txtMask.Enabled = Not mblnReadOnly
'    End If
'  End If

  If miDataType = dtVARCHAR And Not miColumnType = giCOLUMNTYPE_LOOKUP Then
    If chkMultiLine.value = vbChecked Then
      asrSize.Text = VARCHAR_MAX_Size
    Else
      asrSize.Text = 8000
    End If
  End If

  spnDefaultDisplayWidth.Text = asrSize.Text
  txtDefault.MaxLength = Minimum(IIf(chkMultiLine.value = vbChecked, 8000, val(asrSize.Text)), 8000)

  If Not mfLoading Then Changed = True
End Sub

Private Sub chkAFDPostCodeColumn_Click()
  If Not mfReading Then
    'Enable/Disable relevant fields
    AfdToggleControlStatus chkAFDPostCodeColumn.value
  End If
  If Not mfLoading Then Changed = True

End Sub

Private Sub chkQAPostCodeColumn_Click()

  If Not mfReading Then
    'Enable/Disable relevant fields
    QAToggleControlStatus chkQAPostCodeColumn.value
  End If
  If Not mfLoading Then Changed = True

End Sub

'Private Sub chkReadOnly_Click()
'  ' Clear the validation expression details if the column is readonly.
'  If Not mfReading Then
'    If chkReadOnly.Value Then
'      mlngValidationExprID = 0
'      GetValidationExpressionDetails
'      With txtStatusBarMessage
'        .Enabled = False
'        .Text = ""
'        .BackColor = vbButtonFace
'      End With
'    Else
'      With txtStatusBarMessage
'        '.Enabled = True
'        .Enabled = Not mblnReadOnly
'        .BackColor = vbWindowBackground
'      End With
'    End If
'  Else
'    If chkReadOnly.Value Then
'      With txtStatusBarMessage
'        .Enabled = False
'        .Text = ""
'        .BackColor = vbButtonFace
'      End With
'    Else
'      With txtStatusBarMessage
'        '.Enabled = True
'        .Enabled = Not mblnReadOnly
'        .BackColor = vbWindowBackground
'      End With
'    End If
'  End If
'
'  If Not mfLoading Then Changed = True
'
'End Sub

'MH20060928 Fault 11527
Private Sub chkReadOnly_Click()

  ' Clear the validation expression details if the column is readonly.
  If chkReadOnly.value And Not mfReading Then
    mlngValidationExprID = 0
    GetValidationExpressionDetails
  End If
  
  With txtStatusBarMessage
    .Enabled = (chkReadOnly.value = vbUnchecked And Not mblnReadOnly)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With
  
  If Not mfLoading Then Changed = True

End Sub

Private Sub chkUnique_Click()
  If Not mfReading Then
    ' Force the column to be mandatory also.
    If chkUnique.value Then
      'TM05052004 - Unique Not Mandatory
      'chkMandatory.Value = vbChecked
      chkChildUnique.value = vbUnchecked
    End If
    
    RefreshValidationTab
  End If
  If Not mfLoading Then Changed = True
  
End Sub

Private Sub chkUse1000Separator_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub chkZeroBlank_Click()
  If Not mfLoading Then Changed = True
End Sub

Private Sub cmdAddDiaryLink_Click()
  ' Define a new Diary Link.
  Dim frmDiary As frmDiaryLink
  Dim sDfltComment As String
  Dim sOffset As String
  'Dim sSuffix As String
  'Dim iOffset As Integer
  'Dim sBeforeAfter As String
  
  ' Create a default comment for the diary link.
  sDfltComment = ""
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", IIf(IsNull(mobjColumn.Properties("tableID")), 0, mobjColumn.Properties("tableID"))

    If Not .NoMatch Then
      sDfltComment = .Fields("tableName")
    End If
  End With
  sDfltComment = sDfltComment & IIf(Len(sDfltComment) > 0, ".", "")
  sDfltComment = sDfltComment & Trim(txtColumnName.Text)
  
  Set frmDiary = New frmDiaryLink
  
  With frmDiary
    ' Initialise the diary link.
    .TableID = mobjColumn.TableID
    .DiaryComment = sDfltComment
    .DiaryOffset = "0"
    .DiaryPeriod = iTimePeriodDays
    .DiaryReminder = False
    .FilterID = 0
    '.EffectiveDate = #1/1/1980#
    .EffectiveDate = Date
    .CheckLeavingDate = True

    ' Display the diary link form.
    .Show vbModal
    
    ' Read the new diary link.
    If Not .Cancelled Then
      ' Read the diary link information from the form.
'      iOffset = frmDiary.DiaryOffset
'
'      If iOffset = 0 Then
'        sOffset = "No offset"
'      Else
'        If iOffset < 0 Then
'          sBeforeAfter = " before"
'          sSuffix = IIf(iOffset = -1, "", "s")
'          sOffset = Trim(Str(iOffset * -1))
'        Else
'          sBeforeAfter = " after"
'          sSuffix = IIf(iOffset = 1, "", "s")
'          sOffset = Trim(Str(iOffset))
'        End If
'
'        Select Case frmDiary.DiaryPeriod
'          Case iTimePeriodDays
'            sOffset = sOffset & " day" & sSuffix & sBeforeAfter
'          Case iTimePeriodMonths
'            sOffset = sOffset & " month" & sSuffix & sBeforeAfter
'          Case iTimePeriodWeeks
'            sOffset = sOffset & " week" & sSuffix & sBeforeAfter
'          Case iTimePeriodYears
'            sOffset = sOffset & " year" & sSuffix & sBeforeAfter
'        End Select
      
        sOffset = GetOffset(frmDiary.DiaryOffset, frmDiary.DiaryPeriod, False)
      
'      End If
          
      ' Add the diary link to the grid.
      ssGrdDiaryLinks.AddItem frmDiary.DiaryComment & _
        vbTab & sOffset & _
        vbTab & frmDiary.DiaryReminder & _
        vbTab & frmDiary.DiaryOffset & _
        vbTab & frmDiary.DiaryPeriod & _
        vbTab & frmDiary.FilterID & _
        vbTab & frmDiary.EffectiveDate & _
        vbTab & frmDiary.CheckLeavingDate
        'vbTab & Format(frmDiary.EffectiveDate, "mm/dd/yyyy")

      ' Select the new row.
      'ssGrdDiaryLinks.Bookmark = (ssGrdDiaryLinks.Rows -1)
      With ssGrdDiaryLinks
        .Bookmark = .AddItemBookmark(.Rows - 1)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With

      ' Refesh the diary link page controls.
      RefreshDiaryLinkTab
      Changed = True

    End If
  End With
  
  ' Disassociate object variables.
  Set frmDiary = Nothing

End Sub

'Private Sub cmdAddEmailLink_Click()
'
'  Dim frmEmail As frmEmailLink
'  Dim objNewLink As clsEmailLink
'  Dim strKey As String
'  'Dim iLoop As Integer
'
'  Set frmEmail = New frmEmailLink
'  Set objNewLink = New clsEmailLink
'
'
'  'Set up defaults for new link
'  With objNewLink
'
'    .LinkID = .GetNewLinkID(mvarEmailLinks)
'    .ColumnID = mobjColumn.ColumnID
'
'    .Title = vbNullString
'    .FilterID = 0
'    .Offset = 0
'    .OffsetPeriod = iTimePeriodDays
'    .EffectiveDate = Date
'
'    .Subject = vbNullString
'    '.Importance = 1
'    '.Sensitivity = 0
'    .IncRecordDesc = True
'    .IncColumnDetails = True
'    .IncUsername = True
'
'    .Text = vbNullString
'    .Attachment = vbNullString
'
'  End With
'
'  frmEmail.TableID = mobjColumn.TableID
'  frmEmail.EmailLink = objNewLink
'  frmEmail.EmailLink.Recipients = objNewLink.Recipients
'  frmEmail.AllowOffset = (miDataType = dtTIMESTAMP)
'  frmEmail.PopulateControls
'  frmEmail.Show vbModal
'
'  If Not frmEmail.Cancelled Then
'
'    Set objNewLink = frmEmail.EmailLink
'
'    'With objNewLink.Recipients
'      strKey = "ID" & objNewLink.LinkID
'      mvarEmailLinks.Add objNewLink, strKey
'      'For iLoop = 1 To objNewLink.Recipients.Count
'      '  mvarEmailLinks.Item(strKey).Recipients.Add objNewLink.Recipients(iLoop)
'      'Next
'
'
'    With ssGrdEmailLinks
'
'      .AddItem _
'        objNewLink.Title & vbTab & _
'        GetOffset(objNewLink.Offset, objNewLink.OffsetPeriod, objNewLink.Immediate) & vbTab & _
'        objNewLink.Subject & vbTab & _
'        objNewLink.LinkID
'
'      '.Bookmark = (.Rows - 1)
'      .Bookmark = .AddItemBookmark(.Rows - 1)
'      .SelBookmarks.RemoveAll
'      .SelBookmarks.Add .Bookmark
'
'    End With
'
'    RefreshEmailLinkTab
'    Changed = True
'
'  End If
'
'  Set objNewLink = Nothing
'
'  UnLoad frmEmail
'  Set frmEmail = Nothing
'
'
'
''  ' Define a new email Link.
''  Dim frmEmail As frmEmailLink
''  Dim sDfltComment As String
''  Dim sOffset As String
''
''  ' Create a default comment for the email link.
''  sDfltComment = ""
''  With recTabEdit
''    .Index = "idxTableID"
''    .Seek "=", IIf(IsNull(mobjColumn.Properties("tableID")), 0, mobjColumn.Properties("tableID"))
''
''    If Not .NoMatch Then
''      sDfltComment = .Fields("tableName")
''    End If
''  End With
''  sDfltComment = sDfltComment & IIf(Len(sDfltComment) > 0, ".", "")
''  sDfltComment = sDfltComment & Trim(txtColumnName.Text)
''
''  Set frmEmail = New frmEmailLink
''
''  With frmEmail
''    ' Initialise the email link.
''    .emailComment = sDfltComment
''    .emailOffset = "0"
''    .emailPeriod = iTimePeriodDays
''    .emailReminder = False
''
''    ' Display the email link form.
''    .Show vbModal
''
''    ' Read the new email link.
''    If Not .Cancelled Then
''      sOffset = GetOffset(frmEmail.emailOffset, frmEmail.emailPeriod)
''
''      ' Add the email link to the grid.
''      ssGrdEmailLinks.AddItem frmEmail.emailComment & _
''        vbTab & sOffset & _
''        vbTab & frmEmail.emailReminder & _
''        vbTab & frmEmail.emailOffset & _
''        vbTab & frmEmail.emailPeriod
''
''      ' Select the new row.
''      ssGrdEmailLinks.Bookmark = (ssGrdEmailLinks.Rows - 1)
''      ssGrdEmailLinks.SelBookmarks.Add ssGrdEmailLinks.Bookmark
''
''      ' Refesh the email link page controls.
''      RefreshEmailLinkTab
''    End If
''  End With
''
''  ' Disassociate object variables.
''  Set frmEmail = Nothing
''
'End Sub

Private Sub cmdCalculation_Click()
  
  Dim objExpr As CExpression
  Dim fDataTypeChanged As Boolean
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    Select Case miDataType
      Case dtVARCHAR
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_CHARACTER
      Case dtTIMESTAMP
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_DATE
      Case dtLONGVARBINARY
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_OLE
      Case dtVARBINARY
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_PHOTO
      Case dtINTEGER
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_NUMERIC
      Case dtBIT
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_LOGIC
      Case dtNUMERIC
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_NUMERIC
      Case dtLONGVARCHAR
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_CHARACTER
      Case Else
        .Initialise mobjColumn.TableID, mlngCalcExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_UNDEFINED
    End Select
    
    .CalculatedColumnID = mobjColumn.ColumnID
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      
      mlngCalcExprID = .ExpressionID
      ' Read the selected expression info.
      GetCalculationExpressionDetails
      If Not mfLoading Then Changed = True
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", mlngCalcExprID, False

        If .NoMatch Then
          ' Read the selected expression info.
          mlngCalcExprID = 0
          GetCalculationExpressionDetails
          If Not mfLoading Then Changed = True
        End If
        
        ' RH 24/01/01 - BUG 1566.  Only clear the selection if the user has
        ' changed the datatype of the calculation.
        If mlngCalcExprID > 0 Then
          Select Case miDataType
            Case dtNUMERIC, dtINTEGER
              fDataTypeChanged = (.Fields("returntype").value <> giEXPRVALUE_NUMERIC)
            Case dtVARCHAR
              fDataTypeChanged = (.Fields("returntype").value <> giEXPRVALUE_CHARACTER)
            Case adDate, dtTIMESTAMP
              fDataTypeChanged = (.Fields("returntype").value <> giEXPRVALUE_DATE)
            Case dtBIT
              fDataTypeChanged = (.Fields("returntype").value <> giEXPRVALUE_LOGIC)
            'JDM - 13/09/01 - Fault 2364 - Error when cancelling on working pattern
            Case dtLONGVARCHAR
              fDataTypeChanged = (.Fields("returntype").value <> giEXPRVALUE_CHARACTER)
            Case Else
              MsgBox "Warning ! Unknown miDatatype !"
          End Select
          If fDataTypeChanged Then
            mlngCalcExprID = 0
            GetCalculationExpressionDetails
            If Not mfLoading Then Changed = True
          End If
        End If
      End With
    End If
  End With
  
  Set objExpr = Nothing
  
  'If Not mfLoading Then Changed =  True

End Sub

Private Sub cmdCancel_Click()
  Dim pintAnswer As Integer
    If Changed = True And cmdOK.Enabled Then
      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
      If pintAnswer = vbYes Then
        Me.MousePointer = vbHourglass
        cmdOK_Click 'This is just like saving
        Me.MousePointer = vbNormal
        Exit Sub
      ElseIf pintAnswer = vbCancel Then
        Exit Sub
      End If
    End If
TidyUpAndExit:
  UnLoad Me
End Sub


'Private Sub cmdDfltColourSelect_Click()
'
'  Dim lngPictureID As Long
'  Dim sFileName As String
'
'  frmPictSel.PictureType = vbPicTypeBitmap
'  frmPictSel.SelectedPicture = mlngDeskTopBitmapID
'  frmPictSel.ExcludedExtensions = ".gif"
'  frmPictSel.Show vbModal
'
'If (frmPictSel.SelectedPicture <> mlngDeskTopBitmapID) And (Not mblnLoading) Then Changed = True
'
'  If frmPictSel.SelectedPicture > 0 Then
'    With recPictEdit
'      .Index = "idxID"
'      .Seek "=", frmPictSel.SelectedPicture
'      If Not .NoMatch Then
'        mlngDeskTopBitmapID = !PictureID
'        txtDeskTopBitmapName.Text = !Name
'        cboBitmapLocation.Enabled = True
'        cmdPictureClear.Enabled = True
'        sFileName = ReadPicture
'        picWork.Picture = LoadPicture(sFileName)
'        picWork.Move 0, 0, picHolder.ScaleWidth, picHolder.ScaleHeight
'        SizeImage picWork
'        picWork.Top = (picHolder.ScaleHeight - picWork.Height) \ 2
'        picWork.Left = (picHolder.ScaleWidth - picWork.Width) \ 2
'        Kill sFileName
'        picHolder.Visible = True
'      Else
'        cboBitmapLocation.Enabled = False
'        cmdPictureClear.Enabled = False
'        picHolder.Visible = False
'      End If
'    End With
'
'  End If
'
'End Sub

Private Sub selDefaultColour_Click()

  On Error GoTo ErrorTrap

  With COA_ColourPicker1
    .Color = selDefaultColour.BackColor
    .ShowPalette
    If selDefaultColour.BackColor <> .Color Then
      selDefaultColour.BackColor = .Color
      Changed = True
    End If
  End With

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub cmdDfltValueExpression_Click()
  Dim objExpr As CExpression

  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    Select Case miDataType
      Case dtVARCHAR
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_CHARACTER
      Case dtTIMESTAMP
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_DATE
      Case dtLONGVARBINARY
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_OLE
      Case dtVARBINARY
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_PHOTO
      Case dtINTEGER
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_NUMERIC
      Case dtBIT
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_LOGIC
      Case dtNUMERIC
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_NUMERIC
      Case dtLONGVARCHAR
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_CHARACTER
      Case Else
        .Initialise mobjColumn.TableID, mlngDfltValueExprID, giEXPR_DEFAULTVALUE, giEXPRVALUE_UNDEFINED
    End Select
    
    .CalculatedColumnID = mobjColumn.ColumnID
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      mlngDfltValueExprID = .ExpressionID
        
      ' Read the selected expression info.
      GetDfltValueExpressionDetails
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", mlngDfltValueExprID, False

        If .NoMatch Then
          ' Read the selected expression info.
          mlngDfltValueExprID = 0
          GetDfltValueExpressionDetails
        End If
      End With
    End If
  End With
  
  Set objExpr = Nothing
  If Not mfLoading Then Changed = True

  ' Refresh the current tab page
  cboCase.Enabled = False
  RefreshCurrentTab
  
End Sub

Private Sub cmdDiaryLinkProperties_Click()
  ' Define a new Diary Link.
  Dim frmDiary As frmDiaryLink
  Dim sOffset As String
  Dim sSuffix As String
  Dim iOffset As Integer
  Dim sBeforeAfter As String
    
  Dim strRow As String
  Dim lngRow As Long
    
  Set frmDiary = New frmDiaryLink
  
  With frmDiary
    ' Read the existing diary link information from the grid.
    .TableID = mobjColumn.TableID
    .DiaryComment = ssGrdDiaryLinks.Columns(0).value
    .DiaryReminder = ssGrdDiaryLinks.Columns(2).value
    .DiaryOffset = ssGrdDiaryLinks.Columns(3).value
    .DiaryPeriod = ssGrdDiaryLinks.Columns(4).value
    .FilterID = ssGrdDiaryLinks.Columns(5).value
    .EffectiveDate = ssGrdDiaryLinks.Columns(6).value
    .CheckLeavingDate = ssGrdDiaryLinks.Columns(7).value

    ' Display the diary link form.
    .Show vbModal
    
    ' Read the new diary link.
    If Not .Cancelled Then
      ' Read the diary link information from the form.
      iOffset = frmDiary.DiaryOffset
      
      If iOffset = 0 Then
        sOffset = "No offset"
      Else
        If iOffset < 0 Then
          sBeforeAfter = " before"
          sSuffix = IIf(iOffset = -1, "", "s")
          sOffset = Trim(Str(iOffset * -1))
        Else
          sBeforeAfter = " after"
          sSuffix = IIf(iOffset = 1, "", "s")
          sOffset = Trim(Str(iOffset))
        End If

        'Select Case frmDiary.DiaryPeriod
        '  Case iTimePeriodDays
        '    sOffset = sOffset & " day" & sSuffix & sBeforeAfter
        '  Case iTimePeriodMonths
        '    sOffset = sOffset & " month" & sSuffix & sBeforeAfter
        '  Case iTimePeriodWeeks
        '    sOffset = sOffset & " week" & sSuffix & sBeforeAfter
        '  Case iTimePeriodYears
        '    sOffset = sOffset & " year" & sSuffix & sBeforeAfter
        'End Select
        sOffset = sOffset & " " & _
            TimePeriod(frmDiary.DiaryPeriod) & _
            sSuffix & sBeforeAfter

      End If

      'MH20060419 Fault 10843
      ''' Add the diary link to the grid.
      ''ssGrdDiaryLinks.Columns(0).Value = .DiaryComment
      ''ssGrdDiaryLinks.Columns(1).Value = sOffset
      ''ssGrdDiaryLinks.Columns(2).Value = .DiaryReminder
      ''ssGrdDiaryLinks.Columns(3).Value = .DiaryOffset
      ''ssGrdDiaryLinks.Columns(4).Value = .DiaryPeriod
      ''ssGrdDiaryLinks.Columns(5).Value = .FilterID
      ''ssGrdDiaryLinks.Columns(6).Value = .EffectiveDate
      ''ssGrdDiaryLinks.Columns(7).Value = .CheckLeavingDate
      strRow = .DiaryComment & vbTab & _
               sOffset & vbTab & _
               .DiaryReminder & vbTab & _
               .DiaryOffset & vbTab & _
               .DiaryPeriod & vbTab & _
               .FilterID & vbTab & _
               .EffectiveDate & vbTab & _
               .CheckLeavingDate

      With ssGrdDiaryLinks
        lngRow = .AddItemRowIndex(.Bookmark)

        .RemoveItem lngRow
        .AddItem strRow, lngRow
        .Bookmark = .AddItemBookmark(lngRow)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With

      ' Refesh the diary link page controls.
      RefreshDiaryLinkTab
      Changed = True
    End If
  End With
  
  ' Disassociate object variables.
  Set frmDiary = Nothing

End Sub


'Private Sub cmdEmailLinkProperties_Click()
'
'  Dim frmEmail As frmEmailLink
'  Dim objNewLink As clsEmailLink
'  Dim lngOldLinkID As Long
'
'  Dim strRow As String
'  Dim lngRow As Long
'
'  Set frmEmail = New frmEmailLink
'  Set objNewLink = New clsEmailLink
'
'  On Error GoTo LocalErr
'
'  'Get existing object
'  Set objNewLink = mvarEmailLinks.Item("ID" & ssGrdEmailLinks.Columns(3).value)
'  lngOldLinkID = objNewLink.LinkID
'
'  Load frmEmail   'Required!
'  frmEmail.TableID = mobjColumn.TableID
'  frmEmail.EmailLink = objNewLink
'  frmEmail.AllowOffset = (miDataType = dtTIMESTAMP)
'  frmEmail.PopulateControls
'  frmEmail.Show vbModal
'
'  If Not frmEmail.Cancelled Then
'
'    mvarEmailLinks.Remove "ID" & lngOldLinkID
'    Set objNewLink = frmEmail.EmailLink
'    mvarEmailLinks.Add objNewLink, "ID" & objNewLink.LinkID
'
'    strRow = objNewLink.Title & vbTab & _
'             GetOffset(objNewLink.Offset, objNewLink.OffsetPeriod, objNewLink.Immediate) & vbTab & _
'             objNewLink.Subject & vbTab & _
'             CStr(objNewLink.LinkID)
'
'    With ssGrdEmailLinks
'      lngRow = .AddItemRowIndex(.Bookmark)
'
'      .RemoveItem lngRow
'      .AddItem strRow, lngRow
'      .Bookmark = .AddItemBookmark(lngRow)
'      .SelBookmarks.RemoveAll
'      .SelBookmarks.Add .Bookmark
'    End With
'
'    RefreshEmailLinkTab
'    Changed = True
'
'  End If
'
'  Set objNewLink = Nothing
'
'  UnLoad frmEmail
'  Set frmEmail = Nothing
'
'Exit Sub
'
'LocalErr:
'  If ASRDEVELOPMENT Then
'    MsgBox Err.Description, vbCritical, "ASR DEVELOPMENT"
'    Stop
'  End If
'
'End Sub

Private Sub cmdLinkOrder_Click()
  ' Display the Order selection form.
  Dim objOrder As Order
  
  ' Create a new order object.
  Set objOrder = New Order
  
  ' Initialize the order object.
  With objOrder
    .OrderID = mlngLinkOrderID
    .TableID = mlngLinkTableID
    .OrderType = giORDERTYPE_DYNAMIC

    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      mlngLinkOrderID = .OrderID
    Else
      ' Check in case the original order has been deleted.
      With recOrdEdit
        .Index = "idxID"
        .Seek "=", mlngLinkOrderID

        If .NoMatch Then
          mlngLinkOrderID = 0
        Else
          If !Deleted Then
            mlngLinkOrderID = 0
          End If
        End If
      End With
    End If
  End With
  
  GetLinkOrderDetails
  
  ' Disassociate object variables.
  Set objOrder = Nothing

  If Not mfLoading Then Changed = True

End Sub

Private Sub cmdOK_Click()
  
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If ValidateGTMaskDate(ASRDate1) = False Then
    Exit Sub
  End If
  
  
  ' Write the control values into the Column object for writing to the database.
  On Error GoTo ErrorTrap
    
  ' RH 22/08/00 - Start of mblnChanged condition
  If Changed = True Then
    
    Dim iLoop As Integer
    'Dim iLoop2 As Integer
    Dim iIndex As Integer
    Dim lngColumnID As Long
    Dim lngScreenIDs() As Long
    'Dim lngExprParentTableID As Long
    Dim dblMaxValue As Double
    Dim sSQL As String
    'Dim sExprName As String
    'Dim sExprType As String
    Dim sDefault As String
    Dim sValues As String
    Dim sScreens() As String
    Dim sSubString As String
    'Dim sTableName As String
    Dim sColumnName As String
    'Dim sOtherColumnName As String
    'Dim sExprParentTable As String
    Dim objMisc As Misc
    'Dim objExpr As CExpression
    Dim frmControlChange As frmControlChange
    Dim objDiaryLink As cDiaryLink
    Dim objDiaryLinks As Collection
    Dim vNull As Variant
    'Dim vValidatedDate As Variant
    Dim vScreenList As Variant
    Dim rsScreens As DAO.Recordset
    Dim rsOrders As DAO.Recordset
    'Dim rsExpressions As dao.Recordset
    Dim rsOtherColumns As DAO.Recordset
    'Dim objEmailLink As clsEmailLink
    Dim strKey As String
    Dim fOneSelected As Boolean
    Dim fAllSelected As Boolean
    Dim lngParentTableID As Long
    
    Dim rsDAOTemp As DAO.Recordset
    
    'Dim sOtherColumnID As String
    'Dim sTableID As String
    Dim iNewIndex As Integer
    Dim sMsgBoxText As String
    Dim sOtherCols() As String
    Dim objOtherCol As SystemMgr.Column
    
    Dim rsChangedViews As DAO.Recordset   'MH20010320
    Dim rsChangedExprs As DAO.Recordset
    Dim iCount As Integer
    Dim iChangeStatus As Integer
    Dim sChangeStatus As String
      
    Dim psColourUsage As String
    
    ' Get the column name.
    sColumnName = Trim(txtColumnName.Text)
    
    ' Check that a column name has been entered.
    If Len(Trim(sColumnName)) < 1 Then
      MsgBox "A column name must be entered.", vbOKOnly + vbExclamation, Application.Name
      tabColProps.Tab = iPAGE_DEFINITION
      txtColumnName.SetFocus
      Exit Sub
    End If
    
    ' Ensure that the column name is not a database keyword.
    If UCase(sColumnName) = "ID" Or IsKeyword(sColumnName) _
      Or (UCase(Left(sColumnName, 3)) = "ID_" And val(Mid(sColumnName, 4)) > 0) Then
      ' Flag to the user that the column name is a database keyword.
      MsgBox "'" & sColumnName & "' is a reserved word" & vbCr & _
        "and cannot be used for a column name.", _
        vbOKOnly + vbExclamation, Application.Name
      tabColProps.Tab = iPAGE_DEFINITION
      txtColumnName.SetFocus
      Exit Sub
    End If
    
    ' Ensure that the column name is unique.
    With recColEdit
      If Not (.BOF And .EOF) Then
        ' Seek columns table for a column with this name.
        .Index = "idxName"
        .Seek "=", mobjColumn.TableID, sColumnName, False
          
        If Not .NoMatch Then
          If (Not mfIsSaved) Or _
            (.Fields("columnID") <> mobjColumn.ColumnID) Then
            ' Flag to the user if there already exists a column with this name.
            MsgBox "A column named '" & sColumnName & "' already exists!", _
              vbOKOnly + vbExclamation, Application.Name
            tabColProps.Tab = iPAGE_DEFINITION
            txtColumnName.SetFocus
            Exit Sub
          End If
        End If
      End If
    End With
    
    If Me.optColumnType(2).value Then
      If (LCase(cboDataType.Text) = "ole object") Or (LCase(cboDataType.Text) = "photo") Then
        MsgBox cboDataType.Text & " data types are invalid for calculated column types.", vbExclamation + vbOKOnly, App.Title
        Exit Sub
      End If
    End If
    
    
    'MH20010119 Fault 1598
    'This section of code was remmed out but it is required.
    'This code needs to be here for when you select a lookup table
    'which has no columns on it!
    
    ' Ensure that lookup details are specified for lookup columns.
    If miColumnType = giCOLUMNTYPE_LOOKUP Then
      If Not (mlngLookupTableID > 0 And mlngLookupColumnID > 0) Then
        ' Flag to the user that lookup details need to be specified.
        MsgBox "Lookup details incomplete", _
          vbOKOnly + vbExclamation, Application.Name
        tabColProps.Tab = iPAGE_DEFINITION
        cboLookupTables.SetFocus
        Exit Sub
      End If
      
      If chkLookupFilter.value = vbChecked Then
        If cboLookupFilterColumn.ListIndex = 0 Then
          MsgBox "Lookup details incomplete", vbOKOnly + vbExclamation, Application.Name
          tabColProps.Tab = iPAGE_DEFINITION
          cboLookupFilterColumn.SetFocus
          Exit Sub
        End If
        
        If cboLookupFilterValue.ListIndex = 0 Then
          MsgBox "Lookup details incomplete", vbOKOnly + vbExclamation, Application.Name
          tabColProps.Tab = iPAGE_DEFINITION
          cboLookupFilterValue.SetFocus
          Exit Sub
        End If
      End If
      
    End If
  
  '  ' Ensure that a calculation is selected for calculated columns.
  '  If miColumnType = giCOLUMNTYPE_CALCULATED Then
  '     If Not (mlngCalcExprID > 0) Then
  '      ' Flag to the user that lookup details need to be specified.
  '      MsgBox "No calculation selected.", _
  '        vbOKOnly + vbExclamation, Application.Name
  '        tabColProps.Tab = iPAGE_DEFINITION
  '        cmdCalculation.SetFocus
  '      Exit Sub
  '    End If
  '  End If
    
    ' Flag to the user if the column size is invalid.
    If ColumnHasSize(miDataType) And (val(asrSize.Text) < 1) Then
      MsgBox "Invalid column size.", vbOKOnly + vbExclamation, Application.Name
      tabColProps.Tab = iPAGE_DEFINITION
      If asrSize.Enabled Then
        asrSize.SetFocus
      ElseIf cmdCalculation.Enabled Then
        cmdCalculation.SetFocus
      End If
      Exit Sub
    End If
    
    
    ' JPD - check that a parent is selected if the user has
    ' chosen the column to be unique within sibling records.
    If (chkChildUnique.value = vbChecked) And (miParentCount > 1) Then
      fOneSelected = False
      For iLoop = 0 To lstUniqueParents.ListCount - 1
        If lstUniqueParents.Selected(iLoop) Then
          fOneSelected = True
          Exit For
        End If
      Next iLoop
      
      If Not fOneSelected Then
        MsgBox "At least one parent table must be selected if the column is to be unique (within sibling records).", vbOKOnly + vbExclamation, Application.Name
        tabColProps.Tab = iPAGE_VALIDATION
        lstUniqueParents.SetFocus
        Exit Sub
      End If
    End If
    
    ' Validate the Control Values and default values.
    Select Case miDataType
      Case dtVARCHAR
        ' Check that the Control Values are valid.
        
        'JPD 20050623 Fault 10176
        If ((miColumnType = giCOLUMNTYPE_DATA) And _
          ((miControlType = giCTRL_COMBOBOX) Or (miControlType = giCTRL_OPTIONGROUP))) Then
        
          sValues = Trim(txtListValues.Text)
          While Len(sValues) > 0
            iIndex = InStr(txtListValues.Text, vbCr & vbLf)
            If iIndex > 0 Then
              sSubString = Left(sValues, iIndex - 1)
              sValues = Mid(sValues, iIndex + 2)
            Else
              sSubString = sValues
              sValues = ""
            End If
              
            If Len(sSubString) > val(asrSize.Text) Then
              MsgBox "The control values are too long for the defined column size.", _
                vbOKOnly + vbExclamation, Application.Name
              tabColProps.Tab = iPAGE_CONTROL
              txtListValues.SetFocus
              Exit Sub
            End If
          Wend
        
          'JPD 20030901 Fault 6897
          If (miControlType = giCTRL_OPTIONGROUP) And (Len(Trim(txtListValues.Text)) < 1) Then
            MsgBox "There are no control values defined for the option group.", _
              vbOKOnly + vbExclamation, Application.Name
            tabColProps.Tab = iPAGE_CONTROL
            txtListValues.SetFocus
            Exit Sub
          End If
      
          'JPD 20030902 Fault 4918
          If (miColumnType <> giCOLUMNTYPE_LOOKUP) And (miControlType = giCTRL_COMBOBOX) And (Len(Trim(txtListValues.Text)) < 1) Then
            MsgBox "There are no control values defined for the dropdown list.", _
              vbOKOnly + vbExclamation, Application.Name
            tabColProps.Tab = iPAGE_CONTROL
            txtListValues.SetFocus
            Exit Sub
          End If
        End If
        
        'MH20040213 Fault 8088
        If txtDefault.Visible Then
          ' Check that the Default Value is valid.
          If Len(Trim(txtDefault.Text)) > val(asrSize.Text) Then
            MsgBox "The default value is too long for the defined column size.", _
              vbOKOnly + vbExclamation, Application.Name
            tabColProps.Tab = iPAGE_OPTIONS
            txtDefault.SetFocus
            Exit Sub
          End If
        End If
        
      Case dtNUMERIC
        ' Check that the Default Value is valid.
  '''      dblMaxValue = 10 ^ (Val(asrSize.Text) - Val(asrDecimals.Text))
  '''      If (Val(txtDefault.Text) >= dblMaxValue) Or _
  '''        (Val(txtDefault.Text) <= (-1 * dblMaxValue)) Then
  '''        MsgBox "The default value exceeds the defined column size and decimal places.", _
  '''          vbOKOnly + vbExclamation, Application.Name
  '''        tabColProps.Tab = iPAGE_OPTIONS
  '''        txtDefault.SetFocus
  '''        Exit Sub
  '''      End If
  '''      If InStr(1, Trim(txtDefault.Text), ".") > 0 Then
  '''        If Len(Mid(Trim(txtDefault.Text), InStr(1, Trim(txtDefault.Text), ".") + 1)) > Val(asrDecimals.Text) Then
  '''          MsgBox "The default value exceeds the defined column decimal places.", _
  '''            vbOKOnly + vbExclamation, Application.Name
  '''          tabColProps.Tab = iPAGE_OPTIONS
  '''          txtDefault.SetFocus
  '''          Exit Sub
  '''        End If
  '''      End If
    
      Case dtINTEGER
        dblMaxValue = (2 ^ 31) - 1
        If TDBDefaultNumber.value > dblMaxValue Then
          MsgBox "The default value exceeds the maximum of " & CStr(dblMaxValue) & " allowed for an integer column.", vbExclamation, Application.Name
          tabColProps.Tab = iPAGE_OPTIONS
          TDBDefaultNumber.SetFocus
          Exit Sub
        End If

        ' JDM - Fault 2826 - Check for lowest value
        dblMaxValue = -(2 ^ 31)
        If TDBDefaultNumber.value < dblMaxValue Then
          MsgBox "The default value exceeds the minimum of " & CStr(dblMaxValue) & " allowed for an integer column.", vbExclamation, Application.Name
          tabColProps.Tab = iPAGE_OPTIONS
          TDBDefaultNumber.SetFocus
          Exit Sub
        End If

    End Select
    
    ' Validate the default date if there is one.
    'If (miDataType = dtTIMESTAMP) And _
    '  (Len(ASRDate1.Text) <> 0) Then
    '  Set objMisc = New MISC
    '  vValidatedDate = objMisc.ValidateDate(ASRDate1.FormattedText)
    '  Set objMisc = Nothing
    '
    '  If VarType(vValidatedDate) <> vbDate Then
    '    MsgBox "Invalid default date.", vbOKOnly, Application.Name
    '    tabColProps.Tab = iPAGE_OPTIONS
    '    ASRDate1.SetFocus
    '    Exit Sub
    '  End If
    'End If
  
    If mobjColumn.ColumnID > 0 Then
      
      ' Validate any change of datatype.
      If mobjColumn.Properties("dataType") <> miDataType Then
        
        Dim mfrmUse As frmUsage
        Set mfrmUse = New frmUsage
        mfrmUse.ResetList
        If mobjColumn.ColumnIsUsed(mfrmUse) Then
          Screen.MousePointer = vbDefault
          mfrmUse.ShowMessage GetTableName(mobjColumn.TableID) & "." & mobjColumn.Properties("ColumnName").value & " Column", "The data type cannot be changed as the column is used in the following : ", UsageCheckObject.Column
          tabColProps.Tab = 0
          If Me.cboDataType.Enabled Then Me.cboDataType.SetFocus
          Exit Sub
        End If
        UnLoad mfrmUse
        Set mfrmUse = Nothing
  
      End If
      
      ' Check if the column type is a Linked/Embedded OLE/Photo column and has had the embedded option changed.
      ' This option might be required if used in a Workflow Stored Data element.
      If ((miDataType = dtLONGVARBINARY) _
          Or (miDataType = dtVARBINARY)) _
        And (mobjColumn.Properties("OLEType") = 2) Then
        
        If (Not chkEnableOLEMaxSize.value) _
          And (CBool(mobjColumn.Properties("maxOLESizeEnabled"))) Then

          ' Check if the column is used in any Workflow StoredData elements,
          ' as a column updated by a Linked/Embedded OLE/Photo column with the embedded option enabled.
          
          sSQL = "SELECT tmpWorkflows.name, " & _
            "   tmpWorkflowElements.caption" & _
            " FROM tmpWorkflowElementColumns DC," & _
            "   tmpColumns SC," & _
            "   tmpWorkflowElements," & _
            "   tmpWorkflows" & _
            " WHERE DC.columnID = " & CStr(mobjColumn.ColumnID) & _
            "   AND DC.valueType = 2" & _
            "   AND DC.dbColumnID = SC.columnID" & _
            "   AND SC.maxOLESizeEnabled = TRUE" & _
            "   AND DC.elementID = tmpWorkflowElements.ID" & _
            "   AND tmpWorkflowElements.workflowID = tmpWorkflows.ID" & _
            "   AND tmpWorkflows.deleted = FALSE"
          Set rsOtherColumns = daoDb.OpenRecordset(sSQL, _
            dbOpenForwardOnly, dbReadOnly)
  
          If Not (rsOtherColumns.BOF And rsOtherColumns.EOF) Then
            iLoop = 0
            sMsgBoxText = ""
            
            Do Until rsOtherColumns.EOF
              iLoop = iLoop + 1
              sMsgBoxText = sMsgBoxText & vbTab & "Workflow : " & rsOtherColumns!Name & " <'" & rsOtherColumns!Caption & "' stored data element>" & vbNewLine
              
              rsOtherColumns.MoveNext
            Loop
  
            ' Set the first bit of the msgbox text
            sMsgBoxText = "Document embedding cannot be disabled as it may invalidate the following Workflow Stored Data element" & IIf(iLoop = 1, "", "s") & " :" & vbNewLine & vbNewLine _
              & sMsgBoxText
            
              MsgBox sMsgBoxText, vbOKOnly + vbExclamation, Application.Name
  
            tabColProps.Tab = 2
            If Me.chkEnableOLEMaxSize.Enabled Then Me.chkEnableOLEMaxSize.SetFocus
            Exit Sub
          End If
          
        ElseIf (chkEnableOLEMaxSize.value) _
          And (Not CBool(mobjColumn.Properties("maxOLESizeEnabled"))) Then
          ' Check if the column is used in any Workflow StoredData elements,
          ' as a column used to update a Linked/Embedded OLE/Photo column with the embedded option disabled.

          sSQL = "SELECT tmpWorkflows.name, " & _
            "   tmpWorkflowElements.caption" & _
            " FROM tmpWorkflowElementColumns DC," & _
            "   tmpColumns SC," & _
            "   tmpWorkflowElements," & _
            "   tmpWorkflows" & _
            " WHERE DC.dbColumnID = " & CStr(mobjColumn.ColumnID) & _
            "   AND DC.valueType = 2" & _
            "   AND DC.columnID = SC.columnID" & _
            "   AND SC.maxOLESizeEnabled = FALSE" & _
            "   AND DC.elementID = tmpWorkflowElements.ID" & _
            "   AND tmpWorkflowElements.workflowID = tmpWorkflows.ID" & _
            "   AND tmpWorkflows.deleted = FALSE"
          Set rsOtherColumns = daoDb.OpenRecordset(sSQL, _
            dbOpenForwardOnly, dbReadOnly)

          If Not (rsOtherColumns.BOF And rsOtherColumns.EOF) Then
            iLoop = 0
            sMsgBoxText = ""

            Do Until rsOtherColumns.EOF
              iLoop = iLoop + 1
              sMsgBoxText = sMsgBoxText & vbTab & "Workflow : " & rsOtherColumns!Name & " <'" & rsOtherColumns!Caption & "' stored data element>" & vbNewLine

              rsOtherColumns.MoveNext
            Loop

            ' Set the first bit of the msgbox text
            sMsgBoxText = "Document embedding cannot be enabled as it may invalidate the following Workflow Stored Data element" & IIf(iLoop = 1, "", "s") & " :" & vbNewLine & vbNewLine _
              & sMsgBoxText

              MsgBox sMsgBoxText, vbOKOnly + vbExclamation, Application.Name

            tabColProps.Tab = 2
            If Me.chkEnableOLEMaxSize.Enabled Then Me.chkEnableOLEMaxSize.SetFocus
            Exit Sub
          End If
        End If
      End If
      
      ' Check if the size/decimals has changed.
      iChangeStatus = 0
      sChangeStatus = ""
      
      If (val(asrSize.Text) <> mobjColumn.Properties("size")) And _
          (val(asrDecimals.Text) <> mobjColumn.Properties("decimals")) Then
        iChangeStatus = 1
        sChangeStatus = "size or decimals"
      ElseIf (val(asrSize.Text) <> mobjColumn.Properties("size")) Then
        iChangeStatus = 2
        sChangeStatus = "size"
      ElseIf (val(asrDecimals.Text) <> mobjColumn.Properties("decimals")) Then
        iChangeStatus = 3
        sChangeStatus = "decimals"
      End If
      
      ' JPD 2/12/99
      ' Flag to the user if the column size/decimals is invalid.
      If iChangeStatus > 0 Then
       
        ' Check if the column is used as a lookup reference by any other columns.
        sSQL = "SELECT tmpColumns.TableID, tmpColumns.columnID, tmpColumns.columnName, tmpTables.tableName" & _
          " FROM tmpColumns, tmpTables" & _
          " WHERE tmpTables.deleted = FALSE" & _
          " AND tmpColumns.deleted = FALSE" & _
          " AND tmpColumns.tableID = tmpTables.tableID" & _
          " AND tmpColumns.lookupColumnID = " & Trim(Str(mobjColumn.ColumnID)) & _
          " AND tmpColumns.columnType = " & Trim(Str(giCOLUMNTYPE_LOOKUP))
        Set rsOtherColumns = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
            
        If Not (rsOtherColumns.BOF And rsOtherColumns.EOF) Then
          
          ' Initialise the array which stores the columns which will need autochanging
          ' 0 - Table Name
          ' 1 - Column Name
          ' 2 - Column ID
          ' 3 - Table ID
          ReDim sOtherCols(3, 0)
          
          Do Until rsOtherColumns.EOF
          
            ' Load the information into the array
            
            iNewIndex = UBound(sOtherCols, 2) + 1
            ReDim Preserve sOtherCols(3, iNewIndex)
            
            sOtherCols(0, iNewIndex) = rsOtherColumns.Fields("tableName")
            sOtherCols(1, iNewIndex) = rsOtherColumns.Fields("columnName")
            sOtherCols(2, iNewIndex) = rsOtherColumns.Fields("ColumnID")
            sOtherCols(3, iNewIndex) = rsOtherColumns.Fields("TableID")
                
            rsOtherColumns.MoveNext
                  
          Loop
        
          ' Set the first bit of the msgbox text
          sMsgBoxText = "This column is used by the following lookup columns :" & vbCrLf & vbCrLf
          
          ' Loop thru the array, adding the table.column information
          For iNewIndex = 1 To UBound(sOtherCols, 2)
            sMsgBoxText = sMsgBoxText & vbTab & sOtherCols(0, iNewIndex) & "." & sOtherCols(1, iNewIndex) & vbCrLf
          Next iNewIndex
          
          ' Set the final bit of the msgbox text
          ' JDM - 21/08/01 - Fault 2547 - Change the text of the below message
          sMsgBoxText = sMsgBoxText & vbCrLf & "The " & sChangeStatus & " of these columns will automatically be changed." & vbCrLf & "Do you wish to continue ?"
                        
          ' Show the msgbox, asking yes/no whether or not to continue
          If MsgBox(sMsgBoxText, vbYesNo + vbQuestion, Application.Name) = vbYes Then
            
            ' We ARE going to change the other cols
            For iNewIndex = 1 To UBound(sOtherCols, 2)
              
              Set objOtherCol = New SystemMgr.Column
              objOtherCol.TableID = sOtherCols(3, iNewIndex)
              objOtherCol.ColumnID = sOtherCols(2, iNewIndex)
              
              If objOtherCol.ReadColumn Then
                objOtherCol.Properties("size") = val(asrSize.Text)
                objOtherCol.Properties("decimals") = val(asrDecimals.Text)
                objOtherCol.Properties("defaultdisplaywidth") = val(spnDefaultDisplayWidth.value)
                objOtherCol.WriteColumn_Transaction
              End If
            
            Next iNewIndex
                
            Set objOtherCol = Nothing
            
          Else
            ' We ARENT going to change the other cols
            Exit Sub
          End If
        
        End If
      End If
    End If
    
    ' Check, but don't force the user, if a calclation has been defined for Calculated columns.
    If (miColumnType = giCOLUMNTYPE_CALCULATED) And mlngCalcExprID <= 0 Then
      MsgBox "There is no calculation defined for this calculated column." & vbCrLf & _
        "Please ensure one is defined before saving changes to the server.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
    
    If mobjColumn.ColumnID > 0 Then
      ' Validate any change of controlType.
      If mobjColumn.Properties("controlType") <> miControlType Then
        
        ' Check if this Integer column is a ColourPicker control type. If it is check if used in any charts
        ' and bounce if it is.
        If mobjColumn.Properties("dataType") = dtINTEGER And mobjColumn.Properties("controlType") = 2 ^ 15 Then
            ' Check that it is not used in SSI Charting.
            sSQL = "SELECT DISTINCT tmpSSIntranetLinks.ID," & _
              "   tmpSSIntranetLinks.Element_Type," & _
              "   tmpSSIntranetLinks.text" & _
              " FROM tmpSSIntranetLinks" & _
              " WHERE tmpSSIntranetLinks.Chart_ColourID = " & Trim(Str(mobjColumn.ColumnID))
          
            Set rsDAOTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
              psColourUsage = "The control type cannot be changed as the column is used in" & vbCrLf & _
                              "the following :" & vbCrLf & vbCrLf
              Do Until rsDAOTemp.EOF
                    psColourUsage = psColourUsage & "Self Service Intranet Chart : " & rsDAOTemp.Fields("text") & vbCrLf
                rsDAOTemp.MoveNext
              Loop
              
              MsgBox psColourUsage, vbOKOnly + vbExclamation, Application.Name
              'Close temporary recordset
              rsDAOTemp.Close
              Exit Sub
            Else
              'Close temporary recordset
              rsDAOTemp.Close
            End If
        End If
      
        ReDim sScreens(0)
        ReDim lngScreenIDs(0)
        ' Check if the control is used in any screens.
        sSQL = "SELECT DISTINCT tmpScreens.name, tmpScreens.screenID" & _
          " FROM tmpScreens, tmpControls" & _
          " WHERE tmpControls.columnID=" & Trim(Str(mobjColumn.ColumnID)) & _
          " AND tmpControls.screenID = tmpScreens.screenID" & _
          " AND tmpScreens.deleted=FALSE"
        Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        With rsScreens
          ' For each parent table definition ...
          Do While (Not .EOF)
            iIndex = UBound(sScreens) + 1
            ReDim Preserve sScreens(iIndex)
            ReDim Preserve lngScreenIDs(iIndex)
            sScreens(iIndex) = .Fields("name")
            lngScreenIDs(iIndex) = .Fields("screenID")
            .MoveNext
          Loop
          .Close
        End With
        Set rsScreens = Nothing
        
        ' If the column control is already used in a screen then ask the user if they
        ' want to delete or update these instances.
        If UBound(sScreens) > 0 Then
          Set frmControlChange = New frmControlChange
          With frmControlChange
            vScreenList = sScreens
            .ScreenList = vScreenList
            .Show vbModal
          End With
          
          If frmControlChange.Cancelled Then
            Exit Sub
          End If
          
          ' Flag the screens that will be changed.
          For iIndex = 1 To UBound(lngScreenIDs)
            sSQL = "UPDATE tmpScreens" & _
              " SET changed = TRUE" & _
              " WHERE tmpScreens.screenID=" & Trim(Str(lngScreenIDs(iIndex)))
            daoDb.Execute sSQL
          Next iIndex
          
          If frmControlChange.DeleteControls Then
            ' Delete the controls that represent this column in any screens.
            sSQL = "DELETE FROM tmpControls" & _
              " WHERE columnID=" & Trim(Str(mobjColumn.ColumnID))
            daoDb.Execute sSQL
          Else
            
            Set objMisc = New Misc
            
            ' Change the controls that represent this column in any screens to the new type.
            sSQL = "UPDATE tmpControls SET controlType=" & Trim(Str(miControlType))
            
            Select Case miControlType
              Case giCTRL_OPTIONGROUP
                sSQL = sSQL & ", Caption = '" & Trim(objMisc.StrReplace(sColumnName, "_", " ", False)) & "'"
                ' AE20080418 Fault #10170
                'sSQL = sSQL & ", BackColor = " & Str(RGB(192, 192, 192))
                sSQL = sSQL & ", BackColor = " & Str(vbButtonFace)
              
              Case giCTRL_COMBOBOX
                ' AE20080418 Fault #10170
                'sSQL = sSQL & ", BackColor = " & Str(RGB(255, 255, 255))
                sSQL = sSQL & ", BackColor = " & Str(vbWindowBackground)
              
              Case giCTRL_TEXTBOX
                ' AE20080418 Fault #10170
                'sSQL = sSQL & ", BackColor = " & Str(RGB(255, 255, 255))
                sSQL = sSQL & ", BackColor = " & Str(vbWindowBackground)
            
              Case giCTRL_WORKINGPATTERN
                ' AE20080418 Fault #10170
                'sSQL = sSQL & ", BackColor = " & Str(RGB(192, 192, 192))
                sSQL = sSQL & ", BackColor = " & Str(vbButtonFace)
           
              End Select
              
            sSQL = sSQL & " WHERE columnID=" & Trim(Str(mobjColumn.ColumnID))
            daoDb.Execute sSQL
          End If
          
          ' Disassociate object variables.
          Set frmControlChange = Nothing
        End If
      End If

      'MH20010320
      'If column name has changed then mark any views which use this column as changed.
      If mobjColumn.Properties("columnName") <> sColumnName Then
        
        Application.ChangedColumnName = True
        
        sSQL = "SELECT DISTINCT ViewID FROM tmpViewColumns " & _
               "WHERE InView = True AND ColumnID = " & CStr(mobjColumn.ColumnID)
        Set rsChangedViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
        With rsChangedViews
          '.MoveFirst
          Do While Not rsChangedViews.EOF
            With recViewEdit
              .Index = "idxViewID"
              .Seek "=", rsChangedViews!ViewID
              If Not .NoMatch Then
                .Edit
                .Fields("Changed") = True
                .Update
              End If
            End With
            .MoveNext
          Loop
          .Close
        End With
        Set rsChangedViews = Nothing
      
      End If

    End If
    
    ' Update the properties of the associated column object.
    With mobjColumn
      ' Update Definition properties.
      .Properties("columnName") = sColumnName
      .Properties("columnType") = miColumnType
      .Properties("dataType") = miDataType
      
      ' RH 12/03/01 - Write the new default display width to the column object
      .Properties("defaultdisplaywidth") = val(Me.spnDefaultDisplayWidth.value)
      
      'MH20010301 Fault 1931
      'Overwrite the size and decimals for certain data types...
      '.Properties("size") = Val(asrSize.Text)
      '.Properties("decimals") = Val(asrDecimals.Text)
      Select Case LCase(cboDataType.Text)
      Case "integer"
        .Properties("size") = 10
        .Properties("decimals") = 0
      Case "numeric"
        .Properties("size") = val(asrSize.Text)
        .Properties("decimals") = val(asrDecimals.Text)
      Case "character"
        If chkMultiLine.value = vbChecked Then
          .Properties("size") = VARCHAR_MAX_Size
        Else
          .Properties("size") = val(asrSize.Text)
        End If
        .Properties("decimals") = 0
      Case Else
        .Properties("size") = 0
        .Properties("decimals") = 0
      End Select
      
      'NPG20080414 Suggestion S000441
      .Properties("CalculateIfEmpty") = chkCalculateIfEmpty
      
      .Properties("linkTableID") = IIf(miColumnType = giCOLUMNTYPE_LINK, mlngLinkTableID, 0)
      .Properties("linkViewID") = IIf(miColumnType = giCOLUMNTYPE_LINK, mlngLinkViewID, 0)
      .Properties("lookupTableID") = IIf(miColumnType = giCOLUMNTYPE_LOOKUP, mlngLookupTableID, 0)
      .Properties("lookupColumnID") = IIf(miColumnType = giCOLUMNTYPE_LOOKUP, mlngLookupColumnID, 0)
      
      .Properties("AutoUpdateLookupValues") = IIf(miColumnType = giCOLUMNTYPE_LOOKUP, Me.chkAutoUpdateRecords.value, 0)
      
      .Properties("LookupFilterValueID") = IIf(miColumnType = giCOLUMNTYPE_LOOKUP, mlngLookupFilterValueID, 0)
      .Properties("LookupFilterOperator") = IIf(miColumnType = giCOLUMNTYPE_LOOKUP, miLookupFilterOperator, 0)
      .Properties("LookupFilterColumnID") = IIf(miColumnType = giCOLUMNTYPE_LOOKUP, mlngLookupFilterColumnID, 0)
      
      .Properties("calcExprID") = IIf(miColumnType = giCOLUMNTYPE_CALCULATED, mlngCalcExprID, 0)
      .Properties("linkOrderID") = IIf(miColumnType = giCOLUMNTYPE_LINK, mlngLinkOrderID, 0)
      
      ' Update Control properties.
      'TM20030123 Fault 4957 - save the spinner values if it is a Calculated Column also.
      .Properties("controlType") = miControlType
      .ControlValuesString = IIf((miColumnType = giCOLUMNTYPE_DATA) And _
        ((miControlType = giCTRL_COMBOBOX) Or (miControlType = giCTRL_OPTIONGROUP)), _
        Trim(txtListValues.Text), "")
      .Properties("spinnerMinimum") = IIf((((miColumnType = giCOLUMNTYPE_DATA) Or _
                                            (miColumnType = giCOLUMNTYPE_CALCULATED)) _
                                            And (miControlType = giCTRL_SPINNER)), _
                                      val(asrMinVal.Text), 0)
      .Properties("spinnerMaximum") = IIf((((miColumnType = giCOLUMNTYPE_DATA) Or _
                                            (miColumnType = giCOLUMNTYPE_CALCULATED)) _
                                            And (miControlType = giCTRL_SPINNER)), _
                                      val(asrMaxVal.Text), 0)
      .Properties("spinnerIncrement") = IIf((((miColumnType = giCOLUMNTYPE_DATA) Or _
                                            (miColumnType = giCOLUMNTYPE_CALCULATED)) _
                                            And (miControlType = giCTRL_SPINNER)), _
                                      val(asrIncVal.Text), 0)
      
      'MH20060928 Fault 11527
      '.Properties("statusBarMessage") = Trim(txtStatusBarMessage.Text)
      .Properties("statusBarMessage") = IIf(chkReadOnly.value <> vbChecked, Trim(txtStatusBarMessage.Text), "")
      
      ' Update Options properties.
      .Properties("readOnly") = (chkReadOnly.value = vbChecked)
      .Properties("audit") = (chkAudit.value = vbChecked)
      
      ' OLE settings
      For iCount = optOLEStorageType.LBound To optOLEStorageType.UBound
        If optOLEStorageType(iCount).value = True Then
          .Properties("OLEType") = iCount
        End If
      Next iCount
      
      .Properties("MaxOLESizeEnabled") = (chkEnableOLEMaxSize.value = vbChecked)
      .Properties("MaxOLESize") = asrMaxOLESize.value
      
      '.Properties("multiLine") = IIf((miControlType = giCTRL_TEXTBOX Or miControlType = giCTRL_NAVIGATION) And _
        (miDataType = dtVARCHAR), _
        (chkMultiLine.value = vbChecked), False)
      .Properties("multiLine") = (miDataType = dtVARCHAR And chkMultiLine.value = vbChecked)
      
      .Properties("blankIfZero") = IIf((miControlType = giCTRL_TEXTBOX) And _
        (miDataType = dtNUMERIC Or miDataType = dtINTEGER), _
        (chkZeroBlank.value = vbChecked), False)
      
      
      .Properties("convertCase") = miConvertCase
      .Properties("alignment") = miAlignment
      .Properties("Trimming") = miTrimming
      .Properties("Use1000Separator") = (chkUse1000Separator.value = vbChecked)
      .Properties("mandatory") = (chkMandatory.value = vbChecked)
      
      ' Update Default properties.
      If optDfltType(1).value Then
        sDefault = ""
      Else
        Select Case miControlType
          Case giCTRL_TEXTBOX
            If miDataType = dtTIMESTAMP Then
              If IsDate(ASRDate1.Text) Then
                'JPD 20041112 Fault 8970
                'sDefault = Format(ASRDate1.Text, "mm/dd/yyyy")
                sDefault = UI.ConvertDateLocaleToSQL(ASRDate1.Text)
              Else
                sDefault = ""
              End If
            ElseIf ((miDataType = dtINTEGER) Or (miDataType = dtNUMERIC)) Then
              sDefault = Trim(Str(TDBDefaultNumber.value))
            Else
              sDefault = Trim(txtDefault.Text)
            End If
          Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
            If miDataType = dtTIMESTAMP Then
              'JPD 20041115 Fault 8970
              sDefault = IIf(cboDefault.Text = "<None>", "", UI.ConvertDateLocaleToSQL(cboDefault.Text))
            Else
              sDefault = IIf(cboDefault.Text = "<None>", "", Trim(cboDefault.Text))
            End If
          Case giCTRL_CHECKBOX
            sDefault = IIf(optDefault(0).value, "TRUE", "FALSE")
          Case giCTRL_SPINNER
            sDefault = Trim(asrDefault.Text)
          Case giCTRL_WORKINGPATTERN
            sDefault = ASRDefaultWorkingPattern.value
          Case giCTRL_NAVIGATION
            sDefault = Trim(txtDefault.Text)
          Case giCTRL_COLOURPICKER
            sDefault = CStr(selDefaultColour.BackColor)
          Case Else
            sDefault = vbNullString
        End Select
        
        mlngDfltValueExprID = 0
      End If
      .Properties("defaultValue") = sDefault
      .Properties("dfltValueExprID").value = IIf(mlngDfltValueExprID = -1, 0, mlngDfltValueExprID)
      
      ' Update Validation properties.
      .Properties("duplicate") = (chkDuplicate.value = vbChecked)
      .Properties("mandatory") = (chkMandatory.value = vbChecked)
      
      'JPD20010727
      '.Properties("uniqueCheck") = (chkUnique.Value = vbChecked)
      '.Properties("childUniqueCheck") = (chkChildUnique.Value = vbChecked)
      If (chkUnique.value = vbChecked) Then
        .Properties("uniqueCheckType") = giUNIQUECHECKTYPE_ENTIRE
      Else
        If (chkChildUnique.value = vbChecked) Then
          If miParentCount > 1 Then
            fAllSelected = True
            For iLoop = 0 To lstUniqueParents.ListCount - 1
              If Not lstUniqueParents.Selected(iLoop) Then
                fAllSelected = False
              Else
                lngParentTableID = lstUniqueParents.ItemData(iLoop)
              End If
            Next iLoop
            
            If fAllSelected Then
              .Properties("uniqueCheckType") = giUNIQUECHECKTYPE_SIBLINGSALL
            Else
              .Properties("uniqueCheckType") = lngParentTableID
            End If
          Else
            .Properties("uniqueCheckType") = giUNIQUECHECKTYPE_SIBLINGSALL
          End If
        Else
          .Properties("uniqueCheckType") = giUNIQUECHECKTYPE_NONE
        End If
      End If
      
      .Properties("mask") = IIf(Len(Trim(txtMask.Text)) > 0, txtMask.Text, vNull)
      
      .Properties("lostFocusExprID").value = IIf(mlngValidationExprID = -1, 0, mlngValidationExprID)
      .Properties("errorMessage") = Trim(txtErrorMessage.Text)
      
      ' Update Afd properties. RH 7/9/99
      If tabColProps.TabVisible(iPAGE_AFD) Then
        .Properties("Afdenabled") = IIf(chkAFDPostCodeColumn.value = vbChecked, 1, 0)
        .Properties("Afdindividual") = IIf(optAFDAddressType(0).value = True, 1, 0)
        .Properties("Afdforename") = IIf(cboAFDForename.ListIndex > 0, cboAFDForename.ItemData(cboAFDForename.ListIndex), 0)
        .Properties("Afdsurname") = IIf(cboAFDSurname.ListIndex > 0, cboAFDSurname.ItemData(cboAFDSurname.ListIndex), 0)
        .Properties("Afdinitial") = IIf(cboAFDInitial.ListIndex > 0, cboAFDInitial.ItemData(cboAFDInitial.ListIndex), 0)
        .Properties("Afdtelephone") = IIf(cboAFDTelephone.ListIndex > 0, cboAFDTelephone.ItemData(cboAFDTelephone.ListIndex), 0)
        .Properties("Afdaddress") = IIf(cboAFDAddress.ListIndex > 0, cboAFDAddress.ItemData(cboAFDAddress.ListIndex), 0)
        .Properties("Afdproperty") = IIf(cboAFDProperty.ListIndex > 0, cboAFDProperty.ItemData(cboAFDProperty.ListIndex), 0)
        .Properties("Afdstreet") = IIf(cboAFDStreet.ListIndex > 0, cboAFDStreet.ItemData(cboAFDStreet.ListIndex), 0)
        .Properties("Afdlocality") = IIf(cboAFDLocality.ListIndex > 0, cboAFDLocality.ItemData(cboAFDLocality.ListIndex), 0)
        .Properties("Afdtown") = IIf(cboAFDTown.ListIndex > 0, cboAFDTown.ItemData(cboAFDTown.ListIndex), 0)
        .Properties("Afdcounty") = IIf(cboAFDCounty.ListIndex > 0, cboAFDCounty.ItemData(cboAFDCounty.ListIndex), 0)
        'blah blah
      End If
      
      ' Update Quick Address properties.
      If tabColProps.TabVisible(iPAGE_QADDRESS) Then
        .Properties("QAddressEnabled") = IIf(chkQAPostCodeColumn.value = vbChecked, 1, 0)
        .Properties("QAindividual") = IIf(optQAAddressType(0).value = True, 1, 0)
        .Properties("QAaddress") = IIf(cboQAAddress.ListIndex > 0, cboQAAddress.ItemData(cboQAAddress.ListIndex), 0)
        .Properties("QAproperty") = IIf(cboQAProperty.ListIndex > 0, cboQAProperty.ItemData(cboQAProperty.ListIndex), 0)
        .Properties("QAstreet") = IIf(cboQAStreet.ListIndex > 0, cboQAStreet.ItemData(cboQAStreet.ListIndex), 0)
        .Properties("QAlocality") = IIf(cboQALocality.ListIndex > 0, cboQALocality.ItemData(cboQALocality.ListIndex), 0)
        .Properties("QAtown") = IIf(cboQATown.ListIndex > 0, cboQATown.ItemData(cboQATown.ListIndex), 0)
        .Properties("QAcounty") = IIf(cboQACounty.ListIndex > 0, cboQACounty.ItemData(cboQACounty.ListIndex), 0)
      End If
      
      ' Nulls causing problems with views over the metadata
      .Properties("Afdenabled") = IIf(IsNull(.Properties("Afdenabled")), 0, .Properties("Afdenabled"))
      .Properties("Afdindividual") = IIf(IsNull(.Properties("Afdindividual")), 0, .Properties("Afdindividual"))
      .Properties("Afdforename") = IIf(IsNull(.Properties("Afdforename")), 0, .Properties("Afdforename"))
      .Properties("Afdsurname") = IIf(IsNull(.Properties("Afdsurname")), 0, .Properties("Afdsurname"))
      .Properties("Afdinitial") = IIf(IsNull(.Properties("Afdinitial")), 0, .Properties("Afdinitial"))
      .Properties("Afdtelephone") = IIf(IsNull(.Properties("Afdtelephone")), 0, .Properties("Afdtelephone"))
      .Properties("Afdaddress") = IIf(IsNull(.Properties("Afdaddress")), 0, .Properties("Afdaddress"))
      .Properties("Afdproperty") = IIf(IsNull(.Properties("Afdproperty")), 0, .Properties("Afdproperty"))
      .Properties("Afdstreet") = IIf(IsNull(.Properties("Afdstreet")), 0, .Properties("Afdstreet"))
      .Properties("Afdlocality") = IIf(IsNull(.Properties("Afdlocality")), 0, .Properties("Afdlocality"))
      .Properties("Afdtown") = IIf(IsNull(.Properties("Afdtown")), 0, .Properties("Afdtown"))
      .Properties("Afdcounty") = IIf(IsNull(.Properties("Afdcounty")), 0, .Properties("Afdcounty"))
      .Properties("QAddressEnabled") = IIf(IsNull(.Properties("QAddressEnabled")), 0, .Properties("QAddressEnabled"))
      .Properties("QAindividual") = IIf(IsNull(.Properties("QAindividual")), 0, .Properties("QAindividual"))
      .Properties("QAaddress") = IIf(IsNull(.Properties("QAaddress")), 0, .Properties("QAaddress"))
      .Properties("QAproperty") = IIf(IsNull(.Properties("QAproperty")), 0, .Properties("QAproperty"))
      .Properties("QAstreet") = IIf(IsNull(.Properties("QAstreet")), 0, .Properties("QAstreet"))
      .Properties("QAlocality") = IIf(IsNull(.Properties("QAlocality")), 0, .Properties("QAlocality"))
      .Properties("QAtown") = IIf(IsNull(.Properties("QAtown")), 0, .Properties("QAtown"))
      .Properties("QAcounty") = IIf(IsNull(.Properties("QAcounty")), 0, .Properties("QAcounty"))
      
      
    End With
      
    ' Update Diary Link properties.
    With mobjColumn
      lngColumnID = .ColumnID
      ' Delete all existing diary links for this column from the database.
      daoDb.Execute "DELETE FROM tmpDiary WHERE columnID=" & lngColumnID
      ' Write the diary link information into the database if required
      If miDataType = dtTIMESTAMP Then
        .ClearDiaryLinks
        Set objDiaryLinks = .DiaryLinks
  
        ssGrdDiaryLinks.MoveFirst
        'Do While Not ssGrdDiaryLinks.Row = ssGrdDiaryLinks.Rows - 1
  
        For iLoop = 1 To ssGrdDiaryLinks.Rows
         ' ssGrdDiaryLinks.Row = iLoop - 1
  
          Set objDiaryLink = New cDiaryLink
          With objDiaryLink
            .DiaryLinkId = 0
            .ColumnID = lngColumnID
            .Comment = ssGrdDiaryLinks.Columns(0).value
            .Offset = ssGrdDiaryLinks.Columns(3).value
            .Period = ssGrdDiaryLinks.Columns(4).value
            .Reminder = ssGrdDiaryLinks.Columns(2).value
            .FilterID = ssGrdDiaryLinks.Columns(5).value
            .EffectiveDate = ssGrdDiaryLinks.Columns(6).value
            .CheckLeavingDate = ssGrdDiaryLinks.Columns(7).value

          End With
          objDiaryLinks.Add objDiaryLink
          Set objDiaryLink = Nothing
  
          ssGrdDiaryLinks.MoveNext
        Next iLoop
      
        Set objDiaryLinks = Nothing
      End If
    End With
  
  
'    ' Update Email Link properties.
'    'mobjColumn.ClearEmailLinks
'
'    For iLoop = 1 To mvarEmailLinks.Count
'
'      With mvarEmailLinks.Item(iLoop)
'
'        strKey = "ID" & .LinkID
'        mobjColumn.EmailLinks.Add mvarEmailLinks.Item(iLoop), strKey
'        'For iLoop2 = 1 To .Recipients
'        '  mobjColumn.EmailLinks.Item(strKey).Add .Recipients.Item(iLoop2).RecipientID, .Recipients.Item(iLoop2).SendType
'        'Next
'
'      End With
'    Next

    ' JDM - Fault 3252 -17/09/03 - Update other table columns that reference this lookup field
    If mobjColumn.ColumnID > 0 Then
      daoDb.Execute ("UPDATE tmpColumns SET Trimming = " & Str(miTrimming) _
        & " ,ConvertCase = " & Str(miConvertCase) _
        & " ,Alignment = " & Str(miAlignment) _
        & " ,MultiLine = " & IIf(chkMultiLine.value = vbChecked, "1", "0") _
        & " ,BlankIfZero = " & IIf(IIf((miControlType <> giCTRL_TEXTBOX) Or _
          (miDataType <> dtNUMERIC And miDataType <> dtINTEGER), _
          (chkZeroBlank.value = vbChecked), False), "1", "0") _
        & " ,Use1000Separator = " & IIf(chkUse1000Separator.value = vbChecked, "1", "0") _
        & " WHERE LookupColumnID = " & Str(mobjColumn.ColumnID) & " and ColumnType = " & Str(giCOLUMNTYPE_LOOKUP))
    End If

    If mobjColumn.ColumnID > 0 Then
      daoDb.Execute ("UPDATE tmpTables SET changed = 1" _
        & " WHERE TableID IN " _
        & "(SELECT TableID FROM tmpColumns WHERE lookupColumnID = " _
        & Str(mobjColumn.ColumnID) & ")")
    End If

    mfCancelled = False
  Else
    mfCancelled = True
  End If
  ' RH 22/08/00 - End of mblnChanged condition

TidyUpAndExit:
  'Disassociate object variables.
  Set objMisc = Nothing
  Set objDiaryLink = Nothing
  Set objDiaryLinks = Nothing
  Set rsScreens = Nothing
  Set rsOrders = Nothing
  UnLoad Me
  Exit Sub
  
ErrorTrap:
  MsgBox Err.Description, _
    vbOKOnly + vbExclamation, Application.Name
  
  Err = False
  Resume 0
  
End Sub



Private Sub cmdLostFocusClause_Click()
  Dim objExpr As CExpression

  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    .Initialise mobjColumn.TableID, mlngValidationExprID, giEXPR_RECORDVALIDATION, giEXPRVALUE_LOGIC
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      mlngValidationExprID = .ExpressionID
      
      ' Read the selected expression info.
      GetValidationExpressionDetails
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", mlngValidationExprID, False
        
        If .NoMatch Then
          mlngValidationExprID = 0
      
          ' Read the selected expression info.
          GetValidationExpressionDetails
        End If
      End With
    End If
  End With
  
  Set objExpr = Nothing
  If Not mfLoading Then Changed = True
  
  RefreshValidationTab
  
End Sub


Private Sub cmdRemoveAllDiaryLinks_Click()

  '15/08/2001 MH Fault 2679
  If RemoveAllRows("diary links", ssGrdDiaryLinks) Then
    Application.ChangedDiaryLink = True
  End If

  ' Refesh the diary link page controls.
  RefreshDiaryLinkTab
  Changed = True

End Sub

'Private Sub cmdRemoveAllEmailLinks_Click()
'
'  If RemoveAllRows("email links", ssGrdEmailLinks) Then
'
'    Do While mvarEmailLinks.Count > 0
'      mvarEmailLinks.Remove 1
'    Loop
'
'    ' Refesh the diary link page controls.
'    RefreshEmailLinkTab
'    Changed = True
'    Application.ChangedEmailLink = True   '15/08/2001 MH Fault 2679
'
'  End If
'
'End Sub

Private Sub cmdRemoveDiaryLink_Click()
  
  '15/08/2001 MH Fault 2679
  If DeleteRow("diary link", ssGrdDiaryLinks) Then
    Application.ChangedDiaryLink = True
  End If
  
  ' Refesh the diary link page controls.
  RefreshDiaryLinkTab
  Changed = True

End Sub

'Private Sub cmdRemoveEmailLink_Click()
'
'  Dim lngLinkID As Long
'  Dim iLoop As Long
'
'  lngLinkID = ssGrdEmailLinks.Columns(3).value
'
'  If DeleteRow("email link", ssGrdEmailLinks) Then
'    'On Error Resume Next
'
'    iLoop = 1
'    Do While iLoop <= mvarEmailLinks.Count
'      If mvarEmailLinks(iLoop).LinkID = lngLinkID Then
'        mvarEmailLinks.Remove iLoop
'      Else
'        iLoop = iLoop + 1
'      End If
'    Loop
'
'    Changed = True
'    Application.ChangedEmailLink = True     '15/08/2001 MH Fault 2679
'
'  End If
'
'  ' Refesh the email link page controls.
'  RefreshEmailLinkTab
'
'End Sub

Private Function DeleteRow(strType As String, ssgrd As SSDBGrid) As Boolean

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  Dim lRow As Long

  strMBText = "Delete this " & strType & ", are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Delete"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbNo Then
    DeleteRow = False
    Exit Function
  End If
  
  DeleteRow = True

  With ssgrd
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
 
  If Not mfLoading Then Changed = True
 
End Function


Public Function RemoveAllRows(strType As String, ssgrd As SSDBGrid) As Boolean

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer

  strMBText = "Remove all " & strType & " for this column, are you sure ?"
  intMBButtons = vbQuestion + vbYesNo
  strMBTitle = "Confirm Remove All"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse = vbNo Then
    RemoveAllRows = False
    Exit Function
  End If

  RemoveAllRows = True
  ssgrd.RemoveAll
  If Not mfLoading Then Changed = True

End Function



Private Sub Form_Activate()
  
  If mfLoading Then
    ' Initially set the current tab page to first tab page.
    tabColProps.Tab = iPAGE_DEFINITION
    
    ' Refresh the Definition tab page.
    RefreshDefinitionTab
  
    ' Set focus on the column name textbox.
    If txtColumnName.Enabled Then
      txtColumnName.SetFocus
    End If

    mfLoading = False
  End If
  
  ' RH 06/04/01 - BUG 2095 - needed here as well as txtmask_change
  'JPD20010727
'  If Len(txtMask.Text) > 0 Then
'    cboCase.ListIndex = 0
'    cboCase.Enabled = False
'  Else
'    cboCase.Enabled = Not mblnReadOnly
'  End If

End Sub

Private Sub Form_Initialize()
  ' Initialise properties.
  mfCancelled = True
  mfReading = True
  
  Set mvarEmailLinks = New Collection

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
  Dim lngCXBorder As Long
  
  Const iDFLTCONTROLLEFT = 1500
  Const iDFLTCONTROLTOP = 300
  
  lngCXBorder = UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX
  
  mfLoading = True

  tabColProps.Tab = 0
  tabColProps.TabVisible(iPAGE_EMAIL) = False

  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode) Or _
                  mobjColumn.Locked

  If mblnReadOnly Then
    ControlsDisableAll Me
    cmdCalculation.Enabled = True
    cmdLinkOrder.Enabled = True
    cmdLostFocusClause.Enabled = True
    cmdDfltValueExpression.Enabled = True
    cmdDiaryLinkProperties.Enabled = True
    cmdEmailLinkProperties.Enabled = True
    cmdDiaryLinkProperties.Caption = "&View"
    cmdEmailLinkProperties.Caption = "&View"
  End If


  'MH20000920
  'Temporary get rid of the email tab for build!
  'tabColProps.TabVisible(5) = False
  
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts

  'SetDateComboFormat Me.ASRDate1

  'Set maximum column name length
  txtColumnName.MaxLength = MaxColumnNameLength
  
  ' Ensure the frames on each of the tab pages have the same
  ' background colour as the tab pages themselves.
  With tabColProps
    fraDefinitionPage.BackColor = .BackColor
    fraControlPage.BackColor = .BackColor
    fraOptionsPage.BackColor = .BackColor
    fraValidationPage.BackColor = .BackColor
    fraDiaryLinkPage.BackColor = .BackColor
    fraEmailLinkPage.BackColor = .BackColor
    fraAfdPage.BackColor = .BackColor
    fraQAPage.BackColor = .BackColor
  End With
  
  ' Position frames.
  With fraLookup
    .Left = 200
    .Top = 2400
  End With
  With fraCalculation
    .Left = 200
    .Top = 2400
  End With
  With fraLink
    .Left = 200
    .Top = 2400
  End With
  With fraListValues
    .Left = 200
    .Top = 690
  End With
  With fraSpinnerProperties
    .Left = 200
    .Top = 690
  End With
  With fraStatusBarMessage
    .Left = 200
    .Top = 690
  End With
  
  With ssGrdDiaryLinks
    .Columns(0).Width = .Width / 2
    .Columns(1).Width = (.Width / 4)
    .Columns(2).Width = (.Width / 4) - (2 * lngCXBorder)
  End With
  
  ' Position controls.
  txtDefault.Left = iDFLTCONTROLLEFT
  txtDefault.Top = iDFLTCONTROLTOP
  
  cboDefault.Left = iDFLTCONTROLLEFT
  cboDefault.Top = iDFLTCONTROLTOP
  cboDefault.Width = txtDefault.Width
  
  asrDefault.Left = iDFLTCONTROLLEFT
  asrDefault.Top = iDFLTCONTROLTOP
  asrDefault.Width = txtDefault.Width
  
  TDBDefaultNumber.Left = iDFLTCONTROLLEFT
  TDBDefaultNumber.Top = iDFLTCONTROLTOP
  TDBDefaultNumber.Width = txtDefault.Width
  
  selDefaultColour.Left = iDFLTCONTROLLEFT
  selDefaultColour.Top = iDFLTCONTROLTOP
  'cmdDefaultColour.Left = iDFLTCONTROLLEFT + lblDfltColour.Width
  'cmdDefaultColour.Top = iDFLTCONTROLTOP
  
  
  'MH20010130 Fault 1610
  UI.FormatTDBNumberControl TDBDefaultNumber

  'JPD 20041112 Fault 8970
  UI.FormatGTDateControl ASRDate1
  ASRDate1.Left = iDFLTCONTROLLEFT
  ASRDate1.Top = iDFLTCONTROLTOP
  
  fraLogicDefaults.Left = iDFLTCONTROLLEFT
  fraLogicDefaults.Top = iDFLTCONTROLTOP + 60
  
  ASRDefaultWorkingPattern.Left = iDFLTCONTROLLEFT
  ASRDefaultWorkingPattern.Top = iDFLTCONTROLTOP
  
  txtDfltValueExpression.Left = iDFLTCONTROLLEFT
  cmdDfltValueExpression.Left = txtDfltValueExpression.Left + txtDfltValueExpression.Width
  
  ' Populate combos that are not dynamic.
  cboDataType_Initialize
  cboLookupTables_Initialize
  cboLinkTables_Initialize
  
  ' Enable Field Mapping if either potcode modules enabled
  If gbAFDEnabled Or gbQAddressEnabled Then
    FieldMappingInitialiseCombos
  End If
  
  ' Enable postcode pages if necessary
  tabColProps.TabVisible(iPAGE_AFD) = gbAFDEnabled
  tabColProps.TabVisible(iPAGE_QADDRESS) = gbQAddressEnabled

  'required until database structures are changed etc.
  AfdToggleControlStatus False
  QAToggleControlStatus False
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'NHRD28112002 Fault 4331 - Moved to the CmdCancel buton.
'Re-instated the code as at 30/07/2003
  Dim pintAnswer As Integer

  If mfCancelled = True Then
    If UnloadMode <> vbFormCode Then
      If Changed = True And cmdOK.Enabled Then
        pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
        If pintAnswer = vbYes Then
          cmdOK_Click
          Exit Sub
        ElseIf pintAnswer = vbCancel Then
          Cancel = True
          Exit Sub
        End If
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

  ' This needs to be here otherwise when u edit a calculated colum,
  ' the size/decimals labels do not actually appear until you select
  ' another tab and then select the first tab again.
  RefreshDefinitionTab

End Sub

Private Sub Form_Terminate()
  ' Disassociate object variables.
  Set mobjColumn = Nothing
  Set mvarEmailLinks = Nothing
End Sub


Private Sub ASRDate1_GotFocus()
  ' Select all text upon focus.
  UI.txtSelText

End Sub

Private Sub lstUniqueParents_ItemCheck(Item As Integer)
  If Not mfLoading Then Changed = True
End Sub

Private Sub optOLEStorageType_Click(Index As Integer)

  If Not mfLoading Then
    Changed = True
  End If

  ' Enable/Disable the maximum ole size field
  chkEnableOLEMaxSize.Enabled = optOLEStorageType(2).value = True And Not mblnReadOnly
  chkEnableOLEMaxSize.value = IIf(optOLEStorageType(2).value = True, chkEnableOLEMaxSize.value, vbUnchecked)

End Sub

Private Sub optQAAddressType_Click(Index As Integer)
  If Not mfReading Then
    QAToggleControlStatus True
  
    If Index = 0 Then
      cboQAAddress.ListIndex = 0
    Else
      cboQAProperty.ListIndex = 0
      cboQAStreet.ListIndex = 0
      cboQALocality.ListIndex = 0
      cboQATown.ListIndex = 0
      cboQACounty.ListIndex = 0
    End If
  End If
  If Not mfLoading Then Changed = True
  
End Sub

Private Sub optAFDAddressType_Click(Index As Integer)
  If Not mfReading Then
    AfdToggleControlStatus True
  
    If Index = 0 Then
      cboAFDAddress.ListIndex = 0
    Else
      cboAFDProperty.ListIndex = 0
      cboAFDStreet.ListIndex = 0
      cboAFDLocality.ListIndex = 0
      cboAFDTown.ListIndex = 0
      cboAFDCounty.ListIndex = 0
    End If
  End If
  If Not mfLoading Then Changed = True
  
End Sub

Private Sub optDefault_Click(Index As Integer)
  If Not mfLoading Then Changed = True

End Sub

Private Sub optDfltType_Click(Index As Integer)
  If Not mfReading Then
    Select Case Index
      Case 1 ' Calculated default.
        ' Clear the straight value, and disable the controls.
        txtDefault.Text = ""
        
        If cboDefault.ListCount > 0 Then
          cboDefault.ListIndex = 0
        End If
        
        asrDefault.value = 0
'        ASRDate1.Value = Null
        ASRDate1.Text = vbNullString
        optDefault(1).value = True
        ASRDefaultWorkingPattern.value = ""
        TDBDefaultNumber.Text = ""
        
      Case Else ' Straight value default.
        ' Clear the default expression, and disable the controls.
        mlngDfltValueExprID = 0
        GetDfltValueExpressionDetails
    End Select
      
    RefreshOptionsTab
  End If
  If Not mfLoading Then Changed = True

End Sub

Private Sub spnDefaultDisplayWidth_Change()

  If Not mfLoading Then Changed = True

End Sub

Private Sub spnMaxOLESize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  If Not mfLoading Then
    Changed = True
  End If

End Sub

Private Sub ssGrdDiaryLinks_DblClick()
  ' Display the properties form for the current diary link.
  If cmdDiaryLinkProperties.Enabled Then
    cmdDiaryLinkProperties_Click
  End If
  
End Sub


'Private Sub ssGrdEmailLinks_DblClick()
'
'  If cmdEmailLinkProperties.Enabled Then
'    cmdEmailLinkProperties_Click
'  ElseIf cmdAddEmailLink.Enabled Then
'    cmdAddEmailLink_Click
'  End If
'
'End Sub

Private Sub tabColProps_Click(PreviousTab As Integer)
  
  ' Enable, and make visible the selected tab.
  With fraDefinitionPage
    .Enabled = (tabColProps.Tab = iPAGE_DEFINITION)
    .Visible = .Enabled
  End With
  
  With fraControlPage
    .Enabled = (tabColProps.Tab = iPAGE_CONTROL)
    .Visible = .Enabled
  End With
  
  With fraOptionsPage
    .Enabled = (tabColProps.Tab = iPAGE_OPTIONS)
    .Visible = .Enabled
  End With
  
  With fraValidationPage
    .Enabled = (tabColProps.Tab = iPAGE_VALIDATION)
    .Visible = .Enabled
  End With
  
  With fraDiaryLinkPage
    .Enabled = (tabColProps.Tab = iPAGE_DIARY)
    .Visible = .Enabled
  End With
  
  'MH20000731
  With fraEmailLinkPage
    .Enabled = (tabColProps.Tab = iPAGE_EMAIL)
    .Visible = .Enabled
  End With
  
  With fraAfdPage
    .Enabled = (tabColProps.Tab = iPAGE_AFD)
    .Visible = .Enabled
  End With

  With fraQAPage
    .Enabled = (tabColProps.Tab = iPAGE_QADDRESS)
    .Visible = .Enabled
  End With

  ' Refresh the current tab page
  RefreshCurrentTab

End Sub

Private Sub TDBDefaultNumber_Change()
  If Not mfLoading Then Changed = True

End Sub

Private Sub txtColumnName_Change()
  Dim sValidatedName As String
  Dim iSelStart As Integer
  Dim iSelLen As Integer

  If Not mfLoading Then
    'JPD 20090102 Fault 13484
    sValidatedName = ValidateName(txtColumnName.Text)
    
    If sValidatedName <> txtColumnName.Text Then
      iSelStart = txtColumnName.SelStart
      iSelLen = txtColumnName.SelLength
      
      txtColumnName.Text = sValidatedName
      
      txtColumnName.SelStart = iSelStart
      txtColumnName.SelLength = iSelLen
    End If
    
    Changed = True
  End If
  
End Sub

Private Sub txtColumnName_GotFocus()
  ' Select the whole string.
  UI.txtSelText
  
End Sub

Private Sub txtColumnName_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidNameChar(KeyAscii, txtColumnName.SelStart)
  
End Sub

Private Sub optColumnType_Click(piIndex As Integer)
  If Not mfReading Then
    ' Update the Column Type global variable.
    miColumnType = optColumnType(piIndex).Tag
    
    Select Case miColumnType
      Case giCOLUMNTYPE_LOOKUP ' Lookup column.
        ' Refresh the lookup table combo.
        cboLookupTables_Refresh
        ' Set the control type.
        miControlType = giCTRL_COMBOBOX
          
        ' Initialise the column to be read-only.
        chkReadOnly.value = 0
        
        If Not mblnReadOnly Then
          spnDefaultDisplayWidth.Enabled = True
          spnDefaultDisplayWidth.BackColor = vbWindowBackground
        End If

      Case giCOLUMNTYPE_CALCULATED ' Calculated column.
        ' Read the calculation info (return type, size, decimals, etc.) from the
        ' current expression.
        GetCalculationExpressionDetails
        ' Initialise the column to be read-only.
        chkReadOnly.value = 1
        'JPD20010727
        cboCase.ListIndex = 0
        
        If Not mblnReadOnly Then
          spnDefaultDisplayWidth.Enabled = True
          spnDefaultDisplayWidth.BackColor = vbWindowBackground
        End If

      Case giCOLUMNTYPE_LINK ' Link column.
        ' Set a dummy size value.
        asrSize.Text = 1
        asrDecimals.Text = 0
        ' Refresh the link table combo.
        cboLinkTables_Refresh
        ' Set the control type.
        miControlType = giCTRL_LINK
        
        ' Initialise the column to be read-only.
        chkReadOnly.value = 0
        
        spnDefaultDisplayWidth.Enabled = False
        spnDefaultDisplayWidth.BackColor = vbButtonFace
        
      Case Else ' Data Type column.
        If Not mblnReadOnly Then
          spnDefaultDisplayWidth.Enabled = True
          spnDefaultDisplayWidth.BackColor = vbWindowBackground
        End If
        
        ' Initialise the column to be read-only.
        chkReadOnly.value = 0
    End Select
      
    ' Refresh the Definition tab page.
    cboDataType_Refresh
  
    ' Set control types for the new datatype
    cboControl_Refresh
  End If
  
  RefreshCurrentTab
  
  If Not mfLoading Then Changed = True
  
End Sub

Private Sub cboDataType_Click()
  Dim fValidCalcExpr As Boolean
  Dim fValidDfltValueExpr As Boolean
  Dim objExpr As CExpression
  
  If Not mfReading Then
    
    If miDataType = cboDataType.ItemData(cboDataType.ListIndex) Then
      Exit Sub
    End If
    
    If cboDataType.ItemData(cboDataType.ListIndex) <> dtTIMESTAMP Then
      Dim iAnswer As Integer
      'If Me.ssGrdDiaryLinks.Rows > 0 Or Me.ssGrdEmailLinks.Rows > 0 Then
      If Me.ssGrdDiaryLinks.Rows > 0 Then
        iAnswer = MsgBox("Changing the data type of a column will remove all " & _
                  "diary links defined for this column. " & _
                  "Are you sure you want to change the data type of this column?" _
                  , vbYesNo + vbQuestion, App.Title)
        
        If iAnswer = vbNo Then
          SetComboItem cboDataType, miDataType
          Exit Sub
        End If
      End If
    End If
     
    ' Change the column data type.
    miDataType = cboDataType.ItemData(cboDataType.ListIndex)
    
    'MH20010202 Fault 1757
    'Get the default control if the data type changes
    miControlType = 0
    cboControl_Refresh
        
    ' RH 06/04/2000. Fault Log 119
    ' If data type is numeric or integer, default to right justification, otherwise
    ' reset to default of left justification.
    If (miDataType = dtNUMERIC) Or _
      (miDataType = dtINTEGER) Then
      'JPD20010727
      cboTextAlignment.ListIndex = 2
    Else
      cboTextAlignment.ListIndex = 0
    End If
    
    'JPD20020325 Fault 2098. Force Working pattern controls to be uppercase.
    If (miDataType = dtLONGVARCHAR) Then
      cboCase.ListIndex = 1
    End If

    ' Set default storage type for OLE or Image type
    If miDataType = dtLONGVARBINARY Or dtVARBINARY Then
      optOLEStorageType(OLE_SERVER).value = True
    End If


    'Refresh the valdation tab
    RefreshValidationTab
  
    'Refresh whether or not the Afd enabled checkbox should be available.
    AfdControl_Refresh
      
    'Refresh whether or not the Quick Address enabled checkbox should be available.
    QAControl_Refresh
      
    ' Clear the defined calculation if it not of the defined return type.
    ' Instantiate an expression object.
    Set objExpr = New CExpression
    With objExpr
      ' Set the properties of the expression object.
      .ExpressionID = mlngCalcExprID
  
      ' Read the required info from the expression.
      fValidCalcExpr = .ReadExpressionDetails
      If fValidCalcExpr Then
        Select Case miDataType
          Case dtVARCHAR
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_CHARACTER)
          Case dtTIMESTAMP
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_DATE)
          Case dtLONGVARBINARY
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_OLE)
          Case dtVARBINARY
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_PHOTO)
          Case dtINTEGER
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_NUMERIC)
          Case dtBIT
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_LOGIC)
          Case dtNUMERIC
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_NUMERIC)
          Case dtLONGVARCHAR
            fValidCalcExpr = (.ReturnType = giEXPRVALUE_CHARACTER)
          Case Else
            fValidCalcExpr = False
        End Select
      End If
    End With
    If Not fValidCalcExpr Then
      mlngCalcExprID = 0
      GetCalculationExpressionDetails
    End If
    Set objExpr = Nothing
  
    ' Refresh the Definition tab page.
    RefreshDefinitionTab
        
    ' Set default display width for certain datatypes
    spnDefaultDisplayWidth_Refresh miDataType
   
    ' Clear the defined Default Value if it not of the defined return type.
    ' Instantiate an expression object.
    Set objExpr = New CExpression
    With objExpr
      ' Set the properties of the expression object.
      .ExpressionID = mlngDfltValueExprID
  
      ' Read the required info from the expression.
      fValidDfltValueExpr = .ReadExpressionDetails
      If fValidDfltValueExpr Then
        Select Case miDataType
          Case dtVARCHAR
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_CHARACTER)
          Case dtTIMESTAMP
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_DATE)
          Case dtLONGVARBINARY
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_OLE)
          Case dtVARBINARY
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_PHOTO)
          Case dtINTEGER
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_NUMERIC)
          Case dtBIT
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_LOGIC)
          Case dtNUMERIC
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_NUMERIC)
          Case dtLONGVARCHAR
            fValidDfltValueExpr = (.ReturnType = giEXPRVALUE_CHARACTER)
          Case Else
            fValidDfltValueExpr = False
        End Select
      End If
    End With
    If Not fValidDfltValueExpr Then
      mlngDfltValueExprID = 0
      GetDfltValueExpressionDetails
    End If
    Set objExpr = Nothing
  
    If miDataType <> dtTIMESTAMP Then
      'cmdRemoveAllDiaryLinks_Click
      
      'TM20020724 Fault 2883
'      ssGrdDiaryLinks.RemoveAll
'      ssGrdEmailLinks.RemoveAll

      With Me.ssGrdDiaryLinks
        If .Rows > 0 Then
          .RemoveAll
          Application.ChangedDiaryLink = True
        End If
      End With
      
      With Me.ssGrdEmailLinks
        If .Rows > 0 Then
          .RemoveAll
          
          Do While mvarEmailLinks.Count > 0
            mvarEmailLinks.Remove 1
          Loop

          Application.ChangedEmailLink = True
        End If
      End With
    End If
        
    ' Clear any default that may have been set up.
    txtDefault.Text = ""
    ASRDate1.Text = ""
    TDBDefaultNumber.Text = ""
    cboDefault_Refresh
    optDefault(1).value = True
    asrDefault.Text = 0
    ASRDefaultWorkingPattern.value = ""
  End If
  
  If Not mfLoading Then Changed = True

  
End Sub

Private Sub spnDefaultDisplayWidth_Refresh(piDataType As DataTypes)
  Select Case piDataType
    Case dtVARCHAR:
      spnDefaultDisplayWidth.MaximumValue = asrSize.MaximumValue
      If (miColumnType <> giCOLUMNTYPE_LOOKUP) And (miColumnType <> giCOLUMNTYPE_CALCULATED) Then
        asrSize.value = 1
      End If
      spnDefaultDisplayWidth.value = asrSize.value
    
    Case dtTIMESTAMP:
      spnDefaultDisplayWidth.MaximumValue = 10
      spnDefaultDisplayWidth.value = 10
    
    Case dtBIT:
      spnDefaultDisplayWidth.MaximumValue = 1
      spnDefaultDisplayWidth.value = 1
    
    Case dtINTEGER:
      spnDefaultDisplayWidth.MaximumValue = 10
      spnDefaultDisplayWidth.value = 10
    
    Case dtNUMERIC:
      spnDefaultDisplayWidth.MaximumValue = 15
      If (miColumnType <> giCOLUMNTYPE_LOOKUP) And (miColumnType <> giCOLUMNTYPE_CALCULATED) Then
        asrSize.value = 1
      End If
      spnDefaultDisplayWidth.value = asrSize.value
    
    Case dtLONGVARCHAR:
      spnDefaultDisplayWidth.MaximumValue = 14
      spnDefaultDisplayWidth.value = 14
      asrSize.value = 14
    
    Case dtLONGVARBINARY:
      spnDefaultDisplayWidth.MaximumValue = 255
      spnDefaultDisplayWidth.value = 255
    
    Case dtVARBINARY:
      spnDefaultDisplayWidth.MaximumValue = 255
      spnDefaultDisplayWidth.value = 255
  End Select
        
End Sub

Private Sub txtColumnName_LostFocus()
  txtColumnName.Text = Trim(txtColumnName.Text)
  
End Sub



Private Sub cboLookupColumns_Click()
  
  Dim objExpr As CExpression
  Dim fValidCalcExpr As Boolean
  
  ' Set the lookup column.
  mlngLookupColumnID = cboLookupColumns.ItemData(cboLookupColumns.ListIndex)
      
  ' Get the column data type, etc.
  GetLookupColumn
      

  ' Clear the defined calculation if it not of the defined return type.
  Set objExpr = New CExpression
  With objExpr
    ' Set the properties of the expression object.
    .ExpressionID = mlngCalcExprID

    ' Read the required info from the expression.
    fValidCalcExpr = .ReadExpressionDetails
    If fValidCalcExpr Then
      Select Case miDataType
        Case dtVARCHAR
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_CHARACTER)
        Case dtTIMESTAMP
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_DATE)
        Case dtLONGVARBINARY
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_OLE)
        Case dtVARBINARY
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_PHOTO)
        Case dtINTEGER
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_NUMERIC)
        Case dtBIT
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_LOGIC)
        Case dtNUMERIC
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_NUMERIC)
        Case dtLONGVARCHAR
          fValidCalcExpr = (.ReturnType = giEXPRVALUE_CHARACTER)
        Case Else
          fValidCalcExpr = False
      End Select
    End If
  End With
  If Not fValidCalcExpr Then
    mlngCalcExprID = 0
    GetCalculationExpressionDetails
  End If
  Set objExpr = Nothing
   
  
  If Not mfLoading Then Changed = True
  
End Sub


Private Sub cboControl_Click()
  If Not mfReading Then
    ' Get the current control type.
    miControlType = cboControl.ItemData(cboControl.ListIndex)
      
    If miControlType = giCTRL_TEXTBOX Then
      chkMultiLine.value = vbUnchecked
    End If
      
    ' Refresh Control tab page.
    RefreshControlTab
    RefreshOptionsTab

    ' JDM - 29/03/01 - Fault 1824 - Refresh the defaults
    cboDefault_Refresh
  End If
  
  If Not mfLoading Then Changed = True
  
End Sub



Private Sub txtDefault_Change()
  If Not mfLoading Then Changed = True

End Sub

Private Sub txtDefault_GotFocus()
  ' Select the whole string.
  UI.txtSelText
  
End Sub



Private Sub ReadColumnProperties()
  Dim iLoop As Integer
  Dim iCount As Integer
  Dim sComment As String
  Dim sOffset As String
  Dim sSuffix As String
  Dim sFormat As String
  Dim iDiaryPeriod As TimePeriods
  Dim fReminder As Boolean
  Dim iOffset As Integer
  Dim lFilterID As Long
  Dim dtEffectiveDate As Date
  Dim fCheckLeavingDate As Boolean
  Dim iUniqueCheckType As Integer
  Dim rsInfo As DAO.Recordset
  Dim sSQL As String
  
  Dim sBeforeAfter As String
  Dim objDiaryLinks As Collection
  Dim objDiaryLink As cDiaryLink
  Dim objEmailLink As clsEmailLink
  Dim objMisc As Misc
  'Dim sDateFormat As String
  
  ' Read the column properties from the column object.
  mfReading = True
  
  ' Determine if the column requires saving.
  mfIsSaved = (mobjColumn.IsChanged Or Not mobjColumn.IsNew)
  
  ' Set the form caption.
  If mobjColumn.IsNew And (Not mfIsSaved) Then
    Me.Caption = "New Column"
  Else
    Me.Caption = "Column Properties : " & mobjColumn.Properties("columnName") + IIf(mobjColumn.Locked, " (Locked)", "")
  End If
  
  
  With mobjColumn
    ' Initialize member variables with column property values.
    ' Definition variables.
    miColumnType = IIf(IsNull(.Properties("columnType")), 0, .Properties("columnType"))
    miDataType = IIf(IsNull(.Properties("dataType")), 0, .Properties("dataType"))
    mlngLinkTableID = IIf(IsNull(.Properties("linkTableID")), 0, .Properties("linkTableID"))
    mlngLinkViewID = IIf(IsNull(.Properties("linkViewID")), 0, .Properties("linkViewID"))
    mlngLookupTableID = IIf(IsNull(.Properties("lookupTableID")), 0, .Properties("lookupTableID"))
    mlngLookupColumnID = IIf(IsNull(.Properties("lookupColumnID")), 0, .Properties("lookupColumnID"))
    mlngCalcExprID = IIf(IsNull(.Properties("calcExprID")), 0, .Properties("calcExprID"))
    mlngLinkOrderID = IIf(IsNull(.Properties("linkOrderID")), 0, .Properties("linkOrderID"))
    
    If IsNull(.Properties("AutoUpdateLookupValues")) Then
      chkAutoUpdateRecords.value = vbUnchecked
    Else
      chkAutoUpdateRecords.value = IIf(.Properties("AutoUpdateLookupValues"), vbChecked, vbUnchecked)
    End If
     
    ' Lookup Filter Options
    mlngLookupFilterValueID = IIf(IsNull(.Properties("lookupFilterValueID")), 0, .Properties("lookupFilterValueID"))
    miLookupFilterOperator = IIf(IsNull(.Properties("lookupFilterOperator")), 0, .Properties("lookupFilterOperator"))
    mlngLookupFilterColumnID = IIf(IsNull(.Properties("lookupFilterColumnID")), 0, .Properties("lookupFilterColumnID"))
    chkLookupFilter.value = IIf(mlngLookupFilterValueID = 0, vbUnchecked, vbChecked)
    cboLookupFilterColumn_Refresh
    cboLookupFilterValue_Refresh
    cboLookupFilterValue.ListIndex = SetCombo(cboLookupFilterValue, IIf(IsNull(.Properties("lookupFilterValueID")), 0, .Properties("lookupFilterValueID")))

    ' Control variables.
    miControlType = IIf(IsNull(.Properties("controlType")), 0, .Properties("controlType"))
    ' Option variables.
    miConvertCase = IIf(IsNull(.Properties("convertCase")), 0, .Properties("convertCase"))
    miAlignment = IIf(IsNull(.Properties("alignment")), 0, .Properties("alignment"))

    ' Fault 5606 - Default new field to trim left and right
'    If mobjColumn.IsNew Then
'      miTrimming = 1
'    Else
      miTrimming = IIf(IsNull(.Properties("trimming")), 0, .Properties("trimming"))
'    End If
    
    ' Validation variables.
    mlngValidationExprID = IIf(IsNull(.Properties("lostFocusExprID")), 0, .Properties("lostFocusExprID"))
    mlngDfltValueExprID = IIf(IsNull(.Properties("dfltValueExprID")), 0, .Properties("dfltValueExprID"))
    
    If IsNull(.Properties("DefaultValue")) Then
      msDefault = vbNullString
    ElseIf .Properties("defaultValue") = "__/__/____" Then
      msDefault = vbNullString
    Else
      msDefault = .Properties("defaultValue")
    End If
    
    'msDefault = IIf(IsNull(.Properties("defaultValue")), vbNullString, .Properties("defaultValue"))

    '
    ' Initialize all of the form controls with the column property values.
    ' Initialize the Definition page controls.
    txtColumnName.Text = IIf(IsNull(.Properties("columnName")), "", .Properties("columnName"))
    
    For iLoop = optColumnType.LBound To optColumnType.UBound
      If optColumnType(iLoop).Tag = miColumnType Then
        optColumnType.Item(iLoop).value = True
      End If
    Next iLoop
    
    cboDataType_Refresh
    
    asrSize.Text = Trim(Str(IIf(IsNull(.Properties("size")), 0, .Properties("size"))))
    asrDecimals.Text = Trim(Str(IIf(IsNull(.Properties("decimals")), 0, .Properties("decimals"))))
    cboLookupTables_Refresh
    cboLookupColumns_Refresh
    GetCalculationExpressionDetails
    chkCalculateIfEmpty.value = IIf(IIf(IsNull(.Properties("CalculateIfEmpty")), 0, .Properties("CalculateIfEmpty")), 1, 0)
    GetLinkOrderDetails
    cboLinkTables_Refresh
    
    ' RH 19/09/00 - BUG 797. Should onle be able to select link if the column is
    '                        from a child table, not a parent/lookup
    'Me.optColumnType(3).Enabled = (cboLinkTables.ListCount <> 0)
    Me.optColumnType(3).Enabled = (cboLinkTables.ListCount <> 0 And Not mblnReadOnly)
    
    spnDefaultDisplayWidth.value = Trim(IIf(IsNull(.Properties("defaultdisplaywidth")), 0, .Properties("defaultdisplaywidth")))
    
    ' Initialize the Control page.
    cboControl_Refresh
    txtListValues.Text = .ControlValuesString
    asrMinVal.Text = Trim(Str(IIf(IsNull(.Properties("spinnerMinimum")), 0, .Properties("spinnerMinimum"))))
    asrMaxVal.Text = Trim(Str(IIf(IsNull(.Properties("spinnerMaximum")), 0, .Properties("spinnerMaximum"))))
    asrIncVal.Text = Trim(Str(IIf(IsNull(.Properties("spinnerIncrement")), 0, .Properties("spinnerIncrement"))))
    txtStatusBarMessage.Text = Trim(IIf(IsNull(.Properties("statusBarMessage")), "", .Properties("statusBarMessage")))

    ' Initialize the Options page.
    chkReadOnly.value = IIf(IIf(IsNull(.Properties("readOnly")), 0, .Properties("readOnly")), 1, 0)
    chkAudit.value = IIf(IIf(IsNull(.Properties("audit")), 0, .Properties("audit")), 1, 0)
    
    ' OLE storage types
    For iCount = optOLEStorageType.LBound To optOLEStorageType.UBound
      If .Properties("OLEType") = iCount Then
        optOLEStorageType(iCount).value = True
      End If
    Next iCount
    
    chkEnableOLEMaxSize.value = IIf(IsNull(.Properties("MaxOLESizeEnabled")), vbUnchecked, IIf(.Properties("MaxOLESizeEnabled") = True, vbChecked, vbUnchecked))
    asrMaxOLESize.value = IIf(IsNull(.Properties("MaxOLESize")), 100, .Properties("MaxOLESize"))
    
    ' These options are taken from the lookup table settings
    If Not miColumnType = giCOLUMNTYPE_LOOKUP Then
      chkMultiLine.value = IIf(IIf(IsNull(.Properties("multiLine")), 0, .Properties("multiLine")), 1, 0)
      chkZeroBlank.value = IIf(IIf(IsNull(.Properties("blankIfZero")), 0, .Properties("blankIfZero")), 1, 0)
      cboCase.ListIndex = miConvertCase
      For iLoop = 0 To cboTextAlignment.ListCount - 1
        If cboTextAlignment.ItemData(iLoop) = miAlignment Then
          cboTextAlignment.ListIndex = iLoop
        End If
      Next iLoop
      
      chkUse1000Separator.value = IIf(IIf(IsNull(.Properties("Use1000Separator")), 0, .Properties("Use1000Separator")), 1, 0)
      
      ' Set the trimming option
      cboTrimming.ListIndex = miTrimming
    End If
                             
    GetDfltValueExpressionDetails
    If (mlngDfltValueExprID > 0) Then
      optDfltType(1).value = True
    Else
      optDfltType(0).value = True
    End If
    
    txtDefault.Text = Trim(msDefault)
    If miDataType = dtTIMESTAMP Then
      msDefault = Trim(msDefault)
      If Len(msDefault) = 8 Then
        ' Previous version saved the defult dates in the format mmddyyyy.
        ' If the default is in this format, convert to mm/dd/yyyy format.
        msDefault = Left(msDefault, 2) & "/" & Mid(msDefault, 3, 2) & "/" & Mid(msDefault, 5)
      End If

      Set objMisc = New Misc
      ASRDate1.Text = IIf(Len(msDefault) > 0, objMisc.ConvertSQLDateToLocale(msDefault), "")
      Set objMisc = Nothing
    ElseIf miDataType = dtINTEGER Then
      TDBDefaultNumber.Format = "##########"
      TDBDefaultNumber.DisplayFormat = TDBDefaultNumber.Format
      TDBDefaultNumber.MaxValue = 2147483647#
      TDBDefaultNumber.MinValue = -2147483648#
      TDBDefaultNumber.Text = msDefault
      selDefaultColour.BackColor = val(msDefault)

    ElseIf miDataType = dtNUMERIC Then
      sFormat = ""
      For iCount = 1 To (asrSize.value - asrDecimals.value)
        sFormat = sFormat & "#"
      Next iCount

      If Len(sFormat) > 0 Then
        sFormat = Left(sFormat, Len(sFormat) - 1) & "0"
      End If
                  
      If asrDecimals.value > 0 Then
        sFormat = sFormat & "."
        For iCount = 1 To asrDecimals.value
          sFormat = sFormat & "0"
        Next iCount
      End If
      
      If Len(sFormat) = 0 Then
        sFormat = "0"
      End If
      TDBDefaultNumber.Format = sFormat
      TDBDefaultNumber.DisplayFormat = TDBDefaultNumber.Format
      TDBDefaultNumber.Text = msDefault
    End If
    cboDefault_Refresh
    If msDefault = "FALSE" Then
      optDefault(1).value = True
    Else
      optDefault(0).value = True
    End If
    asrDefault.Text = Trim(msDefault)
    ASRDefaultWorkingPattern.value = msDefault

    ' Initialize the Validation page.
    chkDuplicate.value = IIf(IIf(IsNull(.Properties("duplicate")), 0, .Properties("duplicate")), 1, 0)
    chkMandatory.value = IIf(IIf(IsNull(.Properties("mandatory")), 0, .Properties("mandatory")), 1, 0)
    'JPD20010727
'    chkUnique.Value = IIf(IIf(IsNull(.Properties("uniqueCheck")), 0, .Properties("uniqueCheck")), 1, 0)
'    chkChildUnique.Value = IIf(IIf(IsNull(.Properties("childUniqueCheck")), 0, .Properties("childUniqueCheck")), 1, 0)
    iUniqueCheckType = IIf(IsNull(.Properties("uniqueCheckType")), 0, .Properties("uniqueCheckType"))
    
    sSQL = "SELECT COUNT(*) AS recCount" & _
      " FROM tmpRelations" & _
      " WHERE childID = " & mobjColumn.TableID
    Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    miParentCount = rsInfo!reccount
    rsInfo.Close
    Set rsInfo = Nothing
  
    chkUnique.value = IIf(iUniqueCheckType = giUNIQUECHECKTYPE_ENTIRE, vbChecked, vbUnchecked)
    chkChildUnique.value = IIf((miParentCount > 0) And _
      ((iUniqueCheckType = giUNIQUECHECKTYPE_SIBLINGSALL) Or (iUniqueCheckType > 0)), _
      vbChecked, vbUnchecked)
      
    lstUniqueParents.Clear
    
    If miParentCount > 1 Then
      sSQL = "SELECT tmpTables.tableName, tmpTables.tableID" & _
        " FROM tmpTables, tmpRelations" & _
        " WHERE tmptables.deleted = FALSE" & _
        " AND tmpRelations.childID = " & Trim(Str(mobjColumn.Properties("tableID"))) & _
        " AND tmpTables.tableID = tmpRelations.parentID" & _
        " ORDER BY tmpTables.tableName"
      
      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      Do While Not (rsInfo.EOF)
        lstUniqueParents.AddItem (rsInfo!TableName)
        lstUniqueParents.ItemData(lstUniqueParents.NewIndex) = rsInfo!TableID
        lstUniqueParents.Selected(lstUniqueParents.NewIndex) = (iUniqueCheckType = giUNIQUECHECKTYPE_SIBLINGSALL) Or _
          (iUniqueCheckType = rsInfo!TableID)
        
        rsInfo.MoveNext
      Loop
      
      rsInfo.Close
      Set rsInfo = Nothing
    End If
      
    txtMask.Text = IIf(IsNull(.Properties("mask")), "", .Properties("mask"))
    
    GetValidationExpressionDetails
    txtErrorMessage.Text = Trim(IIf(IsNull(.Properties("errorMessage")), "", .Properties("errorMessage")))
    
    ' Initialize the Afd Page. RH 7/9/99
    ' (Only if Afd module is enabled)
    ' Check whether or not the Afd enabled checkbox should be available.
    AfdControl_Refresh
    
    ' Check whether or not the QA enabled checkbox should be available.
    QAControl_Refresh
    
    'If tabColProps.TabVisible(5) Then
    If tabColProps.TabVisible(iPAGE_AFD) Then
      chkAFDPostCodeColumn.value = IIf(IIf(IsNull(.Properties("Afdenabled")), 0, .Properties("Afdenabled")), 1, 0)
      If chkAFDPostCodeColumn.value = 1 Then
        cboAFDForename.ListIndex = SetCombo(cboAFDForename, .Properties("Afdforename"))
        cboAFDSurname.ListIndex = SetCombo(cboAFDSurname, .Properties("Afdsurname"))
        cboAFDInitial.ListIndex = SetCombo(cboAFDInitial, .Properties("Afdinitial"))
        cboAFDTelephone.ListIndex = SetCombo(cboAFDTelephone, .Properties("Afdtelephone"))

        If .Properties("Afdindividual") = True Then
          optAFDAddressType(0).value = True
          cboAFDProperty.ListIndex = SetCombo(cboAFDProperty, .Properties("Afdproperty"))
          cboAFDStreet.ListIndex = SetCombo(cboAFDStreet, .Properties("Afdstreet"))
          cboAFDLocality.ListIndex = SetCombo(cboAFDLocality, .Properties("Afdlocality"))
          cboAFDTown.ListIndex = SetCombo(cboAFDTown, .Properties("Afdtown"))
          cboAFDCounty.ListIndex = SetCombo(cboAFDCounty, .Properties("Afdcounty"))
        Else
          optAFDAddressType(1).value = True
          cboAFDAddress.ListIndex = SetCombo(cboAFDAddress, .Properties("Afdaddress"))
        End If
      End If
    End If
    'Enable/Disable relevant fields
    AfdToggleControlStatus chkAFDPostCodeColumn.value
    
    ' Quick Address fields
    If tabColProps.TabVisible(iPAGE_QADDRESS) Then
      chkQAPostCodeColumn.value = IIf(IIf(IsNull(.Properties("QAddressEnabled")), 0, .Properties("QAddressEnabled")), 1, 0)
      If chkQAPostCodeColumn.value = 1 Then
        If .Properties("QAindividual") = True Then
          optQAAddressType(0).value = True
          cboQAProperty.ListIndex = SetCombo(cboQAProperty, .Properties("QAproperty"))
          cboQAStreet.ListIndex = SetCombo(cboQAStreet, .Properties("QAstreet"))
          cboQALocality.ListIndex = SetCombo(cboQALocality, .Properties("QAlocality"))
          cboQATown.ListIndex = SetCombo(cboQATown, .Properties("QAtown"))
          cboQACounty.ListIndex = SetCombo(cboQACounty, .Properties("QAcounty"))
        Else
          optQAAddressType(1).value = True
          cboQAAddress.ListIndex = SetCombo(cboQAAddress, IIf(IsNull(.Properties("QAaddress")), 0, .Properties("QAaddress")))
        End If
      End If
    End If
    'Enable/Disable relevant fields
    QAToggleControlStatus chkQAPostCodeColumn.value
    
    ' Initialize the Diary Link page.
    ssGrdDiaryLinks.RemoveAll
    
    'Only do the diary links for date columns...
    If miDataType = dtTIMESTAMP Then
    
      Set objDiaryLinks = mobjColumn.DiaryLinks
      For Each objDiaryLink In objDiaryLinks
        sComment = objDiaryLink.Comment
        iOffset = objDiaryLink.Offset
        sOffset = Trim(Str(iOffset))
        iDiaryPeriod = objDiaryLink.Period
        lFilterID = objDiaryLink.FilterID
        dtEffectiveDate = objDiaryLink.EffectiveDate
        fCheckLeavingDate = objDiaryLink.CheckLeavingDate
        
        sOffset = GetOffset(iOffset, iDiaryPeriod, False)
        
        If iOffset = 0 Then
          sOffset = "No offset"
        Else
          If iOffset < 0 Then
            sBeforeAfter = " before"
            sSuffix = IIf(iOffset = -1, "", "s")
            sOffset = Trim(Str(iOffset * -1))
          Else
            sBeforeAfter = " after"
            sSuffix = IIf(iOffset = 1, "", "s")
            sOffset = Trim(Str(iOffset))
          End If
          
          'Select Case iDiaryPeriod
          '  Case iTimePeriodDays
          '    sOffset = sOffset & " day" & sSuffix & sBeforeAfter
          '  Case iTimePeriodMonths
          '    sOffset = sOffset & " month" & sSuffix & sBeforeAfter
          '  Case iTimePeriodWeeks
          '    sOffset = sOffset & " week" & sSuffix & sBeforeAfter
          '  Case iTimePeriodYears
          '    sOffset = sOffset & " year" & sSuffix & sBeforeAfter
          'End Select
          sOffset = sOffset & " " & _
              TimePeriod(iDiaryPeriod) & _
              sSuffix & sBeforeAfter
        
        End If
          
        fReminder = objDiaryLink.Reminder
            
        ' Add the diary link to the grid.
        ssGrdDiaryLinks.AddItem sComment & _
          vbTab & sOffset & _
          vbTab & fReminder & _
          vbTab & iOffset & _
          vbTab & iDiaryPeriod & _
          vbTab & lFilterID & _
          vbTab & dtEffectiveDate & _
          vbTab & fCheckLeavingDate
          'vbTab & Format(dtEffectiveDate, "mm/dd/yyyy")

      Next objDiaryLink
      Set objDiaryLink = Nothing
      Set objDiaryLinks = Nothing

    End If

    RefreshCurrentTab
  End With
  
  
'  ' Initialize the Email Link page.
'  ssGrdEmailLinks.RemoveAll
'  'Set mvarEmailLinks = mobjColumn.EmailLinks
'  For iLoop = 1 To mobjColumn.EmailLinks.Count
'    mvarEmailLinks.Add mobjColumn.EmailLinks.Item(iLoop), "ID" & mobjColumn.EmailLinks.Item(iLoop).LinkID
'  Next
'  For Each objEmailLink In mvarEmailLinks
'
'    ' Add the Email link to the grid.
'    ssGrdEmailLinks.AddItem _
'      objEmailLink.Title & vbTab & _
'      GetOffset(objEmailLink.Offset, objEmailLink.OffsetPeriod, objEmailLink.Immediate) & vbTab & _
'      objEmailLink.Subject & vbTab & objEmailLink.LinkID
'
'  Next objEmailLink
'  Set objEmailLink = Nothing
  
  mfReading = False

  Changed = False

End Sub

Private Sub RefreshCurrentTab()
  
  mfLoading = True
  
  'Refresh the controls on the active tab page
  Select Case tabColProps.Tab
  Case iPAGE_DEFINITION ' Definition tab.
    RefreshDefinitionTab
      
  Case iPAGE_CONTROL    ' Control tab.
    RefreshControlTab
    
  Case iPAGE_OPTIONS    ' Options tab.
    RefreshOptionsTab
      
  Case iPAGE_VALIDATION ' Validation tab.
    RefreshValidationTab
      
  Case iPAGE_DIARY      ' Diary Link tab.
    RefreshDiaryLinkTab
  
  Case iPAGE_EMAIL      ' Email Link tab.
    RefreshEmailLinkTab
  
  End Select
  
  mfLoading = False
  
  Me.Refresh
  
End Sub

Private Sub cboLookupColumns_Refresh()
  ' Refresh the Lookup Columns combo.
  Dim iIndex As Integer
  
  iIndex = 0
  
  ' Clear lookup columns combo.
  cboLookupColumns.Clear
  
  If (miColumnType = giCOLUMNTYPE_LOOKUP) And _
    (mlngLookupTableID > 0) Then
  
    ' Loop through columns for selected lookup table.
    With recColEdit
      .Index = "idxName"
      .Seek ">=", mlngLookupTableID
      
      If Not .NoMatch Then
        Do While Not .EOF
          If .Fields("tableID") <> mlngLookupTableID Then
            Exit Do
          End If
          
          ' Add each column name to the lookup columns combo.
          ' NB. We only want to add certain types of column. There's not use in
          ' looking up OLE or logic values.
          If (.Fields("columnType") <> giCOLUMNTYPE_SYSTEM) And _
            (.Fields("columnType") <> giCOLUMNTYPE_LINK) And _
            (Not .Fields("deleted")) And _
            (.Fields("dataType") <> dtLONGVARBINARY) And _
            (.Fields("dataType") <> dtVARBINARY) And _
            (.Fields("dataType") <> dtBIT) Then
            
            'MH20071116 Fault 12458
            If .Fields("ColumnID") <> mobjColumn.ColumnID Then
              cboLookupColumns.AddItem .Fields("columnName")
              cboLookupColumns.ItemData(cboLookupColumns.NewIndex) = .Fields("columnID")
        
              If .Fields("columnID") = mlngLookupColumnID Then
                iIndex = cboLookupColumns.NewIndex
              End If
            End If
          End If
      
          .MoveNext
        Loop
      End If
    End With
  End If
  
  ' Enable the combo if there are items.
  With cboLookupColumns
    If .ListCount > 0 Then
      .ListIndex = iIndex
      '.Enabled = True
'      .Enabled = Not mblnReadOnly
    Else
      .Enabled = False
    End If
  End With
  
  Exit Sub
  
End Sub
Private Sub cboDataType_Refresh()
  Dim iLoop As Integer
  Dim iIndex As Integer
    
  cboDataType_Initialize
  
  ' Add/remove the 'Link' data type as required.
  If miColumnType = giCOLUMNTYPE_LINK Then
    With cboDataType
      .AddItem "Link"
      .ItemData(.NewIndex) = dtBINARY
    End With
  End If

  ' Select the current data type in the data type combo.
  iIndex = 0
  With cboDataType
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = miDataType Then
        iIndex = iLoop
      End If
    Next iLoop
  
    ' If the current data type is not in the combo
    ' select the first combo item.
    .ListIndex = iIndex
  End With

End Sub
Private Sub cboControl_Refresh()
  ' Populate the control type combo with the valid
  ' controls for the current data type.
  Dim iLoop As Integer
  Dim iNewIndex As Integer
  Dim iDefaultIndex As Integer
  Dim iDefaultControl As ControlTypes
  
  With cboControl
    ' Clear the combos list items.
    .Clear
  
    If miColumnType = giCOLUMNTYPE_LOOKUP Then
      ' If the column is a lookup column then the control must be 'combo'.
      .AddItem "Dropdown List"
      .ItemData(.NewIndex) = giCTRL_COMBOBOX
      iDefaultControl = giCTRL_COMBOBOX
      
    ElseIf miColumnType = giCOLUMNTYPE_LINK Then
      ' If the column is a Link column then the control must be 'link button'.
      .AddItem "Link button"
      .ItemData(.NewIndex) = giCTRL_LINK
      iDefaultControl = giCTRL_LINK
      
    Else
      ' If the column is neither a lookup nor link column then then control type is
      ' determined by the column data type.
      iDefaultControl = DefaultControl(miDataType)
    
      Select Case miDataType
        ' Logic.
        Case dtBIT
          .AddItem "Check Box"
          .ItemData(.NewIndex) = giCTRL_CHECKBOX
          
        ' Character.
        Case dtVARCHAR
          If miColumnType <> giCOLUMNTYPE_CALCULATED Then
            .AddItem "Dropdown List"
            .ItemData(.NewIndex) = giCTRL_COMBOBOX
          End If
          If miColumnType = giCOLUMNTYPE_DATA Then
            .AddItem "Option Group"
            .ItemData(.NewIndex) = giCTRL_OPTIONGROUP
          End If
          If miColumnType <> giCOLUMNTYPE_LOOKUP Then
            .AddItem "Text Box"
            .ItemData(.NewIndex) = giCTRL_TEXTBOX
            
            .AddItem "Navigation"
            .ItemData(.NewIndex) = giCTRL_NAVIGATION
                        
          End If
          
        ' OLE.
        Case dtLONGVARBINARY
          .AddItem "OLE"
          .ItemData(.NewIndex) = giCTRL_OLE
  
        ' Photo.
        Case dtVARBINARY
          .AddItem "Photo"
          .ItemData(.NewIndex) = giCTRL_PHOTO
          
        ' Integer.
        Case dtINTEGER
          .AddItem "Spinner"
          .ItemData(.NewIndex) = giCTRL_SPINNER
          .AddItem "Text Box"
          .ItemData(.NewIndex) = giCTRL_TEXTBOX
          .AddItem "Colour Picker"
          .ItemData(.NewIndex) = giCTRL_COLOURPICKER
          
        ' Working Pattern.
        Case dtLONGVARCHAR
          .AddItem "Working Pattern"
          .ItemData(.NewIndex) = giCTRL_WORKINGPATTERN
          
        ' Else ?
        Case Else
          .AddItem "Text Box"
          .ItemData(.NewIndex) = giCTRL_TEXTBOX
          
      End Select
    End If
      
    iNewIndex = -1
    iDefaultIndex = -1

    ' Select the current item in the list of control types.
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = miControlType Then
        iNewIndex = iLoop
      End If

      If .ItemData(iLoop) = iDefaultControl Then
        iDefaultIndex = iLoop
      End If
    Next iLoop

    If iNewIndex < 0 Then
      ' The current control is not in the combo so
      ' select the default item.
      If iDefaultIndex >= 0 Then
        iNewIndex = iDefaultIndex
      Else
        ' The default control is not in the combo so
        ' select the first combo item.
        iNewIndex = 0
      End If
    End If

    .ListIndex = iNewIndex
    
    ' Get the type of selected control.
    miControlType = .ItemData(.ListIndex)
  End With
  
End Sub

Private Sub QAControl_Refresh()
  
  If (miDataType = dtVARCHAR) Then
    chkQAPostCodeColumn.Enabled = Not mblnReadOnly
  Else
    chkQAPostCodeColumn.value = False
    chkQAPostCodeColumn.Enabled = False
  End If

End Sub

Private Sub AfdControl_Refresh()
'  If (miDataType = dtVARCHAR) And _
    (miColumnType <> giCOLUMNTYPE_WORKINGPATTERN) Then
  If (miDataType = dtVARCHAR) Then
    'chkAFDPostCodeColumn.Enabled = True
    chkAFDPostCodeColumn.Enabled = Not mblnReadOnly
  Else
    chkAFDPostCodeColumn.value = False
    chkAFDPostCodeColumn.Enabled = False
  End If

End Sub

Private Sub cboDefault_Refresh()
  Dim sDefaultValue As String
  Dim sListValues As String
  Dim iIndex As Integer
  Dim iRecCount As Integer
  Dim sValue As String
  Dim iSelection As Integer
  Dim fDefaultChanged As Boolean
  Dim sSQL As String
  Dim rsLookupValues As New ADODB.Recordset
  Dim rsInfo As New ADODB.Recordset
  Dim vValue As Variant
  Dim objMisc As Misc

  Set objMisc = New Misc
  
  ' RH 31/07/00 - To store the longest item in the option/combo box options
  Dim iMaxLength As Integer

  sDefaultValue = msDefault
  iSelection = 0
  fDefaultChanged = True
  
  ' Clear the combo.
  cboDefault.Clear
  
  If miColumnType = giCOLUMNTYPE_LOOKUP Then
    ' Populate the default values combo with the values in the lookup table.
    cboDefault.AddItem "<None>"
    If (UCase(Trim("<None>")) = UCase(Trim(sDefaultValue))) Or _
      (Len(sDefaultValue) = 0) Then
      iSelection = cboDefault.NewIndex
      fDefaultChanged = False
    End If
    
    If cboLookupColumns.ListIndex >= 0 Then
      ' Check that the server-side definition of the selected column in the lookup table
      ' exists, and matches with the local version.
      sSQL = "SELECT COUNT(ASRSysColumns.columnID) AS recCount" & _
        " FROM ASRSysColumns" & _
        " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
        " WHERE ASRSysColumns.columnName = '" & cboLookupColumns.Text & "'" & _
        " AND ASRSysTables.tableName = '" & cboLookupTables.Text & "'" & _
        " AND ASRSysColumns.dataType = " & Trim(Str(miDataType))
      rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      iRecCount = rsInfo!reccount
      rsInfo.Close
      Set rsInfo = Nothing
        
      If (iRecCount > 0) Then
        sSQL = "SELECT DISTINCT TOP 30000 " & cboLookupColumns.Text & " AS lookupValue" & _
          " FROM " & cboLookupTables.Text & _
          " ORDER BY lookupValue"
        
        rsLookupValues.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rsLookupValues.EOF
          vValue = rsLookupValues!LookupValue
          
          If Not IsNull(vValue) Then
            Select Case miDataType
              Case dtNUMERIC, dtINTEGER
                cboDefault.AddItem Trim(Str(vValue))
                If vValue = val(sDefaultValue) Then
                  iSelection = cboDefault.NewIndex
                  fDefaultChanged = False
                End If
        
              Case dtTIMESTAMP
                If IsDate(vValue) Then
                  'JPD 20041115 Fault 8970
                  cboDefault.AddItem Format(vValue, objMisc.DateFormat)
                  If Replace(Format(vValue, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") = sDefaultValue Then
                    iSelection = cboDefault.NewIndex
                    fDefaultChanged = False
                  End If
                End If
        
              Case Else
                cboDefault.AddItem Trim(vValue)
                If UCase(Trim(vValue)) = UCase(Trim(sDefaultValue)) Then
                  iSelection = cboDefault.NewIndex
                  fDefaultChanged = False
                End If
            End Select
          End If
          
          rsLookupValues.MoveNext
        Loop
        
        rsLookupValues.Close
        Set rsLookupValues = Nothing
      End If
    End If
    
    'JPD 20050812 Fault 10256
    If Not mfClearDefault Then
      If (iSelection = 0) And (sDefaultValue <> "") And (UCase(Trim(sDefaultValue)) <> "<NONE>") Then
        If miDataType = dtTIMESTAMP Then
          'JPD 20041115 Fault 8970
          cboDefault.AddItem objMisc.ConvertSQLDateToLocale(sDefaultValue)
        Else
          cboDefault.AddItem Trim(sDefaultValue)
        End If
        iSelection = cboDefault.NewIndex
        fDefaultChanged = False
      End If
    End If

  ElseIf (miControlType = giCTRL_COMBOBOX) Or (miControlType = giCTRL_OPTIONGROUP) Then
    ' JPD 10/4/01 - Add a blank entry if the column is not mandatory.
    ' JPD20020429 Fault 3041 Force optiongroups to have a default value.
    If (chkMandatory.value = vbUnchecked) And _
      (miControlType = giCTRL_COMBOBOX) Then
      cboDefault.AddItem ""
        
      If sDefaultValue = "" Then
        iSelection = cboDefault.NewIndex
        fDefaultChanged = False
      End If
    End If
    
    sListValues = txtListValues.Text
    
    ' Add combo items for each control value in the string.
    While Len(sListValues) > 0
      iIndex = InStr(sListValues, vbCr & vbLf)
      
      If iIndex > 0 Then
        sValue = Left(sListValues, iIndex - 1)
        sListValues = Mid(sListValues, iIndex + 2)
      Else
        sValue = sListValues
        sListValues = ""
      End If
          
      If Len(sValue) > 0 Then
        
        ' RH 31/07/00 - store the longest item in the option/combo box options
        If Len(sValue) > iMaxLength Then
          iMaxLength = Len(sValue)
        End If
        
        cboDefault.AddItem sValue
        
        If UCase(sValue) = UCase(Trim(sDefaultValue)) Then
          iSelection = cboDefault.NewIndex
          fDefaultChanged = False
        End If
      End If
    Wend
  
    ' RH 12/01/01 - Automatically set the field size to the longest option if lookup
    'If miColumnType = giCOLUMNTYPE_LOOKUP Then
    If iMaxLength <> 0 Then
      Me.asrSize.value = iMaxLength
    End If
    'Me.asrSize.Value = IIf(iMaxLength = 0, IIf(Me.txtListValues.Text = "", 1, iMaxLength), iMaxLength)
  End If

  With cboDefault
    If .ListCount > 0 Then
    
      'JPD 20050812 Fault 10256
      If mfClearDefault Then iSelection = 0
      
      .ListIndex = iSelection
      
      '.Enabled = True
      'TM20011121 Fault 3169
      .Enabled = Not mblnReadOnly
    Else
      .Enabled = False
    End If
  End With

  Set objMisc = Nothing

End Sub

Private Sub cboLookupTables_Refresh()
  ' Refresh the Lookup Tables combo.
  Dim iIndex As Integer
  Dim iLoop As Integer
  
  If miColumnType = giCOLUMNTYPE_LOOKUP Then
    iIndex = 0
    
    ' Find the current lookup table in the combo.
    For iLoop = 0 To cboLookupTables.ListCount - 1
      If cboLookupTables.ItemData(iLoop) = mlngLookupTableID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
        
    ' Select the current lookup table in the combo.
    ' If the current table is not in the combo then
    ' select the first item.
    If cboLookupTables.ListCount > 0 Then
      If cboLookupTables.ListIndex = iIndex Then
        cboLookupTables_Click
      Else
        cboLookupTables.ListIndex = iIndex
      End If
    End If
  End If
  
End Sub
Private Sub cboLinkTables_Refresh()
  ' Refresh the Link Tables combo.
  Dim iIndex As Integer
  Dim iLoop As Integer
  
  If miColumnType = giCOLUMNTYPE_LINK Then
    iIndex = 0
    
    ' Find the current lookup table in the combo.
    For iLoop = 0 To cboLinkTables.ListCount - 1
      If cboLinkTables.ItemData(iLoop) = mlngLinkTableID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
        
    ' Select the current lookup table in the combo.
    ' If the current table is not in the combo then
    ' select the first item.
    If cboLinkTables.ListCount > 0 Then
      If cboLinkTables.ListIndex = iIndex Then
        cboLinkTables_Click
      Else
        cboLinkTables.ListIndex = iIndex
      End If
    End If
  End If
  
End Sub

Private Sub cboLookupTables_Initialize()
  ' Populate the lookup tables combo.
  
  ' Clear the combo.
  cboLookupTables.Clear
  
  With recTabEdit
    .Index = "idxName"
    .MoveFirst
    
    Do While Not .EOF
      ' Add items to the combo for each lookup tables that
      ' has not been deleted.
      
      'JPD 20031216 Islington changes
      'If !TableType = iTabLookup And Not !deleted Then
      If Not !Deleted Then
        cboLookupTables.AddItem !TableName
        cboLookupTables.ItemData(cboLookupTables.NewIndex) = !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
End Sub


Private Sub cboLinkTables_Initialize()
  ' Populate the link tables combo.
  Dim sSQL As String
  Dim rsParents As DAO.Recordset
  
  ' Clear the combo.
  cboLinkTables.Clear
  
  ' Get the names of parent tables.
  sSQL = "SELECT tmpTables.tableName, tmpTables.tableID" & _
    " FROM tmpTables, tmpRelations" & _
    " WHERE tmptables.deleted = FALSE" & _
    " AND tmpRelations.childID = " & Trim(Str(mobjColumn.Properties("tableID"))) & _
    " AND tmpTables.tableID = tmpRelations.parentID"
  Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  With rsParents
    Do While Not (.EOF)
      cboLinkTables.AddItem !TableName
      cboLinkTables.ItemData(cboLinkTables.NewIndex) = !TableID
      
      rsParents.MoveNext
    Loop
      
    .Close
  End With
  
  Set rsParents = Nothing
  
End Sub


Private Sub cboDataType_Initialize()
  ' Populate the data type combo.
  With cboDataType
    ' Clear the combo.
    .Clear
    
    ' Add an item for each data type.
    .AddItem "Character"
    .ItemData(.NewIndex) = dtVARCHAR
    
    .AddItem "Date"
    .ItemData(.NewIndex) = dtTIMESTAMP
       
    .AddItem "OLE object"
    .ItemData(.NewIndex) = dtLONGVARBINARY
    
    .AddItem "Integer"
    .ItemData(.NewIndex) = dtINTEGER
    
    .AddItem "Logic"
    .ItemData(.NewIndex) = dtBIT
    
    .AddItem "Numeric"
    .ItemData(.NewIndex) = dtNUMERIC
          
    .AddItem "Photo"
    .ItemData(.NewIndex) = dtVARBINARY
    
'    .AddItem "Unique Identifier"
'    .ItemData(.NewIndex) = dtGUID
    
    .AddItem "Working Pattern"
    .ItemData(.NewIndex) = dtLONGVARCHAR
    
  End With

End Sub


Private Function GetLookupColumn() As Boolean
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim iOriginalDataType As DataTypes
  
  'JPD 20050812 Fault 10256
  iOriginalDataType = miDataType
  
  If miColumnType = giCOLUMNTYPE_LOOKUP Then
    With recColEdit
      ' Find the required column's record.
      .Index = "idxColumnID"
      .Seek "=", mlngLookupColumnID
      
      If .NoMatch Then
        ' Re-set the column info controls.
        asrSize.Text = ""
        asrDecimals.Text = ""
        miDataType = dtVARCHAR
        
        ' JPD20021106 Fault 4690
        spnDefaultDisplayWidth.Text = ""

        GetLookupColumn = False
      Else
        ' Read the column info from the database.
        asrSize.Text = Trim(Str(.Fields("size")))
        asrDecimals.Text = Trim(Str(.Fields("decimals")))
        miDataType = .Fields("dataType")

        spnDefaultDisplayWidth_Refresh miDataType

        'JDM - 16/09/03 - Fault 3252 - Pull across details from the lookup
        chkMultiLine.value = IIf(IsNull(.Fields("multiline")), vbUnchecked, IIf(.Fields("multiline"), vbChecked, vbUnchecked))
        cboCase.ListIndex = .Fields("convertcase")
        
        For iLoop = 0 To cboTextAlignment.ListCount - 1
          If cboTextAlignment.ItemData(iLoop) = .Fields("alignment") Then
            cboTextAlignment.ListIndex = iLoop
          End If
        Next iLoop

        cboTrimming.ListIndex = IIf(IsNull(.Fields("Trimming")), 0, .Fields("Trimming"))
        
        chkZeroBlank.value = IIf(IsNull(.Fields("blankifZero")), vbUnchecked, IIf(.Fields("blankifZero"), vbChecked, vbUnchecked))
        chkUse1000Separator.value = IIf(IsNull(.Fields("Use1000Separator")), vbUnchecked, IIf(.Fields("Use1000Separator"), vbChecked, vbUnchecked))

        GetLookupColumn = True
      End If
    
      ' JPD20021107 Fault 4708
      ' Refresh the Definition tab page.
      If Not mfReading Then
        RefreshDefinitionTab
      End If
    End With
    
    ' Refresh the data type combo with the new data type selected.
    'JPD 20050812 Fault 10256
    If iOriginalDataType <> miDataType Then mfClearDefault = True
    
    cboDataType_Refresh
    cboDefault_Refresh
    'RefreshControlTab
    mfClearDefault = False
  End If
  
  Exit Function
  
ErrorTrap:
  GetLookupColumn = False
  Err = False

End Function






Private Sub txtErrorMessage_Change()
  If Not mfLoading Then Changed = True

End Sub

Private Sub txtErrorMessage_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub


Private Sub txtListValues_Change()

  If Not mfLoading Then Changed = True

End Sub

Private Sub txtListValues_GotFocus()
  ' Disable the 'Default' property of the 'OK' button as the return key is
  ' used by this textbox.
  cmdOK.Default = False
  
End Sub

Private Sub txtListValues_LostFocus()
  ' Enable the 'Default' property of the OK button.
  cmdOK.Default = True

  ' Refresh the list of possible default values.
  cboDefault_Refresh

End Sub


Private Sub RefreshDefinitionTab()
  'Dim iLoop As Integer
  Dim fEnableDataType As Boolean
  Dim bUnlimitedSize As Boolean
  
  ' Only allow the user to change the column name
  ' if the column is new.
  'lblColumnName.Enabled = True
  'txtColumnName.Enabled = True
  'fraColumnType.Enabled = True
  lblColumnName.Enabled = Not mblnReadOnly
  txtColumnName.Enabled = Not mblnReadOnly
  fraColumnType.Enabled = Not mblnReadOnly
  spnDefaultDisplayWidth.Enabled = Not mblnReadOnly
      
  ' Only allow the user to change the data type if the
  ' column is a 'Data' or 'Calculated' type column.
  fEnableDataType = (miColumnType = giCOLUMNTYPE_DATA) Or _
    (miColumnType = giCOLUMNTYPE_CALCULATED)
        
  bUnlimitedSize = (chkMultiLine.value = vbChecked)
        
  fraDataType.Enabled = fEnableDataType And Not mblnReadOnly
  cboDataType.Enabled = fraDataType.Enabled And Not mblnReadOnly
  cboDataType.BackColor = IIf(cboDataType.Enabled, vbWindowBackground, vbButtonFace)

  ' RH - We should still allow the user to change default display width for lookups,
  '      so enable the frame.
  If miColumnType = giCOLUMNTYPE_LOOKUP And Not mblnReadOnly Then
    fraDataType.Enabled = True
  End If
  
  ' Only allow the user to change the data size control
  ' if the selected data type requires a size.
  asrSize.Enabled = fEnableDataType And Not mblnReadOnly
  asrSize.Visible = ColumnHasSize(miDataType) And Not (chkMultiLine.value = vbChecked And miDataType = dtVARCHAR)
  lblSize.Enabled = asrSize.Enabled And Not mblnReadOnly
  lblSize.Visible = asrSize.Visible
  asrSize.BackColor = IIf(asrSize.Enabled, vbWindowBackground, vbButtonFace)
          
  ' RH 30/01/01 - BUG Numerics must be a maximum of 15 otherwise the
  '               TDBNumber control cannot handle it.
  If (miDataType = dtNUMERIC) And fEnableDataType Then
    If asrSize.value > 15 Then asrSize.value = 15
    asrSize.MaximumValue = 15
  Else
    asrSize.Enabled = Not bUnlimitedSize And Not miColumnType = giCOLUMNTYPE_LOOKUP
    asrDecimals.Enabled = Not bUnlimitedSize
    spnDefaultDisplayWidth.Enabled = Not bUnlimitedSize

    If bUnlimitedSize Then
      asrSize.MaximumValue = VARCHAR_MAX_Size
      asrSize.value = VARCHAR_MAX_Size
      spnDefaultDisplayWidth.MaximumValue = VARCHAR_MAX_Size
      spnDefaultDisplayWidth.value = VARCHAR_MAX_Size
    Else
      asrSize.MaximumValue = 8000
      asrSize.value = IIf(asrSize.value > 8000, 8000, asrSize.value)
      spnDefaultDisplayWidth.MaximumValue = asrSize.MaximumValue
    End If
  End If
  
 
  ' Only allow the user to change the data decimals
  ' control if the selected data type requires decimals.
  asrDecimals.Enabled = fEnableDataType And Not mblnReadOnly
  asrDecimals.Visible = ColumnHasScale(miDataType)
  lblDecimals.Enabled = asrDecimals.Enabled And Not mblnReadOnly
  lblDecimals.Visible = asrDecimals.Visible
  asrDecimals.BackColor = IIf(asrDecimals.Enabled, vbWindowBackground, vbButtonFace)
        
  ' Lookup frame.
  fraLookup.Visible = (miColumnType = giCOLUMNTYPE_LOOKUP)
  If fraLookup.Visible Then
    cboLookupTables.Enabled = (cboLookupTables.ListCount > 0 And (Not mblnReadOnly))
    cboLookupTables.BackColor = IIf(cboLookupTables.Enabled, vbWindowBackground, vbButtonFace)
    lblTable.Enabled = cboLookupTables.Enabled
    
    cboLookupColumns.Enabled = (cboLookupColumns.ListCount > 0 And Not mblnReadOnly)
    cboLookupColumns.BackColor = IIf(cboLookupColumns.Enabled, vbWindowBackground, vbButtonFace)
    lblColumn.Enabled = cboLookupColumns.Enabled
    
    cboLookupFilterColumn.Enabled = (chkLookupFilter.value = vbChecked)
    cboLookupFilterColumn.BackColor = IIf(cboLookupFilterColumn.Enabled, vbWindowBackground, vbButtonFace)
    txtLookupFilterField.Enabled = cboLookupFilterColumn.Enabled
    
    cboLookupFilterOperator.Enabled = (chkLookupFilter.value = vbChecked)
    cboLookupFilterOperator.BackColor = IIf(cboLookupFilterOperator.Enabled, vbWindowBackground, vbButtonFace)
    Label1.Enabled = cboLookupFilterOperator.Enabled
    
    cboLookupFilterValue.Enabled = (chkLookupFilter.value = vbChecked)
    cboLookupFilterValue.BackColor = IIf(cboLookupFilterValue.Enabled, vbWindowBackground, vbButtonFace)
    txtLookupFilterValue.Enabled = cboLookupFilterValue.Enabled
    
    fraLookup.Enabled = cboLookupTables.Enabled And Not mblnReadOnly
  End If
      
  ' Calculation frame.
  fraCalculation.Visible = (miColumnType = giCOLUMNTYPE_CALCULATED)
  If fraCalculation.Visible Then
    
    ' RH 12/03/01 - If frame is disabled, then enabling the button wont work
    '               so need to set the frames enabled state first
    fraCalculation.Enabled = (miDataType <> dtVARBINARY) And _
      (miDataType <> dtLONGVARBINARY)
    
    cmdCalculation.Enabled = fraCalculation.Enabled '(miDataType <> dtVARBINARY) And _
      (miDataType <> dtLONGVARBINARY)
      
    If Not cmdCalculation.Enabled Then
      mlngCalcExprID = 0
      GetCalculationExpressionDetails
    End If
    'fraCalculation.Enabled = cmdCalculation.Enabled
  End If

  ' Link frame.
  fraLink.Visible = (miColumnType = giCOLUMNTYPE_LINK)
  If fraLink.Visible Then
    cboLinkTables.Enabled = (cboLinkTables.ListCount > 0 And Not mblnReadOnly)
    lblLinkTable.Enabled = cboLinkTables.Enabled
    cboLinkTables.BackColor = IIf(cboLinkTables.Enabled, vbWindowBackground, vbButtonFace)
    cmdLinkOrder.Enabled = cboLinkTables.Enabled
    lblLinkOrder.Enabled = cboLinkTables.Enabled And Not mblnReadOnly
    fraLink.Enabled = cboLinkTables.Enabled
  End If
      
  ' JDM - 12/09/01 - Fault 2810 - No display width if OLE object
  'TM20011005 Fault 2917
  'Make the DefaultDislpayWidth spinner and label invisible if the column is
  ' a Link column or the data type is Photo, OLE. Otherwise set the visible and
  'enabled properties to True if not readonly.
  If (miDataType = dtLONGVARBINARY) Or (miDataType = dtVARBINARY) _
    Or (miColumnType = giCOLUMNTYPE_LINK) Or (chkMultiLine.value = vbChecked And miDataType = dtVARCHAR) Then
    Me.lblDefaultDisplayWidth.Enabled = False
    Me.lblDefaultDisplayWidth.Visible = False
    Me.spnDefaultDisplayWidth.Enabled = False
    Me.spnDefaultDisplayWidth.Visible = False
    Me.fraDataType.Refresh
  Else
    Me.lblDefaultDisplayWidth.Enabled = Not mblnReadOnly
    Me.lblDefaultDisplayWidth.Visible = True
    Me.spnDefaultDisplayWidth.Enabled = Not mblnReadOnly And Not bUnlimitedSize
    Me.spnDefaultDisplayWidth.Visible = True
    Me.fraDataType.Refresh
  End If
  
End Sub
Private Sub RefreshControlTab()
  Dim fEnableControlType As Boolean
  Dim fManualListEntry As Boolean
  Dim fSpinnerPropertiesEntry As Boolean
  
  ' Only allow the user to change the control type parameters
  ' if it is a 'Data' or 'Calculated' type column.
'  fEnableControlType = (miColumnType = giCOLUMNTYPE_DATA) Or _
    (miColumnType = giCOLUMNTYPE_CALCULATED) Or _
    (miColumnType = giCOLUMNTYPE_WORKINGPATTERN)
  fEnableControlType = (miColumnType = giCOLUMNTYPE_DATA) Or _
    (miColumnType = giCOLUMNTYPE_CALCULATED)
  
  ' Allow the user to change the control type only if
  ' there is more than one type available.
  If cboControl.ListCount = 1 Or mblnReadOnly Then
    cboControl.Enabled = False
  Else
    cboControl.Enabled = fEnableControlType
  End If
  lblControlType.Enabled = cboControl.Enabled
  cboControl.BackColor = IIf(cboControl.Enabled, vbWindowBackground, vbButtonFace)
      
  fManualListEntry = False
  fSpinnerPropertiesEntry = False
      
  If miColumnType = giCOLUMNTYPE_DATA Then
    Select Case miControlType
      Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
        ' Only display the list item textbox if the selected
        ' control type is 'combo' or 'radio'.
        fManualListEntry = fEnableControlType
      
      Case giCTRL_SPINNER
        ' Only display the spinner properties frame if the selected
        ' control type is 'spinner'.
        fSpinnerPropertiesEntry = fEnableControlType
    End Select
    
  'TM20030123 Fault 4957 - show the spinner controls if it is a Calculated Column also.
  ElseIf miColumnType = giCOLUMNTYPE_CALCULATED Then
    If miControlType = giCTRL_SPINNER Then
      fSpinnerPropertiesEntry = fEnableControlType
    End If
    
  End If
      
  fraListValues.Visible = fManualListEntry
  fraSpinnerProperties.Visible = fSpinnerPropertiesEntry
                  
  If fraListValues.Visible Then
    fraStatusBarMessage.Top = fraListValues.Top + fraListValues.Height + 200
  Else
    If fraSpinnerProperties.Visible Then
      fraStatusBarMessage.Top = fraSpinnerProperties.Top + fraSpinnerProperties.Height + 200
    Else
      fraStatusBarMessage.Top = 690
    End If
  End If

  '02/08/2001 MH Fault 2233
  'If optColumnType(2).Value Or mblnReadOnly Then
  If chkReadOnly.value = vbChecked Or mblnReadOnly Then
    Me.fraStatusBarMessage.Enabled = False
    Me.txtStatusBarMessage.Enabled = False
    Me.txtStatusBarMessage.BackColor = vbButtonFace
  Else
    Me.fraStatusBarMessage.Enabled = True
    Me.txtStatusBarMessage.Enabled = True
    Me.txtStatusBarMessage.BackColor = vbWindowBackground
  End If

End Sub
Private Sub RefreshOptionsTab()
  Dim fEnableOptions As Boolean
  Dim fEnableCase As Boolean
  Dim fEnableAlignment As Boolean
  Dim fEnableTrimming As Boolean
  'Dim iLoop As Integer
  Dim iCount As Integer
  Dim dblOldIntValue As Double
  Dim dblOldNumValue As Double
  Dim dblControlBottom As Double
  Dim sFormat As String
  Dim fEnableDefault As Boolean
  Dim fEnable1000Sep As Boolean
  Dim bEnableStorage As Boolean
 
  Const iGAPY = 200
  
  ' Options frame
  fraOptions.Enabled = True
  
  fEnableOptions = (miColumnType <> giCOLUMNTYPE_LINK) _
      And (miDataType <> dtVARBINARY) _
      And (miDataType <> dtLONGVARBINARY) _
      And Not mblnReadOnly
  
  chkAudit.Enabled = fEnableOptions
  chkAudit.value = IIf(chkAudit.Enabled, chkAudit.value, vbUnchecked)
  
  chkReadOnly.Enabled = fEnableOptions
'  chkReadOnly.value = chkReadOnly.value Or IIf(miControlType = giCTRL_NAVIGATION, 1, 0)
      
  fraOptions.Enabled = fEnableOptions
        
  
  ' Format frame.
  fraFormat.Enabled = True
  
  chkMultiLine.Enabled = (miControlType = giCTRL_TEXTBOX) And _
    (miDataType = dtVARCHAR) And _
    Not mblnReadOnly
  chkZeroBlank.Enabled = (miControlType = giCTRL_TEXTBOX) And _
    (miDataType = dtNUMERIC Or miDataType = dtINTEGER) And _
    Not mblnReadOnly
  
  'JPD 20031017 Fault 7294
  If (miControlType <> giCTRL_TEXTBOX) Or _
    (miDataType <> dtNUMERIC And miDataType <> dtINTEGER) Then
    If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
      chkZeroBlank.value = vbUnchecked
    End If
  End If
  If (miControlType <> giCTRL_TEXTBOX) Or _
    (miDataType <> dtVARCHAR) Then
    If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
      chkMultiLine.value = vbUnchecked
    End If
    If (miControlType = giCTRL_NAVIGATION) Then
      chkMultiLine.value = vbChecked
    End If
  End If
  
  If miControlType = giCTRL_TEXTBOX Then
    fEnableCase = (miDataType = dtVARCHAR)
    fEnableAlignment = True
    fEnableTrimming = True
  ElseIf (miControlType = giCTRL_COMBOBOX) Then
    'NHRD - 15042003 - Fault 4934
    fEnableCase = False ' (miDataType = dtVARCHAR)
    fEnableAlignment = False
    fEnableTrimming = False
  Else
    fEnableCase = False
    fEnableAlignment = (miControlType = giCTRL_SPINNER)
    fEnableTrimming = False
  End If
  
  If Len(txtMask.Text) > 0 Then
    If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
      cboCase.ListIndex = 0
    End If
    cboCase.Enabled = False
    chkMultiLine.Enabled = False
    If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
      chkMultiLine.value = vbUnchecked
    End If
  Else
    If Not fEnableCase Then
      If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
        cboCase.ListIndex = 0
      End If
      cboCase.Enabled = False
    End If
    
    cboCase.Enabled = fEnableCase And Not mblnReadOnly
    chkMultiLine.Enabled = chkMultiLine.Enabled And Not mblnReadOnly
  End If
  lblCase.Enabled = cboCase.Enabled
  
  If Not fEnableAlignment Then
    If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
      cboTextAlignment.ListIndex = 0
    End If
    cboTextAlignment.Enabled = False
  End If
  cboTextAlignment.Enabled = fEnableAlignment And Not mblnReadOnly
  lblTextAlignment.Enabled = cboTextAlignment.Enabled
  
  cboCase.BackColor = IIf(cboCase.Enabled, vbWindowBackground, vbButtonFace)
  cboTextAlignment.BackColor = IIf(cboTextAlignment.Enabled, vbWindowBackground, vbButtonFace)
  
  'MH20030911 Fault 6125
  'chkUse1000Separator.Enabled = (miDataType = dtNUMERIC) And asrSize.Value > 3
  'fEnable1000Sep = (((miDataType = dtNUMERIC) And asrSize.value > 3) Or miDataType = dtinteger) And (Not miColumnType = giCOLUMNTYPE_LOOKUP)
  fEnable1000Sep = (((miDataType = dtNUMERIC) And asrSize.value > 3) Or miDataType = dtINTEGER) _
          And (Not miColumnType = giCOLUMNTYPE_LOOKUP) And (Not miControlType = giCTRL_COLOURPICKER)
  chkUse1000Separator.Enabled = fEnable1000Sep
  If Not fEnable1000Sep And (Not miColumnType = giCOLUMNTYPE_LOOKUP) Then
    chkUse1000Separator.value = vbUnchecked
  End If

  ' Trimming
  If Len(txtMask.Text) > 0 Then
    If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
      cboTrimming.ListIndex = 0
    End If
    cboTrimming.Enabled = False
  Else
    If Not fEnableTrimming Then
      If (miColumnType <> giCOLUMNTYPE_LOOKUP) Then
        If miControlType = giCTRL_WORKINGPATTERN Then
          cboTrimming.ListIndex = 3
        Else
          cboTrimming.ListIndex = 1
        End If
      End If
      cboTrimming.Enabled = False
    End If
    
    cboTrimming.Enabled = fEnableTrimming And Not mblnReadOnly
  End If
  lblTrimming.Enabled = cboTrimming.Enabled
  cboTrimming.BackColor = IIf(cboTrimming.Enabled, vbWindowBackground, vbButtonFace)
   
  fraFormat.Enabled = chkMultiLine.Enabled Or _
    chkZeroBlank.Enabled Or _
    cboCase.Enabled Or _
    cboTextAlignment.Enabled
  
  ' Default frame.
  fraDefault.Enabled = True
  
  fEnableDefault = True
  txtDefault.Visible = False
  ASRDate1.Visible = False
  TDBDefaultNumber.Visible = False
  cboDefault.Visible = False
  fraLogicDefaults.Visible = False
  asrDefault.Visible = False
  ASRDefaultWorkingPattern.Visible = False
  selDefaultColour.Visible = False
  'cmdDefaultColour.Visible = False

  Select Case miControlType
    Case giCTRL_TEXTBOX
      If miDataType = dtTIMESTAMP Then
        ASRDate1.Visible = True
        dblControlBottom = ASRDate1.Top + ASRDate1.Height
      ElseIf miDataType = dtINTEGER Then
        TDBDefaultNumber.Visible = True
        dblOldIntValue = TDBDefaultNumber.value
        TDBDefaultNumber.Format = "##########"
        TDBDefaultNumber.DisplayFormat = TDBDefaultNumber.Format
        TDBDefaultNumber.MaxValue = 2147483647#
        TDBDefaultNumber.MinValue = -2147483648#
        TDBDefaultNumber.value = dblOldIntValue
        dblControlBottom = TDBDefaultNumber.Top + TDBDefaultNumber.Height
      ElseIf miDataType = dtNUMERIC Then
        TDBDefaultNumber.Visible = True
                    
        sFormat = ""
        For iCount = 1 To (Minimum(asrSize.value, 15) - asrDecimals.value)
          sFormat = sFormat & "#"
        Next iCount

        If Len(sFormat) > 0 Then
          sFormat = Left(sFormat, Len(sFormat) - 1) & "0"
        End If
                    
        If asrDecimals.value > 0 Then
          sFormat = sFormat & "."
          For iCount = 1 To asrDecimals.value
            sFormat = sFormat & "0"
          Next iCount
        End If
        
        If Len(sFormat) = 0 Then
          sFormat = "0"
        End If
                    
        dblOldNumValue = TDBDefaultNumber.value
        TDBDefaultNumber.Format = sFormat
        TDBDefaultNumber.DisplayFormat = TDBDefaultNumber.Format
        TDBDefaultNumber.value = dblOldNumValue
        dblControlBottom = TDBDefaultNumber.Top + TDBDefaultNumber.Height
      Else
        txtDefault.Visible = True
        txtDefault.MaxLength = Minimum(val(asrSize.Text), 8000)
        dblControlBottom = txtDefault.Top + txtDefault.Height
      End If
      
    Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
      cboDefault.Visible = True
      cboDefault_Refresh
      dblControlBottom = cboDefault.Top + cboDefault.Height
            
    Case giCTRL_CHECKBOX
      fraLogicDefaults.Visible = True
      dblControlBottom = fraLogicDefaults.Top + fraLogicDefaults.Height
    
    Case giCTRL_SPINNER
      asrDefault.Visible = True
      asrDefault.MinimumValue = val(asrMinVal.Text)
      asrDefault.MaximumValue = val(asrMaxVal.Text)
      asrDefault.Increment = val(asrIncVal.Text)
      dblControlBottom = asrDefault.Top + asrDefault.Height
            
    Case giCTRL_WORKINGPATTERN
      ASRDefaultWorkingPattern.Visible = True
      dblControlBottom = ASRDefaultWorkingPattern.Top + ASRDefaultWorkingPattern.Height
      
    Case giCTRL_NAVIGATION
      fEnableDefault = True
      txtDefault.Visible = True
      txtDefault.MaxLength = Minimum(val(asrSize.Text), 8000)
      dblControlBottom = txtDefault.Top + txtDefault.Height
      
    Case giCTRL_COLOURPICKER
      fEnableDefault = True
      selDefaultColour.Visible = True
      'cmdDefaultColour.Visible = True
      dblControlBottom = selDefaultColour.Top + selDefaultColour.Height
      
    Case Else
      fEnableDefault = False
      txtDefault.Visible = True
      dblControlBottom = txtDefault.Top + txtDefault.Height
  End Select
  
  txtDfltValueExpression.Top = dblControlBottom + 85
  cmdDfltValueExpression.Top = txtDfltValueExpression.Top
  optDfltType(1).Top = txtDfltValueExpression.Top + 60
  dblControlBottom = txtDfltValueExpression.Top + txtDfltValueExpression.Height
  
  fraDefault.Height = dblControlBottom + iGAPY
    
  optDfltType(0).Enabled = fEnableDefault And Not mblnReadOnly
  optDfltType(1).Enabled = optDfltType(0).Enabled

  cmdDfltValueExpression.Enabled = optDfltType(0).Enabled And (optDfltType(1).value)
  txtDefault.Enabled = optDfltType(0).Enabled And optDfltType(0).value
  TDBDefaultNumber.Enabled = optDfltType(0).Enabled And optDfltType(0).value
  ASRDate1.Enabled = optDfltType(0).Enabled And optDfltType(0).value
  cboDefault.Enabled = optDfltType(0).Enabled And optDfltType(0).value
  asrDefault.Enabled = optDfltType(0).Enabled And optDfltType(0).value
  fraLogicDefaults.Enabled = optDfltType(0).Enabled And optDfltType(0).value
  ASRDefaultWorkingPattern.Enabled = optDfltType(0).Enabled And optDfltType(0).value

  txtDefault.BackColor = IIf(txtDefault.Enabled, vbWindowBackground, vbButtonFace)
  TDBDefaultNumber.BackColor = IIf(TDBDefaultNumber.Enabled, vbWindowBackground, vbButtonFace)
  ASRDate1.BackColor = IIf(ASRDate1.Enabled, vbWindowBackground, vbButtonFace)
  cboDefault.BackColor = IIf(cboDefault.Enabled, vbWindowBackground, vbButtonFace)
  asrDefault.BackColor = IIf(asrDefault.Enabled, vbWindowBackground, vbButtonFace)

  fraDefault.Enabled = optDfltType(0).Enabled

  If Not fraDefault.Enabled Then
    mlngDfltValueExprID = 0
    GetDfltValueExpressionDetails
  End If

  
  ' OLE storage frame
  bEnableStorage = (miControlType = giCTRL_OLE Or miControlType = giCTRL_PHOTO)
  fraStorage.Visible = bEnableStorage
  fraDefault.Visible = Not bEnableStorage
  fraStorage.Enabled = bEnableStorage And Not mblnReadOnly
  
  If bEnableStorage Then
  
    ' Only 2 types of OLE for photo
    If miControlType = giCTRL_PHOTO Then
      optOLEStorageType(OLE_SERVER).Caption = "Co&pied to photo directory"
      optOLEStorageType(OLE_LOCAL).Enabled = False
    Else
      optOLEStorageType(OLE_SERVER).Caption = "Co&pied to server OLE directory"
      optOLEStorageType(OLE_LOCAL).Enabled = True
    End If
  
    ' If column has been saved don't allow type to be changed
    If mfIsSaved Then
      For iCount = optOLEStorageType.LBound To optOLEStorageType.UBound
        optOLEStorageType(iCount).Enabled = (optOLEStorageType(iCount).value = True) And Not mblnReadOnly
      Next iCount
    End If
      
    chkEnableOLEMaxSize.Enabled = optOLEStorageType(OLE_EMBEDDED).value = True And Not mblnReadOnly
    lblMaximumOLESize.Enabled = chkEnableOLEMaxSize.value = vbChecked
    asrMaxOLESize.Enabled = chkEnableOLEMaxSize.value = vbChecked
    asrMaxOLESize.BackColor = IIf(asrMaxOLESize.Enabled, vbWindowBackground, vbButtonFace)
    lblMb.Enabled = chkEnableOLEMaxSize.value = vbChecked
  
  End If


End Sub
Private Sub RefreshValidationTab()
  Dim fEnableValidation As Boolean
  Dim fEnableFormat As Boolean
  Dim iTableType As Integer
  Dim fEnableCustomValidation As Boolean
  Dim iLoop As Integer
  
  fraValidationPage.Enabled = True
  
  ' Get the column's table type.
  iTableType = 0
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", mobjColumn.TableID
    
    If Not .NoMatch Then
      iTableType = !TableType
    End If
  End With
  
  ' Standard Validation frame
  fEnableValidation = (miColumnType <> giCOLUMNTYPE_LINK) _
      And (miDataType <> dtVARBINARY) _
      And (miDataType <> dtLONGVARBINARY) _
      And Not mblnReadOnly
  
  fraStandardValidation.Enabled = True
  
  chkDuplicate.Enabled = fEnableValidation
  chkDuplicate.value = IIf(chkDuplicate.Enabled, chkDuplicate.value, vbUnchecked)
  
'TM05052004 - Unique Not Mandatory
'  chkMandatory.Enabled = fEnableValidation And _
'    (chkUnique.Value = vbUnchecked) And _
'    (chkChildUnique.Value = vbUnchecked)
  chkMandatory.Enabled = fEnableValidation
  chkMandatory.value = IIf(chkMandatory.Enabled, chkMandatory.value, vbUnchecked)
  
    'TM20011212 Fault 2761 - Unique checkboxes should be enabled for
    'Lookups and Calculated columns.
'  chkUnique.Enabled = fEnableValidation And _
'    (miColumnType = giCOLUMNTYPE_DATA) And _
'    (miControlType = giCTRL_TEXTBOX)
'  chkChildUnique.Enabled = fEnableValidation And _
'    (miColumnType = giCOLUMNTYPE_DATA) And _
'    (miControlType = giCTRL_TEXTBOX) And _
'    (iTableType = iTabChild)
  chkUnique.Enabled = fEnableValidation And _
    (miColumnType <> giCOLUMNTYPE_LINK) And _
    ((miControlType = giCTRL_TEXTBOX) Or (miControlType = giCTRL_COMBOBOX))
  chkUnique.value = IIf(chkUnique.Enabled, chkUnique.value, vbUnchecked)
    
  chkChildUnique.Enabled = fEnableValidation And _
    (miColumnType <> giCOLUMNTYPE_LINK) And _
    ((miControlType = giCTRL_TEXTBOX) Or (miControlType = giCTRL_COMBOBOX)) And _
    (iTableType = iTabChild)
  chkChildUnique.value = IIf(chkChildUnique.Enabled, chkChildUnique.value, vbUnchecked)

  lstUniqueParents.Enabled = fEnableValidation And _
    (chkChildUnique.value = vbChecked) And _
    (miParentCount > 1)
  lblSiblingParentList.Enabled = lstUniqueParents.Enabled
  
  lstUniqueParents.BackColor = IIf(lstUniqueParents.Enabled, vbWindowBackground, vbButtonFace)
  If (chkChildUnique.value = vbUnchecked) Then
    For iLoop = 0 To lstUniqueParents.ListCount - 1
      lstUniqueParents.Selected(iLoop) = False
    Next iLoop
  End If
  
  fraStandardValidation.Enabled = chkDuplicate.Enabled Or _
    chkMandatory.Enabled Or _
    chkUnique.Enabled Or _
    chkChildUnique.Enabled Or _
    lstUniqueParents.Enabled

  ' Mask validation frame
  fraMask.Enabled = True
  fEnableFormat = (miControlType = giCTRL_TEXTBOX) And Not mblnReadOnly
  
  txtMask.Enabled = fEnableFormat And _
    (miDataType = dtVARCHAR) And IIf(chkMultiLine.value = 1, False, True)

'TM20020403 Fault 3524
  If Not txtMask.Enabled Then
    txtMask.Text = ""
  End If

  txtMask.BackColor = IIf(txtMask.Enabled, vbWindowBackground, vbButtonFace)
  fraMaskKey.Enabled = txtMask.Enabled
  lblMaskKey1.Enabled = txtMask.Enabled
  lblMaskKey2.Enabled = txtMask.Enabled
  lblMaskKey3.Enabled = txtMask.Enabled
  lblMaskKey4.Enabled = txtMask.Enabled
  lblMaskKey5.Enabled = txtMask.Enabled
  lblMaskKey6.Enabled = txtMask.Enabled
    
  fraMask.Enabled = txtMask.Enabled
        
  ' Validation Clause frame.
  fraLostFocusClause.Enabled = True
  fEnableCustomValidation = (chkReadOnly.value = 0) And _
    fEnableValidation And _
    Not mblnReadOnly
  cmdLostFocusClause.Enabled = fEnableCustomValidation
  
  If mlngValidationExprID <= 0 Then txtErrorMessage.Text = ""
  
  txtErrorMessage.Enabled = fEnableCustomValidation And (mlngValidationExprID > 0)
  txtErrorMessage.BackColor = IIf(txtErrorMessage.Enabled, vbWindowBackground, vbButtonFace)
  lblErrorMessage.Enabled = txtErrorMessage.Enabled
  fraLostFocusClause.Enabled = cmdLostFocusClause.Enabled Or _
    txtErrorMessage.Enabled
    
End Sub

Private Sub RefreshDiaryLinkTab()
  Dim fDiaryEnabled As Boolean
  Dim fLinksExist As Boolean
  
  ' Only enable the diary link options if the column data type is 'date'.
  fDiaryEnabled = (miDataType = dtTIMESTAMP)
  fLinksExist = (ssGrdDiaryLinks.Rows > 0)
  
  fraDiary.Enabled = fDiaryEnabled
  ssGrdDiaryLinks.Enabled = fDiaryEnabled   'MH20000724
  cmdDiaryLinkProperties.Enabled = fDiaryEnabled And fLinksExist

  If Not mblnReadOnly Then
    cmdAddDiaryLink.Enabled = fDiaryEnabled
    cmdDiaryLinkProperties.Enabled = fDiaryEnabled And fLinksExist
    cmdRemoveDiaryLink.Enabled = fDiaryEnabled And fLinksExist
    cmdRemoveAllDiaryLinks.Enabled = fDiaryEnabled And fLinksExist
  End If
  CheckIfScrollBarRequiredDiary
  
End Sub

Private Sub RefreshEmailLinkTab()

  Dim fEmailEnabled As Boolean
  Dim fLinksExist As Boolean
  
  'Allow Links for any type of column!
  'Disable this tab if column type is a a link.  'NHRD02122002 Fault 4654
  'TM06082003 Fault 6526 - need the Email Link tab on Link columns to now be enabled.
  
  ' JDM - 02/06/2004 - Fault 8622 - Disable email links for OLE/Photo
  fEmailEnabled = (miColumnType <> giCOLUMNTYPE_LINK) _
      And (miDataType <> dtVARBINARY) _
      And (miDataType <> dtLONGVARBINARY) _
      And Not mblnReadOnly
      
  fLinksExist = (ssGrdEmailLinks.Rows > 0)

  fraEmail.Enabled = fEmailEnabled
  
  ssGrdEmailLinks.Enabled = fEmailEnabled
  cmdEmailLinkProperties.Enabled = fEmailEnabled And fLinksExist
  
  If Not mblnReadOnly Then
    cmdAddEmailLink.Enabled = fEmailEnabled
    cmdEmailLinkProperties.Enabled = fEmailEnabled And fLinksExist
    cmdRemoveEmailLink.Enabled = fEmailEnabled And fLinksExist
    cmdRemoveAllEmailLinks.Enabled = fEmailEnabled And fLinksExist
  End If
  CheckIfScrollBarRequiredEmail

End Sub


Private Function DefaultControl(pDatatype As DataTypes) As ControlTypes
  ' Return the default control for the given data type.
  Select Case pDatatype
    Case dtBIT
      DefaultControl = giCTRL_CHECKBOX
      
    Case dtLONGVARBINARY
      DefaultControl = giCTRL_OLE
      
    Case dtVARBINARY
      DefaultControl = giCTRL_PHOTO
    
    Case dtLONGVARCHAR
      DefaultControl = giCTRL_WORKINGPATTERN
    
    Case Else
      DefaultControl = giCTRL_TEXTBOX
  End Select

End Function

Private Function GetDataType(piReturnType As ExpressionValueTypes) As DataTypes
  ' Return the database type associated with the given
  ' expression return type.
  Select Case piReturnType
    Case giEXPRVALUE_NUMERIC
      GetDataType = dtNUMERIC
    
    Case giEXPRVALUE_LOGIC
      GetDataType = dtBIT
    
    Case giEXPRVALUE_DATE
      GetDataType = dtTIMESTAMP
    
    Case Else
      GetDataType = dtVARCHAR
  End Select

End Function

Private Sub GetValidationExpressionDetails()
  ' Read the Got Focus expression info from the current expression.
  Dim sExprName As String
  Dim objExpr As CExpression
  
  ' Initialise the default values.
  sExprName = vbNullString
    
  ' Instantiate the expression class.
  Set objExpr = New CExpression
    
  With objExpr
    ' Set the expression id.
    .ExpressionID = mlngValidationExprID
      
    ' Read the required info from the expression.
    If .ReadExpressionDetails Then
      sExprName = .Name
    End If
  End With
  
  ' Disassociate object variables.
  Set objExpr = Nothing
    
  ' Update the clause controls properties.
  txtLostFocusClause.Text = sExprName
  
End Sub
Private Sub GetCalculationExpressionDetails()
  
  Dim objExpr As CExpression
  
  ' Read the calculation name from the
  ' current expression, if it is a calculated column.
  If miColumnType = giCOLUMNTYPE_CALCULATED Then
    
    If mlngCalcExprID > 0 Then
      
      ' Instantiate the expression class.
      Set objExpr = New CExpression
        
      With objExpr
        ' Set the expression id.
        .ExpressionID = mlngCalcExprID
        
        ' Read the required info from the expression.
        If .ReadExpressionDetails Then
          ' Update the calculation controls properties.
          txtCalculation.Text = .Name
        End If
       
      End With
    
    Else
      txtCalculation.Text = vbNullString
      
    End If
    
  End If

  ' Disassociate object variables.
  Set objExpr = Nothing

End Sub

Private Sub GetLinkOrderDetails()
  Dim objOrder As Order
  
  If mlngLinkOrderID > 0 Then
  
    ' Instantiate a new Order object.
    Set objOrder = New Order
    objOrder.OrderID = mlngLinkOrderID
    
    ' Read the name of the current order.
    If objOrder.ConstructOrder Then
      txtLinkOrder.Text = objOrder.OrderName
    End If
    
    ' Disassociate object variables.
    Set objOrder = Nothing
  Else
    txtLinkOrder.Text = ""
  End If
  
End Sub


Private Sub GetDfltValueExpressionDetails()
  Dim sExprName As String
  Dim objExpr As CExpression
  
  ' Read the calculation name from the
  ' current expression.
    
  ' Initialise the default values.
  sExprName = vbNullString
  
  ' Instantiate the expression class.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the expression id.
    .ExpressionID = mlngDfltValueExprID
    
    ' Read the required info from the expression.
    If .ReadExpressionDetails Then
      sExprName = .Name
    End If
  End With

  ' Disassociate object variables.
  Set objExpr = Nothing
  
  ' Update the calculation controls properties.
  txtDfltValueExpression.Text = sExprName

End Sub


Private Sub txtMask_Change()
  
  If Not mfLoading Then Changed = True

'  ' RH 06/04/01 - BUG 2095
'  If Len(txtMask.Text) > 0 Then
'    'JPD20010727
'    cboCase.ListIndex = 0
'    cboCase.Enabled = False
'  Else
'    cboCase.ListIndex = 0
'    cboCase.Enabled = Not mblnReadOnly
'    chkMultiLine.Enabled = Not mblnReadOnly
'  End If
  
End Sub

Private Sub txtMask_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub


Private Sub txtMask_KeyUp(KeyCode As Integer, Shift As Integer)

'NHRD20022003 Fault 2712
If Len(txtMask.Text) > asrSize.value Then
  MsgBox "The mask cannot be longer than the Data Type Size." + Chr(13) + Chr(13) & _
    "Data Type Size is currently set at " + CStr(asrSize.value), vbOKOnly + vbExclamation, Application.Name
    txtMask.Text = Left(txtMask.Text, asrSize.value)
End If

End Sub

Private Sub txtMask_LostFocus()
  On Error GoTo ErrCheck
  
  If Len(txtMask.Text) > 0 Then
    If Len(txtMask.Text) > 128 Then
      'Me.tabColProps.Tab = 2
      Me.tabColProps.Tab = iPAGE_VALIDATION
      MsgBox "The mask can be no longer than 128 characters.", vbOKOnly + vbExclamation, Application.Name
      txtMask.SetFocus
      Exit Sub
    Else
      ' Try setting the mask of the hidden field to the mask the user has specified.
      ' If it errors, then the handler captures the error as necessary
      txtMaskTest.Format = txtMask.Text
      
      ' If we get to here, then the mask is ok, so reset the Mask checkbox
      ' as Mulitline/Mask are mutually exclusive.
      chkMultiLine.value = False
      asrSize.value = Len(txtMask.Text)
    End If
  End If
     
  Exit Sub
  
ErrCheck:
  If Err.Number = 380 Then
    MsgBox "You must have at least one user enterable character in the mask field !", vbExclamation + vbOKOnly, "Validation Error"
    txtMaskTest.Format = ""
    'tabColProps.Tab = 2
    tabColProps.Tab = iPAGE_VALIDATION
    txtMask.SetFocus
  Else
    MsgBox "Warning : " & Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, "Validation Error"
    txtMaskTest.Format = ""
    'tabColProps.Tab = 2
    tabColProps.Tab = iPAGE_VALIDATION
    txtMask.SetFocus
  End If
  
End Sub


Private Sub txtStatusBarMessage_Change()

  If Not mfLoading Then Changed = True

End Sub

Private Sub txtStatusBarMessage_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub

'Public Function AfdEnabled() As Boolean
'  ' Returns TRUE if the AFD Module has been enabled.
'  Dim lngCustNo As Long
'  Dim sSQL As String
'  Dim sAuth As String
'  Dim objLicence As ASRLicense.CLicense
'  Dim rsConfig As rdoResultset
'
'  ' Get the Customer number and Module authorisation code.
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsConfig = rdoCon.OpenResultset(sSQL)
'
'  If Not rsConfig.BOF And Not rsConfig.EOF Then
'    lngCustNo = IIf(IsNull(rsConfig!CustNo), 0, rsConfig!CustNo)
'    sAuth = IIf(IsNull(rsConfig!ModuleCode), "", rsConfig!ModuleCode)
'  End If
'
'  rsConfig.Close
'  Set rsConfig = Nothing
'
'  ' Validate the Afd module.
'  Set objLicence = New ASRLicense.CLicense
'  AfdEnabled = objLicence.GetModule(Afd, sAuth, lngCustNo)
'  Set objLicence = Nothing
'
'End Function

Private Function AfdToggleControlStatus(pfValue As Boolean)
  ' Enables/Disables the Afd control fields depending on Value
  ' Two uses:
  '
  ' 1. To enable/disable all Afd controls
  ' 2. To enable/disable just the address fields (either individual or one)
  '    depending on the option button status
  Dim objControl As Control
  'Dim mbAddressType As Boolean
    
  For Each objControl In Me.Controls
    If Not TypeOf objControl Is COA_ColourPicker Then
    If ((TypeOf objControl Is ComboBox) Or (TypeOf objControl Is Label)) And _
      (objControl.Container.Name = "fraFieldMapping") Then
      
      If objControl.Tag = "AFD" Then
        If pfValue Then
          Select Case objControl.Name
            Case "cboAFDProperty", "cboAFDStreet", "cboAFDLocality", "cboAFDTown", "cboAFDCounty", "cboAFDPostcode"
              If optAFDAddressType(0).value Then
                'objControl.Enabled = True
                objControl.Enabled = Not mblnReadOnly
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H80000005
              Else
                objControl.Enabled = False
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H8000000F
              End If
            Case "cboAFDAddress"
              If optAFDAddressType(0).value Then
                objControl.Enabled = False
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H8000000F
              Else
                'objControl.Enabled = True
                objControl.Enabled = Not mblnReadOnly
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H80000005
              End If
            Case Else
              'objControl.Enabled = True
              objControl.Enabled = Not mblnReadOnly
              If TypeOf objControl Is ComboBox Then objControl.BackColor = &H80000005
          End Select
        Else
          objControl.Enabled = False
          If TypeOf objControl Is ComboBox Then objControl.BackColor = &H8000000F
        End If
      End If
    End If
    End If
  Next objControl
  Set objControl = Nothing
  
  optAFDAddressType(0).Enabled = pfValue And Not mblnReadOnly
  optAFDAddressType(1).Enabled = pfValue And Not mblnReadOnly
  
  fraFieldMapping(0).Enabled = pfValue And Not mblnReadOnly
  
End Function

Private Sub FieldMappingInitialiseCombos()
  Dim rsColumns As DAO.Recordset
  Dim sSQL As String
  Dim objControl As Control
  
  For Each objControl In Me.Controls
    If Not TypeOf objControl Is COA_ColourPicker Then
      If TypeOf objControl Is ComboBox And objControl.Container.Name = "fraFieldMapping" Then
        ' Add <None> to the combos.
        With objControl
          .AddItem "<None>"
          .ItemData(.NewIndex) = "0"
          .ListIndex = 0
        End With
      End If
    End If
  Next objControl

  ' Load the columns from the temp tables.
  sSQL = "SELECT tmpcolumns.ColumnID, tmpcolumns.columnName" & _
    " FROM tmpColumns" & _
    " WHERE tmpcolumns.TableID = " & mobjColumn.TableID & _
    " AND tmpcolumns.deleted = False" & _
    " AND tmpcolumns.datatype = " & Trim(Str(dtVARCHAR)) & _
    " AND tmpcolumns.controltype = " & Trim(Str(giCTRL_TEXTBOX)) & _
    " ORDER BY tmpcolumns.columnname"
  
  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  If Not rsColumns.BOF And Not rsColumns.EOF Then
    Do Until rsColumns.EOF
      For Each objControl In Me.Controls
        If Not TypeOf objControl Is COA_ColourPicker Then
          If TypeOf objControl Is ComboBox And objControl.Container.Name = "fraFieldMapping" Then
            'add columnname to the combos
            With objControl
              .AddItem rsColumns("ColumnName")
              .ItemData(.NewIndex) = rsColumns("ColumnID")
            End With
          End If
        End If
      Next objControl
      
      rsColumns.MoveNext
    Loop
  End If
  
  rsColumns.Close
  Set rsColumns = Nothing

End Sub

Private Function SetCombo(combo As ComboBox, colmap As Integer) As Integer
  Dim i As Integer
  
  For i = 0 To combo.ListCount
    If combo.ItemData(i) = colmap Then
      SetCombo = i
      Exit For
    End If
  Next i

End Function


Private Function GetOffset(intOffset As Integer, intTimePeriod As TimePeriods, blnImmediate As Boolean) As String
  
  If blnImmediate = True Then
    GetOffset = "Immediate"
  Else
    If intOffset = 0 Then
      GetOffset = "No offset"
    Else
      GetOffset = _
        CStr(Abs(intOffset)) & " " & _
        TimePeriod(intTimePeriod) & _
        IIf(Abs(intOffset) = 1, "", "s") & _
        IIf(intOffset < 0, " before", " after")
    End If
  End If

End Function


Private Function cboLinkViews_Initialize()

  ' Populate the link tables combo.
  Dim sSQL As String
  Dim rsLinkViews As DAO.Recordset
  
  ' Clear the combo.
  With cboLinkViews
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With

  'NHRD28112002 Fault 4650 - ordered the results of the query
  ' Get the names of parent tables.
  sSQL = "SELECT tmpViews.ViewName, tmpViews.ViewID" & _
    " FROM tmpViews" & _
    " WHERE tmpViews.deleted = FALSE" & _
    " AND tmpViews.viewtableID = " & CStr(mlngLinkTableID) & _
    " ORDER BY tmpViews.ViewName ASC"
  Set rsLinkViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  With rsLinkViews
    Do While Not (.EOF)
      cboLinkViews.AddItem !ViewName
      cboLinkViews.ItemData(cboLinkViews.NewIndex) = !ViewID
      
      rsLinkViews.MoveNext
    Loop
      
    .Close
  End With

  SetComboItem cboLinkViews, mlngLinkViewID
  Set rsLinkViews = Nothing

End Function


Private Sub CheckIfScrollBarRequiredDiary()
  
  With ssGrdDiaryLinks
    If .Rows > 15 Then
      .ScrollBars = ssScrollBarsVertical
      .Columns(0).Width = 3500
    Else
      .ScrollBars = ssScrollBarsNone
      .Columns(0).Width = 3720
    End If
  End With

End Sub

Private Sub CheckIfScrollBarRequiredEmail()

  With ssGrdEmailLinks
    If .Rows > 15 Then
      .ScrollBars = ssScrollBarsVertical
      .Columns("Subject").Width = 3330
    Else
      .ScrollBars = ssScrollBarsNone
      .Columns("Subject").Width = 3550
    End If
  End With

End Sub

Public Sub SetAsNew()

    Changed = True

End Sub

Public Sub PrintDefinition()

  Dim objPrinter As SystemMgr.clsPrintDef
  Dim strColumnType As String
  Dim strTemp As String
  Dim iCount As Integer
  Dim strDefaultValue As String
  Dim bOK As Boolean
  Dim strSize As String
  Dim iTabs As Integer
  'Dim strDigit As String
  Dim fEnableDefault As Boolean
  Dim iTemp As Integer
  
  bOK = True
  On Error GoTo ErrorTrap

  ' Load the printer object
  Set objPrinter = New SystemMgr.clsPrintDef
  With objPrinter
    If .IsOK Then
      If .PrintStart(True) Then
    
        .TabsOnPage = 2
      
        ' Name
        .PrintHeader "Column Name : " & txtColumnName.Text
      
        ' Column type
        For iCount = 0 To optColumnType.Count - 1
          If optColumnType(iCount).value = True Then strColumnType = Replace(optColumnType(iCount).Caption, "&", "")
        Next iCount
    
        .PrintTitle "Definition"
        .PrintNormal "Column Type : " & strColumnType
        .PrintNormal "Data Type : " & cboDataType.Text
        
        iTabs = 0
        If ColumnHasSize(miDataType) Then
          strSize = "Size : " & asrSize.Text & vbTab
          iTabs = 1
        End If
        
        If ColumnHasScale(miDataType) Then
          strSize = strSize & "Decimals : " & asrDecimals.Text
          iTabs = iTabs + 1
        End If
            
        If (miDataType <> dtLONGVARBINARY) And (miDataType <> dtVARBINARY) _
          And (miColumnType <> giCOLUMNTYPE_LINK) Then
          strSize = strSize & IIf(Len(strSize) > 0, vbTab, "") & "Display Width: " & spnDefaultDisplayWidth.Text
          iTabs = iTabs + 1
        End If
        .TabsOnPage = iTabs
          
        If iTabs > 0 Then
          .PrintNormal strSize
        End If
        .TabsOnPage = 2
        
        Select Case miColumnType
          Case giCOLUMNTYPE_LOOKUP
            .PrintNormal "Lookup Table : " & cboLookupTables.Text
            .PrintNormal "Lookup Column : " & cboLookupColumns.Text
            .PrintNormal "Filter Lookup Values : " & IIf(chkLookupFilter.value = vbChecked, "Yes", "No")
            If (chkLookupFilter.value = vbChecked) Then
              .PrintNormal "Filter Column : " & cboLookupFilterColumn.Text
              .PrintNormal "Filter Operator : " & cboLookupFilterOperator.Text
              .PrintNormal "Filter Value : " & cboLookupFilterValue.Text
            End If
            
          Case giCOLUMNTYPE_CALCULATED
            .PrintNormal "Calculation : " & IIf(Len(txtCalculation.Text) = 0, "<None>", txtCalculation.Text)
            'NPG20080502 Fault 13142
            .PrintNormal "Calculate Only If Empty : " & IIf(chkCalculateIfEmpty.value, "Yes", "No")
          
          Case giCOLUMNTYPE_LINK
            .PrintNormal "Link To Parent Table : " & cboLinkTables.Text
            .PrintNormal "Default View : " & cboLinkViews.Text
            .PrintNormal "Default Order : " & txtLinkOrder.Text
      
        End Select
      
        ' Screen Control Tab
        .PrintTitle "Screen Control"
        .TabsOnPage = 2
        
        .PrintNormal "Control Type : " & cboControl.Text
        
        Select Case miControlType
          Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
            .PrintNormal "Control Values : " & IIf(Len(txtListValues.Text) = 0, "<None>", Replace(txtListValues.Text, vbCrLf, ", "))
          
          Case giCTRL_SPINNER
            .PrintNormal "Minimum Value : " & asrMinVal.value
            .PrintNormal "Maximum Value : " & asrMaxVal.value
            .PrintNormal "Increment Value : " & asrIncVal.value
        End Select
      
        If chkReadOnly.value = vbChecked Then
          .PrintNormal "Status Bar Message : N/A"
        Else
          .PrintNormal "Status Bar Message : " & IIf(Len(txtStatusBarMessage.Text) = 0, "<None>", txtStatusBarMessage.Text)
        End If
        
        ' Options Tab
        If miColumnType <> giCOLUMNTYPE_LINK Then
          .PrintTitle "Options"
          .TabsOnPage = 2
          
          If Not (miControlType = giCTRL_OLE Or miControlType = giCTRL_PHOTO) Then
            .PrintNormal "Read Only : " & IIf(chkReadOnly.value = vbChecked, "Yes", "No") & vbTab & "Audit : " & IIf(chkAudit.value = vbChecked, "Yes", "No")
          End If
          
          If miControlType = giCTRL_OLE Or miControlType = giCTRL_PHOTO Then
            
            ' Local
            If optOLEStorageType(0).value = True Then
              .PrintNormal "Storage Type : " & "Copied to local OLE directory"
            
            ' Server / Photo
            ElseIf optOLEStorageType(1).value = True Then
              If miControlType = giCTRL_PHOTO Then
                .PrintNormal "Storage Type : " & "Copied to photo directory"
              Else
                .PrintNormal "Storage Type : " & "Copied to server OLE directory"
              End If
      
            ' Linked / Embedded
            ElseIf optOLEStorageType(2).value = True Then
              .PrintNormal "Storage Type : Linked / Embedded in database"
            
              If chkEnableOLEMaxSize.value = vbChecked Then
                .PrintNormal "Maximum embedding size : " & asrMaxOLESize.value & "KB"
              Else
                .PrintNormal "Maximum embedding size : 0 KB"
              End If
            
            End If
          End If
         
          
          strTemp = ""
          If (miControlType = giCTRL_TEXTBOX) And (miDataType = dtVARCHAR) Then
            strTemp = "Multi-line : " & IIf(chkMultiLine.value = vbChecked, "Yes", "No")
            iCount = iCount + 1
          End If
          If (miControlType = giCTRL_TEXTBOX) And (miDataType = dtNUMERIC Or miDataType = dtINTEGER) Then
            strTemp = strTemp & IIf(Len(strTemp) > 0, vbTab, "") & _
              "Blank if zero : " & IIf(chkZeroBlank.value = vbChecked, "Yes", "No")
          End If
          If Len(strTemp) > 0 Then
            .PrintNormal strTemp
          End If
  
          strTemp = ""
          If (miDataType = dtVARCHAR) And (miControlType = giCTRL_TEXTBOX) Then
            strTemp = "Case : " & cboCase.Text
          End If
          If (miDataType = dtNUMERIC) And (asrSize.value > 3) Then
            strTemp = strTemp & IIf(Len(strTemp) > 0, vbTab, "") & _
              "Use 1000 Separator(,) : " & IIf(chkUse1000Separator.value = vbChecked, "Yes", "No")
          End If
          If Len(strTemp) > 0 Then
            .PrintNormal strTemp
          End If
  
          strTemp = ""
          If (miControlType = giCTRL_SPINNER) Or (miControlType = giCTRL_TEXTBOX) Then
            strTemp = "Alignment : " & cboTextAlignment.Text
          End If
          If miControlType = giCTRL_TEXTBOX Then
            strTemp = strTemp & IIf(Len(strTemp) > 0, vbTab, "") & _
              "Trimming : " & cboTrimming.Text
          End If
          If Len(strTemp) > 0 Then
            .PrintNormal strTemp
          End If
        
          fEnableDefault = True
          strDefaultValue = ""
          
          Select Case miControlType
            Case giCTRL_TEXTBOX
              If miDataType = dtTIMESTAMP Then
                If IsDate(ASRDate1.Text) Then
                  strDefaultValue = ASRDate1.Text
                Else
                  strDefaultValue = "<None>"
                End If
              ElseIf (miDataType = dtINTEGER) Or (miDataType = dtNUMERIC) Then
                strDefaultValue = TDBDefaultNumber.value
              Else
                strDefaultValue = txtDefault.Text
              End If
            Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
              strDefaultValue = cboDefault.Text
            Case giCTRL_CHECKBOX
              strDefaultValue = IIf(optDefault(0).value = True, "True", "False")
            Case giCTRL_SPINNER
              strDefaultValue = TDBDefaultNumber.value
            Case giCTRL_WORKINGPATTERN
              strDefaultValue = ASRDefaultWorkingPattern.value
            Case Else
              fEnableDefault = False
          End Select
  
          If fEnableDefault Then
            If optDfltType(0).value Then
              .PrintNormal "Default Value : " & IIf(Len(strDefaultValue) = 0, "<None>", strDefaultValue)
            Else
              .PrintNormal "Default Calculation : " & IIf(Len(txtDfltValueExpression.Text) = 0, "<None>", txtDfltValueExpression.Text)
            End If
          End If
        End If
        
        ' Validation Tab
        If (miColumnType <> giCOLUMNTYPE_LINK And miControlType <> giCTRL_OLE And miControlType <> giCTRL_PHOTO) Then
          .PrintTitle "Validation"
          
          .PrintNormal "Duplicate Check : " & IIf(chkDuplicate.value = vbChecked, "Yes", "No") & vbTab & "Mandatory : " & IIf(chkMandatory.value = vbChecked, "Yes", "No")
          
          If ((miControlType = giCTRL_TEXTBOX) Or (miControlType = giCTRL_COMBOBOX)) Then
            .PrintNormal "Unique within entire table : " & IIf(chkUnique.value = vbChecked, "Yes", "No") & IIf(lstUniqueParents.ListCount = 0, "", vbTab & "Unique within sibling records : " & IIf(chkChildUnique.value = vbChecked, "Yes", "No"))
          
            If chkChildUnique.value = vbChecked Then
              iTemp = 0
              For iCount = 0 To lstUniqueParents.ListCount - 1
                If lstUniqueParents.Selected(iCount) Then
                  .PrintNormal vbTab & IIf(iTemp = 0, "   Related to : ", "   and : ") & lstUniqueParents.List(iCount)
                  iTemp = iTemp + 1
                End If
              Next iCount
            End If
          End If
        
          If (miDataType = dtVARCHAR) Then
            .PrintNormal "Mask : " & IIf(Len(txtMask.Text) = 0, "<None>", txtMask.Text)
          End If
          
          .PrintNormal "Custom Validation : " & IIf(Len(txtLostFocusClause.Text) = 0, "<None>", txtLostFocusClause.Text)
        
          .PrintNormal "Error Message : " & IIf(Len(txtErrorMessage.Text) = 0, "<None>", txtErrorMessage.Text)
        End If
        
        ' Diary Links Tab
        If (miDataType = dtTIMESTAMP) Then
          .PrintTitle "Diary Links"
        
          If ssGrdDiaryLinks.Rows > 0 Then
            .TabsOnPage = 6
            .PrintBold "Comment" & vbTab & vbTab & vbTab & "Offset" & vbTab & "Alarmed Events"
            
            ssGrdDiaryLinks.MoveFirst
            For iCount = 1 To ssGrdDiaryLinks.Rows
              .PrintNonBold ssGrdDiaryLinks.Columns(0).value & vbTab & vbTab & vbTab _
                  & ssGrdDiaryLinks.Columns(1).value & vbTab & ssGrdDiaryLinks.Columns(2).value
              ssGrdDiaryLinks.MoveNext
            Next iCount
          Else
            .PrintNonBold "<None>"
          End If
        End If
        
'        ' Email Links Tab
'        .PrintTitle "Email Links"
'        If ssGrdEmailLinks.Rows > 0 Then
'          .TabsOnPage = 6
'
'          .PrintBold "Title" & vbTab & vbTab & "Offset" & vbTab & "Subject"
'
'          ssGrdEmailLinks.MoveFirst
'          For iCount = 1 To ssGrdEmailLinks.Rows
'            .PrintNonBold ssGrdEmailLinks.Columns(0).value & vbTab & vbTab _
'                & ssGrdEmailLinks.Columns(1).value & vbTab & ssGrdEmailLinks.Columns(2).value
'            ssGrdDiaryLinks.MoveNext
'          Next iCount
'        Else
'          .PrintNonBold "<None>"
'        End If
        
        ' AFD Postcode
        If IsModuleEnabled(modAFD) Then
          .PrintTitle "Afd"
          
          .PrintNormal "Afd enabled : " & IIf(chkAFDPostCodeColumn.value = vbChecked, "Yes", "No")
          
          If chkAFDPostCodeColumn.value = vbChecked Then
            .TabsOnPage = 2
            
            .PrintNormal "Forename : " & IIf(Len(cboAFDForename.Text) = 0, "<None>", cboAFDForename.Text) & vbTab & "Surname : " & IIf(Len(cboAFDSurname.Text) = 0, "<None>", cboAFDSurname.Text)
            .PrintNormal "Initial(s) : " & IIf(Len(cboAFDInitial.Text) = 0, "<None>", cboAFDInitial.Text) & vbTab & "Telephone : " & IIf(Len(cboAFDTelephone.Text) = 0, "<None>", cboAFDTelephone.Text)
          
            If optAFDAddressType(0).value = True Then
              .PrintNormal "Property : " & IIf(Len(cboAFDProperty.Text) = 0, "<None>", cboAFDProperty.Text) & vbTab & "Street : " & IIf(Len(cboAFDStreet.Text) = 0, "<None>", cboAFDStreet.Text)
              .PrintNormal "Locality : " & IIf(Len(cboAFDLocality.Text) = 0, "<None>", cboAFDLocality.Text) & vbTab & "Town : " & IIf(Len(cboAFDTown.Text) = 0, "<None>", cboAFDTown.Text)
              .PrintNormal "County : " & IIf(Len(cboAFDCounty.Text) = 0, "<None>", cboAFDCounty.Text)
            Else
              .PrintNormal "Address : " & IIf(Len(cboAFDAddress.Text) = 0, "<None>", cboAFDAddress.Text)
            End If
          End If
        End If
        
        ' Quick Address Postcode
        If IsModuleEnabled(modQAddress) Then
          .PrintTitle "Quick Address"
          
          .PrintNormal "Quick Address enabled : " & IIf(chkQAPostCodeColumn.value = vbChecked, "Yes", "No")
          
          If chkQAPostCodeColumn.value = vbChecked Then
            .TabsOnPage = 2
          
            If optQAAddressType(0).value = True Then
              .PrintNormal "Property : " & IIf(Len(cboQAProperty.Text) = 0, "<None>", cboQAProperty.Text) & vbTab & "Street : " & IIf(Len(cboAFDStreet.Text) = 0, "<None>", cboQAStreet.Text)
              .PrintNormal "Locality : " & IIf(Len(cboQALocality.Text) = 0, "<None>", cboQALocality.Text) & vbTab & "Town : " & IIf(Len(cboAFDTown.Text) = 0, "<None>", cboQATown.Text)
              .PrintNormal "County : " & IIf(Len(cboQACounty.Text) = 0, "<None>", cboQACounty.Text)
            Else
              .PrintNormal "Address : " & IIf(Len(cboQAAddress.Text) = 0, "<None>", cboQAAddress.Text)
            End If
          End If
        End If
        
        ' Print the footer
        .PrintEnd       ' this adds the correct footer
      End If
    End If
  End With
  Set objPrinter = Nothing

TidyUpAndExit:
  If Not bOK Then
    MsgBox "Unable to print the table definition." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit


End Sub

Public Sub CopyDefinitionToClipboard()

  Dim strClipboardText As String
  Dim strColumnType As String
  'Dim strTemp As String
  Dim iCount As Integer
  Dim strDefaultValue As String
  Dim fEnableDefault As Boolean
  Dim iTemp As Integer
  
  ' Name
  strClipboardText = "Column Name : " & txtColumnName.Text & vbCrLf

  ' Table type
  For iCount = 0 To optColumnType.Count - 1
    If optColumnType(iCount).value = True Then strColumnType = Replace(optColumnType(iCount).Caption, "&", "")
  Next iCount
  
  strClipboardText = strClipboardText & vbCrLf & "Definition" & vbCrLf
  strClipboardText = strClipboardText & "----------" & vbCrLf & vbCrLf
  strClipboardText = strClipboardText & "Column Type : " & strColumnType & vbCrLf
  strClipboardText = strClipboardText & "Data Type : " & cboDataType.Text & vbCrLf
  
  If ColumnHasSize(miDataType) Then
    strClipboardText = strClipboardText & "Size : " & asrSize.Text & vbCrLf
  End If
  
  If ColumnHasScale(miDataType) Then
    strClipboardText = strClipboardText & "Decimals : " & asrDecimals.Text & vbCrLf
  End If
  
  If (miDataType <> dtLONGVARBINARY) And (miDataType <> dtVARBINARY) _
    And (miColumnType <> giCOLUMNTYPE_LINK) Then
    strClipboardText = strClipboardText & "Display Width: " & spnDefaultDisplayWidth.Text & vbCrLf
  End If
  
  Select Case miColumnType
    Case giCOLUMNTYPE_LOOKUP
      strClipboardText = strClipboardText & "Lookup Table : " & cboLookupTables.Text & vbCrLf
      strClipboardText = strClipboardText & "Lookup Column : " & cboLookupColumns.Text & vbCrLf
      strClipboardText = strClipboardText & "Filter Lookup Values : " & IIf(chkLookupFilter.value = vbChecked, "Yes", "No") & vbCrLf
      If (chkLookupFilter.value = vbChecked) Then
        strClipboardText = strClipboardText & "Filter Column : " & cboLookupFilterColumn.Text & vbCrLf
        strClipboardText = strClipboardText & "Filter Operator : " & cboLookupFilterOperator.Text & vbCrLf
        strClipboardText = strClipboardText & "Filter Value : " & cboLookupFilterValue.Text & vbCrLf
      End If
    
    Case giCOLUMNTYPE_CALCULATED
      strClipboardText = strClipboardText & "Calculation : " & IIf(Len(txtCalculation.Text) = 0, "<None>", txtCalculation.Text) & vbCrLf
      'NPG20080502 Fault 13142
      strClipboardText = strClipboardText & "Calculate Only If Empty : " & IIf(chkCalculateIfEmpty, "Yes", "No") & vbCrLf
      
    Case giCOLUMNTYPE_LINK
      strClipboardText = strClipboardText & "Link To Parent Table : " & cboLinkTables.Text & vbCrLf
      strClipboardText = strClipboardText & "Default View : " & cboLinkViews.Text & vbCrLf
      strClipboardText = strClipboardText & "Default Order : " & txtLinkOrder.Text & vbCrLf

  End Select

  ' Screen Control Tab
  strClipboardText = strClipboardText & vbCrLf & "Screen Control" & vbCrLf
  strClipboardText = strClipboardText & "--------------" & vbCrLf
  
  strClipboardText = strClipboardText & vbCrLf & "Control Type : " & cboControl.Text & vbCrLf
  
  Select Case miControlType
    Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
      strClipboardText = strClipboardText & "Control Values : " & IIf(Len(txtListValues.Text) = 0, "<None>", Replace(txtListValues.Text, vbCrLf, ", ")) & vbCrLf
    
    Case giCTRL_SPINNER
      strClipboardText = strClipboardText & "Minimum Value : " & asrMinVal.value & vbCrLf
      strClipboardText = strClipboardText & "Maximum Value : " & asrMaxVal.value & vbCrLf
      strClipboardText = strClipboardText & "Increment Value : " & asrIncVal.value & vbCrLf
  End Select

  If chkReadOnly.value = vbChecked Then
    strClipboardText = strClipboardText & "Status Bar Message : N/A" & vbCrLf
  Else
    strClipboardText = strClipboardText & "Status Bar Message : " & IIf(Len(txtStatusBarMessage.Text) = 0, "<None>", txtStatusBarMessage.Text) & vbCrLf
  End If
  
  ' Options Tab
  If miColumnType <> giCOLUMNTYPE_LINK Then
    strClipboardText = strClipboardText & vbCrLf & "Options" & vbCrLf
    strClipboardText = strClipboardText & "-------" & vbCrLf
    
    strClipboardText = strClipboardText & "Read Only : " & IIf(chkReadOnly.value = vbChecked, "Yes", "No") & vbCrLf
    strClipboardText = strClipboardText & "Audit : " & IIf(chkAudit.value = vbChecked, "Yes", "No") & vbCrLf
    
    
    If miControlType = giCTRL_OLE Or miControlType = giCTRL_PHOTO Then
      
      ' Local
      If optOLEStorageType(0).value = True Then
        strClipboardText = strClipboardText & "Storage Type : " & "Copied to local OLE directory" & vbCrLf
      
      ' Server / Photo
      ElseIf optOLEStorageType(1).value = True Then
        If miControlType = giCTRL_PHOTO Then
          strClipboardText = strClipboardText & "Storage Type : " & "Copied to photo directory" & vbCrLf
        Else
          strClipboardText = strClipboardText & "Storage Type : " & "Copied to server OLE directory" & vbCrLf
        End If

      ' Linked / Embedded
      ElseIf optOLEStorageType(2).value = True Then
        strClipboardText = strClipboardText & "Storage Type : Linked / Embedded in database" & vbCrLf
      
        If chkEnableOLEMaxSize.value = vbChecked Then
          strClipboardText = strClipboardText & "Maximum embedding size : " & asrMaxOLESize.value & "KB" & vbCrLf
        Else
          strClipboardText = strClipboardText & "Maximum embedding size : 0 KB" & vbCrLf
        End If
      
      End If
    End If
      
    If (miControlType = giCTRL_TEXTBOX) And (miDataType = dtVARCHAR) Then
      strClipboardText = strClipboardText & vbCrLf & "Multi-line : " & IIf(chkMultiLine.value = vbChecked, "Yes", "No") & vbCrLf
    End If
    If (miControlType = giCTRL_TEXTBOX) And (miDataType = dtNUMERIC Or miDataType = dtINTEGER) Then
      strClipboardText = strClipboardText & "Blank if zero : " & IIf(chkZeroBlank.value = vbChecked, "Yes", "No") & vbCrLf
    End If
    If (miDataType = dtVARCHAR) And (miControlType = giCTRL_TEXTBOX) Then
      strClipboardText = strClipboardText & "Case : " & cboCase.Text & vbCrLf
    End If
    If (miDataType = dtNUMERIC) And (asrSize.value > 3) Then
      strClipboardText = strClipboardText & "Use 1000 Separator : " & IIf(chkUse1000Separator.value = vbChecked, "Yes", "No")
    End If
    If (miControlType = giCTRL_SPINNER) Or (miControlType = giCTRL_TEXTBOX) Then
      strClipboardText = strClipboardText & "Alignment : " & cboTextAlignment.Text & vbCrLf
    End If
    If miControlType = giCTRL_TEXTBOX Then
      strClipboardText = strClipboardText & "Trimming : " & cboTrimming.Text & vbCrLf
    End If
    
    fEnableDefault = True
    strDefaultValue = ""
    
    Select Case miControlType
      Case giCTRL_TEXTBOX
        If miDataType = dtTIMESTAMP Then
          If IsDate(ASRDate1.Text) Then
            strDefaultValue = ASRDate1.Text
          Else
            strDefaultValue = "<None>"
          End If
        ElseIf (miDataType = dtINTEGER) Or (miDataType = dtNUMERIC) Then
          strDefaultValue = TDBDefaultNumber.value
        Else
          strDefaultValue = txtDefault.Text
        End If
      Case giCTRL_COMBOBOX, giCTRL_OPTIONGROUP
        strDefaultValue = cboDefault.Text
      Case giCTRL_CHECKBOX
        strDefaultValue = IIf(optDefault(0).value = True, "True", "False")
      Case giCTRL_SPINNER
        strDefaultValue = TDBDefaultNumber.value
      Case giCTRL_WORKINGPATTERN
        strDefaultValue = ASRDefaultWorkingPattern.value
      Case Else
        fEnableDefault = False
    End Select

    If fEnableDefault Then
      If optDfltType(0).value Then
        strClipboardText = strClipboardText & vbCrLf & "Default Value : " & IIf(Len(strDefaultValue) = 0, "<None>", strDefaultValue) & vbCrLf
      Else
        strClipboardText = strClipboardText & vbCrLf & "Default Calculation : " & IIf(Len(txtDfltValueExpression.Text) = 0, "<None>", txtDfltValueExpression.Text) & vbCrLf
        
      End If
    End If
  End If
  
  ' Validation Tab
  If (miColumnType <> giCOLUMNTYPE_LINK) Then
    strClipboardText = strClipboardText & vbCrLf & "Validation" & vbCrLf
    strClipboardText = strClipboardText & "----------" & vbCrLf
    
    strClipboardText = strClipboardText & vbCrLf & "Duplicate Check : " & IIf(chkDuplicate.value = vbChecked, "Yes", "No") & vbCrLf
    strClipboardText = strClipboardText & "Mandatory : " & IIf(chkMandatory.value = vbChecked, "Yes", "No") & vbCrLf
    
    If ((miControlType = giCTRL_TEXTBOX) Or (miControlType = giCTRL_COMBOBOX)) Then
      strClipboardText = strClipboardText & "Unique within entire table : " & IIf(chkUnique.value = vbChecked, "Yes", "No") & vbCrLf
      
      If lstUniqueParents.ListCount > 0 Then
        strClipboardText = strClipboardText & "Unique within sibling records : " & IIf(chkChildUnique.value = vbChecked, "Yes", "No") & vbCrLf
      
        If chkChildUnique.value = vbChecked Then
          For iCount = 0 To lstUniqueParents.ListCount - 1
            If lstUniqueParents.Selected(iCount) Then
              strClipboardText = strClipboardText & IIf(iTemp = 0, "   Related to : ", "   and : ") & lstUniqueParents.List(iCount) & vbCrLf
              iTemp = iTemp + 1
            End If
          Next iCount
        End If
      End If
    End If
    
    If (miDataType = dtVARCHAR) Then
      strClipboardText = strClipboardText & "Mask : " & IIf(Len(txtMask.Text) = 0, "<None>", txtMask.Text) & vbCrLf
    End If
    
    strClipboardText = strClipboardText & "Custom Validation : " & IIf(Len(txtLostFocusClause.Text) = 0, "<None>", txtLostFocusClause.Text) & vbCrLf
  
    strClipboardText = strClipboardText & "Error Message : " & IIf(Len(txtErrorMessage.Text) = 0, "<None>", txtErrorMessage.Text) & vbCrLf
  End If
  
  ' Diary Links Tab
  If (miDataType = dtTIMESTAMP) Then
    strClipboardText = strClipboardText & vbCrLf & "Diary Links" & vbCrLf
    strClipboardText = strClipboardText & "-----------" & vbCrLf
    
    If ssGrdDiaryLinks.Rows > 0 Then
      
      strClipboardText = strClipboardText & vbCrLf & "Comment" & vbTab & "Offset" & vbTab & "Alarmed Events" & vbCrLf
      
      ssGrdDiaryLinks.MoveFirst
      For iCount = 1 To ssGrdDiaryLinks.Rows
        strClipboardText = strClipboardText & ssGrdDiaryLinks.Columns(0).value & vbTab _
            & ssGrdDiaryLinks.Columns(1).value & vbTab & ssGrdDiaryLinks.Columns(2).value & vbCrLf
        ssGrdDiaryLinks.MoveNext
      Next iCount
    Else
      strClipboardText = strClipboardText & vbCrLf & "<None>" & vbCrLf
    End If
  End If
  
'  ' Email Links Tab
'  strClipboardText = strClipboardText & vbCrLf & "Email Links" & vbCrLf
'  strClipboardText = strClipboardText & "-----------" & vbCrLf
'
'  If ssGrdEmailLinks.Rows > 0 Then
'
'    strClipboardText = strClipboardText & vbCrLf & "Title" & vbTab & "Offset" & vbTab & "Subject" & vbCrLf
'
'    ssGrdEmailLinks.MoveFirst
'    For iCount = 1 To ssGrdEmailLinks.Rows
'      strClipboardText = strClipboardText & ssGrdEmailLinks.Columns(0).value & vbTab _
'          & ssGrdEmailLinks.Columns(1).value & vbTab & ssGrdEmailLinks.Columns(2).value & vbCrLf
'      ssGrdDiaryLinks.MoveNext
'    Next iCount
'  Else
'    strClipboardText = strClipboardText & vbCrLf & "<None>" & vbCrLf
'  End If
  
  ' AFD Postcode
  If gbAFDEnabled Then
    strClipboardText = strClipboardText & vbCrLf & "Afd" & vbCrLf
    strClipboardText = strClipboardText & "---" & vbCrLf & vbCrLf
    strClipboardText = strClipboardText & "Afd enabled : " & IIf(chkAFDPostCodeColumn.value = vbChecked, "Yes", "No") & vbCrLf
  
    If chkAFDPostCodeColumn.value = vbChecked Then
      strClipboardText = strClipboardText & "Forename : " & IIf(Len(cboAFDForename.Text) = 0, "<None>", cboAFDForename.Text) & vbCrLf
      strClipboardText = strClipboardText & "Surname : " & IIf(Len(cboAFDSurname.Text) = 0, "<None>", cboAFDSurname.Text) & vbCrLf
      strClipboardText = strClipboardText & "Initial(s) : " & IIf(Len(cboAFDInitial.Text) = 0, "<None>", cboAFDInitial.Text) & vbCrLf
      strClipboardText = strClipboardText & "Telephone : " & IIf(Len(cboAFDTelephone.Text) = 0, "<None>", cboAFDTelephone.Text) & vbCrLf
    
      If optAFDAddressType(0).value = True Then
        strClipboardText = strClipboardText & "Property : " & IIf(Len(cboAFDProperty.Text) = 0, "<None>", cboAFDProperty.Text) & vbCrLf
        strClipboardText = strClipboardText & "Street : " & IIf(Len(cboAFDStreet.Text) = 0, "<None>", cboAFDStreet.Text) & vbCrLf
        strClipboardText = strClipboardText & "Locality : " & IIf(Len(cboAFDLocality.Text) = 0, "<None>", cboAFDLocality.Text) & vbCrLf
        strClipboardText = strClipboardText & "Town : " & IIf(Len(cboAFDTown.Text) = 0, "<None>", cboAFDTown.Text) & vbCrLf
        strClipboardText = strClipboardText & "County : " & IIf(Len(cboAFDCounty.Text) = 0, "<None>", cboAFDCounty.Text) & vbCrLf
      Else
        strClipboardText = strClipboardText & "Address : " & IIf(Len(cboAFDAddress.Text) = 0, "<None>", cboAFDAddress.Text) & vbCrLf
      End If
    
    End If
  End If

  ' Quick Address Postcode
  If gbQAddressEnabled Then
    strClipboardText = strClipboardText & vbCrLf & "Quick Address" & vbCrLf
    strClipboardText = strClipboardText & "---" & vbCrLf & vbCrLf
    strClipboardText = strClipboardText & "Quick Address enabled : " & IIf(chkQAPostCodeColumn.value = vbChecked, "Yes", "No") & vbCrLf
  
    If chkQAPostCodeColumn.value = vbChecked Then
      If optQAAddressType(0).value = True Then
        strClipboardText = strClipboardText & "Property : " & IIf(Len(cboQAProperty.Text) = 0, "<None>", cboQAProperty.Text) & vbCrLf
        strClipboardText = strClipboardText & "Street : " & IIf(Len(cboQAStreet.Text) = 0, "<None>", cboQAStreet.Text) & vbCrLf
        strClipboardText = strClipboardText & "Locality : " & IIf(Len(cboQALocality.Text) = 0, "<None>", cboQALocality.Text) & vbCrLf
        strClipboardText = strClipboardText & "Town : " & IIf(Len(cboQATown.Text) = 0, "<None>", cboQATown.Text) & vbCrLf
        strClipboardText = strClipboardText & "County : " & IIf(Len(cboQACounty.Text) = 0, "<None>", cboQACounty.Text) & vbCrLf
      Else
        strClipboardText = strClipboardText & "Address : " & IIf(Len(cboQAAddress.Text) = 0, "<None>", cboQAAddress.Text) & vbCrLf
      End If
    
    End If
  End If

  ' Put the info in the clipboard
  Clipboard.Clear
  Clipboard.SetText strClipboardText

End Sub

Private Sub cboLookupFilterColumn_Refresh()
  ' Refresh the Lookup Columns filter combo.
  Dim iIndex As Integer
  
  iIndex = 0
  
  With cboLookupFilterColumn
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = "0"
    .ListIndex = 0
  End With
  
  If (miColumnType = giCOLUMNTYPE_LOOKUP) And _
    (mlngLookupColumnID > 0) Then
  
    ' Loop through columns for selected lookup table.
    With recColEdit
      .Index = "idxName"
      .Seek ">=", mlngLookupTableID
      
      If Not .NoMatch Then
        Do While Not .EOF
          If .Fields("tableID") <> mlngLookupTableID Then
            Exit Do
          End If
          
          ' Add each column name to the lookup columns combo.
          ' NB. We only want to add certain types of column. There's not use in
          ' looking up OLE or logic values.
          'If (.Fields("columnType") <> giCOLUMNTYPE_SYSTEM) And _
            (.Fields("columnType") <> giCOLUMNTYPE_LINK) And _
            (Not .Fields("deleted")) And _
            (.Fields("dataType") <> dtLONGVARBINARY) And _
            (.Fields("dataType") <> dtVARBINARY) Then
          If (.Fields("columnType") <> giCOLUMNTYPE_SYSTEM) And _
            (.Fields("columnType") <> giCOLUMNTYPE_LINK) And _
            (Not .Fields("deleted")) And _
            (.Fields("dataType") <> dtLONGVARBINARY) And _
            (.Fields("dataType") <> dtVARBINARY) And _
            (.Fields("controlType") <> ControlTypes.giCTRL_COLOURPICKER) Then
            
            cboLookupFilterColumn.AddItem .Fields("columnName")
            cboLookupFilterColumn.ItemData(cboLookupFilterColumn.NewIndex) = .Fields("columnID")
        
            If .Fields("columnID") = mlngLookupFilterColumnID Then
              iIndex = cboLookupFilterColumn.NewIndex
              mlngLookupFilterColumnType = .Fields("DataType")
            End If
          End If
      
          .MoveNext
        Loop
      End If
    End With
  End If
  
  ' Enable the combo if there are items.
  With cboLookupFilterColumn
    If .ListCount > 0 Then
      .ListIndex = iIndex
      '.Enabled = True
      '.Enabled = Not mblnReadOnly
    Else
      .Enabled = False
    End If
  End With
  
  Exit Sub
  
End Sub

Private Sub chkLookupFilter_Click()
  If Not mfLoading Then
    cboLookupFilterColumn_Refresh
    cboLookupFilterValue_Refresh
    RefreshCurrentTab
  End If
End Sub

Private Sub cboLookupFilterValue_Refresh()
  Dim rsColumns As DAO.Recordset
  Dim sSQL As String
  'Dim objControl As Control

  With cboLookupFilterValue
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = "0"
    .ListIndex = 0
  End With

  ' Load the columns from the temp tables.
  sSQL = "SELECT tmpcolumns.ColumnID, tmpcolumns.columnName" & _
    " FROM tmpColumns" & _
    " WHERE tmpcolumns.TableID = " & mobjColumn.TableID & _
    " AND tmpcolumns.deleted = False" & _
    " AND tmpcolumns.datatype = " & Trim(Str(mlngLookupFilterColumnType)) & _
    " AND columnType <> " & CStr(giCOLUMNTYPE_SYSTEM) & _
    " AND columnType <> " & CStr(giCOLUMNTYPE_LINK) & _
    " ORDER BY tmpcolumns.columnname"
  
  Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  If Not rsColumns.BOF And Not rsColumns.EOF Then
    Do Until rsColumns.EOF
      'add columnname to the combos
      With cboLookupFilterValue
        .AddItem rsColumns("ColumnName")
        .ItemData(.NewIndex) = rsColumns("ColumnID")
      End With

      rsColumns.MoveNext
    Loop
  End If
  
  rsColumns.Close
  Set rsColumns = Nothing

End Sub


Private Function QAToggleControlStatus(pfValue As Boolean)
  ' Enables/Disables the QA control fields depending on Value
  ' Two uses:
  '
  ' 1. To enable/disable all QA controls
  ' 2. To enable/disable just the address fields (either individual or one)
  '    depending on the option button status
  Dim objControl As Control
  
  For Each objControl In Me.Controls
    If Not TypeOf objControl Is COA_ColourPicker Then
    If ((TypeOf objControl Is ComboBox) Or (TypeOf objControl Is Label)) And _
      (objControl.Container.Name = "fraFieldMapping") Then
      
      If objControl.Tag = "QA" Then
        If pfValue Then
          Select Case objControl.Name
            Case "cboQAProperty", "cboQAStreet", "cboQALocality", "cboQATown", "cboQACounty", "cboQAPostcode"
              If optQAAddressType(0).value Then
                'objControl.Enabled = True
                objControl.Enabled = Not mblnReadOnly
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H80000005
              Else
                objControl.Enabled = False
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H8000000F
              End If
            Case "cboQAAddress"
              If optQAAddressType(0).value Then
                objControl.Enabled = False
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H8000000F
              Else
                'objControl.Enabled = True
                objControl.Enabled = Not mblnReadOnly
                If TypeOf objControl Is ComboBox Then objControl.BackColor = &H80000005
              End If
            Case Else
              'objControl.Enabled = True
              objControl.Enabled = Not mblnReadOnly
              If TypeOf objControl Is ComboBox Then objControl.BackColor = &H80000005
          End Select
        Else
          objControl.Enabled = False
          If TypeOf objControl Is ComboBox Then objControl.BackColor = &H8000000F
        End If
      End If
    End If
    End If
  Next objControl
  Set objControl = Nothing
  
  optQAAddressType(0).Enabled = pfValue And Not mblnReadOnly
  optQAAddressType(1).Enabled = pfValue And Not mblnReadOnly
  
  fraFieldMapping(1).Enabled = pfValue And Not mblnReadOnly
  
End Function


