VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "COA_ColourPicker.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Object = "{96E404DC-B217-4A2D-A891-C73A92A628CC}#1.0#0"; "COA_WorkingPattern.ocx"
Begin VB.Form frmWorkflowTimeout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web Form Item Properties"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5075
   Icon            =   "frmWorkflowTimeout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin COAColourPicker.COA_ColourPicker colPickDlg 
      Left            =   2280
      Top             =   10320
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   6195
      TabIndex        =   62
      Top             =   10320
      Width           =   2600
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   64
         Top             =   0
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab ssTabStrip 
      Height          =   10140
      Left            =   90
      TabIndex        =   65
      Top             =   105
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   17886
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frmWorkflowTimeout.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTabContainer(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Appea&rance"
      TabPicture(1)   =   "frmWorkflowTimeout.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picTabContainer(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Data"
      TabPicture(2)   =   "frmWorkflowTimeout.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picTabContainer(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picTabContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   9210
         Index           =   2
         Left            =   -74850
         ScaleHeight     =   9210
         ScaleWidth      =   8400
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   400
         Width           =   8400
         Begin VB.Frame fraRecordIdentification 
            Caption         =   "Record Identification :"
            Height          =   3200
            Left            =   0
            TabIndex        =   121
            Top             =   0
            Width           =   2800
            Begin VB.ComboBox cboRecordIdentificationRecordSelector 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0060
               Left            =   1800
               List            =   "frmWorkflowTimeout.frx":0062
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   1500
               Width           =   500
            End
            Begin VB.ComboBox cboRecordIdentificationElement 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0064
               Left            =   1800
               List            =   "frmWorkflowTimeout.frx":0066
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   127
               Top             =   1100
               Width           =   500
            End
            Begin VB.ComboBox cboRecordIdentificationRecord 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0068
               Left            =   1800
               List            =   "frmWorkflowTimeout.frx":006A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   125
               Top             =   700
               Width           =   500
            End
            Begin VB.ComboBox cboRecordIdentificationTable 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   123
               Top             =   300
               Width           =   500
            End
            Begin VB.ComboBox cboRecordIdentificationRecordTable 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   131
               Top             =   1900
               Width           =   500
            End
            Begin VB.CommandButton cmdRecordIdentificationOrder 
               Caption         =   "..."
               Height          =   315
               Left            =   2300
               TabIndex        =   134
               Top             =   2300
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtRecordIdentificationOrder 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   2300
               Width           =   500
            End
            Begin VB.CommandButton cmdRecordIdentificationFilter 
               Caption         =   "..."
               Height          =   315
               Left            =   2300
               TabIndex        =   137
               Top             =   2700
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtRecordIdentificationFilter 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   2700
               Width           =   500
            End
            Begin VB.Label lblRecordIdentificationRecordSelector 
               Caption         =   "Record Selector :"
               Height          =   195
               Left            =   195
               TabIndex        =   128
               Top             =   1560
               Width           =   1515
            End
            Begin VB.Label lblRecordIdentificationElement 
               Caption         =   "Element :"
               Height          =   195
               Left            =   200
               TabIndex        =   126
               Top             =   1160
               Width           =   840
            End
            Begin VB.Label lblRecordIdentificationRecord 
               Caption         =   "Record :"
               Height          =   195
               Left            =   195
               TabIndex        =   124
               Top             =   765
               Width           =   930
            End
            Begin VB.Label lblRecordIdentificationTable 
               Caption         =   "Table :"
               Height          =   195
               Left            =   195
               TabIndex        =   122
               Top             =   360
               Width           =   810
            End
            Begin VB.Label lblRecordIdentificationRecordTable 
               Caption         =   "Record Table :"
               Height          =   195
               Left            =   195
               TabIndex        =   130
               Top             =   1965
               Width           =   1320
            End
            Begin VB.Label lblRecordIdentificationFilter 
               Caption         =   "Filter :"
               Height          =   195
               Left            =   195
               TabIndex        =   135
               Top             =   2760
               Width           =   735
            End
            Begin VB.Label lblRecordIdentificationOrder 
               Caption         =   "Order :"
               Height          =   195
               Left            =   195
               TabIndex        =   132
               Top             =   2355
               Width           =   795
            End
         End
         Begin VB.Frame fraValidation 
            Caption         =   "Validation :"
            Height          =   2535
            Left            =   120
            TabIndex        =   176
            Top             =   4200
            Width           =   5500
            Begin VB.TextBox txtFileExtensions 
               Height          =   375
               Left            =   4320
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   184
               Top             =   2040
               Width           =   975
            End
            Begin VB.CommandButton cmdValidationEdit 
               Caption         =   "&Edit"
               Enabled         =   0   'False
               Height          =   400
               Left            =   1500
               TabIndex        =   179
               Top             =   1605
               Width           =   1200
            End
            Begin VB.CommandButton cmdValidationDeleteAll 
               Caption         =   "Delete &All"
               Enabled         =   0   'False
               Height          =   400
               Left            =   4100
               TabIndex        =   181
               Top             =   1605
               Width           =   1200
            End
            Begin VB.CommandButton cmdValidationAdd 
               Caption         =   "&New"
               Enabled         =   0   'False
               Height          =   400
               Left            =   200
               TabIndex        =   178
               Top             =   1605
               Width           =   1200
            End
            Begin VB.CommandButton cmdValidationDelete 
               Caption         =   "&Delete"
               Enabled         =   0   'False
               Height          =   400
               Left            =   2800
               TabIndex        =   180
               Top             =   1605
               Width           =   1200
            End
            Begin VB.CheckBox chkValidationMandatory 
               Caption         =   "&Mandatory"
               Height          =   195
               Left            =   200
               TabIndex        =   182
               Top             =   2160
               Width           =   1410
            End
            Begin SSDataWidgets_B.SSDBGrid grdValidation 
               Height          =   1140
               Left            =   200
               TabIndex        =   177
               Top             =   360
               Width           =   5100
               _Version        =   196617
               DataMode        =   2
               RecordSelectors =   0   'False
               Col.Count       =   5
               DefColWidth     =   26458
               CheckBox3D      =   0   'False
               AllowUpdate     =   0   'False
               MultiLine       =   0   'False
               AllowRowSizing  =   0   'False
               AllowGroupSizing=   0   'False
               AllowColumnSizing=   0   'False
               AllowGroupMoving=   0   'False
               AllowColumnMoving=   2
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
               ExtraHeight     =   79
               Columns.Count   =   5
               Columns(0).Width=   26458
               Columns(0).Visible=   0   'False
               Columns(0).Caption=   "ExprID"
               Columns(0).Name =   "ExprID"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               Columns(1).Width=   3519
               Columns(1).Caption=   "Calculation"
               Columns(1).Name =   "Calculation"
               Columns(1).DataField=   "Column 1"
               Columns(1).DataType=   8
               Columns(1).FieldLen=   256
               Columns(2).Width=   26458
               Columns(2).Visible=   0   'False
               Columns(2).Caption=   "Type"
               Columns(2).Name =   "Type"
               Columns(2).DataField=   "Column 2"
               Columns(2).DataType=   8
               Columns(2).FieldLen=   256
               Columns(3).Width=   1905
               Columns(3).Caption=   "Type"
               Columns(3).Name =   "TypeDescription"
               Columns(3).DataField=   "Column 3"
               Columns(3).DataType=   8
               Columns(3).FieldLen=   256
               Columns(4).Width=   3519
               Columns(4).Caption=   "Message"
               Columns(4).Name =   "Message"
               Columns(4).DataField=   "Column 4"
               Columns(4).DataType=   8
               Columns(4).FieldLen=   256
               UseDefaults     =   0   'False
               TabNavigation   =   1
               _ExtentX        =   8996
               _ExtentY        =   2011
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
            Begin VB.Label lblFileExtensions 
               AutoSize        =   -1  'True
               Caption         =   "File Extensions :"
               Height          =   195
               Left            =   3000
               TabIndex        =   183
               Top             =   2160
               Width           =   1170
            End
            Begin VB.Label lblFileExtensionNote 
               Caption         =   "(Leave empty to include all file extensions)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   3000
               TabIndex        =   186
               Top             =   1920
               Width           =   1410
            End
         End
         Begin VB.Frame fraControlValues 
            Caption         =   "Control Values : "
            Height          =   855
            Left            =   0
            TabIndex        =   143
            Top             =   3240
            Width           =   1500
            Begin VB.TextBox txtControlValues 
               Height          =   375
               Left            =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   144
               Top             =   300
               Width           =   500
            End
         End
         Begin VB.Frame fraLookup 
            Caption         =   "Lookup :"
            Height          =   3160
            Left            =   2880
            TabIndex        =   145
            Top             =   900
            Width           =   5265
            Begin VB.TextBox txtLookupOrder 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   1100
               Width           =   500
            End
            Begin VB.CommandButton cmdLookupOrder 
               Caption         =   "..."
               Height          =   315
               Left            =   2295
               TabIndex        =   152
               Top             =   1100
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.ComboBox cboLookupFilterColumn 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   155
               Top             =   1860
               Width           =   2000
            End
            Begin VB.CheckBox chkLookupFilter 
               Caption         =   "&Filter Lookup Values"
               Height          =   255
               Left            =   200
               TabIndex        =   153
               Top             =   1560
               Width           =   3555
            End
            Begin VB.ComboBox cboLookupFilterValue 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   159
               Top             =   2670
               Width           =   2000
            End
            Begin VB.ComboBox cboLookupFilterOperator 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   157
               Top             =   2265
               Width           =   2000
            End
            Begin VB.ComboBox cboLookupTable 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   147
               Top             =   300
               Width           =   500
            End
            Begin VB.ComboBox cboLookupColumn 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   149
               Top             =   700
               Width           =   500
            End
            Begin VB.Label lblLookupOrder 
               Caption         =   "Order :"
               Height          =   195
               Left            =   200
               TabIndex        =   150
               Top             =   1160
               Width           =   795
            End
            Begin VB.Label lblLookupFilterColumn 
               Caption         =   "Filter Column :"
               Height          =   285
               Left            =   200
               TabIndex        =   154
               Top             =   1920
               Width           =   1395
            End
            Begin VB.Label lblLookupFilterValue 
               Caption         =   "Filter Value :"
               Height          =   270
               Left            =   200
               TabIndex        =   158
               Top             =   2715
               Width           =   1260
            End
            Begin VB.Label lblLookupFilterOperator 
               Caption         =   "Filter Operator :"
               Height          =   270
               Left            =   200
               TabIndex        =   156
               Top             =   2325
               Width           =   1425
            End
            Begin VB.Label lblLookupTable 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Table :"
               Height          =   195
               Left            =   200
               TabIndex        =   146
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblLookupColumn 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Column :"
               Height          =   195
               Left            =   200
               TabIndex        =   148
               Top             =   760
               Width           =   630
            End
         End
         Begin VB.Frame fraSize 
            Caption         =   "Size :"
            Height          =   855
            Left            =   2900
            TabIndex        =   138
            Top             =   0
            Width           =   5000
            Begin COASpinner.COA_Spinner spnSize 
               Height          =   315
               Left            =   840
               TabIndex        =   140
               Top             =   300
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   999
               MinimumValue    =   1
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnDecimals 
               Height          =   315
               Left            =   3480
               TabIndex        =   142
               Top             =   300
               Width           =   1410
               _ExtentX        =   2487
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
               MaximumValue    =   999
               Text            =   "0"
            End
            Begin VB.Label lblSize 
               AutoSize        =   -1  'True
               Caption         =   "Size :"
               Height          =   195
               Left            =   200
               TabIndex        =   139
               Top             =   360
               Width           =   390
            End
            Begin VB.Label lblDecimals 
               Caption         =   "Decimals :"
               Height          =   195
               Left            =   2505
               TabIndex        =   141
               Top             =   360
               Width           =   990
            End
         End
         Begin VB.Frame fraDefaultValue 
            Caption         =   "Default Value :"
            Height          =   2100
            Left            =   120
            TabIndex        =   160
            Top             =   6960
            Width           =   7815
            Begin VB.TextBox txtDefaultValue 
               Height          =   315
               Left            =   1485
               TabIndex        =   165
               Top             =   300
               Width           =   5730
            End
            Begin VB.ComboBox cboDefaultValue 
               Height          =   315
               Left            =   2835
               Style           =   2  'Dropdown List
               TabIndex        =   171
               Top             =   1500
               Width           =   1035
            End
            Begin VB.CommandButton cmdDefaultValueExpression 
               Caption         =   "..."
               Height          =   315
               Left            =   6900
               TabIndex        =   175
               Top             =   700
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtDefaultValueExpression 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1470
               TabIndex        =   174
               Top             =   700
               Width           =   5430
            End
            Begin VB.OptionButton optDefaultValueType 
               Caption         =   "&Value"
               Height          =   195
               Index           =   0
               Left            =   200
               TabIndex        =   163
               Top             =   360
               Value           =   -1  'True
               Width           =   930
            End
            Begin VB.OptionButton optDefaultValueType 
               Caption         =   "Ca&lculation"
               Height          =   195
               Index           =   3
               Left            =   200
               TabIndex        =   164
               Top             =   760
               Width           =   1320
            End
            Begin VB.Frame fraLogicDefaults 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   195
               Left            =   200
               TabIndex        =   166
               Top             =   1160
               Width           =   2000
               Begin VB.OptionButton optDefaultValue 
                  Caption         =   "&False"
                  Height          =   195
                  Index           =   1
                  Left            =   1130
                  TabIndex        =   168
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton optDefaultValue 
                  Caption         =   "&True"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   167
                  Top             =   0
                  Width           =   810
               End
            End
            Begin COAWorkingPattern.COA_WorkingPattern wpDefaultValue 
               Height          =   765
               Left            =   4020
               TabIndex        =   172
               Top             =   1100
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   1349
            End
            Begin COASpinner.COA_Spinner spnDefaultValue 
               Height          =   315
               Left            =   1695
               TabIndex        =   170
               Top             =   1500
               Width           =   1500
               _ExtentX        =   2646
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
            Begin TDBNumberCtrl.TDBNumber numDefaultValue 
               Height          =   315
               Left            =   5900
               TabIndex        =   173
               Top             =   1500
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
               MouseIcon       =   "frmWorkflowTimeout.frx":006C
               MousePointer    =   0
            End
            Begin GTMaskDate.GTMaskDate dtDefaultValue 
               Height          =   315
               Left            =   200
               TabIndex        =   169
               Top             =   1500
               Width           =   1500
               _Version        =   65537
               _ExtentX        =   2646
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
            Begin VB.Label lblDefaultValueCalculation 
               Caption         =   "Calculation :"
               Height          =   195
               Left            =   195
               TabIndex        =   162
               Top             =   600
               Width           =   1065
            End
            Begin VB.Label lblDefaultValueValue 
               Caption         =   "Value :"
               Height          =   195
               Left            =   195
               TabIndex        =   161
               Top             =   195
               Width           =   495
            End
         End
      End
      Begin VB.PictureBox picTabContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   9690
         Index           =   0
         Left            =   150
         ScaleHeight     =   9690
         ScaleWidth      =   8400
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   400
         Width           =   8400
         Begin VB.Frame fraBehaviour 
            Caption         =   "Behaviour : "
            Height          =   2670
            Left            =   0
            TabIndex        =   41
            Top             =   6960
            Width           =   8000
            Begin VB.CheckBox chkRequireAuthentication 
               Caption         =   "Form requires authenticating before proceeding"
               Height          =   300
               Left            =   180
               TabIndex        =   188
               Top             =   2235
               Width           =   4650
            End
            Begin VB.ComboBox cboFollowOnFormsMessageType 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0088
               Left            =   2700
               List            =   "frmWorkflowTimeout.frx":008A
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   1840
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.ComboBox cboSavedForLaterMessageType 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":008C
               Left            =   2700
               List            =   "frmWorkflowTimeout.frx":008E
               Style           =   2  'Dropdown List
               TabIndex        =   54
               Top             =   1450
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.ComboBox cboCompletionMessageType 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0090
               Left            =   2700
               List            =   "frmWorkflowTimeout.frx":0092
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   1060
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.CommandButton cmdFollowOnFormsMessage 
               Caption         =   "..."
               Height          =   315
               Left            =   6105
               TabIndex        =   58
               Top             =   1840
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.CommandButton cmdSavedForLaterMessage 
               Caption         =   "..."
               Height          =   315
               Left            =   6105
               TabIndex        =   55
               Top             =   1450
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.CommandButton cmdCompletionMessage 
               Caption         =   "..."
               Height          =   315
               Left            =   6105
               TabIndex        =   52
               Top             =   1060
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.CheckBox chkExcludeWeekends 
               Caption         =   "E&xclude Weekends"
               Height          =   255
               Left            =   4560
               TabIndex        =   45
               Top             =   330
               Width           =   2040
            End
            Begin VB.OptionButton optButtonAction 
               Caption         =   "Su&bmit (no validation)"
               Height          =   195
               Index           =   2
               Left            =   4185
               TabIndex        =   48
               Top             =   700
               Width           =   2250
            End
            Begin VB.OptionButton optButtonAction 
               Caption         =   "&Submit (with validation)"
               Height          =   195
               Index           =   0
               Left            =   1800
               TabIndex        =   47
               Top             =   700
               Value           =   -1  'True
               Width           =   2400
            End
            Begin VB.OptionButton optButtonAction 
               Caption         =   "Save for &later"
               Height          =   195
               Index           =   1
               Left            =   6400
               TabIndex        =   49
               Top             =   700
               Width           =   1530
            End
            Begin VB.ComboBox cboTimeoutPeriod 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0094
               Left            =   2715
               List            =   "frmWorkflowTimeout.frx":00A4
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   300
               Width           =   1515
            End
            Begin COASpinner.COA_Spinner spnTimeoutFrequency 
               Height          =   315
               Left            =   1800
               TabIndex        =   43
               Top             =   300
               Width           =   705
               _ExtentX        =   1244
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
               MaximumValue    =   999
               Text            =   "1"
            End
            Begin VB.Label lblCompletionMessage 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Completion Message :"
               Height          =   195
               Left            =   195
               TabIndex        =   50
               Top             =   1125
               Width           =   2115
            End
            Begin VB.Label lblFollowOnFormsMessage 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Follow On Forms Message :"
               Height          =   195
               Left            =   195
               TabIndex        =   56
               Top             =   1905
               Width           =   2505
            End
            Begin VB.Label lblSavedForLaterMessage 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Save For Later Message :"
               Height          =   195
               Left            =   195
               TabIndex        =   53
               Top             =   1515
               Width           =   2385
            End
            Begin VB.Label lblButtonAction 
               Caption         =   "Action :"
               Height          =   195
               Left            =   195
               TabIndex        =   46
               Top             =   705
               Width           =   825
            End
            Begin VB.Label lblTimeoutPeriod 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Timeout Period :"
               Height          =   195
               Left            =   200
               TabIndex        =   42
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame fraDisplay 
            Caption         =   "Display :"
            Height          =   3060
            Left            =   0
            TabIndex        =   17
            Top             =   4335
            Width           =   7400
            Begin VB.ComboBox cboWidthBehaviour 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":00CC
               Left            =   1800
               List            =   "frmWorkflowTimeout.frx":00D6
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   2670
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.ComboBox cboHeightBehaviour 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":00E7
               Left            =   1800
               List            =   "frmWorkflowTimeout.frx":00F1
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   2280
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.ComboBox cboHOffsetBehaviour 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0102
               Left            =   5295
               List            =   "frmWorkflowTimeout.frx":010C
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   1110
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.ComboBox cboVOffsetBehaviour 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":011D
               Left            =   5295
               List            =   "frmWorkflowTimeout.frx":0127
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   720
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.OptionButton optOrientation 
               Caption         =   "&Vertical"
               Height          =   195
               Index           =   1
               Left            =   3165
               TabIndex        =   20
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton optOrientation 
               Caption         =   "Hori&zontal"
               Height          =   195
               Index           =   0
               Left            =   1800
               TabIndex        =   19
               Top             =   360
               Value           =   -1  'True
               Width           =   1410
            End
            Begin COASpinner.COA_Spinner spnTop 
               Height          =   315
               Left            =   1800
               TabIndex        =   22
               Top             =   1500
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   2000
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnHeight 
               Height          =   315
               Left            =   1800
               TabIndex        =   36
               Top             =   1890
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   2000
               MinimumValue    =   1
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnLeft 
               Height          =   315
               Left            =   5295
               TabIndex        =   24
               Top             =   1500
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   2000
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnWidth 
               Height          =   315
               Left            =   5295
               TabIndex        =   40
               Top             =   1890
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   2000
               MinimumValue    =   1
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnVOffset 
               Height          =   315
               Left            =   1800
               TabIndex        =   26
               Top             =   720
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   2000
               Text            =   "0"
            End
            Begin COASpinner.COA_Spinner spnHOffset 
               Height          =   315
               Left            =   1800
               TabIndex        =   30
               Top             =   1110
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
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
               MaximumValue    =   2000
               Text            =   "0"
            End
            Begin VB.Label lblHOffset 
               Caption         =   "Horizontal Offset :"
               Height          =   195
               Left            =   195
               TabIndex        =   29
               Top             =   1140
               Visible         =   0   'False
               Width           =   1680
            End
            Begin VB.Label lblVOffset 
               Caption         =   "Vertical Offset :"
               Height          =   195
               Left            =   195
               TabIndex        =   25
               Top             =   750
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.Label lblWidthValue 
               AutoSize        =   -1  'True
               Caption         =   "Value :"
               Height          =   195
               Left            =   3705
               TabIndex        =   39
               Top             =   2700
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lblHeightValue 
               AutoSize        =   -1  'True
               Caption         =   "Value :"
               Height          =   195
               Left            =   3705
               TabIndex        =   35
               Top             =   2310
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lblVOffsetFrom 
               AutoSize        =   -1  'True
               Caption         =   "From : "
               Height          =   195
               Left            =   3705
               TabIndex        =   27
               Top             =   750
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.Label lblTop 
               Caption         =   "Top :"
               Height          =   195
               Left            =   195
               TabIndex        =   21
               Top             =   1530
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.Label lblHeight 
               Caption         =   "Height :"
               Height          =   195
               Left            =   195
               TabIndex        =   33
               Top             =   1920
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lblLeft 
               Caption         =   "Left :"
               Height          =   195
               Left            =   3705
               TabIndex        =   23
               Top             =   1530
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label lblWidth 
               Caption         =   "Width :"
               Height          =   195
               Left            =   3705
               TabIndex        =   37
               Top             =   1920
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblOrientation 
               Caption         =   "Orientation :"
               Height          =   195
               Left            =   195
               TabIndex        =   18
               Top             =   360
               Width           =   1185
            End
            Begin VB.Label lblHOffsetFrom 
               AutoSize        =   -1  'True
               Caption         =   "From : "
               Height          =   195
               Left            =   3705
               TabIndex        =   31
               Top             =   1140
               Visible         =   0   'False
               Width           =   510
            End
         End
         Begin VB.Frame fraHotspot 
            Caption         =   "Hotspot :"
            Height          =   1095
            Left            =   1380
            TabIndex        =   59
            Top             =   3510
            Width           =   6615
            Begin VB.ComboBox cboHotSpotIdentifier 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0138
               Left            =   3360
               List            =   "frmWorkflowTimeout.frx":0142
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   360
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.Label lblHotSpotIdentifier 
               AutoSize        =   -1  'True
               Caption         =   "Hotspot Identifier : "
               Height          =   195
               Left            =   360
               TabIndex        =   60
               Top             =   480
               Visible         =   0   'False
               Width           =   1680
            End
         End
         Begin VB.Frame fraIdentification 
            Caption         =   "Identification :"
            Height          =   4035
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   7400
            Begin VB.CheckBox chkUseAsTargetIdentifier 
               Caption         =   "Use as workflow target &identifier"
               Height          =   405
               Left            =   150
               TabIndex        =   187
               Top             =   3570
               Width           =   3225
            End
            Begin VB.Frame fraCaption 
               Caption         =   "Caption :"
               Height          =   1200
               Left            =   200
               TabIndex        =   6
               Top             =   1100
               Width           =   6495
               Begin VB.CommandButton cmdCaptionTypeExpression 
                  Caption         =   "..."
                  Height          =   315
                  Left            =   5900
                  TabIndex        =   11
                  Top             =   700
                  UseMaskColor    =   -1  'True
                  Width           =   315
               End
               Begin VB.TextBox txtCaptionTypeExpression 
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1600
                  Locked          =   -1  'True
                  TabIndex        =   10
                  TabStop         =   0   'False
                  Top             =   700
                  Width           =   4300
               End
               Begin VB.TextBox txtCaptionTypeValue 
                  Height          =   315
                  Left            =   1600
                  MaxLength       =   200
                  TabIndex        =   9
                  Top             =   300
                  Width           =   4600
               End
               Begin VB.OptionButton optCaptionType 
                  Caption         =   "Calc&ulation"
                  Height          =   255
                  Index           =   3
                  Left            =   200
                  TabIndex        =   8
                  Top             =   760
                  Width           =   1410
               End
               Begin VB.OptionButton optCaptionType 
                  Caption         =   "V&alue"
                  Height          =   255
                  Index           =   0
                  Left            =   200
                  TabIndex        =   7
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
            Begin VB.CheckBox chkDescriptionHasElementCaption 
               Caption         =   "Description prefixed with &Element Caption"
               Height          =   195
               Left            =   1800
               TabIndex        =   16
               Top             =   3300
               Width           =   4110
            End
            Begin VB.CheckBox chkDescriptionHasWorkflowName 
               Caption         =   "Description prefixed with &Workflow Name "
               Height          =   195
               Left            =   1800
               TabIndex        =   15
               Top             =   2975
               Width           =   4110
            End
            Begin VB.TextBox txtIdentifier 
               Height          =   315
               Left            =   1800
               MaxLength       =   200
               TabIndex        =   3
               Top             =   300
               Width           =   4600
            End
            Begin VB.TextBox txtDescriptionExpression 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   2500
               Width           =   4300
            End
            Begin VB.CommandButton cmdDescriptionExpression 
               Caption         =   "..."
               Height          =   315
               Left            =   6100
               TabIndex        =   14
               Top             =   2500
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtCaption 
               Height          =   315
               Left            =   1800
               MaxLength       =   200
               TabIndex        =   5
               Top             =   700
               Width           =   4600
            End
            Begin VB.Label lblIdentifier 
               Caption         =   "Identifier :"
               Height          =   195
               Left            =   195
               TabIndex        =   2
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label lblDescription 
               Caption         =   "Description :"
               Height          =   195
               Left            =   195
               TabIndex        =   12
               Top             =   2565
               Width           =   1260
            End
            Begin VB.Label lblCaption 
               Caption         =   "Caption :"
               Height          =   195
               Left            =   195
               TabIndex        =   4
               Top             =   765
               Width           =   1125
            End
         End
      End
      Begin VB.PictureBox picTabContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   7300
         Index           =   1
         Left            =   -74850
         ScaleHeight     =   7305
         ScaleWidth      =   8400
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   400
         Width           =   8400
         Begin VB.Frame fraForeground 
            Caption         =   "Foreground :"
            Height          =   2400
            Left            =   0
            TabIndex        =   82
            Top             =   3000
            Width           =   2900
            Begin VB.CommandButton cmdForegroundHighlightColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   97
               Top             =   1900
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtForegroundHighlightColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   96
               Top             =   1900
               Width           =   500
            End
            Begin VB.CommandButton cmdForegroundOddColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   94
               Top             =   1500
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtForegroundOddColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   93
               Top             =   1500
               Width           =   500
            End
            Begin VB.CommandButton cmdForegroundEvenColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   91
               Top             =   1100
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtForegroundEvenColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   90
               Top             =   1100
               Width           =   500
            End
            Begin VB.CommandButton cmdForegroundColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   88
               Top             =   700
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtForegroundColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   87
               Top             =   700
               Width           =   500
            End
            Begin VB.CommandButton cmdForegroundFont 
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   85
               Top             =   300
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtForegroundFont 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   84
               Top             =   300
               Width           =   500
            End
            Begin VB.Label lblForegroundHighlightColour 
               Caption         =   "Highlighted Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   95
               Top             =   1965
               Width           =   1725
            End
            Begin VB.Label lblForegroundOddColour 
               Caption         =   "Odd Row Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   92
               Top             =   1560
               Width           =   1590
            End
            Begin VB.Label lblForegroundEvenColour 
               Caption         =   "Even Row Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   89
               Top             =   1155
               Width           =   1635
            End
            Begin VB.Label lblForegroundColour 
               Caption         =   "Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   86
               Top             =   765
               Width           =   705
            End
            Begin VB.Label lblForegroundFont 
               Caption         =   "Font :"
               Height          =   195
               Left            =   195
               TabIndex        =   83
               Top             =   360
               Width           =   570
            End
         End
         Begin VB.Frame fraOptions 
            Caption         =   "Options :"
            Height          =   1225
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   7400
            Begin VB.CheckBox chkPasswordType 
               Caption         =   "Hide &Text"
               Height          =   195
               Left            =   200
               TabIndex        =   71
               Top             =   760
               Width           =   1230
            End
            Begin VB.CheckBox chkBorder 
               Caption         =   "&Border"
               Height          =   195
               Left            =   200
               TabIndex        =   68
               Top             =   360
               Width           =   1725
            End
            Begin VB.ComboBox cboAlignment 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":0153
               Left            =   5300
               List            =   "frmWorkflowTimeout.frx":0163
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   300
               Width           =   1725
            End
            Begin VB.Label lblAlignment 
               Caption         =   "Alignment :"
               Height          =   195
               Left            =   3705
               TabIndex        =   69
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame fraHeader 
            Caption         =   "Header :"
            Height          =   1625
            Left            =   0
            TabIndex        =   72
            Top             =   1300
            Width           =   7400
            Begin VB.CommandButton cmdHeaderFont 
               Caption         =   "..."
               Height          =   315
               Left            =   6100
               TabIndex        =   78
               Top             =   700
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtHeaderFont 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   315
               Left            =   2025
               TabIndex        =   77
               Top             =   700
               Width           =   4080
            End
            Begin VB.CheckBox chkColumnHeaders 
               Caption         =   "Column &Headers"
               Height          =   195
               Left            =   200
               TabIndex        =   73
               Top             =   360
               Width           =   1860
            End
            Begin VB.TextBox txtHeaderBackgroundColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   2025
               TabIndex        =   80
               Top             =   1100
               Width           =   4080
            End
            Begin VB.CommandButton cmdHeaderBackgroundColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   6100
               TabIndex        =   81
               Top             =   1100
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin COASpinner.COA_Spinner spnHeaderLines 
               Height          =   315
               Left            =   5300
               TabIndex        =   75
               Top             =   300
               Width           =   1065
               _ExtentX        =   1879
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
               Text            =   "0"
            End
            Begin VB.Label lblHeaderLines 
               Caption         =   "Header Lines :"
               Height          =   195
               Left            =   3705
               TabIndex        =   74
               Top             =   360
               Width           =   1395
            End
            Begin VB.Label lblHeaderFont 
               Caption         =   "Font :"
               Height          =   195
               Left            =   195
               TabIndex        =   76
               Top             =   765
               Width           =   660
            End
            Begin VB.Label lblHeaderBackgroundColour 
               Caption         =   "Background Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   79
               Top             =   1155
               Width           =   1815
            End
         End
         Begin VB.Frame fraBackground 
            Caption         =   "Background :"
            Height          =   4000
            Left            =   3050
            TabIndex        =   98
            Top             =   3000
            Width           =   4700
            Begin VB.CommandButton cmdPictureClear 
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
               Left            =   2745
               MaskColor       =   &H000000FF&
               TabIndex        =   116
               ToolTipText     =   "Clear Path"
               Top             =   2300
               UseMaskColor    =   -1  'True
               Width           =   330
            End
            Begin VB.TextBox txtPicture 
               BackColor       =   &H8000000F&
               ForeColor       =   &H80000011&
               Height          =   315
               Left            =   1935
               Locked          =   -1  'True
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   2300
               Width           =   500
            End
            Begin VB.CommandButton cmdPictureSelect 
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   115
               ToolTipText     =   "Select Path"
               Top             =   2300
               Width           =   330
            End
            Begin VB.ComboBox cboPictureLocation 
               Height          =   315
               Left            =   1935
               Style           =   2  'Dropdown List
               TabIndex        =   119
               Top             =   2700
               Width           =   500
            End
            Begin VB.PictureBox picPictureHolder 
               Height          =   1470
               Left            =   3090
               ScaleHeight     =   1410
               ScaleWidth      =   1410
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   2300
               Width           =   1470
               Begin VB.Image picPicture 
                  Height          =   855
                  Left            =   255
                  Stretch         =   -1  'True
                  Top             =   270
                  Width           =   930
               End
            End
            Begin VB.ComboBox cboBackgroundStyle 
               Height          =   315
               ItemData        =   "frmWorkflowTimeout.frx":018B
               Left            =   1935
               List            =   "frmWorkflowTimeout.frx":019B
               Style           =   2  'Dropdown List
               TabIndex        =   100
               Top             =   300
               Width           =   1500
            End
            Begin VB.TextBox txtBackgroundColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   102
               Top             =   700
               Width           =   500
            End
            Begin VB.CommandButton cmdBackgroundColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   103
               Top             =   700
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtBackgroundEvenColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   105
               Top             =   1100
               Width           =   500
            End
            Begin VB.CommandButton cmdBackgroundEvenColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   106
               Top             =   1100
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtBackgroundOddColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   108
               Top             =   1500
               Width           =   500
            End
            Begin VB.CommandButton cmdBackgroundOddColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   109
               Top             =   1500
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.TextBox txtBackgroundHighlightColour 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1935
               TabIndex        =   111
               Top             =   1900
               Width           =   500
            End
            Begin VB.CommandButton cmdBackgroundHighlightColour 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   315
               Left            =   2430
               TabIndex        =   112
               Top             =   1900
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.Label lblPicture 
               Caption         =   "Picture : "
               Height          =   195
               Left            =   195
               TabIndex        =   113
               Top             =   2355
               Width           =   825
            End
            Begin VB.Label lblPictureLocation 
               Caption         =   "Location : "
               Height          =   195
               Left            =   195
               TabIndex        =   118
               Top             =   2760
               Width           =   930
            End
            Begin VB.Label lblBackgroundStyle 
               Caption         =   "Style :"
               Height          =   195
               Left            =   195
               TabIndex        =   99
               Top             =   360
               Width           =   990
            End
            Begin VB.Label lblBackgroundColour 
               Caption         =   "Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   101
               Top             =   765
               Width           =   750
            End
            Begin VB.Label lblBackgroundEvenColour 
               Caption         =   "Even Row Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   104
               Top             =   1155
               Width           =   1635
            End
            Begin VB.Label lblBackgroundOddColour 
               Caption         =   "Odd Row Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   107
               Top             =   1560
               Width           =   1635
            End
            Begin VB.Label lblBackgroundHighlightColour 
               Caption         =   "Highlighted Colour :"
               Height          =   195
               Left            =   195
               TabIndex        =   110
               Top             =   1965
               Width           =   1725
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   240
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
   End
   Begin VB.Label lblSizeTester 
      Caption         =   "<Size Tester>"
      Height          =   255
      Left            =   840
      TabIndex        =   185
      Top             =   10320
      Visible         =   0   'False
      Width           =   1305
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWorkflowTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form handling variables
Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfLoading As Boolean
Private msngCurrentFrameTop As Single
Private msngMaxFrameBottom As Single
Private msngFrameWidth As Single
Private mfReadOnly As Boolean

Private mlngPersonnelTableID As Long

Private maWFPrecedingElements() As VB.Control
Private maWFPrecedingAndCurrentElements() As VB.Control
Private maWFAllElements() As VB.Control

' Form formatting variables
Private Const YGAP_TAB_FRAME = 400
Private Const YGAP_FRAME_CONTROL = 350
Private Const YGAP_CONTROL_LABEL = 60
Private Const YGAP_CONTROL_CONTROL = 400
Private Const YGAP_CONTROL_FRAME = 125
Private Const YGAP_FRAME_FRAME = 100
Private Const YGAP_FRAME_TAB = 200
Private Const YGAP_TAB_BUTTONS = 150
Private Const YGAP_BUTTONS_FORM = 600
Private Const Y_STANDARDCONTROLHEIGHT = 315
Private Const Y_GRIDCONTROLHEIGHT = 2000

Private Const XGAP_TAB_FRAME = 150
Private Const X_COLUMN1 = 200
Private Const X_COLUMN2 = 2100
Private Const X_COLUMN2POINT5 = 2800
Private Const X_COLUMN3 = 4000
Private Const X_COLUMN4 = 5600
Private Const XGAP_CONTROL_CONTROL = 300

' Page number constants.
Private Const miPAGE_GENERAL = 0
Private Const miPAGE_APPEARANCE = 1
Private Const miPAGE_DATA = 2

' PROPERTY VARIABLES
Private mctlSelectedControl As Control
Private mfrmCallingForm As frmWorkflowWFDesigner
Private miItemType As WorkflowWebFormItemTypes

' GENERAL TAB
  ' IDENTIFICATION FRAME
    ' Identifier - held in the control
    ' Caption - held in the control
    ' Description expression
    Private mlngDescriptionExprID As Long
    Private mlngCaptionExprID As Long
  ' DISPLAY FRAME
    ' Orientation - held in the control
    ' Top - held in the control
    ' Left - held in the control
    ' Height - held in the control
    ' Width - held in the control
    Private miVOffsetBehaviour As Integer
    Private miHOffsetBehaviour As Integer
    Private miHeightBehaviour As Integer
    Private miWidthBehaviour As Integer
  ' BEHAVIOUR FRAME
    ' TimeoutFrequency - held in the control
    ' TimeoutPeriod - held in the control
    ' TimeoutExcludeWeekend
    ' Completion Message type - held in the control
    ' Saved For Later Message type - held in the control
    ' Follow On Forms Message type - held in the control
    Private msCompletionMessage As String
    Private msSavedForLaterMessage As String
    Private msFollowOnFormsMessage As String
    
' APPEARANCE TAB
  ' OPTIONS FRAME
    ' Border - held in the control
    ' Alignment - held in the control
  ' HEADER FRAME
    ' HeaderLines - held in the control
    ' ColumnHeaders - held in the control
    Private mObjHeadFont As StdFont
    Private mColHeaderBackColor As OLE_COLOR
  ' FOREGROUND FRAME
    Private mObjFont As StdFont
    Private mColForeColor As OLE_COLOR
    Private mColForeColorEven As OLE_COLOR
    Private mColForeColorOdd As OLE_COLOR
    Private mColForeColorHighlight As OLE_COLOR
  ' BACKGROUND FRAME
    ' BackgroundStyle - held in the control
    Private mColBackColor As OLE_COLOR
    Private mColBackColorEven As OLE_COLOR
    Private mColBackColorOdd As OLE_COLOR
    Private mColBackColorHighlight As OLE_COLOR
    Private mlngPictureID As Long
    ' PictureLocation - held in the control
' DATA TAB
'   RECORDIDENTIFICATION FRAME
    ' Table - held in the control
    ' Record - held in the control
    ' Element - held in the control
    ' RecordSelector - held in the control
    ' RecordTable - held in the control
    Private mlngRecordIdentificationOrderID As Long
    Private mlngRecordIdentificationFilterID As Long
'   SIZE FRAME
    ' Size - held in the control
    ' Decimals - held in the control
'   LOOKUP FRAME
    ' LookupTable - held in the control
    ' LookupColumn - held in the control
    Private mlngLookupOrderID As Long
    ' LookupFilterColumn - held in the control
    ' LookupFilterOperator - held in the control
    ' LookupFilterValue - held in the control
'   CONTROLVALUES FRAME
    ' ControlValues - held in the control
'   DEFAULTVALUES FRAME
    Private mlngDefaultValueExprID As Long
'   VALIDATION FRAME
    ' FileExtensions - held in the control

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
Private Sub RefreshValidationControls()
  If WebFormItemHasProperty(miItemType, WFITEMPROP_VALIDATION) Then
    With grdValidation
      If .Rows > 0 And .SelBookmarks.Count = 0 Then
        .SelBookmarks.Add .AddItemBookmark(0)
      ElseIf .Rows = 0 Then
        .SelBookmarks.RemoveAll
      End If
  
      cmdValidationAdd.Enabled = (Not mfReadOnly)
      cmdValidationEdit.Enabled = (.SelBookmarks.Count > 0)
      cmdValidationDelete.Enabled = (.SelBookmarks.Count > 0 And (Not mfReadOnly))
      cmdValidationDeleteAll.Enabled = (.Rows > 0 And (Not mfReadOnly))
    End With
    
    ResizeValidationColumns
  End If

End Sub

Private Sub ResizeValidationColumns()
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_VALIDATION) Then
    ResizeGridColumns grdValidation
  End If

End Sub

Private Sub ResizeGridColumns(pctlGrid As SSDBGrid)
  ' Size the visible columns in the given grid to fit the text.
  ' If the columns are then not as wide as the grid, stretch out the last visible column.

  Dim iLastVisibleColumn As Integer
  Dim iColumn As Integer
  Dim iRow As Integer
  Dim lngTextWidth As Long
  Dim varBookMark As Variant
  Dim varOriginalPos As Variant
  Dim fVerticalScrollRequired As Boolean
  Dim fHorizontalScrollRequired As Boolean
  
  Const SCROLLWIDTH = 255
  
  iLastVisibleColumn = -1
  lngTextWidth = 0
  
  With pctlGrid
    varOriginalPos = .Bookmark

      .Refresh
    .Redraw = False
    .MoveFirst
    
    For iColumn = 0 To .Columns.Count - 1 Step 1
      lngTextWidth = TextWidth(.Columns(iColumn).Caption)

      If .Columns(iColumn).Visible Then
        iLastVisibleColumn = iColumn
        
        For iRow = 0 To .Rows - 1 Step 1
          varBookMark = .AddItemBookmark(iRow)

          If TextWidth(Trim(.Columns(iColumn).CellText(varBookMark))) > lngTextWidth Then
            lngTextWidth = TextWidth(Trim(.Columns(iColumn).CellText(varBookMark)))
          End If
        Next iRow

        .Columns(iColumn).Width = lngTextWidth + 195
      End If
      lngTextWidth = 0
    Next iColumn

    If iLastVisibleColumn >= 0 Then
      ' Stretch out the last column if required
      fVerticalScrollRequired = (.Rows > .VisibleRows)
      
      If .Columns(iLastVisibleColumn).Left + .Columns(iLastVisibleColumn).Width _
        < (.Width - IIf(fVerticalScrollRequired, SCROLLWIDTH, 0)) Then
      
        .Columns(iLastVisibleColumn).Width = _
          (.Width - IIf(fVerticalScrollRequired, SCROLLWIDTH, 0)) - .Columns(iLastVisibleColumn).Left - 25
      End If
    End If
    
    .Bookmark = varOriginalPos
    .Redraw = True
  End With

End Sub



Private Sub AutoResizeControl()
  ' Adjust dimensions of labels if the font/caption change.
  If (Not mObjFont Is Nothing) Then
    Select Case miItemType
      Case giWFFORMITEM_LABEL
        If optCaptionType(giWFDATAVALUE_FIXED).value Then
          lblSizeTester.Width = PixelsToTwips(spnWidth.value)
          lblSizeTester.Height = PixelsToTwips(spnHeight.value)
          lblSizeTester.Caption = ""
          Set lblSizeTester.Font = mObjFont
          lblSizeTester.WordWrap = True
          lblSizeTester.BorderStyle = IIf(chkBorder.value = vbChecked, vbFixedSingle, vbBSNone)
        
          lblSizeTester.AutoSize = True
          lblSizeTester.Caption = txtCaptionTypeValue.Text
          lblSizeTester.AutoSize = False
        
          spnWidth.value = TwipsToPixels(lblSizeTester.Width)
          spnHeight.value = TwipsToPixels(lblSizeTester.Height)
        End If
        
      Case giWFFORMITEM_DBVALUE, _
        giWFFORMITEM_WFVALUE, _
        giWFFORMITEM_DBFILE, _
        giWFFORMITEM_WFFILE
        
        lblSizeTester.Caption = ""
        Set lblSizeTester.Font = mObjFont
        lblSizeTester.WordWrap = False
        lblSizeTester.BorderStyle = IIf(chkBorder.value = vbChecked, vbFixedSingle, vbBSNone)

        lblSizeTester.AutoSize = True
        lblSizeTester.Caption = "A"
        lblSizeTester.AutoSize = False

        If spnHeight.value < TwipsToPixels(lblSizeTester.Height) Then
          spnHeight.value = TwipsToPixels(lblSizeTester.Height)
        End If
    End Select
  End If

End Sub

Private Sub cboDefaultValue_refresh(ByVal psCurrentValue As String)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  Dim lngColumnID As Long
  Dim iDataType As Integer
  Dim sLookupTableName As String
  Dim sLookupColumnName As String
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
  Dim iRecCount As Integer
  Dim rsLookupValues As New ADODB.Recordset
  Dim vValue As Variant
  Dim objMisc As Misc
  Dim asControlValues() As String
  Dim sList As String
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboDefaultValue
    .Clear
    
    ' Populate the combo
    If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LIST) Then
      sList = MergeControlValues(txtControlValues.Text)
      asControlValues() = Split(sList, vbTab)

      If (miItemType <> giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Then
        .AddItem "<None>"
      End If
      
      For iLoop = 0 To UBound(asControlValues)
        If Len(Trim(asControlValues(iLoop))) > 0 Then
          .AddItem asControlValues(iLoop)
        End If
      Next iLoop
    
    ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOOKUP) Then
      .AddItem "<None>"

      ' Populate the default values combo with the values in the lookup table.
      ' Get the selected table.
      If cboLookupTable.ListCount > 0 Then
        lngTableID = cboLookupTable.ItemData(cboLookupTable.ListIndex)
      End If
    
      ' Get the selected column.
      If cboLookupColumn.ListCount > 0 Then
        lngColumnID = cboLookupColumn.ItemData(cboLookupColumn.ListIndex)
      End If
    
      If (lngTableID > 0) And (lngColumnID > 0) Then
        ' Check that the server-side definition of the selected column in the lookup table
        ' exists, and matches with the local version.
        iDataType = GetColumnDataType(lngColumnID)
        sLookupTableName = GetTableName(lngTableID)
        sLookupColumnName = GetColumnName(lngColumnID, True)

        sSQL = "SELECT COUNT(ASRSysColumns.columnID) AS recCount" & _
          " FROM ASRSysColumns" & _
          " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
          " WHERE ASRSysColumns.columnName = '" & sLookupColumnName & "'" & _
          " AND ASRSysTables.tableName = '" & sLookupTableName & "'" & _
          " AND ASRSysColumns.dataType = " & Trim(Str(iDataType))
        rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        iRecCount = rsInfo!reccount
        rsInfo.Close
        Set rsInfo = Nothing

        If (iRecCount > 0) Then
          sSQL = "SELECT DISTINCT TOP 30000 " & sLookupColumnName & " AS [lookupValue]" & _
            " FROM " & sLookupTableName

          If Len(psCurrentValue) > 0 Then
            sSQL = sSQL & _
              " UNION" & _
              " SELECT " & sLookupColumnName & " AS [lookupValue]" & _
              " FROM " & sLookupTableName & _
              " WHERE " & sLookupColumnName

            Select Case iDataType
              Case dtNUMERIC, dtINTEGER
                sSQL = sSQL & _
                  " = " & UI.ConvertNumberForSQL(val(psCurrentValue))
              Case dtTIMESTAMP
                If IsDate(psCurrentValue) Then
                  sSQL = sSQL & _
                    " = '" & UI.ConvertDateLocaleToSQL(psCurrentValue) & "'"
                Else
                  sSQL = sSQL & _
                    " is null"
                End If
              Case Else
                sSQL = sSQL & _
                  " = '" & Replace(psCurrentValue, "'", "''") & "'"
            End Select
          End If

          sSQL = sSQL & _
            " ORDER BY lookupValue"

          Set objMisc = New Misc
          
          rsLookupValues.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
          Do While Not rsLookupValues.EOF
            vValue = rsLookupValues!LookupValue

            If Not IsNull(vValue) Then
              Select Case iDataType
                Case dtNUMERIC, dtINTEGER
                  .AddItem UI.ConvertNumberForDisplay(Trim(Str(vValue)))

                Case dtTIMESTAMP
                  If IsDate(vValue) Then
                    .AddItem Format(vValue, objMisc.DateFormat)
                  End If

                Case Else
                  .AddItem vValue
              End Select
            End If

            rsLookupValues.MoveNext
          Loop

          rsLookupValues.Close
          Set rsLookupValues = Nothing
        
          Set objMisc = Nothing
        End If
      End If
    End If
    
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .List(iLoop) = psCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    If .ListCount > 0 Then
      .ListIndex = iIndex
    End If
  End With
  
End Sub

Private Sub cboPictureLocation_refresh(ByVal piCurrentValue As Integer)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboPictureLocation
    .Clear
    
    ' Populate the combo
    .AddItem "Top Left"
    .ItemData(.NewIndex) = 0
  
    .AddItem "Top Right"
    .ItemData(.NewIndex) = 1
  
    .AddItem "Centre"
    .ItemData(.NewIndex) = 2
  
    .AddItem "Left Tile"
    .ItemData(.NewIndex) = 3
  
    .AddItem "Right Tile"
    .ItemData(.NewIndex) = 4
  
    .AddItem "Top Tile"
    .ItemData(.NewIndex) = 5
  
    .AddItem "Bottom Tile"
    .ItemData(.NewIndex) = 6
  
    .AddItem "Tile"
    .ItemData(.NewIndex) = 7
  
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
  End With
  
End Sub


Private Sub cboRecordIdentificationElement_refresh(ByVal psCurrentElement As String)
  ' Populate the combo and select the current or default value.
  Dim iRecord As WorkflowRecordSelectorTypes
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim aWFPrecedingElements() As VB.Control
  Dim aLngTableIds() As Long
  Dim alngValidTables() As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim fDone As Boolean
  Dim fFound As Boolean
  
  ReDim aWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements aWFPrecedingElements
  
  iIndex = -1
  iDefaultIndex = 0
  iRecord = giWFRECSEL_UNKNOWN
  
  If cboRecordIdentificationRecord.ListCount > 0 Then
    iRecord = cboRecordIdentificationRecord.ItemData(cboRecordIdentificationRecord.ListIndex)
  End If
  
  With cboRecordIdentificationElement
    .Clear
  
    If (iRecord = giWFRECSEL_IDENTIFIEDRECORD) _
      And (UBound(aWFPrecedingElements) > 1) Then
      
      ReDim aLngTableIds(0)

      If miItemType = giWFFORMITEM_INPUTVALUE_GRID Then
        If cboRecordIdentificationTable.ListCount > 0 Then
          lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
        End If
      
        If lngTableID > 0 Then
          sSQL = "SELECT tmpRelations.parentID" & _
            " FROM tmpRelations" & _
            " WHERE tmpRelations.childID = " & CStr(lngTableID)
          Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

          Do While Not (rsTables.BOF Or rsTables.EOF)
            ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
            aLngTableIds(UBound(aLngTableIds)) = rsTables!parentID

            rsTables.MoveNext
          Loop
          rsTables.Close
          Set rsTables = Nothing
        End If
      ElseIf ((miItemType = giWFFORMITEM_DBVALUE) _
        Or (miItemType = giWFFORMITEM_DBFILE)) Then
        
        lngTableID = GetTableIDFromColumnID(mctlSelectedControl.ColumnID)
        
        ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
        aLngTableIds(UBound(aLngTableIds)) = lngTableID
      End If

      ' Add  an item to the combo for each preceding element that can be used to identify the required table.
      For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
        Set wfTemp = aWFPrecedingElements(iLoop)

        If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
          ' Add  an item to the combo for each grid item in the preceding web form.
          asItems = wfTemp.Items

          fDone = False
          For iLoop2 = 1 To UBound(asItems, 2)
            If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
              For iLoop3 = 1 To UBound(aLngTableIds)
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, iLoop2)), alngValidTables

                fFound = False
                For lngLoop = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop) = aLngTableIds(iLoop3) Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop

                If fFound Then
                  fDone = True
                  .AddItem wfTemp.Identifier
                  Exit For
                End If
              Next iLoop3
            End If

            If fDone Then
              Exit For
            End If
          Next iLoop2

        ElseIf aWFPrecedingElements(iLoop).ElementType = elem_StoredData Then
          For iLoop3 = 1 To UBound(aLngTableIds)
            ReDim alngValidTables(0)
            TableAscendants wfTemp.DataTableID, alngValidTables

            'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
            If miItemType = giWFFORMITEM_INPUTVALUE_GRID Then
              If wfTemp.DataAction = DATAACTION_DELETE Then
                ' Cannot do anything with a Deleted record, but can use its ascendants.
                ' Remove the table itself from the array of valid tables.
                alngValidTables(1) = 0
              End If
            End If

            fFound = False
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = aLngTableIds(iLoop3) Then
                fFound = True
                Exit For
              End If
            Next lngLoop

            If fFound Then
              .AddItem wfTemp.Identifier
              Exit For
            End If
          Next iLoop3
        End If

        Set wfTemp = Nothing
      Next iLoop
    End If

    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .List(iLoop) = psCurrentElement Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop

    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If

    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboRecordIdentificationRecordSelector_refresh ""
    End If
  End With
  
End Sub


Private Sub cboRecordIdentificationRecordSelector_refresh(ByVal psCurrentRecSel As String)
  ' Populate the combo and select the current or default value.
  Dim iRecord As WorkflowRecordSelectorTypes
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim aWFPrecedingElements() As VB.Control
  Dim aLngTableIds() As Long
  Dim alngValidTables() As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim fDone As Boolean
  Dim fFound As Boolean
  Dim fNeeded As Boolean
  Dim sElement As String
  
  ReDim aWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements aWFPrecedingElements

  iIndex = -1
  iDefaultIndex = 0
  iRecord = giWFRECSEL_UNKNOWN

  If cboRecordIdentificationRecord.ListCount > 0 Then
    iRecord = cboRecordIdentificationRecord.ItemData(cboRecordIdentificationRecord.ListIndex)
  End If

  With cboRecordIdentificationRecordSelector
    .Clear

    fNeeded = (iRecord = giWFRECSEL_IDENTIFIEDRECORD) _
      And (cboRecordIdentificationElement.ListCount > 0) _
      And (UBound(aWFPrecedingElements) > 1)
      
    If fNeeded Then
      sElement = cboRecordIdentificationElement.List(cboRecordIdentificationElement.ListIndex)
      
      Set wfTemp = GetElementByIdentifier(sElement)
      fNeeded = (Not wfTemp Is Nothing)
      
      If fNeeded Then
        fNeeded = (wfTemp.ElementType = elem_WebForm)
      End If
      
      Set wfTemp = Nothing
    End If

    If fNeeded Then
      ReDim aLngTableIds(0)

      If miItemType = giWFFORMITEM_INPUTVALUE_GRID Then
        If cboRecordIdentificationTable.ListCount > 0 Then
          lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
        End If

        If lngTableID > 0 Then
          sSQL = "SELECT tmpRelations.parentID" & _
            " FROM tmpRelations" & _
            " WHERE tmpRelations.childID = " & CStr(lngTableID)
          Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

          Do While Not (rsTables.BOF Or rsTables.EOF)
            ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
            aLngTableIds(UBound(aLngTableIds)) = rsTables!parentID

            rsTables.MoveNext
          Loop
          rsTables.Close
          Set rsTables = Nothing
        End If
      ElseIf ((miItemType = giWFFORMITEM_DBVALUE) _
        Or (miItemType = giWFFORMITEM_DBFILE)) Then
        
        lngTableID = GetTableIDFromColumnID(mctlSelectedControl.ColumnID)
        
        ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
        aLngTableIds(UBound(aLngTableIds)) = lngTableID
      End If

      ' Add an item to the combo for each recSel in the identified web form.
      fDone = False
      For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
        If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
          Set wfTemp = aWFPrecedingElements(iLoop)

          If UCase(wfTemp.Identifier) = UCase(sElement) Then
            asItems = wfTemp.Items

            For iLoop2 = 1 To UBound(asItems, 2)
              If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) Then
                For iLoop3 = 1 To UBound(aLngTableIds)
                  ReDim alngValidTables(0)
                  TableAscendants CLng(asItems(44, iLoop2)), alngValidTables

                  fFound = False
                  For lngLoop = 1 To UBound(alngValidTables)
                    If alngValidTables(lngLoop) = aLngTableIds(iLoop3) Then
                      fFound = True
                      Exit For
                    End If
                  Next lngLoop

                  If fFound Then
                    .AddItem asItems(9, iLoop2)
                    Exit For
                  End If
                Next iLoop3
              End If
            Next iLoop2

            fDone = True
            Exit For
          End If
        End If

        If fDone Then
          Exit For
        End If
      Next iLoop
    End If

    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .List(iLoop) = psCurrentRecSel Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop

    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If

    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboRecordIdentificationRecordTable_refresh 0
    End If
  End With
  
End Sub



Private Function CheckExpression(plngFilterID As Long, _
  plngTableID As Long, _
  pfCheckTable As Boolean) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If pfCheckTable And (plngTableID <= 0) Then
    fOK = False
  Else
    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", plngFilterID, False

      If .NoMatch Then
        fOK = False
      Else
        If pfCheckTable _
          And !TableID <> plngTableID Then
          
          fOK = False
        End If
      End If
    End With
  End If
  
TidyUpAndExit:
  CheckExpression = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function CheckOrder(plngOrderID As Long, plngTableID As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If plngTableID <= 0 Then
    fOK = False
  Else
    With recOrdEdit
      .Index = "idxID"
      .Seek "=", plngOrderID

      If .NoMatch Then
        fOK = False
      Else
        If !Deleted Then
          fOK = False
        ElseIf !TableID <> plngTableID Then
          fOK = False
        End If
      End If
    End With
  End If
  
TidyUpAndExit:
  CheckOrder = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function GetElementByIdentifier(psIdentifier As String) As VB.Control
  ' Return the element with the given identifier.
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim aWFPrecedingElements() As VB.Control
  
  If Len(Trim(psIdentifier)) = 0 Then
    Exit Function
  End If
    
  ReDim aWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements aWFPrecedingElements
  
  For lngLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfTemp = aWFPrecedingElements(lngLoop)

    If (UCase(Trim(wfTemp.Identifier)) = UCase(Trim(psIdentifier))) Then
      Set GetElementByIdentifier = wfTemp
      Exit For
    End If
    
    Set wfTemp = Nothing
  Next lngLoop

End Function

Private Sub cboRecordIdentificationRecord_refresh(ByVal piCurrentRecord As WorkflowRecordSelectorTypes)
  ' Populate the combo and select the current or default value.
  Dim lngTableID As Long
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim aWFPrecedingElements() As VB.Control
  Dim alngValidTables() As Long
  Dim fTableOK As Boolean
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sTableIDs As String
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim sValidTableIDs As String
  
  iIndex = -1
  iDefaultIndex = 0

  ReDim aWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements aWFPrecedingElements

  With cboRecordIdentificationRecord
    .Clear

    If (miItemType = giWFFORMITEM_DBVALUE) _
      Or (miItemType = giWFFORMITEM_DBFILE) Then
      
      lngTableID = GetTableIDFromColumnID(mctlSelectedControl.ColumnID)
      
      If lngTableID > 0 Then
        If UBound(aWFPrecedingElements) > 1 Then
          fTableOK = False

          For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
            Set wfTemp = aWFPrecedingElements(iLoop)
  
            If wfTemp.ElementType = elem_WebForm Then
              ' Add  an item to the combo for each grid in the preceding web form.
              asItems = wfTemp.Items
              For iLoop2 = 1 To UBound(asItems, 2)
                If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                  ReDim alngValidTables(0)
                  TableAscendants CLng(asItems(44, iLoop2)), alngValidTables

                  For lngLoop = 1 To UBound(alngValidTables)
                    If (lngTableID = alngValidTables(lngLoop)) Then
                      fTableOK = True
                      Exit For
                    End If
                  Next lngLoop

                  If fTableOK Then
                    Exit For
                  End If
                End If
              Next iLoop2
            ElseIf wfTemp.ElementType = elem_StoredData Then
              ReDim alngValidTables(0)
              TableAscendants wfTemp.DataTableID, alngValidTables

              'JPD 20061227 DBValues can now be from DELETE StoredData elements
              For lngLoop = 1 To UBound(alngValidTables)
                If (lngTableID = alngValidTables(lngLoop)) Then
                  fTableOK = True
                  Exit For
                End If
              Next lngLoop
            End If
  
            Set wfTemp = Nothing
  
            If fTableOK Then
              Exit For
            End If
          Next iLoop

          If fTableOK Then
            .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
            .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
          End If
        End If
        
        ReDim alngValidTables(0)
        TableAscendants mfrmCallingForm.BaseTable, alngValidTables

        For lngLoop = 1 To UBound(alngValidTables)
          If (lngTableID = alngValidTables(lngLoop)) Then
            If (mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_MANUAL) Then
              .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
              .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
            ElseIf (mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) Then
              .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
              .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD
            End If
            
            Exit For
          End If
        Next lngLoop
      End If
    Else
      If cboRecordIdentificationTable.ListCount > 0 Then
        lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
      End If
    
      If lngTableID > 0 Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_ALL)
        .ItemData(.NewIndex) = giWFRECSEL_ALL

        ' Only add 'Identified' as an option if the selected table is a child of a preceding WebForm's record selector
        sTableIDs = "0"

        If UBound(aWFPrecedingElements) > 1 Then
          ' Add  an item to the combo for each valid preceding WebForm or StroedData element.
          For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
            Set wfTemp = aWFPrecedingElements(iLoop)

            If wfTemp.ElementType = elem_WebForm Then
              ' Add  an item to the combo for each input item in the preceding web form.
              asItems = wfTemp.Items
              For iLoop2 = 1 To UBound(asItems, 2)
                If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                  ReDim alngValidTables(0)
                  TableAscendants CLng(asItems(44, iLoop2)), alngValidTables

                  For lngLoop = 1 To UBound(alngValidTables)
                    sTableIDs = sTableIDs & "," & CStr(alngValidTables(lngLoop))
                  Next lngLoop
                End If
              Next iLoop2
            ElseIf wfTemp.ElementType = elem_StoredData Then
              ReDim alngValidTables(0)
              TableAscendants wfTemp.DataTableID, alngValidTables

              'JPD 20061227 RecSels still CANNOT be from DELETE StoredData elements
              If wfTemp.DataAction = DATAACTION_DELETE Then
                ' Cannot do anything with a Deleted record, but can use its ascendants.
                ' Remove the table itself from the array of valid tables.
                alngValidTables(1) = 0
              End If

              For lngLoop = 1 To UBound(alngValidTables)
                sTableIDs = sTableIDs & "," & CStr(alngValidTables(lngLoop))
              Next lngLoop
            End If

            Set wfTemp = Nothing
          Next iLoop
        End If

        sSQL = "SELECT COUNT(*) AS [result]" & _
          " FROM tmpRelations" & _
          " WHERE tmpRelations.parentID IN(" & sTableIDs & ")" & _
          " AND tmpRelations.childID = " & CStr(lngTableID)
        Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

        If rsTables!result > 0 Then
          .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
          .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
        End If

        rsTables.Close
        Set rsTables = Nothing

        If mfrmCallingForm.BaseTable > 0 Then
          sValidTableIDs = "0"
          ReDim alngValidTables(0)
          TableAscendants mfrmCallingForm.BaseTable, alngValidTables
  
          For lngLoop = 1 To UBound(alngValidTables)
            sValidTableIDs = sValidTableIDs & "," & CStr(alngValidTables(lngLoop))
          Next lngLoop
  
          sSQL = "SELECT COUNT(*) AS [result]" & _
            " FROM tmpRelations" & _
            " WHERE tmpRelations.parentID IN (" & sValidTableIDs & ")" & _
            " AND tmpRelations.childID = " & CStr(lngTableID)
          Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
          If rsTables!result > 0 Then
            If mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
              .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
              .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
            ElseIf mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
              .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
              .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD
            End If
          End If
          
          rsTables.Close
          Set rsTables = Nothing
        End If
      End If
    End If

    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentRecord Then
        iIndex = iLoop
        Exit For
      End If

      If .ItemData(iLoop) = mfrmCallingForm.BaseTable Then
        iDefaultIndex = iLoop
      End If
    Next iLoop

    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If

    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboRecordIdentificationElement_refresh ""
    End If
  End With
  
End Sub



Private Sub cboRecordIdentificationRecordTable_refresh(ByVal plngCurrentRecordTableID As Long)
  ' Populate the combo and select the current or default value.
  Dim iRecord As WorkflowRecordSelectorTypes
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim sValidTableIDs As String
  Dim lngTemp As Long
  Dim aWFPrecedingElements() As VB.Control
  Dim alngValidTables() As Long
  Dim fDone As Boolean
  Dim wfTemp As VB.Control
  Dim sElement As String
  Dim sRecSel As String
  Dim asItems() As String
  Dim lngTableID As Long
  
  iIndex = -1
  iDefaultIndex = 0
  iRecord = giWFRECSEL_UNKNOWN

  ReDim aWFPrecedingElements(0)
  mfrmCallingForm.PrecedingElements aWFPrecedingElements

  If cboRecordIdentificationRecord.ListCount > 0 Then
    iRecord = cboRecordIdentificationRecord.ItemData(cboRecordIdentificationRecord.ListIndex)
  End If
  
  With cboRecordIdentificationRecordTable
    .Clear

    ' Populate the combo
    If iRecord <> giWFRECSEL_ALL Then
      sValidTableIDs = "0"
      lngTemp = 0

      If miItemType = giWFFORMITEM_INPUTVALUE_GRID Then
        If cboRecordIdentificationTable.ListCount > 0 Then
          lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
        End If
      Else
        lngTableID = GetTableIDFromColumnID(mctlSelectedControl.ColumnID)
      End If
      
      ReDim alngValidTables(0)
      If iRecord = giWFRECSEL_INITIATOR Then
        lngTemp = mlngPersonnelTableID
        TableAscendants lngTemp, alngValidTables
      ElseIf iRecord = giWFRECSEL_TRIGGEREDRECORD Then
        lngTemp = mfrmCallingForm.BaseTable
        TableAscendants lngTemp, alngValidTables
      ElseIf iRecord = giWFRECSEL_IDENTIFIEDRECORD Then
        fDone = False
        
        If cboRecordIdentificationElement.ListCount > 0 Then
          sElement = cboRecordIdentificationElement.List(cboRecordIdentificationElement.ListIndex)
        
          If cboRecordIdentificationRecordSelector.ListCount > 0 Then
            sRecSel = cboRecordIdentificationRecordSelector.List(cboRecordIdentificationRecordSelector.ListIndex)
          End If
        
          For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
            Set wfTemp = aWFPrecedingElements(iLoop)
  
            If UCase(wfTemp.Identifier) = UCase(sElement) Then
              If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
                asItems = wfTemp.Items

                For iLoop2 = 1 To UBound(asItems, 2)
                  If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) Then
                    If (asItems(9, iLoop2) = sRecSel) Then
                      lngTemp = CLng(asItems(44, iLoop2))
                      TableAscendants lngTemp, alngValidTables
                    End If
                  End If
                Next iLoop2

              ElseIf aWFPrecedingElements(iLoop).ElementType = elem_StoredData Then
                lngTemp = aWFPrecedingElements(iLoop).DataTableID
                TableAscendants lngTemp, alngValidTables

                'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
                If aWFPrecedingElements(iLoop).DataAction = DATAACTION_DELETE Then
                  ' Cannot do anything with a Deleted record, but can use its ascendants.
                  ' Remove the table itself from the array of valid tables.
                  alngValidTables(1) = 0
                End If
              End If

              fDone = True
              Exit For
            End If

            If fDone Then
              Exit For
            End If
          Next iLoop
        End If
      End If

      For lngLoop = 1 To UBound(alngValidTables)
        sValidTableIDs = sValidTableIDs & "," & CStr(alngValidTables(lngLoop))
      Next lngLoop

      sSQL = "SELECT tmpRelations.parentID, tmpTables.tableName" & _
        " FROM tmpRelations, tmpTables" & _
        " WHERE tmpRelations.parentID IN (" & sValidTableIDs & ")" & _
        " AND tmpRelations.childID = " & CStr(lngTableID) & _
        " AND tmpRelations.parentID = tmpTables.tableID" & _
        " AND tmpTables.deleted = FALSE" & _
        " ORDER BY tmpTables.tableName"
      Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      While Not rsTables.EOF
        .AddItem rsTables!TableName
        .ItemData(.NewIndex) = rsTables!parentID

        rsTables.MoveNext
      Wend
      rsTables.Close
      Set rsTables = Nothing
    End If

    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = plngCurrentRecordTableID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop

    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If

    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    End If
  End With
  
End Sub



Private Sub cboRecordIdentificationTable_refresh(ByVal pLngCurrentTableID As Long)
  ' Populate the combo and select the current or default value.
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer

  iIndex = -1
  iDefaultIndex = 0

  With cboRecordIdentificationTable
    .Clear

    ' Populate the combo
    sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
      " FROM tmpTables" & _
      " WHERE (tmpTables.deleted = FALSE)" & _
      " ORDER BY tmpTables.tableName"
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    Do While Not rsTables.EOF
      .AddItem rsTables!TableName
      .ItemData(.NewIndex) = rsTables!TableID

      rsTables.MoveNext
    Loop

    rsTables.Close
    Set rsTables = Nothing

    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = pLngCurrentTableID Then
        iIndex = iLoop
        Exit For
      End If
    
      If .ItemData(iLoop) = mfrmCallingForm.BaseTable Then
        iDefaultIndex = iLoop
      End If
    Next iLoop

    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If

    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboRecordIdentificationRecord_refresh giWFRECSEL_UNKNOWN
    End If
  End With
  
End Sub




Private Sub cboLookupTable_refresh(ByVal pLngCurrentTableID As Long)
  ' Populate the combo and select the current or default value.
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboLookupTable
    .Clear
    
    ' Populate the combo
    sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
      " FROM tmpTables" & _
      " WHERE (tmpTables.deleted = FALSE)" & _
      " ORDER BY tmpTables.tableName"
    Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    Do While Not rsTables.EOF
      .AddItem rsTables!TableName
      .ItemData(.NewIndex) = rsTables!TableID
      
      rsTables.MoveNext
    Loop

    rsTables.Close
    Set rsTables = Nothing
  
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = pLngCurrentTableID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .Enabled = (.ListCount > 0) And (Not mfReadOnly)
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      cboLookupColumn_refresh 0
      cboLookupFilterColumn_Refresh 0
    End If
  End With
  
End Sub



Private Sub cboLookupColumn_refresh(ByVal plngCurrentColumnID As Long)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboLookupColumn
    .Clear
    
    ' Get the selected table.
    If cboLookupTable.ListCount > 0 Then
      lngTableID = cboLookupTable.ItemData(cboLookupTable.ListIndex)
    End If
    
    If lngTableID > 0 Then
      ' Populate the combo
      recColEdit.Index = "idxName"
      recColEdit.Seek ">=", lngTableID

      If Not recColEdit.NoMatch Then
        Do While Not recColEdit.EOF
          If recColEdit.Fields("tableID") <> lngTableID Then
            Exit Do
          End If

          ' Add each column name to the lookup columns combo.
          ' NB. We only want to add certain types of column. There's not use in
          ' looking up OLE or logic values.
          If (recColEdit.Fields("columnType") <> giCOLUMNTYPE_SYSTEM) And _
            (recColEdit.Fields("columnType") <> giCOLUMNTYPE_LINK) And _
            (Not recColEdit.Fields("deleted")) And _
            (recColEdit.Fields("dataType") <> dtLONGVARBINARY) And _
            (recColEdit.Fields("dataType") <> dtVARBINARY) And _
            (recColEdit.Fields("dataType") <> dtBIT) Then

            .AddItem recColEdit.Fields("columnName").value
            .ItemData(.NewIndex) = recColEdit.Fields("columnID")
          End If

          recColEdit.MoveNext
        Loop
      
        ' Get the indexes of the required/default values
        For iLoop = 0 To .ListCount - 1
          If .ItemData(iLoop) = plngCurrentColumnID Then
            iIndex = iLoop
            Exit For
          End If
        Next iLoop

        If iIndex < 0 Then
          iIndex = iDefaultIndex
        End If

        .Enabled = (.ListCount > 0) And (Not mfReadOnly)
        If .ListCount > 0 Then
          .ListIndex = iIndex
        End If
      End If
    End If
  End With
  
End Sub

Private Sub cboLookupFilterValue_Refresh(ByVal psCurrentValue As String)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim asItems() As String
  Dim lngLookupFilterColumnType As DataTypes
  Dim fItemOK As Boolean
  Dim sItemDescription As String
  
  iIndex = -1
  iDefaultIndex = 0

  With cboLookupFilterValue
    .Clear

    .AddItem "<None>"
    .ItemData(.NewIndex) = "0"
    .ListIndex = 0

    If (chkLookupFilter.value = vbChecked) And (cboLookupFilterOperator.ListIndex >= 0) Then
      ' Get the filter column data type, etc.
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", cboLookupFilterColumn.ItemData(cboLookupFilterColumn.ListIndex)
    
        If .NoMatch Then
          lngLookupFilterColumnType = sqlUnknown
        Else
          lngLookupFilterColumnType = .Fields("DataType")
        End If
      End With
      
      asItems = mfrmCallingForm.CurrentElementDefinition.Items
      
      For iLoop = 1 To UBound(asItems, 2)

        Select Case asItems(2, iLoop)
          Case giWFFORMITEM_BUTTON
            fItemOK = False

          Case giWFFORMITEM_INPUTVALUE_CHAR
            fItemOK = (lngLookupFilterColumnType = dtVARCHAR) _
              Or (lngLookupFilterColumnType = dtLONGVARCHAR)
          
          Case giWFFORMITEM_INPUTVALUE_NUMERIC
            fItemOK = (lngLookupFilterColumnType = dtINTEGER) _
              Or (lngLookupFilterColumnType = dtNUMERIC)

          Case giWFFORMITEM_INPUTVALUE_LOGIC
            fItemOK = (lngLookupFilterColumnType = dtBIT)

          Case giWFFORMITEM_INPUTVALUE_DATE
            fItemOK = (lngLookupFilterColumnType = dtTIMESTAMP)

          Case giWFFORMITEM_INPUTVALUE_DROPDOWN
            fItemOK = (lngLookupFilterColumnType = dtVARCHAR) _
              Or (lngLookupFilterColumnType = dtLONGVARCHAR)

          Case giWFFORMITEM_INPUTVALUE_LOOKUP
            Select Case GetColumnDataType(CLng(asItems(49, iLoop)))
              Case dtLONGVARCHAR
                fItemOK = (lngLookupFilterColumnType = dtVARCHAR) _
                  Or (lngLookupFilterColumnType = dtLONGVARCHAR)
              Case dtNUMERIC
                fItemOK = (lngLookupFilterColumnType = dtINTEGER) _
                  Or (lngLookupFilterColumnType = dtNUMERIC)
              Case dtINTEGER
                fItemOK = (lngLookupFilterColumnType = dtINTEGER) _
                  Or (lngLookupFilterColumnType = dtNUMERIC)
              Case dtTIMESTAMP
                fItemOK = (lngLookupFilterColumnType = dtTIMESTAMP)
              Case dtVARCHAR
                fItemOK = (lngLookupFilterColumnType = dtVARCHAR) _
                  Or (lngLookupFilterColumnType = dtLONGVARCHAR)
              Case Else
                fItemOK = False
            End Select

          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
            fItemOK = (lngLookupFilterColumnType = dtVARCHAR) _
              Or (lngLookupFilterColumnType = dtLONGVARCHAR)

          Case Else
            fItemOK = False
        
        End Select

        If fItemOK _
          And (mctlSelectedControl.WFIdentifier <> asItems(9, iLoop)) Then
          .AddItem asItems(9, iLoop)
        End If
      Next iLoop

      ' Get the indexes of the required/default values
      For iLoop = 0 To .ListCount - 1
        If .List(iLoop) = psCurrentValue Then
          iIndex = iLoop
          Exit For
        End If
      Next iLoop

      If iIndex < 0 Then
        iIndex = iDefaultIndex
      End If

      .Enabled = (.ListCount > 0) And (Not mfReadOnly)
      If .ListCount > 0 Then
        .ListIndex = iIndex
      End If
    Else
      .Enabled = False
    End If

    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    lblLookupFilterValue.Enabled = .Enabled
  End With
  
End Sub


Private Sub cboLookupFilterColumn_Refresh(ByVal plngCurrentColumnID As Long)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long

  iIndex = -1
  iDefaultIndex = 0

  With cboLookupFilterColumn
    .Clear

    .AddItem "<None>"
    .ItemData(.NewIndex) = "0"
    .ListIndex = 0

    If chkLookupFilter.value = vbChecked Then
      ' Get the selected table.
        If cboLookupTable.ListCount > 0 Then
        lngTableID = cboLookupTable.ItemData(cboLookupTable.ListIndex)
      End If

      If lngTableID > 0 Then
        ' Populate the combo
        recColEdit.Index = "idxName"
        recColEdit.Seek ">=", lngTableID
  
        If Not recColEdit.NoMatch Then
          Do While Not recColEdit.EOF
            If recColEdit.Fields("tableID") <> lngTableID Then
              Exit Do
            End If
  
            ' Add each column name to the lookup filter columns combo.
            ' NB. We only want to add certain types of column. There's not use in
            ' looking up OLE or logic values.
            If (recColEdit.Fields("columnType") <> giCOLUMNTYPE_SYSTEM) And _
              (recColEdit.Fields("columnType") <> giCOLUMNTYPE_LINK) And _
              (Not recColEdit.Fields("deleted")) And _
              (recColEdit.Fields("dataType") <> dtLONGVARBINARY) And _
              (recColEdit.Fields("dataType") <> dtVARBINARY) Then
  
              .AddItem recColEdit.Fields("columnName").value
              .ItemData(.NewIndex) = recColEdit.Fields("columnID")
            End If
  
            recColEdit.MoveNext
          Loop
  
          ' Get the indexes of the required/default values
          For iLoop = 0 To .ListCount - 1
            If .ItemData(iLoop) = plngCurrentColumnID Then
              iIndex = iLoop
              Exit For
            End If
          Next iLoop
  
          If iIndex < 0 Then
            iIndex = iDefaultIndex
          End If
  
          .Enabled = (.ListCount > 0) And (Not mfReadOnly)
          If .ListCount > 0 Then
            .ListIndex = iIndex
          End If
        End If
      End If
    Else
      .Enabled = False
    End If
    
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    lblLookupFilterColumn.Enabled = .Enabled
  End With
  
End Sub


Private Sub cboLookupFilterOperator_Refresh(ByVal piCurrentOperatorID As Integer)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLookupFilterColumnType As DataTypes

  iIndex = -1
  iDefaultIndex = 0

  With cboLookupFilterOperator
    .Clear

    If chkLookupFilter.value = vbChecked Then

      ' Get the filter column data type, etc.
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", cboLookupFilterColumn.ItemData(cboLookupFilterColumn.ListIndex)
    
        If .NoMatch Then
          lngLookupFilterColumnType = sqlUnknown
        Else
          lngLookupFilterColumnType = .Fields("DataType")
        End If
      End With
              
      Select Case lngLookupFilterColumnType
        Case sqlOle  ' Not required as OLEs are not permitted in the Lookup Filter Column selection.
        
        Case sqlBoolean ' Logic columns.
          .AddItem OperatorDescription(giFILTEROP_EQUALS)
          .ItemData(.NewIndex) = giFILTEROP_EQUALS
          .AddItem OperatorDescription(giFILTEROP_NOTEQUALTO)
          .ItemData(.NewIndex) = giFILTEROP_NOTEQUALTO

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

      ' Get the indexes of the required/default values
      For iLoop = 0 To .ListCount - 1
        If .ItemData(iLoop) = piCurrentOperatorID Then
          iIndex = iLoop
          Exit For
        End If
      Next iLoop

      If iIndex < 0 Then
        iIndex = iDefaultIndex
      End If

      .Enabled = (.ListCount > 0) And (Not mfReadOnly)
      If .ListCount > 0 Then
        .ListIndex = iIndex
      End If
    Else
      .Enabled = False
    End If
    
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    lblLookupFilterOperator.Enabled = .Enabled

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



Private Sub cboVOffsetBehaviour_Refresh(ByVal piCurrentValue As VerticalOffset)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboVOffsetBehaviour
    .Clear
    
    ' Populate the combo
    .AddItem "Top"
    .ItemData(.NewIndex) = offsetTop
  
    .AddItem "Bottom"
    .ItemData(.NewIndex) = offsetBottom
    
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
      
      If .ItemData(iLoop) = offsetTop Then
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
        
  End With
  
  
End Sub


Private Sub cboHOffsetBehaviour_Refresh(ByVal piCurrentValue As HorizontalOffset)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboHOffsetBehaviour
    .Clear
    
    ' Populate the combo
    .AddItem "Left"
    .ItemData(.NewIndex) = offsetLeft
  
    .AddItem "Right"
    .ItemData(.NewIndex) = offsetRight
    
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
      
      If .ItemData(iLoop) = offsetLeft Then
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
        
    .ListIndex = iIndex
        
  End With
End Sub

Private Sub cboHeightBehaviour_Refresh(ByVal piCurrentValue As ControlSizeBehaviour)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboHeightBehaviour
    .Clear
    
    ' Populate the combo
    .AddItem "Fixed"
    .ItemData(.NewIndex) = behaveFixed
  
    .AddItem "Full"
    .ItemData(.NewIndex) = behaveFull
    
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
      
      If .ItemData(iLoop) = behaveFixed Then
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
  End With
End Sub

Private Sub cboWidthBehaviour_Refresh(ByVal piCurrentValue As ControlSizeBehaviour)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboWidthBehaviour
    .Clear
    
    ' Populate the combo
    .AddItem "Fixed"
    .ItemData(.NewIndex) = behaveFixed
  
    .AddItem "Full"
    .ItemData(.NewIndex) = behaveFull
    
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
      
      If .ItemData(iLoop) = behaveFixed Then
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
  End With
End Sub

Private Sub cboAlignment_refresh(ByVal piCurrentValue As AlignmentConstants)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboAlignment
    .Clear
    
    ' Populate the combo
    .AddItem "Left Alignment"
    .ItemData(.NewIndex) = vbLeftJustify
  
    .AddItem "Right Alignment"
    .ItemData(.NewIndex) = vbRightJustify
    
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
      
      If .ItemData(iLoop) = vbLeftJustify Then
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
  End With
End Sub



Private Sub cboBackgroundStyle_refresh(ByVal piCurrentValue As ASRBackStyleConstants)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboBackgroundStyle
    .Clear
    
    ' Populate the combo
    .AddItem "Opaque"
    .ItemData(.NewIndex) = BACKSTYLE_OPAQUE

    .AddItem "Transparent"
    .ItemData(.NewIndex) = BACKSTYLE_TRANSPARENT
  
    ' Get the indexes of the required/default values
    For iLoop = 0 To cboBackgroundStyle.ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
  
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If

    .ListIndex = iIndex
  End With

End Sub


Private Sub cboHotSpotIdentifier_refresh(ByVal psCurrentValue As String)
  
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim asItems() As String
  ' Dim lngLookupFilterColumnType As DataTypes
  Dim fItemOK As Boolean
  Dim sItemDescription As String
  
  iIndex = -1
  iDefaultIndex = 0

  With cboHotSpotIdentifier
    .Clear

    .AddItem "<None>"
    .ItemData(.NewIndex) = "0"
    .ListIndex = 0
      
      asItems = mfrmCallingForm.CurrentElementDefinition.Items
      
      For iLoop = 1 To UBound(asItems, 2)
        
        Select Case asItems(2, iLoop)
          Case giWFFORMITEM_FRAME
            fItemOK = False

          Case giWFFORMITEM_IMAGE
            fItemOK = False
          
          Case giWFFORMITEM_PAGETAB
            fItemOK = False

          Case giWFFORMITEM_INPUTVALUE_GRID
            fItemOK = False

          Case giWFFORMITEM_LABEL
            fItemOK = False

          Case giWFFORMITEM_LINE
            fItemOK = False
          
          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
            fItemOK = False
            
          Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
            fItemOK = False
          Case Else
        
            'Item is ok to add to the list if it is on the same tab page as the selected control
            fItemOK = (asItems(78, iLoop) = mfrmCallingForm.GetControlPageNo(mctlSelectedControl))
                
        End Select

        If fItemOK _
          And (mctlSelectedControl.WFIdentifier <> asItems(9, iLoop)) Then
          .AddItem asItems(9, iLoop)
        End If
      Next iLoop

      ' Get the indexes of the required/default values
      For iLoop = 0 To .ListCount - 1
        If .List(iLoop) = psCurrentValue Then
          iIndex = iLoop
          Exit For
        End If
      Next iLoop

      If iIndex < 0 Then
        iIndex = iDefaultIndex
      End If

      .Enabled = (.ListCount > 0) And (Not mfReadOnly)
      If .ListCount > 0 Then
        .ListIndex = iIndex
      End If

    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    lblHotSpotIdentifier.Enabled = .Enabled
  End With
  
  
End Sub


Private Sub cboTimeoutPeriod_refresh(ByVal piCurrentValue As TimeoutPeriod)
  ' Populate the combo and select the current or default value.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  
  iIndex = -1
  iDefaultIndex = 0
  
  With cboTimeoutPeriod
    .Clear
    
    ' Populate the combo
    .AddItem "Minute(s)"
    .ItemData(.NewIndex) = TIMEOUT_MINUTE
  
    .AddItem "Hour(s)"
    .ItemData(.NewIndex) = TIMEOUT_HOUR
  
    .AddItem "Day(s)"
    .ItemData(.NewIndex) = TIMEOUT_DAY
  
    .AddItem "Week(s)"
    .ItemData(.NewIndex) = TIMEOUT_WEEK
  
    .AddItem "Month(s)"
    .ItemData(.NewIndex) = TIMEOUT_MONTH
  
    .AddItem "Year(s)"
    .ItemData(.NewIndex) = TIMEOUT_YEAR
  
    ' Get the indexes of the required/default values
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = piCurrentValue Then
        iIndex = iLoop
        Exit For
      End If
      
      If .ItemData(iLoop) = TIMEOUT_DAY Then
        iDefaultIndex = iLoop
      End If
    Next iLoop
    
    If iIndex < 0 Then
      iIndex = iDefaultIndex
    End If
    
    .ListIndex = iIndex
  End With
  
End Sub


Private Sub RefreshBackgroundControls()
  Dim fEnable As Boolean
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKSTYLE) _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLOR) Then
    ' Disable the BackgroundColour controls if the BackgroundStyle is Transparent.
    If cboBackgroundStyle.ListIndex >= 0 Then
      fEnable = (cboBackgroundStyle.ItemData(cboBackgroundStyle.ListIndex) = BACKSTYLE_OPAQUE) _
        And (Not mfReadOnly)
    End If
    
    EnableControl lblBackgroundColour, fEnable
    EnableControl cmdBackgroundColour, fEnable
    
    ' Cannot use EnableControl for the txtBackgroundColour control
    ' as the BackColor is non-standard (used to show the selected colour)
    'EnableControl txtBackgroundColour, fEnable
    With txtBackgroundColour
      .BackColor = IIf((cboBackgroundStyle.ItemData(cboBackgroundStyle.ListIndex) = BACKSTYLE_OPAQUE), mColBackColor, vbButtonFace)
      'JPD 20071206 Fault 12581
      '.Locked = Not fEnable
      '.TabStop = fEnable
      '.Enabled = fEnable And (Not mfReadOnly)
    End With
  End If

End Sub

Private Sub RefreshBehaviourControls()
  Dim fEnable As Boolean
    
  If WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGETYPE) _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGE) Then
    
    If cboCompletionMessageType.ListIndex >= 0 Then
      fEnable = (cboCompletionMessageType.ItemData(cboCompletionMessageType.ListIndex) = MESSAGE_CUSTOM)
    End If

    EnableControl cmdCompletionMessage, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGETYPE) _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGE) Then
    
    If cboSavedForLaterMessageType.ListIndex >= 0 Then
      fEnable = (cboSavedForLaterMessageType.ItemData(cboSavedForLaterMessageType.ListIndex) = MESSAGE_CUSTOM)
    End If

    EnableControl cmdSavedForLaterMessage, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGETYPE) _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGE) Then
    
    If cboFollowOnFormsMessageType.ListIndex >= 0 Then
      fEnable = (cboFollowOnFormsMessageType.ItemData(cboFollowOnFormsMessageType.ListIndex) = MESSAGE_CUSTOM)
    End If

    EnableControl cmdFollowOnFormsMessage, fEnable
  End If
  
End Sub


Private Sub DefineValidation(blnNew As Boolean)

  Dim frmValidation As frmWorkflowWFValidation
  Dim lngRow As Long
  Dim lngExprID As Long
  Dim iType As WorkflowWebFormValidationTypes
  Dim sTypeDesc As String
  Dim sMessage As String
  Dim sTemp As String
  Dim fRowModified As Boolean
  
  Set frmValidation = New frmWorkflowWFValidation

  lngExprID = 0
  iType = WORKFLOWWFVALIDATIONTYPE_ERROR
  sMessage = ""
  sTypeDesc = ""
  
  If Not blnNew Then
    lngExprID = val(grdValidation.Columns("ExprID").CellText(grdValidation.Bookmark))
    iType = val(grdValidation.Columns("Type").CellText(grdValidation.Bookmark))
    sMessage = grdValidation.Columns("Message").CellText(grdValidation.Bookmark)
    sTypeDesc = grdValidation.Columns("TypeDescription").CellText(grdValidation.Bookmark)
  End If

  With frmValidation
    .Initialise _
      mfReadOnly, _
      lngExprID, _
      iType, _
      sMessage, _
      mfrmCallingForm.CallingForm.WorkflowID, _
      mfrmCallingForm.BaseTable, _
      mfrmCallingForm.InitiationType, _
      maWFPrecedingAndCurrentElements, _
      maWFAllElements
      
    .Show vbModal
    
    With grdValidation
      .Redraw = False
    
      fRowModified = False
      
      If frmValidation.Cancelled Then
        ' Expression may have changed name or been deleted.
        If Not blnNew Then
          fRowModified = True
          lngRow = .AddItemRowIndex(.Bookmark)
          .RemoveItem lngRow
      
          sTemp = CStr(lngExprID) _
            & vbTab & GetExpressionName(lngExprID) _
            & vbTab & CStr(iType) _
            & vbTab & sTypeDesc _
            & vbTab & sMessage
        End If
      Else
        fRowModified = True
        
        If Not blnNew Then
          lngRow = .AddItemRowIndex(.Bookmark)
          .RemoveItem lngRow
        Else
          lngRow = .Rows
        End If

        sTemp = CStr(frmValidation.ValidationExprID) _
          & vbTab & GetExpressionName(frmValidation.ValidationExprID) _
          & vbTab & CStr(frmValidation.ValidationType) _
          & vbTab & WorkflowWebFormValidationTypeDescription(frmValidation.ValidationType) _
          & vbTab & frmValidation.Message
      End If
      
      If fRowModified Then
        .AddItem sTemp, lngRow
      End If
      
      .Redraw = True

      .Redraw = False
      .Bookmark = .AddItemBookmark(lngRow)
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      .Redraw = True

      RefreshValidationControls
      If Not frmValidation.Cancelled Then
        Changed = True
      End If
    End With
  End With
  
  UnLoad frmValidation
  Set frmValidation = Nothing

End Sub




Private Sub DefineRichTextMessage(piWhichMessage As WorkflowWebFormMessageType)
  Dim ctlMessageTypeCombo As Control
  Dim fOK As Boolean
  Dim frmMessage As frmRichTextEntry
  Dim sMessage As String
  Dim sDefaultMessage As String
  Dim fApplyDefault As Boolean
  
  fOK = True
  
  Select Case piWhichMessage
    Case WORKFLOWWEBFORMMESSAGE_COMPLETION
      Set ctlMessageTypeCombo = cboCompletionMessageType
      sMessage = msCompletionMessage
      sDefaultMessage = "Workflow step completed." & vbNewLine & vbNewLine & _
        "Click \ul here\ulnone  to close this form."
    Case WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER
      Set ctlMessageTypeCombo = cboSavedForLaterMessageType
      sMessage = msSavedForLaterMessage
      sDefaultMessage = "Workflow step saved for later." & vbNewLine & vbNewLine & _
        "Click \ul here\ulnone  to close this form."
    Case WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS
      Set ctlMessageTypeCombo = cboFollowOnFormsMessageType
      sMessage = msFollowOnFormsMessage
      sDefaultMessage = "Workflow step completed." & vbNewLine & vbNewLine & _
        "Click \ul here\ulnone  to complete the follow-on Workflow form(s)."
    Case Else
      fOK = False
  End Select

  If fOK Then
    fApplyDefault = ((Len(sMessage) = 0) And (Not mfReadOnly))
    If fApplyDefault Then
      sMessage = sDefaultMessage
    End If
    
    Set frmMessage = New frmRichTextEntry

    With frmMessage
      .Initialise _
        piWhichMessage, _
        sMessage, _
        mfReadOnly
        
      If fApplyDefault Then
        .Changed = True
      End If
      
      .Show vbModal

      If Not .Cancelled Then
        Select Case piWhichMessage
          Case WORKFLOWWEBFORMMESSAGE_COMPLETION
            Set ctlMessageTypeCombo = cboCompletionMessageType
            msCompletionMessage = .RichText
          Case WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER
            msSavedForLaterMessage = .RichText
          Case WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS
            msFollowOnFormsMessage = .RichText
        End Select
        
        cboMessage_refresh piWhichMessage, MESSAGE_CUSTOM
      End If

      Changed = True
    End With
  
    UnLoad frmMessage
    Set frmMessage = Nothing
  End If
End Sub





Private Sub RefreshIdentificationControls()
  Dim fEnable As Boolean
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION) _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION_WORKFLOWNAME) Then
    ' Disable the DescriptionHasWorkflowName controls if there is no Description Expression.
    fEnable = (mlngDescriptionExprID > 0) _
        And (Not mfReadOnly)
    
    EnableControl chkDescriptionHasWorkflowName, fEnable
    
    If (mlngDescriptionExprID = 0) Then
      chkDescriptionHasWorkflowName.value = vbUnchecked
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION) _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION_ELEMENTCAPTION) Then
    ' Disable the DescriptionHasWorkflowName controls if there is no Description Expression.
    fEnable = (mlngDescriptionExprID > 0) _
        And (Not mfReadOnly)

    EnableControl chkDescriptionHasElementCaption, fEnable
    
    If (mlngDescriptionExprID = 0) Then
      chkDescriptionHasElementCaption.value = vbUnchecked
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTION) Then
    If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTIONTYPE) Then
      
      fEnable = optCaptionType(giWFDATAVALUE_FIXED).value _
        And (Not mfReadOnly)
      EnableControl txtCaptionTypeValue, fEnable
      
      If WebFormItemHasProperty(miItemType, WFITEMPROP_CALCULATION) Then
        fEnable = optCaptionType(giWFDATAVALUE_CALC).value _
          And (Not mfReadOnly)
        EnableControl cmdCaptionTypeExpression, fEnable
        
        If optCaptionType(giWFDATAVALUE_CALC).value Then
          txtCaptionTypeValue.Text = ""
        Else
          mlngCaptionExprID = 0
          txtCaptionTypeExpression.Text = ""
        End If
      End If
    End If
  End If

End Sub


Private Sub RefreshDefaultValueControls()
  Dim fEnable As Boolean
  Dim fFixedDefault As Boolean
  Dim fCalcDefault As Boolean
  Dim lngOriginalID As Long
  Dim iExprType As Integer
  Dim iDataType As Integer
  Dim lngColumnID As Long
  Dim objExpr As CExpression
  Dim fValidExpr As Boolean
  Dim fDfltComboEnabled As Boolean
  
  fFixedDefault = True
  fCalcDefault = True
  
  fDfltComboEnabled = False
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
    fFixedDefault = optDefaultValueType(giWFDATAVALUE_FIXED).value
    fCalcDefault = optDefaultValueType(giWFDATAVALUE_CALC).value
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_CHAR) Then
    fEnable = fFixedDefault _
      And (Not mfReadOnly)
    
    EnableControl txtDefaultValue, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOGIC) Then
    fEnable = fFixedDefault _
      And (Not mfReadOnly)

    EnableControl fraLogicDefaults, fEnable
    EnableControl optDefaultValue(0), fEnable
    EnableControl optDefaultValue(1), fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_DATE) Then
    fEnable = fFixedDefault _
      And (Not mfReadOnly)

    EnableControl dtDefaultValue, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
    fEnable = fFixedDefault _
      And (Not mfReadOnly)

    EnableControl numDefaultValue, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LIST) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOOKUP) Then
    fEnable = fFixedDefault _
      And (Not mfReadOnly)

    fDfltComboEnabled = fEnable
    EnableControl cboDefaultValue, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_WORKPATTERN) Then
    ' FUTURE DEV
    fEnable = fFixedDefault _
      And (Not mfReadOnly)

    EnableControl wpDefaultValue, fEnable
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
    fEnable = fCalcDefault _
      And (Not mfReadOnly)

    EnableControl cmdDefaultValueExpression, fEnable
    
    If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPCOLUMNID) _
      And mlngDefaultValueExprID > 0 Then
      ' Check the Expression return type is still good.
      lngOriginalID = mlngDefaultValueExprID
       
      If cboLookupColumn.ListCount > 0 Then
        lngColumnID = cboLookupColumn.ItemData(cboLookupColumn.ListIndex)
        iExprType = giEXPRVALUE_UNDEFINED
        iDataType = GetColumnDataType(lngColumnID)

        Select Case iDataType
          Case dtVARCHAR
            iExprType = giEXPRVALUE_CHARACTER

          Case dtTIMESTAMP
            iExprType = giEXPRVALUE_DATE

          Case dtINTEGER
            iExprType = giEXPRVALUE_NUMERIC

          Case dtBIT
            iExprType = giEXPRVALUE_LOGIC

          Case dtNUMERIC
            iExprType = giEXPRVALUE_NUMERIC

          Case dtLONGVARCHAR
            iExprType = giEXPRVALUE_CHARACTER
        End Select
      
        If iExprType = giEXPRVALUE_UNDEFINED Then
          MsgBox "Unable to determine the Lookup Column data type. Default Value calculation selection has been cleared.", vbInformation + vbOKOnly, App.ProductName
          mlngDefaultValueExprID = 0
        Else
          ' Instantiate an expression object.
          Set objExpr = New CExpression

          With objExpr
            .ExpressionID = mlngDefaultValueExprID
            
            fValidExpr = .ReadExpressionDetails
            If fValidExpr Then
              If iExprType <> .ReturnType Then
                MsgBox "Default Value calculation data type no longer matches the Lookup Column data type. Default Value calculation selection has been cleared.", vbInformation + vbOKOnly, App.ProductName
                mlngDefaultValueExprID = 0
              End If
            Else
              MsgBox "Unable to determine the Default Value calculation data type. Default Value calculation selection has been cleared.", vbInformation + vbOKOnly, App.ProductName
              mlngDefaultValueExprID = 0
            End If
          End With
      
          Set objExpr = Nothing
        End If
      Else
        MsgBox "No Lookup Column selected. Default Value calculation selection has been cleared.", vbInformation + vbOKOnly, App.ProductName
        mlngDefaultValueExprID = 0
      End If
    
      If lngOriginalID <> mlngDefaultValueExprID Then
        ' Read the selected expression info.
        txtDefaultValueExpression.Text = GetExpressionName(mlngDefaultValueExprID)
      End If
    End If
  End If

  If Not fFixedDefault Then
    txtDefaultValue.Text = ""
    'JPD 20071031 Fault 12479
    'optDefaultValue(0).Value = True
    dtDefaultValue.Text = ""
    'spnDefaultValue.Value = 0
    numDefaultValue.value = 0
    cboDefaultValue.Clear
    
    'JPD 20070927 Fault 12496
    cboDefaultValue_refresh ""
    'wpDefaultValue.Value = 0
    
    EnableControl cboDefaultValue, fDfltComboEnabled
  End If
  
  If Not fCalcDefault Then
    mlngDefaultValueExprID = 0
    txtDefaultValueExpression.Text = ""
  End If

End Sub



Private Sub RefreshHeaderControls()
  Dim fEnable As Boolean
  Dim iHeaderLines As Integer
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_COLUMNHEADERS) Then
    If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADLINES) Then
      ' Disable the Headlines controls if ColumnHeaders is False.
      fEnable = (chkColumnHeaders.value = vbChecked) _
        And (Not mfReadOnly)

      EnableControl lblHeaderLines, fEnable
      EnableControl spnHeaderLines, fEnable
      
      If Not (chkColumnHeaders.value = vbChecked) Then
        spnHeaderLines.value = 0
      End If
      iHeaderLines = spnHeaderLines.value
    End If

    If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADFONT) Then
      ' Disable the Headlines controls if ColumnHeaders is False.
      fEnable = (chkColumnHeaders.value = vbChecked) _
        And (iHeaderLines > 0) _
        And (Not mfReadOnly)

      EnableControl lblHeaderFont, fEnable
      ' NPG20100428 Fault HRPRO-718
      ' EnableControl txtHeaderFont, fEnable
      EnableControl txtHeaderFont, vbFalse
      EnableControl cmdHeaderFont, fEnable
    End If
    
    If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADERBACKCOLOR) Then
      ' Cannot use EnableControl for the txtHeaderBackgroundColour control
      ' as the BackColor is non-standard (used to show the selected colour)
      fEnable = (chkColumnHeaders.value = vbChecked) _
        And (iHeaderLines > 0) _
        And (Not mfReadOnly)
      
      EnableControl lblHeaderBackgroundColour, fEnable
      EnableControl cmdHeaderBackgroundColour, fEnable
      
      With txtHeaderBackgroundColour
        .BackColor = IIf((chkColumnHeaders.value = vbChecked) _
          And (iHeaderLines > 0), mColHeaderBackColor, vbButtonFace)
        'JPD 20071206 Fault 12275
        '.Locked = Not fEnable
        '.TabStop = fEnable
        '.Enabled = fEnable
      End With
    End If
  End If

End Sub


Private Sub RefreshPictureControls()
  ' Refresh the Picture controls depending on the selected picture.
  Dim sFileName As String

  If mlngPictureID > 0 Then
    With recPictEdit
      .Index = "idxID"
      .Seek "=", mlngPictureID
      
      If Not .NoMatch Then
        txtPicture.Text = !Name
        sFileName = ReadPicture
        picPicture.Picture = LoadPicture(sFileName)
        picPicture.Move 0, 0, picPictureHolder.ScaleWidth, picPictureHolder.ScaleHeight
        SizeImage picPicture
        picPicture.Top = (picPictureHolder.ScaleHeight - picPicture.Height) \ 2
        picPicture.Left = (picPictureHolder.ScaleWidth - picPicture.Width) \ 2
        Kill sFileName
        picPictureHolder.Visible = True
      Else
        mlngPictureID = 0
      End If
    End With
  End If

  If mlngPictureID = 0 Then
    picPicture.Picture = LoadPicture("")
    txtPicture.Text = ""
  End If
  
  cmdPictureClear.Enabled = (mlngPictureID > 0) And (Not mfReadOnly)
  
End Sub


Private Sub RefreshScreen()
  ' Refresh the screen controls.
  Dim fOKToSave As Boolean
  
  fOKToSave = mfChanged And (Not mfReadOnly)
  
  cmdOK.Enabled = fOKToSave

End Sub

Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = mfCancelled
  
End Property
Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
  
End Property

Private Sub RefreshSizeControls()
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DECIMALS) Then
    With spnDecimals
      .MaximumValue = spnSize.value - 1
      
      If .value > .MaximumValue Then
        .value = .MaximumValue
      End If
    End With
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_CHAR) Then
    txtDefaultValue.MaxLength = IIf(spnSize.value > 65535, 0, spnSize.value)
  
    If Len(txtDefaultValue.Text) > spnSize.value Then
      txtDefaultValue.Text = Left(txtDefaultValue.Text, spnSize.value)
    End If
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
    RefreshDecimalsControls
  End If

End Sub

Private Sub RefreshDecimalsControls()
  Dim dblValue As Double
  Dim sFormat As String
  Dim iCount As Integer
  Dim dblMax As Double
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
    dblMax = 0
    
    'If spnDefaultValue.Visible Then
    '  dblValue = spnDefaultValue.Value
    'Else
      dblValue = numDefaultValue.value
    'End If
    
    'spnDefaultValue.Visible = (spnDecimals.Value = 0)
    'numDefaultValue.Visible = (spnDecimals.Value > 0)
    
    'If spnDecimals.Value = 0 Then
    '  spnDefaultValue.Value = CLng(dblValue)
    '
    '  For iCount = 1 To spnSize.Value
    '    dblMax = (dblMax * 10) + 9
    '  Next iCount
    '
    '  spnDefaultValue.MaximumValue = dblMax
    '  spnDefaultValue.MinimumValue = -dblMax
    '  If spnDefaultValue.Value > spnDefaultValue.MaximumValue Then
    '    spnDefaultValue.Value = spnDefaultValue.MaximumValue
    '  End If
    '  If spnDefaultValue.Value < spnDefaultValue.MinimumValue Then
    '    spnDefaultValue.Value = spnDefaultValue.MinimumValue
    '  End If
    'Else
      sFormat = ""
      For iCount = 1 To (spnSize.value - spnDecimals.value)
        sFormat = sFormat & "#"
      Next iCount

      If Len(sFormat) > 0 Then
        sFormat = Left(sFormat, Len(sFormat) - 1) & "0"
      End If

      If spnDecimals.value > 0 Then
        sFormat = sFormat & "."
        For iCount = 1 To spnDecimals.value
          sFormat = sFormat & "0"
        Next iCount
      End If

      If Len(sFormat) = 0 Then
        sFormat = "0"
      End If

      numDefaultValue.Format = sFormat
      numDefaultValue.DisplayFormat = numDefaultValue.Format
      numDefaultValue.value = dblValue
    'End If
  End If

End Sub


Private Sub SaveProperties()
  ' Write the properties to the selected control/form.
  Dim varControl As Variant
  Dim objFont As StdFont
  Dim sFileName As String
  Dim sList As String
  Dim sDefaultValue As String
  Dim asControlValues() As String
  Dim asFileExtensions() As String
  Dim iLoop As Integer
  Dim sNewList As String
  Dim sCaption As String
  Dim sExprCaption As String
  Dim varBookMark As Variant
  Dim lngExprID As Long
  Dim iValidationType As WorkflowWebFormValidationTypes
  Dim sMessage As String
  Dim asValidations() As String
  Dim fFixedDefault As Boolean
  Dim fCalcDefault As Boolean
  Dim sTemp As String
  Dim sOriginalIdentifier As String
  Dim lngOldParameter As Long
  Dim sNewIdentifier As String
  Dim lngNewParameter As Long
  Dim fMessageOK As Boolean
  Dim iIndex As Integer
  Dim objForm As Form
  
  sOriginalIdentifier = ""
  lngOldParameter = 0
  sNewIdentifier = ""
  lngNewParameter = 0
  
  If miItemType = giWFFORMITEM_FORM Then
    Set varControl = mfrmCallingForm
  Else
    Set varControl = mctlSelectedControl
  End If
  
  '++++++++++++++++++++++++++++++++++++++++++++++++++
  ' GENERAL TAB
  '++++++++++++++++++++++++++++++++++++++++++++++++++
  ' Identification frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_WFIDENTIFIER) Then
    sOriginalIdentifier = varControl.WFIdentifier
    varControl.WFIdentifier = txtIdentifier.Text
    sNewIdentifier = varControl.WFIdentifier
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_USEASTARGETIDENTIFIER) Then
    varControl.UseAsTargetIdentifier = chkUseAsTargetIdentifier.value
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTION) Then
    If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTIONTYPE) Then
      sCaption = txtCaptionTypeValue.Text
      
      If WebFormItemHasProperty(miItemType, WFITEMPROP_CALCULATION) Then
        If optCaptionType(giWFDATAVALUE_CALC).value Then
          sCaption = txtCaptionTypeExpression.Text
          
          If Len(Trim(sCaption)) = 0 Then
            sCaption = "<Calculated>"
          Else
            sCaption = "<" & sCaption & ">"
          End If
        End If
        
        varControl.CalculationID = mlngCaptionExprID
      End If
      
      varControl.CaptionType = IIf(optCaptionType(giWFDATAVALUE_CALC).value, giWFDATAVALUE_CALC, giWFDATAVALUE_FIXED)
      
    Else
      sCaption = txtCaption.Text
    End If
        
    'JPD 20081209 Fault 13410
    If miItemType = giWFFORMITEM_FORM Then
      Set objForm = varControl
      SetFormCaption objForm, sCaption
    Else
      varControl.Caption = Replace(sCaption, "&", "&&")
    End If
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION) Then
    varControl.DescriptionExprID = mlngDescriptionExprID
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION_WORKFLOWNAME) Then
    varControl.DescriptionHasWorkflowName = (chkDescriptionHasWorkflowName.value = vbChecked)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION_ELEMENTCAPTION) Then
    varControl.DescriptionHasElementCaption = (chkDescriptionHasElementCaption.value = vbChecked)
  End If
  
  '--------------------------------------------------
  ' Display frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_ORIENTATION) Then
    If optOrientation(0) Then
      varControl.Alignment = wfItemPropertyOrientation_Horizontal
    Else
      varControl.Alignment = wfItemPropertyOrientation_Vertical
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_VERTICALOFFSET) Then
    varControl.VerticalOffsetBehaviour = cboVOffsetBehaviour.ItemData(cboVOffsetBehaviour.ListIndex)
    varControl.VerticalOffset = PixelsToTwips(spnVOffset.value)
    
    If cboVOffsetBehaviour.ItemData(cboVOffsetBehaviour.ListIndex) = offsetTop Then
      varControl.Top = PixelsToTwips(spnVOffset.value)
    Else
      varControl.Top = PixelsToTwips(TwipsToPixels(mfrmCallingForm.ScaleHeight) - (spnVOffset.value + spnHeight.value))
    End If
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_TOP) Then
    varControl.Top = PixelsToTwips(spnTop.value)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HORIZONTALOFFSET) Then
    varControl.HorizontalOffsetBehaviour = cboHOffsetBehaviour.ItemData(cboHOffsetBehaviour.ListIndex)
    varControl.HorizontalOffset = PixelsToTwips(spnHOffset.value)
    
    If cboHOffsetBehaviour.ItemData(cboHOffsetBehaviour.ListIndex) = offsetLeft Then
      varControl.Left = PixelsToTwips(spnHOffset.value)
    Else
      varControl.Left = PixelsToTwips(TwipsToPixels(mfrmCallingForm.ScaleWidth) - (spnHOffset.value + spnWidth.value))
    End If
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_LEFT) Then
    varControl.Left = PixelsToTwips(spnLeft.value)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
    varControl.HeightBehaviour = cboHeightBehaviour.ItemData(cboHeightBehaviour.ListIndex)
    
    ' AE20080306 Fault #12985
    If varControl.HeightBehaviour <> behaveFixed Then
      varControl.Top = 0
      varControl.Height = PixelsToTwips(mfrmCallingForm.ScaleHeight)
    End If
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
    varControl.WidthBehaviour = cboWidthBehaviour.ItemData(cboWidthBehaviour.ListIndex)
    
    ' AE20080306 Fault #12985
    If varControl.WidthBehaviour <> behaveFixed Then
      varControl.Left = 0
      varControl.Width = PixelsToTwips(mfrmCallingForm.ScaleWidth)
    End If
  End If
   
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEIGHT) Then
    varControl.Height = PixelsToTwips(spnHeight.value)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTH) Then
    varControl.Width = PixelsToTwips(spnWidth.value)
  End If

  '--------------------------------------------------
  ' Behaviour frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_TIMEOUT) Then
    varControl.TimeoutFrequency = spnTimeoutFrequency.value
    If cboTimeoutPeriod.ListCount > 0 Then
      varControl.TimeoutPeriod = cboTimeoutPeriod.ItemData(cboTimeoutPeriod.ListIndex)
    Else
      varControl.TimeoutPeriod = TIMEOUT_DAY
    End If
    varControl.TimeoutExcludeWeekend = chkExcludeWeekends
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_SUBMITTYPE) Then
    varControl.Behaviour = IIf(optButtonAction(WORKFLOWBUTTONACTION_SAVEFORLATER).value, _
      WORKFLOWBUTTONACTION_SAVEFORLATER, _
      IIf(optButtonAction(WORKFLOWBUTTONACTION_CANCEL).value, _
        WORKFLOWBUTTONACTION_CANCEL, _
        WORKFLOWBUTTONACTION_SUBMIT))
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGETYPE) Then
    varControl.WFCompletionMessageType = cboCompletionMessageType.ItemData(cboCompletionMessageType.ListIndex)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGE) Then
    fMessageOK = True
    If WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGETYPE) Then
      fMessageOK = (cboCompletionMessageType.ItemData(cboCompletionMessageType.ListIndex) = MESSAGE_CUSTOM)
    End If
    
    varControl.WFCompletionMessage = IIf(fMessageOK, msCompletionMessage, "")
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGETYPE) Then
    varControl.WFSavedForLaterMessageType = cboSavedForLaterMessageType.ItemData(cboSavedForLaterMessageType.ListIndex)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGE) Then
    fMessageOK = True
    If WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGETYPE) Then
      fMessageOK = (cboSavedForLaterMessageType.ItemData(cboSavedForLaterMessageType.ListIndex) = MESSAGE_CUSTOM)
    End If
    
    varControl.WFSavedForLaterMessage = IIf(fMessageOK, msSavedForLaterMessage, "")
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGETYPE) Then
    varControl.WFFollowOnFormsMessageType = cboFollowOnFormsMessageType.ItemData(cboFollowOnFormsMessageType.ListIndex)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGE) Then
    fMessageOK = True
    If WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGETYPE) Then
      fMessageOK = (cboFollowOnFormsMessageType.ItemData(cboFollowOnFormsMessageType.ListIndex) = MESSAGE_CUSTOM)
    End If
    
    varControl.WFFollowOnFormsMessage = IIf(fMessageOK, msFollowOnFormsMessage, "")
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_REQUIRESAUTHENTICATION) Then
    varControl.RequiresAuthentication = chkRequireAuthentication
  End If
  
  
  '--------------------------------------------------
  ' HotSpot frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HOTSPOT) Then
    
    If cboHotSpotIdentifier.ListCount > 0 Then
      varControl.HotSpotIdentifier = cboHotSpotIdentifier.List(cboHotSpotIdentifier.ListIndex)
    Else
      varControl.HotSpotIdentifier = ""
    End If
    
  End If
  
  
  
  '++++++++++++++++++++++++++++++++++++++++++++++++++
  ' APPEARANCE TAB
  '++++++++++++++++++++++++++++++++++++++++++++++++++
  ' Options frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BORDERSTYLE) Then
    varControl.BorderStyle = IIf(chkBorder.value = vbChecked, vbFixedSingle, vbBSNone)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_ALIGNMENT) Then
    If cboAlignment.ListCount > 0 Then
      varControl.Alignment = cboAlignment.ItemData(cboAlignment.ListIndex)
    Else
      varControl.Alignment = vbLeftJustify
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_PASSWORDTYPE) Then
    varControl.PasswordType = (chkPasswordType.value = vbChecked)
  End If

  '--------------------------------------------------
  ' Header frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADLINES) Then
    varControl.HeadLines = spnHeaderLines.value
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_COLUMNHEADERS) Then
    varControl.ColumnHeaders = (chkColumnHeaders.value = vbChecked)
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADFONT) Then
    Set objFont = New StdFont
    objFont.Name = mObjHeadFont.Name
    objFont.Size = mObjHeadFont.Size
    objFont.Bold = mObjHeadFont.Bold
    objFont.Italic = mObjHeadFont.Italic
    objFont.Strikethrough = mObjHeadFont.Strikethrough
    objFont.Underline = mObjHeadFont.Underline
    Set varControl.HeadFont = objFont
    Set objFont = Nothing
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADERBACKCOLOR) Then
    varControl.HeaderBackColor = mColHeaderBackColor
  End If
  
  '--------------------------------------------------
  ' Foreground frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FONT) Then
    Set objFont = New StdFont
    objFont.Name = mObjFont.Name
    objFont.Size = mObjFont.Size
    objFont.Bold = mObjFont.Bold
    objFont.Italic = mObjFont.Italic
    objFont.Strikethrough = mObjFont.Strikethrough
    objFont.Underline = mObjFont.Underline
    Set varControl.Font = objFont
    Set objFont = Nothing
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLOR) Then
    varControl.ForeColor = mColForeColor
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLOREVEN) Then
    varControl.ForeColorEven = mColForeColorEven
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLORODD) Then
    varControl.ForeColorOdd = mColForeColorOdd
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
    varControl.ForeColorHighlight = mColForeColorHighlight
  End If
  
  '--------------------------------------------------
  ' Background frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKSTYLE) Then
    If cboBackgroundStyle.ListCount > 0 Then
      varControl.BackStyle = cboBackgroundStyle.ItemData(cboBackgroundStyle.ListIndex)
    Else
      varControl.BackStyle = BACKSTYLE_OPAQUE
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLOR) Then
    varControl.BackColor = mColBackColor
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLOREVEN) Then
    varControl.BackColorEven = mColBackColorEven
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLORODD) Then
    varControl.BackColorOdd = mColBackColorOdd
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
    varControl.BackColorHighlight = mColBackColorHighlight
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_PICTURE) Then
    If miItemType = giWFFORMITEM_FORM Then
      If mlngPictureID > 0 Then
        recPictEdit.Index = "idxID"
        recPictEdit.Seek "=", mlngPictureID

        If Not recPictEdit.NoMatch Then
          sFileName = ReadPicture
          varControl.PictureID = mlngPictureID
          varControl.Picture = LoadPicture(sFileName)
          Kill sFileName
        End If
      Else
        varControl.PictureID = 0
        varControl.Picture = LoadPicture("")
      End If
    Else
      If mlngPictureID > 0 Then
        recPictEdit.Index = "idxID"
        recPictEdit.Seek "=", mlngPictureID

        If Not recPictEdit.NoMatch Then
          sFileName = ReadPicture
          varControl.PictureID = mlngPictureID
          varControl.Picture = sFileName
          Kill sFileName
        End If
      Else
        'varControl.Picture = ""
        varControl.PictureID = 0
      End If
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_PICTURELOCATION) Then
    If cboPictureLocation.ListCount > 0 Then
      varControl.PictureLocation = cboPictureLocation.ItemData(cboPictureLocation.ListIndex)
    Else
      varControl.PictureLocation = 0
    End If
  End If

  '++++++++++++++++++++++++++++++++++++++++++++++++++
  ' DATA TAB
  '++++++++++++++++++++++++++++++++++++++++++++++++++
  ' RecordIdentification frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_TABLEID) Then
    lngOldParameter = varControl.TableID
    
    If cboRecordIdentificationTable.ListCount > 0 Then
      varControl.TableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
    Else
      varControl.TableID = 0
    End If
  
    lngNewParameter = varControl.TableID
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DBRECORD) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_RECSELTYPE) Then
    
    If cboRecordIdentificationRecord.ListCount > 0 Then
      varControl.WFDatabaseRecord = cboRecordIdentificationRecord.ItemData(cboRecordIdentificationRecord.ListIndex)
    Else
      varControl.WFDatabaseRecord = giWFRECSEL_UNKNOWN
    End If
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_ELEMENTIDENTIFIER) Then
    If cboRecordIdentificationElement.ListCount > 0 Then
      varControl.WFWorkflowForm = cboRecordIdentificationElement.List(cboRecordIdentificationElement.ListIndex)
    Else
      varControl.WFWorkflowForm = ""
    End If
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDSELECTOR) Then
    If cboRecordIdentificationRecordSelector.ListCount > 0 Then
      varControl.WFWorkflowValue = cboRecordIdentificationRecordSelector.List(cboRecordIdentificationRecordSelector.ListIndex)
    Else
      varControl.WFWorkflowValue = ""
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDTABLEID) Then
    If cboRecordIdentificationRecordTable.ListCount > 0 Then
      varControl.WFRecordTableID = cboRecordIdentificationRecordTable.ItemData(cboRecordIdentificationRecordTable.ListIndex)
    Else
      varControl.WFRecordTableID = 0
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDORDER) Then
    varControl.WFRecordOrderID = mlngRecordIdentificationOrderID
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDFILTER) Then
    varControl.WFRecordFilterID = mlngRecordIdentificationFilterID
  End If
  
  '--------------------------------------------------
  ' Size frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_SIZE) Then
    varControl.WFInputSize = spnSize.value
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DECIMALS) Then
    varControl.WFInputDecimals = spnDecimals.value
  End If

  '--------------------------------------------------
  ' Lookup frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPTABLEID) Then
    If cboLookupTable.ListCount > 0 Then
      varControl.LookupTableID = cboLookupTable.ItemData(cboLookupTable.ListIndex)
    Else
      varControl.LookupTableID = 0
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPCOLUMNID) Then
    If cboLookupColumn.ListCount > 0 Then
      lngOldParameter = GetColumnDataType(varControl.LookupColumnID)
      
      If cboLookupColumn.ListCount > 0 Then
        varControl.LookupColumnID = cboLookupColumn.ItemData(cboLookupColumn.ListIndex)
      Else
        varControl.LookupColumnID = 0
      End If

      lngNewParameter = GetColumnDataType(varControl.LookupColumnID)
    Else
      varControl.LookupColumnID = 0
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPORDER) Then
    varControl.LookupOrderID = mlngLookupOrderID
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTERCOLUMN) Then
    If cboLookupFilterColumn.ListCount > 0 Then
      varControl.LookupFilterColumn = cboLookupFilterColumn.ItemData(cboLookupFilterColumn.ListIndex)
    Else
      varControl.LookupFilterColumn = 0
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTEROPERATOR) Then
    If cboLookupFilterOperator.ListCount > 0 Then
      varControl.LookupFilterOperator = cboLookupFilterOperator.ItemData(cboLookupFilterOperator.ListIndex)
    Else
      varControl.LookupFilterOperator = 0
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTERVALUE) Then
    If cboLookupFilterValue.ListCount > 0 Then
      varControl.LookupFilterValue = cboLookupFilterValue.List(cboLookupFilterValue.ListIndex)
    Else
      varControl.LookupFilterValue = ""
    End If
  End If

  '--------------------------------------------------
  ' ControlValues frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_CONTROLVALUELIST) Then
    sList = MergeControlValues(txtControlValues.Text)
    
    'JPD 20070416 Fault 12127
    asControlValues() = Split(sList, vbTab)
    sList = ""
    
    For iLoop = 0 To UBound(asControlValues)
      If Len(Trim(asControlValues(iLoop))) > 0 Then
        sList = sList & IIf(Len(sList) > 0, vbTab, "") & asControlValues(iLoop)
      End If
    Next iLoop

    If miItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
      If Len(sList) > 0 Then
        varControl.ClearOptions
        varControl.NoOptions = False
        varControl.SetOptions Split(sList, vbTab)
      Else
        varControl.NoOptions = True
        varControl.Height = 0 ' Dummy call to get the control to resize itself.
      End If
    Else
      varControl.ControlValueList = sList
    End If
  End If
  
  '--------------------------------------------------
  ' Validation frame
  '--------------------------------------------------
  If WebFormItemHasProperty(miItemType, WFITEMPROP_MANDATORY) Then
    varControl.Mandatory = (chkValidationMandatory.value = vbChecked)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_VALIDATION) Then
    ReDim asValidations(4, 0)
  
    With grdValidation
      For iLoop = 0 To (.Rows - 1)
        varBookMark = .AddItemBookmark(iLoop)
  
        lngExprID = val(.Columns("ExprID").CellText(varBookMark))
        iValidationType = val(.Columns("Type").CellText(varBookMark))
        sMessage = Trim(.Columns("Message").CellText(varBookMark))
        
        ReDim Preserve asValidations(4, UBound(asValidations, 2) + 1)
        asValidations(1, UBound(asValidations, 2)) = CStr(lngExprID)
        asValidations(2, UBound(asValidations, 2)) = CStr(iValidationType)
        asValidations(3, UBound(asValidations, 2)) = sMessage
      Next iLoop
    End With
    
    varControl.Validations = asValidations
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_FILEEXTENSIONS) Then
    sList = MergeControlValues(txtFileExtensions.Text)

    asFileExtensions() = Split(sList, vbTab)
    sList = ""

    For iLoop = 0 To UBound(asFileExtensions)
      sTemp = Trim(asFileExtensions(iLoop))
      iIndex = InStrRev(sTemp, ".")
      If iIndex > 0 Then
        sTemp = Mid(sTemp, iIndex + 1)
      End If
      
      If Len(sTemp) > 0 Then
        sList = sList & IIf(Len(sList) > 0, vbTab, "") & sTemp
      End If
    Next iLoop

    varControl.WFFileExtensions = sList
  End If

  '--------------------------------------------------
  ' DefaultValue frame
  '--------------------------------------------------
  fFixedDefault = True
  fCalcDefault = True
  sExprCaption = ""

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
    fFixedDefault = optDefaultValueType(giWFDATAVALUE_FIXED).value
    fCalcDefault = optDefaultValueType(giWFDATAVALUE_CALC).value
  
    varControl.DefaultValueType = IIf(optDefaultValueType(giWFDATAVALUE_CALC).value, giWFDATAVALUE_CALC, giWFDATAVALUE_FIXED)
  End If
  
  If fCalcDefault _
    And WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
  
    sExprCaption = txtDefaultValueExpression.Text
  
    If Len(Trim(sExprCaption)) = 0 Then
      sExprCaption = "<Calculated>"
    Else
      sExprCaption = "<" & sExprCaption & ">"
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_CHAR) Then
    sTemp = Left(txtDefaultValue.Text, Minimum(spnSize.value, 65535))
    varControl.WFDefaultCharValue = IIf(fFixedDefault, sTemp, "")
    varControl.Caption = " " & IIf(fFixedDefault, sTemp, sExprCaption)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOGIC) Then
    varControl.WFDefaultValue = IIf(fFixedDefault, optDefaultValue(0).value, True)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_DATE) Then
    sTemp = dtDefaultValue.Text
    If Not IsDate(sTemp) Then
      sTemp = ""
    End If

    varControl.WFDefaultValueDateString = IIf(fFixedDefault, sTemp, "")
    varControl.Caption = " " & IIf(fFixedDefault, sTemp, sExprCaption)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
    ' TODO - future dev
    'If spnDecimals.Value = 0 Then
    '  varControl.WFDefaultNumericValue = CDbl(spnDefaultValue.Value)
    'Else
      varControl.WFDefaultNumericValue = IIf(fFixedDefault, CDbl(numDefaultValue.value), 0)
      varControl.Caption = " " & IIf(fFixedDefault, varControl.WFDefaultNumericValue, sExprCaption)
    'End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LIST) Then
    sDefaultValue = vbNullString

    If cboDefaultValue.ListCount > 0 Then
      If miItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
        sDefaultValue = cboDefaultValue.List(cboDefaultValue.ListIndex)
      ElseIf cboDefaultValue.ListIndex > 0 Then
        sDefaultValue = cboDefaultValue.List(cboDefaultValue.ListIndex)
      End If
    End If

    If miItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
      If Not varControl.SelectOption(sDefaultValue) Then
        varControl.DefaultStringValue = ""
      Else
        varControl.DefaultStringValue = IIf(fFixedDefault, sDefaultValue, "")
      End If
    Else
      sList = varControl.ControlValueList
      If (InStr(1, sList, sDefaultValue) = 0) Or _
        (sDefaultValue = vbNullString) Then
        sDefaultValue = ""
      End If

      varControl.DefaultStringValue = IIf(fFixedDefault, sDefaultValue, "")
      varControl.Caption = IIf(fFixedDefault, sDefaultValue, sExprCaption)
    End If
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOOKUP) Then
    sDefaultValue = vbNullString

    If cboDefaultValue.ListCount > 0 Then
      If cboDefaultValue.ListIndex > 0 Then
        sDefaultValue = cboDefaultValue.List(cboDefaultValue.ListIndex)
      End If
    End If

    varControl.DefaultStringValue = IIf(fFixedDefault, sDefaultValue, "")
    varControl.Caption = IIf(fFixedDefault, sDefaultValue, sExprCaption)
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_WORKPATTERN) Then
    ' TODO future dev
    'wpDefaultValue
  End If
  
  ' Format the DefaultValue (expression) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
    varControl.CalculationID = IIf(fCalcDefault, mlngDefaultValueExprID, 0)
  End If

  ' Update identifiers used in expressions
  mfrmCallingForm.UpdateIdentifiers (miItemType = giWFFORMITEM_FORM), _
    sOriginalIdentifier, _
    sNewIdentifier, _
    lngOldParameter, _
    lngNewParameter

End Sub

Private Function ValidateProperties() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fContinue As Boolean
  Dim frmUsage As frmUsage
  Dim asMessages() As String
  Dim iLoop As Integer
  Dim lngExprID As Long
  Dim sMessage As String
  Dim fInvalidValidationExpr As Boolean
  Dim fInvalidValidationMessage As Boolean
  Dim varBookMark As Variant
  Dim fCalcDefault As Boolean
  Dim fValid As Boolean
  
  '-----------------------------------------------------------------------
  ' First do the validation that CANNOT be overridden.
  '-----------------------------------------------------------------------
  ReDim asMessages(0)
  
  ' Ensure that lookup filtering is fully defined if required.
  fValid = True
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTER) Then
    If chkLookupFilter.value = vbChecked Then
      If (cboLookupFilterColumn.ListCount = 0) _
        Or (cboLookupFilterOperator.ListCount = 0) _
        Or (cboLookupFilterValue.ListCount = 0) Then
        
        fValid = False
      End If
      
      If fValid Then
        fValid = (cboLookupFilterColumn.ListIndex > 0) _
          And (cboLookupFilterOperator.ItemData(cboLookupFilterOperator.ListIndex) > 0) _
          And (cboLookupFilterValue.ListIndex > 0)
      End If
    End If
  End If
  
  If Not fValid Then
    ReDim Preserve asMessages(UBound(asMessages) + 1)
    asMessages(UBound(asMessages)) = "Lookup filter details incomplete."
  End If
  
  If (UBound(asMessages) > 0) Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList

    For iLoop = 1 To UBound(asMessages)
      frmUsage.AddToList (asMessages(iLoop))
    Next iLoop

    Screen.MousePointer = vbDefault
    frmUsage.ShowMessage "Workflow", "The " & GetWebFormItemTypeName(CInt(miItemType)) & " definition is invalid for the reasons listed below.", UsageCheckObject.Workflow, _
      USAGEBUTTONS_OK, "validation"

    UnLoad frmUsage
    Set frmUsage = Nothing
    
    ValidateProperties = False
    Exit Function
  End If
  
  '-----------------------------------------------------------------------
  ' Now do the validation that CAN be overridden.
  '-----------------------------------------------------------------------
  fContinue = True
  ReDim asMessages(0)
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTION) Then
    If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTIONTYPE) Then
      If WebFormItemHasProperty(miItemType, WFITEMPROP_CALCULATION) Then
        If optCaptionType(giWFDATAVALUE_CALC).value _
          And mlngCaptionExprID = 0 Then
        
          ReDim Preserve asMessages(UBound(asMessages) + 1)
          asMessages(UBound(asMessages)) = "No Caption calculation selected."
        End If
      End If
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
    fCalcDefault = True
  
    If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
      If Not optDefaultValueType(giWFDATAVALUE_CALC).value Then
        fCalcDefault = False
      End If
    End If

    If fCalcDefault _
      And mlngDefaultValueExprID = 0 Then

      ReDim Preserve asMessages(UBound(asMessages) + 1)
      asMessages(UBound(asMessages)) = "No Default Value calculation selected."
    End If
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_VALIDATION) Then
    fInvalidValidationExpr = False
    fInvalidValidationMessage = False
    
    With grdValidation
      For iLoop = 0 To (.Rows - 1)
        varBookMark = .AddItemBookmark(iLoop)
  
        lngExprID = val(.Columns("ExprID").CellText(varBookMark))
        sMessage = Trim(.Columns("Message").CellText(varBookMark))
        
        If lngExprID <= 0 Then
          fInvalidValidationExpr = True
        End If
        
        If Len(sMessage) = 0 Then
          fInvalidValidationMessage = True
        End If
      Next iLoop
    End With

    If fInvalidValidationExpr Then
      ReDim Preserve asMessages(UBound(asMessages) + 1)
      asMessages(UBound(asMessages)) = "Validation(s) with no calculation selected."
    End If
    
    If fInvalidValidationMessage Then
      ReDim Preserve asMessages(UBound(asMessages) + 1)
      asMessages(UBound(asMessages)) = "Validation(s) with no message."
    End If
  End If

  ' Display the validity failures to the user.
  fContinue = (UBound(asMessages) = 0)

  If Not fContinue Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList

    For iLoop = 1 To UBound(asMessages)
      frmUsage.AddToList (asMessages(iLoop))
    Next iLoop

    Screen.MousePointer = vbDefault
    frmUsage.ShowMessage "Workflow", "The " & GetWebFormItemTypeName(CInt(miItemType)) & " definition is invalid for the reasons listed below." & _
      vbCrLf & "Do you wish to continue?", UsageCheckObject.Workflow, _
      USAGEBUTTONS_YES + USAGEBUTTONS_NO + USAGEBUTTONS_PRINT, "validation"

    fContinue = (frmUsage.Choice = vbYes)

    UnLoad frmUsage
    Set frmUsage = Nothing
  End If

TidyUpAndExit:
  ValidateProperties = fContinue
  Exit Function
  
ErrorTrap:
  fContinue = True
  Resume TidyUpAndExit
  
End Function

Private Sub cboAlignment_Click()
    Changed = True

End Sub


Private Sub cboBackgroundStyle_Click()
  RefreshBackgroundControls
  Changed = True

End Sub


Private Sub cboCompletionMessageType_Click()
  RefreshBehaviourControls
  Changed = True

End Sub


Private Sub cboDefaultValue_Click()
  Changed = True

End Sub

Private Sub cboFollowOnFormsMessageType_Click()
  RefreshBehaviourControls
  Changed = True

End Sub


Private Sub cboHeightBehaviour_Click()
  If (mfLoading Or cboHeightBehaviour.ListIndex = -1) Then Exit Sub
  
  If cboHeightBehaviour.ItemData(cboHeightBehaviour.ListIndex) <> miHeightBehaviour Then
    EnableControl lblHeightValue, (Not lblHeightValue.Enabled)
    EnableControl spnHeight, (Not spnHeight.Enabled)
                
    spnHeight.value = TwipsToPixels(mfrmCallingForm.ScaleHeight)
    spnTop.value = 0
    
    miHeightBehaviour = cboHeightBehaviour.ItemData(cboHeightBehaviour.ListIndex)
    Changed = True
  End If
End Sub

Private Sub cboHotSpotIdentifier_Click()
  Changed = True
    
End Sub

Private Sub cboLookupColumn_Click()
  Dim sCurrentDefault As String
  
  sCurrentDefault = ""
  
  If cboDefaultValue.ListCount > 0 Then
    If cboDefaultValue.ListIndex > 0 Then
      sCurrentDefault = cboDefaultValue.List(cboDefaultValue.ListIndex)
    End If
  End If
  
  cboDefaultValue_refresh sCurrentDefault
  Changed = True
  RefreshDefaultValueControls
  
End Sub


Private Sub cboLookupFilterColumn_Click()
  Dim lngOperator As Long
  Dim sValue As String
  
  lngOperator = 0
  If cboLookupFilterOperator.ListIndex >= 0 Then
    lngOperator = cboLookupFilterOperator.ItemData(cboLookupFilterOperator.ListIndex)
  End If
  
  sValue = ""
  If cboLookupFilterValue.ListIndex >= 0 Then
    sValue = cboLookupFilterValue.List(cboLookupFilterValue.ListIndex)
  End If
  
  cboLookupFilterOperator_Refresh lngOperator
  cboLookupFilterValue_Refresh sValue
  
  Changed = True
End Sub

Private Sub cboLookupFilterOperator_Click()
  Changed = True
End Sub

Private Sub cboLookupFilterValue_Click()
  Changed = True
End Sub

Private Sub cboHOffsetBehaviour_Click()
  
  If (mfLoading Or cboHOffsetBehaviour.ListIndex = -1) Then Exit Sub
  
  If cboHOffsetBehaviour.ItemData(cboHOffsetBehaviour.ListIndex) <> miHOffsetBehaviour Then
    spnHOffset.value = TwipsToPixels(mfrmCallingForm.ScaleWidth) - (spnHOffset.value + spnWidth.value)
    miHOffsetBehaviour = cboHOffsetBehaviour.ItemData(cboHOffsetBehaviour.ListIndex)
    
    Changed = True
  End If
End Sub

Private Sub cboSavedForLaterMessageType_Click()
  RefreshBehaviourControls
  Changed = True

End Sub


Private Sub cboVOffsetBehaviour_Click()
    
  If (mfLoading Or cboVOffsetBehaviour.ListIndex = -1) Then Exit Sub
  
  If cboVOffsetBehaviour.ItemData(cboVOffsetBehaviour.ListIndex) <> miVOffsetBehaviour Then
    spnVOffset.value = TwipsToPixels(mfrmCallingForm.ScaleHeight) - spnHeight.value - spnVOffset.value
    miVOffsetBehaviour = cboVOffsetBehaviour.ItemData(cboVOffsetBehaviour.ListIndex)
    
    Changed = True
  End If

End Sub

Private Sub cboPictureLocation_Click()
  Changed = True
End Sub


Private Sub cboRecordIdentificationElement_Click()
  Dim sCurrentRecordSelector As String

  If cboRecordIdentificationRecordSelector.ListCount > 0 Then
    sCurrentRecordSelector = cboRecordIdentificationRecordSelector.List(cboRecordIdentificationRecordSelector.ListIndex)
  Else
    sCurrentRecordSelector = ""
  End If

  cboRecordIdentificationRecordSelector_refresh sCurrentRecordSelector
  
  Changed = True

End Sub


Private Sub cboRecordIdentificationRecord_Click()
  Dim sCurrentElement As String

  If cboRecordIdentificationElement.ListCount > 0 Then
    sCurrentElement = cboRecordIdentificationElement.List(cboRecordIdentificationElement.ListIndex)
  Else
    sCurrentElement = ""
  End If

  cboRecordIdentificationElement_refresh sCurrentElement
  
  Changed = True

End Sub


Private Sub cboRecordIdentificationRecordSelector_Click()
  Dim lngCurrentRecordTable As Long

  If cboRecordIdentificationRecordTable.ListCount > 0 Then
    lngCurrentRecordTable = cboRecordIdentificationRecordTable.ItemData(cboRecordIdentificationRecordTable.ListIndex)
  Else
    lngCurrentRecordTable = 0
  End If

  cboRecordIdentificationRecordTable_refresh lngCurrentRecordTable
  
  Changed = True

End Sub

Private Sub cboRecordIdentificationRecordTable_Click()
  Changed = True
End Sub

Private Sub cboRecordIdentificationTable_Click()
  Dim lngTableID As Long
  Dim iCurrentRecord As WorkflowRecordSelectorTypes

  If cboRecordIdentificationRecord.ListCount > 0 Then
    iCurrentRecord = cboRecordIdentificationRecord.ItemData(cboRecordIdentificationRecord.ListIndex)
  Else
    iCurrentRecord = giWFRECSEL_UNKNOWN
  End If

  cboRecordIdentificationRecord_refresh iCurrentRecord
  
  If cboRecordIdentificationTable.ListCount > 0 Then
    lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
  End If
  
  If Not CheckOrder(mlngRecordIdentificationOrderID, lngTableID) Then
    mlngRecordIdentificationOrderID = 0
    txtRecordIdentificationOrder.Text = ""
  End If
  
  If Not CheckExpression(mlngRecordIdentificationFilterID, lngTableID, True) Then
    mlngRecordIdentificationFilterID = 0
    txtRecordIdentificationFilter.Text = ""
  End If
  
  Changed = True

End Sub

Private Sub cboLookupTable_Click()
  Dim lngTableID As Long
  Dim lngCurrentColumnID As Long
  
  If cboLookupColumn.ListCount > 0 Then
    lngCurrentColumnID = cboLookupColumn.ItemData(cboLookupColumn.ListIndex)
  Else
    lngCurrentColumnID = 0
  End If
  
  cboLookupColumn_refresh lngCurrentColumnID
  
  If cboLookupFilterColumn.ListCount > 0 Then
    lngCurrentColumnID = cboLookupFilterColumn.ItemData(cboLookupFilterColumn.ListIndex)
  Else
    lngCurrentColumnID = 0
  End If
  
  cboLookupFilterColumn_Refresh lngCurrentColumnID
  
  If cboLookupTable.ListCount > 0 Then
    lngTableID = cboLookupTable.ItemData(cboLookupTable.ListIndex)
  End If
  
  If Not CheckOrder(mlngLookupOrderID, lngTableID) Then
    mlngLookupOrderID = 0
    txtLookupOrder.Text = ""
  End If
  
  Changed = True
  
End Sub

Private Sub cboTimeoutPeriod_Click()
  If Not mfReadOnly Then
    chkExcludeWeekends.Enabled = (cboTimeoutPeriod.ItemData(cboTimeoutPeriod.ListIndex) = TIMEOUT_DAY)
    ' AE20080317 Fault #13017
    chkExcludeWeekends = IIf(chkExcludeWeekends.Enabled, chkExcludeWeekends, vbUnchecked)
  End If
  Changed = True
End Sub

Private Sub cboWidthBehaviour_Click()
    If (mfLoading Or cboWidthBehaviour.ListIndex = -1) Then Exit Sub
  
  If cboWidthBehaviour.ItemData(cboWidthBehaviour.ListIndex) <> miWidthBehaviour Then
    EnableControl lblWidthValue, (Not lblWidthValue.Enabled)
    EnableControl spnWidth, (Not spnWidth.Enabled)
                
    spnWidth.value = TwipsToPixels(mfrmCallingForm.ScaleWidth)
    spnLeft.value = 0
    
    miWidthBehaviour = cboWidthBehaviour.ItemData(cboWidthBehaviour.ListIndex)
    Changed = True
  End If
End Sub

Private Sub chkBorder_Click()
  AutoResizeControl
  Changed = True
  
End Sub

Private Sub chkColumnHeaders_Click()
  RefreshHeaderControls
  Changed = True

End Sub


Private Sub chkDescriptionHasElementCaption_Click()
  Changed = True

End Sub


Private Sub chkDescriptionHasWorkflowName_Click()
  Changed = True

End Sub


Private Sub chkExcludeWeekends_Click()
  Changed = True
End Sub

Private Sub chkLookupFilter_Click()
  cboLookupFilterColumn_Refresh 0
  cboLookupFilterOperator_Refresh 0
  cboLookupFilterValue_Refresh ""
    
  Changed = True

End Sub

Private Sub chkPasswordType_Click()
  Changed = True

End Sub

Private Sub chkRequireAuthentication_Click()
  Changed = True
End Sub

Private Sub chkUseAsTargetIdentifier_Click()
  Changed = True
End Sub

Private Sub chkValidationMandatory_Click()
  Changed = True
End Sub

Private Sub cmdBackgroundColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColBackColor
'    .ShowColor
'
'    If mColBackColor <> .Color Then
'      mColBackColor = .Color
'      Changed = True
'
'      txtBackgroundColour.BackColor = mColBackColor
'    End If
'  End With
  
  With colPickDlg
    .Color = mColBackColor
    .ShowPalette
    
    If mColBackColor <> .Color Then
      mColBackColor = .Color
      Changed = True
      
      txtBackgroundColour.BackColor = mColBackColor
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False
  
End Sub

Private Sub cmdBackgroundEvenColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColBackColorEven
'    .ShowColor
'
'    If mColBackColorEven <> .Color Then
'      mColBackColorEven = .Color
'      Changed = True
'
'      txtBackgroundEvenColour.BackColor = mColBackColorEven
'    End If
'  End With

  With colPickDlg
    .Color = mColBackColorEven
    .ShowPalette
    
    If mColBackColorEven <> .Color Then
      mColBackColorEven = .Color
      Changed = True
      
      txtBackgroundEvenColour.BackColor = mColBackColorEven
    End If
  End With
  
ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub


Private Sub cmdBackgroundHighlightColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColBackColorHighlight
'    .ShowColor
'
'    If mColBackColorHighlight <> .Color Then
'      mColBackColorHighlight = .Color
'      Changed = True
'
'      txtBackgroundHighlightColour.BackColor = mColBackColorHighlight
'    End If
'  End With

  With colPickDlg
    .Color = mColBackColorHighlight
    .ShowPalette
    
    If mColBackColorHighlight <> .Color Then
      mColBackColorHighlight = .Color
      Changed = True
      
      txtBackgroundHighlightColour.BackColor = mColBackColorHighlight
    End If
  End With
  
ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub


Private Sub cmdBackgroundOddColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColBackColorOdd
'    .ShowColor
'
'    If mColBackColorOdd <> .Color Then
'      mColBackColorOdd = .Color
'      Changed = True
'
'      txtBackgroundOddColour.BackColor = mColBackColorOdd
'    End If
'  End With

  With colPickDlg
    .Color = mColBackColorOdd
    .ShowPalette
    
    If mColBackColorOdd <> .Color Then
      mColBackColorOdd = .Color
      Changed = True
      
      txtBackgroundOddColour.BackColor = mColBackColorOdd
    End If
  End With
  
ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub


Private Sub cmdCancel_Click()
Dim iAnswer As Integer

  'Check if any changes have been made.
  If mfChanged Then
    iAnswer = MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
    If iAnswer = vbYes Then
      Call cmdOK_Click
      ' Flag that the copy has been cancelled..
      mfCancelled = False
    ElseIf iAnswer = vbNo Then
      Me.Cancelled = True
    ElseIf iAnswer = vbCancel Then
      ' Flag that the copy has been cancelled..
      mfCancelled = False
      Exit Sub
    End If
  Else
    Me.Cancelled = True
    mfCancelled = True
  End If
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdCaptionTypeExpression_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long
  
  lngOriginalID = mlngCaptionExprID

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise 0, mlngCaptionExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_CHARACTER
    .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
    .UtilityBaseTable = mfrmCallingForm.BaseTable
    .WorkflowInitiationType = mfrmCallingForm.InitiationType
    .PrecedingWorkflowElements = maWFPrecedingElements
    .AllWorkflowElements = maWFAllElements

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression(mfReadOnly) Then
      mlngCaptionExprID = .ExpressionID
    Else
      ' Check in case the original expression has been deleted.
      If Not CheckExpression(mlngCaptionExprID, 0, False) Then
        mlngCaptionExprID = 0
      End If
    End If

    ' Read the selected expression info.
    txtCaptionTypeExpression.Text = GetExpressionName(mlngCaptionExprID)
  End With

  Set objExpr = Nothing

  If lngOriginalID <> mlngCaptionExprID Then
    RefreshIdentificationControls
    Changed = True
  End If

End Sub

Private Sub cmdCompletionMessage_Click()
  
  DefineRichTextMessage WORKFLOWWEBFORMMESSAGE_COMPLETION
  
End Sub

Private Sub cmdDefaultValueExpression_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long
  Dim iExprType As Integer
  Dim iDataType As Integer
  Dim lngColumnID As Long
  
  lngOriginalID = mlngDefaultValueExprID
  iExprType = giEXPRVALUE_UNDEFINED

  Select Case miItemType
    Case giWFFORMITEM_INPUTVALUE_CHAR
      iExprType = giEXPRVALUE_CHARACTER
      
    Case giWFFORMITEM_INPUTVALUE_DATE
      iExprType = giEXPRVALUE_DATE
      
    Case giWFFORMITEM_INPUTVALUE_LOGIC
      iExprType = giEXPRVALUE_LOGIC
    
    Case giWFFORMITEM_INPUTVALUE_DROPDOWN
      iExprType = giEXPRVALUE_CHARACTER
    
    Case giWFFORMITEM_INPUTVALUE_LOOKUP
      If cboLookupColumn.ListCount > 0 Then
        lngColumnID = cboLookupColumn.ItemData(cboLookupColumn.ListIndex)
      
        iDataType = GetColumnDataType(lngColumnID)
        
        Select Case iDataType
          Case dtVARCHAR
            iExprType = giEXPRVALUE_CHARACTER
          
          Case dtTIMESTAMP
            iExprType = giEXPRVALUE_DATE
          
          Case dtINTEGER
            iExprType = giEXPRVALUE_NUMERIC
 
          Case dtBIT
            iExprType = giEXPRVALUE_LOGIC

          Case dtNUMERIC
            iExprType = giEXPRVALUE_NUMERIC

          Case dtLONGVARCHAR
            iExprType = giEXPRVALUE_CHARACTER
        End Select
      Else
        MsgBox "No lookup column selected. Unable to select calculation as required data type cannot be determined.", vbInformation + vbOKOnly, App.ProductName
      End If
    
    Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
      iExprType = giEXPRVALUE_CHARACTER
    
    Case giWFFORMITEM_INPUTVALUE_NUMERIC
      iExprType = giEXPRVALUE_NUMERIC
  End Select
  
  If iExprType <> giEXPRVALUE_UNDEFINED Then
    ' Instantiate an expression object.
    Set objExpr = New CExpression
  
    With objExpr
      ' Set the properties of the expression object.
      .Initialise 0, mlngDefaultValueExprID, giEXPR_WORKFLOWCALCULATION, iExprType
      .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
      .UtilityBaseTable = mfrmCallingForm.BaseTable
      .WorkflowInitiationType = mfrmCallingForm.InitiationType
      .PrecedingWorkflowElements = maWFPrecedingElements
      .AllWorkflowElements = maWFAllElements
  
      ' Instruct the expression object to display the
      ' expression selection form.
      If .SelectExpression(mfReadOnly) Then
        mlngDefaultValueExprID = .ExpressionID
      Else
        ' Check in case the original expression has been deleted.
        If Not CheckExpression(mlngDefaultValueExprID, 0, False) Then
          mlngDefaultValueExprID = 0
        End If
      End If
  
      ' Read the selected expression info.
      txtDefaultValueExpression.Text = GetExpressionName(mlngDefaultValueExprID)
    End With
  
    Set objExpr = Nothing
  
    If lngOriginalID <> mlngDefaultValueExprID Then
      RefreshDefaultValueControls
      Changed = True
    End If
  End If
  
End Sub


Private Sub cmdDescriptionExpression_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long
  
  lngOriginalID = mlngDescriptionExprID
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    .Initialise 0, mlngDescriptionExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_CHARACTER
    .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
    .UtilityBaseTable = mfrmCallingForm.BaseTable
    .WorkflowInitiationType = mfrmCallingForm.InitiationType
    .PrecedingWorkflowElements = maWFPrecedingElements
    .AllWorkflowElements = maWFAllElements

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression(mfReadOnly) Then
      mlngDescriptionExprID = .ExpressionID
    Else
      ' Check in case the original expression has been deleted.
      If Not CheckExpression(mlngDescriptionExprID, 0, False) Then
        mlngDescriptionExprID = 0
      End If
    End If

    ' Read the selected expression info.
    txtDescriptionExpression.Text = GetExpressionName(mlngDescriptionExprID)
  End With
  
  Set objExpr = Nothing
  
  If lngOriginalID <> mlngDescriptionExprID Then
    RefreshIdentificationControls
    Changed = True
  End If
  
End Sub

Private Sub cmdFollowOnFormsMessage_Click()
  
  DefineRichTextMessage WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS

End Sub

Private Sub cmdForegroundColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColForeColor
'    .ShowColor
'
'    If mColForeColor <> .Color Then
'      mColForeColor = .Color
'      Changed = True
'
'      txtForegroundColour.BackColor = mColForeColor
'    End If
'  End With

  With colPickDlg
    .Color = mColForeColor
    .ShowPalette
    
    If mColForeColor <> .Color Then
      mColForeColor = .Color
      Changed = True
      
      txtForegroundColour.BackColor = mColForeColor
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub

Private Sub cmdForegroundEvenColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColForeColorEven
'    .ShowColor
'
'    If mColForeColorEven <> .Color Then
'      mColForeColorEven = .Color
'      Changed = True
'
'      txtForegroundEvenColour.BackColor = mColForeColorEven
'    End If
'  End With

  With colPickDlg
    .Color = mColForeColorEven
    .ShowPalette
    
    If mColForeColorEven <> .Color Then
      mColForeColorEven = .Color
      Changed = True
      
      txtForegroundEvenColour.BackColor = mColForeColorEven
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub

Private Sub cmdForegroundFont_Click()
  On Error GoTo ErrorTrap
  
  With comDlgBox
    .FontName = mObjFont.Name
    .FontSize = mObjFont.Size
    .FontBold = mObjFont.Bold
    .FontItalic = mObjFont.Italic
    .FontUnderline = mObjFont.Underline
    .FontStrikethru = mObjFont.Strikethrough
    .Flags = cdlCFScreenFonts Or cdlCFEffects
    .ShowFont
      
    If mObjFont.Name <> .FontName _
      Or mObjFont.Size <> .FontSize _
      Or mObjFont.Bold <> .FontBold _
      Or mObjFont.Italic <> .FontItalic _
      Or mObjFont.Underline <> .FontUnderline _
      Or mObjFont.Strikethrough <> .FontStrikethru Then
      
      mObjFont.Name = .FontName
      mObjFont.Size = .FontSize
      mObjFont.Bold = .FontBold
      mObjFont.Italic = .FontItalic
      mObjFont.Underline = .FontUnderline
      mObjFont.Strikethrough = .FontStrikethru
      
      Changed = True
      AutoResizeControl

      txtForegroundFont.Text = GetFontDescription(mObjFont)
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False
  
End Sub

Private Sub cmdForegroundHighlightColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColForeColorHighlight
'    .ShowColor
'
'    If mColForeColorHighlight <> .Color Then
'      mColForeColorHighlight = .Color
'      Changed = True
'
'      txtForegroundHighlightColour.BackColor = mColForeColorHighlight
'    End If
'  End With

  With colPickDlg
    .Color = mColForeColorHighlight
    .ShowPalette
    
    If mColForeColorHighlight <> .Color Then
      mColForeColorHighlight = .Color
      Changed = True
      
      txtForegroundHighlightColour.BackColor = mColForeColorHighlight
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub

Private Sub cmdForegroundOddColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColForeColorOdd
'    .ShowColor
'
'    If mColForeColorOdd <> .Color Then
'      mColForeColorOdd = .Color
'      Changed = True
'
'      txtForegroundOddColour.BackColor = mColForeColorOdd
'    End If
'  End With

  With colPickDlg
    .Color = mColForeColorOdd
    .ShowPalette
    
    If mColForeColorOdd <> .Color Then
      mColForeColorOdd = .Color
      Changed = True
      
      txtForegroundOddColour.BackColor = mColForeColorOdd
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub

Private Sub cmdHeaderBackgroundColour_Click()
  ' Display the Colour dialogue box.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'    .Flags = cdlCCRGBInit
'    .Color = mColHeaderBackColor
'    .ShowColor
'
'    If mColHeaderBackColor <> .Color Then
'      mColHeaderBackColor = .Color
'      Changed = True
'
'      txtHeaderBackgroundColour.BackColor = mColHeaderBackColor
'    End If
'  End With

  With colPickDlg
    .Color = mColHeaderBackColor
    .ShowPalette
    
    If mColHeaderBackColor <> .Color Then
      mColHeaderBackColor = .Color
      Changed = True
      
      txtHeaderBackgroundColour.BackColor = mColHeaderBackColor
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False
  
End Sub

Private Sub cmdHeaderFont_Click()
  On Error GoTo ErrorTrap
  
  With comDlgBox
    .FontName = mObjHeadFont.Name
    .FontSize = mObjHeadFont.Size
    .FontBold = mObjHeadFont.Bold
    .FontItalic = mObjHeadFont.Italic
    .FontUnderline = mObjHeadFont.Underline
    .FontStrikethru = mObjHeadFont.Strikethrough
    .Flags = cdlCFScreenFonts Or cdlCFEffects
    .ShowFont
      
    If mObjHeadFont.Name <> .FontName _
      Or mObjHeadFont.Size <> .FontSize _
      Or mObjHeadFont.Bold <> .FontBold _
      Or mObjHeadFont.Italic <> .FontItalic _
      Or mObjHeadFont.Underline <> .FontUnderline _
      Or mObjHeadFont.Strikethrough <> .FontStrikethru Then
      
      mObjHeadFont.Name = .FontName
      mObjHeadFont.Size = .FontSize
      mObjHeadFont.Bold = .FontBold
      mObjHeadFont.Italic = .FontItalic
      mObjHeadFont.Underline = .FontUnderline
      mObjHeadFont.Strikethrough = .FontStrikethru
      
      Changed = True

      txtHeaderFont.Text = GetFontDescription(mObjHeadFont)
    End If
  End With

ErrorTrap:
  ' Dialog was cancelled?
  Err = False

End Sub

Private Sub cmdOK_Click()
  Dim fOK As Boolean
  
  fOK = True
  
  If Changed Then
    fOK = ValidateProperties
    
    If fOK Then
      SaveProperties
      mfrmCallingForm.IsChanged = True
    End If
  End If
  
  If fOK Then
    ' Flag that the change/deletion has been confirmed.
    mfCancelled = False
  
    ' Unload the form.
    UnLoad Me
  End If

End Sub

Private Sub cmdPictureClear_Click()
  mlngPictureID = 0
  RefreshPictureControls

  Changed = True

End Sub

Private Sub cmdPictureSelect_Click()
  ' Display the Picture selection form.
  Dim lngOriginalID As Long
  
  lngOriginalID = mlngPictureID
  
  frmPictSel.SelectedPicture = mlngPictureID
  frmPictSel.ExcludedExtensions = ""
  frmPictSel.Show vbModal
  
  mlngPictureID = frmPictSel.SelectedPicture
  RefreshPictureControls

  If lngOriginalID <> mlngPictureID Then
    Changed = True
  End If

End Sub


Private Sub cmdRecordIdentificationFilter_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long
  Dim lngTableID As Long
  
  lngOriginalID = mlngRecordIdentificationFilterID
  
  If cboRecordIdentificationTable.ListCount > 0 Then
    lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
  End If
  
  If lngTableID > 0 Then
    ' Instantiate an expression object.
    Set objExpr = New CExpression
    
    With objExpr
      ' Set the properties of the expression object.
      .Initialise lngTableID, mlngRecordIdentificationFilterID, giEXPR_WORKFLOWRUNTIMEFILTER, giEXPRVALUE_LOGIC
      .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
      .UtilityBaseTable = mfrmCallingForm.BaseTable
      .WorkflowInitiationType = mfrmCallingForm.InitiationType
      .PrecedingWorkflowElements = maWFPrecedingElements
      .AllWorkflowElements = maWFAllElements
  
      ' Instruct the expression object to display the
      ' expression selection form.
      If .SelectExpression(mfReadOnly) Then
        mlngRecordIdentificationFilterID = .ExpressionID
      Else
        ' Check in case the original expression has been deleted.
        If Not CheckExpression(mlngRecordIdentificationFilterID, lngTableID, True) Then
          mlngRecordIdentificationFilterID = 0
        End If
      End If
  
      ' Read the selected expression info.
      txtRecordIdentificationFilter.Text = GetExpressionName(mlngRecordIdentificationFilterID)
    End With
    
    Set objExpr = Nothing
  Else
    mlngRecordIdentificationFilterID = 0
  End If
  
  If lngOriginalID <> mlngRecordIdentificationFilterID Then
    Changed = True
  End If
  
End Sub


Private Sub cmdRecordIdentificationOrder_Click()
  ' Display the Order selection form.
  Dim objOrder As Order
  Dim lngTableID As Long
  Dim lngOriginalID As Long
  
  lngOriginalID = mlngRecordIdentificationOrderID
  
  If cboRecordIdentificationTable.ListCount > 0 Then
    lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
  End If
  
  If lngTableID > 0 Then
    Set objOrder = New Order
    With objOrder
      .OrderID = mlngRecordIdentificationOrderID
      .TableID = lngTableID
      .OrderType = giORDERTYPE_DYNAMIC
      
      .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
      .AllWorkflowElements = maWFAllElements
  
      If .SelectOrder(mfReadOnly) Then
        mlngRecordIdentificationOrderID = .OrderID
      Else
        If Not CheckOrder(mlngRecordIdentificationOrderID, lngTableID) Then
          mlngRecordIdentificationOrderID = 0
        End If
      End If
    End With
  
    Set objOrder = Nothing
  Else
    mlngRecordIdentificationOrderID = 0
  End If
  
  txtRecordIdentificationOrder.Text = GetOrderName(mlngRecordIdentificationOrderID)
  
  If lngOriginalID <> mlngRecordIdentificationOrderID Then
    Changed = True
  End If
  
End Sub

Private Sub cmdLookupOrder_Click()
  ' Display the Order selection form.
  Dim objOrder As Order
  Dim lngTableID As Long
  Dim lngOriginalID As Long
  
  lngOriginalID = mlngLookupOrderID
  
  If cboLookupTable.ListCount > 0 Then
    lngTableID = cboLookupTable.ItemData(cboLookupTable.ListIndex)
  End If
  
  If lngTableID > 0 Then
    Set objOrder = New Order
    With objOrder
      .OrderID = mlngLookupOrderID
      .TableID = lngTableID
      .OrderType = giORDERTYPE_DYNAMIC
      
      .UtilityID = mfrmCallingForm.CallingForm.WorkflowID
      .AllWorkflowElements = maWFAllElements
  
      If .SelectOrder(mfReadOnly) Then
        mlngLookupOrderID = .OrderID
      Else
        If Not CheckOrder(mlngLookupOrderID, lngTableID) Then
          mlngLookupOrderID = 0
        End If
      End If
    End With
  
    Set objOrder = Nothing
  Else
    mlngLookupOrderID = 0
  End If
  
  txtLookupOrder.Text = GetOrderName(mlngLookupOrderID)
  
  If lngOriginalID <> mlngLookupOrderID Then
    Changed = True
  End If
  
End Sub

Private Sub cmdSavedForLaterMessage_Click()
  
  DefineRichTextMessage WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER

End Sub

Private Sub cmdValidationAdd_Click()
  DefineValidation True

End Sub

Private Sub cmdValidationDelete_Click()
  Dim lRow As Long

  If MsgBox("Delete this validation, are you sure ?", _
    vbQuestion + vbYesNo, _
    "Confirm Delete") = vbYes Then
    
    With grdValidation
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
  
    RefreshValidationControls
    Changed = True
  End If

End Sub

Private Sub cmdValidationDeleteAll_Click()
  If MsgBox("Remove all validation for this Web Form, are you sure ?", _
    vbQuestion + vbYesNo, _
    "Confirm Remove All") = vbYes Then
    
    grdValidation.RemoveAll

    RefreshValidationControls
    Changed = True
  End If

End Sub

Private Sub cmdValidationEdit_Click()
  DefineValidation False

End Sub

Private Sub dtDefaultValue_Change()
  Changed = True
  
End Sub

Private Sub dtDefaultValue_GotFocus()
  UI.txtSelText

End Sub

Private Sub dtDefaultValue_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    dtDefaultValue.DateValue = Date
  End If

End Sub

Private Sub dtDefaultValue_LostFocus()
  ValidateGTMaskDate dtDefaultValue

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
  Dim iLoop As Integer
  
  Const FORM_WIDTH = 9600
  Const GRIDROWHEIGHT = 239
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

  Me.Width = FORM_WIDTH

  grdValidation.RowHeight = GRIDROWHEIGHT

  UI.FormatGTDateControl dtDefaultValue

  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)

  With ssTabStrip
    .Width = Me.ScaleWidth - (2 * .Left)
    
    msngFrameWidth = .Width - (2 * XGAP_TAB_FRAME)
    
    For iLoop = 0 To .Tabs - 1
      With picTabContainer(iLoop)
        .BackColor = vbButtonFace
        .Top = YGAP_TAB_FRAME
        .Left = XGAP_TAB_FRAME
        .Width = msngFrameWidth
      End With
    Next iLoop
    
    fraButtons.Left = .Left _
      + .Width _
      - fraButtons.Width
  End With
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iAnswer As Integer
  
  If UnloadMode <> vbFormCode Then

    'Check if any changes have been made.
    If mfChanged Then
      iAnswer = MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If iAnswer = vbYes Then
        Call cmdOK_Click
        If Me.Cancelled Then Cancel = 1
      ElseIf iAnswer = vbNo Then
        Me.Cancelled = True
      ElseIf iAnswer = vbCancel Then
        Cancel = 1
      End If
    Else
      Me.Cancelled = True
    End If
  End If

End Sub

Public Property Let Changed(ByVal pfNewValue As Boolean)
  If Not mfLoading Then
    mfChanged = pfNewValue
    RefreshScreen
  End If
  
End Property

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mObjFont = Nothing
  Set mObjHeadFont = Nothing

End Sub

Private Sub grdValidation_Click()
  RefreshValidationControls

End Sub

Private Sub grdValidation_DblClick()
  If cmdValidationEdit.Enabled Then
    cmdValidationEdit_Click
  ElseIf cmdValidationAdd.Enabled Then
    cmdValidationAdd_Click
  End If

End Sub

Private Sub numDefaultValue_Change()
  Changed = True

End Sub

Private Sub optButtonAction_Click(Index As Integer)
  Changed = True

End Sub

Private Sub optCaptionType_Click(Index As Integer)
  Changed = True
  RefreshIdentificationControls
  
End Sub


Private Sub optDefaultValue_Click(Index As Integer)
  Changed = True

End Sub

Private Sub optDefaultValueType_Click(Index As Integer)
  Changed = True
  RefreshDefaultValueControls

End Sub


Private Sub optOrientation_Click(Index As Integer)

  Changed = True
  RefreshOrientationControls
        
End Sub

Private Sub spnDecimals_Change()
  RefreshDecimalsControls
  Changed = True

End Sub

Private Sub spnDecimals_Click()
  spnDecimals.SetFocus
  
End Sub


Private Sub spnDefaultValue_Change()
  Changed = True

End Sub

Private Sub spnDefaultValue_Click()
  spnDefaultValue.SetFocus

End Sub


Private Sub spnHeaderLines_Change()
  RefreshHeaderControls
  Changed = True

End Sub

Private Sub spnHeaderLines_Click()
  spnHeaderLines.SetFocus

End Sub


Private Sub spnHeight_Change()
  Changed = True

End Sub

Private Sub spnHeight_Click()
  spnHeight.SetFocus

End Sub

Private Sub spnHOffset_Change()
  'spnLeft.Value = (mfrmCallingForm.Height) + (spnTop.Value) + (spnHeight.Value)
  Changed = True
End Sub

Private Sub spnHOffset_Click()
    spnHOffset.SetFocus
End Sub

Private Sub spnLeft_Change()
  Changed = True

End Sub

Private Sub spnLeft_Click()
  spnLeft.SetFocus

End Sub


Private Sub spnSize_Change()
  RefreshSizeControls
  Changed = True

End Sub

Private Sub spnSize_Click()
  spnSize.SetFocus

End Sub


Private Sub spnTimeoutFrequency_Change()
  Changed = True

End Sub

Private Sub FormatScreen()
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim iFirstTab As Integer
  
  ' Format the screen display
  msngMaxFrameBottom = 0
  
  With ssTabStrip
    iFirstTab = -1
    ssTabStrip_Click 1
    
    For iLoop = 0 To .Tabs - 1
      If FormatScreen_Tab(iLoop) _
        And iFirstTab < 0 Then
        
        iFirstTab = iLoop
      End If
    Next iLoop
    
    For iLoop = 0 To .Tabs - 1
      picTabContainer(iLoop).Height = msngMaxFrameBottom
    Next iLoop
    
    If iFirstTab >= 0 Then
      .Tab = iFirstTab
    End If
    
    .Height = YGAP_TAB_FRAME _
      + msngMaxFrameBottom _
      + YGAP_FRAME_TAB
      
    fraButtons.Top = .Top _
      + .Height _
      + YGAP_TAB_BUTTONS
  End With

  With Me
    .Height = fraButtons.Top _
      + fraButtons.Height _
      + YGAP_BUTTONS_FORM
    
    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2
  End With
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub



Private Function FormatScreen_Frame_Display() As Boolean
  ' Format the DISPLAY frame on the GENERAL tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the Orientation controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_ORIENTATION) Then
    fFrameNeeded = True
    
    With lblOrientation
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With optOrientation(0)
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN2
      .Visible = True
    End With
    
    With optOrientation(1)
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = optOrientation(0).Left + optOrientation(0).Width + XGAP_CONTROL_CONTROL
      .Visible = True
    End With
  
    If mctlSelectedControl.Alignment = wfItemPropertyOrientation_Vertical Then
      optOrientation(1) = True
      optOrientation(0) = False
    Else
      optOrientation(0) = True
      optOrientation(1) = False
    End If

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblOrientation.Visible = False
    optOrientation(0).Visible = False
    optOrientation(1).Visible = False
  End If

  ' Format the Top controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_VERTICALOFFSET) Then
    fFrameNeeded = True
    
    With lblVOffset
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With spnVOffset
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
      
      .value = TwipsToPixels(mctlSelectedControl.VerticalOffset)
    End With
    
    With lblVOffsetFrom
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN3
      .Visible = True
    End With
    
    cboVOffsetBehaviour_Refresh mctlSelectedControl.VerticalOffsetBehaviour
    miVOffsetBehaviour = mctlSelectedControl.VerticalOffsetBehaviour
    
    With cboVOffsetBehaviour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN4 '+ (XGAP_CONTROL_CONTROL * 2)
      .Visible = True
    End With
    
    spnTop.value = TwipsToPixels(mctlSelectedControl.Top)
    
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_TOP) Then
    fFrameNeeded = True
    
    With lblTop
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With spnTop
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True

      If miItemType = giWFFORMITEM_FORM Then
        .value = TwipsToPixels(mfrmCallingForm.Top)
      Else
        .value = TwipsToPixels(mctlSelectedControl.Top)
      End If
    End With
  Else
    lblTop.Visible = False
    spnTop.Visible = False
  End If
  
  ' Format the Left controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HORIZONTALOFFSET) Then
    fFrameNeeded = True
    
    With lblHOffset
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With spnHOffset
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
    
      .value = TwipsToPixels(mctlSelectedControl.HorizontalOffset)
    End With
    
     With lblHOffsetFrom
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN3
      .Visible = True
    End With
    
    cboHOffsetBehaviour_Refresh mctlSelectedControl.HorizontalOffsetBehaviour
    miHOffsetBehaviour = mctlSelectedControl.HorizontalOffsetBehaviour
    
    With cboHOffsetBehaviour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN4 '+ (XGAP_CONTROL_CONTROL * 2)
      .Visible = True
    End With
    
    spnLeft.value = TwipsToPixels(mctlSelectedControl.Left)
    
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_LEFT) Then
    fFrameNeeded = True
    
    With lblLeft
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN3
      .Visible = True
    End With
    
    With spnLeft
      .Top = sngCurrentControlTop
      .Left = X_COLUMN4
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = TwipsToPixels(mfrmCallingForm.Left)
      Else
        .value = TwipsToPixels(mctlSelectedControl.Left)
      End If
    End With
  Else
    lblLeft.Visible = False
    spnLeft.Visible = False
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_TOP) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_LEFT) Then
    
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If
  
  ' Format the Height controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEIGHT) Then
    fFrameNeeded = True

    With lblHeight
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With spnHeight
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = TwipsToPixels(mfrmCallingForm.Height)
      Else
        .value = TwipsToPixels(mctlSelectedControl.Height)
      End If
      
      If WebFormItemHasProperty(miItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
        .Left = X_COLUMN4 '+ (XGAP_CONTROL_CONTROL * 2)
      End If
    End With
  
    If WebFormItemHasProperty(miItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
      With lblHeightValue
        .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
        .Left = X_COLUMN3
        .Visible = True
      End With
      
      cboHeightBehaviour_Refresh mctlSelectedControl.HeightBehaviour
      miHeightBehaviour = mctlSelectedControl.HeightBehaviour

      With cboHeightBehaviour
        .Top = sngCurrentControlTop
        .Left = X_COLUMN2
        .Visible = True
      End With

      If cboHeightBehaviour.ItemData(cboHeightBehaviour.ListIndex) = behaveFull Then
        EnableControl lblHeightValue, False
        EnableControl spnHeight, False
      End If
      
      sngCurrentControlTop = sngCurrentControlTop _
        + YGAP_CONTROL_CONTROL
    End If
    
    If (miItemType = giWFFORMITEM_INPUTVALUE_DATE) _
      Or (miItemType = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
      Or (miItemType = giWFFORMITEM_INPUTVALUE_LOOKUP) _
      Or (miItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
      Or (miItemType = giWFFORMITEM_LINE) Then
      
      EnableControl lblHeight, False
      EnableControl spnHeight, False
      
      ' AE20080311 Fault #12989
      'If miItemType = giWFFORMITEM_LINE And mctlSelectedControl.Alignment = wfItemPropertyOrientation_Vertical Then
      If miItemType = giWFFORMITEM_LINE And WebFormItemHasProperty(miItemType, WFITEMPROP_ORIENTATION) Then
        If mctlSelectedControl.Alignment = wfItemPropertyOrientation_Vertical Then
          EnableControl lblHeight, True
          EnableControl spnHeight, True
        End If
      End If
    End If
  Else
    lblHeight.Visible = False
    spnHeight.Visible = False
  End If

  ' Format the Width controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTH) Then
    fFrameNeeded = True

    With lblWidth
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN3
      .Visible = True
      
      If WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
        .Left = X_COLUMN1
      End If
    End With

    With spnWidth
      .Top = sngCurrentControlTop
      .Left = X_COLUMN4
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = TwipsToPixels(mfrmCallingForm.Width)
      Else
        .value = TwipsToPixels(mctlSelectedControl.Width)
      End If
      
'      If WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
'        .Left = X_COLUMN3 + (XGAP_CONTROL_CONTROL * 2)
'      End If
    End With
    
    If WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
      With lblWidthValue
        .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
        .Left = X_COLUMN3
        .Visible = True
      End With
      
      cboWidthBehaviour_Refresh mctlSelectedControl.WidthBehaviour
      miWidthBehaviour = mctlSelectedControl.WidthBehaviour
      
      With cboWidthBehaviour
        .Top = sngCurrentControlTop
        .Left = X_COLUMN2
        .Visible = True
      End With
      
      If cboWidthBehaviour.ItemData(cboWidthBehaviour.ListIndex) = behaveFull Then
        EnableControl lblWidthValue, False
        EnableControl spnWidth, False
      End If
    End If
    
    Select Case miItemType
      Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
        EnableControl lblWidth, False
        EnableControl spnWidth, False
      Case giWFFORMITEM_LINE
        If mctlSelectedControl.Alignment = wfItemPropertyOrientation_Vertical Then
          EnableControl lblWidth, False
          EnableControl spnWidth, False
        End If
    End Select
  Else
    lblWidth.Visible = False
    spnWidth.Visible = False
  End If
  
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEIGHT) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_WIDTH) Then
    
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If
  
  ' Format the frame
  With fraDisplay
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Display = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Err.Clear
  Resume TidyUpAndExit
  
End Function




Private Function FormatScreen_Frame_Size() As Boolean
  ' Format the SIZE frame on the DATA tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the Size controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_SIZE) Then
    fFrameNeeded = True

    With lblSize
      If miItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD Then
        .Caption = "Max. Size (Kb) :"
      Else
        .Caption = "Size :"
      End If
      
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With spnSize
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = 0
      Else
        If miItemType = giWFFORMITEM_INPUTVALUE_NUMERIC Then
          .MaximumValue = WORKFLOWWEBFORM_MAXSIZE_NUMINPUT
        ElseIf miItemType = giWFFORMITEM_INPUTVALUE_FILEUPLOAD Then
          .MinimumValue = WORKFLOWWEBFORM_MINSIZE_FILEUPLOAD
          .MaximumValue = WORKFLOWWEBFORM_MAXSIZE_FILEUPLOAD
        Else
          .MaximumValue = WORKFLOWWEBFORM_MAXSIZE_CHARINPUT
        End If
        
        .value = mctlSelectedControl.WFInputSize
        RefreshSizeControls
      End If
    End With
  Else
    lblSize.Visible = False
    spnSize.Visible = False
  End If

  ' Format the Decimals controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DECIMALS) Then
    fFrameNeeded = True

    With lblDecimals
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN3
      .Visible = True
    End With

    With spnDecimals
      .Top = sngCurrentControlTop
      .Left = X_COLUMN4
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = 0 ' Not available for forms
      Else
        .value = mctlSelectedControl.WFInputDecimals
        RefreshDecimalsControls
      End If
    End With
  Else
    lblDecimals.Visible = False
    spnDecimals.Visible = False
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_SIZE) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_DECIMALS) Then

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If
  
  ' Format the frame
  With fraSize
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Size = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function





Private Function FormatScreen_Frame_Validation() As Boolean
  ' Format the VALIDATION frame on the DATA tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim sngGapBetweenButtons As Single
  Dim asValidations() As String
  Dim lngLoop As Long
  Dim sTemp As String
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the ValidationExpressions controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_VALIDATION) Then
    fFrameNeeded = True

    With grdValidation
      .Top = sngCurrentControlTop
      .Left = X_COLUMN1
      .Height = Y_GRIDCONTROLHEIGHT
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      sngCurrentControlTop = .Top _
        + .Height _
        - Y_STANDARDCONTROLHEIGHT _
        + YGAP_CONTROL_CONTROL
        
      If miItemType = giWFFORMITEM_FORM Then
        ' Only available for forms
        asValidations = mfrmCallingForm.Validations
        
        .RemoveAll
        For lngLoop = 1 To UBound(asValidations, 2)
          sTemp = asValidations(1, lngLoop) _
            & vbTab & GetExpressionName(CLng(asValidations(1, lngLoop))) _
            & vbTab & asValidations(2, lngLoop) _
            & vbTab & WorkflowWebFormValidationTypeDescription(CInt(asValidations(2, lngLoop))) _
            & vbTab & asValidations(3, lngLoop)
          .AddItem sTemp
        Next lngLoop
      End If
        
    End With

    sngGapBetweenButtons = (msngFrameWidth - cmdValidationAdd.Width - cmdValidationEdit.Width - cmdValidationDelete.Width - cmdValidationDeleteAll.Width) / 3

    With cmdValidationAdd
      .Top = sngCurrentControlTop
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cmdValidationEdit
      .Top = sngCurrentControlTop
      .Left = cmdValidationAdd.Left _
        + cmdValidationAdd.Width _
        + sngGapBetweenButtons
      .Visible = True
    End With

    With cmdValidationDelete
      .Top = sngCurrentControlTop
      .Left = cmdValidationEdit.Left _
        + cmdValidationEdit.Width _
        + sngGapBetweenButtons
      .Visible = True
    End With

    With cmdValidationDeleteAll
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = cmdValidationDeleteAll.Top _
      + cmdValidationDeleteAll.Height _
      - Y_STANDARDCONTROLHEIGHT _
      + YGAP_CONTROL_CONTROL
  Else
    grdValidation.Visible = False
    cmdValidationAdd.Visible = False
    cmdValidationEdit.Visible = False
    cmdValidationDelete.Visible = False
    cmdValidationDeleteAll.Visible = False
  End If

  ' Format the Mandatory controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_MANDATORY) Then
    fFrameNeeded = True

    With chkValidationMandatory
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = vbUnchecked ' Not available for forms
      Else
        .value = IIf(mctlSelectedControl.Mandatory, vbChecked, vbUnchecked)
      End If
    End With
  
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    chkValidationMandatory.Visible = False
  End If

  ' Format the File Extensions controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FILEEXTENSIONS) Then
    fFrameNeeded = True

    With lblFileExtensions
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    'NPG20090303 Fault 13578
    With lblFileExtensionNote
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL + lblFileExtensions.Height - Y_STANDARDCONTROLHEIGHT + YGAP_CONTROL_CONTROL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtFileExtensions
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Height = Y_GRIDCONTROLHEIGHT
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True

      .Text = SplitControlValues(mctlSelectedControl.WFFileExtensions)

      sngCurrentControlTop = .Top _
        + .Height _
        - Y_STANDARDCONTROLHEIGHT _
        + YGAP_CONTROL_CONTROL
    End With
        
  Else
    lblFileExtensions.Visible = False
    lblFileExtensionNote.Visible = False
    txtFileExtensions.Visible = False
  End If

  ' Format the frame
  With fraValidation
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth

      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME

      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If

    .Visible = fFrameNeeded
        
    RefreshValidationControls
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Validation = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function






Private Function FormatScreen_Frame_ControlValues() As Boolean
  ' Format the CONTROLVALUES frame on the DATA tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the Control Values controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_CONTROLVALUELIST) Then
    fFrameNeeded = True
    
    With txtControlValues
      .Top = sngCurrentControlTop
      .Left = X_COLUMN1
      .Height = Y_GRIDCONTROLHEIGHT
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
  
      If miItemType = giWFFORMITEM_FORM Then
        .Text = "" ' Not available for forms
      Else
        If miItemType = giWFFORMITEM_INPUTVALUE_OPTIONGROUP Then
          If mctlSelectedControl.NoOptions Then
            .Text = ""
          Else
            .Text = SplitControlValues(mctlSelectedControl.ControlValueList)
          End If
        Else
          .Text = SplitControlValues(mctlSelectedControl.ControlValueList)
        End If
      End If

      sngCurrentControlTop = .Top _
        + .Height _
        - Y_STANDARDCONTROLHEIGHT _
        + YGAP_CONTROL_CONTROL
    End With
  Else
    txtControlValues.Visible = False
  End If

  ' Format the frame
  With fraControlValues
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth

      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME

      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If

    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_ControlValues = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function


Private Function FormatScreen_Frame_HotSpot() As Boolean
  ' Format the HOTSPOT frame on the GENERAL tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the HotsSpot controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HOTSPOT) Then
    fFrameNeeded = True

    With lblHotSpotIdentifier
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboHotSpotIdentifier
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
      'TODO
        cboHotSpotIdentifier_refresh mctlSelectedControl.HotSpotIdentifier
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblHotSpotIdentifier.Visible = False
    cboHotSpotIdentifier.Visible = False
  End If


  ' Format the frame
  With fraHotspot
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth

      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME

      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If

    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_HotSpot = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function




Private Function FormatScreen_Frame_Options() As Boolean
  ' Format the OPTIONS frame on the APPEARANCE tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim iLastColumnUsed As Integer
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  iLastColumnUsed = 0
  
  ' Format the Border controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BORDERSTYLE) Then
    fFrameNeeded = True
    
    With chkBorder
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = vbUnchecked ' Not available for form items
      Else
        .value = IIf(mctlSelectedControl.BorderStyle = vbBSNone, vbUnchecked, vbChecked)
      End If
    End With

    iLastColumnUsed = 1
  Else
    chkBorder.Visible = False
  End If
  
  ' Format the Alignment controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_ALIGNMENT) Then
    fFrameNeeded = True
    
    With lblAlignment
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = IIf(iLastColumnUsed = 1, X_COLUMN3, X_COLUMN1)
      .Visible = True
    End With
    
    With cboAlignment
      .Top = sngCurrentControlTop
      .Left = IIf(iLastColumnUsed = 1, X_COLUMN4, X_COLUMN2)
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboAlignment_refresh mctlSelectedControl.Alignment
      End If
    End With
  
    iLastColumnUsed = iLastColumnUsed + 1
  Else
    lblAlignment.Visible = False
    cboAlignment.Visible = False
  End If
  
  ' Format the Password Type (Hide Text) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_PASSWORDTYPE) Then
    fFrameNeeded = True

    If iLastColumnUsed = 2 Then
      sngCurrentControlTop = sngCurrentControlTop _
        + YGAP_CONTROL_CONTROL
    End If
    
    With chkPasswordType
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = IIf(iLastColumnUsed = 1, X_COLUMN3, X_COLUMN1)
      .Visible = True

      If miItemType = giWFFORMITEM_FORM Then
        .value = vbUnchecked ' Not available for form items
      Else
        .value = IIf(mctlSelectedControl.PasswordType = True, vbChecked, vbUnchecked)
      End If
    End With
  
    iLastColumnUsed = IIf(iLastColumnUsed = 2, 1, iLastColumnUsed + 1)
  Else
    chkPasswordType.Visible = False
  End If
  
  ' Format the frame
  With fraOptions
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_CONTROL + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Options = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function





Private Function FormatScreen_Frame_Header() As Boolean
  ' Format the HEADER frame on the APPEARANCE tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim fColumn1Used As Boolean
  Dim objCtlFont As StdFont
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  fColumn1Used = False
  
  ' Format the ColumnHeaders controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_COLUMNHEADERS) Then
    fFrameNeeded = True

    With chkColumnHeaders
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = vbUnchecked ' Not available for forms
      Else
        .value = IIf(mctlSelectedControl.ColumnHeaders, vbChecked, vbUnchecked)
      End If
    End With

    fColumn1Used = True
  Else
    chkColumnHeaders.Visible = False
  End If

  ' Format the HeaderLines controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADLINES) Then
    fFrameNeeded = True

    With lblHeaderLines
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = IIf(fColumn1Used, X_COLUMN3, X_COLUMN1)
      .Visible = True
    End With

    With spnHeaderLines
      .Top = sngCurrentControlTop
      .Left = IIf(fColumn1Used, X_COLUMN4, X_COLUMN2)
      .Visible = True

      If miItemType = giWFFORMITEM_FORM Then
        .value = 0 ' Not available for forms
      Else
        .value = mctlSelectedControl.HeadLines
      End If
    End With
  Else
    lblHeaderLines.Visible = False
    spnHeaderLines.Visible = False
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADLINES) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_COLUMNHEADERS) Then
    
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If
  
  ' Format the HeaderFont controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADFONT) Then
    fFrameNeeded = True

    With lblHeaderFont
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtHeaderFont
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdHeaderFont.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        Set objCtlFont = mfrmCallingForm.Font ' Not available for forms
      Else
        Set objCtlFont = mctlSelectedControl.HeadFont
      End If
      
      If mObjHeadFont Is Nothing Then
        Set mObjHeadFont = New StdFont
      End If
      mObjHeadFont.Name = objCtlFont.Name
      mObjHeadFont.Size = objCtlFont.Size
      mObjHeadFont.Bold = objCtlFont.Bold
      mObjHeadFont.Italic = objCtlFont.Italic
      mObjHeadFont.Strikethrough = objCtlFont.Strikethrough
      mObjHeadFont.Underline = objCtlFont.Underline
      Set objCtlFont = Nothing
    
      .Text = GetFontDescription(mObjHeadFont)
    End With

    With cmdHeaderFont
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With
  
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblHeaderFont.Visible = False
    txtHeaderFont.Visible = False
    cmdHeaderFont.Visible = False
  End If

  ' Format the HeaderBackgroundColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_HEADERBACKCOLOR) Then
    fFrameNeeded = True

    With lblHeaderBackgroundColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtHeaderBackgroundColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdHeaderBackgroundColour.Width
      .Visible = True

      If miItemType = giWFFORMITEM_FORM Then
        mColHeaderBackColor = vbWhite ' Not available for form items
      Else
        mColHeaderBackColor = mctlSelectedControl.HeaderBackColor
      End If
      .BackColor = mColHeaderBackColor
    End With

    With cmdHeaderBackgroundColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With
  
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblHeaderBackgroundColour.Visible = False
    txtHeaderBackgroundColour.Visible = False
    cmdHeaderBackgroundColour.Visible = False
  End If
  
  ' Format the frame
  With fraHeader
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    
      RefreshHeaderControls
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Header = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function






Private Function FormatScreen_Frame_Foreground() As Boolean
  ' Format the FOREGROUND frame on the APPEARANCE tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim objCtlFont As StdFont
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the ForegroundFont controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FONT) Then
    fFrameNeeded = True

    With lblForegroundFont
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtForegroundFont
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdForegroundFont.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        Set objCtlFont = mfrmCallingForm.Font
      Else
        Set objCtlFont = mctlSelectedControl.Font
      End If
      
      If mObjFont Is Nothing Then
        Set mObjFont = New StdFont
      End If
      mObjFont.Name = objCtlFont.Name
      mObjFont.Size = objCtlFont.Size
      mObjFont.Bold = objCtlFont.Bold
      mObjFont.Italic = objCtlFont.Italic
      mObjFont.Strikethrough = objCtlFont.Strikethrough
      mObjFont.Underline = objCtlFont.Underline
      Set objCtlFont = Nothing
    
      .Text = GetFontDescription(mObjFont)
    End With

    With cmdForegroundFont
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblForegroundFont.Visible = False
    txtForegroundFont.Visible = False
    cmdForegroundFont.Visible = False
  End If

  ' Format the ForegroundColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLOR) Then
    fFrameNeeded = True

    With lblForegroundColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtForegroundColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdForegroundColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColForeColor = mfrmCallingForm.ForeColor
      Else
        mColForeColor = mctlSelectedControl.ForeColor
      End If
      .BackColor = mColForeColor
    End With

    With cmdForegroundColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblForegroundColour.Visible = False
    txtForegroundColour.Visible = False
    cmdForegroundColour.Visible = False
  End If
  
  ' Format the ForegroundEvenColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLOREVEN) Then
    fFrameNeeded = True

    With lblForegroundEvenColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtForegroundEvenColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdForegroundEvenColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColForeColorEven = vbWhite ' Not available for forms
      Else
        mColForeColorEven = mctlSelectedControl.ForeColorEven
      End If
      .BackColor = mColForeColorEven
    End With

    With cmdForegroundEvenColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblForegroundEvenColour.Visible = False
    txtForegroundEvenColour.Visible = False
    cmdForegroundEvenColour.Visible = False
  End If
  
  ' Format the ForegroundOddColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLORODD) Then
    fFrameNeeded = True

    With lblForegroundOddColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtForegroundOddColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdForegroundOddColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColForeColorOdd = vbWhite ' Not available for forms
      Else
        mColForeColorOdd = mctlSelectedControl.ForeColorOdd
      End If
      .BackColor = mColForeColorOdd
    End With

    With cmdForegroundOddColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblForegroundOddColour.Visible = False
    txtForegroundOddColour.Visible = False
    cmdForegroundOddColour.Visible = False
  End If
  
  ' Format the ForegroundHighlightColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
    fFrameNeeded = True

    With lblForegroundHighlightColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtForegroundHighlightColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdForegroundHighlightColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColForeColorHighlight = vbWhite ' Not available for forms
      Else
        mColForeColorHighlight = mctlSelectedControl.ForeColorHighlight
      End If
      .BackColor = mColForeColorHighlight
    End With

    With cmdForegroundHighlightColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblForegroundHighlightColour.Visible = False
    txtForegroundHighlightColour.Visible = False
    cmdForegroundHighlightColour.Visible = False
  End If
  
  ' Format the frame
  With fraForeground
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Foreground = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function







Private Function FormatScreen_Frame_Background() As Boolean
  ' Format the BACKGROUND frame on the APPEARANCE tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim fPictureUsed As Boolean
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  fPictureUsed = False
  
  ' Format the BackgroundStyle controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKSTYLE) Then
    fFrameNeeded = True

    With lblBackgroundStyle
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboBackgroundStyle
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboBackgroundStyle_refresh mctlSelectedControl.BackStyle
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblBackgroundStyle.Visible = False
    cboBackgroundStyle.Visible = False
  End If

  ' Format the BackgroundColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLOR) Then
    fFrameNeeded = True

    With lblBackgroundColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtBackgroundColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdBackgroundColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColBackColor = mfrmCallingForm.BackColor
      Else
        mColBackColor = mctlSelectedControl.BackColor
      End If
      .BackColor = mColBackColor
    End With

    With cmdBackgroundColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblBackgroundColour.Visible = False
    txtBackgroundColour.Visible = False
    cmdBackgroundColour.Visible = False
  End If

  ' Format the BackgroundEvenColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLOREVEN) Then
    fFrameNeeded = True
    
    With lblBackgroundEvenColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtBackgroundEvenColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdBackgroundEvenColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColBackColorEven = vbWhite ' Not available for forms
      Else
        mColBackColorEven = mctlSelectedControl.BackColorEven
      End If
      .BackColor = mColBackColorEven
    End With

    With cmdBackgroundEvenColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblBackgroundEvenColour.Visible = False
    txtBackgroundEvenColour.Visible = False
    cmdBackgroundEvenColour.Visible = False
  End If

  ' Format the BackgroundOddColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLORODD) Then
    fFrameNeeded = True

    With lblBackgroundOddColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtBackgroundOddColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdBackgroundOddColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColBackColorOdd = vbWhite ' Not available for forms
      Else
        mColBackColorOdd = mctlSelectedControl.BackColorOdd
      End If
      .BackColor = mColBackColorOdd
    End With

    With cmdBackgroundOddColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblBackgroundOddColour.Visible = False
    txtBackgroundOddColour.Visible = False
    cmdBackgroundOddColour.Visible = False
  End If

  ' Format the BackgroundHighlightColour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
    fFrameNeeded = True

    With lblBackgroundHighlightColour
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtBackgroundHighlightColour
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdBackgroundHighlightColour.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mColBackColorHighlight = vbWhite ' Not available for forms
      Else
        mColBackColorHighlight = mctlSelectedControl.BackColorHighlight
      End If
      .BackColor = mColBackColorHighlight
    End With

    With cmdBackgroundHighlightColour
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblBackgroundHighlightColour.Visible = False
    txtBackgroundHighlightColour.Visible = False
    cmdBackgroundHighlightColour.Visible = False
  End If
  
  ' Format the Picture controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_PICTURE) Then
    fFrameNeeded = True
    fPictureUsed = True

    With lblPicture
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtPicture
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - picPictureHolder.Width _
        - XGAP_CONTROL_CONTROL _
        - cmdPictureClear.Width _
        - cmdPictureSelect.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        mlngPictureID = mfrmCallingForm.PictureID
      Else
        mlngPictureID = mctlSelectedControl.PictureID
      End If
      RefreshPictureControls
    End With

    With cmdPictureSelect
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - X_COLUMN1 _
        - picPictureHolder.Width _
        - XGAP_CONTROL_CONTROL _
        - cmdPictureClear.Width _
        - .Width
      .Visible = True
    End With

    With cmdPictureClear
      .Font.Size = 8
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - X_COLUMN1 _
        - picPictureHolder.Width _
        - XGAP_CONTROL_CONTROL _
        - .Width
      .Font.Size = 20
      .Visible = True
    End With

    With picPictureHolder
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblPicture.Visible = False
    txtPicture.Visible = False
    cmdPictureSelect.Visible = False
    cmdPictureClear.Visible = False
    picPictureHolder.Visible = False
  End If
  
  ' Format the Location controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_PICTURELOCATION) Then
    fFrameNeeded = True

    With lblPictureLocation
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboPictureLocation
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = IIf(fPictureUsed, _
        msngFrameWidth _
          - .Left _
          - X_COLUMN1 _
          - picPictureHolder.Width _
          - XGAP_CONTROL_CONTROL, _
        msngFrameWidth _
          - .Left _
          - X_COLUMN1)
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        cboPictureLocation_refresh mfrmCallingForm.PictureLocation
      End If
    End With
  
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblPictureLocation.Visible = False
    cboPictureLocation.Visible = False
  End If
  
  If fPictureUsed Then
    If sngCurrentControlTop < (picPictureHolder.Top + picPictureHolder.Height - txtPicture.Height + YGAP_CONTROL_CONTROL) Then
      sngCurrentControlTop = picPictureHolder.Top _
        + picPictureHolder.Height _
        - txtPicture.Height _
        + YGAP_CONTROL_CONTROL
    End If
  End If
  
  ' Format the frame
  With fraBackground
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    
      RefreshBackgroundControls
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Background = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Private Function FormatScreen_Frame_Behaviour() As Boolean
  ' Format the BEHAVIOUR frame on the GENERAL tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the Behaviour controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_TIMEOUT) Then
    fFrameNeeded = True
    
    With lblTimeoutPeriod
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With spnTimeoutFrequency
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = mfrmCallingForm.TimeoutFrequency
      Else
        .value = 0 ' Not available for non-form items
      End If
    End With
    
    With cboTimeoutPeriod
      .Top = sngCurrentControlTop
      .Left = spnTimeoutFrequency.Left + spnTimeoutFrequency.Width + XGAP_CONTROL_CONTROL
      .Visible = True
      
      If miItemType = giWFFORMITEM_FORM Then
        cboTimeoutPeriod_refresh mfrmCallingForm.TimeoutPeriod
      End If
    End With

    With chkExcludeWeekends
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = cboTimeoutPeriod.Left + cboTimeoutPeriod.Width + XGAP_CONTROL_CONTROL
      .Visible = True
      
      If miItemType = giWFFORMITEM_FORM Then
        .value = IIf(mfrmCallingForm.TimeoutExcludeWeekend, vbChecked, vbUnchecked)
      Else
        .value = vbUnchecked
      End If
    End With
    
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblTimeoutPeriod.Visible = False
    spnTimeoutFrequency.Visible = False
    cboTimeoutPeriod.Visible = False
    chkExcludeWeekends.Visible = False
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_SUBMITTYPE) Then
    fFrameNeeded = True

    With lblButtonAction
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With optButtonAction(WORKFLOWBUTTONACTION_SUBMIT)
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN2
      .Visible = True
    End With

    With optButtonAction(WORKFLOWBUTTONACTION_CANCEL)
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = optButtonAction(WORKFLOWBUTTONACTION_SUBMIT).Left + optButtonAction(WORKFLOWBUTTONACTION_SUBMIT).Width + XGAP_CONTROL_CONTROL
      .Visible = True
    End With

    With optButtonAction(WORKFLOWBUTTONACTION_SAVEFORLATER)
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = optButtonAction(WORKFLOWBUTTONACTION_CANCEL).Left + optButtonAction(WORKFLOWBUTTONACTION_CANCEL).Width + XGAP_CONTROL_CONTROL
      .Visible = True
    End With

    optButtonAction(WORKFLOWBUTTONACTION_SUBMIT).value = (mctlSelectedControl.Behaviour = WORKFLOWBUTTONACTION_SUBMIT)
    optButtonAction(WORKFLOWBUTTONACTION_SAVEFORLATER).value = (mctlSelectedControl.Behaviour = WORKFLOWBUTTONACTION_SAVEFORLATER)
    optButtonAction(WORKFLOWBUTTONACTION_CANCEL).value = (mctlSelectedControl.Behaviour = WORKFLOWBUTTONACTION_CANCEL)

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblButtonAction.Visible = False
    optButtonAction(WORKFLOWBUTTONACTION_SUBMIT).Visible = False
    optButtonAction(WORKFLOWBUTTONACTION_SAVEFORLATER).Visible = False
    optButtonAction(WORKFLOWBUTTONACTION_CANCEL).Visible = False
  End If


  If WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGETYPE) Then
    fFrameNeeded = True

    With lblCompletionMessage
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboCompletionMessageType
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2POINT5
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - IIf(WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGE), cmdCompletionMessage.Width, 0)
      .Visible = True
    End With

    If WebFormItemHasProperty(miItemType, WFITEMPROP_COMPLETIONMESSAGE) Then
      With cmdCompletionMessage
        .Top = sngCurrentControlTop
        .Left = cboCompletionMessageType.Left + cboCompletionMessageType.Width
        .Visible = True
      End With
    End If
    
    If miItemType = giWFFORMITEM_FORM Then
      msCompletionMessage = mfrmCallingForm.WFCompletionMessage
      cboMessage_refresh WORKFLOWWEBFORMMESSAGE_COMPLETION, mfrmCallingForm.WFCompletionMessageType
    Else
      ' Not available for non-form items.
      msCompletionMessage = ""
    End If

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGETYPE) Then
    fFrameNeeded = True

    With lblSavedForLaterMessage
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboSavedForLaterMessageType
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2POINT5
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - IIf(WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGE), cmdSavedForLaterMessage.Width, 0)
      .Visible = True
    End With

    If WebFormItemHasProperty(miItemType, WFITEMPROP_SAVEDFORLATERMESSAGE) Then
      With cmdSavedForLaterMessage
        .Top = sngCurrentControlTop
        .Left = cboSavedForLaterMessageType.Left + cboSavedForLaterMessageType.Width
        .Visible = True
      End With
    End If
    
    If miItemType = giWFFORMITEM_FORM Then
      msSavedForLaterMessage = mfrmCallingForm.WFSavedForLaterMessage
      cboMessage_refresh WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER, mfrmCallingForm.WFSavedForLaterMessageType
    Else
      ' Not available for non-form items.
      msSavedForLaterMessage = ""
    End If

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If

  If WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGETYPE) Then
    fFrameNeeded = True

    With lblFollowOnFormsMessage
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboFollowOnFormsMessageType
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2POINT5
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - IIf(WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGE), cmdFollowOnFormsMessage.Width, 0)
      .Visible = True
    End With

    If WebFormItemHasProperty(miItemType, WFITEMPROP_FOLLOWONFORMSMESSAGE) Then
      With cmdFollowOnFormsMessage
        .Top = sngCurrentControlTop
        .Left = cboFollowOnFormsMessageType.Left + cboFollowOnFormsMessageType.Width
        .Visible = True
      End With
    End If
    
    If miItemType = giWFFORMITEM_FORM Then
      msFollowOnFormsMessage = mfrmCallingForm.WFFollowOnFormsMessage
      cboMessage_refresh WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS, mfrmCallingForm.WFFollowOnFormsMessageType
    Else
      ' Not available for non-form items.
      msFollowOnFormsMessage = ""
    End If

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  End If
    
    
  If WebFormItemHasProperty(miItemType, WFITEMPROP_REQUIRESAUTHENTICATION) Then
  
    With chkRequireAuthentication
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
      
      If miItemType = giWFFORMITEM_FORM Then
        .value = IIf(mfrmCallingForm.RequiresAuthentication, vbChecked, vbUnchecked)
      Else
        .value = vbUnchecked
      End If
    End With
    
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  
  End If
  
  ' Format the frame
  With fraBehaviour
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
    
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
    
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Behaviour = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Private Function FormatScreen_Frame_Identification() As Boolean
  ' Format the IDENTIFICATION frame on the GENERAL tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim sngSubCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the Identifier controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_WFIDENTIFIER) Then
    fFrameNeeded = True
    
    With lblIdentifier
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With txtIdentifier
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
      
      If miItemType = giWFFORMITEM_FORM Then
        .Text = mfrmCallingForm.WFIdentifier
      Else
        .Text = mctlSelectedControl.WFIdentifier
      End If
    End With
  
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblIdentifier.Visible = False
    txtIdentifier.Visible = False
  End If

  ' Format the Target Identifier if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_USEASTARGETIDENTIFIER) Then
    fFrameNeeded = True
    
    With chkUseAsTargetIdentifier
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
      .value = IIf(mctlSelectedControl.UseAsTargetIdentifier, vbChecked, vbUnchecked)
    End With
    
    sngCurrentControlTop = sngCurrentControlTop + YGAP_CONTROL_CONTROL
    
  End If

  ' Format the Caption controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTION) Then
    fFrameNeeded = True
    
    If WebFormItemHasProperty(miItemType, WFITEMPROP_CAPTIONTYPE) Then
      With fraCaption
        .Top = sngCurrentControlTop
        .Left = X_COLUMN1
      End With

      sngSubCurrentControlTop = YGAP_FRAME_CONTROL

      With optCaptionType(giWFDATAVALUE_FIXED)
        .Top = sngSubCurrentControlTop + YGAP_CONTROL_LABEL
        .Left = X_COLUMN1
        .Visible = True
      
        If miItemType = giWFFORMITEM_FORM Then
          .value = False
        Else
          .value = (mctlSelectedControl.CaptionType = giWFDATAVALUE_FIXED)
        End If
      End With
      
      With txtCaptionTypeValue
        .Top = sngSubCurrentControlTop
        .Left = X_COLUMN2 - X_COLUMN1
        .Width = msngFrameWidth _
          - .Left _
          - (3 * X_COLUMN1)
        .Visible = True
      
        If miItemType = giWFFORMITEM_FORM Then
          .Text = ""
        Else
          If (mctlSelectedControl.CaptionType = giWFDATAVALUE_FIXED) Then
            .Text = Replace(mctlSelectedControl.Caption, "&&", "&")
          Else
            .Text = ""
          End If
        End If
      End With
      
      sngSubCurrentControlTop = sngSubCurrentControlTop _
        + YGAP_CONTROL_CONTROL
        
      If WebFormItemHasProperty(miItemType, WFITEMPROP_CALCULATION) Then
        With optCaptionType(giWFDATAVALUE_CALC)
          .Top = sngSubCurrentControlTop + YGAP_CONTROL_LABEL
          .Left = X_COLUMN1
          .Visible = True
        
          If miItemType = giWFFORMITEM_FORM Then
            .value = False
          Else
            .value = (mctlSelectedControl.CaptionType = giWFDATAVALUE_CALC)
          End If
        End With
      
        With txtCaptionTypeExpression
          .Top = sngSubCurrentControlTop
          .Left = X_COLUMN2 - X_COLUMN1
          .Width = msngFrameWidth _
            - .Left _
            - (3 * X_COLUMN1) _
            - cmdCaptionTypeExpression.Width
          .Visible = True
        End With
      
        With cmdCaptionTypeExpression
          .Top = sngSubCurrentControlTop
          .Left = txtCaptionTypeExpression.Left + txtCaptionTypeExpression.Width
          .Visible = True
        End With
              
        If miItemType = giWFFORMITEM_FORM Then
          mlngCaptionExprID = 0
        Else
          ' Not available for form items.
          mlngCaptionExprID = mctlSelectedControl.CalculationID
        End If
        txtCaptionTypeExpression.Text = GetExpressionName(mlngCaptionExprID)
              
        sngSubCurrentControlTop = sngSubCurrentControlTop _
          + YGAP_CONTROL_CONTROL
      End If
      
      With fraCaption
        .Width = msngFrameWidth - (2 * X_COLUMN1)
        .Height = sngSubCurrentControlTop + YGAP_CONTROL_FRAME
        .Visible = True
      End With
    
      sngCurrentControlTop = sngCurrentControlTop _
        + sngSubCurrentControlTop + YGAP_CONTROL_FRAME _
        + YGAP_FRAME_FRAME
    
      lblCaption.Visible = False
      txtCaption.Visible = False
    Else
      With lblCaption
        .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
        .Left = X_COLUMN1
        .Visible = True
      End With
  
      With txtCaption
        .Top = sngCurrentControlTop
        .Left = X_COLUMN2
        .Width = msngFrameWidth _
          - .Left _
          - X_COLUMN1
        .Visible = True
  
        If miItemType = giWFFORMITEM_FORM Then
          .Text = mfrmCallingForm.Caption
        Else
          .Text = Replace(mctlSelectedControl.Caption, "&&", "&")
        End If
      End With
  
      sngCurrentControlTop = sngCurrentControlTop _
        + YGAP_CONTROL_CONTROL
    
      fraCaption.Visible = False
    End If
  Else
    lblCaption.Visible = False
    txtCaption.Visible = False
    fraCaption.Visible = False
  End If
  
  ' Format the Description controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION) Then
    fFrameNeeded = True
    
    With lblDescription
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
    
    With txtDescriptionExpression
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdDescriptionExpression.Width
      .Visible = True
    End With
  
    With cmdDescriptionExpression
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
      .Enabled = True
    End With
  
    If miItemType = giWFFORMITEM_FORM Then
      mlngDescriptionExprID = mfrmCallingForm.DescriptionExprID
    Else
      ' Not available for form items.
      mlngDescriptionExprID = 0
    End If
    txtDescriptionExpression.Text = GetExpressionName(mlngDescriptionExprID)
  
    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblDescription.Visible = False
    txtDescriptionExpression.Visible = False
    cmdDescriptionExpression.Visible = False
  End If

  ' Format the Description (has Workflow Name) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION_WORKFLOWNAME) Then
    fFrameNeeded = True

    With chkDescriptionHasWorkflowName
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN2
      .Visible = True
  
      If miItemType = giWFFORMITEM_FORM Then
        .value = IIf(mfrmCallingForm.DescriptionHasWorkflowName, vbChecked, vbUnchecked)
      Else
        .value = vbUnchecked ' Only available for forms
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    chkDescriptionHasWorkflowName.Visible = False
  End If

  ' Format the Description (has Element Caption) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DESCRIPTION_ELEMENTCAPTION) Then
    fFrameNeeded = True

    With chkDescriptionHasElementCaption
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN2
      .Visible = True

      If miItemType = giWFFORMITEM_FORM Then
        .value = IIf(mfrmCallingForm.DescriptionHasElementCaption, vbChecked, vbUnchecked)
      Else
        .value = vbUnchecked ' Only available for forms
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    chkDescriptionHasElementCaption.Visible = False
  End If

  ' Format the frame
  With fraIdentification
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
      
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
        
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    
      RefreshIdentificationControls
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Identification = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function






Private Function FormatScreen_Frame_DefaultValue() As Boolean
  ' Format the DEFAULTVALUE frame on the DATA tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim lngMax As Long
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the first DefaultValueType controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
    fFrameNeeded = True
    
    With optDefaultValueType(giWFDATAVALUE_FIXED)
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
      
      If miItemType = giWFFORMITEM_FORM Then
        .value = False
      Else
        .value = (mctlSelectedControl.DefaultValueType = giWFDATAVALUE_FIXED)
      End If
    End With
    
    lblDefaultValueValue.Visible = False
    lblDefaultValueCalculation.Visible = False
  
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
    
    fFrameNeeded = True
    
    With lblDefaultValueCalculation
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
      
    optDefaultValueType(giWFDATAVALUE_FIXED).Visible = False
    optDefaultValueType(giWFDATAVALUE_CALC).Visible = False
    lblDefaultValueValue.Visible = False
  
  ElseIf WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_CHAR) _
      Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOGIC) _
      Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_DATE) _
      Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_NUMERIC) _
      Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LIST) _
      Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOOKUP) _
      Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_WORKPATTERN) Then
    
    fFrameNeeded = True
    
    With lblDefaultValueValue
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With
  
    optDefaultValueType(giWFDATAVALUE_FIXED).Visible = False
    optDefaultValueType(giWFDATAVALUE_CALC).Visible = False
    lblDefaultValueCalculation.Visible = False
  End If
  
  ' Format the DefaultValue (character) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_CHAR) Then
    fFrameNeeded = True

    With txtDefaultValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True

      .Text = Left(mctlSelectedControl.WFDefaultCharValue, Minimum(spnSize.value, 8000))
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    txtDefaultValue.Visible = False
  End If

  ' Format the DefaultValue (logic) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOGIC) Then
    fFrameNeeded = True

    With fraLogicDefaults
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN2
      .Visible = True
    
      optDefaultValue(0).value = mctlSelectedControl.WFDefaultValue
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    fraLogicDefaults.Visible = False
  End If

  ' Format the DefaultValue (date) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_DATE) Then
    fFrameNeeded = True

    With dtDefaultValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
    
      .Text = IIf(Len(mctlSelectedControl.WFDefaultValueDateString) > 0, _
        mctlSelectedControl.WFDefaultValueDateString, "")
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    dtDefaultValue.Visible = False
  End If

  ' Format the DefaultValue (numeric) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_NUMERIC) Then
    fFrameNeeded = True

    With spnDefaultValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = False
      ' TODO - future dev
      '.Visible = (spnDecimals.Value = 0)
    
      'If miItemType <> giWFFORMITEM_FORM Then
      '  If (spnDecimals.Value = 0) Then
      '    .Value = CLng(mctlSelectedControl.WFDefaultNumericValue)
      '  End If
      'End If
    End With

    With numDefaultValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
      '.Visible = (spnDecimals.Value > 0)
    
      If miItemType <> giWFFORMITEM_FORM Then
        'If (spnDecimals.Value > 0) Then
          .value = mctlSelectedControl.WFDefaultNumericValue
        'End If
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    spnDefaultValue.Visible = False
    numDefaultValue.Visible = False
  End If

  ' Format the DefaultValue (dropdown) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LIST) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_LOOKUP) Then
    
    fFrameNeeded = True

    With cboDefaultValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboDefaultValue_refresh mctlSelectedControl.DefaultStringValue
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    cboDefaultValue.Visible = False
  End If

  ' Format the DefaultValue (working pattern) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_WORKPATTERN) Then
    fFrameNeeded = True

    With wpDefaultValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Visible = True
      
      ' TODO - future dev
    End With

    sngCurrentControlTop = wpDefaultValue.Top _
      + wpDefaultValue.Height _
      - Y_STANDARDCONTROLHEIGHT _
      + YGAP_CONTROL_CONTROL
  Else
    wpDefaultValue.Visible = False
  End If
  
  ' Format the DefaultValue (expression) controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUE_EXPRID) Then
    fFrameNeeded = True

    If WebFormItemHasProperty(miItemType, WFITEMPROP_DEFAULTVALUETYPE) Then
      With optDefaultValueType(giWFDATAVALUE_CALC)
        .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
        .Left = X_COLUMN1
        .Visible = True
      
      
        If miItemType = giWFFORMITEM_FORM Then
          .value = False
        Else
          .value = (mctlSelectedControl.DefaultValueType = giWFDATAVALUE_CALC)
        End If
      End With
    End If
    
    With txtDefaultValueExpression
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdDefaultValueExpression.Width
      .Visible = True
    End With

    With cmdDefaultValueExpression
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
    End With

    If miItemType = giWFFORMITEM_FORM Then
      mlngDefaultValueExprID = 0
    Else
      ' Not available for form items.
      mlngDefaultValueExprID = mctlSelectedControl.CalculationID
    End If
    txtDefaultValueExpression.Text = GetExpressionName(mlngDefaultValueExprID)

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    txtDefaultValueExpression.Visible = False
    cmdDefaultValueExpression.Visible = False
  End If

  ' Format the frame
  With fraDefaultValue
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
      
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
        
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
      
      RefreshDefaultValueControls
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_DefaultValue = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function







Private Function FormatScreen_Frame_Lookup() As Boolean
  ' Format the LOOKUP frame on the DATA tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the LookupTable controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPTABLEID) Then
    fFrameNeeded = True

    With lblLookupTable
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboLookupTable
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboLookupTable_refresh mctlSelectedControl.LookupTableID
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblLookupTable.Visible = False
    cboLookupTable.Visible = False
  End If

  ' Format the LookupColumn controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPCOLUMNID) Then
    fFrameNeeded = True

    With lblLookupColumn
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboLookupColumn
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboLookupColumn_refresh mctlSelectedControl.LookupColumnID
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblLookupColumn.Visible = False
    cboLookupColumn.Visible = False
  End If
  
  ' Format the Order controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPORDER) Then
    fFrameNeeded = True

    With lblLookupOrder
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtLookupOrder
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdLookupOrder.Width
      .Visible = True
    End With

    With cmdLookupOrder
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
      .Enabled = True

      If miItemType = giWFFORMITEM_FORM Then
        ' Not available for form items.
        mlngLookupOrderID = 0
      Else
        mlngLookupOrderID = mctlSelectedControl.LookupOrderID
      End If
      txtLookupOrder.Text = GetOrderName(mlngLookupOrderID)
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblLookupOrder.Visible = False
    txtLookupOrder.Visible = False
    cmdLookupOrder.Visible = False
  End If
  
  ' Format the LookupFilter controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTER) Then
    fFrameNeeded = True

    With chkLookupFilter
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        .value = vbUnchecked ' Not available for forms
      Else
        .value = IIf((mctlSelectedControl.LookupFilterColumn > 0) And (Len(mctlSelectedControl.LookupFilterValue) > 0), vbChecked, vbUnchecked)
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    chkLookupFilter.Visible = False
  End If
  
  ' Format the LookupFilterColumn controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTERCOLUMN) Then
    fFrameNeeded = True

    With lblLookupFilterColumn
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboLookupFilterColumn
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True

      If miItemType <> giWFFORMITEM_FORM Then
        cboLookupFilterColumn_Refresh mctlSelectedControl.LookupFilterColumn
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblLookupFilterColumn.Visible = False
    cboLookupFilterColumn.Visible = False
  End If

  ' Format the LookupFilterOperator controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTEROPERATOR) Then
    fFrameNeeded = True

    With lblLookupFilterOperator
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboLookupFilterOperator
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True

      If miItemType <> giWFFORMITEM_FORM Then
        cboLookupFilterOperator_Refresh mctlSelectedControl.LookupFilterOperator
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblLookupFilterOperator.Visible = False
    cboLookupFilterOperator.Visible = False
  End If

  ' Format the LookupFilterValue controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_LOOKUPFILTERVALUE) Then
    fFrameNeeded = True

    With lblLookupFilterValue
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboLookupFilterValue
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True

      If miItemType <> giWFFORMITEM_FORM Then
        cboLookupFilterValue_Refresh mctlSelectedControl.LookupFilterValue
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblLookupFilterValue.Visible = False
    cboLookupFilterValue.Visible = False
  End If

  ' Format the frame
  With fraLookup
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
      
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
        
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_Lookup = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function







Private Function FormatScreen_Frame_RecordIdentification() As Boolean
  ' Format the RECORDIDENTIFICATION frame on the DATA tab.
  ' Return TRUE if the frame needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fFrameNeeded As Boolean
  Dim sngCurrentControlTop As Single
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim lngTableID As Long
  
  fFrameNeeded = False
  sngCurrentControlTop = YGAP_FRAME_CONTROL
  
  ' Format the Table controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_TABLEID) Then
    fFrameNeeded = True

    With lblRecordIdentificationTable
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboRecordIdentificationTable
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        If mctlSelectedControl.TableID = 0 Then
          If (mfrmCallingForm.BaseTable > 0) Then
            mctlSelectedControl.TableID = mfrmCallingForm.BaseTable
          Else
            sSQL = "SELECT TOP 1 tmpTables.tableID" & _
              " FROM tmpTables" & _
              " WHERE (tmpTables.deleted = FALSE)" & _
              " ORDER BY tmpTables.tableName"
            Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

            If Not (rsTables.EOF And rsTables.BOF) Then
              mctlSelectedControl.TableID = rsTables!TableID
            End If

            rsTables.Close
            Set rsTables = Nothing
          End If

          mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_ALL
        End If

        cboRecordIdentificationTable_refresh mctlSelectedControl.TableID
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblRecordIdentificationTable.Visible = False
    cboRecordIdentificationTable.Visible = False
  End If

  ' Format the Record controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_DBRECORD) _
    Or WebFormItemHasProperty(miItemType, WFITEMPROP_RECSELTYPE) Then
    
    fFrameNeeded = True

    With lblRecordIdentificationRecord
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboRecordIdentificationRecord
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        ' Ensure the currently selected record is valid.
        If miItemType = giWFFORMITEM_INPUTVALUE_GRID Then
          If cboRecordIdentificationTable.ListCount > 0 Then
            lngTableID = cboRecordIdentificationTable.ItemData(cboRecordIdentificationTable.ListIndex)
          End If
        Else
          lngTableID = GetTableIDFromColumnID(mctlSelectedControl.ColumnID)
        End If
        
        If (mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_MANUAL) _
          And (mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_TRIGGEREDRECORD) Then

          mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_INITIATOR
        End If

        If (mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) _
          And (mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_INITIATOR) Then

          mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_TRIGGEREDRECORD
        End If

        If (mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL) _
          And ((mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_INITIATOR) _
            Or (mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_TRIGGEREDRECORD)) Then

          mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD
        End If

        If WebFormItemHasProperty(miItemType, WFITEMPROP_DBRECORD) Then
          If (mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_INITIATOR) Then
            ReDim alngValidTables(0)
            TableAscendants mlngPersonnelTableID, alngValidTables
  
            fFound = False
            For lngLoop = 1 To UBound(alngValidTables)
              If lngTableID = alngValidTables(lngLoop) Then
                fFound = True
                Exit For
              End If
            Next lngLoop
  
            If Not fFound Then
              mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD
            End If
          End If
  
          If (mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_TRIGGEREDRECORD) Then
            ReDim alngValidTables(0)
            TableAscendants mfrmCallingForm.BaseTable, alngValidTables
  
            fFound = False
            For lngLoop = 1 To UBound(alngValidTables)
              If lngTableID = alngValidTables(lngLoop) Then
                fFound = True
                Exit For
              End If
            Next lngLoop
  
            If Not fFound Then
              mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD
            End If
          End If
  
          If (mfrmCallingForm.InitiationType <> WORKFLOWINITIATIONTYPE_EXTERNAL) _
            And (mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD) _
            And Len(Trim(mctlSelectedControl.WFWorkflowForm)) = 0 Then
  
            ReDim alngValidTables(0)
            TableAscendants mfrmCallingForm.BaseTable, alngValidTables
  
            fFound = False
            For lngLoop = 1 To UBound(alngValidTables)
              If lngTableID = alngValidTables(lngLoop) Then
                fFound = True
                Exit For
              End If
            Next lngLoop
  
            If fFound Then
              If (mfrmCallingForm.InitiationType = WORKFLOWINITIATIONTYPE_MANUAL) Then
                mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_INITIATOR
              Else
                mctlSelectedControl.WFDatabaseRecord = giWFRECSEL_TRIGGEREDRECORD
              End If
            End If
          End If
        End If

        cboRecordIdentificationRecord_refresh mctlSelectedControl.WFDatabaseRecord
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblRecordIdentificationRecord.Visible = False
    cboRecordIdentificationRecord.Visible = False
  End If

  ' Format the Element controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_ELEMENTIDENTIFIER) Then
    fFrameNeeded = True

    With lblRecordIdentificationElement
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboRecordIdentificationElement
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboRecordIdentificationElement_refresh mctlSelectedControl.WFWorkflowForm
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblRecordIdentificationElement.Visible = False
    cboRecordIdentificationElement.Visible = False
  End If

  ' Format the Record Selector controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDSELECTOR) Then
    fFrameNeeded = True

    With lblRecordIdentificationRecordSelector
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboRecordIdentificationRecordSelector
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboRecordIdentificationRecordSelector_refresh mctlSelectedControl.WFWorkflowValue
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
            
  Else
    lblRecordIdentificationRecordSelector.Visible = False
    cboRecordIdentificationRecordSelector.Visible = False
  End If

  ' Format the Record Table controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDTABLEID) Then
    fFrameNeeded = True

    With lblRecordIdentificationRecordTable
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With cboRecordIdentificationRecordTable
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1
      .Visible = True
    
      If miItemType <> giWFFORMITEM_FORM Then
        cboRecordIdentificationRecordTable_refresh mctlSelectedControl.WFRecordTableID
      End If
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblRecordIdentificationRecordTable.Visible = False
    cboRecordIdentificationRecordTable.Visible = False
  End If

  ' Format the Order controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDORDER) Then
    fFrameNeeded = True

    With lblRecordIdentificationOrder
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtRecordIdentificationOrder
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdRecordIdentificationOrder.Width
      .Visible = True
    End With

    With cmdRecordIdentificationOrder
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
      .Enabled = True

      If miItemType = giWFFORMITEM_FORM Then
        ' Not available for form items.
        mlngRecordIdentificationOrderID = 0
      Else
        mlngRecordIdentificationOrderID = mctlSelectedControl.WFRecordOrderID
      End If
      txtRecordIdentificationOrder.Text = GetOrderName(mlngRecordIdentificationOrderID)
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblRecordIdentificationOrder.Visible = False
    txtRecordIdentificationOrder.Visible = False
    cmdRecordIdentificationOrder.Visible = False
  End If

  ' Format the Filter controls if required.
  If WebFormItemHasProperty(miItemType, WFITEMPROP_RECORDFILTER) Then
    fFrameNeeded = True

    With lblRecordIdentificationFilter
      .Top = sngCurrentControlTop + YGAP_CONTROL_LABEL
      .Left = X_COLUMN1
      .Visible = True
    End With

    With txtRecordIdentificationFilter
      .Top = sngCurrentControlTop
      .Left = X_COLUMN2
      .Width = msngFrameWidth _
        - .Left _
        - X_COLUMN1 _
        - cmdRecordIdentificationFilter.Width
      .Visible = True
    
      If miItemType = giWFFORMITEM_FORM Then
        ' Not available for form items.
        mlngRecordIdentificationFilterID = 0
      Else
        mlngRecordIdentificationFilterID = mctlSelectedControl.WFRecordFilterID
      End If
      txtRecordIdentificationFilter.Text = GetExpressionName(mlngRecordIdentificationFilterID)
    End With

    With cmdRecordIdentificationFilter
      .Top = sngCurrentControlTop
      .Left = msngFrameWidth _
        - .Width _
        - X_COLUMN1
      .Visible = True
      .Enabled = True
    End With

    sngCurrentControlTop = sngCurrentControlTop _
      + YGAP_CONTROL_CONTROL
  Else
    lblRecordIdentificationFilter.Visible = False
    txtRecordIdentificationFilter.Visible = False
    cmdRecordIdentificationFilter.Visible = False
  End If

  ' Format the frame
  With fraRecordIdentification
    If fFrameNeeded Then
      .Height = sngCurrentControlTop + YGAP_CONTROL_FRAME
      .Top = msngCurrentFrameTop
      .Left = 0
      .Width = msngFrameWidth
      
      msngCurrentFrameTop = msngCurrentFrameTop _
        + .Height _
        + YGAP_FRAME_FRAME
        
      If msngMaxFrameBottom < (.Top + .Height) Then
        msngMaxFrameBottom = (.Top + .Height)
      End If
    End If
    
    .Visible = fFrameNeeded
  End With
  
TidyUpAndExit:
  FormatScreen_Frame_RecordIdentification = fFrameNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function







Private Function FormatScreen_Tab(piTab As Integer) As Boolean
  ' Format the given tab.
  ' Return TRUE if the tab needs to be displayed
  On Error GoTo ErrorTrap
  
  Dim fTabNeeded As Boolean
  
  fTabNeeded = False
  msngCurrentFrameTop = 0
  
  Select Case piTab
    Case miPAGE_GENERAL
      fTabNeeded = FormatScreen_Frame_Identification _
        Or FormatScreen_Frame_Display _
        Or FormatScreen_Frame_Behaviour _
        Or FormatScreen_Frame_HotSpot
        
    Case miPAGE_APPEARANCE
      fTabNeeded = FormatScreen_Frame_Options _
        Or FormatScreen_Frame_Header _
        Or FormatScreen_Frame_Foreground _
        Or FormatScreen_Frame_Background
    
    Case miPAGE_DATA
      fTabNeeded = FormatScreen_Frame_RecordIdentification _
        Or FormatScreen_Frame_Size _
        Or FormatScreen_Frame_ControlValues _
        Or FormatScreen_Frame_Lookup _
        Or FormatScreen_Frame_DefaultValue _
        Or FormatScreen_Frame_Validation
  End Select
  
  ssTabStrip.TabVisible(piTab) = fTabNeeded

TidyUpAndExit:
  FormatScreen_Tab = fTabNeeded
  
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Private Sub spnTimeoutFrequency_Click()
  spnTimeoutFrequency.SetFocus

End Sub

Private Sub spnTop_Change()
  Changed = True
End Sub

Private Sub spnTop_Click()
  spnTop.SetFocus
End Sub


Private Sub spnVOffset_Change()
  Changed = True
End Sub

Private Sub spnVOffset_Click()
  spnVOffset.SetFocus
End Sub

Private Sub spnWidth_Change()
  Changed = True
End Sub

Private Sub spnWidth_Click()
  spnWidth.SetFocus
End Sub

Private Sub ssTabStrip_Click(PreviousTab As Integer)
  ' Enable only the frames on the selected tab
  Dim iLoop As Integer
  
  ' Hide the controls that are NOT on the currently selected tab.
  For iLoop = 0 To ssTabStrip.Tabs - 1
    picTabContainer(iLoop).Visible = (ssTabStrip.Tab = iLoop)
  Next iLoop
  
  If ssTabStrip.Tab = miPAGE_DATA Then
    ResizeValidationColumns
  End If
  
End Sub

Private Sub txtCaption_Change()
  AutoResizeControl
  Changed = True
End Sub

Private Sub txtCaption_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtCaptionTypeValue_Change()
  AutoResizeControl
  Changed = True
End Sub

Private Sub txtCaptionTypeValue_GotFocus()
  UI.txtSelText
End Sub


Private Sub txtControlValues_Change()
  Dim sCurrentDefault As String
  Dim sList As String
  Dim sNewList As String
  Dim asControlValues() As String
  Dim iLoop As Integer
  Dim fValid As Boolean
  Dim iSelStart As Integer
  
  Const MAX_LENGTH = 200

  fValid = True
  sList = MergeControlValues(txtControlValues.Text)
  asControlValues() = Split(sList, vbTab)
  iSelStart = txtControlValues.SelStart
  
  For iLoop = 0 To UBound(asControlValues)
    If Len(asControlValues(iLoop)) > MAX_LENGTH Then
      asControlValues(iLoop) = Left(asControlValues(iLoop), MAX_LENGTH)
      fValid = False
    End If
  
    sNewList = sNewList & IIf(Len(sNewList) > 0, vbTab, "") & Left(asControlValues(iLoop), MAX_LENGTH)
  Next iLoop

  If Not fValid Then
    txtControlValues.Text = SplitControlValues(sNewList)
    txtControlValues.SelStart = iSelStart
  End If
  
  'JPD 20070927 Fault 12495
  If optDefaultValueType(giWFDATAVALUE_FIXED).value Then
    sCurrentDefault = ""
    
    If cboDefaultValue.ListCount > 0 Then
      sCurrentDefault = cboDefaultValue.List(cboDefaultValue.ListIndex)
    End If
    
    cboDefaultValue_refresh sCurrentDefault
  Else
    RefreshDefaultValueControls
  End If
  
  Changed = True

End Sub

Private Sub txtControlValues_GotFocus()
  ' Disable the 'Default' property of the 'OK' button as the return key is
  ' used by this textbox.
  cmdOK.Default = False

End Sub

Private Sub txtControlValues_LostFocus()
  ' Enable the 'Default' property of the OK button.
  cmdOK.Default = True

End Sub

Public Sub Initialise(pctlSelectedControl As Control, _
  pfrmCallingForm As frmWorkflowWFDesigner)
  
  Dim sCaption As String
  Dim wfOriginalCurrentElement As VB.Control
  Dim wfCurrentElement As VB.Control
  Dim iLoop As Integer
  Dim fFound As Boolean
  
  mfLoading = True
  
  Set mctlSelectedControl = pctlSelectedControl
  Set mfrmCallingForm = pfrmCallingForm
  
  mfReadOnly = mfrmCallingForm.ReadOnly
  
  Set wfCurrentElement = mfrmCallingForm.CurrentElementDefinition
  Set wfOriginalCurrentElement = mfrmCallingForm.Element
  
  ' Get the complete set of elements in the workflow.
  ' Replace the saved version of the current element with the current version as it has been edited.
  mfrmCallingForm.AllElements maWFAllElements
  If (Not wfCurrentElement Is Nothing) _
    And (Not wfOriginalCurrentElement Is Nothing) Then
  
    For iLoop = 1 To UBound(maWFAllElements)
      If (maWFAllElements(iLoop).ControlIndex = wfOriginalCurrentElement.ControlIndex) Then
        Set maWFAllElements(iLoop) = wfCurrentElement
      End If
    Next iLoop
  End If
  
  ' Get the elements that preceed the current one.
  ' NB. If the current element is looped back to, it will be in the set of preceeding elements.
  ' Replace the saved version of the current element with the current version as it has been edited.
  ' Also create an array of elements that preceed the current one, and include the current element
  ' even if its not looped back to.
  mfrmCallingForm.PrecedingElements maWFPrecedingElements
  ReDim Preserve maWFPrecedingAndCurrentElements(UBound(maWFPrecedingElements))
  
  fFound = False

  For iLoop = 1 To UBound(maWFPrecedingElements)
    If (iLoop > 1) _
      And (Not wfCurrentElement Is Nothing) _
      And (Not wfOriginalCurrentElement Is Nothing) Then
      
      If (maWFPrecedingElements(iLoop).ControlIndex = wfOriginalCurrentElement.ControlIndex) Then
        Set maWFPrecedingElements(iLoop) = wfCurrentElement
        fFound = True
      End If
    End If
    
    Set maWFPrecedingAndCurrentElements(iLoop) = maWFPrecedingElements(iLoop)
  Next iLoop

  If (Not fFound) And (Not wfCurrentElement Is Nothing) Then
    ReDim Preserve maWFPrecedingAndCurrentElements(UBound(maWFPrecedingAndCurrentElements) + 1)
    Set maWFPrecedingAndCurrentElements(UBound(maWFPrecedingAndCurrentElements)) = wfCurrentElement
  End If

  If pctlSelectedControl Is Nothing Then
    miItemType = giWFFORMITEM_FORM
  Else
    miItemType = mfrmCallingForm.WebFormControl_Type(pctlSelectedControl)
  End If
  
  sCaption = GetWebFormItemTypeName(CInt(miItemType))
  Select Case miItemType
    Case giWFFORMITEM_DBVALUE, _
      giWFFORMITEM_DBFILE
      
      sCaption = sCaption & " (" & GetColumnName(mctlSelectedControl.ColumnID) & ")"
    
    Case giWFFORMITEM_WFVALUE, _
      giWFFORMITEM_WFFILE
      
      sCaption = sCaption & " (" & mctlSelectedControl.WFWorkflowForm & "." & mctlSelectedControl.WFWorkflowValue & ")"
  End Select
  sCaption = sCaption & " - Properties"
  Me.Caption = sCaption
  
  If mfReadOnly Then
    ControlsDisableAll Me
  
    grdValidation.Enabled = True
  End If
  
  FormatScreen
  
  mfLoading = False
  
End Sub

Private Sub txtDefaultValue_Change()
  Changed = True

End Sub

Private Sub txtFileExtensions_Change()
  Dim sCurrentDefault As String
  Dim sList As String
  Dim sNewList As String
  Dim asFileExtensions() As String
  Dim iLoop As Integer
  Dim fValid As Boolean
  Dim iSelStart As Integer
  Dim sTemp As String
  Dim iIndex As Integer
  
  Const MAX_LENGTH = 10

  fValid = True
  sList = MergeControlValues(txtFileExtensions.Text)
  asFileExtensions() = Split(sList, vbTab)
  iSelStart = txtFileExtensions.SelStart

  For iLoop = 0 To UBound(asFileExtensions)
    sTemp = asFileExtensions(iLoop)
    iIndex = InStrRev(sTemp, ".")
    If iIndex > 0 Then
      sTemp = Mid(sTemp, iIndex + 1)
    End If
    
    If Len(sTemp) > MAX_LENGTH Then
      asFileExtensions(iLoop) = Left(asFileExtensions(iLoop), MAX_LENGTH)
      fValid = False
    End If

    sNewList = sNewList & IIf(Len(sNewList) > 0, vbTab, "") & Left(sTemp, MAX_LENGTH)
  Next iLoop

  If Not fValid Then
    txtFileExtensions.Text = SplitControlValues(sNewList)
    txtFileExtensions.SelStart = iSelStart
  End If

  Changed = True

End Sub

Private Sub txtFileExtensions_GotFocus()
  ' Disable the 'Default' property of the 'OK' button as the return key is
  ' used by this textbox.
  cmdOK.Default = False

End Sub


Private Sub txtFileExtensions_LostFocus()
  ' Enable the 'Default' property of the OK button.
  cmdOK.Default = True

End Sub

Private Sub txtIdentifier_Change()
  Changed = True

End Sub


Private Sub txtIdentifier_GotFocus()
  UI.txtSelText

End Sub


Private Sub wpDefaultValue_Click()
  Changed = True
  
End Sub

Private Function RefreshOrientationControls()
  Dim fEnable As Boolean
  
  Dim lngHeight As Long
  Dim lngWidth As Long
  
  lngHeight = spnHeight.value
  lngWidth = spnWidth.value
      
  If WebFormItemHasProperty(miItemType, WFITEMPROP_ORIENTATION) Then
    If miItemType = giWFFORMITEM_LINE Then
      spnHeight.value = lngWidth
      spnWidth.value = lngHeight
      
      fEnable = (optOrientation(0).value = True) And (Not mfReadOnly)
      
      EnableControl lblWidth, fEnable
      EnableControl spnWidth, fEnable
      EnableControl lblHeight, Not fEnable
      EnableControl spnHeight, Not fEnable

    End If
  End If
End Function


Private Sub cboMessage_refresh(piWhichMessage As WorkflowWebFormMessageType, _
  piCurrentMessageType As MessageType)
  
  Dim ctlMessageTypeCombo As Control
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim sTemp As String
  Dim sMessage As String
  Dim sTemp_pt1 As String
  Dim sTemp_pt2 As String
  Dim sTemp_pt3 As String
  
  fOK = True
  
  Select Case piWhichMessage
    Case WORKFLOWWEBFORMMESSAGE_COMPLETION
      Set ctlMessageTypeCombo = cboCompletionMessageType
      sMessage = msCompletionMessage
    Case WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER
      Set ctlMessageTypeCombo = cboSavedForLaterMessageType
      sMessage = msSavedForLaterMessage
    Case WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS
      Set ctlMessageTypeCombo = cboFollowOnFormsMessageType
      sMessage = msFollowOnFormsMessage
    Case Else
      fOK = False
  End Select
    
  If fOK Then
    iIndex = -1
    iDefaultIndex = 0

    With ctlMessageTypeCombo
      .Clear

      ' Populate the combo
      .AddItem "Default"
      .ItemData(.NewIndex) = MESSAGE_SYSTEMDEFAULT
  
      sTemp = "Custom"
      If Len(sMessage) > 0 Then
        ParseWebFormMessage sMessage, _
          sTemp_pt1, _
          sTemp_pt2, _
          sTemp_pt3
        sTemp = sTemp & " - """ & sTemp_pt1 & sTemp_pt2 & sTemp_pt3 & """"
        sTemp = Replace(Replace(sTemp, vbCr, ""), vbLf, "")
      End If
      .AddItem sTemp
      .ItemData(.NewIndex) = MESSAGE_CUSTOM
  
      .AddItem "None"
      .ItemData(.NewIndex) = MESSAGE_NONE
  
      ' Get the indexes of the required/default values
      For iLoop = 0 To .ListCount - 1
        If .ItemData(iLoop) = piCurrentMessageType Then
          iIndex = iLoop
          Exit For
        End If
      
        If .ItemData(iLoop) = MESSAGE_SYSTEMDEFAULT Then
          iDefaultIndex = iLoop
        End If
      Next iLoop

      If iIndex < 0 Then
        iIndex = iDefaultIndex
      End If

      .Enabled = (.ListCount > 0) And (Not mfReadOnly)
      If .ListCount > 0 Then
        .ListIndex = iIndex
      End If
    End With
  End If
  
End Sub



