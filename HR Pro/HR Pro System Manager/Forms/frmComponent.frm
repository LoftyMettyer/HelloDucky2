VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expression Component"
   ClientHeight    =   9915
   ClientLeft      =   525
   ClientTop       =   990
   ClientWidth     =   13125
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
   Icon            =   "frmComponent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraComponent 
      Caption         =   "Workflow Identified Field :"
      Height          =   4900
      Index           =   12
      Left            =   8160
      TabIndex        =   75
      Tag             =   "12"
      Top             =   3720
      Width           =   4700
      Begin VB.Frame fraWorkflowFieldRecord 
         Caption         =   "Record Identification :"
         Height          =   1900
         Left            =   100
         TabIndex        =   84
         Top             =   1400
         Width           =   4425
         Begin VB.ComboBox cboWorkflowFieldRecordTable 
            Height          =   315
            ItemData        =   "frmComponent.frx":000C
            Left            =   1725
            List            =   "frmComponent.frx":000E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   1440
            Width           =   1000
         End
         Begin VB.ComboBox cboWorkflowFieldRecord 
            Height          =   315
            ItemData        =   "frmComponent.frx":0010
            Left            =   1725
            List            =   "frmComponent.frx":0012
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   240
            Width           =   1000
         End
         Begin VB.ComboBox cboWorkflowFieldElement 
            Height          =   315
            ItemData        =   "frmComponent.frx":0014
            Left            =   1725
            List            =   "frmComponent.frx":0016
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   640
            Width           =   1000
         End
         Begin VB.ComboBox cboWorkflowFieldRecordSelector 
            Height          =   315
            ItemData        =   "frmComponent.frx":0018
            Left            =   1725
            List            =   "frmComponent.frx":001A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   1040
            Width           =   1000
         End
         Begin VB.Label lblWorkflowFieldRecordTable 
            Caption         =   "Table :"
            Height          =   195
            Left            =   105
            TabIndex        =   91
            Top             =   1500
            Width           =   1515
         End
         Begin VB.Label lblWorkflowFieldRecord 
            Caption         =   "Record :"
            Height          =   195
            Left            =   105
            TabIndex        =   85
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblWorkflowFieldElement 
            Caption         =   "Element :"
            Height          =   195
            Left            =   105
            TabIndex        =   87
            Top             =   705
            Width           =   1110
         End
         Begin VB.Label lblWorkflowFieldRecordSelector 
            Caption         =   "Record Selector :"
            Height          =   195
            Left            =   105
            TabIndex        =   89
            Top             =   1095
            Width           =   1515
         End
      End
      Begin VB.Frame fraWorkflowFldSelOptions 
         Caption         =   "Child Field Options :"
         Height          =   1500
         Left            =   100
         TabIndex        =   93
         Top             =   3300
         Width           =   4425
         Begin VB.CommandButton cmdWorkflowFldSelFilter 
            Caption         =   "..."
            Height          =   315
            Left            =   2270
            TabIndex        =   104
            Top             =   1000
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txtWorkflowFldSelFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   103
            Top             =   1000
            Width           =   735
         End
         Begin VB.TextBox txtWorkflowFldSelOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   100
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton cmdWorkflowFldSelOrder 
            Caption         =   "..."
            Height          =   315
            Left            =   2270
            TabIndex        =   101
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Frame fraWorkflowFieldSel 
            Height          =   315
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton optWorkflowFieldSel 
               Caption         =   "Specific"
               Height          =   195
               Index           =   2
               Left            =   2400
               TabIndex        =   97
               Top             =   60
               Width           =   990
            End
            Begin VB.OptionButton optWorkflowFieldSel 
               Caption         =   "Last"
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   96
               Top             =   60
               Width           =   735
            End
            Begin VB.OptionButton optWorkflowFieldSel 
               Caption         =   "First"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   95
               Top             =   60
               Value           =   -1  'True
               Width           =   700
            End
            Begin COASpinner.COA_Spinner asrWorkflowFldSelLine 
               Height          =   315
               Left            =   3400
               TabIndex        =   98
               Top             =   0
               Width           =   1000
               _ExtentX        =   1746
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
               MaximumValue    =   99999
               MinimumValue    =   1
               Text            =   "1"
            End
         End
         Begin VB.Label lblWorkflowFldFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter :"
            Height          =   195
            Left            =   105
            TabIndex        =   102
            Top             =   1060
            Width           =   465
         End
         Begin VB.Label lblWorkflowFldOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order :"
            Height          =   195
            Left            =   105
            TabIndex        =   99
            Top             =   660
            Width           =   525
         End
      End
      Begin VB.Frame fraWorkflowField 
         Height          =   315
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   4455
         Begin VB.OptionButton optWorkflowField 
            Caption         =   "Field"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   77
            Top             =   60
            Value           =   -1  'True
            Width           =   700
         End
         Begin VB.OptionButton optWorkflowField 
            Caption         =   "Count"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   78
            Top             =   60
            Width           =   975
         End
         Begin VB.OptionButton optWorkflowField 
            Caption         =   "Total"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   79
            Top             =   60
            Width           =   885
         End
      End
      Begin VB.ComboBox cboWorkflowFieldColumn 
         Height          =   315
         Left            =   1400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   1015
         Width           =   1000
      End
      Begin VB.ComboBox cboWorkflowFieldTable 
         Height          =   315
         Left            =   1400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   615
         Width           =   1000
      End
      Begin VB.Label lblWorkflowFieldColumn 
         Caption         =   "Column :"
         Height          =   195
         Left            =   105
         TabIndex        =   82
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblWorkflowFieldTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   105
         TabIndex        =   80
         Top             =   675
         Width           =   900
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Workflow Value :"
      Height          =   1100
      Index           =   11
      Left            =   8400
      TabIndex        =   70
      Tag             =   "11"
      Top             =   8760
      Width           =   2250
      Begin VB.ComboBox cboWFValueItem 
         Height          =   315
         ItemData        =   "frmComponent.frx":001C
         Left            =   1000
         List            =   "frmComponent.frx":001E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   650
         Width           =   1000
      End
      Begin VB.ComboBox cboWFValueElement 
         Height          =   315
         ItemData        =   "frmComponent.frx":0020
         Left            =   1000
         List            =   "frmComponent.frx":0022
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   250
         Width           =   1000
      End
      Begin VB.Label lblWFValueItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   105
         TabIndex        =   73
         Top             =   705
         Width           =   765
      End
      Begin VB.Label lblWFValueElement 
         BackStyle       =   0  'Transparent
         Caption         =   "Element :"
         Height          =   195
         Left            =   105
         TabIndex        =   71
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Dummy Frame"
      Height          =   390
      Index           =   9
      Left            =   5760
      TabIndex        =   143
      Tag             =   "9"
      Top             =   8760
      Width           =   1080
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Dummy Frame"
      Height          =   390
      Index           =   8
      Left            =   5400
      TabIndex        =   142
      Tag             =   "8"
      Top             =   8880
      Width           =   1080
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Filter :"
      Height          =   780
      Index           =   10
      Left            =   2520
      TabIndex        =   140
      Tag             =   "10"
      Top             =   8640
      Width           =   1395
      Begin VB.ListBox listCalcFilters 
         Height          =   255
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   141
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Prompted Value :"
      Height          =   4920
      Index           =   6
      Left            =   2520
      TabIndex        =   117
      Tag             =   "7"
      Top             =   3600
      Width           =   5505
      Begin VB.Frame fraPValFormat 
         Caption         =   "Format :"
         Height          =   1260
         Left            =   105
         TabIndex        =   132
         Top             =   1395
         Width           =   5250
         Begin VB.TextBox txtPValFormat 
            Height          =   315
            Left            =   195
            MaxLength       =   128
            TabIndex        =   48
            Top             =   270
            Width           =   1000
         End
         Begin VB.Label lblMaskKey1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A - Upper case"
            Height          =   195
            Left            =   195
            TabIndex        =   138
            Top             =   645
            Width           =   1380
         End
         Begin VB.Label lblMaskKey3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "9 - Numbers (0-9)"
            Height          =   195
            Left            =   1545
            TabIndex        =   137
            Top             =   645
            Width           =   1680
         End
         Begin VB.Label lblMaskKey4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "# - Numbers, Symbols"
            Height          =   195
            Left            =   1545
            TabIndex        =   136
            Top             =   900
            Width           =   2085
         End
         Begin VB.Label lblMaskKey2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a - Lower case"
            Height          =   195
            Left            =   195
            TabIndex        =   135
            Top             =   885
            Width           =   1275
         End
         Begin VB.Label lblMaskKey5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B - Binary (0 or 1)"
            Height          =   195
            Left            =   3525
            TabIndex        =   134
            Top             =   645
            Width           =   1725
         End
         Begin VB.Label lblMaskKey6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "\ - Follow by literal"
            Height          =   195
            Left            =   3570
            TabIndex        =   133
            Top             =   900
            Width           =   1605
         End
      End
      Begin VB.Frame fraPValTable 
         Caption         =   "Lookup Table Value :"
         Height          =   975
         Left            =   135
         TabIndex        =   129
         Top             =   2730
         Width           =   2235
         Begin VB.ComboBox cboPValColumn 
            Height          =   315
            Left            =   975
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   555
            Width           =   1000
         End
         Begin VB.ComboBox cboPValTable 
            Height          =   315
            Left            =   855
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   225
            Width           =   1000
         End
         Begin VB.Label lblPValColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Column :"
            Height          =   195
            Left            =   105
            TabIndex        =   131
            Top             =   615
            Width           =   810
         End
         Begin VB.Label lblPValTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.Frame fraPValDefaultValue 
         Caption         =   "Default Value :"
         Height          =   960
         Left            =   135
         TabIndex        =   124
         Top             =   3795
         Width           =   4545
         Begin TDBNumber6Ctl.TDBNumber TDBPValDefaultNumeric 
            Height          =   315
            Left            =   2235
            TabIndex        =   51
            Top             =   225
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            Calculator      =   "frmComponent.frx":0024
            Caption         =   "frmComponent.frx":0044
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmComponent.frx":00A9
            Keys            =   "frmComponent.frx":00C7
            Spin            =   "frmComponent.frx":0111
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,##0.#######; -##,###,##0.#######"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,##0.#######; -##,###,##0.#######"
            HighlightText   =   -1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.OptionButton optPValDefaultLogic 
            Caption         =   "&No"
            Height          =   225
            Index           =   1
            Left            =   1200
            TabIndex        =   54
            Top             =   600
            Width           =   675
         End
         Begin VB.OptionButton optPValDefaultLogic 
            Caption         =   "&Yes"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   53
            Top             =   600
            Value           =   -1  'True
            Width           =   630
         End
         Begin VB.ComboBox cboPValDefaultTabVal 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   255
            Width           =   1000
         End
         Begin VB.TextBox txtPValDefaultCharacter 
            Height          =   315
            Left            =   75
            TabIndex        =   49
            Top             =   225
            Width           =   1000
         End
         Begin GTMaskDate.GTMaskDate asrPValDefaultDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   50
            Top             =   225
            Width           =   1050
            _Version        =   65537
            _ExtentX        =   1852
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
      End
      Begin VB.Frame fraPValValueType 
         Caption         =   "Type :"
         Height          =   705
         Left            =   100
         TabIndex        =   121
         Top             =   600
         Width           =   4980
         Begin VB.ComboBox cboPValReturnType 
            Height          =   315
            ItemData        =   "frmComponent.frx":0139
            Left            =   120
            List            =   "frmComponent.frx":014C
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   240
            Width           =   1000
         End
         Begin COASpinner.COA_Spinner asrPValReturnDecimals 
            Height          =   315
            Left            =   3825
            TabIndex        =   45
            Top             =   270
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
            MaximumValue    =   4
            Text            =   "0"
         End
         Begin COASpinner.COA_Spinner asrPValReturnSize 
            Height          =   315
            Left            =   1770
            TabIndex        =   44
            Top             =   255
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
            MinimumValue    =   1
            Text            =   "1"
         End
         Begin VB.Label lblPValSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            Height          =   195
            Left            =   1230
            TabIndex        =   123
            Top             =   285
            Width           =   525
         End
         Begin VB.Label lblPValDecimals 
            BackStyle       =   0  'Transparent
            Caption         =   "Decimals :"
            Height          =   195
            Left            =   2865
            TabIndex        =   122
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.TextBox txtPValPrompt 
         Height          =   315
         Left            =   885
         MaxLength       =   40
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   250
         Width           =   1140
      End
      Begin VB.Label lblPValPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt :"
         Height          =   195
         Left            =   105
         TabIndex        =   120
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Operator :"
      Height          =   1035
      Index           =   4
      Left            =   10800
      TabIndex        =   113
      Tag             =   "5"
      Top             =   8760
      Width           =   1200
      Begin SSActiveTreeView.SSTree ssTreeOpOperator 
         Height          =   645
         Left            =   105
         TabIndex        =   59
         Top             =   255
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1138
         _Version        =   65538
         LabelEdit       =   1
         Style           =   6
         Indentation     =   525
         HideSelection   =   0   'False
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   1
      End
   End
   Begin VB.Frame fraComponent 
      BackColor       =   &H80000010&
      Caption         =   "Custom Calculation :"
      Height          =   2800
      Index           =   7
      Left            =   6990
      TabIndex        =   118
      Tag             =   "8"
      Top             =   75
      Width           =   2200
      Begin VB.Frame fraCustParameters 
         Caption         =   "Parameters :"
         Height          =   2100
         Left            =   100
         TabIndex        =   125
         Top             =   600
         Width           =   1750
         Begin VB.ComboBox cboCustField 
            Height          =   315
            ItemData        =   "frmComponent.frx":017E
            Left            =   650
            List            =   "frmComponent.frx":018E
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1600
            Width           =   1000
         End
         Begin VB.ComboBox cboCustTable 
            Height          =   315
            ItemData        =   "frmComponent.frx":01B5
            Left            =   650
            List            =   "frmComponent.frx":01C5
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1200
            Width           =   1000
         End
         Begin VB.ListBox lstCustParameters 
            Height          =   840
            Left            =   100
            TabIndex        =   56
            Top             =   250
            Width           =   1000
         End
         Begin VB.Label lblCustField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Column :"
            Height          =   195
            Left            =   105
            TabIndex        =   127
            Top             =   1665
            Width           =   630
         End
         Begin VB.Label lblCustTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   100
            TabIndex        =   126
            Top             =   1260
            Width           =   495
         End
      End
      Begin VB.ComboBox cboCustCalculation 
         Height          =   315
         ItemData        =   "frmComponent.frx":01EC
         Left            =   1100
         List            =   "frmComponent.frx":01FC
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   250
         Width           =   1000
      End
      Begin VB.Label lblCustCalculation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculation :"
         Height          =   195
         Left            =   105
         TabIndex        =   119
         Top             =   315
         Width           =   885
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Lookup Table Value :"
      Height          =   2160
      Index           =   5
      Left            =   120
      TabIndex        =   114
      Tag             =   "6"
      Top             =   4680
      Width           =   2445
      Begin VB.ComboBox cboTabValValue 
         Height          =   315
         ItemData        =   "frmComponent.frx":0223
         Left            =   975
         List            =   "frmComponent.frx":0225
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1050
         Width           =   1275
      End
      Begin VB.ComboBox cboTabValColumn 
         Height          =   315
         ItemData        =   "frmComponent.frx":0227
         Left            =   975
         List            =   "frmComponent.frx":0229
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   650
         Width           =   1275
      End
      Begin VB.ComboBox cboTabValTable 
         Height          =   315
         ItemData        =   "frmComponent.frx":022B
         Left            =   975
         List            =   "frmComponent.frx":022D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   250
         Width           =   1275
      End
      Begin VB.Label lblLookupValNotFound 
         Caption         =   "Original lookup value was not found"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   120
         TabIndex        =   144
         Top             =   1500
         Visible         =   0   'False
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTabValValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   105
         TabIndex        =   128
         Top             =   1110
         Width           =   765
      End
      Begin VB.Label lblTabValColumn 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   105
         TabIndex        =   116
         Top             =   705
         Width           =   900
      End
      Begin VB.Label lblTabValTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   105
         TabIndex        =   115
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Calculation :"
      Height          =   650
      Index           =   3
      Left            =   4080
      TabIndex        =   112
      Tag             =   "3"
      Top             =   8760
      Width           =   1200
      Begin VB.ListBox listCalcCalculation 
         Height          =   255
         Left            =   100
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   250
         Width           =   1000
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Function :"
      Height          =   3435
      Index           =   2
      Left            =   9360
      TabIndex        =   111
      Tag             =   "2"
      Top             =   120
      Width           =   3120
      Begin VB.Frame fraWorkflowFunctionRecord 
         Caption         =   "Record Identification :"
         Height          =   1900
         Left            =   120
         TabIndex        =   61
         Top             =   1320
         Width           =   2835
         Begin VB.ComboBox cboWorkflowFunctionRecordSelector 
            Height          =   315
            ItemData        =   "frmComponent.frx":022F
            Left            =   1680
            List            =   "frmComponent.frx":0231
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1040
            Width           =   1000
         End
         Begin VB.ComboBox cboWorkflowFunctionElement 
            Height          =   315
            ItemData        =   "frmComponent.frx":0233
            Left            =   1680
            List            =   "frmComponent.frx":0235
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   640
            Width           =   1000
         End
         Begin VB.ComboBox cboWorkflowFunctionRecord 
            Height          =   315
            ItemData        =   "frmComponent.frx":0237
            Left            =   1680
            List            =   "frmComponent.frx":0239
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   240
            Width           =   1000
         End
         Begin VB.ComboBox cboWorkflowFunctionRecordTable 
            Height          =   315
            ItemData        =   "frmComponent.frx":023B
            Left            =   1680
            List            =   "frmComponent.frx":023D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   1440
            Width           =   1000
         End
         Begin VB.Label lblWorkflowFunctionRecordSelector 
            Caption         =   "Record Selector :"
            Height          =   195
            Left            =   105
            TabIndex        =   66
            Top             =   1095
            Width           =   1515
         End
         Begin VB.Label lblWorkflowFunctionElement 
            Caption         =   "Element :"
            Height          =   195
            Left            =   100
            TabIndex        =   64
            Top             =   700
            Width           =   840
         End
         Begin VB.Label lblWorkflowFunctionRecord 
            Caption         =   "Record :"
            Height          =   195
            Left            =   105
            TabIndex        =   62
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblWorkflowFunctionRecordTable 
            Caption         =   "Table :"
            Height          =   195
            Left            =   100
            TabIndex        =   68
            Top             =   1500
            Width           =   1245
         End
      End
      Begin SSActiveTreeView.SSTree ssTreeFuncFunction 
         Height          =   1005
         Left            =   105
         TabIndex        =   60
         Top             =   255
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   1773
         _Version        =   65538
         LabelEdit       =   1
         Style           =   6
         Indentation     =   525
         HideSelection   =   0   'False
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   1
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Field :"
      Height          =   3375
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      Tag             =   "1"
      Top             =   120
      Width           =   4665
      Begin VB.ComboBox cboFldDummyColumn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Frame fraField 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4455
         Begin VB.OptionButton optField 
            Caption         =   "Total"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   14
            Top             =   60
            Width           =   840
         End
         Begin VB.OptionButton optField 
            Caption         =   "Count"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   13
            Top             =   60
            Width           =   930
         End
         Begin VB.OptionButton optField 
            Caption         =   "Field"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   60
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin VB.ComboBox cboFldColumn 
         Height          =   315
         Left            =   1000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1005
         Width           =   1000
      End
      Begin VB.ComboBox cboFldTable 
         Height          =   315
         ItemData        =   "frmComponent.frx":023F
         Left            =   1000
         List            =   "frmComponent.frx":0241
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   615
         Width           =   1000
      End
      Begin VB.Frame fraFldSelOptions 
         Caption         =   "Child Field Options :"
         Height          =   1920
         Left            =   100
         TabIndex        =   20
         Top             =   1300
         Width           =   4425
         Begin VB.Frame fraFieldSel 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton optFieldSel 
               Caption         =   "First"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   22
               Top             =   60
               Value           =   -1  'True
               Width           =   700
            End
            Begin VB.OptionButton optFieldSel 
               Caption         =   "Last"
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   23
               Top             =   60
               Width           =   735
            End
            Begin VB.OptionButton optFieldSel 
               Caption         =   "Specific"
               Height          =   195
               Index           =   2
               Left            =   2400
               TabIndex        =   24
               Top             =   60
               Width           =   990
            End
            Begin COASpinner.COA_Spinner asrFldSelLine 
               Height          =   315
               Left            =   3400
               TabIndex        =   25
               Top             =   0
               Width           =   1000
               _ExtentX        =   1746
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
               MaximumValue    =   99999
               MinimumValue    =   1
               Text            =   "1"
            End
         End
         Begin VB.CommandButton cmdFldSelOrder 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   2270
            TabIndex        =   28
            Top             =   1035
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txtFldSelOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   27
            Top             =   1035
            Width           =   735
         End
         Begin VB.TextBox txtFldSelFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   30
            Top             =   1395
            Width           =   735
         End
         Begin VB.CommandButton cmdFldSelFilter 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   2270
            TabIndex        =   31
            Top             =   1395
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Label lblFldOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order :"
            Height          =   195
            Left            =   105
            TabIndex        =   26
            Top             =   1095
            Width           =   525
         End
         Begin VB.Label lblFldFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter :"
            Height          =   195
            Left            =   105
            TabIndex        =   29
            Top             =   1455
            Width           =   465
         End
      End
      Begin VB.Label lblFldField 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   1065
         Width           =   810
      End
      Begin VB.Label lblFldDatabase 
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   675
         Width           =   675
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Value :"
      Height          =   2250
      Index           =   1
      Left            =   120
      TabIndex        =   108
      Tag             =   "4"
      Top             =   6960
      Width           =   2385
      Begin TDBNumber6Ctl.TDBNumber TDBValNumericValue 
         Height          =   315
         Left            =   780
         TabIndex        =   35
         Top             =   1050
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   556
         Calculator      =   "frmComponent.frx":0243
         Caption         =   "frmComponent.frx":0263
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmComponent.frx":02C8
         Keys            =   "frmComponent.frx":02E6
         Spin            =   "frmComponent.frx":0330
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##,###,##0.#######; -##,###,##0.#######"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##,###,##0.#######; -##,###,##0.#######"
         HighlightText   =   -1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   6750213
         MinValueVT      =   3538949
      End
      Begin VB.ComboBox cboValType 
         Height          =   315
         ItemData        =   "frmComponent.frx":0358
         Left            =   780
         List            =   "frmComponent.frx":0368
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   250
         Width           =   1000
      End
      Begin VB.OptionButton optValLogicValue 
         Caption         =   "&True"
         Height          =   315
         Index           =   0
         Left            =   765
         TabIndex        =   37
         Top             =   1850
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optValLogicValue 
         Caption         =   "&False"
         Height          =   315
         Index           =   1
         Left            =   1530
         TabIndex        =   38
         Top             =   1850
         Width           =   810
      End
      Begin VB.TextBox txtValCharacterValue 
         Height          =   315
         Left            =   780
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   645
         Width           =   1000
      End
      Begin GTMaskDate.GTMaskDate asrValDateValue 
         Height          =   315
         Left            =   780
         TabIndex        =   36
         Top             =   1440
         Width           =   1050
         _Version        =   65537
         _ExtentX        =   1852
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
      Begin VB.Label lblValType 
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   105
         TabIndex        =   110
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblValValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   105
         TabIndex        =   109
         Top             =   705
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7080
      TabIndex        =   105
      Top             =   9360
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7080
      TabIndex        =   106
      Top             =   8880
      Width           =   1200
   End
   Begin VB.Frame fraComponentType 
      Caption         =   "Type :"
      Height          =   4500
      Left            =   45
      TabIndex        =   107
      Top             =   45
      Width           =   2580
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Workflow Identified Field"
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   1
         Tag             =   "COMP_FIELD"
         Top             =   650
         Width           =   2415
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "Wo&rkflow Value"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   2
         Tag             =   "COMP_FIELD"
         Top             =   1000
         Width           =   1815
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "F&ilter"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   139
         Tag             =   "COMP_FILTER"
         Top             =   3450
         Width           =   1035
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "F&unction"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Tag             =   "COMP_FUNCTION"
         Top             =   1700
         Width           =   1320
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "C&alculation"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Tag             =   "COMP_CALCULATION"
         Top             =   3100
         Width           =   1410
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Value"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Tag             =   "COMP_VALUE"
         Top             =   2050
         Width           =   1020
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "Op&erator"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Tag             =   "COMP_OPERATOR"
         Top             =   1350
         Width           =   1320
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "Loo&kup Table Value"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Tag             =   "COMP_LOOKUPVALUE"
         Top             =   2400
         Width           =   1995
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Prompted Value"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Tag             =   "COMP_PROMPTED"
         Top             =   2750
         Width           =   1815
      End
      Begin VB.OptionButton optComponentType 
         BackColor       =   &H80000010&
         Caption         =   "Custo&m Calculation"
         Height          =   315
         Index           =   8
         Left            =   105
         TabIndex        =   9
         Tag             =   "COMP_CUSTOMCALC"
         Top             =   3800
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Field"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Tag             =   "COMP_FIELD"
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mobjComponent As CExprComponent
Private miComponentType As ExpressionComponentTypes
Private mavColumns() As Variant
Private mDataType As Long

' Form handling variables.
Private mfCancelled As Boolean
Private mfFunctionsPopulated As Boolean
Private mfCalculationsPopulated As Boolean
Private mfFiltersPopulated As Boolean
Private mfOperatorsPopulated As Boolean
Private mfTabValTablesPopulated As Boolean
Private mfPValTablesPopulated As Boolean

Private mfFieldByValue As Boolean

Private mfInitializing As Boolean
Private msInitializeMessage As String

Private maWFPrecedingElements() As VB.Control

Private mlngHierarchyTableID As Long
Private mlngPersonnelTableID As Long
Private mlngWFPersonnelTableID As Long
Private mlngDependantsTableID As Long
Private mlngMaternityTableID As Long



Private Sub cboWorkflowFieldColumn_Refresh()
  ' Populate the Workflow Field component - Column combo, and then select the current field.
  Dim iIndex As Integer
  Dim iNextIndex As Integer
  Dim lngTableID As Long
  Dim lngColumnID As Long

  ' Get the current component's table and column id.
  lngTableID = mobjComponent.Component.TableID
  lngColumnID = mobjComponent.Component.ColumnID

  ' Clear the current contents of the combo.
  cboWorkflowFieldColumn.Clear

  If Not optWorkflowField(1).value Then
    ' Create an array of column info.
    ' Column 1 = column ID
    ' Column 2 = data type.
    ReDim mavColumns(2, 0)
  
    With recColEdit
      .Index = "idxName"
      .Seek ">=", lngTableID
  
      If Not .NoMatch Then
        If Not (.BOF And .EOF) Then
          .MoveFirst
        End If
  
        ' Add items to the combo for each field that has not been deleted
        Do While Not .EOF
          ' Do not allow the user to select system columns, deleted columns, or
          ' OLE or Photo type columns.
          If (!TableID = lngTableID) And _
            (!Deleted = False) And _
            (!DataType <> dtLONGVARBINARY) And _
            (!DataType <> dtVARBINARY) And _
            (!ColumnType <> giCOLUMNTYPE_LINK) And _
            (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
            ((optWorkflowField(2).value = False) Or _
              (!DataType = dtNUMERIC) Or _
              (!DataType = dtINTEGER)) Then
  
            cboWorkflowFieldColumn.AddItem .Fields("columnName")
            cboWorkflowFieldColumn.ItemData(cboWorkflowFieldColumn.NewIndex) = .Fields("columnID")
  
            iNextIndex = UBound(mavColumns, 2) + 1
            ReDim Preserve mavColumns(2, iNextIndex)
            mavColumns(1, iNextIndex) = !ColumnID
            mavColumns(2, iNextIndex) = !DataType
  
            If .Fields("columnID") = lngColumnID Then
              iIndex = cboWorkflowFieldColumn.NewIndex
            End If
          End If
  
          .MoveNext
        Loop
      End If
    End With
  End If
  
  ' Enable the combo if there are items.
  With cboWorkflowFieldColumn

    If .ListCount > 0 Then
      .ListIndex = iIndex
      .Enabled = True
    Else
      .Enabled = False
      optWorkflowFieldSel_Refresh

      If optWorkflowField(2).value Then
        cboWorkflowFieldColumn.AddItem "<no numeric columns>"
      ElseIf optWorkflowField(0).value Then
        cboWorkflowFieldColumn.AddItem "<no columns>"
      Else
        cboWorkflowFieldColumn.AddItem ""
      End If
      cboWorkflowFieldColumn.ItemData(cboWorkflowFieldColumn.NewIndex) = 0
      cboWorkflowFieldColumn.ListIndex = 0
    End If
  End With

End Sub
Private Sub cboFldColumn_Refresh()
  ' Populate the Field component - Field combo, and then select the current field.
  Dim iIndex As Integer
  Dim iNextIndex As Integer
  Dim lngTableID As Long
  Dim lngColumnID As Long
  
  ' Get the current component's table and column id.
  lngTableID = mobjComponent.Component.TableID
  lngColumnID = mobjComponent.Component.ColumnID
  
  ' Clear the current contents of the combo.
  cboFldColumn.Clear
  
  ' Create an array of column info.
  ' Column 1 = column ID
  ' Column 2 = data type.
  ReDim mavColumns(2, 0)
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngTableID
    
    If Not .NoMatch Then
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If
      
      ' Add items to the combo for each field that has not been deleted
      ' and is not a calculated column.
      iIndex = 0
      
      Do While Not .EOF
        
        ' Do not allow the user to select system columns, deleted columns, or
        ' OLE or Photo type columns.
        If (!TableID = lngTableID) And _
          (!Deleted = False) And _
          (!DataType <> dtLONGVARBINARY) And _
          (!DataType <> dtVARBINARY) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
          ((optField(2).value = False) Or _
            (!DataType = dtNUMERIC) Or _
            (!DataType = dtINTEGER)) Then
                          
          cboFldColumn.AddItem .Fields("columnName")
          cboFldColumn.ItemData(cboFldColumn.NewIndex) = .Fields("columnID")
            
          iNextIndex = UBound(mavColumns, 2) + 1
          ReDim Preserve mavColumns(2, iNextIndex)
          mavColumns(1, iNextIndex) = !ColumnID
          mavColumns(2, iNextIndex) = !DataType
          
          If .Fields("columnID") = lngColumnID Then
            iIndex = cboFldColumn.NewIndex
          End If
        End If
        
        .MoveNext
        
      Loop
    
    End If
    
  End With
  
  ' Enable the combo if there are items.
  With cboFldColumn
    If .ListCount > 0 Then
      .ListIndex = iIndex
      .Enabled = True
    Else
      .Enabled = False
      optFieldSel_Refresh
      
      If optField(2).value Then
        cboFldColumn.AddItem "<no numeric columns>"
      Else
        cboFldColumn.AddItem "<no columns>"
      End If
      cboFldColumn.ItemData(cboFldColumn.NewIndex) = 0
      cboFldColumn.ListIndex = 0
    End If
        
    cmdOK.Enabled = .Enabled
  End With

End Sub

Private Sub cboPValReturnType_Refresh()
  ' Prompted Value component - Return Type combo.
  ' Select the current return type.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iType As ExpressionValueTypes
  
  iIndex = 0
  iType = mobjComponent.Component.ValueType
  
  ' Loop through the available return types.
  For iLoop = 0 To cboPValReturnType.ListCount - 1
    ' Select the current return type if it is in the combo's list.
    If cboPValReturnType.ItemData(iLoop) = iType Then
      iIndex = iLoop
      Exit For
    End If
  Next iLoop
  
  cboPValReturnType.ListIndex = iIndex
  
End Sub

Private Sub cboPValTable_Refresh()
  ' Prompted Value component - Table combo.
  ' Select the current table.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim lngTableID As Long
  
  iIndex = 0
  lngTableID = mobjComponent.Component.LookupTable
  
  If cboPValTable.Enabled Then
    For iLoop = 0 To cboPValTable.ListCount - 1
      If cboPValTable.ItemData(iLoop) = lngTableID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
      
    cboPValTable.ListIndex = iIndex
  Else
    cboPValColumn_Refresh
  End If
  
End Sub
Private Sub cboWorkflowFieldTable_Refresh()
  ' Populate the Workflow Field component - Table combo and
  ' select the current table if it is still valid.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long

  lngTableID = mobjComponent.Component.TableID
  iIndex = -1
  iDefaultIndex = -1

  ' Clear the current contents of the combo.
  cboWorkflowFieldTable.Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    ' Add  an item to the combo for each table that has not been deleted.
    Do While Not .EOF
      If (Not .Fields("deleted")) Then
        If optWorkflowField(0).value _
          Or (!TableType = iTabChild) Then
        
          cboWorkflowFieldTable.AddItem !TableName
          cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.NewIndex) = !TableID
        End If
      End If

      .MoveNext
    Loop
  End With

  ' Enable the combo if there are items.
  With cboWorkflowFieldTable
    For iLoop = 0 To .ListCount - 1
      If .ItemData(iLoop) = lngTableID Then
        iIndex = iLoop
      End If
    
      If (.ItemData(iLoop) = mobjComponent.ParentExpression.UtilityBaseTable) Then
        iDefaultIndex = iLoop
      End If
    Next iLoop

    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        If iDefaultIndex >= 0 Then
          iIndex = iDefaultIndex
        Else
          iIndex = 0
        End If
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no tables>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0

      mobjComponent.Component.TableID = 0

      cboWorkflowFieldColumn_Refresh
      cboWorkflowFieldRecord_Refresh
      fldSelOrder_Refresh
      fldSelFilter_Refresh
    End If
  End With
    
End Sub

Private Sub cboFldTable_Refresh()
  ' Populate the Field component - Database combo and
  ' select the current table if it is still valid.
  Dim fOK As Boolean
  Dim fTableOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  Dim lngRootTableID As Long
  
  ' Determine if the field component is passed by value.
  lngTableID = mobjComponent.Component.TableID
  lngRootTableID = mobjComponent.ParentExpression.BaseTableID
  iIndex = -1
  iDefaultIndex = -1
  
  ' Clear the current contents of the combo.
  cboFldTable.Clear
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    ' Add  an item to the combo for each table that has not been deleted.
    ' If the field component is 'by value' then only the expression's
    ' root table, and its parents and children are valid.
    Do While Not .EOF
      fTableOK = False
      
      If (Not .Fields("deleted")) Then
        If mfFieldByValue Then
          If optField(0).value Then
            If .Fields("tableID") = lngRootTableID Then
              ' The table is the root table.
              If mobjComponent.ParentExpression.ExpressionType <> giEXPR_DEFAULTVALUE Then
                fTableOK = True
              End If
              
            'JPD 20031216 Islington changes
            'ElseIf (mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) Then
            Else
              recRelEdit.Index = "idxParentID"
              recRelEdit.Seek "=", .Fields("tableID"), lngRootTableID
              
              If Not recRelEdit.NoMatch Then
                ' The table is the parent of the root table.
            
                'JPD 20031216 Islington changes
                If (mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) Then
                  fTableOK = True
                End If
              Else
                recRelEdit.Seek "=", lngRootTableID, .Fields("tableID")
                If Not recRelEdit.NoMatch Then
                  ' The table is the child of the root table.
                  If mobjComponent.ParentExpression.ExpressionType <> giEXPR_DEFAULTVALUE Then
                    fTableOK = True
                  End If
                End If
              End If
            End If
          Else
            ' Only add child tables
            recRelEdit.Index = "idxParentID"
            recRelEdit.Seek "=", lngRootTableID, .Fields("tableID")
            If Not recRelEdit.NoMatch Then
              ' The table is the child of the root table.
              If mobjComponent.ParentExpression.ExpressionType <> giEXPR_DEFAULTVALUE Then
                fTableOK = True
              End If
            End If
          End If
        Else
          fTableOK = True
        End If
        
        If fTableOK Then
          cboFldTable.AddItem !TableName
          cboFldTable.ItemData(cboFldTable.NewIndex) = !TableID
          
          If !TableID = lngTableID Then
            iIndex = cboFldTable.NewIndex
          End If
          
          If !TableID = lngRootTableID Then
            iDefaultIndex = cboFldTable.NewIndex
          End If
        End If
      End If
      
      .MoveNext
    Loop
  End With
  
  ' Enable the combo if there are items.
  With cboFldTable
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        If iDefaultIndex >= 0 Then
          iIndex = iDefaultIndex
        Else
          iIndex = 0
        End If
      End If
      cboFldTable.ListIndex = iIndex
    Else
      .Enabled = False
      
      cboFldTable.AddItem "<no tables>"
      cboFldTable.ItemData(cboFldTable.NewIndex) = 0
      cboFldTable.ListIndex = 0
      
      mobjComponent.Component.TableID = 0
    
      cboFldColumn_Refresh
      fldSelOrder_Refresh
      fldSelFilter_Refresh
    End If
  End With
    
End Sub


Private Sub cboValType_Refresh()
  ' Value component - Type combo.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iType As ExpressionValueTypes
  
  iIndex = 0
  iType = mobjComponent.Component.ReturnType
  
  For iLoop = 0 To cboValType.ListCount - 1
    
    If cboValType.ItemData(iLoop) = iType Then
      iIndex = iLoop
      Exit For
    End If
        
  Next iLoop
    
  cboValType.ListIndex = iIndex

End Sub

Private Sub listCalcCalculation_Refresh()
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  If listCalcCalculation.Enabled Then
    iIndex = 0
    
    For iLoop = 0 To listCalcCalculation.ListCount - 1
      If listCalcCalculation.ItemData(iLoop) = mobjComponent.Component.CalculationID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
      
    listCalcCalculation.ListIndex = iIndex
  End If
  
End Sub


Private Sub listCalcFilters_Refresh()
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  If listCalcFilters.Enabled Then
    iIndex = 0
    
    For iLoop = 0 To listCalcFilters.ListCount - 1
      If listCalcFilters.ItemData(iLoop) = mobjComponent.Component.FilterID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
      
    listCalcFilters.ListIndex = iIndex
  End If
  
End Sub

Private Sub cboTabValTable_Initialize()
  ' Populate the Table Value component - Table combo.
  On Error GoTo ErrorTrap
  
  ' Clear the contents of the combo.
  cboTabValTable.Clear
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      ' Add items to the combo for each lookup table that is not deleted.
      If (!TableType = iTabLookup) And _
        (Not !Deleted) Then
      
        cboTabValTable.AddItem !TableName
        cboTabValTable.ItemData(cboTabValTable.NewIndex) = !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
  ' Enable the combo if there are items.
  With cboTabValTable
    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = 0      'MH20021105 Fault 4696
    Else
      .Enabled = False
      cboTabValColumn_Refresh
    End If
  End With
  
  ' Set the flag to show that the combo has been populated.
  mfTabValTablesPopulated = True
  Exit Sub
  
ErrorTrap:
  cboTabValTable.Enabled = False
  cboTabValColumn_Refresh
  Err = False

End Sub
Private Sub cboPValTable_Initialize()
  ' Populate the Prompted Value component - Table combo.
  On Error GoTo ErrorTrap
  
  ' Clear the contents of the combo.
  cboPValTable.Clear
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      ' Add items to the combo for each lookup table that is not deleted.
      If (!TableType = iTabLookup) And _
        (Not !Deleted) Then
      
        cboPValTable.AddItem !TableName
        cboPValTable.ItemData(cboPValTable.NewIndex) = !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
  ' Enable the combo if there are items.
  With cboPValTable
    .Enabled = (.ListCount > 0)
    .ListIndex = 0
  End With
  
  ' Set the flag to show that the combo has been populated.
  mfPValTablesPopulated = True

  Exit Sub
  
ErrorTrap:
  cboPValTable.Enabled = False
  
  Err = False

End Sub

Private Sub cboTabValColumn_Refresh()
  ' Populate the Table Value component - Column combo, and
  ' select the first item.
  On Error GoTo ErrorTrap
  
  Dim iNextIndex As Integer
  Dim lngTableID As Long
  
  ' Clear the current contents of the combo.
  cboTabValColumn.Clear
  
  ' Create an array of column info.
  ' Column 1 = column ID
  ' Column 2 = data type.
  ReDim mavColumns(2, 0)
  
  If cboTabValTable.Enabled Then
  
    lngTableID = cboTabValTable.ItemData(cboTabValTable.ListIndex)
  
    ' Loop through columns for selected lookup table.
    With recColEdit
      .Index = "idxName"
      .Seek ">=", lngTableID
      
      If Not .NoMatch Then
        Do While Not .EOF
          If !TableID <> lngTableID Then
            Exit Do
          End If
          
          ' Add each column name to the lookup columns combo.
          ' NB. We only want to add certain types of column. There's not use in
          ' looking up OLE or logic values.
          If (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
            (!ColumnType <> giCOLUMNTYPE_LINK) And _
            (Not !Deleted) And _
            (!DataType <> dtLONGVARBINARY) And _
            (!DataType <> dtVARBINARY) And _
            (!DataType <> dtBIT) Then
            
            cboTabValColumn.AddItem .Fields("columnName")
            cboTabValColumn.ItemData(cboTabValColumn.NewIndex) = .Fields("columnID")
          
            iNextIndex = UBound(mavColumns, 2) + 1
            ReDim Preserve mavColumns(2, iNextIndex)
            mavColumns(1, iNextIndex) = !ColumnID
            mavColumns(2, iNextIndex) = !DataType
          End If
      
          .MoveNext
        Loop
      End If
    End With
    
    ' Select a list item, and enable the combo, if there are items in the list.
    With cboTabValColumn
      If .ListCount > 0 Then
        .Enabled = True
        .ListIndex = 0
      Else
        .Enabled = False
        cboTabValValue_Refresh
      End If
    End With
  Else
    cboTabValColumn.Enabled = False
    cboTabValValue_Refresh
  End If
  
  Exit Sub

ErrorTrap:
  cboTabValColumn.Enabled = False
  cboTabValValue_Refresh
  Err = False

End Sub
Private Sub cboPValColumn_Refresh()
  ' Populate the Prompted Value - Column combo, and then
  ' select the current column.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iNextIndex As Integer
  Dim lngTableID As Long
  Dim lngColumnID As Long
  
  iIndex = 0
  iLoop = 0
  
  ' Clear the current contents of the combo.
  cboPValColumn.Clear
  
  ' Create an array of column info.
  ' Column 1 = column ID
  ' Column 2 = data type.
  ReDim mavColumns(2, 0)
  
  If cboPValTable.Enabled Then
    ' Get the current component's table and column id.
    lngTableID = cboPValTable.ItemData(cboPValTable.ListIndex)
    lngColumnID = mobjComponent.Component.LookupColumn
    
    With recColEdit
      .Index = "idxName"
      .Seek ">=", lngTableID
      
      If Not .NoMatch Then
        If Not (.BOF And .EOF) Then
          .MoveFirst
        End If
        
        ' Add items to the combo for each field that has not been deleted
        ' and is not a calculated column.
        ' NB. We only want to add certain types of column. There's not use in
        ' looking up OLE or logic values.
        Do While Not .EOF
          If (!TableID = lngTableID) And _
            (!Deleted = False) And _
            (!ColumnType <> giCOLUMNTYPE_SYSTEM) And _
            (!ColumnType <> giCOLUMNTYPE_LINK) And _
            (!DataType <> dtLONGVARBINARY) And _
            (!DataType <> dtVARBINARY) And _
            (!DataType <> dtBIT) Then
                            
            cboPValColumn.AddItem .Fields("columnName")
            cboPValColumn.ItemData(cboPValColumn.NewIndex) = !ColumnID
            
            iNextIndex = UBound(mavColumns, 2) + 1
            ReDim Preserve mavColumns(2, iNextIndex)
            mavColumns(1, iNextIndex) = !ColumnID
            mavColumns(2, iNextIndex) = !DataType
              
            If .Fields("columnID") = lngColumnID Then
              iIndex = iLoop
            End If
            
            iLoop = iLoop + 1
          End If
          
          .MoveNext
        Loop
      End If
    End With
    
    ' Enable the combo if there are items.
    With cboPValColumn
      If .ListCount > 0 Then
        .ListIndex = iIndex
        .Enabled = True
      Else
        .Enabled = False
        cboPValDefaultTabVal_Refresh
      End If
       
    End With
  Else
    cboPValColumn.Enabled = False
    cboPValDefaultTabVal_Refresh
  End If
  
  cmdOK.Enabled = (Len(Trim(mobjComponent.Component.Prompt)) > 0) And _
    ((mobjComponent.Component.ValueType <> giEXPRVALUE_TABLEVALUE) Or (cboPValColumn.Enabled))

  Exit Sub
  
ErrorTrap:
  cboPValColumn.Enabled = False
  cboPValDefaultTabVal_Refresh
  Err = False

End Sub

Private Sub cboPValDefaultTabVal_Refresh()
  ' Populate the Prompted Value - Default Table Value combo, and then
  ' select the current column.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsLookupValues As New ADODB.Recordset
  Dim sDfltValue As String
  Dim iIndex As Integer
  Dim objMisc As Misc

  Set objMisc = New Misc
  
  iIndex = 0
  
  sDfltValue = mobjComponent.Component.DefaultValue
  
  ' Clear the current contents of the combo.
  cboPValDefaultTabVal.Clear

  If cboPValTable.Enabled And cboPValColumn.Enabled Then
    ' Get the values from the lookup table.
    sSQL = "SELECT " & cboPValColumn.List(cboPValColumn.ListIndex) & " AS lookUpValue" & _
      " FROM " & cboPValTable.List(cboPValTable.ListIndex)
      
    rsLookupValues.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsLookupValues
      ' Add an item to the combo for each lookup value.
      While Not .EOF
        Select Case mDataType
          Case dtNUMERIC, dtINTEGER
            cboPValDefaultTabVal.AddItem Trim(Str(!LookupValue))
            If !LookupValue = Val(sDfltValue) Then
              iIndex = cboPValDefaultTabVal.NewIndex
            End If

          Case dtTIMESTAMP
            If IsDate(!LookupValue) Then
              'JPD 20041115 Fault 8970
              'cboPValDefaultTabVal.AddItem Format(!LookupValue, "long date")
              cboPValDefaultTabVal.AddItem Format(!LookupValue, objMisc.DateFormat)
              If Replace(Format(!LookupValue, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") = sDfltValue Then
                iIndex = cboPValDefaultTabVal.NewIndex
              End If
            End If

          Case Else
            cboPValDefaultTabVal.AddItem Trim(!LookupValue)
            If !LookupValue = sDfltValue Then
              iIndex = cboPValDefaultTabVal.NewIndex
            End If
        End Select
          
        .MoveNext
      Wend
  
      .Close
    End With
      
    Set rsLookupValues = Nothing
  End If
               
  ' Select a list item, and enable the combo, if there are items in the list.
  With cboPValDefaultTabVal
    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = iIndex
    Else
      .Enabled = False
    End If
  End With
  
  Set objMisc = Nothing
  
  Exit Sub
  
ErrorTrap:
  Set rsLookupValues = Nothing
  cboPValDefaultTabVal.Enabled = False
  Err = False

End Sub


Private Sub cboTabValValue_Refresh()
  ' Populate the Table Value component - Value combo, and
  ' select the first item.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim iColumnType As Integer
  Dim rsLookupValues As New ADODB.Recordset
  
  ' RH 03/10/00 - Order the values combo by the lookup tables default order
  Dim objOrder As Order
  Dim colOrderItems As Collection
  Dim sOrderString As String
  Dim iCount As Integer
  Dim objMisc As Misc

  Set objMisc = New Misc
  
  ' Clear the current contents of the combo.
  cboTabValValue.Clear

  If cboTabValTable.Enabled And cboTabValColumn.Enabled Then
    ' Get the values from the lookup table.
    
''    ' JPD 22/02/01 - Removed use of the default order. Order by the lookup column values.
''    ' RH 03/10/00 - Order the values combo by the lookup tables def order
''    sSQL = "SELECT DefaultOrderID FROM AsrSysTables WHERE TableID = " & cboTabValTable.ItemData(cboTabValTable.ListIndex)
''    Set rsLookupValues = rdoCon.OpenResultset(sSQL, _
''      rdOpenForwardOnly, rdConcurReadOnly, rdExecDirect)
''    If Not (rsLookupValues.BOF And rsLookupValues.EOF) Then
''      Set objOrder = New Order
''      objOrder.OrderID = rsLookupValues!defaultOrderID
''      objOrder.ConstructOrder
''      Set colOrderItems = objOrder.OrderItems
''      For iCount = 1 To colOrderItems.Count - 1
''       sOrderString = sOrderString & IIf(Len(sOrderString) > 0, ", ", "") & colOrderItems.Item(iCount).ColumnName & " " & IIf(colOrderItems.Item(iCount).Ascending, "ASC", "DESC")
''      Next iCount
''    End If
    
    ' Note : The following line is not added by RH and needs to remain should
    ' the changes made on 03/10 be undone.
    sSQL = "SELECT " & cboTabValColumn.List(cboTabValColumn.ListIndex) & " AS lookUpValue" & _
      " FROM " & cboTabValTable.List(cboTabValTable.ListIndex) & _
      " ORDER BY lookUpValue"
         
''    ' RH 03/10/00 - Order the values combo by the lookup tables def order
''    If Len(sOrderString) > 0 Then sSQL = sSQL & " ORDER BY " & sOrderString
    
    rsLookupValues.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsLookupValues
      ' Add an item to the combo for each function.
      While Not .EOF
        Select Case mDataType
          Case dtNUMERIC, dtINTEGER
            cboTabValValue.AddItem Trim(Str(!LookupValue))
          Case dtTIMESTAMP
            If IsDate(!LookupValue) Then
              'JPD 20041115 Fault 8970
              'cboTabValValue.AddItem Format(!LookupValue, "long date")
              cboTabValValue.AddItem Format(!LookupValue, objMisc.DateFormat)
            End If
          Case Else
            cboTabValValue.AddItem Trim(!LookupValue)
        End Select
             
        .MoveNext
      Wend
     
      .Close
    End With
         
  End If
                 
  ' Select a list item, and enable the combo, if there are items in the list.
  With cboTabValValue
    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = 0
    Else
      .Enabled = False
    End If
  End With
  
TidyUpAndExit:
  Set rsLookupValues = Nothing
  Set objMisc = Nothing
  cmdOK.Enabled = cboTabValValue.Enabled
     
  Exit Sub
  
ErrorTrap:
  cboTabValValue.Enabled = False
  Err = False
  Resume TidyUpAndExit
  
End Sub


Private Sub ReadModuleParameters()
  
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
  mlngWFPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  mlngDependantsTableID = GetModuleSetting(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE, 0)
  mlngMaternityTableID = GetModuleSetting(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYTABLE, 0)
  mlngHierarchyTableID = GetModuleSetting(gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE, 0)

End Sub

Private Function RemoveEmptyCategories() As Boolean

  Dim i As Integer
  Dim objNode As SSNode

  On Error GoTo Error_Trap
  
  With ssTreeFuncFunction
    For Each objNode In .Nodes
      If (objNode.Parent Is objNode.Root) And (objNode.Children = 0) Then
        .Nodes.Remove objNode.Index
      End If
    Next objNode
  End With
  RemoveEmptyCategories = True
  
TidyUpAndExit:
  Set objNode = Nothing
  Exit Function
  
Error_Trap:
  MsgBox "Error validating function categories.", vbExclamation + vbOKOnly, App.Title
  RemoveEmptyCategories = False
  GoTo TidyUpAndExit
  
End Function

Private Function FunctionTableOK(plngFunctionID As Long, _
  plngBaseTableID As Long) As Boolean
  
  Dim fOK As Boolean

  fOK = True

  Select Case plngFunctionID
    Case 17
      ' 17 - Current User
      If (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
        (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
        (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
    
        fOK = (plngBaseTableID = mlngWFPersonnelTableID)
      End If
    Case 30
      ' 30 - Absence Duration
      fOK = (plngBaseTableID = mlngPersonnelTableID) _
        Or IsChildOfTable(mlngPersonnelTableID, plngBaseTableID)
    Case 46
      ' 46 - Working Days Between Two Dates
      fOK = (plngBaseTableID = mlngPersonnelTableID) _
        Or IsChildOfTable(mlngPersonnelTableID, plngBaseTableID)
    Case 47
      ' 47 - Absence Between Two Dates
      fOK = (plngBaseTableID = mlngPersonnelTableID) _
        Or IsChildOfTable(mlngPersonnelTableID, plngBaseTableID)
    Case 62
      ' 62 - Parental Leave Entitlement
      fOK = (plngBaseTableID = mlngDependantsTableID)
    Case 63
      ' 63 - Parental Leave Taken
      fOK = (plngBaseTableID = mlngDependantsTableID)
    Case 64
      ' 64 - Maternity Return Date
      fOK = (plngBaseTableID = mlngMaternityTableID)
    Case 66
      ' 66 - Post that Reports to Current User
      fOK = (plngBaseTableID = mlngHierarchyTableID)
    Case 68
      ' 68 - Personnel that Reports to Current User
      fOK = (plngBaseTableID = mlngPersonnelTableID)
    Case 70
      ' 70 - Post that Current User Reports to
      fOK = (plngBaseTableID = mlngHierarchyTableID)
    Case 72
      ' 72 - Personnel that Current User Reports to
      fOK = (plngBaseTableID = mlngPersonnelTableID)
    Case 73
      ' 73 - Bradford Factor
      fOK = (plngBaseTableID = mlngPersonnelTableID) _
        Or IsChildOfTable(mlngPersonnelTableID, plngBaseTableID)
  End Select

  FunctionTableOK = fOK
  
End Function

Private Function FunctionRequiredTableID(plngFunctionID As Long) As Long
  ' Workflow record selection is only required for the following functions:
  ' Return:
  '   The ID of the required table
  '   0 if no specific table AND no record identification is required
  '   -1 if no specific table is required BUT record identification IS required
  Dim lngRequiredTableID As Long
  
  lngRequiredTableID = 0
  
  Select Case plngFunctionID
    Case 17
      ' 17 - Current User
      If (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
        (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
        (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
    
        lngRequiredTableID = mlngWFPersonnelTableID
      End If
    Case 30
      ' 30 - Absence Duration
      lngRequiredTableID = mlngPersonnelTableID
    Case 46
      ' 46 - Working Days Between Two Dates
      lngRequiredTableID = mlngPersonnelTableID
    Case 47
      ' 47 - Absence Between Two Dates
      lngRequiredTableID = mlngPersonnelTableID
    Case 62
      ' 62 - Parental Leave Entitlement
      lngRequiredTableID = mlngDependantsTableID
    Case 63
      ' 63 - Parental Leave Taken
      lngRequiredTableID = mlngDependantsTableID
    Case 64
      ' 64 - Maternity Return Date
      lngRequiredTableID = mlngMaternityTableID
    Case 66
      ' 66 - Post that Reports to Current User
      lngRequiredTableID = mlngPersonnelTableID
    Case 68
      ' 68 - Personnel that Reports to Current User
      lngRequiredTableID = mlngPersonnelTableID
    Case 70
      ' 70 - Post that Current User Reports to
      lngRequiredTableID = mlngPersonnelTableID
    Case 72
      ' 72 - Personnel that Current User Reports to
      lngRequiredTableID = mlngPersonnelTableID
    Case 73
      ' 73 - Bradford Factor
      lngRequiredTableID = mlngPersonnelTableID
  
    Case 74
      ' 74 - Does Record Exist
      lngRequiredTableID = -1
  End Select

  FunctionRequiredTableID = lngRequiredTableID
  
End Function


Private Function ssTreeFuncFunction_SelectedFunctionID() As Long
  Dim lngFunctionID As Long
  
  lngFunctionID = 0
  
  With ssTreeFuncFunction
    If .SelectedNodes.Count > 0 Then
      If .SelectedItem.key <> "FUNCTION_ROOT" Then
        If .SelectedItem.Parent.key <> "FUNCTION_ROOT" Then
          lngFunctionID = .SelectedItem.key
        End If
      End If
    End If
  End With
  
  ssTreeFuncFunction_SelectedFunctionID = lngFunctionID
  
End Function

Private Sub ssTreeOpOperator_Initialize()
  ' Populate the Operators tree view with the standard operators.
  On Error GoTo ErrorTrap
  
  Dim fCategoryDone As Boolean
  Dim iLoop As Integer
  Dim sCategory As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim objOperatorDef As clsOperatorDef
  Dim sDisplayName As String
  
  ' Clear the treeview.
  ssTreeOpOperator.Nodes.Clear
  
  ' Create the root node.
  Set objNode = ssTreeOpOperator.Nodes.Add(, , "OPERATOR_ROOT", "Operators")
  With objNode
    .Expanded = True
    .Font.Bold = True
    .Sorted = ssatSortAscending
  End With
  Set objNode = Nothing
    
  UI.LockWindow Me.hWnd
  
  For Each objOperatorDef In gobjOperatorDefs
    ' Add a category node if required.
    sCategory = objOperatorDef.Category
    fCategoryDone = False
    For iLoop = 1 To ssTreeOpOperator.Nodes.Count
      If ssTreeOpOperator.Nodes(iLoop).key = sCategory Then
        fCategoryDone = True
        Exit For
      End If
    Next iLoop
    
    If Not fCategoryDone Then
      Set objNode = ssTreeOpOperator.Nodes.Add("OPERATOR_ROOT", tvwChild, sCategory, sCategory)
      objNode.Font.Bold = True
      objNode.Sorted = ssatSortAscending
      Set objNode = Nothing
    End If
    
    ' Add the operator node.
    sDisplayName = objOperatorDef.Name
    If Len(objOperatorDef.ShortcutKeys) > 0 Then
      sDisplayName = sDisplayName & " (" & objOperatorDef.ShortcutKeys & ")"
    End If
    
    Set objNode = ssTreeOpOperator.Nodes.Add(sCategory, tvwChild, Trim(Str(objOperatorDef.id)), sDisplayName)
    Set objNode = Nothing
  Next objOperatorDef
  Set objOperatorDef = Nothing
  
  ' Enable the treeview only if there are items.
  With ssTreeOpOperator
    If .Nodes.Count > 0 Then
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With

  ' Set the flag to show that the treeview has been populated.
  mfOperatorsPopulated = True

TidyUpAndExit:
  UI.UnlockWindow
  Set objNode = Nothing
  ssTreeOpOperator.Refresh
  Exit Sub
  
ErrorTrap:
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  ssTreeOpOperator.Enabled = False
  Resume TidyUpAndExit

End Sub

Private Sub ssTreeFuncFunction_Initialize()
  ' Populate the Functions tree view with the standard functions.
  On Error GoTo ErrorTrap
  
  Dim fCategoryDone As Boolean
  Dim iLoop As Integer
  Dim sCategory As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim sSPName As String
  Dim objFunctionDef As clsFunctionDef
  Dim sDisplayName As String
  Dim fWorkflowExpression As Boolean
  
  fWorkflowExpression = (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)

  ' Clear the treeview.
  ssTreeFuncFunction.Nodes.Clear
  
  ' Create the root node.
  Set objNode = ssTreeFuncFunction.Nodes.Add(, , "FUNCTION_ROOT", "Functions")
  With objNode
    .Expanded = True
    .Font.Bold = True
    .Sorted = ssatSortAscending
  End With
  Set objNode = Nothing
      
  UI.LockWindow Me.hWnd

  ' Add an item to the treeview for each function.
  For Each objFunctionDef In gobjFunctionDefs
    If ((mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMECALCULATION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMEFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_LINKFILTER)) Or _
      objFunctionDef.Runtime Then

      'MH20040213 Fault 8086
      'If InStr(objFunctionDef.ExcludeTypes, CStr(mobjComponent.ParentExpression.ExpressionType)) = 0 Then
      If (InStr(" " & objFunctionDef.ExcludeTypes & " ", " " & CStr(mobjComponent.ParentExpression.ExpressionType) & " ") = 0) _
        And ((Len(objFunctionDef.IncludeTypes) = 0) Or (InStr(" " & objFunctionDef.IncludeTypes & " ", " " & CStr(mobjComponent.ParentExpression.ExpressionType) & " ") > 0)) Then

        sSPName = LCase(objFunctionDef.SPName)
        sDisplayName = objFunctionDef.Name
        If Len(objFunctionDef.ShortcutKeys) > 0 Then
          sDisplayName = sDisplayName & " " & objFunctionDef.ShortcutKeys
        End If
    
        ' Add a category node if required.
        sCategory = objFunctionDef.Category
        fCategoryDone = False
        For iLoop = 1 To ssTreeFuncFunction.Nodes.Count
          If ssTreeFuncFunction.Nodes(iLoop).key = sCategory Then
            fCategoryDone = True
            Exit For
          End If
        Next iLoop
      
        If Not fCategoryDone Then
          Set objNode = ssTreeFuncFunction.Nodes.Add("FUNCTION_ROOT", tvwChild, sCategory, sCategory)
          With objNode
            .Font.Bold = True
            .Sorted = ssatSortAscending
          End With
          Set objNode = Nothing
        End If
  
        'JPD 20040107 Use Ids rather than SPName
        Select Case objFunctionDef.id
          'Case "sp_asrfn_absenceduration", _
               "sp_asrfn_absencebetweentwodates", _
               "sp_asrfn_workingdaysbetweentwodates"
          Case 30, 46, 47, 73
            If (mlngPersonnelTableID > 0) _
              And (fWorkflowExpression _
                Or mobjComponent.ParentExpression.BaseTableID = mlngPersonnelTableID _
                Or IsChildOfTable(mlngPersonnelTableID, mobjComponent.ParentExpression.BaseTableID)) Then
              
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
            End If
    
          'Case "spasrsysfnparentalleaveentitlement", _
               "spasrsysfnparentalleavetaken"
          Case 62, 63
            If (mlngDependantsTableID > 0) _
              And (fWorkflowExpression _
                Or mobjComponent.ParentExpression.BaseTableID = mlngDependantsTableID) _
              And (mobjComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWRUNTIMEFILTER) Then
                
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
            End If
    
          'Case "spasrsysfnmaternityexpectedreturn"
          Case 64
            If (mlngMaternityTableID > 0) _
              And (fWorkflowExpression _
                Or mobjComponent.ParentExpression.BaseTableID = mlngMaternityTableID) _
              And (mobjComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWRUNTIMEFILTER) Then
                
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
            End If
    
          'JPD 20040127 Hierarchy performance modifications
          'Case 67, 68, 71, 72
          Case 68, 72
            If (mlngPersonnelTableID > 0) _
              And (modHierarchySpecifics.HierarchyFunctionConfigured(objFunctionDef.id)) _
              And (((mobjComponent.ParentExpression.ExpressionType = giEXPR_VIEWFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMECALCULATION) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMEFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_LINKFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)) And _
                  (mobjComponent.ParentExpression.BaseTableID = mlngPersonnelTableID)) Then
              
              If fWorkflowExpression And (objFunctionDef.id = 68) Then
                ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), "Is Personnel That Reports To Identified Person"
              ElseIf fWorkflowExpression And (objFunctionDef.id = 72) Then
                ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), "Is Personnel That Identified Person Reports To"
              Else
                ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
              End If
            End If
    
          'JPD 20040127 Hierarchy performance modifications
          'Case 65, 66, 69, 70
          Case 66, 70
            If (mlngHierarchyTableID > 0) _
              And (modHierarchySpecifics.HierarchyFunctionConfigured(objFunctionDef.id)) _
              And (((mobjComponent.ParentExpression.ExpressionType = giEXPR_VIEWFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMECALCULATION) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMEFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_LINKFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
                  (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)) And _
                  (mobjComponent.ParentExpression.BaseTableID = mlngHierarchyTableID)) Then
              
              If fWorkflowExpression And (objFunctionDef.id = 66) Then
                ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), "Is Post That Reports To Identified Person"
              ElseIf fWorkflowExpression And (objFunctionDef.id = 70) Then
                ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), "Is Post That Identified Person Reports To"
              Else
                ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
              End If
            End If
          
          Case 17 ' Current user
            If ((mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
              (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
              (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)) Then
              
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), "Login Name"
            Else
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
            End If
            
          Case 52, 53
            ' Some functions not available for Workflow expressions:
            '   52 - Field Last Change Date
            '   53 - Field Changed between Two Dates
            If ((mobjComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWCALCULATION) And _
              (mobjComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWSTATICFILTER) And _
              (mobjComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWRUNTIMEFILTER)) Then
              
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
            End If
            
          'JPD 20081103 Fault 13411
          Case 41, 43
            ' Some functions not available for Workflow runtime filter expressions:
            '   41 - Statutory Redundancy Pay
            '   43 - Unique Code
            If (mobjComponent.ParentExpression.ExpressionType <> giEXPR_WORKFLOWRUNTIMEFILTER) Then
              ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
            End If
            
          Case Else
            ssTreeFuncFunction.Nodes.Add sCategory, tvwChild, CStr(objFunctionDef.id), sDisplayName
  
        End Select
      End If
    End If
    
  Next objFunctionDef
  Set objFunctionDef = Nothing
  
  'TM20020121 Fault 3367
  'Remove the nodes(categories) if there are no sub nodes.
  RemoveEmptyCategories
  
  ' Enable the treeview only if there are items.
  With ssTreeFuncFunction
    If .Nodes.Count > 0 Then
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With

  ' Set the flag to show that the treeview has been populated.
  mfFunctionsPopulated = True

TidyUpAndExit:
  UI.UnlockWindow
  Set objNode = Nothing
  ssTreeFuncFunction.Refresh
  Exit Sub
  
ErrorTrap:
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  ssTreeFuncFunction.Enabled = False
  Resume TidyUpAndExit

End Sub


Private Sub listCalcCalculation_Initialize()
  ' Populate the Calculation component - Calculation listbox.
  On Error GoTo ErrorTrap
  Dim lngExpressionID As Long
  
  lngExpressionID = mobjComponent.ParentExpression.ExpressionID
  
  ' Clear the current contents of the listbox.
  listCalcCalculation.Clear
  
  UI.LockWindow Me.hWnd
  
  ' Add an item to the listbox for each calculation.
  With recExprEdit
    .Index = "idxExprName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      ' Add items to the listbox for each expression that is not deleted or hidden by another user.

      'MH20000727 Added Email
      If (Not !ExprID = lngExpressionID) And _
        ((!Type = giEXPR_COLUMNCALCULATION) Or _
        (!Type = giEXPR_RECORDDESCRIPTION) Or _
        (!Type = giEXPR_OUTLOOKFOLDER) Or _
        (!Type = giEXPR_OUTLOOKSUBJECT) Or _
        (!Type = giEXPR_EMAIL) Or _
        (!Type = giEXPR_RECORDVALIDATION)) And _
        (!ParentComponentID = 0) And _
        (Not !Deleted) And _
        ((!UserName = gsUserName) Or (!Access <> ACCESS_HIDDEN)) And _
        (!TableID = mobjComponent.ParentExpression.BaseTableID) Then

        listCalcCalculation.AddItem !Name
        listCalcCalculation.ItemData(listCalcCalculation.NewIndex) = !ExprID
      End If
      
      .MoveNext
    Loop
  End With
  
  ' Enable the combo if there are items.
  With listCalcCalculation
    If .ListCount > 0 Then
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With

  ' Set the flag to show that the listbox has been populated.
  mfCalculationsPopulated = True

TidyUpAndExit:
  UI.UnlockWindow
  listCalcCalculation.Refresh
  Exit Sub
  
ErrorTrap:
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  Resume TidyUpAndExit
  
End Sub
Private Sub listCalcFilters_Initialize()
  ' Populate the Calculation component - Calculation listbox.
  On Error GoTo ErrorTrap
  Dim lngExpressionID As Long
  
  lngExpressionID = mobjComponent.ParentExpression.ExpressionID
  
  ' Clear the current contents of the listbox.
  listCalcFilters.Clear
  
  UI.LockWindow Me.hWnd
  
  ' Add an item to the listbox for each filter.
  With recExprEdit
    .Index = "idxExprName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      ' Add items to the listbox for each expression that is not deleted or hidden by another user.

      'MH20000727 Added Email
      If (Not !ExprID = lngExpressionID) And _
        ((!Type = giEXPR_STATICFILTER)) And _
        (!ParentComponentID = 0) And _
        (Not !Deleted) And _
        ((!UserName = gsUserName) Or (!Access <> ACCESS_HIDDEN)) And _
        (!TableID = mobjComponent.ParentExpression.BaseTableID) Then

        listCalcFilters.AddItem !Name
        listCalcFilters.ItemData(listCalcFilters.NewIndex) = !ExprID
      End If
      
      .MoveNext
    Loop
  End With
  
  ' Enable the combo if there are items.
  With listCalcFilters
    If .ListCount > 0 Then
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With

  ' Set the flag to show that the listbox has been populated.
  mfFiltersPopulated = True

TidyUpAndExit:
  UI.UnlockWindow
  listCalcFilters.Refresh
  Exit Sub
  
ErrorTrap:
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  Err = False
  Resume TidyUpAndExit
  
End Sub



Private Sub asrFldSelLine_Change()
  ' Update the component object.
  mobjComponent.Component.SelectionLine = Val(asrFldSelLine.Text)

End Sub



Private Sub asrPValDefaultDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    asrPValDefaultDate.DateValue = Date
  End If

End Sub

Private Sub asrPValDefaultDate_LostFocus()

'  If IsNull(asrPValDefaultDate.DateValue) And Not _
'     IsDate(asrPValDefaultDate.DateValue) And _
'     asrPValDefaultDate.Text <> "  /  /" Then
'
'     MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     asrPValDefaultDate.DateValue = Null
'     asrPValDefaultDate.SetFocus
'     Exit Sub
'  End If
  
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate asrPValDefaultDate
  
End Sub


Private Sub asrPValReturnDecimals_Change()
  ' Update the component object with the new value.
  mobjComponent.Component.ReturnDecimals = Val(asrPValReturnDecimals.Text)
  
End Sub


Private Sub asrPValReturnSize_Change()
  ' Update the component object with the new value.
  mobjComponent.Component.ReturnSize = Val(asrPValReturnSize.Text)

End Sub


Private Sub asrValDateValue_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    asrValDateValue.DateValue = Date
  End If

End Sub

Private Sub asrValDateValue_LostFocus()

'  If IsNull(asrValDateValue.DateValue) And Not _
'     IsDate(asrValDateValue.DateValue) And _
'     asrValDateValue.Text <> "  /  /" Then
'
'     MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     asrValDateValue.DateValue = Null
'     asrValDateValue.SetFocus
'     Exit Sub
'  End If

  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate asrValDateValue

End Sub


Private Sub asrWorkflowFldSelLine_Change()
  ' Update the component object.
  mobjComponent.Component.SelectionLine = Val(asrWorkflowFldSelLine.Text)

End Sub

Private Sub cboFldTable_Click()
  ' Update the component object.
  mobjComponent.Component.TableID = cboFldTable.ItemData(cboFldTable.ListIndex)

  ' Populate the field combo with the relevant fields.
  cboFldColumn_Refresh
  fldSelOrder_Refresh
  fldSelFilter_Refresh

  FormatFieldControls
  
End Sub


Private Sub cboFldColumn_Click()
  ' Update the component object.
  mobjComponent.Component.ColumnID = cboFldColumn.ItemData(cboFldColumn.ListIndex)

  optFieldSel_Refresh

End Sub













Private Sub cboPValColumn_Click()
  Dim iLoop As Integer
  
  ' Update the component object with the new value.
  mobjComponent.Component.LookupColumn = cboPValColumn.ItemData(cboPValColumn.ListIndex)
  
  For iLoop = 1 To UBound(mavColumns, 2)
    If mavColumns(1, iLoop) = cboPValColumn.ItemData(cboPValColumn.ListIndex) Then
      mDataType = mavColumns(2, iLoop)
      Exit For
    End If
  Next iLoop
  
  ' Populate the default table value combo with the relevant fields.
  cboPValDefaultTabVal_Refresh

End Sub


Private Sub cboPValReturnType_Click()
  ' Update the component object with the new value.
  mobjComponent.Component.ValueType = cboPValReturnType.ItemData(cboPValReturnType.ListIndex)
  
  ' Display only the required controls.
  FormatPromptedValueControls

End Sub

Private Sub cboPValTable_Click()
  ' Populate the Column combo with the relevant fields.
  cboPValColumn_Refresh

End Sub


Private Sub cboTabValColumn_Click()
  ' Populate the value combo with the relevant fields.
  Dim iLoop As Integer
  
  For iLoop = 1 To UBound(mavColumns, 2)
    If mavColumns(1, iLoop) = cboTabValColumn.ItemData(cboTabValColumn.ListIndex) Then
      mDataType = mavColumns(2, iLoop)
      
      Select Case mavColumns(2, iLoop)
        Case dtNUMERIC
          mobjComponent.Component.ReturnType = giEXPRVALUE_NUMERIC
        Case dtINTEGER
          mobjComponent.Component.ReturnType = giEXPRVALUE_NUMERIC
        Case dtTIMESTAMP
          mobjComponent.Component.ReturnType = giEXPRVALUE_DATE
        Case dtVARCHAR, dtLONGVARCHAR
          mobjComponent.Component.ReturnType = giEXPRVALUE_CHARACTER
      End Select
      
      Exit For
    End If
  Next iLoop

  cboTabValValue_Refresh

End Sub


Private Sub cboTabValTable_Click()
  ' Populate the value combo with the relevant fields.
  cboTabValColumn_Refresh
  
End Sub


Private Sub cboTabValValue_Click()
  ' Update the component.
  With mobjComponent.Component
    Select Case .ReturnType
      Case giEXPRVALUE_NUMERIC
        .value = Val(cboTabValValue.List(cboTabValValue.ListIndex))
    
      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(cboTabValValue.List(cboTabValValue.ListIndex)) Then
          .value = CDate(cboTabValValue.List(cboTabValValue.ListIndex))
        Else
          .value = 0
        End If
  
      Case Else
        .value = cboTabValValue.List(cboTabValValue.ListIndex)
    End Select
  End With

End Sub

Private Sub cboValType_Click()
  ' Update the component object with the new value.
  mobjComponent.Component.ReturnType = cboValType.ItemData(cboValType.ListIndex)
    
  ' Display only the required controls.
  FormatValueControls

End Sub


Private Sub cboWFValueElement_Click()
  mobjComponent.Component.WorkflowElement = cboWFValueElement.List(cboWFValueElement.ListIndex)
  cboWFValueItem_Refresh

End Sub


Private Sub cboWFValueItem_Click()
  Dim asItems() As String
  Dim wfTemp As VB.Control
  Dim sItemIdentifier As String
  Dim iElementProperty As WorkflowElementProperties
  
  sItemIdentifier = ""
  iElementProperty = WORKFLOWELEMENTPROP_ITEMVALUE

  If cboWFValueElement.Enabled _
    And cboWFValueItem.Enabled Then
    
    Select Case cboWFValueItem.ItemData(cboWFValueItem.ListIndex)
      Case (WORKFLOWELEMENTPROP_COMPETIONCOUNT * -1)
        iElementProperty = WORKFLOWELEMENTPROP_COMPETIONCOUNT
        
      Case (WORKFLOWELEMENTPROP_FAILURECOUNT * -1)
        iElementProperty = WORKFLOWELEMENTPROP_FAILURECOUNT
        
      Case (WORKFLOWELEMENTPROP_TIMEOUTCOUNT * -1)
        iElementProperty = WORKFLOWELEMENTPROP_TIMEOUTCOUNT
        
      Case (WORKFLOWELEMENTPROP_MESSAGE * -1)
        iElementProperty = WORKFLOWELEMENTPROP_MESSAGE

      Case Else
        If cboWFValueItem.ItemData(cboWFValueItem.ListIndex) >= 0 Then
          Set wfTemp = maWFPrecedingElements(cboWFValueElement.ItemData(cboWFValueElement.ListIndex))
          asItems = wfTemp.Items
          Set wfTemp = Nothing
      
          sItemIdentifier = asItems(9, cboWFValueItem.ItemData(cboWFValueItem.ListIndex))
        End If
    End Select
  End If
  
  mobjComponent.Component.WorkflowItem = sItemIdentifier
  mobjComponent.Component.WorkflowElementProperty = iElementProperty

End Sub


Private Sub cboWorkflowFieldColumn_Click()
  ' Update the component object.
  mobjComponent.Component.ColumnID = cboWorkflowFieldColumn.ItemData(cboWorkflowFieldColumn.ListIndex)

  optWorkflowFieldSel_Refresh

End Sub


Private Sub cboWorkflowFieldElement_Click()
  Dim sElementIdentifier As String
  
  sElementIdentifier = ""
  
  If cboWorkflowFieldElement.ListCount > 0 Then
    If cboWorkflowFieldElement.ItemData(cboWorkflowFieldElement.ListIndex) > 0 Then
      sElementIdentifier = cboWorkflowFieldElement.List(cboWorkflowFieldElement.ListIndex)
    End If
  End If
  
  mobjComponent.Component.WorkflowElement = sElementIdentifier
  
  cboWorkflowFieldRecordSelector_Refresh

End Sub


Private Sub cboWorkflowFieldRecord_Click()
  ' Update the component object.
  mobjComponent.Component.RecordSelectionType = cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex)

  cboWorkflowFieldElement_Refresh
  cboWorkflowFieldRecordTable_Refresh

End Sub


Private Sub cboWorkflowFieldRecordTable_Refresh()
  ' Populate the Workflow Field Record Table combo and select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim fFound As Boolean
  Dim alngValidTables() As Long
  Dim lngBaseTable As Long
  Dim lngCurrentTable As Long
  Dim fTableOK As Boolean
  Dim lngTableID As Long
  Dim lngExcludedTableID As Long
  Dim sElementIdentifier As String
  Dim fRecordTableRequired As Boolean
  
  iIndex = -1
  iDefaultIndex = -1
  lngExcludedTableID = 0
  lngCurrentTable = mobjComponent.Component.RecordTableID
  fRecordTableRequired = False
  
  ' Clear the current contents of the combo.
  cboWorkflowFieldRecordTable.Clear

  If (cboWorkflowFieldTable.ListIndex >= 0) Then
    lngTableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)

    If cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD Then
      fRecordTableRequired = True
      sElementIdentifier = cboWorkflowFieldElement.List(cboWorkflowFieldElement.ListIndex)
      Set wfTemp = Nothing
      
      If Len(Trim(sElementIdentifier)) > 0 Then
        For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
          Set wfTemp = maWFPrecedingElements(lngLoop)

          If (UCase(Trim(wfTemp.Identifier)) = UCase(Trim(sElementIdentifier))) Then
            Exit For
          End If

          Set wfTemp = Nothing
        Next lngLoop
      End If

      If Not wfTemp Is Nothing Then
        If wfTemp.ElementType = elem_WebForm Then
          lngBaseTable = cboWorkflowFieldRecordSelector.ItemData(cboWorkflowFieldRecordSelector.ListIndex)
        ElseIf wfTemp.ElementType = elem_StoredData Then
          lngBaseTable = wfTemp.DataTableID
        End If
      End If
    ElseIf (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_INITIATOR) _
      Or (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_TRIGGEREDRECORD) Then
      
      fRecordTableRequired = True
      lngBaseTable = mobjComponent.ParentExpression.UtilityBaseTable
    End If

    If fRecordTableRequired Then
      ' Get an array of the valid table IDs (base table and it's ascendants)
      ReDim alngValidTables(0)
      TableAscendants lngBaseTable, alngValidTables
  
      With recTabEdit
        .Index = "idxName"
  
        If Not (.BOF And .EOF) Then
          .MoveFirst
        End If
  
        ' Add  an item to the combo for each table that has not been deleted.
        Do While Not .EOF
          fTableOK = (Not .Fields("deleted"))
  
          If fTableOK Then
            fFound = False
  
            For lngLoop = 1 To UBound(alngValidTables)
              If (alngValidTables(lngLoop) = !TableID) _
                And (lngExcludedTableID <> !TableID) Then
  
                fFound = IsChildOfTable(alngValidTables(lngLoop), lngTableID)
                If (Not fFound) And optWorkflowField(0).value Then
                  fFound = (alngValidTables(lngLoop) = lngTableID)
                End If
                
                Exit For
              End If
            Next lngLoop
  
            fTableOK = fFound
          End If
  
          If fTableOK Then
            cboWorkflowFieldRecordTable.AddItem !TableName
            cboWorkflowFieldRecordTable.ItemData(cboWorkflowFieldRecordTable.NewIndex) = !TableID
  
            If !TableID = lngCurrentTable Then
              iIndex = cboWorkflowFieldRecordTable.NewIndex
            End If
  
            If !TableID = lngBaseTable Then
              iDefaultIndex = cboWorkflowFieldRecordTable.NewIndex
            End If
          End If
  
          .MoveNext
        Loop
      End With
    End If
  End If

  cboWorkflowFieldRecordTable.Enabled = (cboWorkflowFieldRecordTable.ListCount > 0)

  If cboWorkflowFieldRecordTable.ListCount > 0 Then
    If iIndex < 0 Then
      If iDefaultIndex < 0 Then
        iIndex = 0
      Else
        iIndex = iDefaultIndex
      End If
    End If

    cboWorkflowFieldRecordTable.ListIndex = iIndex
  Else
    If fRecordTableRequired Then
      cboWorkflowFieldRecordTable.AddItem "<no values>"
      cboWorkflowFieldRecordTable.ItemData(cboWorkflowFieldRecordTable.NewIndex) = 0
      cboWorkflowFieldRecordTable.ListIndex = 0
    Else
      cboWorkflowFieldRecordTable_Click
    End If
  End If
    
End Sub



Private Sub cboWorkflowFunctionRecordTable_Refresh()
  ' Populate the Workflow Function Record Table combo and select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim fFound As Boolean
  Dim alngValidTables() As Long
  Dim lngBaseTable As Long
  Dim lngCurrentTable As Long
  Dim fTableOK As Boolean
  Dim lngExcludedTableID As Long
  Dim sElementIdentifier As String
  Dim lngRequiredTableID As Long
  Dim lngFunctionID As Long

  iIndex = -1
  iDefaultIndex = -1
  lngExcludedTableID = 0
  lngCurrentTable = mobjComponent.Component.WorkflowRecordTableID

  ' Clear the current contents of the combo.
  cboWorkflowFunctionRecordTable.Clear

  lngFunctionID = ssTreeFuncFunction_SelectedFunctionID
  lngRequiredTableID = FunctionRequiredTableID(lngFunctionID)

  If cboWorkflowFunctionRecord.ListCount > 0 Then
    If cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD _
      Or cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_INITIATOR _
      Or cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_TRIGGEREDRECORD Then
      
      sElementIdentifier = cboWorkflowFunctionElement.List(cboWorkflowFunctionElement.ListIndex)
      Set wfTemp = Nothing
  
      If Len(Trim(sElementIdentifier)) > 0 Then
        For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
          Set wfTemp = maWFPrecedingElements(lngLoop)
  
          If (UCase(Trim(wfTemp.Identifier)) = UCase(Trim(sElementIdentifier))) Then
            Exit For
          End If
  
          Set wfTemp = Nothing
        Next lngLoop
      End If
  
      If Not wfTemp Is Nothing Then
        If wfTemp.ElementType = elem_WebForm Then
          lngBaseTable = cboWorkflowFunctionRecordSelector.ItemData(cboWorkflowFunctionRecordSelector.ListIndex)
        ElseIf wfTemp.ElementType = elem_StoredData Then
          lngBaseTable = wfTemp.DataTableID
  
          'JPD 20061227
          If wfTemp.DataAction = DATAACTION_DELETE Then
            ' Exclude deleted records (but include their parent records)
            lngExcludedTableID = wfTemp.DataTableID
          End If
        End If
      Else
        lngBaseTable = mobjComponent.ParentExpression.UtilityBaseTable
      End If
  
      ' Get an array of the valid table IDs (base table and it's ascendants)
      ReDim alngValidTables(0)
      TableAscendants lngBaseTable, alngValidTables
  
      With recTabEdit
        .Index = "idxName"
  
        If Not (.BOF And .EOF) Then
          .MoveFirst
        End If
  
        ' Add  an item to the combo for each table that has not been deleted.
        Do While Not .EOF
          fTableOK = (Not .Fields("deleted"))
  
          If fTableOK Then
            fFound = False
  
            For lngLoop = 1 To UBound(alngValidTables)
              If (alngValidTables(lngLoop) = !TableID) _
                And (lngExcludedTableID <> !TableID) Then
  
                fFound = (alngValidTables(lngLoop) = lngRequiredTableID) _
                  Or (lngRequiredTableID < 0)
  
                Exit For
              End If
            Next lngLoop
  
            fTableOK = fFound
          End If
  
          If fTableOK Then
            cboWorkflowFunctionRecordTable.AddItem !TableName
            cboWorkflowFunctionRecordTable.ItemData(cboWorkflowFunctionRecordTable.NewIndex) = !TableID
  
            If !TableID = lngCurrentTable Then
              iIndex = cboWorkflowFunctionRecordTable.NewIndex
            End If
  
            If !TableID = lngBaseTable Then
              iDefaultIndex = cboWorkflowFunctionRecordTable.NewIndex
            End If
          End If
  
          .MoveNext
        Loop
      End With
    End If
  End If
  
  cboWorkflowFunctionRecordTable.Enabled = (cboWorkflowFunctionRecordTable.ListCount > 0)

  If cboWorkflowFunctionRecordTable.ListCount > 0 Then
    If iIndex < 0 Then
      If iDefaultIndex < 0 Then
        iIndex = 0
      Else
        iIndex = iDefaultIndex
      End If
    End If

    cboWorkflowFunctionRecordTable.ListIndex = iIndex
  Else
    cboWorkflowFunctionRecordTable_Click
  End If
    
End Sub




Private Sub cboWorkflowFieldRecordSelector_Click()
  Dim sItemIdentifier As String
  
  sItemIdentifier = ""
  
  If cboWorkflowFieldRecordSelector.ListCount > 0 Then
    If cboWorkflowFieldRecordSelector.ItemData(cboWorkflowFieldRecordSelector.ListIndex) > 0 Then
      sItemIdentifier = cboWorkflowFieldRecordSelector.List(cboWorkflowFieldRecordSelector.ListIndex)
    End If
  End If
  
  mobjComponent.Component.WorkflowItem = sItemIdentifier
  
  cboWorkflowFieldRecordTable_Refresh

End Sub

Private Sub cboWorkflowFieldRecordTable_Click()
  Dim lngTableID As Long
  
  lngTableID = 0
  
  If cboWorkflowFieldRecordTable.ListCount > 0 Then
    lngTableID = cboWorkflowFieldRecordTable.ItemData(cboWorkflowFieldRecordTable.ListIndex)
  End If
  
  mobjComponent.Component.RecordTableID = lngTableID
  
  ' Display only the required controls.
  FormatWorkflowFieldControls

End Sub


Private Sub cboWorkflowFieldTable_Click()
  ' Update the component object.
  mobjComponent.Component.TableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)

  ' Populate the field combo with the relevant fields.
  cboWorkflowFieldColumn_Refresh
  cboWorkflowFieldRecord_Refresh
  fldSelOrder_Refresh
  fldSelFilter_Refresh

  FormatWorkflowFieldControls

End Sub

Private Sub cboWorkflowFieldRecord_Refresh()
  ' Populate the Workflow Field Record combo and
  ' select the current value.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim fIdentifyingElement As Boolean
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim lngTableID As Long
  Dim alngValidTables() As Long
  Dim fFound As Boolean

  ' Get an array of the valid table IDs (base table and it's descendants)
  ReDim alngValidTables(0)

  lngTableID = -1
  If cboWorkflowFieldTable.ListIndex >= 0 Then
    lngTableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)
  End If

  With cboWorkflowFieldRecord
    ' Clear the current contents of the combo.
    .Clear

    For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
      fIdentifyingElement = False
      Set wfTemp = maWFPrecedingElements(lngLoop)

      If wfTemp.ElementType = elem_WebForm Then
        asItems = wfTemp.Items

        For lngLoop2 = 1 To UBound(asItems, 2)
          If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
            fFound = False

            ' Get an array of the valid table IDs (base table and it's ascendants)
            ReDim alngValidTables(0)
            TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables

            For lngLoop3 = 1 To UBound(alngValidTables)
              ' Dealing with a history aggregate (count or total)
              If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
                fFound = True
                Exit For
              ElseIf optWorkflowField(0).value Then
                If alngValidTables(lngLoop3) = lngTableID Then
                  fFound = True
                  Exit For
                End If
              End If
            Next lngLoop3

            If fFound Then
              fIdentifyingElement = True
              Exit For
            End If
          End If
        Next lngLoop2
      ElseIf wfTemp.ElementType = elem_StoredData Then
        fFound = False

        ReDim alngValidTables(0)
        TableAscendants wfTemp.DataTableID, alngValidTables

        For lngLoop3 = 1 To UBound(alngValidTables)
          If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
            fFound = True
            Exit For
          ElseIf optWorkflowField(0).value Then
            If alngValidTables(lngLoop3) = lngTableID Then
              fFound = True
              Exit For
            End If
          End If
        Next lngLoop3

        If fFound Then
          fIdentifyingElement = True
        End If
      End If

      If fIdentifyingElement Then
        Exit For
      End If

      Set wfTemp = Nothing
    Next lngLoop

    If fIdentifyingElement Then
      .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
      .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
    End If

    fFound = False
    If mobjComponent.ParentExpression.UtilityBaseTable > 0 Then
      ReDim alngValidTables(0)
      TableAscendants mobjComponent.ParentExpression.UtilityBaseTable, alngValidTables

      For lngLoop3 = 1 To UBound(alngValidTables)
        If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
          fFound = True
          Exit For
        ElseIf optWorkflowField(0).value Then
          If alngValidTables(lngLoop3) = lngTableID Then
            fFound = True
            Exit For
          End If
        End If
      Next lngLoop3
    End If

    If fFound Then
      If mobjComponent.ParentExpression.WorkflowInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
        .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
      End If

      If mobjComponent.ParentExpression.WorkflowInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
        .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD
      End If
    End If

    iIndex = -1
    iDefaultIndex = 0
    For lngLoop = 0 To .ListCount - 1
      If .ItemData(lngLoop) = mobjComponent.Component.RecordSelectionType Then
        iIndex = lngLoop
        Exit For
      End If

      If (.ItemData(lngLoop) = giWFRECSEL_INITIATOR) _
        Or (.ItemData(lngLoop) = giWFRECSEL_TRIGGEREDRECORD) Then
        iDefaultIndex = lngLoop
      End If
    Next lngLoop

    ' Enable the combo if there are items.
    If .ListCount > 0 Then
      .Enabled = True

      If iIndex < 0 Then
        iIndex = iDefaultIndex
      End If

      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no values>"
      .ItemData(.NewIndex) = giWFRECSEL_UNKNOWN
      .ListIndex = 0
    End If
  End With
    
End Sub



Private Sub cboWorkflowFunctionRecord_Refresh()
  ' Populate the Workflow Function Record combo and
  ' select the current value.
  ' NB. The Workflow record selection controls are only required for functions that are
  ' implicitly record based.
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim fIdentifyingElement As Boolean
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngFunctionID As Long
  Dim lngRequiredTableID As Long
  
  lngFunctionID = ssTreeFuncFunction_SelectedFunctionID
  lngRequiredTableID = FunctionRequiredTableID(lngFunctionID)
  
  ' Get an array of the valid table IDs (base table and it's descendants)
  ReDim alngValidTables(0)

  With cboWorkflowFunctionRecord
    ' Clear the current contents of the combo.
    .Clear

    If lngRequiredTableID <> 0 Then
      For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
        fIdentifyingElement = False
        Set wfTemp = maWFPrecedingElements(lngLoop)

        If wfTemp.ElementType = elem_WebForm Then
          asItems = wfTemp.Items

          For lngLoop2 = 1 To UBound(asItems, 2)
            If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
              fFound = False

              ' Get an array of the valid table IDs (base table and it's ascendants)
              ReDim alngValidTables(0)
              TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables

              For lngLoop3 = 1 To UBound(alngValidTables)
                If (alngValidTables(lngLoop3) = lngRequiredTableID) _
                  Or (lngRequiredTableID < 0) Then
                  
                  fFound = True
                  Exit For
                End If
              Next lngLoop3

              If fFound Then
                fIdentifyingElement = True
                Exit For
              End If
            End If
          Next lngLoop2
        ElseIf wfTemp.ElementType = elem_StoredData Then
          fFound = False

          ReDim alngValidTables(0)
          TableAscendants wfTemp.DataTableID, alngValidTables

          For lngLoop3 = 1 To UBound(alngValidTables)
            If (alngValidTables(lngLoop3) = lngRequiredTableID) _
              Or (lngRequiredTableID < 0) Then
              
              fFound = True
              Exit For
            End If
          Next lngLoop3

          If fFound Then
            fIdentifyingElement = True
          End If
        End If

        If fIdentifyingElement Then
          Exit For
        End If

        Set wfTemp = Nothing
      Next lngLoop
  
      If fIdentifyingElement Then
        .AddItem GetRecordSelectionDescription(giWFRECSEL_IDENTIFIEDRECORD)
        .ItemData(.NewIndex) = giWFRECSEL_IDENTIFIEDRECORD
      End If

      fFound = False
      If mobjComponent.ParentExpression.UtilityBaseTable > 0 Then
        ReDim alngValidTables(0)
        TableAscendants mobjComponent.ParentExpression.UtilityBaseTable, alngValidTables

        For lngLoop3 = 1 To UBound(alngValidTables)
          If (alngValidTables(lngLoop3) = lngRequiredTableID) _
            Or (lngRequiredTableID < 0) Then
            
            fFound = True
            Exit For
          End If
        Next lngLoop3
      End If

      If fFound Then
        If mobjComponent.ParentExpression.WorkflowInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
          .AddItem GetRecordSelectionDescription(giWFRECSEL_INITIATOR)
          .ItemData(.NewIndex) = giWFRECSEL_INITIATOR
        End If

        If mobjComponent.ParentExpression.WorkflowInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
          .AddItem GetRecordSelectionDescription(giWFRECSEL_TRIGGEREDRECORD)
          .ItemData(.NewIndex) = giWFRECSEL_TRIGGEREDRECORD
        End If
      End If

      If ((mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
        Or (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)) _
          And FunctionTableOK(lngFunctionID, mobjComponent.ParentExpression.BaseTableID) Then
        
        'JPD 20070515 Fault 12228
        ' 66 - Post that Reports to Current User
        ' 68 - Personnel that Reports to Current User
        ' 70 - Post that Current User Reports to
        ' 72 - Personnel that Current User Reports to
        ' 74 - Does Record Exist
        If lngFunctionID <> 66 _
          And lngFunctionID <> 68 _
          And lngFunctionID <> 70 _
          And lngFunctionID <> 72 _
          And lngFunctionID <> 74 Then
        
          .AddItem GetRecordSelectionDescription(giWFRECSEL_UNIDENTIFIED)
          .ItemData(.NewIndex) = giWFRECSEL_UNIDENTIFIED
        End If
      End If
      
      iIndex = -1
      iDefaultIndex = -1
      
      For lngLoop = 0 To .ListCount - 1
        If .ItemData(lngLoop) = mobjComponent.Component.WorkflowRecordSelectionType Then
          iIndex = lngLoop
          Exit For
        End If
        
        If (.ItemData(lngLoop) = giWFRECSEL_UNIDENTIFIED) Then
          iDefaultIndex = lngLoop
        End If
        
        If (iDefaultIndex < 0) _
          And ((.ItemData(lngLoop) = giWFRECSEL_INITIATOR) _
            Or (.ItemData(lngLoop) = giWFRECSEL_TRIGGEREDRECORD)) Then
            
          iDefaultIndex = lngLoop
        End If
      Next lngLoop

      ' Enable the combo if there are items.
      If .ListCount > 0 Then
        .Enabled = True
  
        If iIndex < 0 Then
          iIndex = IIf(iDefaultIndex < 0, 0, iDefaultIndex)
        End If
  
        .ListIndex = iIndex
      Else
        .Enabled = False
        .AddItem "<no values>"
        .ItemData(.NewIndex) = giWFRECSEL_UNKNOWN
        .ListIndex = 0

        cboWorkflowFunctionElement_Refresh
        cboWorkflowFunctionRecordTable_Refresh
      End If
    Else
      .Enabled = False
      
      mobjComponent.Component.WorkflowRecordSelectionType = giWFRECSEL_UNKNOWN
      mobjComponent.Component.WorkflowElement = ""
      mobjComponent.Component.WorkflowItem = ""
      mobjComponent.Component.WorkflowRecordTableID = 0
      
      cboWorkflowFunctionElement_Refresh
      cboWorkflowFunctionRecordTable_Refresh
    End If
  End With
    
End Sub




Private Sub cboWorkflowFunctionElement_Refresh()
  ' Populate the Workflow Function Element combo and select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim fIdentifyingElement As Boolean
  Dim asItems() As String
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngRequiredTableID As Long
  Dim sElementIdentifier As String
  Dim lngFunctionID As Long
  
  sElementIdentifier = mobjComponent.Component.WorkflowElement

  With cboWorkflowFunctionElement
    ' Clear the current contents of the combo.
    .Clear

    If cboWorkflowFunctionRecord.ListIndex >= 0 Then
      If (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then

        lngFunctionID = ssTreeFuncFunction_SelectedFunctionID
        lngRequiredTableID = FunctionRequiredTableID(lngFunctionID)

        For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
          fIdentifyingElement = False
          Set wfTemp = maWFPrecedingElements(lngLoop)

          If wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items

            For lngLoop2 = 1 To UBound(asItems, 2)
              If (asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID) Then
                ' Get an array of the valid table IDs (base table and it's descendants)
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables

                fFound = False
                For lngLoop3 = 1 To UBound(alngValidTables)
                  If (alngValidTables(lngLoop3) = lngRequiredTableID) _
                    Or (lngRequiredTableID < 0) Then
                    
                    fFound = True
                    Exit For
                  End If
                Next lngLoop3

                If fFound Then
                  fIdentifyingElement = True
                  Exit For
                End If
              End If
            Next lngLoop2

          ElseIf wfTemp.ElementType = elem_StoredData Then
            ' Get an array of the valid table IDs (base table and it's descendants)
            ReDim alngValidTables(0)
            TableAscendants wfTemp.DataTableID, alngValidTables

            'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
            If wfTemp.DataAction = DATAACTION_DELETE Then
              ' Cannot do anything with a Deleted record, but can use its ascendants.
              ' Remove the table itself from the array of valid tables.
              alngValidTables(1) = 0
            End If

            fFound = False
            For lngLoop3 = 1 To UBound(alngValidTables)
              If (alngValidTables(lngLoop3) = lngRequiredTableID) _
                Or (lngRequiredTableID < 0) Then
              
                fFound = True
                Exit For
              End If
            Next lngLoop3

            If fFound Then
              fIdentifyingElement = True
            End If
          End If

          If fIdentifyingElement Then
            .AddItem wfTemp.Identifier
            .ItemData(.NewIndex) = lngLoop
          End If

          Set wfTemp = Nothing
        Next lngLoop
      End If
    End If

    iIndex = -1
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = sElementIdentifier Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(sElementIdentifier)) > 0) Then
        sMsg = "The previously selected Function Element is no longer valid."

        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Function Element has been selected."
        End If

        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If

'''        mfForcedChanged = True
      End If

      iIndex = 0
    End If

    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = iIndex
    Else
      .Enabled = False
      
      If cboWorkflowFunctionRecord.ListIndex >= 0 Then
        If (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then

          .AddItem "<no elements>"
          .ItemData(.NewIndex) = 0
          .ListIndex = 0
        Else
          mobjComponent.Component.WorkflowElement = ""
          mobjComponent.Component.WorkflowItem = ""
        End If
      End If

      cboWorkflowFunctionRecordSelector_Refresh
    End If
  End With
    
End Sub





Private Sub cboWorkflowFieldElement_Refresh()
  ' Populate the Workflow Field Element combo and select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim fIdentifyingElement As Boolean
  Dim asItems() As String
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngTableID As Long
  Dim sElementIdentifier As String
  Dim fElementRequired As Boolean
  
  sElementIdentifier = mobjComponent.Component.WorkflowElement

  With cboWorkflowFieldElement
    ' Clear the current contents of the combo.
    .Clear

    fElementRequired = False
    If cboWorkflowFieldRecord.ListIndex >= 0 Then
      If (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then
        fElementRequired = True
      End If
    End If
    
    If fElementRequired Then
      lngTableID = -1
      If cboWorkflowFieldTable.ListIndex >= 0 Then
        lngTableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)
      End If

      For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as its for the current element
        fIdentifyingElement = False
        Set wfTemp = maWFPrecedingElements(lngLoop)

        If wfTemp.ElementType = elem_WebForm Then
          asItems = wfTemp.Items

          For lngLoop2 = 1 To UBound(asItems, 2)
            If (asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID) Then
              ' Get an array of the valid table IDs (base table and it's descendants)
              ReDim alngValidTables(0)
              TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables

              fFound = False
              For lngLoop3 = 1 To UBound(alngValidTables)
                If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
                  fFound = True
                  Exit For
                ElseIf optWorkflowField(0).value Then
                  If alngValidTables(lngLoop3) = lngTableID Then
                    fFound = True
                    Exit For
                  End If
                End If
              Next lngLoop3

              If fFound Then
                fIdentifyingElement = True
                Exit For
              End If
            End If
          Next lngLoop2

        ElseIf wfTemp.ElementType = elem_StoredData Then
          ' Get an array of the valid table IDs (base table and it's descendants)
          ReDim alngValidTables(0)
          TableAscendants wfTemp.DataTableID, alngValidTables

          fFound = False
          For lngLoop3 = 1 To UBound(alngValidTables)
            If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
              fFound = True
              Exit For
            ElseIf optWorkflowField(0).value Then
              If alngValidTables(lngLoop3) = lngTableID Then
                fFound = True
                Exit For
              End If
            End If
          Next lngLoop3

          If fFound Then
            fIdentifyingElement = True
          End If
        End If

        If fIdentifyingElement Then
          .AddItem wfTemp.Identifier
          .ItemData(.NewIndex) = lngLoop
        End If

        Set wfTemp = Nothing
      Next lngLoop

      iIndex = -1
      For lngLoop = 0 To .ListCount - 1
        If .List(lngLoop) = sElementIdentifier Then
          iIndex = lngLoop
          Exit For
        End If
      Next lngLoop

      If (iIndex < 0) Then
        If (Len(Trim(sElementIdentifier)) > 0) Then
          sMsg = "The previously selected Workflow Identified Field Element is no longer valid."
  
          If .ListCount > 0 Then
            sMsg = sMsg & vbCrLf & "A default Workflow Identified Field Element has been selected."
          End If
  
          If mfInitializing Then
            If Len(msInitializeMessage) = 0 Then
              msInitializeMessage = sMsg
            End If
          End If
  
  '''        mfForcedChanged = True
        End If

        iIndex = 0
      End If
    End If

    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = iIndex
    Else
      .Enabled = False
      
      If fElementRequired Then
        .AddItem "<no elements>"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
      Else
        mobjComponent.Component.WorkflowElement = ""
        mobjComponent.Component.WorkflowItem = ""
        
        cboWorkflowFieldRecordSelector_Refresh
      End If
    End If
  End With
    
End Sub






Private Sub cboWorkflowFunctionElement_Click()
  Dim sElementIdentifier As String
  
  sElementIdentifier = ""
  
  If cboWorkflowFunctionElement.ListCount > 0 Then
    If cboWorkflowFunctionElement.ItemData(cboWorkflowFunctionElement.ListIndex) > 0 Then
      sElementIdentifier = cboWorkflowFunctionElement.List(cboWorkflowFunctionElement.ListIndex)
    End If
  End If
  
  mobjComponent.Component.WorkflowElement = sElementIdentifier
  
  cboWorkflowFunctionRecordSelector_Refresh

End Sub


Private Sub cboWorkflowFunctionRecord_Click()
  ' Update the component object.
  mobjComponent.Component.WorkflowRecordSelectionType = cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex)

  cboWorkflowFunctionElement_Refresh
  cboWorkflowFunctionRecordTable_Refresh

End Sub


Private Sub cboWorkflowFunctionRecordSelector_Click()
  Dim sItemIdentifier As String
  
  sItemIdentifier = ""
  
  If cboWorkflowFunctionRecordSelector.ListCount > 0 Then
    If cboWorkflowFunctionRecordSelector.ItemData(cboWorkflowFunctionRecordSelector.ListIndex) > 0 Then
      sItemIdentifier = cboWorkflowFunctionRecordSelector.List(cboWorkflowFunctionRecordSelector.ListIndex)
    End If
  End If
  
  mobjComponent.Component.WorkflowItem = sItemIdentifier
  
  cboWorkflowFunctionRecordTable_Refresh

End Sub


Private Sub cboWorkflowFunctionRecordTable_Click()
  Dim lngTableID As Long
  
  lngTableID = 0
  
  If cboWorkflowFunctionRecordTable.ListCount > 0 Then
    lngTableID = cboWorkflowFunctionRecordTable.ItemData(cboWorkflowFunctionRecordTable.ListIndex)
  End If
  
  mobjComponent.Component.WorkflowRecordTableID = lngTableID
  
  ' Display only the required controls.
  FormatFunctionControls

End Sub


Private Sub cmdCancel_Click()
  ' Set the cancelled flag.
  mfCancelled = True
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdFldSelFilter_Click()
  ' Display the 'Field Selection Filter' expression selection form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    
    'MH20000727 Added Email
    If (mobjComponent.ParentExpression.ExpressionType = giEXPR_COLUMNCALCULATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDVALIDATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDDESCRIPTION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_OUTLOOKFOLDER) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_OUTLOOKSUBJECT) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_STATICFILTER) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_EMAIL) Or _
      ((mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) And (mobjComponent.ComponentType <> giCOMPONENT_WORKFLOWFIELD)) Then
      
      .Initialise _
        mobjComponent.Component.TableID, _
        mobjComponent.Component.SelectionFilter, _
        giEXPR_STATICFILTER, _
        giEXPRVALUE_LOGIC
      
    ElseIf (mobjComponent.ParentExpression.ExpressionType = giEXPR_LINKFILTER) Then
    
      .Initialise _
        mobjComponent.Component.TableID, _
        mobjComponent.Component.SelectionFilter, _
        giEXPR_LINKFILTER, _
        giEXPRVALUE_LOGIC
      
    ElseIf (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
      ((mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) And (mobjComponent.ComponentType = giCOMPONENT_WORKFLOWFIELD)) Then
      
      .Initialise _
        mobjComponent.Component.TableID, _
        mobjComponent.Component.SelectionFilter, _
        giEXPR_WORKFLOWSTATICFILTER, _
        giEXPRVALUE_LOGIC
      .UtilityID = mobjComponent.ParentExpression.UtilityID
      .UtilityBaseTable = mobjComponent.ParentExpression.UtilityBaseTable
      .WorkflowInitiationType = mobjComponent.ParentExpression.WorkflowInitiationType
      .PrecedingWorkflowElements = mobjComponent.ParentExpression.PrecedingWorkflowElements
      .AllWorkflowElements = mobjComponent.ParentExpression.AllWorkflowElements
      
    ElseIf (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
      
      .Initialise _
        mobjComponent.Component.TableID, _
        mobjComponent.Component.SelectionFilter, _
        giEXPR_WORKFLOWRUNTIMEFILTER, _
        giEXPRVALUE_LOGIC
      .UtilityID = mobjComponent.ParentExpression.UtilityID
      .UtilityBaseTable = mobjComponent.ParentExpression.UtilityBaseTable
      .WorkflowInitiationType = mobjComponent.ParentExpression.WorkflowInitiationType
      .PrecedingWorkflowElements = mobjComponent.ParentExpression.PrecedingWorkflowElements
      .AllWorkflowElements = mobjComponent.ParentExpression.AllWorkflowElements
      
    Else
      .Initialise _
        mobjComponent.Component.TableID, _
        mobjComponent.Component.SelectionFilter, _
        giEXPR_RUNTIMEFILTER, _
        giEXPRVALUE_LOGIC
    End If
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      mobjComponent.Component.SelectionFilter = .ExpressionID
        
      ' Read the selected expression info.
      GetFldSelFilterDetails
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", mobjComponent.Component.SelectionFilter, False

        If .NoMatch Then
          ' Read the selected expression info.
          mobjComponent.Component.SelectionFilter = 0
          GetFldSelFilterDetails
        End If
      End With
    End If
  End With
  
TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing expression ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Function GetFldSelFilterDetails() As Boolean
  ' Get the 'Field Selection Filter' expression details.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sExprName As String
  Dim objExpr As CExpression
  
  fOK = True
  
  ' Initialise the default values.
  sExprName = ""
    
  ' Instantiate the expression class.
  Set objExpr = New CExpression
    
  With objExpr
    ' Set the expression id.
    .ExpressionID = mobjComponent.Component.SelectionFilter
      
    ' Read the required info from the expression.
    If .ReadExpressionDetails Then
      sExprName = .Name
    End If
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objExpr = Nothing
  If Not fOK Then
    sExprName = ""
  End If
  
  ' Update the controls properties.
  If mobjComponent.ComponentType = giCOMPONENT_WORKFLOWFIELD Then
    txtWorkflowFldSelFilter.Text = sExprName
  Else
    txtFldSelFilter.Text = sExprName
  End If
  GetFldSelFilterDetails = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function GetFldSelOrderDetails() As Boolean
  ' Get the 'Field Selection Order' details.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sOrderName As String
  Dim objOrder As Order
  
  fOK = True
  
  ' Initialise the default values.
  sOrderName = ""
    
  ' Instantiate a new Order object.
  Set objOrder = New Order
  With objOrder
    .OrderID = mobjComponent.Component.SelectionOrderID
    
    ' Read the name of the current order.
    If .ConstructOrder Then
      sOrderName = .OrderName
    End If
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objOrder = Nothing
  If Not fOK Then
    sOrderName = ""
  End If
  
  ' Update the controls properties.
  If mobjComponent.ComponentType = giCOMPONENT_WORKFLOWFIELD Then
    txtWorkflowFldSelOrder.Text = sOrderName
  Else
    txtFldSelOrder.Text = sOrderName
  End If
  
  GetFldSelOrderDetails = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Sub cmdFldSelOrder_Click()
  ' Display the 'Field Selection Order' selection form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objOrder As Order

  fOK = True
  
  ' Instantiate an expression object.
  Set objOrder = New Order
  
  With objOrder
  
    ' Initialize the order object.
    .OrderID = mobjComponent.Component.SelectionOrderID
    .TableID = mobjComponent.Component.TableID
    
    'MH20000727 Added Email
    If (mobjComponent.ParentExpression.ExpressionType = giEXPR_COLUMNCALCULATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDDESCRIPTION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_OUTLOOKFOLDER) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_OUTLOOKSUBJECT) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDVALIDATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_STATICFILTER) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_EMAIL) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
      .OrderType = giORDERTYPE_STATIC
    Else
      .OrderType = giORDERTYPE_DYNAMIC
    End If
    
    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      mobjComponent.Component.SelectionOrderID = .OrderID
        
      ' Read the selected order info.
      GetFldSelOrderDetails
    Else
      ' Check in case the original expression has been deleted.
      With recOrdEdit
        .Index = "idxID"
        .Seek "=", mobjComponent.Component.SelectionOrderID

        If .NoMatch Then
          ' Read the selected expression info.
          mobjComponent.Component.SelectionOrderID = 0
          GetFldSelOrderDetails
        End If
      End With
    End If
  End With
  
TidyUpAndExit:
  Set objOrder = Nothing
  If Not fOK Then
    MsgBox "Error changing order ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdOK_Click()
  
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  ''If ValidateGTMaskDate(asrValDateValue) = False Or _
  ''   ValidateGTMaskDate(asrPValDefaultDate) = False Then
  ''    Exit Sub
  ''End If
  cmdOK.SetFocus
  DoEvents

  ' Write the displayed control values to the component.
  If SaveComponent Then
    ' Set the cancelled flag.
    mfCancelled = False
    
    ' Unload the form.
    UnLoad Me
  End If
  
End Sub



Private Sub cmdWorkflowFldSelFilter_Click()
  cmdFldSelFilter_Click

End Sub

Private Sub cmdWorkflowFldSelOrder_Click()
  cmdFldSelOrder_Click
  
  ' Display only the required controls.
  FormatWorkflowFieldControls
  
End Sub

Private Sub Form_Initialize()
  ' Initialize the 'cancelled' property.
  mfCancelled = True
  mfFunctionsPopulated = False
  mfCalculationsPopulated = False
  mfOperatorsPopulated = False
  mfTabValTablesPopulated = False
  mfPValTablesPopulated = False
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  ' JDM - 15/02/01 - Fault 1869 - Error when pressing CTRL-X on treeview control
  ' For some reason the Sheridan treeview control wants to fire off it own cut'n'paste functionality
  ' must trap it here not in it's own keydown event
  If ActiveControl.Name = "ssTreeFuncFunction" Or ActiveControl.Name = "ssTreeOpOperator" Then
    KeyCode = 0
    Shift = 0
  End If

End Sub


Private Sub Form_Load()
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  ' Format the form's frames and controls
'  SetDateComboFormat Me.asrValDateValue
  'SetDateComboFormat Me.asrPValDefaultDate
  
  'JPD 20041115 Fault 8970
  UI.FormatGTDateControl asrValDateValue
  UI.FormatGTDateControl asrPValDefaultDate
  
  ReadModuleParameters
  
  FormatScreen

End Sub



Public Property Get Cancelled() As Boolean
  ' Return the Cancelled property.
  Cancelled = mfCancelled
  
End Property

Public Property Set Component(pobjComponent As CExprComponent)
  ' Set the component property.
  mfInitializing = True
  msInitializeMessage = ""
  ReDim maWFPrecedingElements(0)
  
  Set mobjComponent = pobjComponent
  
  ' Set the component type.
  miComponentType = mobjComponent.ComponentType
  mfFieldByValue = (mobjComponent.ParentExpression.ReturnType < 100)
  If mobjComponent.ComponentType = giCOMPONENT_FIELD Then
    mobjComponent.Component.FieldPassType = IIf(mfFieldByValue, giPASSBY_VALUE, giPASSBY_REFERENCE)
  End If
  
  ' Format the controls within the frames.
  If mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION _
    Or mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER _
    Or mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER Then
    
    maWFPrecedingElements = mobjComponent.ParentExpression.PrecedingWorkflowElements
  End If
  
  If (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
    And mfFieldByValue Then
    
    FormatWorkflowFieldFrame
  Else
    FormatFieldFrame
  End If
  
  ' Format the Component Type frame for the new component.
  FormatComponentTypeFrame
  
  ' Format the Component frame for the new component.
  If optComponentType(miComponentType).value Then
    DisplayComponentFrame
  Else
    optComponentType(miComponentType).value = True
  End If
  
  If Len(msInitializeMessage) > 0 Then
    MsgBox msInitializeMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  mfInitializing = False
  
End Property








Private Sub DisplayComponentFrame()
  Dim iLoop As Integer

  ' Initialize the displayed controls.
  InitializeComponentControls
  
  ' Display only the frame that defines the selected component type.
  For iLoop = fraComponent.LBound To fraComponent.UBound
    fraComponent(iLoop).Visible = (fraComponent(iLoop).Tag = miComponentType)
  Next iLoop
  
End Sub

Private Sub FormatScreen()
  ' Position and size controls.
  Dim iLoop As Integer
  Dim iFRAMEHEIGHT As Single
  Dim iCOMPONENTFRAMEWIDTH As Single
  
  Const iXGAP = 200
  Const iYGAP = 200
  Const iXFRAMEGAP = 150
  Const iYFRAMEGAP = 100
  Const iFRAMEWIDTH = 5400
  
  Const iFRAMEHEIGHT_DEFAULT = 3900
  Const iFRAMEHEIGHT_WORKFLOW = 5500
  
  Const iCOMPONENTFRAMEWIDTH_DEFAULT = 2300
  Const iCOMPONENTFRAMEWIDTH_WORKFLOW = 3000
  
  iFRAMEHEIGHT = IIf((mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
    Or (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
    Or (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER), _
    iFRAMEHEIGHT_WORKFLOW, _
    iFRAMEHEIGHT_DEFAULT)
  
  iCOMPONENTFRAMEWIDTH = IIf((mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
    Or (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) _
    Or (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER), _
    iCOMPONENTFRAMEWIDTH_WORKFLOW, _
    iCOMPONENTFRAMEWIDTH_DEFAULT)
  
  ' Position and size the component type frame.
  With fraComponentType
    .Left = iXFRAMEGAP
    .Top = iYFRAMEGAP
    .Width = iCOMPONENTFRAMEWIDTH
    .Height = iFRAMEHEIGHT
  End With
  
  ' Position and size the component definition frames.
  For iLoop = fraComponent.LBound To fraComponent.UBound
    With fraComponent(iLoop)
      .Left = fraComponentType.Left + iCOMPONENTFRAMEWIDTH + iXFRAMEGAP
      .Top = iYFRAMEGAP
      .Width = iFRAMEWIDTH
      .Height = iFRAMEHEIGHT
    End With
  Next iLoop
  
  ' Format the controls within the frames.
  FormatFunctionFrame
  FormatCalculationFrame
  FormatOperatorFrame
  FormatValueFrame
  FormatTableValueFrame
  FormatPromptedValueFrame
  FormatFilterFrame
  FormatWorkflowFieldFrame
  FormatWorkflowValueFrame
  
  ' Position and size the OK/Cancel command controls.
  With cmdCancel
    .Top = iYFRAMEGAP + iYGAP + iFRAMEHEIGHT
    .Left = fraComponent(fraComponent.LBound).Left + _
      fraComponent(fraComponent.LBound).Width - .Width
    cmdOK.Top = .Top
    cmdOK.Left = .Left - iXGAP - cmdOK.Width
  End With
  
  ' Size the form.
  Me.Width = fraComponent(fraComponent.UBound).Left + _
    fraComponent(fraComponent.UBound).Width + iXFRAMEGAP + _
    (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
  Me.Height = cmdOK.Top + cmdOK.Height + iXFRAMEGAP + _
    (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    mfCancelled = True
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Terminate()
'Private Sub Form_Unload(Cancel As Integer)
  ' Disassociate object variables.
  Set mobjComponent = Nothing

End Sub


Private Sub listCalcCalculation_Click()
  ' Update the component value.
  mobjComponent.Component.CalculationID = listCalcCalculation.ItemData(listCalcCalculation.ListIndex)

End Sub

Private Sub listCalcCalculation_DblClick()
  ' Confirm the selection.
  If cmdOK.Enabled Then
    cmdOK_Click
  End If
  
End Sub













Private Sub listCalcFilters_Click()

  mobjComponent.Component.FilterID = listCalcFilters.ItemData(listCalcFilters.ListIndex)


End Sub

Private Sub optComponentType_Click(piIndex As Integer)
  ' Set the component type property.
  miComponentType = piIndex
  mobjComponent.ComponentType = piIndex
  
  DisplayComponentFrame
  
End Sub



Private Sub InitializeValueControls()
  Dim sCharacterValue As String
  Dim dblNumericValue As Double
  Dim fLogicValue As Boolean
  Dim dDateValue As Date

  ' Select the current value type.
  cboValType_Refresh
  
  ' Initialise the value controls as required.
  sCharacterValue = vbNullString
  dblNumericValue = 0
  fLogicValue = True
  dDateValue = Date
  
  With mobjComponent.Component
    Select Case .ReturnType
      Case giEXPRVALUE_CHARACTER
        sCharacterValue = .value
      Case giEXPRVALUE_NUMERIC
        dblNumericValue = .value
      Case giEXPRVALUE_LOGIC
        fLogicValue = .value
      Case giEXPRVALUE_DATE
        dDateValue = .value
    End Select
  End With
  
  txtValCharacterValue.Text = sCharacterValue
  TDBValNumericValue.value = dblNumericValue
  optValLogicValue(0).value = fLogicValue
  optValLogicValue(1).value = Not optValLogicValue(0).value
'  asrValDateValue.Value = dDateValue
  asrValDateValue.Text = dDateValue
  
  ' Ensure the user can confirm the component definition.
  cmdOK.Enabled = True

End Sub
Private Sub InitializeWorkflowValueControls()
  ' Initialize the Workflow Value Component controls.
  
  ' Populate the Element combo if it is not already populated.
  cboWFValueElement_Refresh
  
End Sub
Private Sub cboWFValueElement_Refresh()
  ' Populate the WorkflowValue WebForm combo and
  ' select the current webform if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim sMsg As String
  Dim sElementIdentifier As String
  
  iIndex = -1
  sElementIdentifier = mobjComponent.Component.WorkflowElement
  
  ' Clear the current contents of the combo.
  cboWFValueElement.Clear

  ' Add  an item to the combo for each preceding web form.
  ' Ignore the first item, as it will be the current web form.
  For iLoop = 2 To UBound(maWFPrecedingElements)
    If (maWFPrecedingElements(iLoop).ElementType = elem_WebForm) _
      Or (maWFPrecedingElements(iLoop).ElementType = elem_StoredData) Then
      cboWFValueElement.AddItem maWFPrecedingElements(iLoop).Identifier
      cboWFValueElement.ItemData(cboWFValueElement.NewIndex) = iLoop
    End If
  Next iLoop

  For iLoop = 0 To cboWFValueElement.ListCount - 1
    If cboWFValueElement.List(iLoop) = sElementIdentifier Then
      iIndex = iLoop
    End If
  Next iLoop

  If (iIndex < 0) Then
    If (Len(Trim(sElementIdentifier)) > 0) Then
      sMsg = "The previously selected Workflow Value Element is no longer valid."

      If cboWFValueElement.ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Workflow Value Element has been selected."
      End If

      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If

'''      mfForcedChanged = True
    End If

    iIndex = 0
  End If

  ' Enable the combo if there are items.
  With cboWFValueElement
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no preceding elements>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0

      cboWFValueItem_Refresh
    End If
  End With

End Sub


Private Sub cboWFValueItem_Refresh()
  ' Populate the cboWFValueItem combo and
  ' select the current value if it is still valid.
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim fItemOK As Boolean
  Dim sMsg As String
  Dim sItemIdentifier As String
  Dim sItemDescription As String
  Dim sDataTypeDescription As String
  Dim objWFValue As CExprWorkflowValue
  
  iIndex = -1
  sItemIdentifier = mobjComponent.Component.WorkflowItem
  
  ' Clear the current contents of the combo.
  cboWFValueItem.Clear

  If cboWFValueElement.Enabled Then
    ' Add  an item to the combo for each input item in the preceding web form.
    Set wfTemp = maWFPrecedingElements(cboWFValueElement.ItemData(cboWFValueElement.ListIndex))

    Set objWFValue = New CExprWorkflowValue
    
    cboWFValueItem.AddItem objWFValue.ElementPropertyDescription(WORKFLOWELEMENTPROP_COMPETIONCOUNT)
    cboWFValueItem.ItemData(cboWFValueItem.NewIndex) = (WORKFLOWELEMENTPROP_COMPETIONCOUNT * -1)

    If wfTemp.ElementType = elem_StoredData Then
      cboWFValueItem.AddItem objWFValue.ElementPropertyDescription(WORKFLOWELEMENTPROP_FAILURECOUNT)
      cboWFValueItem.ItemData(cboWFValueItem.NewIndex) = (WORKFLOWELEMENTPROP_FAILURECOUNT * -1)
    End If
    
    If wfTemp.ElementType = elem_WebForm Then
      cboWFValueItem.AddItem objWFValue.ElementPropertyDescription(WORKFLOWELEMENTPROP_TIMEOUTCOUNT)
      cboWFValueItem.ItemData(cboWFValueItem.NewIndex) = (WORKFLOWELEMENTPROP_TIMEOUTCOUNT * -1)
    End If
    
    If (wfTemp.ElementType = elem_StoredData) _
      Or (wfTemp.ElementType = elem_Email) Then
      
      cboWFValueItem.AddItem objWFValue.ElementPropertyDescription(WORKFLOWELEMENTPROP_MESSAGE)
      cboWFValueItem.ItemData(cboWFValueItem.NewIndex) = (WORKFLOWELEMENTPROP_MESSAGE * -1)
    End If

    Set objWFValue = Nothing
    
    If wfTemp.ElementType = elem_WebForm Then
      asItems = wfTemp.Items
  
      Set wfTemp = Nothing
      
      For iLoop = 1 To UBound(asItems, 2)
        
        Select Case asItems(2, iLoop)
          Case giWFFORMITEM_BUTTON
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (button pressed)"
            
          Case giWFFORMITEM_INPUTVALUE_CHAR
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (character input)"
            
          Case giWFFORMITEM_INPUTVALUE_NUMERIC
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (numeric input)"
            
          Case giWFFORMITEM_INPUTVALUE_LOGIC
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (logic input)"
            
          Case giWFFORMITEM_INPUTVALUE_DATE
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (date input)"
            
          Case giWFFORMITEM_INPUTVALUE_DROPDOWN
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (dropdown character input)"
            
          Case giWFFORMITEM_INPUTVALUE_LOOKUP
            fItemOK = True
            
            Select Case GetColumnDataType(CLng(asItems(49, iLoop)))
              Case dtLONGVARCHAR
                sDataTypeDescription = "working pattern"
              Case dtNUMERIC
                sDataTypeDescription = "numeric"
              Case dtINTEGER
                sDataTypeDescription = "integer"
              Case dtTIMESTAMP
                sDataTypeDescription = "date"
              Case dtVARCHAR
                sDataTypeDescription = "character"
              Case Else
                fItemOK = False
            End Select
            
            sItemDescription = asItems(9, iLoop) & " (lookup " & sDataTypeDescription & " input)"
            
          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
            fItemOK = True
            sItemDescription = asItems(9, iLoop) & " (option group character input)"
              
          Case Else
            fItemOK = False
          
        End Select
        
        If fItemOK Then
          cboWFValueItem.AddItem sItemDescription
          cboWFValueItem.ItemData(cboWFValueItem.NewIndex) = iLoop
        End If
      Next iLoop
    End If
    
    For iLoop = 0 To cboWFValueItem.ListCount - 1
      Select Case mobjComponent.Component.WorkflowElementProperty
        Case WORKFLOWELEMENTPROP_COMPETIONCOUNT, _
          WORKFLOWELEMENTPROP_FAILURECOUNT, _
          WORKFLOWELEMENTPROP_TIMEOUTCOUNT, _
          WORKFLOWELEMENTPROP_MESSAGE
          
          If cboWFValueItem.ItemData(iLoop) = (mobjComponent.Component.WorkflowElementProperty * -1) Then
            iIndex = iLoop
            Exit For
          End If
        
        Case Else
          If cboWFValueItem.ItemData(iLoop) >= 0 Then
            If asItems(9, cboWFValueItem.ItemData(iLoop)) = sItemIdentifier Then
              iIndex = iLoop
              Exit For
            End If
          End If
      End Select
    Next iLoop
  End If

  If (iIndex < 0) Then
    If (mobjComponent.Component.WorkflowElementProperty = WORKFLOWELEMENTPROP_ITEMVALUE) And _
      (Len(Trim(sItemIdentifier)) > 0) Then
      
      sMsg = "The previously selected Workflow Value is no longer valid."

      If cboWFValueItem.ListCount > 0 Then
        sMsg = sMsg & vbCrLf & "A default Workflow Value has been selected."
      End If

      If mfInitializing Then
        If Len(msInitializeMessage) = 0 Then
          msInitializeMessage = sMsg
        End If
      End If

'''      mfForcedChanged = True
    End If

    iIndex = 0
  End If

  ' Enable the combo if there are items.
  With cboWFValueItem
    cmdOK.Enabled = (.ListCount > 0)
    
    If .ListCount > 0 Then
      .Enabled = True
      If iIndex < 0 Then
        iIndex = 0
      End If
      .ListIndex = iIndex
    Else
      .Enabled = False

      .AddItem "<no values>"
      .ItemData(.NewIndex) = -1
      .ListIndex = 0
    End If
  End With
    
End Sub


Private Sub cboWorkflowFieldRecordSelector_Refresh()
  ' Populate the Workflow Field RecordSelector combo and select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngTableID As Long
  Dim sItemIdentifier As String
  Dim fItemRequired As Boolean
  
  fItemRequired = False
  sItemIdentifier = mobjComponent.Component.WorkflowItem

  With cboWorkflowFieldRecordSelector
    ' Clear the current contents of the combo.
    .Clear

    If cboWorkflowFieldElement.ListIndex >= 0 Then
      lngTableID = -1
      If cboWorkflowFieldTable.ListIndex >= 0 Then
        lngTableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)
      End If

      If cboWorkflowFieldElement.ItemData(cboWorkflowFieldElement.ListIndex) > 0 Then
        Set wfTemp = maWFPrecedingElements(cboWorkflowFieldElement.ItemData(cboWorkflowFieldElement.ListIndex))

        If wfTemp.ElementType = elem_WebForm Then
          asItems = wfTemp.Items
          fItemRequired = True

          For lngLoop2 = 1 To UBound(asItems, 2)
            If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then

              ' Get an array of the valid table IDs (base table and it's descendants)
              ReDim alngValidTables(0)
              TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables

              fFound = False
              For lngLoop3 = 1 To UBound(alngValidTables)
                If IsChildOfTable(alngValidTables(lngLoop3), lngTableID) Then
                  fFound = True
                  Exit For
                ElseIf optWorkflowField(0).value Then
                  If alngValidTables(lngLoop3) = lngTableID Then
                    fFound = True
                    Exit For
                  End If
                End If
              Next lngLoop3

              If fFound Then
                .AddItem asItems(9, lngLoop2)
                .ItemData(.NewIndex) = asItems(44, lngLoop2)
              End If
            End If
          Next lngLoop2
        End If

        Set wfTemp = Nothing
      End If
    End If

    iIndex = -1
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = sItemIdentifier Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(sItemIdentifier)) > 0) Then
        sMsg = "The previously selected Workflow Identified Field Record Selector is no longer valid."

        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Workflow Identified Field Record Selector has been selected."
        End If

        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If

'''        mfForcedChanged = True
      End If

      iIndex = 0
    End If

    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = iIndex
    Else
      .Enabled = False
      
      If fItemRequired Then
        .AddItem "<no values>"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
      Else
        mobjComponent.Component.WorkflowItem = ""
      
        cboWorkflowFieldRecordTable_Refresh
      End If
    End If
  End With

End Sub





Private Sub cboWorkflowFunctionRecordSelector_Refresh()
  ' Populate the Workflow Function RecordSelector combo and select the current value.
  Dim iIndex As Integer
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim lngLoop3 As Long
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim iElementType As ElementType
  Dim sMsg As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngRequiredTableID As Long
  Dim lngFunctionID As Long
  Dim sItemIdentifier As String

  sItemIdentifier = mobjComponent.Component.WorkflowItem

  With cboWorkflowFunctionRecordSelector
    ' Clear the current contents of the combo.
    .Clear

    If cboWorkflowFunctionElement.ListIndex >= 0 Then
      lngFunctionID = ssTreeFuncFunction_SelectedFunctionID
      lngRequiredTableID = FunctionRequiredTableID(lngFunctionID)

      If cboWorkflowFunctionElement.ItemData(cboWorkflowFunctionElement.ListIndex) > 0 Then
        Set wfTemp = maWFPrecedingElements(cboWorkflowFunctionElement.ItemData(cboWorkflowFunctionElement.ListIndex))

        iElementType = wfTemp.ElementType

        If iElementType = elem_WebForm Then
          asItems = wfTemp.Items

          For lngLoop2 = 1 To UBound(asItems, 2)
            If asItems(2, lngLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then

              ' Get an array of the valid table IDs (base table and it's descendants)
              ReDim alngValidTables(0)
              TableAscendants CLng(asItems(44, lngLoop2)), alngValidTables

              fFound = False
              For lngLoop3 = 1 To UBound(alngValidTables)
                If (alngValidTables(lngLoop3) = lngRequiredTableID) _
                  Or (lngRequiredTableID < 0) Then
                
                  fFound = True
                  Exit For
                End If
              Next lngLoop3

              If fFound Then
                .AddItem asItems(9, lngLoop2)
                .ItemData(.NewIndex) = asItems(44, lngLoop2)
              End If
            End If
          Next lngLoop2
        End If

        Set wfTemp = Nothing
      End If
    End If

    iIndex = -1
    For lngLoop = 0 To .ListCount - 1
      If .List(lngLoop) = sItemIdentifier Then
        iIndex = lngLoop
        Exit For
      End If
    Next lngLoop

    If (iIndex < 0) Then
      If (Len(Trim(sItemIdentifier)) > 0) Then
        sMsg = "The previously selected Workflow Function Record Selector is no longer valid."

        If .ListCount > 0 Then
          sMsg = sMsg & vbCrLf & "A default Workflow Function Record Selector has been selected."
        End If

        If mfInitializing Then
          If Len(msInitializeMessage) = 0 Then
            msInitializeMessage = sMsg
          End If
        End If

'''        mfForcedChanged = True
      End If

      iIndex = 0
    End If

    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = iIndex
    Else
      .Enabled = False
      
      If cboWorkflowFunctionRecord.ListIndex >= 0 Then
        If (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then
          If iElementType = elem_WebForm Then
            .AddItem "<no values>"
            .ItemData(.NewIndex) = 0
            .ListIndex = 0
          End If
        Else
          mobjComponent.Component.WorkflowItem = ""
        End If
      End If

    End If
  End With

'''  RefreshScreen

End Sub






Private Sub InitializeWorkflowFieldControls()
  ' Initialize the Workflow Field Component controls.
  Dim objWFFieldComponent As CExprWorkflowField

  Set objWFFieldComponent = mobjComponent.Component

  optWorkflowField_Refresh

  ' Select the current record line number value.
  asrWorkflowFldSelLine.Text = Trim(Str(objWFFieldComponent.SelectionLine))

  ' Disassociate object variables.
  Set objWFFieldComponent = Nothing
  
End Sub

Private Sub InitializeFieldControls()
  ' Initialize the Field Component controls.
  Dim objFieldComponent As CExprField
  
  Set objFieldComponent = mobjComponent.Component
  
  optField_Refresh
    
  ' Select the current record line number value.
  asrFldSelLine.Text = Trim(Str(objFieldComponent.SelectionLine))
  
  ' Disassociate object variables.
  Set objFieldComponent = Nothing
  
End Sub

Private Sub InitializeFunctionControls()
  ' Initialize the Function Component controls.
  On Error GoTo ErrorTrap
  
  Dim sFunctionKey As String
      
  ' Populate the function listbox if it is not already populated.
  If Not mfFunctionsPopulated Then
    ssTreeFuncFunction_Initialize
    ssTreeFuncFunction.Nodes("FUNCTION_ROOT").Expanded = True
  End If
  
  ' Select the current Function in the treeview.
  sFunctionKey = Trim(Str(mobjComponent.Component.FunctionID))
  ssTreeFuncFunction.SelectedItem = ssTreeFuncFunction.Nodes(sFunctionKey)
  cboWorkflowFunctionRecord_Refresh
  Exit Sub
  
ErrorTrap:
  If ssTreeFuncFunction.Nodes.Count > 0 Then
    ssTreeFuncFunction.SelectedItem = ssTreeFuncFunction.Nodes(1)
  End If
  
  cmdOK.Enabled = False
  
End Sub
Private Sub InitializeOperatorControls()
  ' Initialize the Operator Component controls.
  On Error GoTo ErrorTrap
  
  Dim sOperatorKey As String
  
  ' Populate the operator listbox if it is not already populated.
  If Not mfOperatorsPopulated Then
    ssTreeOpOperator_Initialize
    ssTreeOpOperator.Nodes("OPERATOR_ROOT").Expanded = True
  End If
  
  ' Select the current Operator in the treeview.
  sOperatorKey = Trim(Str(mobjComponent.Component.OperatorID))
  ssTreeOpOperator.SelectedItem = ssTreeOpOperator.Nodes(sOperatorKey)
  
  Exit Sub
  
ErrorTrap:
  If ssTreeOpOperator.Nodes.Count > 0 Then
    ssTreeOpOperator.SelectedItem = ssTreeOpOperator.Nodes(1)
  End If
  
  cmdOK.Enabled = False
      
End Sub
Private Sub InitializeTableValueControls()
  ' Initialize the Table Value Component controls.

  Dim iCount As Integer
  Dim iTableID, iColumnID As Integer
  Dim vValue As Variant
  Dim bValueFound As Boolean
  Dim objMisc As Misc

  Set objMisc = New Misc

  ' Save the status of the component
  iTableID = mobjComponent.Component.TableID
  iColumnID = mobjComponent.Component.ColumnID
  vValue = mobjComponent.Component.value
  
  ' Populate the Table Value Combo if it is not already populated.
  If Not mfTabValTablesPopulated Then
    cboTabValTable_Initialize
  End If

  ' Only allow the user to confirm the component definition if a valid
  ' table value is selected.
  cmdOK.Enabled = cboTabValValue.Enabled

  ' Initialize controls.
  If cboTabValTable.Enabled Then
  
    ' Set Lookup Table
    For iCount = 0 To cboTabValTable.ListCount - 1
        If cboTabValTable.ItemData(iCount) = iTableID Then
            cboTabValTable.ListIndex = iCount
        End If
    Next iCount

    ' Set Lookup column
    For iCount = 0 To cboTabValColumn.ListCount - 1
        If cboTabValColumn.ItemData(iCount) = iColumnID Then
            cboTabValColumn.ListIndex = iCount
        End If
    Next iCount

    ' Set lookup value
    bValueFound = False

    ' Format the data so it can be found in the dropdown combo
    With mobjComponent.Component
        Select Case .ReturnType
            Case giEXPRVALUE_LOGIC
                vValue = IIf(vValue = True, "True", "False")
            Case giEXPRVALUE_DATE
              'JPD 20041115 Fault 8970
              vValue = Format(vValue, objMisc.DateFormat)
              'vValue = FormatDateTime(vValue, vbLongDate)
        End Select
    End With

    For iCount = 0 To cboTabValValue.ListCount - 1
        If cboTabValValue.List(iCount) = vValue Then
            'cboTabValValue.Text = vValue
            cboTabValValue.ListIndex = iCount
            bValueFound = True
        End If
    Next iCount

    ' They must have modified/deleted this lookup table entry
    If Not bValueFound Then

        If Not vValue = "" Then
            cboTabValValue.AddItem vValue
            cboTabValValue.Text = vValue
            lblLookupValNotFound.Visible = True
            lblLookupValNotFound.Caption = Trim(vValue) & " does not appear in " + Trim(cboTabValTable.Text) + "." + cboTabValColumn.Text
            lblLookupValNotFound.Width = fraComponent(5).Width - 300
        End If
    End If

  End If
  
  Set objMisc = Nothing

End Sub
Private Sub InitializePromptedValueControls()
  ' Initialize the Prompted Value component controls.
  Dim sDefaultCharacter As String
  Dim dblDefaultNumeric As Double
  Dim fDefaultLogic As Boolean
  Dim dtDefaultDate As Date

  With mobjComponent.Component
    ' Initialise the prompt text box.
    txtPValPrompt.Text = .Prompt
    
    ' Select the current return value type.
    cboPValReturnType_Refresh
    
    ' Initialise the return size and decimals controls.
    asrPValReturnSize.Text = Trim(Str(.ReturnSize))
    asrPValReturnDecimals.Text = Trim(Str(.ReturnDecimals))
  
    ' Initialise the mask text box.
    txtPValFormat.Text = .ValueFormat
  
    ' Populate the Table combo if it not already populated.
    If Not mfPValTablesPopulated Then
      cboPValTable_Initialize
    End If
  
    ' Select the current table in the combo.
    cboPValTable_Refresh
    
    sDefaultCharacter = vbNullString
    dblDefaultNumeric = 0
    fDefaultLogic = True
    dtDefaultDate = Date
    
    Select Case .ValueType
      Case giEXPRVALUE_CHARACTER ' Character
        sDefaultCharacter = .DefaultValue
      Case giEXPRVALUE_NUMERIC ' Numeric
        dblDefaultNumeric = .DefaultValue
      Case giEXPRVALUE_LOGIC ' Logic
        fDefaultLogic = .DefaultValue
      Case giEXPRVALUE_DATE ' Date
        dtDefaultDate = .DefaultValue
    End Select
    
    txtPValDefaultCharacter.Text = sDefaultCharacter
    TDBPValDefaultNumeric.value = dblDefaultNumeric
    optPValDefaultLogic(0).value = fDefaultLogic
    optPValDefaultLogic(1).value = Not optPValDefaultLogic(0).value
    
    If CDbl(dtDefaultDate) <> 0 Then
'      asrPValDefaultDate.Value = dtDefaultDate
      asrPValDefaultDate.Text = dtDefaultDate
    End If
  End With
  
End Sub

Private Sub InitializeComponentControls()

  ' Call the required sub-routine to initialze the component
  ' definition controls.
  Select Case miComponentType
    Case giCOMPONENT_FIELD
      InitializeFieldControls
    
    Case giCOMPONENT_FUNCTION
      InitializeFunctionControls
      FormatFunctionControls
     
    Case giCOMPONENT_CALCULATION
      InitializeCalcControls
    
    Case giCOMPONENT_VALUE
      InitializeValueControls
      
    Case giCOMPONENT_OPERATOR
      InitializeOperatorControls
    
    Case giCOMPONENT_TABLEVALUE
      InitializeTableValueControls
    
    Case giCOMPONENT_PROMPTEDVALUE
      InitializePromptedValueControls
    
    Case giCOMPONENT_CUSTOMCALC
      ' Not required.
      
    Case giCOMPONENT_EXPRESSION
      ' Not handled in this form.

    Case giCOMPONENT_FILTER
        InitializeFilterControls

    Case giCOMPONENT_WORKFLOWFIELD
        InitializeWorkflowFieldControls

    Case giCOMPONENT_WORKFLOWVALUE
        InitializeWorkflowValueControls

  End Select
  
End Sub
Private Function SaveComponent() As Boolean
  ' Call the required sub-routine to save the
  ' control values to the component.
  ' NB. most changes are written directly to the component from
  ' the change events of the controls themselves. We do not do this
  ' for Operators and Functions as updating the OperatorID/FuctionID
  ' property of the component results in the Operator/Function definition
  ' being read from the database. This takes time, so we only do this when
  ' the Operator/Function selection is confirmed.
  SaveComponent = True
  
  Select Case miComponentType
    Case giCOMPONENT_FIELD
      SaveComponent = SaveField
    
    Case giCOMPONENT_FUNCTION
      SaveComponent = SaveFunction
         
    Case giCOMPONENT_OPERATOR
      SaveComponent = SaveOperator
        
    Case giCOMPONENT_VALUE
      SaveComponent = SaveValue
        
    Case giCOMPONENT_PROMPTEDVALUE
      SaveComponent = SavePromptedValue

    Case giCOMPONENT_TABLEVALUE
      SaveComponent = SaveTableValue

    Case giCOMPONENT_WORKFLOWVALUE
      SaveComponent = SaveWorkflowValue

    Case giCOMPONENT_WORKFLOWFIELD
      SaveComponent = SaveWorkflowField
  
  End Select
  
End Function
Private Function SaveWorkflowValue() As Boolean
  ' No validation required.
  Dim fSaveOK As Boolean

  fSaveOK = True
  
  SaveWorkflowValue = fSaveOK
  
End Function

Private Function SaveTableValue() As Boolean

  ' Populate the value combo with the relevant fields.
  Dim iLoop As Integer
  Dim fSaveOK As Boolean

  fSaveOK = True

  ' Set the component table id
  mobjComponent.Component.TableID = cboTabValTable.ItemData(cboTabValTable.ListIndex)

  ' Set the component column id
  mobjComponent.Component.ColumnID = cboTabValColumn.ItemData(cboTabValColumn.ListIndex)

  For iLoop = 1 To UBound(mavColumns, 2)
    If mavColumns(1, iLoop) = cboTabValColumn.ItemData(cboTabValColumn.ListIndex) Then
      mDataType = mavColumns(2, iLoop)
      
      Select Case mavColumns(2, iLoop)
        Case sqlNumeric
          mobjComponent.Component.ReturnType = giEXPRVALUE_NUMERIC
        Case sqlInteger
          mobjComponent.Component.ReturnType = giEXPRVALUE_NUMERIC
        Case sqlDate
          mobjComponent.Component.ReturnType = giEXPRVALUE_DATE
        Case sqlVarChar, sqlLongVarChar
          mobjComponent.Component.ReturnType = giEXPRVALUE_CHARACTER
      End Select
      
      Exit For
    End If
  Next iLoop

  ' Update the component.
  With mobjComponent.Component
    Select Case .ReturnType
      Case giEXPRVALUE_NUMERIC
        .value = Val(cboTabValValue.List(cboTabValValue.ListIndex))
    
      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(cboTabValValue.List(cboTabValValue.ListIndex)) Then
          .value = CDate(cboTabValValue.List(cboTabValValue.ListIndex))
        Else
          .value = 0
        End If
  
      Case Else
        .value = cboTabValValue.List(cboTabValValue.ListIndex)
    End Select
  End With

  SaveTableValue = fSaveOK
  
End Function


Private Function SaveField() As Boolean
  ' Validate the field component definition.
  
  ' Ensure that an order is selected for references to child fields.
  'SaveField = (Not cmdFldSelOrder.Enabled) Or _
    (mobjComponent.Component.SelectionOrderID > 0) Or _
    (mobjComponent.Component.SelectionType = giSELECT_RECORDTOTAL) Or _
    (mobjComponent.Component.SelectionType = giSELECT_RECORDCOUNT)
  
  'MH20010201 Fault 1608
  'If pass by reference then the field selection option isn't even visible
  'so don't show an error forcing an order
  SaveField = (Not mfFieldByValue) Or _
    (Not cmdFldSelOrder.Enabled) Or _
    (mobjComponent.Component.SelectionOrderID > 0) Or _
    (mobjComponent.Component.SelectionType = giSELECT_RECORDTOTAL) Or _
    (mobjComponent.Component.SelectionType = giSELECT_RECORDCOUNT)
  
  If Not SaveField Then
    MsgBox "An order must be specified when referring to child fields.", vbExclamation + vbOKOnly, App.ProductName
  End If
  
End Function

Private Function SaveWorkflowField() As Boolean
  ' Validate the Workflow field component definition.

  ' Ensure that an order is selected for references to child fields.
  SaveWorkflowField = (Not cmdWorkflowFldSelOrder.Enabled) Or _
    (mobjComponent.Component.SelectionOrderID > 0) Or _
    (mobjComponent.Component.SelectionType = giSELECT_RECORDTOTAL) Or _
    (mobjComponent.Component.SelectionType = giSELECT_RECORDCOUNT)

  If Not SaveWorkflowField Then
    MsgBox "An order must be specified when referring to child fields.", vbExclamation + vbOKOnly, App.ProductName
  End If
  
End Function


Private Sub InitializeCalcControls()
  ' Initialize the Calculation Component controls.
  
  ' Populate the Calculation listbox if it is not already populated.
  If Not mfCalculationsPopulated Then
    listCalcCalculation_Initialize
  End If
  
  ' Select the current calculation in the list box.
  listCalcCalculation_Refresh
  
  ' Only allow the user to confirm the component definition if a valid
  ' calculation is selected.
  cmdOK.Enabled = listCalcCalculation.Enabled
    
End Sub


Private Sub InitializeFilterControls()
  ' Initialize the Filter Component controls.
  
  ' Populate the Calculation listbox if it is not already populated.
  If Not mfFiltersPopulated Then
    listCalcFilters_Initialize
  End If
  
  ' Select the current calculation in the list box.
  listCalcFilters_Refresh
  
  ' Only allow the user to confirm the component definition if a valid
  ' calculation is selected.
  cmdOK.Enabled = listCalcFilters.Enabled
    
End Sub

Private Sub FormatValueControls()
  ' Display only the required Value Component controls.
  
  txtValCharacterValue.Visible = False
  TDBValNumericValue.Visible = False
  optValLogicValue(0).Visible = False
  optValLogicValue(1).Visible = False
  asrValDateValue.Visible = False
  
  ' Conditionally display some controls.
  Select Case cboValType.ItemData(cboValType.ListIndex)
    Case giEXPRVALUE_CHARACTER
      txtValCharacterValue.Visible = True
      
    Case giEXPRVALUE_NUMERIC
      TDBValNumericValue.Visible = True
      UI.FormatTDBNumberControl TDBValNumericValue

    Case giEXPRVALUE_LOGIC
      optValLogicValue(0).Visible = True
      optValLogicValue(1).Visible = True
  
    Case giEXPRVALUE_DATE
      asrValDateValue.Visible = True
  
      'MH20001003 Fault 1048
      'After making the greentree control visible you will
      'be unable to use the arrow keys to scroll though the
      'combo box items.  Setting focus to the form seems to
      'fixed this !
      If Me.Visible Then Me.SetFocus
  
  End Select

End Sub
Private Sub FormatPromptedValueControls()
  ' Display only the required Prompted Value Component controls.
  Dim fSizeVisible As Boolean
  Dim fDecimalsVisible As Boolean
  Dim fFormatVisible As Boolean
  Dim fFormatEnabled As Boolean
  Dim fLookupTableEnabled As Boolean
  
  Const YGAP = 75
  
  fSizeVisible = False
  fDecimalsVisible = False
  fFormatVisible = False
  fFormatEnabled = False
  
  txtPValDefaultCharacter.Visible = False
  TDBPValDefaultNumeric.Visible = False
  optPValDefaultLogic(0).Visible = False
  optPValDefaultLogic(1).Visible = False
  asrPValDefaultDate.Visible = False
  cboPValDefaultTabVal.Visible = False
  
  ' Conditionally display some controls.
  Select Case cboPValReturnType.ItemData(cboPValReturnType.ListIndex)
    Case giEXPRVALUE_CHARACTER ' Character
      fSizeVisible = True
      fFormatEnabled = True
      fFormatVisible = True
      txtPValDefaultCharacter.Visible = True
      
    Case giEXPRVALUE_NUMERIC ' Numeric
      fSizeVisible = True
      fDecimalsVisible = True
      fFormatVisible = True
      TDBPValDefaultNumeric.Visible = True
      UI.FormatTDBNumberControl TDBPValDefaultNumeric
    
    Case giEXPRVALUE_LOGIC ' Logic
      optPValDefaultLogic(0).Visible = True
      optPValDefaultLogic(1).Visible = True
      fFormatVisible = True
  
    Case giEXPRVALUE_DATE ' Date
      asrPValDefaultDate.Visible = True
      fFormatVisible = True
    
    Case giEXPRVALUE_TABLEVALUE ' Table Value
      fLookupTableEnabled = True
      cboPValDefaultTabVal.Visible = True
  End Select

  ' Display the Return Size controls if required.
  lblPValSize.Visible = fSizeVisible
  asrPValReturnSize.Visible = fSizeVisible
  
  ' Display the Return Decimals controls if required.
  lblPValDecimals.Visible = fDecimalsVisible
  asrPValReturnDecimals.Visible = fDecimalsVisible
  
  ' Display and enable the Format frame and controls if required.
  fraPValFormat.Visible = fFormatVisible
  fraPValFormat.Enabled = fFormatEnabled
  If (Not fFormatVisible) Or _
    (Not fFormatEnabled) Then
    txtPValFormat.Text = ""
  End If
  
  txtPValFormat.Enabled = fFormatEnabled
  txtPValFormat.BackColor = IIf(fFormatEnabled, vbWhite, &H8000000F)
  lblMaskKey1.Enabled = fFormatEnabled
  lblMaskKey2.Enabled = fFormatEnabled
  lblMaskKey3.Enabled = fFormatEnabled
  lblMaskKey4.Enabled = fFormatEnabled
  lblMaskKey5.Enabled = fFormatEnabled
  lblMaskKey6.Enabled = fFormatEnabled
  
  ' Display the Table frame and controls if required.
  fraPValTable.Visible = fLookupTableEnabled
  
  If fFormatVisible Then
    fraPValDefaultValue.Top = fraPValFormat.Top + fraPValFormat.Height + YGAP
  End If
  
  If fraPValTable.Visible Then
    fraPValDefaultValue.Top = fraPValTable.Top + fraPValTable.Height + YGAP
  End If

End Sub

Private Sub FormatWorkflowValueFrame()
  ' Size and position the Workflow Value component controls.
  Dim lngYCoordinate As Long

  Const lngCOLUMN1 = 200
  Const lngCOLUMN2 = 1200
  Const lngYGAP = 400
  Const lngCONTROLWIDTH = 4000

  lngYCoordinate = 300

  ' Format the Workflow Value - Element controls.
  With cboWFValueElement
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWFValueElement.Left = lngCOLUMN1
    lblWFValueElement.Top = lngYCoordinate + ((.Height - lblWFValueElement.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Value - Item controls.
  With cboWFValueItem
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWFValueItem.Left = lngCOLUMN1
    lblWFValueItem.Top = lngYCoordinate + ((.Height - lblWFValueItem.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
End Sub

Private Sub FormatTableValueFrame()
  ' Size and position the Value component controls.
  Dim lngYCoordinate As Long
  
  Const lngCOLUMN1 = 200
  Const lngCOLUMN2 = 1000
  Const lngYGAP = 400
  Const lngCONTROLWIDTH = 3700
  
  lngYCoordinate = 300
  
  ' Format the Table Value - Table controls.
  With cboTabValTable
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblTabValTable.Left = lngCOLUMN1
    lblTabValTable.Top = lngYCoordinate + ((.Height - lblTabValTable.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Table Value - Column controls.
  With cboTabValColumn
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblTabValColumn.Left = lngCOLUMN1
    lblTabValColumn.Top = lngYCoordinate + ((.Height - lblTabValColumn.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Table Value - Value controls.
  With cboTabValValue
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblTabValValue.Left = lngCOLUMN1
    lblTabValValue.Top = lngYCoordinate + ((.Height - lblTabValValue.Height) / 2)
  End With
  
End Sub


Private Sub FormatWorkflowFieldControls()
  ' Display only the required Workflow Field Component controls.
  Dim fIsChildOfBase As Boolean
  Dim fEnabled As Boolean
  Dim lngBaseTableID As Long
  Dim lngTableID As Long
  Dim wfTemp As VB.Control
  Dim fOK As Boolean
  
  ' Disable the column combo if 'COUNT' is selected.
  fEnabled = Not optWorkflowField(1).value
  If cboWorkflowFieldColumn.Enabled Then
    cboWorkflowFieldColumn.Enabled = fEnabled
  End If
  cboWorkflowFieldColumn.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
  lblWorkflowFieldColumn.Enabled = fEnabled

  ' Disable the element combo if its not required.
  fEnabled = (cboWorkflowFieldRecord.ListIndex >= 0)
  If fEnabled Then
    fEnabled = (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD)
  End If
  If cboWorkflowFieldElement.Enabled Then
    cboWorkflowFieldElement.Enabled = fEnabled
  End If
  cboWorkflowFieldElement.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
  lblWorkflowFieldElement.Enabled = fEnabled
  
  ' Disable the record selector combo if its not required.
  fEnabled = (cboWorkflowFieldRecord.ListIndex >= 0) _
    And (cboWorkflowFieldElement.ListIndex >= 0)
  If fEnabled Then
    fEnabled = (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
      And (cboWorkflowFieldElement.ItemData(cboWorkflowFieldElement.ListIndex) > 0)
    
    If fEnabled Then
      Set wfTemp = GetElementByIdentifier(cboWorkflowFieldElement.List(cboWorkflowFieldElement.ListIndex))
      fEnabled = Not wfTemp Is Nothing
      If fEnabled Then
        fEnabled = (wfTemp.ElementType = elem_WebForm)
      End If
    End If
  End If
  If cboWorkflowFieldRecordSelector.Enabled Then
    cboWorkflowFieldRecordSelector.Enabled = fEnabled
  End If
  cboWorkflowFieldRecordSelector.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
  lblWorkflowFieldRecordSelector.Enabled = fEnabled
  
  ' Disable the record table combo if its not required.
  fEnabled = (cboWorkflowFieldRecord.ListIndex >= 0)
  If fEnabled Then
    fEnabled = (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
      Or (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_TRIGGEREDRECORD) _
      Or (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_INITIATOR)
  End If
  If cboWorkflowFieldRecordTable.Enabled Then
    cboWorkflowFieldRecordTable.Enabled = fEnabled
  End If
  cboWorkflowFieldRecordTable.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
  lblWorkflowFieldRecordTable.Enabled = fEnabled
  
  ' Only enable the field selection options if the selected table
  ' is a child of the base database.
  lngTableID = 0
  If cboWorkflowFieldTable.ListIndex >= 0 Then
    lngTableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)
  End If
  
  lngBaseTableID = 0
  If cboWorkflowFieldRecordTable.ListIndex >= 0 Then
    lngBaseTableID = cboWorkflowFieldRecordTable.ItemData(cboWorkflowFieldRecordTable.ListIndex)
  End If
  fIsChildOfBase = IsChildOfTable(lngBaseTableID, lngTableID)
    
  fraWorkflowFldSelOptions.Enabled = fIsChildOfBase

  optWorkflowFieldSel(0).Enabled = (fIsChildOfBase And optWorkflowField(0).value)
  optWorkflowFieldSel(1).Enabled = (fIsChildOfBase And optWorkflowField(0).value)
  optWorkflowFieldSel(2).Enabled = (fIsChildOfBase And optWorkflowField(0).value)
  asrWorkflowFldSelLine.Enabled = optWorkflowFieldSel(2).Enabled

  If (Not optWorkflowFieldSel(1).Enabled) And (optWorkflowFieldSel(1).value) Then
    optWorkflowFieldSel(0).value = True
  End If
  If (Not optWorkflowFieldSel(2).Enabled) And (optWorkflowFieldSel(2).value) Then
    optWorkflowFieldSel(0).value = True
  End If

  If (Not optWorkflowFieldSel(2).value) Then
    asrWorkflowFldSelLine.Text = vbNullString
  End If

  ' Only enable the line number control if required.
  asrWorkflowFldSelLine.Enabled = (fIsChildOfBase And optWorkflowFieldSel(2).value)
  asrWorkflowFldSelLine.BackColor = IIf(asrWorkflowFldSelLine.Enabled, vbWhite, vbButtonFace)

  lblWorkflowFldOrder.Enabled = (fIsChildOfBase And optWorkflowField(0).value)
  cmdWorkflowFldSelOrder.Enabled = (fIsChildOfBase And optWorkflowField(0).value)
  If Not cmdWorkflowFldSelOrder.Enabled Then
    mobjComponent.Component.SelectionOrderID = 0
    txtWorkflowFldSelOrder.Text = ""
  End If

  lblWorkflowFldFilter.Enabled = fIsChildOfBase
  cmdWorkflowFldSelFilter.Enabled = fIsChildOfBase
  If Not cmdWorkflowFldSelFilter.Enabled Then
    mobjComponent.Component.SelectionFilter = 0
    txtWorkflowFldSelFilter.Text = ""
  End If
  
  ' -----------------------------------------------
  ' Check if it's OK to enable the OK button
  ' -----------------------------------------------
  
  ' OK only if a valid table is selected
  lngTableID = 0
  fOK = (cboWorkflowFieldTable.ListCount > 0)
  If fOK Then
    fOK = (cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex) > 0)
    lngTableID = cboWorkflowFieldTable.ItemData(cboWorkflowFieldTable.ListIndex)
  End If
  
  ' OK only if we want the COUNT of records or a valid column is selected
  If fOK And (Not optWorkflowField(1).value) Then
    fOK = (cboWorkflowFieldColumn.ListCount > 0)
    
    If fOK Then
      fOK = (cboWorkflowFieldColumn.ItemData(cboWorkflowFieldColumn.ListIndex) > 0)
    End If
  End If
  
  ' OK only if a valid record is selected
  If fOK Then
    fOK = (cboWorkflowFieldRecord.ListIndex >= 0)
    
    If fOK Then
      fOK = (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
        Or (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_TRIGGEREDRECORD) _
        Or (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_INITIATOR)
    End If
  End If
  
  ' If we are using an 'identified' record...
  If fOK Then
    If (cboWorkflowFieldRecord.ItemData(cboWorkflowFieldRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then
      ' OK only if a valid element is selected
      fOK = (cboWorkflowFieldElement.ListIndex >= 0)
        
      If fOK Then
        fOK = (cboWorkflowFieldElement.ItemData(cboWorkflowFieldElement.ListIndex) > 0)
      End If
    
      If fOK Then
        Set wfTemp = GetElementByIdentifier(cboWorkflowFieldElement.List(cboWorkflowFieldElement.ListIndex))
        fOK = Not wfTemp Is Nothing
        
        If fOK And (wfTemp.ElementType = elem_WebForm) Then
          ' OK only if a valid record selector is selected
          fOK = (cboWorkflowFieldRecordSelector.ListIndex >= 0)
        
          If fOK Then
            fOK = (cboWorkflowFieldRecordSelector.ItemData(cboWorkflowFieldRecordSelector.ListIndex) > 0)
          End If
        End If
      End If
    End If
  End If
  
  ' OK only if a valid record table is selected
  lngBaseTableID = 0
  If fOK Then
    fOK = (cboWorkflowFieldRecordTable.ListIndex >= 0)
    
    If fOK Then
      fOK = (cboWorkflowFieldRecordTable.ItemData(cboWorkflowFieldRecordTable.ListIndex) > 0)
      lngBaseTableID = cboWorkflowFieldRecordTable.ItemData(cboWorkflowFieldRecordTable.ListIndex)
    End If
  End If
  
  'JPD 20070323 Fault 11975
  ' If we're dealing with a child table...
  'If fOK Then
  '  If IsChildOfTable(lngBaseTableID, lngTableID) And optWorkflowField(0).Value Then
  '    ' If we're dealing with a 'field' from a child table...
  '    ' OK only if a valid order is selected
  '    fOK = (mobjComponent.Component.SelectionOrderID > 0)
  '  End If
  'End If

  cmdOK.Enabled = fOK
  
End Sub
Private Sub FormatFunctionControls()
  ' Enable the Workflow controls if required.
  Dim fFunctionSelected As Boolean
  Dim lngFunctionID As Long
  Dim lngRequiredTableID As Long
  Dim fEnabled As Boolean
  Dim wfTemp As VB.Control
  Dim fOK As Boolean
  
  lngFunctionID = ssTreeFuncFunction_SelectedFunctionID
  fFunctionSelected = (lngFunctionID > 0)
  lngRequiredTableID = FunctionRequiredTableID(lngFunctionID)
  
  fEnabled = (lngRequiredTableID <> 0)
  fraWorkflowFunctionRecord.Enabled = fEnabled
  
  ' Disable the record combo if its not required.
  fOK = True
  
  If (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) Then
    
    If cboWorkflowFunctionRecord.Enabled Then
      cboWorkflowFunctionRecord.Enabled = fEnabled
    End If
    cboWorkflowFunctionRecord.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
    lblWorkflowFunctionRecord.Enabled = fEnabled
    
    fEnabled = (cboWorkflowFunctionRecord.ListIndex >= 0)
    If fEnabled Then
      fEnabled = (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
        Or (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_TRIGGEREDRECORD) _
        Or (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_INITIATOR)
    End If
    If cboWorkflowFunctionRecordTable.Enabled Then
      cboWorkflowFunctionRecordTable.Enabled = fEnabled
    End If
    cboWorkflowFunctionRecordTable.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
    lblWorkflowFunctionRecordTable.Enabled = fEnabled
    
    ' Disable the element combo if its not required.
    fEnabled = (cboWorkflowFunctionRecord.ListIndex >= 0)
    If fEnabled Then
      fEnabled = (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD)
    End If
    If cboWorkflowFunctionElement.Enabled Then
      cboWorkflowFunctionElement.Enabled = fEnabled
    End If
    cboWorkflowFunctionElement.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
    lblWorkflowFunctionElement.Enabled = fEnabled
  
    ' Disable the record selector combo if its not required.
    fEnabled = (cboWorkflowFunctionRecord.ListIndex >= 0) _
      And (cboWorkflowFunctionElement.ListIndex >= 0)
    
    If fEnabled Then
      fEnabled = (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
        And (cboWorkflowFunctionElement.ItemData(cboWorkflowFunctionElement.ListIndex) > 0)
  
      If fEnabled Then
        Set wfTemp = GetElementByIdentifier(cboWorkflowFunctionElement.List(cboWorkflowFunctionElement.ListIndex))
        fEnabled = Not wfTemp Is Nothing
        If fEnabled Then
          fEnabled = (wfTemp.ElementType = elem_WebForm)
        End If
      End If
    End If
    If cboWorkflowFunctionRecordSelector.Enabled Then
      cboWorkflowFunctionRecordSelector.Enabled = fEnabled
    End If
    cboWorkflowFunctionRecordSelector.BackColor = IIf(fEnabled, vbWhite, vbButtonFace)
    lblWorkflowFunctionRecordSelector.Enabled = fEnabled
    
    ' -----------------------------------------------
    ' Check if it's OK to enable the OK button
    ' -----------------------------------------------
    
    ' OK only if a valid function is selected
    ' ie. not when the root node, or one of the category nodes is selected.
    fOK = fFunctionSelected
      
    If fOK And (lngRequiredTableID <> 0) Then
      ' If we're dealing with a function that required record identification...
      
      ' OK only if a valid record is selected
      fOK = (cboWorkflowFunctionRecord.ListIndex >= 0)
    
      If fOK Then
        fOK = (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) _
          Or (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_TRIGGEREDRECORD) _
          Or (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_INITIATOR) _
          Or (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_UNIDENTIFIED)
      End If
    
      ' If we are using an 'identified' record...
      If fOK Then
        If (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) = giWFRECSEL_IDENTIFIEDRECORD) Then
          ' OK only if a valid element is selected
          fOK = (cboWorkflowFunctionElement.ListIndex >= 0)
      
          If fOK Then
            fOK = (cboWorkflowFunctionElement.ItemData(cboWorkflowFunctionElement.ListIndex) > 0)
          End If
    
          If fOK Then
            Set wfTemp = GetElementByIdentifier(cboWorkflowFunctionElement.List(cboWorkflowFunctionElement.ListIndex))
            fOK = Not wfTemp Is Nothing
      
            If fOK And (wfTemp.ElementType = elem_WebForm) Then
              ' OK only if a valid record selector is selected
              fOK = (cboWorkflowFunctionRecordSelector.ListIndex >= 0)
      
              If fOK Then
                fOK = (cboWorkflowFunctionRecordSelector.ItemData(cboWorkflowFunctionRecordSelector.ListIndex) > 0)
              End If
            End If
          End If
        End If
      End If
    
      ' OK only if a valid record table is selected
      If fOK Then
        If (cboWorkflowFunctionRecord.ItemData(cboWorkflowFunctionRecord.ListIndex) <> giWFRECSEL_UNIDENTIFIED) Then
          fOK = (cboWorkflowFunctionRecordTable.ListIndex >= 0)
      
          If fOK Then
            fOK = (cboWorkflowFunctionRecordTable.ItemData(cboWorkflowFunctionRecordTable.ListIndex) > 0)
          End If
        End If
      End If
    End If
  End If


  'MH20070615 Fault 12332
  'cmdOk.Enabled = fOK
  cmdOK.Enabled = (fOK And fFunctionSelected)
  
End Sub

Private Function GetElementByIdentifier(psIdentifier As String) As VB.Control
  ' Return the element with the given identifier.
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  
  If Len(Trim(psIdentifier)) = 0 Then
    Exit Function
  End If
  
  For lngLoop = 2 To UBound(maWFPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfTemp = maWFPrecedingElements(lngLoop)

    If (UCase(Trim(wfTemp.Identifier)) = UCase(Trim(psIdentifier))) Then
      Set GetElementByIdentifier = wfTemp
      Exit For
    End If
    
    Set wfTemp = Nothing
  Next lngLoop
  
End Function



Private Sub FormatFieldControls()
  ' Display only the required Field Component controls.
  Dim fIsChildOfBase As Boolean
  
  ' Disable the column combo if 'COUNT' is selected.
  If cboFldColumn.Enabled Then
    cboFldColumn.Enabled = Not optField(1).value
  End If
  cboFldDummyColumn.Visible = optField(1).value
  cboFldColumn.BackColor = IIf(cboFldColumn.Enabled, vbWhite, vbButtonFace)
  lblFldField.Enabled = Not optField(1).value
  
  If mfFieldByValue Then
    ' Only enable the field selection options if the selected table
    ' is a child of the expression's parent database.
    With recRelEdit
      .Index = "idxParentID"
      .Seek "=", mobjComponent.ParentExpression.BaseTableID, mobjComponent.Component.TableID
      
      fIsChildOfBase = Not .NoMatch
    End With

    fraFldSelOptions.Enabled = fIsChildOfBase
    
    optFieldSel(0).Enabled = (fIsChildOfBase And optField(0).value)
    optFieldSel(1).Enabled = (fIsChildOfBase And optField(0).value)
    optFieldSel(2).Enabled = (fIsChildOfBase And optField(0).value) And _
      ((mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMECALCULATION) And _
        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) And _
        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMEFILTER) And _
        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_LINKFILTER) And _
        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHJOINEXPRESSION) And _
        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHSCOREEXPRESSION))
    asrFldSelLine.Enabled = optFieldSel(2).Enabled
    
    If (Not optFieldSel(1).Enabled) And (optFieldSel(1).value) Then
      optFieldSel(0).value = True
    End If
    If (Not optFieldSel(2).Enabled) And (optFieldSel(2).value) Then
      optFieldSel(0).value = True
    End If
       
    If (Not optFieldSel(2).value) Then
      asrFldSelLine.Text = vbNullString
    End If
    
    ' Only enable the line number control if required.
    asrFldSelLine.Enabled = (fIsChildOfBase And optFieldSel(2).value)
    asrFldSelLine.BackColor = IIf(asrFldSelLine.Enabled, vbWhite, vbButtonFace)
    
    lblFldOrder.Enabled = (fIsChildOfBase And optField(0).value)
    cmdFldSelOrder.Enabled = (fIsChildOfBase And optField(0).value)
    If Not cmdFldSelOrder.Enabled Then
      mobjComponent.Component.SelectionOrderID = 0
      txtFldSelOrder.Text = ""
    End If
    
    lblFldFilter.Enabled = fIsChildOfBase
    cmdFldSelFilter.Enabled = fIsChildOfBase
    If Not cmdFldSelFilter.Enabled Then
      mobjComponent.Component.SelectionFilter = 0
      txtFldSelFilter.Text = ""
    End If
  End If
  
End Sub


Private Sub optWorkflowField_Refresh()
  Dim iSelection As FieldSelectionTypes
  
  iSelection = mobjComponent.Component.SelectionType

  Select Case iSelection
    Case giSELECT_RECORDCOUNT
      optWorkflowField(1).value = True
    Case giSELECT_RECORDTOTAL
      optWorkflowField(2).value = True
    Case Else
      optWorkflowField(0).value = True
  End Select
    
  cboWorkflowFieldTable_Refresh
    
End Sub


Private Sub optField_Refresh()
  Dim iSelection As FieldSelectionTypes
  
  iSelection = mobjComponent.Component.SelectionType

  Select Case iSelection
    Case giSELECT_RECORDCOUNT
      optField(1).value = True
    Case giSELECT_RECORDTOTAL
      optField(2).value = True
    Case Else
      optField(0).value = True
  End Select
    
  cboFldTable_Refresh
    
End Sub



Private Sub optFieldSel_Refresh()
  Dim iSelection As FieldSelectionTypes
  
  iSelection = mobjComponent.Component.SelectionType

  Select Case iSelection
    Case giSELECT_LASTRECORD
      optFieldSel(1).value = True
    Case giSELECT_SPECIFICRECORD
      optFieldSel(2).value = True
    Case Else
      optFieldSel(0).value = True
      mobjComponent.Component.SelectionType = iSelection
      
      If (iSelection <> giSELECT_RECORDCOUNT) And _
        (iSelection <> giSELECT_RECORDTOTAL) Then
        mobjComponent.Component.SelectionType = giSELECT_FIRSTRECORD
      End If
  End Select

End Sub



Private Sub optWorkflowFieldSel_Refresh()
  Dim iSelection As FieldSelectionTypes
  
  iSelection = mobjComponent.Component.SelectionType

  Select Case iSelection
    Case giSELECT_LASTRECORD
      optWorkflowFieldSel(1).value = True
    Case giSELECT_SPECIFICRECORD
      optWorkflowFieldSel(2).value = True
    Case Else
      optWorkflowFieldSel(0).value = True
      mobjComponent.Component.SelectionType = iSelection
      
      If (iSelection <> giSELECT_RECORDCOUNT) And _
        (iSelection <> giSELECT_RECORDTOTAL) Then
        mobjComponent.Component.SelectionType = giSELECT_FIRSTRECORD
      End If
  End Select

End Sub




Private Sub optField_Click(Index As Integer)
  ' Update the component object.
  Select Case Index
    Case 1:
      mobjComponent.Component.SelectionType = giSELECT_RECORDCOUNT
    Case 2:
      mobjComponent.Component.SelectionType = giSELECT_RECORDTOTAL
    Case Else:
      mobjComponent.Component.SelectionType = giSELECT_FIRSTRECORD
  End Select
  
  cboFldTable_Refresh
  
End Sub

Private Sub optFieldSel_Click(Index As Integer)
  ' Update the component object.
  Select Case Index
    Case 0:
      mobjComponent.Component.SelectionType = giSELECT_FIRSTRECORD
    Case 1:
      mobjComponent.Component.SelectionType = giSELECT_LASTRECORD
    Case 2:
      mobjComponent.Component.SelectionType = giSELECT_SPECIFICRECORD
  End Select
  
  ' Display only the required controls.
  FormatFieldControls

End Sub


Private Sub optWorkflowField_Click(Index As Integer)
  ' Update the component object.
  Select Case Index
    Case 1:
      mobjComponent.Component.SelectionType = giSELECT_RECORDCOUNT
    Case 2:
      mobjComponent.Component.SelectionType = giSELECT_RECORDTOTAL
    Case Else:
      mobjComponent.Component.SelectionType = giSELECT_FIRSTRECORD
  End Select
  
  cboWorkflowFieldTable_Refresh

End Sub

Private Sub optWorkflowFieldSel_Click(Index As Integer)
  ' Update the component object.
  If Not mfInitializing Then
    Select Case Index
      Case 0:
        mobjComponent.Component.SelectionType = giSELECT_FIRSTRECORD
      Case 1:
        mobjComponent.Component.SelectionType = giSELECT_LASTRECORD
      Case 2:
        mobjComponent.Component.SelectionType = giSELECT_SPECIFICRECORD
    End Select
    
    ' Display only the required controls.
    FormatWorkflowFieldControls
  End If
  
End Sub

Private Sub ssTreeFuncFunction_Collapse(Node As SSActiveTreeView.SSNode)
  ' If the specified node is the root node keep it expanded.
  If Node.key = "FUNCTION_ROOT" Then
    Node.Expanded = True
  End If

End Sub


Private Sub ssTreeFuncFunction_DblClick()
  ' Confirm the function selection.
  If cmdOK.Enabled Then
    cmdOK_Click
  End If

End Sub


Private Sub ssTreeFuncFunction_NodeClick(Node As SSActiveTreeView.SSNode)
  cboWorkflowFunctionRecord_Refresh
  FormatFunctionControls

End Sub

Private Sub ssTreeOpOperator_Collapse(Node As SSActiveTreeView.SSNode)
  ' If the specified node is the root node keep it expanded.
  If Node.key = "OPERATOR_ROOT" Then
    Node.Expanded = True
  End If
  
End Sub


Private Sub ssTreeOpOperator_DblClick()
  ' Confirm the operator selection.
  If cmdOK.Enabled Then
    cmdOK_Click
  End If
  
End Sub

Private Sub ssTreeOpOperator_NodeClick(Node As SSActiveTreeView.SSNode)
  ' Update the component object if an operator has been selected.
  Dim fOperatorSelected As Boolean
  
  fOperatorSelected = False
  
  If Node.key <> "OPERATOR_ROOT" Then
    If Node.Parent.key <> "OPERATOR_ROOT" Then
      fOperatorSelected = True
    End If
  End If
  
  ' Only enable the OK button if an operator has been selected.
  ' ie. not when the root node, or one of the category nodes is selected.
  cmdOK.Enabled = fOperatorSelected
  
End Sub

Private Sub txtPValDefaultCharacter_GotFocus()
  ' Select all text upon focus.
  UI.txtSelText

End Sub


Private Sub txtPValFormat_Change()
  ' Update the component object with the new value.
  mobjComponent.Component.ValueFormat = txtPValFormat.Text

End Sub

Private Sub txtPValFormat_GotFocus()
  ' Select all text upon focus.
  UI.txtSelText

End Sub


Private Sub txtPValPrompt_Change()
  ' Update the component object with the new value.
  mobjComponent.Component.Prompt = txtPValPrompt.Text
    
  ' Only enable the OK button if a prompt is entered and
  ' there is a valid column selected for table type prompted values.
  cmdOK.Enabled = (Len(Trim(txtPValPrompt.Text)) > 0) And _
    (mobjComponent.Component.ValueType <> giEXPRVALUE_TABLEVALUE Or _
    mobjComponent.Component.LookupColumn > 0)
    
End Sub

Private Sub txtPValPrompt_GotFocus()
  ' Select all text upon focus.
  UI.txtSelText

End Sub


Private Sub txtValCharacterValue_GotFocus()
  ' Select all text upon focus.
  UI.txtSelText

End Sub





Private Function FormatComponentTypeFrame()
  ' Configure controls in the Component Type frame that are dependent on the expression type.
  Dim fFieldEnabled As Boolean
  Dim fFunctionEnabled As Boolean
  Dim fCalculationEnabled As Boolean
  Dim fValueEnabled As Boolean
  Dim fOperatorEnabled As Boolean
  Dim fTableValueEnabled As Boolean
  Dim fPromptedValueEnabled As Boolean
  Dim fCustomCalculationEnabled As Boolean
  Dim fFilterEnabled As Boolean
  Dim fWorkflowValueEnabled As Boolean
  Dim fWorkflowFieldEnabled As Boolean
  Dim fFieldVisible As Boolean
  Dim fFunctionVisible As Boolean
  Dim fCalculationVisible As Boolean
  Dim fValueVisible As Boolean
  Dim fOperatorVisible As Boolean
  Dim fTableValueVisible As Boolean
  Dim fPromptedValueVisible As Boolean
  Dim fCustomCalculationVisible As Boolean
  Dim fFilterVisible As Boolean
  Dim fWorkflowValueVisible As Boolean
  Dim fWorkflowFieldVisible As Boolean
  Dim dblYCoord As Double
  
  Const YSTART = 300
  Const YGAP = 400
  
  ' Initialize default values.
  fFieldEnabled = True
  fFunctionEnabled = True
  fCalculationEnabled = True
  fValueEnabled = True
  fOperatorEnabled = True
  fTableValueEnabled = True
  fPromptedValueEnabled = True
  fCustomCalculationEnabled = False
  fFilterEnabled = True
  fWorkflowValueEnabled = False
  fWorkflowFieldEnabled = False
  
  fFieldVisible = True
  fFunctionVisible = True
  fCalculationVisible = True
  fValueVisible = True
  fOperatorVisible = True
  fTableValueVisible = True
  fPromptedValueVisible = True
  fCustomCalculationVisible = False
  fFilterVisible = True
  fWorkflowValueVisible = False
  fWorkflowFieldVisible = False
    
  ' Disable some component types for some expression types.
  Select Case mobjComponent.ParentExpression.ExpressionType
    Case giEXPR_COLUMNCALCULATION
      fPromptedValueEnabled = False
  
    Case giEXPR_GOTFOCUS
      ' Not used.
  
    Case giEXPR_RECORDVALIDATION
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_DEFAULTVALUE
      fPromptedValueEnabled = False
      fCalculationEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_STATICFILTER
      fPromptedValueEnabled = False
      fCalculationEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_PAGEBREAK
      ' Not used.
    Case giEXPR_ORDER
      ' Not used.
      
    Case giEXPR_RECORDDESCRIPTION
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_OUTLOOKFOLDER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_OUTLOOKSUBJECT
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_VIEWFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
    
    Case giEXPR_RUNTIMECALCULATION
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_RUNTIMEFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = True
      fFilterEnabled = False
  
    'MH20000727 Added Email
    Case giEXPR_EMAIL
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_LINKFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
  
    Case giEXPR_WORKFLOWCALCULATION
      fFieldEnabled = (Not mfFieldByValue)
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
      fWorkflowValueEnabled = True
      fWorkflowFieldEnabled = True

    Case giEXPR_WORKFLOWSTATICFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
      fWorkflowValueEnabled = True
      fWorkflowFieldEnabled = True

    Case giEXPR_WORKFLOWRUNTIMEFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
      fWorkflowValueEnabled = True
      fWorkflowFieldEnabled = True

  End Select
  
  fFieldVisible = fFieldEnabled
  fFunctionVisible = fFunctionEnabled
  fCalculationVisible = fCalculationEnabled
  fValueVisible = fValueEnabled
  fOperatorVisible = fOperatorEnabled
  fTableValueVisible = fTableValueEnabled
  fPromptedValueVisible = fPromptedValueEnabled
  fCustomCalculationVisible = fCustomCalculationEnabled
  fFilterVisible = fFilterEnabled
  fWorkflowValueVisible = fWorkflowValueEnabled
  fWorkflowFieldVisible = fWorkflowFieldEnabled
  
  ' JPD20021121 Fault 4123
  If Not mfFieldByValue Then
    fFunctionEnabled = False
    fCalculationEnabled = False
    fValueEnabled = False
    fOperatorEnabled = False
    fTableValueEnabled = False
    fPromptedValueEnabled = False
    fCustomCalculationEnabled = False
    fFilterEnabled = False
    fWorkflowValueEnabled = False
    fWorkflowFieldEnabled = False
  End If
  
  ' Disable and hide controls as required.
  dblYCoord = YSTART
  
  optComponentType(giCOMPONENT_FIELD).Enabled = fFieldEnabled
  optComponentType(giCOMPONENT_FIELD).Visible = fFieldVisible
  If fFieldVisible Then
    optComponentType(giCOMPONENT_FIELD).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_WORKFLOWFIELD).Enabled = fWorkflowFieldEnabled
  optComponentType(giCOMPONENT_WORKFLOWFIELD).Visible = fWorkflowFieldVisible
  If fWorkflowFieldVisible Then
    optComponentType(giCOMPONENT_WORKFLOWFIELD).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_WORKFLOWVALUE).Enabled = fWorkflowFieldEnabled
  optComponentType(giCOMPONENT_WORKFLOWVALUE).Visible = fWorkflowFieldVisible
  If fWorkflowValueVisible Then
    optComponentType(giCOMPONENT_WORKFLOWVALUE).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_OPERATOR).Enabled = fOperatorEnabled
  optComponentType(giCOMPONENT_OPERATOR).Visible = fOperatorVisible
  If fOperatorVisible Then
    optComponentType(giCOMPONENT_OPERATOR).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_FUNCTION).Enabled = fFunctionEnabled
  optComponentType(giCOMPONENT_FUNCTION).Visible = fFunctionVisible
  If fFunctionVisible Then
    optComponentType(giCOMPONENT_FUNCTION).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_VALUE).Enabled = fValueEnabled
  optComponentType(giCOMPONENT_VALUE).Visible = fValueVisible
  If fValueVisible Then
    optComponentType(giCOMPONENT_VALUE).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_TABLEVALUE).Enabled = fTableValueEnabled
  optComponentType(giCOMPONENT_TABLEVALUE).Visible = fTableValueVisible
  If fTableValueVisible Then
    optComponentType(giCOMPONENT_TABLEVALUE).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_PROMPTEDVALUE).Enabled = fPromptedValueEnabled
  optComponentType(giCOMPONENT_PROMPTEDVALUE).Visible = fPromptedValueVisible
  If fPromptedValueVisible Then
    optComponentType(giCOMPONENT_PROMPTEDVALUE).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_CALCULATION).Enabled = fCalculationEnabled
  optComponentType(giCOMPONENT_CALCULATION).Visible = fCalculationVisible
  If fCalculationVisible Then
    optComponentType(giCOMPONENT_CALCULATION).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
  
  optComponentType(giCOMPONENT_CUSTOMCALC).Enabled = fCustomCalculationEnabled
  optComponentType(giCOMPONENT_CUSTOMCALC).Visible = fCustomCalculationVisible
  If fCustomCalculationVisible Then
    optComponentType(giCOMPONENT_CUSTOMCALC).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If

  optComponentType(giCOMPONENT_FILTER).Enabled = fFilterEnabled
  optComponentType(giCOMPONENT_FILTER).Visible = fFilterVisible
  If fFilterVisible Then
    optComponentType(giCOMPONENT_FILTER).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If

    
End Function

Private Sub FormatFunctionFrame()
  ' Size and position the Function component controls.
  Dim fWorkflowControlsRequired As Boolean
  Dim lngYCoordinate As Long
  Dim lngCOLUMN2 As Long
  
  Const lngCOLUMN1 = 200
  Const lngYGAP = 420
  Const lngCONTROLWIDTH = 2800
  Const lngYFRAMEGAP = 80
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  lngCOLUMN2 = 2000
  
  fWorkflowControlsRequired = (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWCALCULATION) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER)
  
  fraWorkflowFunctionRecord.Visible = fWorkflowControlsRequired
    
  If fWorkflowControlsRequired Then
    lngYCoordinate = 300
    lngCOLUMN2 = lngCOLUMN2 - lngCOLUMN1
    
    ' Format the Workflow Function - Record controls.
    With cboWorkflowFunctionRecord
      .Left = lngCOLUMN2
      .Top = lngYCoordinate
      .Width = lngCONTROLWIDTH
      lblWorkflowFunctionRecord.Left = lngCOLUMN1
      lblWorkflowFunctionRecord.Top = lngYCoordinate + ((.Height - lblWorkflowFunctionRecord.Height) / 2)
      lngYCoordinate = lngYCoordinate + lngYGAP
    End With
      
    ' Format the Workflow Function - Element controls.
    With cboWorkflowFunctionElement
      .Left = lngCOLUMN2
      .Top = lngYCoordinate
      .Width = lngCONTROLWIDTH
      lblWorkflowFunctionElement.Left = lngCOLUMN1
      lblWorkflowFunctionElement.Top = lngYCoordinate + ((.Height - lblWorkflowFunctionElement.Height) / 2)
      lngYCoordinate = lngYCoordinate + lngYGAP
    End With
  
    ' Format the Workflow Function - RecordSelector controls.
    With cboWorkflowFunctionRecordSelector
      .Left = lngCOLUMN2
      .Top = lngYCoordinate
      .Width = lngCONTROLWIDTH
      lblWorkflowFunctionRecordSelector.Left = lngCOLUMN1
      lblWorkflowFunctionRecordSelector.Top = lngYCoordinate + ((.Height - lblWorkflowFunctionRecordSelector.Height) / 2)
      lngYCoordinate = lngYCoordinate + lngYGAP
    End With
  
    ' Format the Workflow Function - Table controls.
    With cboWorkflowFunctionRecordTable
      .Left = lngCOLUMN2
      .Top = lngYCoordinate
      .Width = lngCONTROLWIDTH
      lblWorkflowFunctionRecordTable.Left = lngCOLUMN1
      lblWorkflowFunctionRecordTable.Top = lngYCoordinate + ((.Height - lblWorkflowFunctionRecordTable.Height) / 2)
      lngYCoordinate = lngYCoordinate + lngYGAP + lngYFRAMEGAP
    End With

    With fraWorkflowFunctionRecord
      .Left = lngCOLUMN1
      .Height = lngYCoordinate
      .Top = fraComponent(2).Height - iYBOTTOMGAP - fraWorkflowFunctionRecord.Height
      .Width = fraComponent(2).Width - (2 * lngCOLUMN1)
    End With
  End If

  With ssTreeFuncFunction
    .Left = lngCOLUMN1
    .Top = iYTOPGAP
    .Width = fraComponent(2).Width - (2 * lngCOLUMN1)
    
    If fWorkflowControlsRequired Then
      .Height = fraWorkflowFunctionRecord.Top - (2 * iYBOTTOMGAP)
    Else
      .Height = fraComponent(2).Height - iYTOPGAP - iYBOTTOMGAP
    End If
  End With

End Sub
Private Sub FormatOperatorFrame()
  ' Size and position the Operator component controls.
  Const iXGAP = 200
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  With ssTreeOpOperator
    .Left = iXGAP
    .Top = iYTOPGAP
    .Width = fraComponent(4).Width - (2 * iXGAP)
    .Height = fraComponent(4).Height - iYTOPGAP - iYBOTTOMGAP
  End With
  
End Sub

Private Sub FormatWorkflowFieldFrame()
  ' Size and position the Workflow Field component controls.
  Dim lngCOLUMN2 As Long
  Dim lngYCoordinate As Long

  Const lngCOLUMN1 = 200
  Const lngYGAP = 420
  Const lngYFRAMEGAP = 80
  Const lngXOPTIONGAP = 500
  Const lngCONTROLWIDTH = 2800

  lngCOLUMN2 = 2000
  lngYCoordinate = 300

  With fraWorkflowField
    .Visible = True
    .BorderStyle = vbBSNone

    .Left = lngCOLUMN1
    .Top = lngYCoordinate

    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Field - Table controls.
  With cboWorkflowFieldTable
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWorkflowFieldTable.Left = lngCOLUMN1
    lblWorkflowFieldTable.Top = lngYCoordinate + ((.Height - lblWorkflowFieldTable.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Field - Column controls.
  With cboWorkflowFieldColumn
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWorkflowFieldColumn.Left = lngCOLUMN1
    lblWorkflowFieldColumn.Top = lngYCoordinate + ((.Height - lblWorkflowFieldColumn.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Workflow Field - Record Selection controls.
  With fraWorkflowFieldRecord
    .Visible = True
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(12).Width - (2 * lngCOLUMN1)
  End With
  lngYCoordinate = 300
  lngCOLUMN2 = lngCOLUMN2 - lngCOLUMN1

  ' Format the Workflow Field - Record controls.
  With cboWorkflowFieldRecord
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWorkflowFieldRecord.Left = lngCOLUMN1
    lblWorkflowFieldRecord.Top = lngYCoordinate + ((.Height - lblWorkflowFieldRecord.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Field - Element controls.
  With cboWorkflowFieldElement
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWorkflowFieldElement.Left = lngCOLUMN1
    lblWorkflowFieldElement.Top = lngYCoordinate + ((.Height - lblWorkflowFieldElement.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Field - RecordSelector controls.
  With cboWorkflowFieldRecordSelector
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWorkflowFieldRecordSelector.Left = lngCOLUMN1
    lblWorkflowFieldRecordSelector.Top = lngYCoordinate + ((.Height - lblWorkflowFieldRecordSelector.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Field - Table controls.
  With cboWorkflowFieldRecordTable
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblWorkflowFieldRecordTable.Left = lngCOLUMN1
    lblWorkflowFieldRecordTable.Top = lngYCoordinate + ((.Height - lblWorkflowFieldRecordTable.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP + lngYFRAMEGAP
  End With

  fraWorkflowFieldRecord.Height = lngYCoordinate
  lngYCoordinate = fraWorkflowFieldRecord.Top + fraWorkflowFieldRecord.Height + lngYFRAMEGAP
  lngCOLUMN2 = lngCOLUMN2 + lngCOLUMN1
  
  ' Format the Workflow Field - Selection Options frame.
  With fraWorkflowFldSelOptions
    .Visible = True
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(12).Width - (2 * lngCOLUMN1)
  End With
  lngYCoordinate = 300
  lngCOLUMN2 = lngCOLUMN2 - lngCOLUMN1

  ' Format the Workflow Field - Record Selection controls.
  With fraWorkflowFieldSel
    .BorderStyle = vbBSNone
    .Left = lngCOLUMN1
    .Top = lngYCoordinate

    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  asrWorkflowFldSelLine.Enabled = True
  optWorkflowFieldSel(2).Enabled = asrWorkflowFldSelLine.Enabled

  ' Format the Workflow Field - Order controls.
  With txtWorkflowFldSelOrder
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH - cmdWorkflowFldSelOrder.Width
    cmdWorkflowFldSelOrder.Left = .Left + .Width
    cmdWorkflowFldSelOrder.Top = lngYCoordinate
    lblWorkflowFldOrder.Left = lngCOLUMN1
    lblWorkflowFldOrder.Top = lngYCoordinate + ((.Height - lblWorkflowFldOrder.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Workflow Field - Filter controls.
  With txtWorkflowFldSelFilter
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH - cmdWorkflowFldSelFilter.Width
    cmdWorkflowFldSelFilter.Left = .Left + .Width
    cmdWorkflowFldSelFilter.Top = lngYCoordinate
    lblWorkflowFldFilter.Left = lngCOLUMN1
    lblWorkflowFldFilter.Top = lngYCoordinate + ((.Height - lblWorkflowFldFilter.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP + lngYFRAMEGAP
  End With

  fraWorkflowFldSelOptions.Height = lngYCoordinate
  
End Sub


Private Sub FormatFieldFrame()
  ' Size and position the Field component controls.
  Dim lngCOLUMN2 As Long
  Dim lngYCoordinate As Long
  
  Const lngCOLUMN1 = 200
  Const lngYGAP = 420
  Const lngYFRAMEGAP = 150
  Const lngXOPTIONGAP = 500
  Const lngCONTROLWIDTH = 2800
  
  lngCOLUMN2 = 2000
  lngYCoordinate = 300
  
  With fraField
    .Visible = mfFieldByValue
    .BorderStyle = vbBSNone
    
    If mfFieldByValue Then
      .Left = lngCOLUMN1
      .Top = lngYCoordinate
        
      lngYCoordinate = lngYCoordinate + lngYGAP
    End If
  End With
   
  'JPD 20031216 Islington changes
  'optField(1).Enabled = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
    (mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) And _
    (mobjComponent.ParentExpression.ExpressionType <> giEXPR_DEFAULTVALUE)
  optField(1).Enabled = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
    (mobjComponent.ParentExpression.ExpressionType <> giEXPR_DEFAULTVALUE)
  optField(2).Enabled = optField(1).Enabled
  
  ' Format the Field - Database controls.
  With cboFldTable
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblFldDatabase.Left = lngCOLUMN1
    lblFldDatabase.Top = lngYCoordinate + ((.Height - lblFldDatabase.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Field - Field controls.
  With cboFldColumn
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblFldField.Left = lngCOLUMN1
    lblFldField.Top = lngYCoordinate + ((.Height - lblFldField.Height) / 2)
    
    cboFldDummyColumn.Left = .Left
    cboFldDummyColumn.Top = .Top
    cboFldDummyColumn.Width = .Width
    cboFldDummyColumn.Visible = False
    
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With

  ' Format the Field - Selection Options frame.
  With fraFldSelOptions
    .Visible = mfFieldByValue
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(0).Width - (2 * lngCOLUMN1)
  End With
  lngYCoordinate = 300
  lngCOLUMN2 = lngCOLUMN2 - lngCOLUMN1
  
  ' Format the Field - Record Selection controls.
  With fraFieldSel
    .BorderStyle = vbBSNone
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
        
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  If mfFieldByValue Then
    asrFldSelLine.Enabled = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMECALCULATION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMEFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHJOINEXPRESSION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHSCOREEXPRESSION)
    optFieldSel(2).Enabled = asrFldSelLine.Enabled
  End If

  ' Format the Field - Order controls.
  With txtFldSelOrder
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH - cmdFldSelOrder.Width
    cmdFldSelOrder.Left = .Left + .Width
    cmdFldSelOrder.Top = lngYCoordinate
    lblFldOrder.Left = lngCOLUMN1
    lblFldOrder.Top = lngYCoordinate + ((.Height - lblFldOrder.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Field - Filter controls.
  With txtFldSelFilter
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH - cmdFldSelFilter.Width
    cmdFldSelFilter.Left = .Left + .Width
    cmdFldSelFilter.Top = lngYCoordinate
    lblFldFilter.Left = lngCOLUMN1
    lblFldFilter.Top = lngYCoordinate + ((.Height - lblFldFilter.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP + lngYFRAMEGAP
  End With
  
  fraFldSelOptions.Height = lngYCoordinate
  
End Sub



Private Sub FormatValueFrame()
  ' Size and position the Value component controls.
  Dim lngYCoordinate As Long
  
  Const lngCOLUMN1 = 200
  Const lngCOLUMN2 = 1000
  Const lngYGAP = 400
  Const lngCONTROLWIDTH = 3700
  Const lngDATECONTROLWIDTH = 1400
  
  lngYCoordinate = 300
  
  ' Format the Value - Type controls.
  With cboValType
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblValType.Left = lngCOLUMN1
    lblValType.Top = lngYCoordinate + ((.Height - lblValType.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Value - Character Value controls.
  With txtValCharacterValue
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
    lblValValue.Left = lngCOLUMN1
    lblValValue.Top = lngYCoordinate + ((.Height - lblValValue.Height) / 2)
  End With
  
  ' Format the Field - Numeric Value control.
  With TDBValNumericValue
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngCONTROLWIDTH
  End With
  
  ' Format the Field - Date Value control.
  With asrValDateValue
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = lngDATECONTROLWIDTH
  End With
  
  ' Format the Field - Logic Value controls.
  With optValLogicValue(0)
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    optValLogicValue(1).Left = .Left + .Width + 500
    optValLogicValue(1).Top = lngYCoordinate
  End With
  
End Sub



Private Sub FormatPromptedValueFrame()
  ' Size and position the Prompted Value component controls.
  Dim lngYCoordinate As Long
  
  Const lngCOLUMN1 = 200
  Const lngCOLUMN2 = 900
  Const lngYFRAMEGAP = 50
  Const lngYGAP = 350
  Const lngXGAP = 200
  Const lngCONTROLWIDTH = 3000
  Const lngSPINNERWIDTH = 600
  Const lngDATECONTROLWIDTH = 1200
  
  lngYCoordinate = 300
  
  ' Format the Prompted Value - Prompt controls.
  With txtPValPrompt
    .Left = lngCOLUMN1 + lngCOLUMN2
    .Top = lngYCoordinate
    .Width = fraComponent(0).Width - .Left - lngXGAP
    lblPValPrompt.Left = lngCOLUMN1
    lblPValPrompt.Top = lngYCoordinate + ((.Height - lblPValPrompt.Height) / 2)
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Prompted Value - Value Type frame.
  With fraPValValueType
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(0).Width - (2 * lngCOLUMN1)
  End With
  lngYCoordinate = 300
  
  ' Format the Prompted Value - Value Type controls.
  With cboPValReturnType
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = 1300
  End With
    
  ' Format the Prompted Value - Value Size controls.
  With lblPValSize
    .Left = cboPValReturnType.Left + cboPValReturnType.Width + lngXGAP
    .Top = lngYCoordinate + ((asrPValReturnSize.Height - .Height) / 2)
    asrPValReturnSize.Left = .Left + .Width + (lngXGAP / 2)
    asrPValReturnSize.Top = lngYCoordinate
    asrPValReturnSize.Width = lngSPINNERWIDTH
  End With
  
  ' Format the Prompted Value - Value Decimals controls.
  With lblPValDecimals
    .Left = asrPValReturnSize.Left + asrPValReturnSize.Width + lngXGAP
    .Top = lngYCoordinate + ((asrPValReturnDecimals.Height - .Height) / 2)
    asrPValReturnDecimals.Left = .Left + .Width + (lngXGAP / 2)
    asrPValReturnDecimals.Top = lngYCoordinate
    asrPValReturnDecimals.Width = lngSPINNERWIDTH
  End With
  
  fraPValValueType.Height = lngYCoordinate + cboPValReturnType.Height + 200
  
  lngYCoordinate = fraPValValueType.Top + fraPValValueType.Height + lngYFRAMEGAP

  ' Format the Prompted Value - Format frame.
  With fraPValFormat
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(0).Width - (2 * lngCOLUMN1)
  End With
  
  ' Format the Prompted Value - Table frame.
  With fraPValTable
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(0).Width - (2 * lngCOLUMN1)
  End With
  
  lngYCoordinate = 300
  
  ' Format the Prompted Value - Format controls.
  With txtPValFormat
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraPValFormat.Width - (2 * lngCOLUMN1)
  End With

  lblMaskKey1.Top = txtPValFormat.Top + txtPValFormat.Height + 60
  lblMaskKey3.Top = lblMaskKey1.Top
  lblMaskKey5.Top = lblMaskKey1.Top
  
  lblMaskKey2.Top = lblMaskKey1.Top + lblMaskKey1.Height + 60
  lblMaskKey4.Top = lblMaskKey2.Top
  lblMaskKey6.Top = lblMaskKey2.Top

  fraPValFormat.Height = lblMaskKey6.Top + lblMaskKey6.Height + 150
  
  ' Format the Prompted Value - Table controls.
  With cboPValTable
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = fraPValTable.Width - lngCOLUMN2 - lngCOLUMN1
    lblPValTable.Left = lngCOLUMN1
    lblPValTable.Top = lngYCoordinate + ((.Height - lblPValTable.Height) / 2)
    
    lngYCoordinate = lngYCoordinate + lngYGAP
  End With
  
  ' Format the Prompted Value - Column controls.
  With cboPValColumn
    .Left = lngCOLUMN2
    .Top = lngYCoordinate
    .Width = fraPValTable.Width - lngCOLUMN2 - lngCOLUMN1
    lblPValColumn.Left = lngCOLUMN1
    lblPValColumn.Top = lngYCoordinate + ((.Height - lblPValColumn.Height) / 2)
    
    fraPValTable.Height = .Top + .Height + 200
  End With

  lngYCoordinate = fraPValFormat.Top + fraPValFormat.Height + lngYFRAMEGAP

  ' Format the Prompted Value - Default Value frame.
  With fraPValDefaultValue
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraComponent(0).Width - (2 * lngCOLUMN1)
  End With
  lngYCoordinate = 300
  
  ' Format the Prompted Value - Default Character Value controls.
  With txtPValDefaultCharacter
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraPValDefaultValue.Width - (2 * lngCOLUMN1)
  End With

  ' Format the Prompted Value - Default Numeric Value controls.
  With TDBPValDefaultNumeric
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraPValDefaultValue.Width - (2 * lngCOLUMN1)
  End With

  ' Format the Prompted Value - Default Date Value controls.
  With asrPValDefaultDate
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = lngDATECONTROLWIDTH
  End With

  ' Format the Prompted Value - Default Table Value controls.
  With cboPValDefaultTabVal
    .Left = lngCOLUMN1
    .Top = lngYCoordinate
    .Width = fraPValDefaultValue.Width - (2 * lngCOLUMN1)
  End With

  ' Format the Prompted Value - Default Logic Value controls.
  With optPValDefaultLogic(0)
    .Left = lngCOLUMN1
    .Top = lngYCoordinate + ((txtPValDefaultCharacter.Height - .Height) / 2)
    optPValDefaultLogic(1).Left = .Left + .Width + 500
    optPValDefaultLogic(1).Top = .Top
  End With

  fraPValDefaultValue.Height = txtPValDefaultCharacter.Top + txtPValDefaultCharacter.Height + 200

End Sub




Private Sub FormatCalculationFrame()
  ' Size and position the Calculation component controls.
  Const iXGAP = 200
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  With listCalcCalculation
    .Left = iXGAP
    .Top = iYTOPGAP
    .Width = fraComponent(3).Width - (2 * iXGAP)
    .Height = fraComponent(3).Height - iYTOPGAP - iYBOTTOMGAP
  End With
  
End Sub


Private Sub FormatFilterFrame()
  ' Size and position the Calculation component controls.
  Const iXGAP = 200
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  With listCalcFilters
    .Left = iXGAP
    .Top = iYTOPGAP
    .Width = fraComponent(10).Width - (2 * iXGAP)
    .Height = fraComponent(10).Height - iYTOPGAP - iYBOTTOMGAP
  End With
  
End Sub
Private Function SaveOperator() As Boolean
  ' Write the selected Operator ID to the component.
  Dim lngOperatorID As Long
  
  With ssTreeOpOperator
    If .SelectedNodes.Count > 0 Then
      lngOperatorID = .SelectedItem.key
    Else
      lngOperatorID = 0
    End If
  End With
  
  mobjComponent.Component.OperatorID = lngOperatorID
  
  SaveOperator = True
  
End Function


Private Function SaveValue() As Boolean
  ' Update the component object
  Dim vValidatedDate As Variant
  Dim fSaveOK As Boolean
  
  fSaveOK = True
  
  With mobjComponent.Component
    Select Case .ReturnType
    
      Case giEXPRVALUE_CHARACTER
        .value = txtValCharacterValue.Text
  
      Case giEXPRVALUE_NUMERIC
        .value = TDBValNumericValue.value
    
      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(asrValDateValue.DateValue) Then
          .value = asrValDateValue.DateValue
        Else
          'MH20020425 Fault 3760
          'User should have already had invalid date message...
          
          '''01/08/2001 MH Fault 2579
          '''MsgBox "Invalid date.", vbOKOnly, App.ProductName
          MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
          
          asrValDateValue.SetFocus
          fSaveOK = False
        End If
  
      Case giEXPRVALUE_LOGIC
        .value = optValLogicValue(0).value
    End Select
  End With
  
  SaveValue = fSaveOK
  
End Function



Private Function SavePromptedValue() As Boolean
  ' Update the component object
  Dim vValidatedDate As Variant
  Dim fSaveOK As Boolean
  Dim dtDateValue As Date
  
  fSaveOK = True

  With mobjComponent.Component
    Select Case .ValueType
      Case giEXPRVALUE_CHARACTER
        .DefaultValue = txtPValDefaultCharacter.Text
  
      Case giEXPRVALUE_NUMERIC
        .DefaultValue = TDBPValDefaultNumeric.value
    
      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(asrPValDefaultDate.Text) Then
'          .DefaultValue = asrPValDefaultDate.Value
          .DefaultValue = asrPValDefaultDate.Text
        Else
          .DefaultValue = Null
        End If
      
      Case giEXPRVALUE_LOGIC
        .DefaultValue = optPValDefaultLogic(0).value
        
      Case giEXPRVALUE_TABLEVALUE
        Select Case mDataType
          Case dtNUMERIC, dtINTEGER
            .DefaultValue = Val(cboPValDefaultTabVal.List(cboPValDefaultTabVal.ListIndex))
          Case dtTIMESTAMP
            If IsDate(cboPValDefaultTabVal.List(cboPValDefaultTabVal.ListIndex)) Then
              .DefaultValue = Replace(Format(CDate(cboPValDefaultTabVal.List(cboPValDefaultTabVal.ListIndex)), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
            Else
              .DefaultValue = 0
            End If
          Case Else
            .DefaultValue = cboPValDefaultTabVal.List(cboPValDefaultTabVal.ListIndex)
        End Select
        
    End Select
  End With
  
  SavePromptedValue = fSaveOK
  
End Function




Private Function SaveFunction() As Boolean
  ' Write the selected Function ID to the component.
  Dim lngFunctionID As Long
  
  With ssTreeFuncFunction
    If .SelectedNodes.Count > 0 Then
      lngFunctionID = .SelectedItem.key
    Else
      lngFunctionID = 0
    End If
  End With
  
  mobjComponent.Component.FunctionID = lngFunctionID
  
  SaveFunction = True
  
End Function



Private Function fldSelFilter_Refresh() As Boolean
  ' Refresh the Field Selection Filter controls.
  ' Validate the expression selection at the same time.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngTableID As Long
  Dim lngFilterID As Long
  
  fOK = True
  
  ' Get the current component's table and order id.
  lngTableID = mobjComponent.Component.TableID
  lngFilterID = mobjComponent.Component.SelectionFilter
  
  ' Check if the selected expression is for the current table.
  With recExprEdit
    .Index = "idxExprID"
    .Seek "=", lngFilterID, False
    
    If Not .NoMatch Then
      If (!TableID <> lngTableID) _
        And (mobjComponent.ComponentType <> giCOMPONENT_WORKFLOWFIELD) Then
        mobjComponent.Component.SelectionFilter = 0
      End If
    Else
      mobjComponent.Component.SelectionFilter = 0
    End If
  End With
    
  ' Refresh the filter controls.
  GetFldSelFilterDetails

TidyUpAndExit:
  fldSelFilter_Refresh = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Private Function fldSelOrder_Refresh() As Boolean
  ' Refresh the Field Selection Order controls.
  ' Validate the order selection at the same time.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngTableID As Long
  Dim lngOrderId As Long
  
  fOK = True
  
  ' Get the current component's table and order id.
  lngTableID = mobjComponent.Component.TableID
  lngOrderId = mobjComponent.Component.SelectionOrderID
  
  ' Check if the selected expression is for the current table.
  With recOrdEdit
    .Index = "idxID"
    .Seek "=", lngOrderId
    
    If Not .NoMatch Then
      If (!TableID <> lngTableID) _
        And (mobjComponent.ComponentType <> giCOMPONENT_WORKFLOWFIELD) Then
        
        mobjComponent.Component.SelectionOrderID = 0
      End If
    Else
      mobjComponent.Component.SelectionOrderID = 0
    End If
  End With
    
  ' Refresh the order controls.
  GetFldSelOrderDetails

TidyUpAndExit:
  fldSelOrder_Refresh = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


