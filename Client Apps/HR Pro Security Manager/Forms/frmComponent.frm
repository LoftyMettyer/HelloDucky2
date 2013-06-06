VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmExprComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expression Component"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1038
   Icon            =   "frmComponent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraComponent 
      Caption         =   "Filter :"
      Height          =   1050
      Index           =   8
      Left            =   6300
      TabIndex        =   105
      Tag             =   "10"
      Top             =   6720
      Width           =   795
      Begin VB.CheckBox chkOnlyMyFilters 
         Caption         =   "Only show filters where owner is "
         Height          =   240
         Left            =   60
         TabIndex        =   63
         Top             =   600
         Width           =   1410
      End
      Begin VB.ListBox listCalcFilters 
         Height          =   255
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   62
         Top             =   270
         Width           =   1000
      End
   End
   Begin VB.Frame fraComponentType 
      Caption         =   "Type :"
      Height          =   3570
      Left            =   30
      TabIndex        =   97
      Top             =   0
      Width           =   2250
      Begin VB.CommandButton cmdEditFilter 
         Caption         =   "E&dit Filter..."
         Height          =   400
         Left            =   135
         TabIndex        =   9
         Top             =   75
         Width           =   1860
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "F&ilter"
         Height          =   315
         Index           =   10
         Left            =   105
         TabIndex        =   7
         Tag             =   "COMP_FILTER"
         Top             =   2820
         Width           =   825
      End
      Begin VB.CommandButton CmdEditCalculation 
         Caption         =   "E&dit Calculation..."
         Height          =   400
         Left            =   135
         TabIndex        =   10
         Top             =   75
         Width           =   1860
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Field"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   0
         Tag             =   "COMP_FIELD"
         Top             =   300
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optComponentType 
         BackColor       =   &H80000010&
         Caption         =   "Custo&m Calculation"
         Height          =   315
         Index           =   8
         Left            =   90
         TabIndex        =   8
         Tag             =   "COMP_CUSTOMCALC"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Prompted Value"
         Height          =   315
         Index           =   7
         Left            =   105
         TabIndex        =   5
         Tag             =   "COMP_PROMPTED"
         Top             =   2100
         Width           =   1770
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "Loo&kup Table Value"
         Height          =   315
         Index           =   6
         Left            =   105
         TabIndex        =   4
         Tag             =   "COMP_LOOKUPVALUE"
         Top             =   1740
         Width           =   2055
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "Op&erator"
         Height          =   315
         Index           =   5
         Left            =   105
         TabIndex        =   1
         Tag             =   "COMP_OPERATOR"
         Top             =   660
         Width           =   1275
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Value"
         Height          =   315
         Index           =   4
         Left            =   105
         TabIndex        =   3
         Tag             =   "COMP_VALUE"
         Top             =   1380
         Width           =   975
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "C&alculation"
         Height          =   315
         Index           =   3
         Left            =   105
         TabIndex        =   6
         Tag             =   "COMP_CALCULATION"
         Top             =   2460
         Width           =   1350
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "F&unction"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   2
         Tag             =   "COMP_FUNCTION"
         Top             =   1020
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6195
      TabIndex        =   96
      Top             =   7845
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   95
      Top             =   8310
      Width           =   1200
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Value :"
      Height          =   2250
      Index           =   1
      Left            =   7440
      TabIndex        =   92
      Tag             =   "4"
      Top             =   6480
      Width           =   2250
      Begin VB.TextBox txtValCharacterValue 
         Height          =   315
         Left            =   650
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   645
         Width           =   1455
      End
      Begin VB.OptionButton optValLogicValue 
         Caption         =   "&False"
         Height          =   315
         Index           =   1
         Left            =   1400
         TabIndex        =   42
         Top             =   1850
         Width           =   765
      End
      Begin VB.OptionButton optValLogicValue 
         Caption         =   "&True"
         Height          =   315
         Index           =   0
         Left            =   630
         TabIndex        =   41
         Top             =   1850
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.ComboBox cboValType 
         Height          =   315
         ItemData        =   "frmComponent.frx":000C
         Left            =   650
         List            =   "frmComponent.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   250
         Width           =   1455
      End
      Begin TDBNumber6Ctl.TDBNumber TDBValNumericValue 
         Height          =   315
         Left            =   660
         TabIndex        =   39
         Top             =   1050
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calculator      =   "frmComponent.frx":0041
         Caption         =   "frmComponent.frx":0061
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmComponent.frx":00C6
         Keys            =   "frmComponent.frx":00E4
         Spin            =   "frmComponent.frx":012E
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
      Begin GTMaskDate.GTMaskDate asrValDateValue 
         Height          =   315
         Left            =   645
         TabIndex        =   40
         Top             =   1410
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
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblValValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   100
         TabIndex        =   94
         Top             =   710
         Width           =   495
      End
      Begin VB.Label lblValType 
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   105
         TabIndex        =   93
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Field :"
      Height          =   3375
      Index           =   0
      Left            =   2310
      TabIndex        =   86
      Tag             =   "1"
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cboFldDummyColumn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Frame fraField 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4455
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
            Caption         =   "Total"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   14
            Top             =   60
            Width           =   840
         End
      End
      Begin VB.Frame fraFldSelOptions 
         Caption         =   "Child Field Options :"
         Height          =   1920
         Left            =   100
         TabIndex        =   87
         Top             =   1300
         Width           =   4305
         Begin VB.Frame fraFieldSel 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton optFieldSel 
               Caption         =   "Specific"
               Height          =   195
               Index           =   2
               Left            =   2220
               TabIndex        =   20
               Top             =   60
               Width           =   1035
            End
            Begin VB.OptionButton optFieldSel 
               Caption         =   "Last"
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   19
               Top             =   60
               Width           =   735
            End
            Begin VB.OptionButton optFieldSel 
               Caption         =   "First"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   60
               Value           =   -1  'True
               Width           =   700
            End
            Begin COASpinner.COA_Spinner asrFldSelLine 
               Height          =   315
               Left            =   3400
               TabIndex        =   21
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
         Begin VB.CommandButton cmdFldSelFilter 
            Height          =   315
            Left            =   2270
            Picture         =   "frmComponent.frx":0156
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1395
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txtFldSelFilter 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   24
            Top             =   1395
            Width           =   735
         End
         Begin VB.TextBox txtFldSelOrder 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   22
            Top             =   1035
            Width           =   735
         End
         Begin VB.CommandButton cmdFldSelOrder 
            Height          =   315
            Left            =   2270
            Picture         =   "frmComponent.frx":01CE
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1035
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Label lblFldFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter :"
            Height          =   195
            Left            =   105
            TabIndex        =   89
            Top             =   1455
            Width           =   465
         End
         Begin VB.Label lblFldOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order :"
            Height          =   195
            Left            =   105
            TabIndex        =   88
            Top             =   1095
            Width           =   525
         End
      End
      Begin VB.ComboBox cboFldTable 
         Height          =   315
         ItemData        =   "frmComponent.frx":0246
         Left            =   1000
         List            =   "frmComponent.frx":0248
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   615
         Width           =   1275
      End
      Begin VB.ComboBox cboFldColumn 
         Height          =   315
         Left            =   1000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1005
         Width           =   1275
      End
      Begin VB.Label lblFldDatabase 
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   105
         TabIndex        =   91
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lblFldField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   105
         TabIndex        =   90
         Top             =   1065
         Width           =   630
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Function :"
      Height          =   1400
      Index           =   2
      Left            =   6180
      TabIndex        =   85
      Tag             =   "2"
      Top             =   3720
      Width           =   1200
      Begin SSActiveTreeView.SSTree ssTreeFuncFunction 
         Height          =   1000
         Left            =   100
         TabIndex        =   26
         Top             =   250
         Width           =   1000
         _ExtentX        =   1746
         _ExtentY        =   1746
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
      Caption         =   "Calculation :"
      Height          =   885
      Index           =   3
      Left            =   8040
      TabIndex        =   84
      Tag             =   "3"
      Top             =   3240
      Width           =   1650
      Begin VB.CheckBox chkOnlyMine 
         Caption         =   "Only show calculations where owner is "
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   585
         Width           =   1455
      End
      Begin VB.ListBox listCalcCalculation 
         Height          =   255
         Left            =   100
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   250
         Width           =   1000
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Lookup Table Value :"
      Height          =   2100
      Index           =   5
      Left            =   7680
      TabIndex        =   80
      Tag             =   "6"
      Top             =   4320
      Width           =   2000
      Begin VB.ComboBox cboTabValTable 
         Height          =   315
         ItemData        =   "frmComponent.frx":024A
         Left            =   750
         List            =   "frmComponent.frx":024C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   250
         Width           =   1000
      End
      Begin VB.ComboBox cboTabValColumn 
         Height          =   315
         ItemData        =   "frmComponent.frx":024E
         Left            =   750
         List            =   "frmComponent.frx":0250
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   650
         Width           =   1000
      End
      Begin VB.ComboBox cboTabValValue 
         Height          =   315
         ItemData        =   "frmComponent.frx":0252
         Left            =   750
         List            =   "frmComponent.frx":0254
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1050
         Width           =   1000
      End
      Begin VB.Label lblLookupValNotFound 
         Caption         =   "Original lookup value was not found"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   210
         TabIndex        =   106
         Top             =   1605
         Visible         =   0   'False
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTabValTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   100
         TabIndex        =   83
         Top             =   310
         Width           =   495
      End
      Begin VB.Label lblTabValColumn 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   105
         TabIndex        =   82
         Top             =   705
         Width           =   630
      End
      Begin VB.Label lblTabValValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   100
         TabIndex        =   81
         Top             =   1110
         Width           =   495
      End
   End
   Begin VB.Frame fraComponent 
      BackColor       =   &H80000010&
      Caption         =   "Custom Calculation :"
      Height          =   2800
      Index           =   7
      Left            =   6975
      TabIndex        =   75
      Tag             =   "8"
      Top             =   30
      Width           =   2200
      Begin VB.ComboBox cboCustCalculation 
         Height          =   315
         ItemData        =   "frmComponent.frx":0256
         Left            =   1100
         List            =   "frmComponent.frx":0266
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   250
         Width           =   1000
      End
      Begin VB.Frame fraCustParameters 
         Caption         =   "Parameters :"
         Height          =   2100
         Left            =   100
         TabIndex        =   76
         Top             =   600
         Width           =   1750
         Begin VB.ListBox lstCustParameters 
            Height          =   645
            Left            =   100
            TabIndex        =   31
            Top             =   250
            Width           =   1000
         End
         Begin VB.ComboBox cboCustTable 
            Height          =   315
            ItemData        =   "frmComponent.frx":028D
            Left            =   650
            List            =   "frmComponent.frx":029D
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1200
            Width           =   1000
         End
         Begin VB.ComboBox cboCustField 
            Height          =   315
            ItemData        =   "frmComponent.frx":02C4
            Left            =   650
            List            =   "frmComponent.frx":02D4
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1600
            Width           =   1000
         End
         Begin VB.Label lblCustTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   100
            TabIndex        =   78
            Top             =   1260
            Width           =   495
         End
         Begin VB.Label lblCustField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Column :"
            Height          =   195
            Left            =   105
            TabIndex        =   77
            Top             =   1665
            Width           =   630
         End
      End
      Begin VB.Label lblCustCalculation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculation :"
         Height          =   195
         Left            =   105
         TabIndex        =   79
         Top             =   315
         Width           =   885
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Operator :"
      Height          =   1400
      Index           =   4
      Left            =   6180
      TabIndex        =   74
      Tag             =   "5"
      Top             =   5160
      Width           =   1200
      Begin SSActiveTreeView.SSTree ssTreeOpOperator 
         Height          =   1000
         Left            =   100
         TabIndex        =   27
         Top             =   250
         Width           =   1000
         _ExtentX        =   1746
         _ExtentY        =   1746
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
      Caption         =   "Prompted Value :"
      Height          =   4890
      Index           =   6
      Left            =   0
      TabIndex        =   64
      Tag             =   "7"
      Top             =   3795
      Width           =   6165
      Begin VB.TextBox txtPValPrompt 
         Height          =   315
         Left            =   800
         MaxLength       =   40
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   250
         Width           =   1000
      End
      Begin VB.Frame fraPValValueType 
         Caption         =   "Type :"
         Height          =   705
         Left            =   100
         TabIndex        =   70
         Top             =   600
         Width           =   5025
         Begin VB.ComboBox cboPValReturnType 
            Height          =   315
            ItemData        =   "frmComponent.frx":02FB
            Left            =   120
            List            =   "frmComponent.frx":030E
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   240
            Width           =   1000
         End
         Begin COASpinner.COA_Spinner asrPValReturnDecimals 
            Height          =   315
            Left            =   3870
            TabIndex        =   46
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
            Left            =   1860
            TabIndex        =   45
            Top             =   255
            Width           =   915
            _ExtentX        =   1614
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
         Begin VB.Label lblPValDecimals 
            BackStyle       =   0  'Transparent
            Caption         =   "Decimals :"
            Height          =   195
            Left            =   2910
            TabIndex        =   72
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lblPValSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            Height          =   195
            Left            =   1230
            TabIndex        =   71
            Top             =   285
            Width           =   570
         End
      End
      Begin VB.Frame fraPValDefaultValue 
         Caption         =   "Default Value :"
         Height          =   960
         Left            =   120
         TabIndex        =   69
         Top             =   3675
         Width           =   5940
         Begin VB.OptionButton optPValDefaultDateType 
            Caption         =   "Current"
            Height          =   225
            Index           =   1
            Left            =   1380
            TabIndex        =   59
            Top             =   720
            Width           =   1000
         End
         Begin VB.OptionButton optPValDefaultDateType 
            Caption         =   "Year End"
            Height          =   225
            Index           =   5
            Left            =   4395
            TabIndex        =   61
            Top             =   705
            Width           =   1200
         End
         Begin VB.OptionButton optPValDefaultDateType 
            Caption         =   "Month End"
            Height          =   225
            Index           =   3
            Left            =   2910
            TabIndex        =   60
            Top             =   750
            Width           =   1300
         End
         Begin VB.OptionButton optPValDefaultDateType 
            Caption         =   "Year Start"
            Height          =   225
            Index           =   4
            Left            =   4410
            TabIndex        =   58
            Top             =   480
            Width           =   1200
         End
         Begin VB.OptionButton optPValDefaultDateType 
            Caption         =   "Month Start"
            Height          =   225
            Index           =   2
            Left            =   2910
            TabIndex        =   57
            Top             =   540
            Width           =   1300
         End
         Begin VB.OptionButton optPValDefaultDateType 
            Caption         =   "Explicit"
            Height          =   225
            Index           =   0
            Left            =   1380
            TabIndex        =   56
            Top             =   540
            Value           =   -1  'True
            Width           =   1000
         End
         Begin VB.TextBox txtPValDefaultCharacter 
            Height          =   315
            Left            =   75
            TabIndex        =   50
            Top             =   225
            Width           =   1000
         End
         Begin VB.ComboBox cboPValDefaultTabVal 
            Height          =   315
            Left            =   3315
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   225
            Width           =   1545
         End
         Begin VB.OptionButton optPValDefaultLogic 
            Caption         =   "&Yes"
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   54
            Top             =   600
            Width           =   630
         End
         Begin VB.OptionButton optPValDefaultLogic 
            Caption         =   "&No"
            Height          =   315
            Index           =   1
            Left            =   1380
            TabIndex        =   55
            Top             =   600
            Width           =   855
         End
         Begin TDBNumber6Ctl.TDBNumber TDBPValDefaultNumeric 
            Height          =   315
            Left            =   2235
            TabIndex        =   52
            Top             =   225
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            Calculator      =   "frmComponent.frx":0341
            Caption         =   "frmComponent.frx":0361
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmComponent.frx":03C6
            Keys            =   "frmComponent.frx":03E4
            Spin            =   "frmComponent.frx":042E
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
         Begin GTMaskDate.GTMaskDate asrPValDefaultDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   51
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
      Begin VB.Frame fraPValTable 
         Caption         =   "Lookup Table Value :"
         Height          =   990
         Left            =   135
         TabIndex        =   66
         Top             =   1350
         Width           =   2235
         Begin VB.ComboBox cboPValTable 
            Height          =   315
            Left            =   720
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   240
            Width           =   1000
         End
         Begin VB.ComboBox cboPValColumn 
            Height          =   315
            Left            =   795
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   555
            Width           =   1000
         End
         Begin VB.Label lblPValTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   255
            Width           =   495
         End
         Begin VB.Label lblPValColumn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Column :"
            Height          =   195
            Left            =   105
            TabIndex        =   67
            Top             =   615
            Width           =   630
         End
      End
      Begin VB.Frame fraPValFormat 
         Caption         =   "Format :"
         Height          =   1260
         Left            =   135
         TabIndex        =   65
         Top             =   2355
         Width           =   5925
         Begin TDBMask6Ctl.TDBMask tdbMaskTest 
            Height          =   270
            Left            =   4485
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   930
            Visible         =   0   'False
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   476
            Caption         =   "frmComponent.frx":0456
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmComponent.frx":04BB
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
            Format          =   "&&&&&&&&&&"
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
            Text            =   "TDBMask1__"
            Value           =   "TDBMask1"
         End
         Begin VB.TextBox txtPValFormat 
            Height          =   315
            Left            =   195
            MaxLength       =   128
            TabIndex        =   49
            Top             =   270
            Width           =   1000
         End
         Begin VB.Label lblMaskKey6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "\ - Follow by literal"
            Height          =   195
            Left            =   3705
            TabIndex        =   103
            Top             =   900
            Width           =   1605
         End
         Begin VB.Label lblMaskKey5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B - Binary (0 or 1)"
            Height          =   195
            Left            =   3660
            TabIndex        =   102
            Top             =   645
            Width           =   1680
         End
         Begin VB.Label lblMaskKey2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a - Lower case"
            Height          =   195
            Left            =   195
            TabIndex        =   101
            Top             =   885
            Width           =   1365
         End
         Begin VB.Label lblMaskKey4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "# - Numbers, Symbols"
            Height          =   195
            Left            =   1590
            TabIndex        =   100
            Top             =   900
            Width           =   1995
         End
         Begin VB.Label lblMaskKey3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "9 - Numbers (0-9)"
            Height          =   195
            Left            =   1590
            TabIndex        =   99
            Top             =   645
            Width           =   1590
         End
         Begin VB.Label lblMaskKey1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A - Upper case"
            Height          =   195
            Left            =   195
            TabIndex        =   98
            Top             =   645
            Width           =   1425
         End
      End
      Begin VB.Label lblPValPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt :"
         Height          =   195
         Left            =   105
         TabIndex        =   73
         Top             =   315
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmExprComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Component definition variables.
Private mobjComponent As clsExprComponent
Private miComponentType As ExpressionComponentTypes
Private mavColumns() As Variant
Private mDataType As SQLDataType

' Form handling variables.
Private mfCancelled As Boolean
Private mfFunctionsPopulated As Boolean
Private mfCalculationsPopulated As Boolean
Private mfFiltersPopulated As Boolean
Private mfOperatorsPopulated As Boolean
Private mfTabValTablesPopulated As Boolean
Private mfPValTablesPopulated As Boolean
Private mblnLoading As Boolean
Private mblnApplySystemPermissions As Boolean

Private mbEnableEditCalculation As Boolean
Private mbEnableViewCalculation As Boolean
Private mbEnableEditFilter As Boolean
Private mbEnableViewFilter As Boolean

Private mfFieldByValue As Boolean

Public Property Set Component(pobjComponent As clsExprComponent)
  ' Set the component property.
  Set mobjComponent = pobjComponent
  
  ' Set the component type.
  miComponentType = mobjComponent.ComponentType
  mfFieldByValue = (mobjComponent.ParentExpression.ReturnType < 100)
  If mobjComponent.ComponentType = giCOMPONENT_FIELD Then
    mobjComponent.Component.FieldPassType = IIf(mfFieldByValue, giPASSBY_VALUE, giPASSBY_REFERENCE)
  End If
    
  ' Format the controls within the frames.
  FormatFieldFrame
  
  ' Format the Component Type frame for the new component.
  FormatComponentTypeFrame
  
  ' Format the Component frame for the new component.
  If optComponentType(miComponentType).Value Then
    DisplayComponentFrame
  Else
    optComponentType(miComponentType).Value = True
  End If
  
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

Private Sub InitializeComponentControls()

  ' Call the required sub-routine to initialze the component definition controls and also dynamicallly
  'set the HelpCOntextIDs in order to get the correct one when using the Help Files
  Select Case miComponentType
    Case giCOMPONENT_FIELD
      InitializeFieldControls
      Me.HelpContextID = 1054
      
    Case giCOMPONENT_FUNCTION
      InitializeFunctionControls
      Me.HelpContextID = 1055
      
    Case giCOMPONENT_CALCULATION
      InitializeCalcControls
      Me.HelpContextID = 1056
      
    Case giCOMPONENT_VALUE
      InitializeValueControls
      Me.HelpContextID = 1057

    Case giCOMPONENT_OPERATOR
      InitializeOperatorControls
      Me.HelpContextID = 1058
      
    Case giCOMPONENT_TABLEVALUE
      InitializeTableValueControls
      Me.HelpContextID = 1059
      
    Case giCOMPONENT_PROMPTEDVALUE
      InitializePromptedValueControls
      Me.HelpContextID = 1060
      
    Case giCOMPONENT_CUSTOMCALC
      ' Not required.
      'If this does get reinstated you may need to add a unique HelpcontextID
      'in the Doc-to-Help Security Man Help Files
      Me.HelpContextID = 9999
      
    Case giCOMPONENT_EXPRESSION
      ' Not handled in this form.
      'If this does get reinstated you may need to add a unique HelpcontextID
      'in the Doc-to-Help Security Man Help Files
      Me.HelpContextID = 9999

    'JDM - 12/03/01 - Fault 1219 - Add filter to filter
    Case giCOMPONENT_FILTER
        InitializeFilterControls
        Me.HelpContextID = 1061
      
  End Select
  
  With CmdEditCalculation
    .Visible = (optComponentType(giCOMPONENT_CALCULATION).Value = True)
    .Top = optComponentType(giCOMPONENT_FILTER).Top + optComponentType(giCOMPONENT_FILTER).Height + 100
  End With

  ' Show edit filter button
  With cmdEditFilter
    .Visible = (optComponentType(giCOMPONENT_FILTER).Value = True)
    .Top = optComponentType(giCOMPONENT_FILTER).Top + optComponentType(giCOMPONENT_FILTER).Height + 100
  End With
  
End Sub

Private Sub InitializePromptedValueControls()
  ' Initialize the Prompted Value component controls.
  Dim sDefaultCharacter As String
  Dim dblDefaultNumeric As Double
  Dim fDefaultLogic As Boolean
  Dim dtDefaultDate As Date
  Dim iCount As Integer

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
    
    Select Case .valueType
      Case giEXPRVALUE_CHARACTER
        sDefaultCharacter = .DefaultValue
      Case giEXPRVALUE_NUMERIC
        dblDefaultNumeric = .DefaultValue
      Case giEXPRVALUE_LOGIC
        fDefaultLogic = .DefaultValue
        optPValDefaultLogic(0).Value = fDefaultLogic
        optPValDefaultLogic(1).Value = Not optPValDefaultLogic(0).Value
      Case giEXPRVALUE_DATE
        dtDefaultDate = .DefaultValue
        optPValDefaultDateType(.DefaultDateType).Value = True
    End Select
    
    txtPValDefaultCharacter.Text = sDefaultCharacter
    TDBPValDefaultNumeric.Value = dblDefaultNumeric
   
    If CDbl(dtDefaultDate) <> 0 Then
      'asrPValDefaultDate.Value = dtDefaultDate
      asrPValDefaultDate.Text = dtDefaultDate
    End If
  End With
  
End Sub

Private Sub cboPValTable_Refresh()
  ' Prompted Value component - Table combo.
  ' Select the current table.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim lngTableID As Long
  Dim sSQL As String
  Dim rsTable As Recordset
  
  ' Get the Table ID definition.
  sSQL = "SELECT tableID" & _
    " FROM ASRSysColumns" & _
    " WHERE columnID = " & Trim(Str(mobjComponent.Component.LookupColumn))
  Set rsTable = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTable
    If Not (.EOF And .BOF) Then
      lngTableID = !TableID
    Else
      lngTableID = 0
    End If
  
    .Close
  End With
  Set rsTable = Nothing
  
  iIndex = 0
  
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

Private Sub cboPValColumn_Refresh()
  ' Populate the Prompted Value - Column combo, and then
  ' select the current column.
  On Error GoTo ErrorTrap
  
  Dim iIndex As Integer
  Dim iNextIndex As Integer
  Dim iLoop As Integer
  Dim sSQL As String
  Dim rsColumns As Recordset
  
  iIndex = 0
  iLoop = 0
  
  ' Clear the current contents of the combo.
  cboPValColumn.Clear
  
  ' Create an array of column info.
  ' Column 1 = column ID
  ' Column 2 = data type.
  ReDim mavColumns(2, 0)
  
  If cboPValTable.Enabled Then
    ' Get the column definition.
    sSQL = "SELECT columnName, columnID, dataType" & _
      " FROM ASRSysColumns" & _
      " WHERE tableID = " & Trim(Str(cboPValTable.ItemData(cboPValTable.ListIndex))) & _
      " AND columnType <> " & Trim(Str(colSystem)) & _
      " AND columnType <> " & Trim(Str(colLink)) & _
      " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
      " AND dataType <> " & Trim(Str(sqlBoolean)) & _
      " AND dataType <> " & Trim(Str(sqlTypeOle)) & _
      " ORDER BY columnName"
    Set rsColumns = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsColumns
      Do While Not .EOF
        cboPValColumn.AddItem .Fields("columnName")
        cboPValColumn.ItemData(cboPValColumn.NewIndex) = .Fields("columnID")
        
        iNextIndex = UBound(mavColumns, 2) + 1
        ReDim Preserve mavColumns(2, iNextIndex)
        mavColumns(1, iNextIndex) = rsColumns!ColumnID
        mavColumns(2, iNextIndex) = rsColumns!DataType
          
        If !ColumnID = mobjComponent.Component.LookupColumn Then
          iIndex = iLoop
        End If
        
        iLoop = iLoop + 1
        .MoveNext
      Loop
    
      .Close
    End With
    Set rsColumns = Nothing
    
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
    ((mobjComponent.Component.valueType <> giEXPRVALUE_TABLEVALUE) Or (cboPValColumn.Enabled))

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
  
  Dim iIndex As Integer
  Dim sSQL As String
  Dim sDfltValue As String
  Dim rsLookupValues As Recordset
  
  iIndex = 0
  
  sDfltValue = mobjComponent.Component.DefaultValue
  
  ' Clear the current contents of the combo.
  cboPValDefaultTabVal.Clear

  If cboPValTable.Enabled And cboPValColumn.Enabled Then
    
    sSQL = "SELECT DISTINCT " & cboPValColumn.List(cboPValColumn.ListIndex) & " AS lookUpValue" & _
      " FROM " & cboPValTable.List(cboPValTable.ListIndex) & _
      " ORDER BY lookUpValue"
      
    Set rsLookupValues = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsLookupValues
      Do While Not .EOF
        Select Case mDataType
          Case sqlNumeric, sqlInteger
            cboPValDefaultTabVal.AddItem Trim(Str(!LookupValue))
            If !LookupValue = Val(sDfltValue) Then
              iIndex = cboPValDefaultTabVal.NewIndex
            End If
            
          Case sqlDate
            If IsDate(!LookupValue) Then
              'JPD 20041115 Fault 9484
              'cboPValDefaultTabVal.AddItem Format(!LookupValue, "long date")
              'If Format(!LookupValue, "mm/dd/yyyy") = sDfltValue Then
              cboPValDefaultTabVal.AddItem Format(!LookupValue, DateFormat)
              If Replace(Format(!LookupValue, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") = sDfltValue Then
                iIndex = cboPValDefaultTabVal.NewIndex
              End If
            End If
            
          Case Else
            cboPValDefaultTabVal.AddItem Trim(!LookupValue)

            ' JDM - 15/03/01 - Fault 1897 - Get rid of trailing spaces
            If Trim(!LookupValue) = sDfltValue Then
              iIndex = cboPValDefaultTabVal.NewIndex
            End If
        End Select
        
        .MoveNext
      Loop
      
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
  
  Exit Sub
  
ErrorTrap:
  Set rsLookupValues = Nothing
  cboPValDefaultTabVal.Enabled = False

End Sub

Private Sub cboPValTable_Initialize()
  ' Populate the Prompted Value component Table combo.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsTables As Recordset
  
  ' Clear the contents of the combo.
  cboPValTable.Clear
  
  ' Get the order definition.
  sSQL = "SELECT tableName, tableID" & _
    " FROM ASRSysTables" & _
    " WHERE tableType = " & Trim(Str(tabLookup)) & _
    " ORDER BY tableName"
  Set rsTables = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTables
    Do While Not .EOF
      cboPValTable.AddItem !TableName
      cboPValTable.ItemData(cboPValTable.NewIndex) = !TableID
      
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsTables = Nothing
  
  ' Enable the combo if there are items.
  With cboPValTable
    If .ListCount > 0 Then
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With
  
  ' Set the flag to show that the combo has been populated.
  mfPValTablesPopulated = True
  
  Exit Sub
  
ErrorTrap:
  cboPValTable.Enabled = False
  Err = False

End Sub

Private Sub cboPValReturnType_Refresh()
  ' Prompted Value component - Return Type combo.
  ' Select the current return type.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iType As ExpressionValueTypes
  
  iIndex = 0
  iType = mobjComponent.Component.valueType
  
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

Private Sub InitializeTableValueControls()

  ' Initialize the Table Value Component controls.

  Dim iCount As Integer
  Dim iTableID, iColumnID As Integer
  Dim vValue As Variant
  Dim bValueFound As Boolean

  ' Save the status of the component
  iTableID = mobjComponent.Component.TableID
  iColumnID = mobjComponent.Component.ColumnID
  vValue = mobjComponent.Component.Value

  ' Populate the Table Value Combo if it is not already populated.
  If Not mfTabValTablesPopulated Then
    cboTabValTable_Initialize
  End If

  ' Only allow the user to confirm the component definition if a valid
  ' table value is selected.
  cmdOK.Enabled = cboTabValValue.Enabled

  ' Set the dropdowns to the selected table & column
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
              'JPD 20041115 Fault 9484
              vValue = Format(vValue, DateFormat)
              'vValue = FormatDateTime(vValue, vbLongDate)
        End Select
    End With

    For iCount = 0 To cboTabValValue.ListCount - 1
        If cboTabValValue.List(iCount) = vValue Then
            cboTabValValue.Text = vValue
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
  
End Sub

Private Sub cboTabValTable_Initialize()
  ' Populate the Table Value component - Table combo.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsTables As Recordset
  
  ' Clear the contents of the combo.
  cboTabValTable.Clear
  
  ' Get the order definition.
  sSQL = "SELECT tableName, tableID" & _
    " FROM ASRSysTables" & _
    " WHERE tableType = " & Trim(Str(tabLookup)) & _
    " ORDER BY tableName"
  Set rsTables = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTables
    Do While Not .EOF
        cboTabValTable.AddItem !TableName
        cboTabValTable.ItemData(cboTabValTable.NewIndex) = !TableID
  
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsTables = Nothing
  
  ' Enable the combo if there are items.
  With cboTabValTable
    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = 0
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

Private Sub cboTabValColumn_Refresh()
  ' Populate the Table Value component - Column combo, and
  ' select the first item.
  On Error GoTo ErrorTrap
  
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim rsColumn As Recordset
  Dim iCount As Integer
  
  ' Clear the current contents of the combo.
  cboTabValColumn.Clear
  
  ' Create an array of column info.
  ' Column 1 = column ID
  ' Column 2 = data type.
  ReDim mavColumns(2, 0)
  
  If cboTabValTable.Enabled Then
  
    ' Get the columns for the currently selected lookup table.
    sSQL = "SELECT columnID, columnName, dataType" & _
      " FROM ASRSysColumns" & _
      " WHERE tableID = " & Trim(Str(cboTabValTable.ItemData(cboTabValTable.ListIndex))) & _
      " AND columnType <> " & Trim(Str(colSystem)) & _
      " AND columnType <> " & Trim(Str(colLink)) & _
      " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
      " AND dataType <> " & Trim(Str(sqlBoolean)) & _
      " AND dataType <> " & Trim(Str(sqlTypeOle))
    Set rsColumn = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsColumn
      Do While Not .EOF
        cboTabValColumn.AddItem rsColumn!ColumnName
        cboTabValColumn.ItemData(cboTabValColumn.NewIndex) = rsColumn!ColumnID

        iNextIndex = UBound(mavColumns, 2) + 1
        ReDim Preserve mavColumns(2, iNextIndex)
        mavColumns(1, iNextIndex) = rsColumn!ColumnID
        mavColumns(2, iNextIndex) = rsColumn!DataType
          
        rsColumn.MoveNext
      Loop
          
      .Close
    End With
    Set rsColumn = Nothing
        
  End If
        
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
  
  Exit Sub

ErrorTrap:
  cboTabValColumn.Enabled = False
  cboTabValValue_Refresh
  Err = False

End Sub
Private Sub cboTabValValue_Refresh()
  ' Populate the Table Value component - Value combo, and
  ' select the first item.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim iColumnType As Integer
  Dim rsLookupValues As Recordset
    
  ' Clear the current contents of the combo.
  cboTabValValue.Clear

  ' Clear the value doesn't exist label
  lblLookupValNotFound.Visible = False

  If cboTabValTable.Enabled And cboTabValColumn.Enabled Then
    sSQL = "SELECT DISTINCT " & cboTabValColumn.List(cboTabValColumn.ListIndex) & " AS lookUpValue" & _
      " FROM " & cboTabValTable.List(cboTabValTable.ListIndex) & _
      " ORDER BY lookUpValue"
      
    Set rsLookupValues = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsLookupValues
      Do While Not .EOF
        Select Case mDataType
          Case sqlNumeric, sqlInteger
            ' TM - If lookup column contains no value eg. "" then the item is not added to the Value listbox.
            If Trim(Str(!LookupValue)) <> vbNullString Then cboTabValValue.AddItem Trim(Str(!LookupValue))
          Case sqlDate
            If IsDate(!LookupValue) Then
              'JPD 20041115 Fault 9484
              'cboTabValValue.AddItem Format(!LookupValue, "long date")
              cboTabValValue.AddItem Format(!LookupValue, DateFormat)
            End If
          Case Else
            ' TM - If lookup column contains no value eg. "" then the item is not added to the Value listbox.
            If Trim(!LookupValue) <> vbNullString Then cboTabValValue.AddItem Trim(!LookupValue)
        End Select
        
        .MoveNext
      Loop
        
      .Close
    End With
    Set rsLookupValues = Nothing
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
  cmdOK.Enabled = cboTabValValue.Enabled
  Exit Sub
  
ErrorTrap:
  cboTabValValue.Enabled = False
  Resume TidyUpAndExit
  
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
  
  ' Get a list of expression operators.
  
  For Each objOperatorDef In gobjOperatorDefs
    ' Add a category node if required.
    sCategory = objOperatorDef.Category
    fCategoryDone = False
    For iLoop = 1 To ssTreeOpOperator.Nodes.Count
      If ssTreeOpOperator.Nodes(iLoop).Key = sCategory Then
        fCategoryDone = True
        Exit For
      End If
    Next iLoop
      
    If Not fCategoryDone Then
      Set objNode = ssTreeOpOperator.Nodes.Add("OPERATOR_ROOT", tvwChild, sCategory, sCategory)
      With objNode
        .Font.Bold = True
        .Sorted = ssatSortAscending
      End With
      Set objNode = Nothing
    End If
    
    ' Add the operator node.
    sDisplayName = objOperatorDef.Name
    If Len(objOperatorDef.ShortcutKeys) > 0 Then
      sDisplayName = sDisplayName & " (" & objOperatorDef.ShortcutKeys & ")"
    End If
    
    Set objNode = ssTreeOpOperator.Nodes.Add(sCategory, tvwChild, Trim(Str(objOperatorDef.ID)), sDisplayName)
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
  Set objNode = Nothing
  ssTreeOpOperator.Refresh
  Exit Sub
  
ErrorTrap:
  ssTreeOpOperator.Enabled = False
  Resume TidyUpAndExit

End Sub

Private Sub InitializeValueControls()
  Dim sCharacterValue As String
  Dim dblNumericValue As Double
  Dim fLogicValue As Boolean
  
  'MH20010201 Fault 1576
  'Dim dDateValue As Date
  Dim dDateValue As Variant 'Date

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
        sCharacterValue = .Value
      Case giEXPRVALUE_NUMERIC
        dblNumericValue = .Value
      Case giEXPRVALUE_LOGIC
        fLogicValue = .Value
      Case giEXPRVALUE_DATE
        dDateValue = .Value
    End Select
  End With
  
  txtValCharacterValue.Text = sCharacterValue
  TDBValNumericValue.Value = dblNumericValue
  
  optValLogicValue(0).Value = fLogicValue
  optValLogicValue(1).Value = Not optValLogicValue(0).Value
  
  'MH20010201 Fault 1576
  'asrValDateValue.Text = dDateValue
  asrValDateValue.Text = IIf(IsNull(dDateValue), vbNullString, dDateValue)
  
  ' Ensure the user can confirm the component definition.
  cmdOK.Enabled = True

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
    listCalcFilter_Initialize
  End If
  
  ' Select the current calculation in the list box.
  listCalcFilter_Refresh
  
  ' Only allow the user to confirm the component definition if a valid
  ' calculation is selected.
  cmdOK.Enabled = listCalcFilters.Enabled
    
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

Private Sub listCalcFilter_Refresh()
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
Private Sub listCalcCalculation_Initialize()
  ' Populate the Calculation component - Calculation listbox.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsCalculations As Recordset
  
  ' Clear the current contents of the listbox.
  listCalcCalculation.Clear
  
  ' Add an item to the listbox for each calculation based on the current expression's parent table.
  ' Get the order definition.
  If (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMECALCULATION) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMEFILTER) Then
    sSQL = "SELECT name, exprID" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID <> " & Trim(Str(mobjComponent.ParentExpression.ExpressionID)) & _
      " AND (type = " & Trim(Str(giEXPR_RUNTIMECALCULATION)) & ")" & _
      " AND TableID = " & Trim(Str(mobjComponent.ParentExpression.BaseTableID)) & _
      " AND parentComponentID = 0" & _
      " AND ((Username = '" & Replace(gsUserName, "'", "''") & "'" & _
      " OR access <> '" & ACCESS_HIDDEN & "'))"
  Else
    sSQL = "SELECT name, exprID" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID <> " & Trim(Str(mobjComponent.ParentExpression.ExpressionID)) & _
      " AND ((type = " & Trim(Str(giEXPR_COLUMNCALCULATION)) & ")" & _
      " OR (type = " & Trim(Str(giEXPR_STATICFILTER)) & ")" & _
      " OR (type = " & Trim(Str(giEXPR_RECORDDESCRIPTION)) & ")" & _
      " OR (type = " & Trim(Str(giEXPR_RECORDVALIDATION)) & "))" & _
      " AND TableID = " & Trim(Str(mobjComponent.ParentExpression.BaseTableID)) & _
      " AND parentComponentID = 0" & _
      " AND ((Username = '" & Replace(gsUserName, "'", "''") & "'" & _
      " OR access <> '" & ACCESS_HIDDEN & "'))"
  End If
  
  ' RH 09/11/00 - BUG 1316
  'JPD 20050812 Fault 10166
  If Me.chkOnlyMine.Value Then sSQL = sSQL & " AND Username = '" & Replace(gsUserName, "'", "''") & "'"
  
  Set rsCalculations = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsCalculations
    Do While Not .EOF
      listCalcCalculation.AddItem !Name
      listCalcCalculation.ItemData(listCalcCalculation.NewIndex) = !ExprID
      
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsCalculations = Nothing
  
  ' Enable the combo if there are items.
  With listCalcCalculation
    If .ListCount > 0 Then
      .Enabled = True
      CmdEditCalculation.Enabled = True
    Else
      .Enabled = False
      CmdEditCalculation.Enabled = False
    End If
  End With

  ' Set the flag to show that the listbox has been populated.
  mfCalculationsPopulated = True

TidyUpAndExit:
  listCalcCalculation.Refresh
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Sub listCalcFilter_Initialize()
  ' Populate the Calculation component - Calculation listbox.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsFilters As Recordset
  
  ' Clear the current contents of the listbox.
  listCalcFilters.Clear
  
  ' Add an item to the listbox for each calculation based on the current expression's parent table.
  ' Get the order definition.
  If (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMECALCULATION) Or _
    (mobjComponent.ParentExpression.ExpressionType = giEXPR_RUNTIMEFILTER) Then
    sSQL = "SELECT name, exprID" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID <> " & Trim(Str(mobjComponent.ParentExpression.ExpressionID)) & _
      " AND (type = " & Trim(Str(giEXPR_RUNTIMEFILTER)) & ")" & _
      " AND TableID = " & Trim(Str(mobjComponent.ParentExpression.BaseTableID)) & _
      " AND parentComponentID = 0" & _
      " AND ((Username = '" & Replace(gsUserName, "'", "''") & "'" & _
      " OR access <> '" & ACCESS_HIDDEN & "'))"

      'JPD 20050812 Fault 10166
      If Me.chkOnlyMyFilters.Value Then sSQL = sSQL & " AND Username = '" & Replace(gsUserName, "'", "''") & "'"

  End If
  
  Set rsFilters = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsFilters
    Do While Not .EOF
      listCalcFilters.AddItem !Name
      listCalcFilters.ItemData(listCalcFilters.NewIndex) = !ExprID
      
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsFilters = Nothing
  
  ' Enable the combo if there are items.
  With listCalcFilters
    If .ListCount > 0 Then
      .Enabled = True
      cmdEditFilter.Enabled = True
    Else
      .Enabled = False
      cmdEditFilter.Enabled = False
    End If
  End With

  ' Set the flag to show that the listbox has been populated.
  mfFiltersPopulated = True

TidyUpAndExit:
  listCalcFilters.Refresh
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
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
  
  Exit Sub
  
ErrorTrap:
  If ssTreeFuncFunction.Nodes.Count > 0 Then
    ssTreeFuncFunction.SelectedItem = ssTreeFuncFunction.Nodes(1)
  End If
  
  cmdOK.Enabled = False
  
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
  Dim fValid As Boolean
  
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
  
  For Each objFunctionDef In gobjFunctionDefs
    fValid = ((mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMECALCULATION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMEFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHJOINEXPRESSION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHWHEREEXPRESSION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHSCOREEXPRESSION) And _
      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RECORDINDEPENDANTCALC)) Or _
      (objFunctionDef.Runtime) Or _
      (objFunctionDef.UDF And (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RECORDINDEPENDANTCALC))
    
    If fValid Then
      fValid = (InStr(" " & objFunctionDef.ExcludeTypes & " ", " " & CStr(mobjComponent.ParentExpression.ExpressionType) & " ") = 0) _
        And ((Len(objFunctionDef.IncludeTypes) = 0) Or (InStr(" " & objFunctionDef.IncludeTypes & " ", " " & CStr(mobjComponent.ParentExpression.ExpressionType) & " ") > 0))
    End If
    
    If fValid Then
      Select Case objFunctionDef.ID

        Case 30, 46, 47 ' Absence Duration, Working Days Between Two Dates, Absence Between Two Date
          fValid = (glngPersonnelTableID > 0) _
            And (mobjComponent.ParentExpression.BaseTableID = glngPersonnelTableID _
              Or IsChildOfTable(glngPersonnelTableID, mobjComponent.ParentExpression.BaseTableID))

        Case 73 ' Bradford Factor
          fValid = IsModuleEnabled(modAbsence) _
            And (glngPersonnelTableID > 0) _
            And (mobjComponent.ParentExpression.BaseTableID = glngPersonnelTableID _
              Or IsChildOfTable(glngPersonnelTableID, mobjComponent.ParentExpression.BaseTableID))

        Case 62, 63 ' Parental Leave Entitlement, Parental Leave Taken
          ' Invalidated already as these are not runtime/udf functions and so can't
          ' be used in SecMgr
          
        Case 64 ' Maternity Return Date
          ' Invalidated already as this is not a runtime/udf function and so can't
          ' be used in SecMgr
          
        Case 66, 70 'Is Post That Reports To Current User, Is Post That Current User Reports To
          fValid = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RECORDINDEPENDANTCALC)
          
          If fValid Then
            fValid = (glngHierarchyTableID > 0) _
              And (mobjComponent.ParentExpression.BaseTableID = glngHierarchyTableID) _
              And modGeneral.HierarchyFunctionConfigured(objFunctionDef.ID)
          End If
           
        Case 68, 72 'Is Personnel That Reports To Current User, Is Personnel That Current User Reports To
          fValid = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RECORDINDEPENDANTCALC)
          
          If fValid Then
            fValid = (glngPersonnelTableID > 0) _
              And (mobjComponent.ParentExpression.BaseTableID = glngPersonnelTableID) _
              And modGeneral.HierarchyFunctionConfigured(objFunctionDef.ID)
          End If
           
      End Select
    End If

    If fValid Then
      sSPName = LCase(objFunctionDef.SPName)
      sDisplayName = objFunctionDef.Name
      If Len(objFunctionDef.ShortcutKeys) > 0 Then
        sDisplayName = sDisplayName & " " & objFunctionDef.ShortcutKeys
      End If
      
      ' Add a category node if required.
      sCategory = objFunctionDef.Category
      fCategoryDone = False
      For iLoop = 1 To ssTreeFuncFunction.Nodes.Count
        If ssTreeFuncFunction.Nodes(iLoop).Key = sCategory Then
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
      
      ' Add the function node.
      Set objNode = ssTreeFuncFunction.Nodes.Add(sCategory, tvwChild, Trim(Str(objFunctionDef.ID)), sDisplayName)
      Set objNode = Nothing
    End If
  Next objFunctionDef
  Set objFunctionDef = Nothing
  
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
  Set objNode = Nothing
  ssTreeFuncFunction.Refresh
  Exit Sub
  
ErrorTrap:
  ssTreeFuncFunction.Enabled = False
  Resume TidyUpAndExit

End Sub




Private Sub InitializeFieldControls()
  ' Initialize the Field Component controls.
  Dim objFieldComponent As clsExprField
  
  Set objFieldComponent = mobjComponent.Component
  
  optField_Refresh
    
  ' Select the current record line number value.
  asrFldSelLine.Text = Trim(Str(objFieldComponent.SelectionLine))
  
  ' Disassociate object variables.
  Set objFieldComponent = Nothing
  
End Sub

Private Sub cboFldTable_Refresh()
  ' Populate the Field component Table combo and
  ' select the current table if it is still valid.
  Dim fOK As Boolean
  Dim fTableOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim lngTableID As Long
  Dim lngRootTableID As Long
  Dim sSQL As String
  Dim rsTables As Recordset
  
  ' Determine if the field component is passed by value.
  lngTableID = mobjComponent.Component.TableID
  lngRootTableID = mobjComponent.ParentExpression.BaseTableID
  iIndex = -1
  iDefaultIndex = -1
  
  ' Clear the current contents of the combo.
  cboFldTable.Clear
  
  If (mobjComponent.ParentExpression.ExpressionType = giEXPR_UTILRUNTIMEFILTER) And _
    mfFieldByValue Then
    ' Only do the selected column's tables.
    If UBound(mobjComponent.ParentExpression.ColumnList) > 0 Then
      sSQL = "SELECT DISTINCT ASRSysTables.tableName, ASRSysTables.tableID" & _
        " FROM ASRSysTables" & _
        " INNER JOIN ASRSysColumns ON ASRSysTables.tableID = ASRSysColumns.tableID" & _
        " WHERE ASRSysColumns.columnID IN ("
        
      For iLoop = 1 To UBound(mobjComponent.ParentExpression.ColumnList)
        sSQL = sSQL & IIf(iLoop > 1, ",", "") & _
          Trim(Str(mobjComponent.ParentExpression.ColumnList(iLoop)))
      Next iLoop
      
      sSQL = sSQL & ")" & _
        " ORDER BY ASRSysTables.tableName"
    
      Set rsTables = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsTables
        Do While Not .EOF
          ' Add the table to the combo, and check if its the currently selected table,
          ' or the expression's parent table.
          cboFldTable.AddItem !TableName
          cboFldTable.ItemData(cboFldTable.NewIndex) = !TableID
          
          If !TableID = lngTableID Then
            iIndex = cboFldTable.NewIndex
          End If
          
          If !TableID = lngRootTableID Then
            iDefaultIndex = cboFldTable.NewIndex
          End If
          
          .MoveNext
        Loop
      
        .Close
      End With
      Set rsTables = Nothing
    End If
  Else
    ' Get the required tables.
    If mfFieldByValue Then
      If optField(0).Value Then
        sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
          " FROM ASRSysTables" & _
          " WHERE ASRSysTables.tableId = " & Trim(Str(lngRootTableID))
      
        ' Only add parent and child tables to the combo if the current expression is not
        ' a View Filter.
        If mobjComponent.ParentExpression.ExpressionType <> giEXPR_VIEWFILTER Then
          sSQL = sSQL & _
            " UNION" & _
            " SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
            " FROM ASRSysTables" & _
            " JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.parentID" & _
            " WHERE ASRSysRelations.childID = " & Trim(Str(lngRootTableID)) & _
            " UNION" & _
            " SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
            " FROM ASRSysTables" & _
            " JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID" & _
            " WHERE ASRSysRelations.parentID = " & Trim(Str(lngRootTableID))
        End If
      Else
        sSQL = " SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
                " FROM ASRSysTables" & _
                " JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID" & _
                " WHERE ASRSysRelations.parentID = " & Trim(Str(lngRootTableID))
      End If
    Else
      sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
        " FROM ASRSysTables"
    End If
  
    sSQL = sSQL & " ORDER BY ASRSysTables.tableName"
    
    Set rsTables = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsTables
      Do While Not .EOF
        ' Add the table to the combo, and check if its the currently selected table,
        ' or the expression's parent table.
        cboFldTable.AddItem !TableName
        cboFldTable.ItemData(cboFldTable.NewIndex) = !TableID
        
        If !TableID = lngTableID Then
          iIndex = cboFldTable.NewIndex
        End If
        
        If !TableID = lngRootTableID Then
          iDefaultIndex = cboFldTable.NewIndex
        End If
        
        .MoveNext
      Loop
    
      .Close
    End With
    Set rsTables = Nothing
  End If
    
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
      
      cboFldColumn_Refresh
      fldSelOrder_Refresh
      fldSelFilter_Refresh
    End If
  End With
    
End Sub
Private Sub fldSelFilter_Refresh()
  ' Refresh the Field Selection Filter controls.
  ' Validate the expression selection at the same time.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sFilterName As String
  Dim rsFilter As Recordset
  
  fOK = True
  
  ' Check if the selected filter is for the current table.
  sSQL = "SELECT name" & _
    " FROM ASRSysExpressions" & _
    " WHERE exprID = " & Trim(Str(mobjComponent.Component.SelectionFilterID)) & _
    " AND TableID = " & Trim(Str(mobjComponent.Component.TableID))
  Set rsFilter = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsFilter
    fOK = Not (.EOF And .BOF)
    
    If fOK Then
      sFilterName = !Name
    End If
    
    .Close
  End With
  
TidyUpAndExit:
  Set rsFilter = Nothing
  If Not fOK Then
    mobjComponent.Component.SelectionFilterID = 0
    sFilterName = ""
  End If
  ' Update the control's properties.
  txtFldSelFilter.Text = sFilterName
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub fldSelOrder_Refresh()
  ' Refresh the Field Selection Order controls.
  ' Validate the order selection at the same time.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sOrderName As String
  Dim sSQL As String
  Dim rsOrder As Recordset
  
  ' Check if the selected order is for the current table.
  sSQL = "SELECT name" & _
    " FROM ASRSysOrders" & _
    " WHERE orderID = " & Trim(Str(mobjComponent.Component.SelectionOrderID)) & _
    " AND tableID = " & Trim(Str(mobjComponent.Component.TableID))
  Set rsOrder = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsOrder
    fOK = Not (.EOF And .BOF)
    
    If fOK Then
      sOrderName = !Name
    End If
    
    .Close
  End With

TidyUpAndExit:
  Set rsOrder = Nothing
  If Not fOK Then
    mobjComponent.Component.SelectionOrderID = 0
    sOrderName = ""
  End If
  
  ' Update the control's properties.
  txtFldSelOrder.Text = sOrderName
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub






Private Sub cboFldColumn_Refresh()
  ' Populate the Field component Column combo, and then select the current field.
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim rsColumns As Recordset
  
  ' Clear the current contents of the combo.
  cboFldColumn.Clear
  
  ' Create an array of column info.
  ' Column 1 = column ID
  ' Column 2 = data type.
  ReDim mavColumns(2, 0)
  
  sSQL = "SELECT columnName, columnID, dataType" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & Trim(Str(mobjComponent.Component.TableID)) & _
    " AND dataType <> " & Trim(Str(sqlTypeOle)) & _
    " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
    " AND columnType <> " & Trim(Str(colLink)) & _
    " AND columnType <> " & Trim(Str(colSystem))
    
  If optField(2).Value Then
    ' Total - must be a numeric field.
    sSQL = sSQL & _
      " AND (dataType = " & Trim(Str(sqlNumeric)) & _
      " OR dataType = " & Trim(Str(sqlInteger)) & ")"
  End If
  
  If (mobjComponent.ParentExpression.ExpressionType = giEXPR_UTILRUNTIMEFILTER) And _
    mfFieldByValue Then
    ' Only do the selected utility columns.
    sSQL = sSQL & " AND columnID IN("
    
    If UBound(mobjComponent.ParentExpression.ColumnList) = 0 Then
      sSQL = sSQL & "0"
    Else
      For iLoop = 1 To UBound(mobjComponent.ParentExpression.ColumnList)
        sSQL = sSQL & IIf(iLoop > 1, ",", "") & _
          Trim(Str(mobjComponent.ParentExpression.ColumnList(iLoop)))
      Next iLoop
    End If
    
    sSQL = sSQL & ")"
  End If
    
  sSQL = sSQL & _
    " ORDER BY columnName"
  Set rsColumns = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsColumns
    Do While Not .EOF
      ' Add the column to the combo, and check if its the currently selected column.
      cboFldColumn.AddItem !ColumnName
      cboFldColumn.ItemData(cboFldColumn.NewIndex) = !ColumnID
      
      iNextIndex = UBound(mavColumns, 2) + 1
      ReDim Preserve mavColumns(2, iNextIndex)
      mavColumns(1, iNextIndex) = !ColumnID
      mavColumns(2, iNextIndex) = !DataType
      
      If !ColumnID = mobjComponent.Component.ColumnID Then
        iIndex = cboFldColumn.NewIndex
      End If
      
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsColumns = Nothing
  
  ' Enable the combo if there are items.
  With cboFldColumn
    If .ListCount > 0 Then
      .ListIndex = iIndex
      .Enabled = True
    Else
      .Enabled = False
      optFieldSel_Refresh
      
      If optField(2).Value Then
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
  Dim fFieldVisible As Boolean
  Dim fFunctionVisible As Boolean
  Dim fCalculationVisible As Boolean
  Dim fValueVisible As Boolean
  Dim fOperatorVisible As Boolean
  Dim fTableValueVisible As Boolean
  Dim fPromptedValueVisible As Boolean
  Dim fCustomCalculationVisible As Boolean
  Dim fFilterVisible As Boolean
  Dim dblYCoord As Double
  
  Const YSTART = 300
  Const YGAP = 350
  
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
    
  fFieldVisible = True
  fFunctionVisible = True
  fCalculationVisible = True
  fValueVisible = True
  fOperatorVisible = True
  fTableValueVisible = True
  fPromptedValueVisible = True
  fCustomCalculationVisible = False
  fFilterVisible = True
    
  ' Disable some component types for some expression types.
  Select Case mobjComponent.ParentExpression.ExpressionType
    Case giEXPR_COLUMNCALCULATION
      fPromptedValueEnabled = False
  
    Case giEXPR_GOTFOCUS
      ' Not used.
  
    Case giEXPR_RECORDVALIDATION
      fCalculationEnabled = False
      fPromptedValueEnabled = False
  
    Case giEXPR_DEFAULTVALUE
  
    Case giEXPR_STATICFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
  
    Case giEXPR_PAGEBREAK
      ' Not used.
    Case giEXPR_ORDER
      ' Not used.
      
    Case giEXPR_RECORDDESCRIPTION
      fCalculationEnabled = False
      fPromptedValueEnabled = False
  
    Case giEXPR_VIEWFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
    
    Case giEXPR_RUNTIMECALCULATION
'      fCalculationEnabled = False
      fPromptedValueEnabled = True
  
    Case giEXPR_RUNTIMEFILTER
'      fCalculationEnabled = False
      fPromptedValueEnabled = True
  
    Case giEXPR_UTILRUNTIMEFILTER
      fCalculationEnabled = False
      fPromptedValueEnabled = False
      fFilterEnabled = False
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
  End If
  
  ' Disable and hide controls as required.
  dblYCoord = YSTART
  
  optComponentType(giCOMPONENT_FIELD).Enabled = fFieldEnabled
  optComponentType(giCOMPONENT_FIELD).Visible = fFieldVisible
  If fFieldVisible Then
    optComponentType(giCOMPONENT_FIELD).Top = dblYCoord
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

  ' JDM - 12/03/01 - Added Filter option
  optComponentType(giCOMPONENT_FILTER).Enabled = fFilterEnabled
  optComponentType(giCOMPONENT_FILTER).Visible = fFilterVisible
  If fFilterVisible Then
    optComponentType(giCOMPONENT_FILTER).Top = dblYCoord
    dblYCoord = dblYCoord + YGAP
  End If
 
End Function


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

'  ' JDM - 10/08/01 -  Fault 2672 - Warn user of duff date (P.S. It's my birthday today)
'  If IsNull(asrPValDefaultDate.DateValue) And Not _
'     IsDate(asrPValDefaultDate.DateValue) And _
'     asrPValDefaultDate.Text <> "  /  /" Then
'
'     MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     asrPValDefaultDate.DateValue = Null
'     asrPValDefaultDate.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
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
  
  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate asrValDateValue

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

Private Sub FormatFieldControls()
  ' Display the required Field Component controls.
  Dim fIsChildOfBase As Boolean
  Dim sSQL As String
  Dim rsRelation As Recordset
  
  ' Disable the column combo if 'COUNT' is selected.
  If cboFldColumn.Enabled Then
    cboFldColumn.Enabled = Not optField(1).Value
  End If
  cboFldDummyColumn.Visible = optField(1).Value
  cboFldColumn.BackColor = IIf(cboFldColumn.Enabled, vbWhite, vbButtonFace)
  lblFldField.Enabled = Not optField(1).Value
  
  If (mfFieldByValue And (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER)) Then
    ' Only enable the field selection options if the selected table
    ' is a child of the expression's base table.
    sSQL = "SELECT *" & _
      " FROM ASRSysRelations" & _
      " WHERE parentID = " & Trim(Str(mobjComponent.ParentExpression.BaseTableID)) & _
      " AND childID = " & Trim(Str(mobjComponent.Component.TableID))
    Set rsRelation = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsRelation
      fIsChildOfBase = Not (.EOF And .BOF)
    
      .Close
    End With
    Set rsRelation = Nothing

    fraFldSelOptions.Enabled = fIsChildOfBase
    
    optFieldSel(0).Enabled = (fIsChildOfBase And optField(0).Value)
    optFieldSel(1).Enabled = (fIsChildOfBase And optField(0).Value)
    optFieldSel(2).Enabled = (fIsChildOfBase And optField(0).Value) 'And _
'      ((mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMECALCULATION) And _
'        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMEFILTER) And _
'        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_LINKFILTER) And _
'        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
'        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHJOINEXPRESSION) And _
'        (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHSCOREEXPRESSION))
    asrFldSelLine.Enabled = optFieldSel(2).Enabled
    
    If (Not optFieldSel(1).Enabled) And (optFieldSel(1).Value) Then
      optFieldSel(0).Value = True
    End If
    If (Not optFieldSel(2).Enabled) And (optFieldSel(2).Value) Then
      optFieldSel(0).Value = True
    End If

    If (Not optFieldSel(2).Value) Then
      asrFldSelLine.Text = vbNullString
    End If
    
    ' Only enable the line number control if required.
    asrFldSelLine.Enabled = (fIsChildOfBase And optFieldSel(2).Value)
    asrFldSelLine.BackColor = IIf(asrFldSelLine.Enabled, vbWhite, vbButtonFace)
    
    lblFldOrder.Enabled = (fIsChildOfBase And optField(0).Value)
    cmdFldSelOrder.Enabled = (fIsChildOfBase And optField(0).Value)
    If Not cmdFldSelOrder.Enabled Then
      mobjComponent.Component.SelectionOrderID = 0
      txtFldSelOrder.Text = ""
    End If
    
    lblFldFilter.Enabled = fIsChildOfBase
    cmdFldSelFilter.Enabled = fIsChildOfBase
    If Not cmdFldSelFilter.Enabled Then
      mobjComponent.Component.SelectionFilterID = 0
      txtFldSelFilter.Text = ""
    End If
  End If
  
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
  mobjComponent.Component.valueType = cboPValReturnType.ItemData(cboPValReturnType.ListIndex)
  
  ' Display only the required controls.
  FormatPromptedValueControls

End Sub


Private Sub FormatPromptedValueControls()
  ' Display only the required Prompted Value Component controls.
  Dim fSizeVisible As Boolean
  Dim fDecimalsVisible As Boolean
  Dim fFormatVisible As Boolean
  Dim fFormatEnabled As Boolean
  Dim fLookupTableEnabled As Boolean
  Dim iCount As Integer
  
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
  
  For iCount = 0 To optPValDefaultDateType.Count - 1
    optPValDefaultDateType(iCount).Visible = False
  Next iCount
 
  ' Conditionally display some controls.
  Select Case cboPValReturnType.ItemData(cboPValReturnType.ListIndex)
    Case giEXPRVALUE_CHARACTER
      fSizeVisible = True
      fFormatEnabled = True
      fFormatVisible = True
      txtPValDefaultCharacter.Visible = True
      
    Case giEXPRVALUE_NUMERIC
      fSizeVisible = True
      fDecimalsVisible = True
      fFormatVisible = True
      TDBPValDefaultNumeric.Visible = True

    Case giEXPRVALUE_LOGIC
      optPValDefaultLogic(0).Visible = True
      optPValDefaultLogic(1).Visible = True
      fFormatVisible = True
  
    Case giEXPRVALUE_DATE
      asrPValDefaultDate.Visible = True
      For iCount = 0 To optPValDefaultDateType.Count - 1
        optPValDefaultDateType(iCount).Visible = True
      Next iCount
      fFormatVisible = True
    
    Case giEXPRVALUE_TABLEVALUE
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
        Case sqlNumeric
          mobjComponent.Component.ReturnType = giEXPRVALUE_NUMERIC
        Case sqlInteger
          mobjComponent.Component.ReturnType = giEXPRVALUE_NUMERIC
        Case sqlDate
          mobjComponent.Component.ReturnType = giEXPRVALUE_DATE
        Case sqlVarchar, sqlLongVarChar
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
        .Value = Val(cboTabValValue.List(cboTabValValue.ListIndex))
    
      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(cboTabValValue.List(cboTabValValue.ListIndex)) Then
          .Value = CDate(cboTabValValue.List(cboTabValValue.ListIndex))
        Else
          .Value = 0
        End If
  
      Case Else
        .Value = cboTabValValue.List(cboTabValValue.ListIndex)
    End Select
  End With

End Sub


Private Sub cboValType_Click()
  ' Update the component object with the new value.
  mobjComponent.Component.ReturnType = cboValType.ItemData(cboValType.ListIndex)
    
  ' Display only the required controls.
  FormatValueControls

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

Private Sub chkOnlyMine_Click()

  Dim lOldCalcID As Long
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  If mblnLoading Then Exit Sub
  
  ' Store the previously selected calc
  lOldCalcID = mobjComponent.Component.CalculationID
  
  ' Refresh the list of calcs according to the checkbox state
  listCalcCalculation_Initialize
  
  ' Now reselect the previously selected calc if its there, if not, select the
  ' first one
  If listCalcCalculation.Enabled Then
    iIndex = 0
    For iLoop = 0 To listCalcCalculation.ListCount - 1
      If listCalcCalculation.ItemData(iLoop) = lOldCalcID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    listCalcCalculation.ListIndex = iIndex
  End If

  cmdOK.Enabled = listCalcCalculation.ListCount > 0

End Sub

Private Sub chkOnlyMyFilters_Click()

  Dim lOldFilterID As Long
  Dim iLoop As Integer
  Dim iIndex As Integer
  
  If mblnLoading Then Exit Sub
  
  ' Store the previously selected calc
  lOldFilterID = mobjComponent.Component.FilterID
  
  ' Refresh the list of calcs according to the checkbox state
  listCalcFilter_Initialize
  
  ' Now reselect the previously selected filter if its there, if not, select the
  ' first one
  If listCalcFilters.Enabled Then
    iIndex = 0
    For iLoop = 0 To listCalcFilters.ListCount - 1
      If listCalcFilters.ItemData(iLoop) = lOldFilterID Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    listCalcFilters.ListIndex = iIndex
  End If

  cmdOK.Enabled = listCalcFilters.ListCount > 0


End Sub

Private Sub cmdCancel_Click()
  ' Set the cancelled flag.
  mfCancelled = True
  
  ' Unload the form.
  Unload Me

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
      
    Case giCOMPONENT_CALCULATION
      SaveComponent = SaveCalc
    
    Case giCOMPONENT_FILTER
      SaveComponent = SaveFilter

    Case giCOMPONENT_TABLEVALUE
      SaveComponent = SaveTableValue

  End Select
  
End Function

Private Function SaveCalc() As Boolean

  Dim fSaveOK As Boolean
  Dim sOwner As String
  Dim lngExpressionID As Long
  
  fSaveOK = True

  If Me.listCalcCalculation.ListCount = 0 Then
    fSaveOK = False
    MsgBox "You must select a calculation.", vbExclamation + vbOKOnly, App.Title
  End If

  'TM20011003 Fault 2656
  'Only allow the component to be added if has not been made hidden or is
  'owned by the expression owner OR has since been deleted.
  If ValidComponent(mobjComponent, True) > 1 Then
    fSaveOK = False
    listCalcCalculation_Initialize
  
    'JPD 20030728 Fault 6479
    listCalcCalculation_Refresh
End If
  
'  lngExpressionID = mobjComponent.Component.CalculationID
'  sOwner = GetExprField(lngExpressionID, "Username")
'  If HasHiddenComponents(lngExpressionID) And LCase(sOwner) <> LCase(gsUserName) Then
'    fSaveOK = False
'    MsgBox "The selected expression contains hidden components and is owned by another user." & vbCrLf & vbCrLf & "The expression will now be made hidden.", vbExclamation + vbOKOnly, App.Title
'    listCalcCalculation_Initialize
'  End If
  
  SaveCalc = fSaveOK
  
End Function

Private Function SaveFilter() As Boolean

  Dim fSaveOK As Boolean
  Dim sOwner As String
  Dim lngExpressionID As Long
  
  fSaveOK = True

  If Me.listCalcFilters.ListCount = 0 Then
    fSaveOK = False
    MsgBox "You must select a filter.", vbExclamation + vbOKOnly, App.Title
  End If

  'TM20011003 Fault 2656
  'Only allow the component to be added if has not been made hidden or is
  'owned by the expression owner OR has since been deleted.
  If ValidComponent(mobjComponent, True) > 1 Then
    fSaveOK = False
    listCalcFilter_Initialize
  
    'JPD 20030728 Fault 6479
    listCalcFilter_Refresh
  End If

'  lngExpressionID = mobjComponent.Component.FilterID
'  sOwner = GetExprField(lngExpressionID, "Username")
'  If HasHiddenComponents(lngExpressionID) And LCase(sOwner) <> LCase(gsUserName) Then
'    fSaveOK = False
'    MsgBox "The selected expression contains hidden components and is owned by another user." & vbCrLf & vbCrLf & "The expression will now be made hidden.", vbExclamation + vbOKOnly, App.Title
'    listCalcFilter_Initialize
'  End If
  
  SaveFilter = fSaveOK
  
End Function

Private Function SavePromptedValue() As Boolean
  ' Update the component object
  Dim vValidatedDate As Variant
  Dim fSaveOK As Boolean
  Dim dtDateValue As Date
  
  fSaveOK = True

  With mobjComponent.Component
    Select Case .valueType
      Case giEXPRVALUE_CHARACTER
        .DefaultValue = txtPValDefaultCharacter.Text
  
      Case giEXPRVALUE_NUMERIC
        .DefaultValue = TDBPValDefaultNumeric.Value

      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(asrPValDefaultDate.Text) Then
          '.DefaultValue = asrPValDefaultDate.Value
          .DefaultValue = asrPValDefaultDate.Text
        Else
          .DefaultValue = Null
        End If
        
      Case giEXPRVALUE_LOGIC
        .DefaultValue = optPValDefaultLogic(0).Value
        
      Case giEXPRVALUE_TABLEVALUE
        Select Case mDataType
          Case sqlNumeric, sqlInteger
            .DefaultValue = Val(cboPValDefaultTabVal.List(cboPValDefaultTabVal.ListIndex))
          Case sqlDate
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





Private Function SaveValue() As Boolean
  ' Update the component object
  Dim vValidatedDate As Variant
  Dim fSaveOK As Boolean
  
  fSaveOK = True
  
  With mobjComponent.Component
    Select Case .ReturnType
    
      Case giEXPRVALUE_CHARACTER
        .Value = txtValCharacterValue.Text
  
      Case giEXPRVALUE_NUMERIC
        .Value = TDBValNumericValue.Value

      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        
        'MH20010201
        'Temporarily stop them entering null dates.
        'This is because is causes a problem when comparing to a date
        'e.g "Leaving_date = null" should be "Leaving_date is null"
        If IsDate(asrValDateValue.DateValue) Then
        'If IsDate(asrValDateValue.DateValue) Or IsNull(asrValDateValue.DateValue) Then
          .Value = asrValDateValue.DateValue
        Else
          'MsgBox "Invalid date.", vbOKOnly + vbExclamation, App.ProductName
          MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.ProductName
          
          asrValDateValue.SetFocus
          fSaveOK = False
        End If
  
      Case giEXPRVALUE_LOGIC
        .Value = optValLogicValue(0).Value
    End Select
  End With
  
  SaveValue = fSaveOK
  
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
        Case sqlVarchar, sqlLongVarChar
          mobjComponent.Component.ReturnType = giEXPRVALUE_CHARACTER
      End Select
      
      Exit For
    End If
  Next iLoop

  ' Update the component.
  With mobjComponent.Component
    Select Case .ReturnType
      Case giEXPRVALUE_NUMERIC
        .Value = Val(cboTabValValue.List(cboTabValValue.ListIndex))
    
      Case giEXPRVALUE_DATE
        ' Validate the entered date.
        If IsDate(cboTabValValue.List(cboTabValValue.ListIndex)) Then
          .Value = CDate(cboTabValValue.List(cboTabValValue.ListIndex))
        Else
          .Value = 0
        End If
  
      Case Else
        .Value = cboTabValValue.List(cboTabValValue.ListIndex)
    End Select
  End With

  SaveTableValue = fSaveOK
  
End Function






Private Function SaveOperator() As Boolean
  ' Write the selected Operator ID to the component.
  Dim lngOperatorID As Long
  
  With ssTreeOpOperator
    If .SelectedNodes.Count > 0 Then
      lngOperatorID = .SelectedItem.Key
    Else
      lngOperatorID = 0
    End If
  End With
  
  mobjComponent.Component.OperatorID = lngOperatorID
  
  SaveOperator = True
  
End Function

Private Function SaveField() As Boolean
  ' Validate the field component definition.
  
  If ValidComponent(mobjComponent, True) > 1 Then
    ' Field filter is no longer valid.
    SaveField = False
    txtFldSelFilter.Text = ""
    txtFldSelFilter.Tag = 0
    mobjComponent.Component.SelectionFilterID = 0
    
    Exit Function
  End If

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


Private Function SaveFunction() As Boolean
  ' Write the selected Function ID to the component.
  Dim lngFunctionID As Long
  
  With ssTreeFuncFunction
    If .SelectedNodes.Count > 0 Then
      lngFunctionID = .SelectedItem.Key
    Else
      lngFunctionID = 0
    End If
  End With
  
  mobjComponent.Component.FunctionID = lngFunctionID
  
  SaveFunction = True
  
End Function

Private Sub cmdEditFilter_Click()

  Dim lExprID As Long
  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lExprID = listCalcFilters.ItemData(listCalcFilters.ListIndex)
  
  If lExprID > 0 Then
    Set objExpr = New clsExprExpression
    With objExpr
      .ExpressionID = lExprID
      .EditExpression
      
      ' Refresh the filter list with the name taken from the edit screen.
      'listCalcFilters.List(listCalcFilters.ListIndex) = .Name
      listCalcFilter_Initialize
      listCalcFilter_Refresh

    End With
    Set objExpr = Nothing
  End If

End Sub

Private Sub cmdFldSelFilter_Click()
  ' Display the 'Field Selection Filter' expression selection form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objExpr As clsExprExpression
  Dim sSQL As String
  Dim rsExpressions As Recordset
  Dim lngOldExpressionID As Long

  fOK = True
  
  lngOldExpressionID = mobjComponent.Component.SelectionFilterID
  
  ' Instantiate an expression object.
  Set objExpr = New clsExprExpression
  
  With objExpr
    ' Set the properties of the expression object.
    If (mobjComponent.ParentExpression.ExpressionType = giEXPR_COLUMNCALCULATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDVALIDATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDDESCRIPTION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_STATICFILTER) Then
      .ExpressionType = giEXPR_STATICFILTER
    Else
      .ExpressionType = giEXPR_RUNTIMEFILTER
    End If
    .BaseTableID = mobjComponent.Component.TableID
    .ExpressionID = mobjComponent.Component.SelectionFilterID
    .ReturnType = giEXPRVALUE_LOGIC
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression(True) Then
      If .Access = ACCESS_HIDDEN Then
        If LCase(mobjComponent.ParentExpression.Owner) <> LCase(gsUserName) Then
          MsgBox "Unable to select this filter as it is a hidden filter and you are not the owner of this expression.", vbExclamation + vbOKOnly, App.Title
          If .ExpressionID = mobjComponent.Component.SelectionFilterID Or (.ExpressionID = 0) Then
            txtFldSelFilter.Text = ""
            txtFldSelFilter.Tag = 0
          End If
          mobjComponent.Component.SelectionFilterID = 0
          Set objExpr = Nothing
          Exit Sub
        End If
      End If
      
      mobjComponent.Component.SelectionFilterID = .ExpressionID
      txtFldSelFilter.Text = .Name
    Else
      ' Check if the original expression has been deleted.
      sSQL = sSQL & "SELECT name, access " & _
        " FROM ASRSysExpressions" & _
        " WHERE exprID = " & Trim(Str(lngOldExpressionID))
      Set rsExpressions = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
      If rsExpressions.EOF And rsExpressions.BOF Then
        mobjComponent.Component.SelectionFilterID = 0
        txtFldSelFilter.Text = ""
      Else
        If rsExpressions!Access = ACCESS_HIDDEN Then
          If LCase(mobjComponent.ParentExpression.Owner) <> LCase(gsUserName) Then
            MsgBox "Unable to select this filter as it is a hidden filter and you are not the owner of this expression.", vbExclamation + vbOKOnly, App.Title
            txtFldSelFilter.Text = ""
            txtFldSelFilter.Tag = 0
            mobjComponent.Component.SelectionFilterID = 0
            Set objExpr = Nothing
            Exit Sub
          End If
        End If
      End If
      
      rsExpressions.Close
      Set rsExpressions = Nothing
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

Private Sub cmdFldSelOrder_Click()
'  ' Display the 'Field Selection Order' selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sSQL As String
  Dim objOrder As clsOrder
  Dim rsOrders As ADODB.Recordset

  fOK = True

  ' Instantiate an order object.
  Set objOrder = New clsOrder

  With objOrder
    ' Initialize the order object.
    .OrderID = mobjComponent.Component.SelectionOrderID
    .TableID = mobjComponent.Component.TableID

    If (mobjComponent.ParentExpression.ExpressionType = giEXPR_COLUMNCALCULATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDDESCRIPTION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_RECORDVALIDATION) Or _
      (mobjComponent.ParentExpression.ExpressionType = giEXPR_STATICFILTER) Then
      .OrderType = giORDERTYPE_STATIC
    Else
      .OrderType = giORDERTYPE_DYNAMIC
    End If

    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      mobjComponent.Component.SelectionOrderID = .OrderID
      txtFldSelOrder.Text = .OrderName
    Else
      ' Check in case the original expression has been deleted.
      sSQL = "SELECT *" & _
        " FROM ASRSysOrders" & _
        " WHERE orderID = " & Trim(Str(mobjComponent.Component.SelectionOrderID))
      Set rsOrders = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsOrders
        If (.EOF And .BOF) Then
          mobjComponent.Component.SelectionOrderID = 0
          txtFldSelOrder.Text = ""
        End If

        .Close
      End With
      Set rsOrders = Nothing
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
  If ValidateGTMaskDate(asrValDateValue) = False Or _
     ValidateGTMaskDate(asrPValDefaultDate) = False Then
      Exit Sub
  End If
  'cmdOK.SetFocus
  'DoEvents
  
  ' Write the displayed control values to the component
  If SaveComponent Then
    ' Set the cancelled flag.
    mfCancelled = False
    
    ' Unload the form.
    Unload Me
  End If

End Sub

Private Sub CmdEditCalculation_Click()

  Dim lExprID As Long
  Dim objExpr As New clsExprExpression
  Dim lngOptions As Long
  
  lExprID = listCalcCalculation.ItemData(listCalcCalculation.ListIndex)
  
  If lExprID > 0 Then
    Set objExpr = New clsExprExpression
    With objExpr
      .ExpressionID = lExprID
      .EditExpression
      
      ' Refresh the calculation list with the name taken from the edit screen.
      'listCalcCalculation.List(listCalcCalculation.ListIndex) = .Name
      listCalcCalculation_Initialize
      listCalcCalculation_Refresh
    End With
    Set objExpr = Nothing
  End If
  
End Sub

Private Sub Form_Initialize()
  ' Initialize the 'cancelled' property.
  mblnApplySystemPermissions = False
  mfCancelled = True
  mfFunctionsPopulated = False
  mfCalculationsPopulated = False
  mfOperatorsPopulated = False
  mfTabValTablesPopulated = False
  mfPValTablesPopulated = False
  
  If mblnApplySystemPermissions Then
    mbEnableViewFilter = SystemPermission("FILTERS", "VIEW")
    mbEnableEditFilter = SystemPermission("FILTERS", "EDIT")
    mbEnableViewCalculation = SystemPermission("CALCULATIONS", "VIEW")
    mbEnableEditCalculation = SystemPermission("CALCULATIONS", "EDIT")
  Else
    mbEnableViewFilter = True
    mbEnableEditFilter = True
    mbEnableViewCalculation = True
    mbEnableEditCalculation = True
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


' JDM - 15/02/01 - fault 1868 - Error when pressing CTRL-X on treeview control
' For some reason the Sheridan treeview control wants to fire off it own cutn'paste functionality
' must trap it here not in it's own keydown event
If ActiveControl.Name = "ssTreeFuncFunction" Or ActiveControl.Name = "ssTreeOpOperator" Then
    
    KeyCode = 0
    Shift = 0

End If


End Sub


Private Sub Form_Load()

'  SetDateComboFormat Me.asrValDateValue
'  SetDateComboFormat Me.asrPValDefaultDate
  
  'JPD 20041115 Fault 9484
  UI.FormatGTDateControl asrValDateValue
  UI.FormatGTDateControl asrPValDefaultDate
  
  ' Format the form's frames and controls
  FormatScreen

  ' Set loading flag temporarily
  mblnLoading = True
  
  'chkOnlyMine.Value = GetPCSetting(gsDatabaseName & "\DefSel", "OnlyMine", 0)
  chkOnlyMine.Value = GetUserSetting("DefSel", "OnlyMine Calculations", 0)
  
  mblnLoading = False
  
End Sub


Private Sub FormatScreen()
  ' Position and size controls.
  Dim iLoop As Integer
  
  Const iXGAP = 200
  Const iYGAP = 200
  Const iXFRAMEGAP = 150
  Const iYFRAMEGAP = 100
  Const iCOMPONENTFRAMEWIDTH = 2200
  Const iFRAMEWIDTH = 6000
  Const iFRAMEHEIGHT = 3900
  
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
  Const lngDATECONTROLWIDTH = 1400
  
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
    .Width = 1600
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
    
    'MH20010130 Fault 1610
    FormatTDBNumberControl TDBPValDefaultNumeric
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


  ' Format the location for the prompted date values
  With optPValDefaultDateType(0)
    .Left = lngCOLUMN1 + lngDATECONTROLWIDTH + 150
    .Top = lngYCoordinate - 70
  End With

  With optPValDefaultDateType(1)
    .Left = optPValDefaultDateType(0).Left
    .Top = lngYCoordinate + optPValDefaultDateType(0).Height - 20
  End With

  With optPValDefaultDateType(2)
    .Left = optPValDefaultDateType(0).Left + optPValDefaultDateType(0).Width + 100
    .Top = optPValDefaultDateType(0).Top
  End With

  With optPValDefaultDateType(3)
    .Left = optPValDefaultDateType(2).Left
    .Top = optPValDefaultDateType(1).Top
  End With

  With optPValDefaultDateType(4)
    .Left = optPValDefaultDateType(2).Left + optPValDefaultDateType(2).Width + 100
    .Top = optPValDefaultDateType(0).Top
  End With

  With optPValDefaultDateType(5)
    .Left = optPValDefaultDateType(4).Left
    .Top = optPValDefaultDateType(1).Top
  End With

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
    
    'MH20010130 Fault 1610
    FormatTDBNumberControl TDBValNumericValue
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


Private Sub FormatCalculationFrame()
  ' Size and position the Calculation component controls.
  Const iXGAP = 200
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  With listCalcCalculation
    .Left = iXGAP
    .Top = iYTOPGAP
    .Width = fraComponent(3).Width - (2 * iXGAP)
    '.Height = fraComponent(3).Height - iYTOPGAP - iYBOTTOMGAP
    .Height = fraComponent(3).Height - iYTOPGAP - iYBOTTOMGAP - iYBOTTOMGAP
  End With
  
  With chkOnlyMine
    .Left = listCalcCalculation.Left
    .Top = listCalcCalculation.Top + listCalcCalculation.Height + (iYBOTTOMGAP / 3)
    .Width = listCalcCalculation.Width
    .Caption = Me.chkOnlyMine.Caption & "'" & gsUserName & "'"
  End With
  
End Sub
Private Sub FormatFilterFrame()
  ' Size and position the Filter component controls.
  Const iXGAP = 200
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  With listCalcFilters
    .Left = iXGAP
    .Top = iYTOPGAP
    .Width = fraComponent(3).Width - (2 * iXGAP)
    .Height = fraComponent(3).Height - iYTOPGAP - iYBOTTOMGAP - iYBOTTOMGAP
  End With
  
  With chkOnlyMyFilters
    .Left = listCalcFilters.Left
    .Top = listCalcFilters.Top + listCalcFilters.Height + (iYBOTTOMGAP / 3)
    .Width = listCalcFilters.Width
    .Caption = Me.chkOnlyMyFilters.Caption & "'" & gsUserName & "'"
  End With
  
End Sub

Private Sub FormatFunctionFrame()
  ' Size and position the Function component controls.
  Const iXGAP = 200
  Const iYTOPGAP = 300
  Const iYBOTTOMGAP = 200
  
  With ssTreeFuncFunction
    .Left = iXGAP
    .Top = iYTOPGAP
    .Width = fraComponent(4).Width - (2 * iXGAP)
    .Height = fraComponent(4).Height - iYTOPGAP - iYBOTTOMGAP
  End With

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
  
  optField(1).Enabled = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER)
  optField(2).Enabled = optField(1).Enabled
   
  'Do we display the specific line code
  optFieldSel(2).Visible = gbEnableUDFFunctions
  asrFldSelLine.Visible = gbEnableUDFFunctions
   
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
'    asrFldSelLine.Enabled = (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMECALCULATION) And _
'      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_RUNTIMEFILTER) And _
'      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) And _
'      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHJOINEXPRESSION) And _
'      (mobjComponent.ParentExpression.ExpressionType <> giEXPR_MATCHSCOREEXPRESSION)
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
Public Property Get Cancelled() As Boolean
  ' Return the Cancelled property.
  Cancelled = mfCancelled
  
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode <> vbFormCode Then
    mfCancelled = True
  End If

  'SavePCSetting gsDatabaseName & "\DefSel", "OnlyMine", Abs(chkOnlyMine.Value)
  'SaveUserSetting "DefSel", "OnlyMine Calculations", Abs(chkOnlyMine.Value)

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Disassociate object variables.
  Set mobjComponent = Nothing

End Sub

Private Sub listCalcCalculation_Click()
  ' Update the component value.
  mobjComponent.Component.CalculationID = listCalcCalculation.ItemData(listCalcCalculation.ListIndex)
  
  'JPD 20031009 Fault 6893
  If (Not gfCurrentUserIsSysSecMgr) And _
    (CalcIsReadOnly(mobjComponent.Component.CalculationID)) Then
    Me.CmdEditCalculation.Caption = "&View Calculation..."
    Me.CmdEditCalculation.Enabled = mbEnableViewCalculation
  Else
    Me.CmdEditCalculation.Caption = "E&dit Calculation..."
    Me.CmdEditCalculation.Enabled = mbEnableEditCalculation
  End If

End Sub


Private Sub listCalcCalculation_DblClick()
  ' Confirm the selection.
  If cmdOK.Enabled Then
    cmdOK_Click
  End If

End Sub






Private Sub listCalcFilters_Click()

  ' Update the component value.
  mobjComponent.Component.FilterID = listCalcFilters.ItemData(listCalcFilters.ListIndex)
  
  'JPD 20030909 Fault 6893
  If (Not gfCurrentUserIsSysSecMgr) And _
    (CalcIsReadOnly(mobjComponent.Component.FilterID)) Then
    Me.cmdEditFilter.Caption = "&View Filter..."
    Me.cmdEditFilter.Enabled = mbEnableViewFilter
  Else
    Me.cmdEditFilter.Caption = "E&dit Filter..."
    Me.cmdEditFilter.Enabled = mbEnableEditFilter
  End If

End Sub

Private Sub listCalcFilters_DblClick()
  ' Confirm the selection.
  If cmdOK.Enabled Then
    cmdOK_Click
  End If

End Sub


Private Sub optComponentType_Click(Index As Integer)
  ' Set the component type property.
  miComponentType = Index
  mobjComponent.ComponentType = Index
  
  DisplayComponentFrame

End Sub


Private Sub optField_Refresh()
  Dim iSelection As FieldSelectionTypes
  
  iSelection = mobjComponent.Component.SelectionType

  Select Case iSelection
    Case giSELECT_RECORDCOUNT
      optField(1).Value = True
    Case giSELECT_RECORDTOTAL
      optField(2).Value = True
    Case Else
      optField(0).Value = True
  End Select
    
  cboFldTable_Refresh
    
End Sub

Private Sub optFieldSel_Refresh()
  Dim iSelection As FieldSelectionTypes
  
  iSelection = mobjComponent.Component.SelectionType

  Select Case iSelection
    Case giSELECT_LASTRECORD
      optFieldSel(1).Value = True
    Case giSELECT_SPECIFICRECORD
      optFieldSel(2).Value = True
    Case Else
      optFieldSel(0).Value = True
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


Private Sub optPValDefaultDateType_Click(Index As Integer)

Dim iCount As Integer

For iCount = 0 To optPValDefaultDateType.Count - 1
  optPValDefaultDateType(iCount).Value = (iCount = Index)
Next iCount

' Enable the default date if we have selected explicit date
asrPValDefaultDate.Enabled = (Index = 0)

' Pass in the Prompted date
mobjComponent.Component.DefaultDateType = Index

End Sub

Private Sub ssTreeFuncFunction_Collapse(Node As SSActiveTreeView.SSNode)
  ' If the specified node is the root node keep it expanded.
  If Node.Key = "FUNCTION_ROOT" Then
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
  ' Update the component object if an function has been selected.
  Dim fFunctionSelected As Boolean
  
  fFunctionSelected = False
  
  If Node.Key <> "FUNCTION_ROOT" Then
    If Node.Parent.Key <> "FUNCTION_ROOT" Then
      fFunctionSelected = True
    End If
  End If
  
  ' Only enable the OK button if a function has been selected.
  ' ie. not when the root node, or one of the category nodes is selected.
  cmdOK.Enabled = fFunctionSelected

End Sub


Private Sub ssTreeOpOperator_Collapse(Node As SSActiveTreeView.SSNode)
  ' If the specified node is the root node keep it expanded.
  If Node.Key = "OPERATOR_ROOT" Then
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
  
  If Node.Key <> "OPERATOR_ROOT" Then
    If Node.Parent.Key <> "OPERATOR_ROOT" Then
      fOperatorSelected = True
    End If
  End If
  
  ' Only enable the OK button if an operator has been selected.
  ' ie. not when the root node, or one of the category nodes is selected.
  cmdOK.Enabled = fOperatorSelected

End Sub

Private Sub txtPValDefaultCharacter_GotFocus()
  UI.txtSelText
  
End Sub


Private Sub txtPValFormat_Change()
  ' Update the component object with the new value.
  mobjComponent.Component.ValueFormat = txtPValFormat.Text

End Sub


Private Sub txtPValFormat_GotFocus()
  UI.txtSelText
  
End Sub


Private Sub txtPValFormat_LostFocus()

  On Error GoTo InvalidMask
  
  ' RH 05/10/00 - BUG 679
  
  'If we are using a mask, test its a valid mask by
  'applying it to a hidden mask control. If an error
  'occurs then we know its not an appropriate mask
  'for the size specified.
  'same kinda thing as SYS MGR, frmColEdit)
  
  If Len(Trim(Me.txtPValFormat.Text)) > 0 Then
    
    ' If the following line causes an error, then the format of the mask is bad
    tdbMaskTest.Format = Me.txtPValFormat.Text
    
    ' If the len of the text property isnt the same as what the user has set
    ' the size property to be, then its an error too !
    If Len(tdbMaskTest.Text) <> Me.asrPValReturnSize.Value Then Err.Raise 380
  
  End If
  
  Exit Sub
  
InvalidMask:
  
  If Err.Number = 380 Then
    MsgBox "You must have at least one user enterable character in the mask field" & vbCrLf & "and the mask must correspond with the 'size' setting.", vbExclamation + vbOKOnly, "Validation Error"
    tdbMaskTest.Format = ""
    txtPValFormat.SetFocus
  Else
    MsgBox "Warning : " & Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, "Validation Error"
    tdbMaskTest.Format = ""
    txtPValFormat.SetFocus
  End If
  
End Sub

Private Sub txtPValPrompt_Change()
  ' Update the component object with the new value.
  mobjComponent.Component.Prompt = txtPValPrompt.Text
    
  ' Only enable the OK button if a prompt is entered and
  ' there is a valid column selected for table type prompted values.
  cmdOK.Enabled = (Len(Trim(txtPValPrompt.Text)) > 0) And _
    (mobjComponent.Component.valueType <> giEXPRVALUE_TABLEVALUE Or _
    mobjComponent.Component.LookupColumn > 0)

End Sub


Private Sub txtPValPrompt_GotFocus()
  UI.txtSelText
  
End Sub


Private Sub txtValCharacterValue_GotFocus()
  UI.txtSelText
  
End Sub


