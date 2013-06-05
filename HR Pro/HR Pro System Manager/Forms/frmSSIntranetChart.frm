VERSION 5.00
Begin VB.Form frmSSIntranetChart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Self-service Intranet Chart Data"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   HelpContextID   =   5086
   Icon            =   "frmSSIntranetChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraChartType 
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   150
      TabIndex        =   35
      Top             =   165
      Width           =   2310
      Begin VB.OptionButton optChartType 
         Caption         =   "Three Tables"
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
         Index           =   2
         Left            =   210
         TabIndex        =   2
         Top             =   1065
         Width           =   1680
      End
      Begin VB.OptionButton optChartType 
         Caption         =   "Two Tables"
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
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optChartType 
         Caption         =   "One Table"
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
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   345
         Width           =   1410
      End
   End
   Begin VB.Frame fra_Y_Data 
      Caption         =   "Data Values (Y-Axis) :"
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
      Left            =   2520
      TabIndex        =   30
      Top             =   1530
      Width           =   8940
      Begin VB.ComboBox cboChartColColumn 
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
         Left            =   5970
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1125
         Width           =   2670
      End
      Begin VB.ComboBox cboSortByAgg 
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
         ItemData        =   "frmSSIntranetChart.frx":000C
         Left            =   5970
         List            =   "frmSSIntranetChart.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   750
         Width           =   2670
      End
      Begin VB.ComboBox cboSortOrderAgg 
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
         ItemData        =   "frmSSIntranetChart.frx":0031
         Left            =   5970
         List            =   "frmSSIntranetChart.frx":003B
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboAggregateType 
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
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboColumnY 
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
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1125
         Width           =   2670
      End
      Begin VB.ComboBox cboTableY 
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
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   750
         Width           =   2670
      End
      Begin VB.Label lblChartIntColour 
         Caption         =   "Chart Colour Column :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4530
         TabIndex        =   37
         Top             =   1080
         Width           =   1170
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSortByAgg 
         AutoSize        =   -1  'True
         Caption         =   "Sort By :"
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
         Left            =   4530
         TabIndex        =   36
         Top             =   810
         Width           =   780
      End
      Begin VB.Label lblSortorderAgg 
         AutoSize        =   -1  'True
         Caption         =   "Sort Order :"
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
         Left            =   4530
         TabIndex        =   34
         Top             =   435
         Width           =   1050
      End
      Begin VB.Label lblAggregateType 
         AutoSize        =   -1  'True
         Caption         =   "Aggregate :"
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
         Left            =   210
         TabIndex        =   33
         Top             =   435
         Width           =   1020
      End
      Begin VB.Label lblColumnY 
         AutoSize        =   -1  'True
         Caption         =   "Column :"
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
         Left            =   210
         TabIndex        =   32
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label lblTableY 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
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
         Left            =   210
         TabIndex        =   31
         Top             =   795
         Width           =   600
      End
   End
   Begin VB.Frame fraChartFilter 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2520
      TabIndex        =   24
      Top             =   4575
      Width           =   8925
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
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
         Left            =   3420
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1425
         TabIndex        =   15
         Top             =   360
         Width           =   1995
      End
      Begin VB.CommandButton cmdFilterClear 
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
         Left            =   3750
         MaskColor       =   &H000000FF&
         TabIndex        =   17
         ToolTipText     =   "Clear Path"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         Caption         =   "Filter :"
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
         Left            =   210
         TabIndex        =   29
         Top             =   405
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   8925
      TabIndex        =   19
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Frame fra_X_Data 
      Caption         =   "Column (X-Axis) :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2520
      TabIndex        =   18
      Top             =   165
      Width           =   8955
      Begin VB.ComboBox cboTableX 
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
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboSortOrderX 
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
         ItemData        =   "frmSSIntranetChart.frx":0056
         Left            =   5970
         List            =   "frmSSIntranetChart.frx":0060
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2670
      End
      Begin VB.ComboBox cboColumnX 
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
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   2670
      End
      Begin VB.Label lblSortorderX 
         AutoSize        =   -1  'True
         Caption         =   "Sort Order :"
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
         Left            =   4530
         TabIndex        =   23
         Top             =   420
         Width           =   1050
      End
      Begin VB.Label lblColumnX 
         AutoSize        =   -1  'True
         Caption         =   "Column :"
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
         Left            =   210
         TabIndex        =   22
         Top             =   810
         Width           =   795
      End
      Begin VB.Label lblTableX 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
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
         Left            =   210
         TabIndex        =   21
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   10215
      TabIndex        =   20
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Frame fra_Z_Data 
      Caption         =   "Row (Z-Axis) :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2520
      TabIndex        =   25
      Top             =   3255
      Width           =   8925
      Begin VB.ComboBox cboTableZ 
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
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboSortOrderZ 
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
         ItemData        =   "frmSSIntranetChart.frx":007B
         Left            =   5970
         List            =   "frmSSIntranetChart.frx":0085
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboColumnZ 
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
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   750
         Width           =   2670
      End
      Begin VB.Label lblTableZ 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
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
         Left            =   210
         TabIndex        =   28
         Top             =   435
         Width           =   600
      End
      Begin VB.Label lblColumnZ 
         AutoSize        =   -1  'True
         Caption         =   "Column :"
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
         Left            =   210
         TabIndex        =   27
         Top             =   795
         Width           =   795
      End
      Begin VB.Label lblSortorderZ 
         AutoSize        =   -1  'True
         Caption         =   "Sort Order :"
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
         Left            =   4530
         TabIndex        =   26
         Top             =   420
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmSSIntranetChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean

' Component definition variables.
Private mfChanged As Boolean
Private mfLoading As Boolean
Private mfRefreshing As Boolean
Private mobjComponent As CExprComponent
Private miComponentType As ExpressionComponentTypes
Private mavColumns() As Variant
Private mDataType As Long
Private mlngTableID As Long
Private mvar_lngPersonnelTableID As Long
Private miChartTableID As Long
Private mlngChartColumnID As Long
Private miChartAggregateType As Integer
Private mlngChartFilterID As Long
Private mavTables() As Variant
Private mlngChart_TableID_2 As Long
Private mlngChart_ColumnID_2 As Long
Private mlngChart_TableID_3 As Long
Private mlngChart_ColumnID_3 As Long
Private mlngChart_SortOrderID As Long
Private miChart_SortDirection As Integer
Private mlngChart_ColourID As Long
Private iLoop As Integer
Private miFilterTableID As Long

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property



Private Sub cboAggregateType_Click()
  
  PopulateSortByCombo
  SetComboItemOrTopItem cboSortByAgg, 0
  
  If Not mfLoading Then
  
  mfChanged = True
    
  RefreshControls
 End If
End Sub

Private Sub cboChartColColumn_Click()
  mfChanged = True
  If Not mfLoading Then
    Chart_ColourID = cboChartColColumn.ItemData(cboChartColColumn.ListIndex)
    RefreshControls
  End If

End Sub

Private Sub ChangeChartType()
  
  If optChartType(0).value Then
    ' Set up One Dimension chart options
    fra_Z_Data.Visible = False
    lblTableY.Visible = False
    cboTableY.Visible = False
    lblColumnY.Visible = False
    cboColumnY.Visible = False
    lblSortByAgg.Visible = True
    cboSortByAgg.Visible = True
  ElseIf optChartType(1).value Then
    ' Set up Two Dimension chart options
    fra_Z_Data.Visible = False
    lblTableY.Visible = True
    cboTableY.Visible = True
    lblColumnY.Visible = True
    cboColumnY.Visible = True
    lblSortByAgg.Visible = True
    cboSortByAgg.Visible = True
  Else
    ' set up Three Dimension chart options
    fra_Z_Data.Visible = True
    lblTableY.Visible = True
    cboTableY.Visible = True
    lblColumnY.Visible = True
    cboColumnY.Visible = True
    lblSortByAgg.Visible = False
    cboSortByAgg.Visible = False
  End If
    

End Sub

Private Sub cboColumnX_Click()

If optChartType(0).value Then
  ' set the aggregate
  PopulateAggregateCombo cboColumnX.ItemData(cboColumnX.ListIndex)
  SetComboItemOrTopItem cboAggregateType, ChartAggregateType
  PopulateColourCombo cboTableX.ItemData(cboTableX.ListIndex)
End If

mfChanged = True

RefreshControls

End Sub

Private Sub cboColumnY_Click()
  ' set the aggregate
  PopulateAggregateCombo cboColumnY.ItemData(cboColumnY.ListIndex)
  SetComboItemOrTopItem cboAggregateType, ChartAggregateType
  PopulateColourCombo cboTableY.ItemData(cboTableY.ListIndex)
  
mfChanged = True

RefreshControls
  
End Sub

Private Sub cboColumnZ_Click()
mfChanged = True

RefreshControls

End Sub

Private Sub cboSortByAgg_Click()
  ' Disable combos as required
  If cboSortByAgg.ListIndex = 1 Then
    ' Sort by aggregate, so disable X-Axis sort order
    lblSortorderX.Enabled = False
    cboSortOrderX.Enabled = False
    lblSortorderAgg.Enabled = True
    cboSortOrderAgg.Enabled = True
  Else
    ' Sort by column, so disable Aggregate sort order
    lblSortorderX.Enabled = True
    cboSortOrderX.Enabled = True
    lblSortorderAgg.Enabled = False
    cboSortOrderAgg.Enabled = False
  End If
  
  If Not mfLoading Then
    SetComboItemOrTopItem cboSortOrderAgg, 0
  End If
  
mfChanged = True

RefreshControls
  
End Sub

Private Sub cboSortOrderAgg_Click()

mfChanged = True

RefreshControls

End Sub

Private Sub cboSortOrderX_Click()
mfChanged = True

RefreshControls

End Sub

Private Sub cboSortOrderZ_Click()
mfChanged = True

RefreshControls

End Sub

Private Sub cboTableX_Click()
  
  On Error GoTo ErrorTrap

  ' XC, XO, YT, then values.

  ' Populate the X column combo
  PopulateColumnXCombo cboTableX.ItemData(cboTableX.ListIndex)

  ' reset the X Sort order combo to default (Ascending) as column has changed.
  cboSortOrderX.ListIndex = 0

  ' Populate the Y table
  PopulateTableYCombo

  ' Set top item if not loading.
  If Not mfLoading Then
    SetComboItemOrTopItem cboColumnX, ChartColumnID
    SetComboItemOrTopItem cboTableY, Chart_TableID_3
    SetComboItemOrTopItem cboColumnY, Chart_ColumnID_3
    SetComboItemOrTopItem cboColumnZ, Chart_ColumnID_2
    SetComboItemOrTopItem cboTableZ, Chart_TableID_2
    
    ' clear the filter box as the tables have changed
    ClearFilter
  End If
  

  mfChanged = True

  RefreshControls
  
ErrorTrap:
 
  
End Sub

Public Function SetSortCombos(plngSortOrderID As Long)
  ' convert the decimal sortorderID into binary then apply as follows:
  ' Digit 1 = X-Axis Data Sort order
  ' Digit 2 = Z-Axis Data Sort order
  ' Digit 3 = 'Sort by Aggregate' tickbox
  ' Digit 4 = 'Sort by Aggregate' (Y-Axis) sort order

  Dim pstrBinaryString As String
  pstrBinaryString = DecToBin(plngSortOrderID, 4)
  
  ' Horizontal (X-AXIS) Sort Combo
  If Mid(pstrBinaryString, 1, 1) = "0" Then ' Ascending
    cboSortOrderX.ListIndex = IIf(cboTableX <> vbNullString, 0, -1)
  Else ' Descending
    cboSortOrderX.ListIndex = IIf(cboTableX <> vbNullString, 1, -1)
  End If
      
  ' Vertical (Z-AXIS) Sort Combo
  If Mid(pstrBinaryString, 2, 1) = "0" Then ' Ascending
    cboSortOrderZ.ListIndex = IIf(cboTableZ <> vbNullString, 0, -1)
  Else ' Descending
    cboSortOrderZ.ListIndex = IIf(cboTableZ <> vbNullString, 1, -1)
  End If
      
  ' Aggregate (Y-AXIS) sort tick box
  If val(Mid(pstrBinaryString, 3, 1)) = 0 Then
    cboSortByAgg.ListIndex = 0
  Else
    cboSortByAgg.ListIndex = 1
  End If
  
  ' Aggregate (Y-AXIS) Sort Combo
  If Mid(pstrBinaryString, 4, 1) = "0" Then ' Ascending
    cboSortOrderAgg.ListIndex = 0
  Else ' Descending
    cboSortOrderAgg.ListIndex = 1
  End If
  
  cboSortByAgg_Click
  
End Function

Private Function DecToBin(DeciValue As Long, Optional NoOfBits As Integer = 8) As String

'********************************************************************************
'* Name : DecToBin
'* Date : 2003
'* Author : Alex Etchells
'*********************************************************************************
Dim i As Integer     'make sure there are enough bits to contain the number
Do While DeciValue > (2 ^ NoOfBits) - 1
  NoOfBits = NoOfBits + 8
Loop
DecToBin = vbNullString
'build the string
For i = 0 To (NoOfBits - 1)
  DecToBin = CStr((DeciValue And 2 ^ i) / 2 ^ i) & DecToBin
Next i
End Function

Private Function ConvertCombosToDecimal() As Integer
' Clever old me. I decided to convert the three combo values to a binary, then decimal value for storing.
' e.g.  5 in decimal = 1-0-1 in binary, and can be used to set the 3 combos to Descending, Ascending, Descending (1=Descending)
Dim pstrBinaryString As String
Dim BinaryToDec As Integer

  pstrBinaryString = IIf(cboSortOrderX = "Descending", "1", "0")
  pstrBinaryString = pstrBinaryString & IIf(cboSortOrderZ = "Descending", "1", "0")
  pstrBinaryString = pstrBinaryString & IIf(cboSortByAgg.ListIndex > 0, "1", "0")
  pstrBinaryString = pstrBinaryString & IIf(cboSortOrderAgg = "Descending", "1", "0")
  
  Do
    BinaryToDec = BinaryToDec + (Left(pstrBinaryString, 1) * 2 ^ (Len(pstrBinaryString) - 1))
    pstrBinaryString = Mid(pstrBinaryString, 2)
  Loop Until pstrBinaryString = ""
  
  ConvertCombosToDecimal = BinaryToDec
End Function

Private Sub cboSortOrder_Click(Index As Integer)
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboTableY_Click()
  On Error GoTo ErrorTrap

  ' Populate the Y column combo
  PopulateColumnYCombo cboTableY.ItemData(cboTableY.ListIndex)

  ' Populate the Z table
  PopulateTableZCombo
  
  ' Set top item if not loading.
  If Not mfLoading Then
    SetComboItemOrTopItem cboTableY, Chart_TableID_3
    SetComboItemOrTopItem cboColumnY, Chart_ColumnID_3
    SetComboItemOrTopItem cboTableZ, Chart_TableID_2
    SetComboItemOrTopItem cboColumnZ, Chart_ColumnID_2
    
  End If
  
  mfChanged = True

  RefreshControls
  
ErrorTrap:
End Sub

Private Sub cboTableZ_Click()
  On Error GoTo ErrorTrap

  ' XC, XO, YT, then values.

  ' Populate the Z column combo
  PopulateColumnZCombo cboTableZ.ItemData(cboTableZ.ListIndex)

  ' Set top item if not loading.
  If Not mfLoading Then
    SetComboItemOrTopItem cboColumnZ, Chart_ColumnID_2
  End If

  ' If filter tag doesn't match any tables, clear it
  If Not mfLoading And miFilterTableID <> cboTableX.ItemData(cboTableX.ListIndex) And miFilterTableID <> cboTableY.ItemData(cboTableY.ListIndex) And miFilterTableID <> cboTableZ.ItemData(cboTableZ.ListIndex) Then
    ClearFilter
  End If


  mfChanged = True

  RefreshControls
  
ErrorTrap:
 
End Sub

Private Sub optChartType_Click(Index As Integer)

ChangeChartType
mfChanged = True

RefreshControls

End Sub

'Private Sub chkSortAggregate_Click()
'  cboSortOrderAgg.Enabled = chkSortAggregate
'  Chart_SortOrderID = ConvertCombosToDecimal
'  cboSortOrderAgg.Enabled = (chkSortAggregate.value = 1)
'  lblSortorderAgg.Enabled = (chkSortAggregate.value = 1)
'  mfChanged = True
'  RefreshControls
'End Sub

Private Sub cmdFilterClear_Click()
  ClearFilter
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Cancelled = True
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOk_Click
        Cancel = True
    End Select
  End If
  RefreshControls
End Sub

Private Sub RefreshControls()
  
  
  cmdOK.Enabled = mfChanged
  
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  UnLoad Me
End Sub

Private Sub cmdFilter_Click()

  ' Display the 'Where Clause' expression selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True
  
  ' use the first combo's tableid if the others are empty.
  ' mlngTableID = IIf(mlngTableID = 0, cboTableX.ItemData(cboTableX.ListIndex), mlngTableID)
  
  ' calculate the base table to use as the tableid
  
  If ChartTableID = 0 Then Exit Sub
  
  If Chart_TableID_2 = 0 And Chart_TableID_3 = 0 Then
    ' Single axis chart, so use tableID 1 (X-axis)
    mlngTableID = ChartTableID
  Else
    ' IsChildOfTable(Parent, Child)
    If IsChildOfTable(Chart_TableID_3, Chart_TableID_2) Then
      If IsChildOfTable(Chart_TableID_2, ChartTableID) Then
        mlngTableID = ChartTableID
      Else
        mlngTableID = Chart_TableID_2
      End If
    Else
      If IsChildOfTable(Chart_TableID_3, ChartTableID) Then
        mlngTableID = ChartTableID
      Else
        mlngTableID = Chart_TableID_3
      End If
    End If
  End If
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise mlngTableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      cmdFilterClear.Enabled = True
      mlngChartFilterID = .ExpressionID
      mfChanged = True
      miFilterTableID = mlngTableID
      
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
        cmdFilterClear.Enabled = False
        miFilterTableID = 0
      End If
    End If

  End With

  RefreshControls

TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

'Private Function GetFldSelFilterDetails() As Boolean
'  ' Get the 'Field Selection Filter' expression details.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sExprName As String
'  Dim objExpr As CExpression
'
'  fOK = True
'
'  ' Initialise the default values.
'  sExprName = ""
'
'  ' Instantiate the expression class.
'  Set objExpr = New CExpression
'
'  With objExpr
'    ' Set the expression id.
'    .ExpressionID = mobjComponent.Component.SelectionFilter
'
'    ' Read the required info from the expression.
'    If .ReadExpressionDetails Then
'      sExprName = .Name
'    End If
'  End With
'
'TidyUpAndExit:
'  ' Disassociate object variables.
'  Set objExpr = Nothing
'  If Not fOK Then
'    sExprName = ""
'  End If
'
'  txtfldSelFilter.Text = sExprName
'
'  GetFldSelFilterDetails = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function



Public Sub Initialize(plngChartViewID As Long, _
                        miChartTableID As Long, _
                        plngChartColumnID As Long, _
                        plngChartFilterID As Long, _
                        piChartAggregateType As Integer, _
                        plngChart_TableID_2 As Long, _
                        plngChart_ColumnID_2 As Long, _
                        plngChart_TableID_3 As Long, _
                        plngChart_ColumnID_3 As Long, _
                        plngChart_SortOrderID As Long, _
                        plngChart_SortDirection As Long, _
                        plngChart_ColourID As Long)

'  '  ChartViewID = plngChartViewID
'  ChartTableID = miChartTableID
'  ChartColumnID = plngChartColumnID
'  ChartAggregateType = piChartAggregateType
'  ChartFilterID = plngChartFilterID
'  Chart_TableID_2 = plngChart_TableID_2
'  Chart_ColumnID_2 = plngChart_ColumnID_2
'  Chart_TableID_3 = plngChart_TableID_3
'  Chart_ColumnID_3 = plngChart_ColumnID_3
'  Chart_SortOrderID = plngChart_SortOrderID
'  Chart_SortDirection = plngChart_SortDirection
'  Chart_ColourID = plngChart_ColourID
    
  mfLoading = True
    
  ' Populate the X-Axis table and column combos
  PopulateTableXCombo  ' populate X-Axis
  
  ' Now all combos are populated, set to preset values.
  SetComboItemOrTopItem cboTableX, miChartTableID   ' this also populates table Y
  SetComboItemOrTopItem cboColumnX, plngChartColumnID
  
  SetComboItemOrTopItem cboTableY, plngChart_TableID_3  ' this also populates table Z
  SetComboItemOrTopItem cboColumnY, plngChart_ColumnID_3
  
  SetComboItemOrTopItem cboTableZ, plngChart_TableID_2
  SetComboItemOrTopItem cboColumnZ, plngChart_ColumnID_2
        
  optChartType(0).value = (plngChart_ColumnID_2 = 0 And plngChart_ColumnID_3 = 0)
  optChartType(1).value = (plngChart_ColumnID_3 > 0)
  optChartType(2).value = (plngChart_ColumnID_2 > 0)
  
  ChangeChartType ' display/hide relevant frames and combos
  
  If ChartTableID = 0 Then
    PopulateColourCombo cboTableX.ItemData(cboTableX.ListIndex)
  Else
    PopulateColourCombo IIf(Chart_ColumnID_3 > 0, Chart_TableID_3, ChartTableID)
  End If
      
  SetComboItemOrTopItem cboChartColColumn, plngChart_ColourID
      
  If optChartType(0).value Then   ' one dimension chart
    PopulateAggregateCombo cboColumnX.ItemData(cboColumnX.ListIndex)
  Else
    PopulateAggregateCombo cboColumnY.ItemData(cboColumnY.ListIndex)
  End If
  
  ' PopulateAggregateCombo cboColumnY.ItemData(cboColumnY.ListIndex)
  SetAggregateValue piChartAggregateType
  
  PopulateSortByCombo
  SetSortCombos plngChart_SortOrderID
  
  ' Filter frame
  txtFilter.Tag = plngChartFilterID
  txtFilter.Text = GetExpressionName(txtFilter.Tag)
  If txtFilter.Text = "" Then
    cmdFilterClear.Enabled = False
  Else
    cmdFilterClear.Enabled = True
  End If
  txtFilter.Enabled = False
  txtFilter.BackColor = vbButtonFace
  

  
  
  ' EnableDisableCombos
  
  mfLoading = False
  
  mfChanged = False
  
  RefreshControls
  
End Sub

Private Sub cmdOk_Click()

'  If cboChartType = "Multi Axis Chart" And (cboTableX = vbNullString Or cboColumnZ = vbNullString Or cboColumnY = vbNullString) Then
'    MsgBox "Please populate all required columns before trying to save.", vbExclamation + vbOKOnly, "Chart Data"
'    Exit Sub
'  End If


  Cancelled = False
  Me.Hide
End Sub

Private Function PopulateColourCombo(plngTableID As Long) As Boolean
  Dim i As Integer
  
  ' Clear the contents of the combo
  cboChartColColumn.Clear

  If plngTableID <= 0 Then Exit Function

  ' Add empty value
  cboChartColColumn.AddItem ""

  ' Add the table's columns to the view definition in the local database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  recColEdit.Index = "idxTableID"
  recColEdit.Seek "=", plngTableID

  fOK = Not recColEdit.NoMatch

  If fOK Then

    Do While Not recColEdit.EOF

      ' If no more columns for this table exit loop
      If recColEdit!TableID <> plngTableID Then
        Exit Do
      End If

      ' Don't add deleted or system columns
      If recColEdit!Deleted <> True And recColEdit!columntype <> giCOLUMNTYPE_SYSTEM Then
        ' Add the column to the combo
        ' Making sure it isn't ole, photo, wp or link...
        If recColEdit!DataType <> dtLONGVARCHAR And _
          recColEdit!DataType <> dtBINARY And _
          recColEdit!DataType <> dtVARBINARY And _
          recColEdit!DataType <> dtLONGVARBINARY And _
          recColEdit!ControlType = 2 ^ 15 Then
            cboChartColColumn.AddItem recColEdit.Fields("ColumnName")
            cboChartColColumn.ItemData(cboChartColColumn.NewIndex) = recColEdit.Fields("ColumnID")
        End If
      End If

      recColEdit.MoveNext
    Loop
      
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function PopulateSortByCombo()
  
  If optChartType(0).value Then
    cboSortByAgg.Clear
    cboSortByAgg.AddItem cboColumnX.Text
    cboSortByAgg.AddItem cboAggregateType.Text & " of " & cboColumnX.Text
  ElseIf optChartType(1).value Then
    cboSortByAgg.Clear
    cboSortByAgg.AddItem cboColumnY.Text
    cboSortByAgg.AddItem cboAggregateType.Text & " of " & cboColumnY.Text
  End If
  
  
  
End Function


Private Function PopulateTableXCombo()
  
  Dim i As Integer
  ' Clear the contents of the combo.
  cboTableX.Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If !TableType <> iTabLookup And Not !Deleted Then
        cboTableX.AddItem !TableName
        cboTableX.ItemData(cboTableX.NewIndex) = !TableID
      End If

      .MoveNext
    Loop
  End With

End Function

Private Sub ClearFilter()
  txtFilter.Text = vbNullString
  txtFilter.Tag = 0
  mlngChartFilterID = 0
  miFilterTableID = 0
  cmdFilterClear.Enabled = False
  mfChanged = True
  RefreshControls
End Sub


Private Function PopulateColumnXCombo(plngTableID As Long)
 
  On Error GoTo ErrorTrap
 
  ' Clear the contents of the combo
  cboColumnX.Clear

  Dim fOK As Boolean

  recColEdit.Index = "idxTableID"
  recColEdit.Seek "=", plngTableID

  fOK = Not recColEdit.NoMatch

  If fOK Then

    Do While Not recColEdit.EOF

      ' If no more columns for this table exit loop
      If recColEdit!TableID <> plngTableID Then
        Exit Do
      End If

      ' Don't add deleted or system columns
      If recColEdit!Deleted <> True And recColEdit!columntype <> giCOLUMNTYPE_SYSTEM Then
        ' Add the column to the combo
        ' Making sure it isn't ole, photo, wp or link...
        If recColEdit!DataType <> dtLONGVARCHAR And _
          recColEdit!DataType <> dtBINARY And _
          recColEdit!DataType <> dtVARBINARY And _
          recColEdit!DataType <> dtLONGVARBINARY Then
            cboColumnX.AddItem recColEdit.Fields("ColumnName")
            cboColumnX.ItemData(cboColumnX.NewIndex) = recColEdit.Fields("ColumnID")
        End If
      End If

      recColEdit.MoveNext
    Loop
      
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function PopulateColumnYCombo(plngTableID As Long)
  
  Dim fOK As Boolean
  On Error GoTo ErrorTrap
  
  ' Populate Y-Axis combo with any parent or child of X-Axis combo table, and the X-Axis table itself.
 
  ' Clear the contents of the combo
  cboColumnY.Clear

  'plngTableX_ID = cboTableX.ItemData(cboTableX.ListIndex)
  
  If plngTableID <= 0 Then Exit Function

  recColEdit.Index = "idxTableID"
  recColEdit.Seek "=", plngTableID

  fOK = Not recColEdit.NoMatch

  If fOK Then

    Do While Not recColEdit.EOF

      ' If no more columns for this table exit loop
      If recColEdit!TableID <> plngTableID Then
        Exit Do
      End If

      ' Don't add deleted or system columns
      If recColEdit!Deleted <> True And recColEdit!columntype <> giCOLUMNTYPE_SYSTEM Then
        ' Add the column to the combo
        ' Making sure it isn't ole, photo, wp or link...
        If recColEdit!DataType <> dtLONGVARCHAR And _
          recColEdit!DataType <> dtBINARY And _
          recColEdit!DataType <> dtVARBINARY And _
          recColEdit!DataType <> dtLONGVARBINARY Then
            cboColumnY.AddItem recColEdit.Fields("ColumnName")
            cboColumnY.ItemData(cboColumnY.NewIndex) = recColEdit.Fields("ColumnID")
        End If
      End If

      recColEdit.MoveNext
    Loop
      
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function PopulateColumnZCombo(plngTableID As Long)
 
  On Error GoTo ErrorTrap
 
  ' Clear the contents of the combo
  cboColumnZ.Clear

  If plngTableID <= 0 Then Exit Function

  Dim fOK As Boolean

  recColEdit.Index = "idxTableID"
  recColEdit.Seek "=", plngTableID

  fOK = Not recColEdit.NoMatch

  If fOK Then

    Do While Not recColEdit.EOF

      ' If no more columns for this table exit loop
      If recColEdit!TableID <> plngTableID Then
        Exit Do
      End If

      ' Don't add deleted or system columns
      If recColEdit!Deleted <> True And recColEdit!columntype <> giCOLUMNTYPE_SYSTEM Then
        ' Add the column to the combo
        ' Making sure it isn't ole, photo, wp or link...
        If recColEdit!DataType <> dtLONGVARCHAR And _
          recColEdit!DataType <> dtBINARY And _
          recColEdit!DataType <> dtVARBINARY And _
          recColEdit!DataType <> dtLONGVARBINARY Then
            cboColumnZ.AddItem recColEdit.Fields("ColumnName")
            cboColumnZ.ItemData(cboColumnZ.NewIndex) = recColEdit.Fields("ColumnID")
        End If
      End If

      recColEdit.MoveNext
    Loop
      
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function




Private Function PopulateTableZCombo()
  
  On Error GoTo Error_Trap

  ' Clear Table Combo
  cboTableZ.Clear

  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
          
  If cboTableX.ItemData(cboTableX.ListIndex) = cboTableY.ItemData(cboTableY.ListIndex) Then
    ' Table X = Table Y, so populate table Z with all parent/child of Table X
    ' Add the Horizontal table to the combo
    cboTableZ.AddItem cboTableX
    cboTableZ.ItemData(cboTableZ.NewIndex) = cboTableX.ItemData(cboTableX.ListIndex)
    
    GetRelatedTables cboTableX.ItemData(cboTableX.ListIndex), "PARENT"
    GetRelatedTables cboTableX.ItemData(cboTableX.ListIndex), "CHILD"
  
    ' add related tables to the combo
    For iLoop = 1 To UBound(mavTables, 2)
        cboTableZ.AddItem mavTables(2, iLoop)
        cboTableZ.ItemData(cboTableZ.NewIndex) = mavTables(1, iLoop)
    Next iLoop
  Else
    ' Table X <> Table Y, so populate Table Z with X & Y
    cboTableZ.AddItem cboTableX.Text
    cboTableZ.ItemData(cboTableZ.NewIndex) = cboTableX.ItemData(cboTableX.ListIndex)
    cboTableZ.AddItem cboTableY.Text
    cboTableZ.ItemData(cboTableZ.NewIndex) = cboTableY.ItemData(cboTableY.ListIndex)
  End If
      
TidyUpAndExit:
  Exit Function

Error_Trap:
  MsgBox "Error populating table Z dropdown box.", vbExclamation + vbOKOnly, "Chart Link"
  PopulateTableZCombo = False
  GoTo TidyUpAndExit
  
End Function

Private Function PopulateTableYCombo()
  Dim plngTableXID As Long
  
  On Error GoTo Error_Trap
  
  ' Clear Table Combo
  cboTableY.Clear
  
  ' Get the tables related to the 'X' table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  ' Add defaults first
  ' Add parent 1 as intersection may be same table.
  plngTableXID = cboTableX.ItemData(cboTableX.ListIndex)
  cboTableY.AddItem cboTableX.Text
  cboTableY.ItemData(cboTableY.NewIndex) = plngTableXID

  GetRelatedTables plngTableXID, "PARENT"
  GetRelatedTables plngTableXID, "CHILD"
  
  For iLoop = 1 To UBound(mavTables, 2)
      cboTableY.AddItem mavTables(2, iLoop)
      cboTableY.ItemData(cboTableY.NewIndex) = mavTables(1, iLoop)
  Next iLoop
  
TidyUpAndExit:
  Exit Function

Error_Trap:
  MsgBox "Error populating Y-Axis tables dropdown box.", vbExclamation + vbOKOnly, "Chart Link"
  PopulateTableYCombo = False
  GoTo TidyUpAndExit
  
End Function

Private Sub GetRelatedTables(plngTableID As Long, psRelationship As String)
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim fFound As Boolean
  
  If psRelationship = "CHILD" Then
    sSQL = "SELECT tmpTables.tableName, tmpTables.tableID" & _
      " FROM tmpTables " & _
      " INNER JOIN tmpRelations ON tmpTables.tableID = tmpRelations.childID" & _
      " WHERE tmpRelations.parentID = " & Trim(Str(plngTableID))
  Else
    sSQL = "SELECT tmpTables.tableName, tmpTables.tableID" & _
      " FROM tmpTables " & _
      " INNER JOIN tmpRelations ON tmpTables.tableID = tmpRelations.parentID" & _
      " WHERE tmpRelations.childID = " & Trim(Str(plngTableID))
  End If
  
  Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly, dbReadOnly)
  
  
  ' Set rsTables = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  Do While Not rsTables.EOF
    fFound = False
    
    For iLoop = 1 To UBound(mavTables, 2)
      If mavTables(1, iLoop) = rsTables!TableID Then
        fFound = True
        Exit For
      End If
    Next iLoop
    
    If fFound = False Then
      ReDim Preserve mavTables(3, UBound(mavTables, 2) + 1)
      mavTables(1, UBound(mavTables, 2)) = rsTables!TableID
      mavTables(2, UBound(mavTables, 2)) = rsTables!TableName
      mavTables(3, UBound(mavTables, 2)) = (psRelationship = "PARENT")
      
      ' NPG - unrem the following to include grandchild tables....
      ' GetRelatedTables rsTables!TableID, psRelationship
    End If
    
    rsTables.MoveNext
  Loop
  rsTables.Close
  Set rsTables = Nothing

End Sub



Private Function PopulateAggregateCombo(plngColumnID As Long) As Boolean
  
  Dim piColumnDataType As Integer
  
  cboAggregateType.Clear
  cboAggregateType.AddItem "Count"
  cboAggregateType.ItemData(cboAggregateType.NewIndex) = 0
      
  piColumnDataType = GetColumnDataType(plngColumnID)
  
  If piColumnDataType = dtINTEGER Or piColumnDataType = dtNUMERIC Then
    cboAggregateType.AddItem "Total"
    cboAggregateType.ItemData(cboAggregateType.NewIndex) = 1
    cboAggregateType.AddItem "Average"
    cboAggregateType.ItemData(cboAggregateType.NewIndex) = 2
    cboAggregateType.AddItem "Minimum"
    cboAggregateType.ItemData(cboAggregateType.NewIndex) = 3
    cboAggregateType.AddItem "Maximum"
    cboAggregateType.ItemData(cboAggregateType.NewIndex) = 4
  End If
  
End Function


Private Function SetAggregateValue(piAggregateType As Integer)
  ' Set the correct item as default
  For iLoop = 0 To cboAggregateType.ListCount - 1
    If cboAggregateType.ItemData(iLoop) = piAggregateType Then
      cboAggregateType.ListIndex = iLoop
      Exit For
    End If
  Next

  If cboAggregateType.ListIndex < 0 Then cboAggregateType.ListIndex = 0
  
End Function

Public Sub SetTable_X_Value(plngDefaultID As Long)
    
  For iLoop = 0 To cboTableX.ListCount - 1
    If cboTableX.ItemData(iLoop) = plngDefaultID Then
      cboTableX.ListIndex = iLoop
      Exit For
    End If
  Next

  If cboTableX.ListCount > 0 And cboTableX.ListIndex = -1 Then cboTableX.ListIndex = 0
  
  ChartTableID = cboTableX.ItemData(cboTableX.ListIndex)
  
  PopulateTableYCombo

End Sub

Public Sub SetComboItemOrTopItem(cboCombo As ComboBox, lItem As Long)

  Dim lCount As Long
    
  With cboCombo
    For lCount = 1 To .ListCount
      If .ItemData(lCount - 1) = lItem Then
        .ListIndex = lCount - 1
        Exit For
      End If
    Next
    
    If .ListCount > 0 And .ListIndex < 0 Then
      .ListIndex = 0
    End If
    
  End With

End Sub

Public Property Get ChartTableID() As Long
  If cboTableX.ListIndex >= 0 Then
    ChartTableID = cboTableX.ItemData(cboTableX.ListIndex) ' miChartTableID
  Else
    ChartTableID = 0
  End If
End Property

Public Property Let ChartTableID(ByVal plngNewValue As Long)
  miChartTableID = plngNewValue
End Property

Public Property Get ChartColumnID() As Long
  If cboColumnX.ListIndex >= 0 Then
    ChartColumnID = cboColumnX.ItemData(cboColumnX.ListIndex)
  Else
    ChartColumnID = 0
  End If
End Property

Public Property Let ChartColumnID(ByVal plngNewValue As Long)
  mlngChartColumnID = plngNewValue
End Property

Public Property Get ChartAggregateType() As Integer
  If cboAggregateType.ListIndex >= 0 Then
    ChartAggregateType = cboAggregateType.ItemData(cboAggregateType.ListIndex)
  Else
    ChartAggregateType = 0
  End If
End Property

Public Property Let ChartAggregateType(ByVal piNewValue As Integer)
  miChartAggregateType = piNewValue
End Property

Public Property Get ChartFilterID() As Long
  ChartFilterID = txtFilter.Tag
End Property

Public Property Let ChartFilterID(ByVal plngNewValue As Long)
  mlngChartFilterID = plngNewValue
End Property

Public Property Get Chart_TableID_2() As Long
  If cboTableZ.ListIndex >= 0 And optChartType(2).value Then
    Chart_TableID_2 = cboTableZ.ItemData(cboTableZ.ListIndex)
  Else
    Chart_TableID_2 = 0
  End If
End Property

Public Property Let Chart_TableID_2(ByVal plngNewValue As Long)
  mlngChart_TableID_2 = plngNewValue
End Property

Public Property Get Chart_ColumnID_2() As Long
  If cboColumnZ.ListIndex >= 0 And optChartType(2).value Then
    Chart_ColumnID_2 = cboColumnZ.ItemData(cboColumnZ.ListIndex)
  Else
    Chart_ColumnID_2 = 0
  End If
End Property

Public Property Let Chart_ColumnID_2(ByVal plngNewValue As Long)
  mlngChart_ColumnID_2 = plngNewValue
End Property

Public Property Get Chart_TableID_3() As Long
  If cboTableY.ListIndex >= 0 And Not optChartType(0).value Then
    Chart_TableID_3 = cboTableY.ItemData(cboTableY.ListIndex)
  Else
    Chart_TableID_3 = 0
  End If
End Property

Public Property Let Chart_TableID_3(ByVal plngNewValue As Long)
  mlngChart_TableID_3 = plngNewValue
End Property

Public Property Get Chart_ColumnID_3() As Long
  If cboColumnY.ListIndex >= 0 And Not optChartType(0).value Then
    Chart_ColumnID_3 = cboColumnY.ItemData(cboColumnY.ListIndex)
  Else
    Chart_ColumnID_3 = 0
  End If
End Property

Public Property Let Chart_ColumnID_3(ByVal plngNewValue As Long)
  mlngChart_ColumnID_3 = plngNewValue
End Property

Public Property Get Chart_SortOrderID() As Long
  Chart_SortOrderID = ConvertCombosToDecimal
End Property

Public Property Let Chart_SortOrderID(ByVal plngNewValue As Long)
  mlngChart_SortOrderID = plngNewValue
End Property

Public Property Get Chart_SortDirection() As Integer
  Chart_SortDirection = miChart_SortDirection
End Property

Public Property Let Chart_SortDirection(ByVal piNewValue As Integer)
  miChart_SortDirection = piNewValue
End Property

Public Property Get Chart_ColourID() As Long
  If cboChartColColumn.ListIndex >= 0 Then
    Chart_ColourID = cboChartColColumn.ItemData(cboChartColColumn.ListIndex)
  Else
    Chart_ColourID = 0
  End If
End Property

Public Property Let Chart_ColourID(ByVal plngNewValue As Long)
  mlngChart_ColourID = plngNewValue
End Property


