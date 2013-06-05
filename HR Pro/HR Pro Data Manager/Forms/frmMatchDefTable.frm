VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "CODEJO~1.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmMatchDefTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Match Report Table Comparison"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMatchDefTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNoDrop 
      Height          =   495
      Left            =   930
      Picture         =   "frmMatchDefTable.frx":000C
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4155
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picDocument 
      Height          =   510
      Left            =   1515
      Picture         =   "frmMatchDefTable.frx":08D6
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6570
      TabIndex        =   28
      Top             =   4155
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   5280
      TabIndex        =   27
      Top             =   4155
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   6959
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ta&bles"
      TabPicture(0)   =   "frmMatchDefTable.frx":11A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTables"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraExpressions"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Breakdown Colu&mns"
      TabPicture(1)   =   "frmMatchDefTable.frx":11BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAdd"
      Tab(1).Control(1)=   "cmdRemove"
      Tab(1).Control(2)=   "cmdMoveUp"
      Tab(1).Control(3)=   "cmdMoveDown"
      Tab(1).Control(4)=   "cmdAddAll"
      Tab(1).Control(5)=   "cmdRemoveAll"
      Tab(1).Control(6)=   "fraFieldsAvailable"
      Tab(1).Control(7)=   "fraFieldsSelected"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraExpressions 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   1720
         Width           =   7380
         Begin VB.CommandButton cmdPreferredClear 
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
            Left            =   5280
            MaskColor       =   &H000000FF&
            TabIndex        =   12
            ToolTipText     =   "Clear Path"
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdScoreClear 
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
            Left            =   5280
            MaskColor       =   &H000000FF&
            TabIndex        =   16
            ToolTipText     =   "Clear Path"
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdRequiredClear 
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
            Left            =   5280
            MaskColor       =   &H000000FF&
            TabIndex        =   8
            ToolTipText     =   "Clear Path"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtPreferred 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2205
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   700
            Width           =   2715
         End
         Begin VB.CommandButton cmdPreferred 
            Caption         =   "..."
            Height          =   315
            Left            =   4915
            TabIndex        =   11
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.CommandButton cmdScore 
            Caption         =   "..."
            Height          =   315
            Left            =   4915
            TabIndex        =   15
            Top             =   1100
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtScore 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2205
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1100
            Width           =   2715
         End
         Begin VB.CommandButton cmdRequired 
            Caption         =   "..."
            Height          =   315
            Left            =   4915
            TabIndex        =   7
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   330
         End
         Begin VB.TextBox txtRequired 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2205
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   300
            Width           =   2715
         End
         Begin VB.Label lblPreferred 
            AutoSize        =   -1  'True
            Caption         =   "Preferred Matches :"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   760
            Width           =   1440
         End
         Begin VB.Label lblScore 
            AutoSize        =   -1  'True
            Caption         =   "Score Calculation :"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1155
            Width           =   1335
         End
         Begin VB.Label lblRequired 
            AutoSize        =   -1  'True
            Caption         =   "Required Matches :"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame fraTables 
         Height          =   1305
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   7380
         Begin VB.ComboBox cboMatchTables 
            Height          =   315
            Left            =   2200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   700
            Width           =   3000
         End
         Begin VB.ComboBox cboTables 
            Height          =   315
            Left            =   2200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   300
            Width           =   3000
         End
         Begin VB.Label lblMatchTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Match Table :"
            Height          =   195
            Left            =   225
            TabIndex        =   2
            Top             =   760
            Width           =   975
         End
         Begin VB.Label lblTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table :"
            Height          =   195
            Left            =   225
            TabIndex        =   0
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraFieldsSelected 
         Caption         =   "Columns Selected :"
         Height          =   3435
         Left            =   -70360
         TabIndex        =   19
         Top             =   360
         Width           =   2860
         Begin VB.TextBox txtHeading 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1125
            TabIndex        =   22
            Top             =   2200
            Width           =   1575
         End
         Begin COASpinner.COA_Spinner spnSize 
            Height          =   315
            Left            =   1125
            TabIndex        =   24
            Top             =   2595
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            BackColor       =   -2147483633
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
            MaximumValue    =   2147483647
            Text            =   "0"
         End
         Begin ComctlLib.ListView ListView2 
            Height          =   1800
            Left            =   195
            TabIndex        =   20
            Top             =   300
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   3175
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Column"
               Object.Tag             =   "Column"
               Text            =   "Column"
               Object.Width           =   5644
            EndProperty
         End
         Begin COASpinner.COA_Spinner spnDec 
            Height          =   315
            Left            =   1125
            TabIndex        =   26
            Top             =   3000
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            BackColor       =   -2147483633
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
            MaximumValue    =   9999
            Text            =   "0"
         End
         Begin VB.Label lblProp_ColumnHeading 
            BackStyle       =   0  'Transparent
            Caption         =   "Heading :"
            Height          =   195
            Left            =   195
            TabIndex        =   21
            Top             =   2260
            Width           =   1260
         End
         Begin VB.Label lblProp_Size 
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            Height          =   195
            Left            =   195
            TabIndex        =   23
            Top             =   2655
            Width           =   570
         End
         Begin VB.Label lblProp_Decimals 
            BackStyle       =   0  'Transparent
            Caption         =   "Decimals :"
            Height          =   195
            Left            =   195
            TabIndex        =   25
            Top             =   3060
            Width           =   915
         End
      End
      Begin VB.Frame fraFieldsAvailable 
         Caption         =   "Columns Available :"
         Height          =   3435
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   2895
         Begin ComctlLib.ListView ListView1 
            Height          =   3015
            Left            =   195
            TabIndex        =   18
            Top             =   270
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   5318
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   5644
            EndProperty
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns :"
         Height          =   5200
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   9180
         Begin VB.CommandButton cmdClearColumn 
            Caption         =   "Remo&ve All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   34
            Top             =   2100
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteColumn 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   33
            Top             =   1500
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditColumn 
            Caption         =   "&Edit..."
            Enabled         =   0   'False
            Height          =   400
            Left            =   7800
            TabIndex        =   32
            Top             =   900
            Width           =   1200
         End
         Begin SSDataWidgets_B.SSDBGrid grdColumns 
            Height          =   4725
            Left            =   195
            TabIndex        =   35
            Top             =   300
            Width           =   7425
            ScrollBars      =   0
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            Col.Count       =   3
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
            RowNavigation   =   1
            MaxSelectedRows =   1
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            Columns.Count   =   3
            Columns(0).Width=   4313
            Columns(0).Caption=   "Table"
            Columns(0).Name =   "Table"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   4313
            Columns(1).Caption=   "Match Table"
            Columns(1).Name =   "Match Table"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   4419
            Columns(2).Caption=   "Criteria"
            Columns(2).Name =   "Criteria"
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   13097
            _ExtentY        =   8334
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
      Begin XtremeSuiteControls.PushButton cmdRemoveAll 
         Height          =   405
         Left            =   -71850
         TabIndex        =   43
         Top             =   2025
         Width           =   1305
         _Version        =   851969
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Remo&ve All"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdAddAll 
         Height          =   405
         Left            =   -71850
         TabIndex        =   42
         Top             =   945
         Width           =   1305
         _Version        =   851969
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Add A&ll"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdMoveDown 
         Height          =   405
         Left            =   -71850
         TabIndex        =   41
         Top             =   3225
         Width           =   1305
         _Version        =   851969
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Do&wn"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdMoveUp 
         Height          =   405
         Left            =   -71850
         TabIndex        =   40
         Top             =   2730
         Width           =   1305
         _Version        =   851969
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "&Up"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdRemove 
         Height          =   405
         Left            =   -71850
         TabIndex        =   39
         Top             =   1515
         Width           =   1305
         _Version        =   851969
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "&Remove"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdAdd 
         Height          =   405
         Left            =   -71850
         TabIndex        =   38
         Top             =   450
         Width           =   1305
         _Version        =   851969
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "&Add"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   255
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchDefTable.frx":11D8
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchDefTable.frx":172A
            Key             =   "IMG_CALC"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchDefTable.frx":1C7C
            Key             =   "IMG_MATCH"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMatchDefTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As frmMatchDef
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mblnColumnDrag As Boolean
Private mblnLoading As Boolean
Private datData As clsDataAccess
Private mobjMatchRelation As clsMatchRelation
Private mcolBreakdownCols As Collection
Private mblnChangedName As Boolean


Public Sub NewRelation(pfrmParent As Form, lngBaseSelected() As Long, lngCriteriaSelected() As Long)

  Dim objTemp As clsMatchRelation
  Dim objBookmark As Variant
  Dim lngCurrentRow As Long
  Dim lngTableID As Long
  Dim lngCriteriaTableID As Long
  Dim lngCount As Long
  Dim lngAvailableTables As Long
  Dim strTemp As String

  Set mfrmParent = pfrmParent
  SetIconAndCaption

  Set mobjMatchRelation = New clsMatchRelation
  Set mcolBreakdownCols = New Collection
  mblnReadOnly = False
  
  Screen.MousePointer = vbHourglass
 
  lngAvailableTables = LoadChildTableCombo( _
    cboTables, _
    mfrmParent.cboTable1.ItemData(mfrmParent.cboTable1.ListIndex), _
    -1, lngBaseSelected(), False)

  If lngAvailableTables = 0 Then
    Screen.MousePointer = vbDefault
    COAMsgBox "You have selected the base table and all child tables of " & _
      mfrmParent.cboTable1.List(mfrmParent.cboTable1.ListIndex), vbExclamation, mfrmParent.Caption

  Else
    LoadChildTableCombo _
      cboMatchTables, _
      mfrmParent.cboTable2.ItemData(mfrmParent.cboTable2.ListIndex), _
      -1, lngCriteriaSelected(), True
  
      SetExpressions
      Screen.MousePointer = vbDefault
      Me.Show vbModal
    
  End If

  Set objBookmark = Nothing
  'Set objCurrentBookmark = Nothing
  Set objTemp = Nothing

End Sub

Public Sub EditRelation(pfrmParent As Form, lngBaseSelected() As Long, lngCriteriaSelected() As Long, pobjMatchRelation As clsMatchRelation, pblnReadOnly As Boolean)

  Dim objBookmark As Variant
  Dim objCurrentBookmark As Variant
  Dim objColumn As clsColumn
  Dim lngCount As Long

  Dim lngCriteriaTableID As Long
  Dim lngTableID As Long
  Dim strText As String

  mblnLoading = True
  Screen.MousePointer = vbHourglass

  Set mfrmParent = pfrmParent
  SetIconAndCaption
  
  
  Set mobjMatchRelation = pobjMatchRelation
  Set mcolBreakdownCols = New Collection
  
  For Each objColumn In mobjMatchRelation.BreakdownColumns
    mcolBreakdownCols.Add objColumn, objColumn.ColType & CStr(objColumn.ID)
  Next
  
  mblnReadOnly = pblnReadOnly

  With mfrmParent.grdRelations
    If .Rows > 0 Then
      lngTableID = Val(.Columns("TableID").CellValue(.Bookmark))
      lngCriteriaTableID = Val(.Columns("MatchTableID").CellValue(.Bookmark))
'
'      objCurrentBookmark = .Bookmark
'      ReDim lngBaseSelected(.Rows - 1) As Long
'      ReDim lngCriteriaSelected(.Rows - 1) As Long
'
'      For lngCount = 0 To .Rows - 1
'        objBookmark = .GetBookmark(lngCount)
'        lngBaseSelected(lngCount) = .Columns("TableID").CellText(objBookmark)
'        lngCriteriaSelected(lngCount) = .Columns("MatchTableID").CellText(objBookmark)
'      Next
'
'      .Bookmark = objCurrentBookmark
'      .SelBookmarks.Add .Bookmark
    End If
  End With

  SetExpressions

  LoadChildTableCombo _
    cboTables, _
    mfrmParent.cboTable1.ItemData(mfrmParent.cboTable1.ListIndex), _
    lngTableID, lngBaseSelected(), False

  LoadChildTableCombo _
    cboMatchTables, _
    mfrmParent.cboTable2.ItemData(mfrmParent.cboTable2.ListIndex), _
    lngCriteriaTableID, lngCriteriaSelected(), True

  'Put the columns back in the right sequence...
  For lngCount = 1 To mcolBreakdownCols.Count
    For Each objColumn In mcolBreakdownCols
      If objColumn.Sequence = lngCount Then
        If objColumn.ColType = "C" Then
          strText = GetTableNameFromColumn(objColumn.ID) & "." & datGeneral.GetColumnName(objColumn.ID)
          ListView2.ListItems.Add , objColumn.ColType & CStr(objColumn.ID), strText, , ImageList1.ListImages("IMG_TABLE").Index
        Else
          strText = datGeneral.GetExpression(objColumn.ID)
          ListView2.ListItems.Add , objColumn.ColType & CStr(objColumn.ID), strText, , ImageList1.ListImages("IMG_CALC").Index
        End If
        Exit For
      End If
    Next
  Next

  PopulateAvailable

  Set objBookmark = Nothing
  Set objCurrentBookmark = Nothing
  Set objColumn = Nothing
  mblnLoading = False

  Screen.MousePointer = vbDefault
  
  Me.Show vbModal

End Sub

Public Property Get MatchRelation() As clsMatchRelation
  Set MatchRelation = mobjMatchRelation
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let ReadOnly(blnNewValue As Boolean)
  mblnReadOnly = blnNewValue
End Property

Private Property Let Changed(blnNewValue As Boolean)
  cmdOK.Enabled = blnNewValue
End Property

Private Sub cboMatchTables_Click()
  
  If Not mblnLoading Then
    If ListView2.ListItems.Count > 0 Or _
      Val(txtRequired.Tag) > 0 Or _
      Val(txtPreferred.Tag) > 0 Or _
      Val(txtScore.Tag) > 0 Then

      If COAMsgBox("Warning: Changing a table will result in all table/column " & _
            "specific aspects of this breakdown being cleared." & vbCrLf & _
            "Are you sure you wish to continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        ClearBreakdown
      Else
        mblnLoading = True
        SetComboItem cboMatchTables, cboMatchTables.Tag
        mblnLoading = False
      End If

    End If
  End If
  
  cboMatchTables.Tag = cboMatchTables.ItemData(cboMatchTables.ListIndex)
  PopulateAvailable

End Sub

Private Sub cboTables_Click()
  
  If Not mblnLoading Then
    If ListView2.ListItems.Count > 0 Or _
      Val(txtRequired.Tag) > 0 Or _
      Val(txtPreferred.Tag) > 0 Or _
      Val(txtScore.Tag) > 0 Then

      If COAMsgBox("Warning: Changing a table will result in all table/column " & _
            "specific aspects of this breakdown being cleared." & vbCrLf & _
            "Are you sure you wish to continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        ClearBreakdown
      Else
        mblnLoading = True
        SetComboItem cboTables, cboTables.Tag
        mblnLoading = False
      End If

    End If
  End If

  cboTables.Tag = cboTables.ItemData(cboTables.ListIndex)
  PopulateAvailable

End Sub

Private Sub ClearBreakdown()

  txtRequired.Text = vbNullString
  txtRequired.Tag = vbNullString
  cmdRequiredClear.Enabled = False

  txtPreferred.Text = vbNullString
  txtPreferred.Tag = vbNullString
  cmdPreferredClear.Enabled = False

  txtScore.Text = vbNullString
  txtScore.Tag = vbNullString
  cmdScoreClear.Enabled = False

  ListView2.ListItems.Clear
  Set mcolBreakdownCols = Nothing
  Set mcolBreakdownCols = New Collection

  Changed = True

End Sub


Private Sub cmdAdd_LostFocus()
  'JPD 20040105 Fault 5827
  cmdAdd.Picture = cmdAdd.Picture

End Sub

Private Sub cmdAddAll_LostFocus()
  'JPD 20040105 Fault 5827
  cmdAddAll.Picture = cmdAddAll.Picture

End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdMoveDown_LostFocus()
  'JPD 20040105 Fault 5827
  cmdMoveDown.Picture = cmdMoveDown.Picture

End Sub

Private Sub cmdMoveUp_LostFocus()
  'JPD 20040105 Fault 5827
  cmdMoveUp.Picture = cmdMoveUp.Picture

End Sub

Private Sub cmdPreferred_Click()
  
  EditExpression _
    cboTables.ItemData(cboTables.ListIndex), _
    cboMatchTables.ItemData(cboMatchTables.ListIndex), _
    txtPreferred, giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC
  cmdPreferredClear.Enabled = (Val(txtPreferred.Tag) > 0)

End Sub

Private Sub cmdRemove_LostFocus()
  'JPD 20040105 Fault 5827
  cmdRemove.Picture = cmdRemove.Picture

End Sub

Private Sub cmdRemoveAll_LostFocus()
  'JPD 20040105 Fault 5827
  cmdRemoveAll.Picture = cmdRemoveAll.Picture

End Sub

Private Sub cmdRequiredClear_Click()
  
  Dim intMBResponse As Integer
  
  intMBResponse = COAMsgBox("Are you sure you want to clear the Required expression?", vbExclamation + vbYesNo, mfrmParent.Caption)
  
  If intMBResponse = vbYes Then
    'mfrmParent.ExprDeleteOnOK cboTables.ItemData(cboTables.ListIndex), Val(txtRequired.Tag), giEXPR_MATCHWHEREEXPRESSION
    mfrmParent.ExprDeleteOnOK cboMatchTables.ItemData(cboMatchTables.ListIndex), Val(txtRequired.Tag), giEXPR_MATCHWHEREEXPRESSION
    
    cmdRequiredClear.Enabled = False
    txtRequired.Text = "<None>"
    txtRequired.Tag = 0
    Changed = True
  End If

End Sub

Private Sub cmdPreferredClear_Click()
  
  Dim intMBResponse As Integer
  
  intMBResponse = COAMsgBox("Are you sure you want to clear the Preferred expression?", vbExclamation + vbYesNo, mfrmParent.Caption)
  
  If intMBResponse = vbYes Then
    mfrmParent.ExprDeleteOnOK cboTables.ItemData(cboTables.ListIndex), Val(txtPreferred.Tag), giEXPR_MATCHJOINEXPRESSION
    
    cmdPreferredClear.Enabled = False
    txtPreferred.Text = "<None>"
    txtPreferred.Tag = 0
    Changed = True
  End If

End Sub

Private Sub cmdScoreClear_Click()
  
  Dim lngCount As Long
  Dim intMBResponse As Integer
  
  intMBResponse = COAMsgBox("Are you sure you want to clear the Match Score expression?", vbExclamation + vbYesNo, mfrmParent.Caption)
  
  If intMBResponse = vbYes Then
    mfrmParent.ExprDeleteOnOK cboTables.ItemData(cboTables.ListIndex), Val(txtScore.Tag), giEXPR_MATCHSCOREEXPRESSION


    'Loop backwards as we are removing items thus changing indexes!
    For lngCount = ListView2.ListItems.Count To 1 Step -1
      If Left(ListView2.ListItems(lngCount).Key, 1) = "E" Then
        ListView2.ListItems.Remove lngCount
      End If
    Next

    'Loop backwards as we are removing items thus changing indexes!
    For lngCount = mcolBreakdownCols.Count To 1 Step -1
      If mcolBreakdownCols(lngCount).ColType = "E" Then
        mcolBreakdownCols.Remove lngCount
      End If
    Next


    cmdScoreClear.Enabled = False
    txtScore.Text = "<None>"
    txtScore.Tag = 0
    Changed = True
  End If

End Sub

Private Sub cmdScore_Click()

  Dim objItem As ListItem
  Dim strText As String
  Dim lngCount As Long
  Dim lngIndex As Long

  Dim lngSeq As Long
  Dim lngSize As Long
  Dim lngDec As Long

  Dim objTemp As clsColumn
  Dim objColumn As clsColumn

  EditExpression _
    cboTables.ItemData(cboTables.ListIndex), _
    cboMatchTables.ItemData(cboMatchTables.ListIndex), _
    txtScore, giEXPR_MATCHSCOREEXPRESSION, giEXPRVALUE_NUMERIC
  cmdScoreClear.Enabled = (Val(txtScore.Tag) > 0)

  PopulateAvailable

  For lngCount = 1 To ListView2.ListItems.Count
    If Left(ListView2.ListItems(lngCount).Key, 1) = "E" Then
      
      'Update Collection
      For lngIndex = 1 To mcolBreakdownCols.Count
        Set objColumn = mcolBreakdownCols(lngIndex)
        If objColumn.ColType = "E" Then
          objColumn.ID = Val(txtScore.Tag)
          objColumn.Heading = txtScore.Text
          mcolBreakdownCols.Remove lngIndex
          mcolBreakdownCols.Add objColumn, objColumn.ColType & CStr(objColumn.ID)
        End If
      Next
      
      'Update Listview2
      ListView2.ListItems.Remove lngCount
      If Val(txtScore.Tag) > 0 Then
        strText = datGeneral.GetExpression(Val(txtScore.Tag))
        ListView2.ListItems.Add lngCount, "E" & txtScore.Tag, strText, , ImageList1.ListImages("IMG_CALC").Index
        EnableColProperties Not mblnReadOnly
      End If
      Exit For

    End If
  Next

  Set objItem = Nothing

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
  SSTab1.Tab = 0
  SSTab1_Click 0
  mblnCancelled = True
  Set datData = New clsDataAccess
End Sub

Private Function SetExpressions()

  With mobjMatchRelation

    If .RequiredExprID > 0 Then
      txtRequired.Text = datGeneral.GetExpression(.RequiredExprID)
      txtRequired.Tag = .RequiredExprID
    Else
      txtRequired.Text = "<None>"
      txtRequired.Tag = 0
    End If
    cmdRequiredClear.Enabled = (.RequiredExprID > 0)
    
    If .PreferredExprID > 0 Then
      txtPreferred.Text = datGeneral.GetExpression(.PreferredExprID)
      txtPreferred.Tag = .PreferredExprID
    Else
      txtPreferred.Text = "<None>"
      txtPreferred.Tag = 0
    End If
    cmdPreferredClear.Enabled = (.PreferredExprID > 0)
  
    If .MatchScoreID > 0 Then
      txtScore.Text = datGeneral.GetExpression(.MatchScoreID)
      txtScore.Tag = .MatchScoreID
    Else
      txtScore.Text = "<None>"
      txtScore.Tag = 0
    End If
    cmdScoreClear.Enabled = (.MatchScoreID > 0)

  End With

End Function


Private Sub cmdOK_Click()

  Dim objColumn As clsColumn
  Dim strDuplicateHeading As String
  Dim lngCount As Long

  Dim lngParentTable1 As Long
  Dim lngParentTable2 As Long
  Dim lngChildTable1 As Long
  Dim lngChildTable2 As Long


  If Val(txtRequired.Tag) = 0 And Val(txtPreferred.Tag) = 0 Then
    SSTab1.Tab = 0
    COAMsgBox "You must select either a Required Match expression or Preferred Match expression.", vbExclamation + vbOKOnly, mfrmParent.Caption
    Exit Sub
  End If
  
  If ListView2.ListItems.Count = 0 Then
    SSTab1.Tab = 1
    COAMsgBox "You must select at least one column to show in the breakdown.", vbExclamation + vbOKOnly, mfrmParent.Caption
    Exit Sub
  End If
  
  
  strDuplicateHeading = mfrmParent.CheckForDuplicateHeadings(mcolBreakdownCols)
  If strDuplicateHeading <> vbNullString Then
    SSTab1.Tab = 1
    COAMsgBox "More than one column has a heading of '" & strDuplicateHeading & "'" & vbCrLf & "Column headings must be unique.", vbExclamation + vbOKOnly, mfrmParent.Caption
    Exit Sub
  End If


  lngParentTable1 = mfrmParent.cboTable1.ItemData(mfrmParent.cboTable1.ListIndex)
  lngParentTable2 = mfrmParent.cboTable2.ItemData(mfrmParent.cboTable2.ListIndex)
  lngChildTable1 = cboTables.ItemData(cboTables.ListIndex)
  lngChildTable2 = cboMatchTables.ItemData(cboMatchTables.ListIndex)
  
  'MH20040416 Fault 8473
  If datGeneral.IsAChildOf(lngChildTable1, lngParentTable1) And _
     datGeneral.IsAChildOf(lngChildTable1, lngParentTable2) Then
        SSTab1.Tab = 0
        COAMsgBox "Cannot use the '" & cboTables.Text & "' table as it is a child table of both the '" & mfrmParent.cboTable1.Text & "' and the '" & mfrmParent.cboTable2.Text & "' tables.", vbExclamation + vbOKOnly, mfrmParent.Caption
        Exit Sub
  End If

  If datGeneral.IsAChildOf(lngChildTable2, lngParentTable1) And _
     datGeneral.IsAChildOf(lngChildTable2, lngParentTable2) Then
        SSTab1.Tab = 0
        COAMsgBox "Cannot use the '" & cboMatchTables.Text & "' table as it is a child table of both the '" & mfrmParent.cboTable1.Text & "' and the '" & mfrmParent.cboTable2.Text & "' tables.", vbExclamation + vbOKOnly, mfrmParent.Caption
        Exit Sub
  End If


  Screen.MousePointer = vbHourglass
  
  mobjMatchRelation.Table1ID = lngChildTable1
  mobjMatchRelation.Table2ID = lngChildTable2
  mobjMatchRelation.RequiredExprID = Val(txtRequired.Tag)
  mobjMatchRelation.PreferredExprID = Val(txtPreferred.Tag)
  mobjMatchRelation.MatchScoreID = Val(txtScore.Tag)
  
  
  'Re-populate the breakdown column array
  mobjMatchRelation.BreakdownColumns = mcolBreakdownCols
  
  'Remember the column sequence
  For lngCount = 1 To ListView2.ListItems.Count
    mcolBreakdownCols(ListView2.ListItems(lngCount).Key).Sequence = lngCount
  Next

  mblnCancelled = False
  Me.Hide
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdRequired_Click()
  
  Dim lngTable1ID As Long
  Dim lngTable2ID As Long
  
  lngTable1ID = cboTables.ItemData(cboTables.ListIndex)
  If cboMatchTables.ListIndex > 0 Then
    lngTable2ID = cboMatchTables.ItemData(cboMatchTables.ListIndex)
  Else
    lngTable2ID = 0
  End If
  
  EditExpression _
    lngTable1ID, _
    lngTable2ID, _
    txtRequired, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC
  cmdRequiredClear.Enabled = (Val(txtRequired.Tag) > 0)

End Sub

Private Sub EditExpression(lngBaseTable As Long, lngSecondTable As Long, ctlTarget As TextBox, intExprType As Integer, intReturnType As Integer)

  Dim objExpression As clsExprExpression
  
  Dim fOK As Boolean
  Dim intLoop As Integer
  Dim alngColumns() As Long
  Dim varBookmark As Variant
  Dim fCancelled As Boolean
  Dim lngExprID As Long
  'Dim lngBaseTable As Long
  Dim strTemp As String
  
  
  Set objExpression = New clsExprExpression
  
  lngExprID = Val(ctlTarget.Tag)
  'lngBaseTable = cboTables.ItemData(cboTables.ListIndex)
  
  With objExpression
    ' Initialise the expression object.
    
    fOK = .Initialise(lngBaseTable, lngExprID, intExprType, intReturnType, lngSecondTable)
    
    If fOK Then
      ' Construct an array of the columns in the import definition.
      ReDim alngColumns(0)
      grdColumns.MoveFirst
      
      Do Until intLoop = grdColumns.Rows
        varBookmark = grdColumns.GetBookmark(intLoop)
        
        If grdColumns.Columns("ColExprID").CellText(varBookmark) > 0 Then
          ReDim Preserve alngColumns(UBound(alngColumns) + 1)
          alngColumns(UBound(alngColumns)) = grdColumns.Columns("ColExprID").CellText(varBookmark)
        End If
        
        intLoop = intLoop + 1
      Loop
      
      .ColumnList = alngColumns
    End If
    
    If fOK Then
      If Val(ctlTarget.Tag) > 0 Then
        .EditExpression fCancelled
      Else
        .NewExpression fCancelled
      End If
      
      ' Read the selected expression info.
      strTemp = IIf(.ExpressionID > 0, .Name, "<None>")
      If ctlTarget.Text <> strTemp Then
        ctlTarget.Text = strTemp
        mblnChangedName = True
      End If
      
      If (Val(ctlTarget.Tag) <> .ExpressionID) Or (Not fCancelled) Then
        Changed = True
        
        If Val(ctlTarget.Tag) = 0 And .ExpressionID > 0 Then
          'Remember new expression
          mfrmParent.ExprDeleteOnCancel lngBaseTable, .ExpressionID, intExprType
        End If
      End If
      
      ctlTarget.Tag = .ExpressionID
      'mlngOriginalFilterID = .ExpressionID
    End If
  End With

  Set objExpression = Nothing

End Sub


Private Sub PopulateAvailable()
  
  Dim objColumnPrivileges As CColumnPrivileges
  Dim objColumn As CColumnPrivilege
  Dim rsCalculations As New Recordset
  Dim sSQL As String
  Dim intCount As Integer
  Dim fOK As Boolean
  Dim objItem As ListItem
  Dim blnFound As Boolean
  
  'If Me.cboTables.ListIndex = -1 Then
  '  Exit Sub
  'End If

  'ListView1.ListItems.Clear
  
  'With ListView1.ListItems
  '  .Clear
  
  Dim fColumnOK As Boolean
  Dim sSource As String
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  
  Dim sRealSource As String
  Dim sCaseStatement As String
  Dim sWhereColumn As String
  Dim sBaseIDColumn As String

  Dim asViews() As String
  
  'On Error GoTo LocalErr
  
  
  With ListView1.ListItems
    .Clear

    'PopulateAvailableFromTable
    
    Set objColumnPrivileges = GetColumnPrivileges(cboTables.Text)
    For Each objColumn In objColumnPrivileges
      If objColumn.ColumnType <> colSystem And objColumn.ColumnType <> colLink Then
        If objColumn.DataType <> sqlVarBinary And objColumn.DataType <> sqlOle Then
          If Not AlreadyUsed(CStr("C" & objColumn.ColumnID)) Then
            .Add , "C" & CStr(objColumn.ColumnID), cboTables.Text & "." & objColumn.ColumnName, , ImageList1.ListImages("IMG_TABLE").Index
          End If
        End If
      End If
    Next

    If cboMatchTables.ListIndex > 0 Then
      If cboTables.Text <> cboMatchTables.Text Then
        Set objColumnPrivileges = GetColumnPrivileges(cboMatchTables.Text)
        For Each objColumn In objColumnPrivileges
          If objColumn.ColumnType <> colSystem And objColumn.ColumnType <> colLink Then
            If objColumn.DataType <> sqlVarBinary And objColumn.DataType <> sqlOle Then
              If Not AlreadyUsed(CStr("C" & objColumn.ColumnID)) Then
                .Add , "C" & CStr(objColumn.ColumnID), cboMatchTables.Text & "." & objColumn.ColumnName, , ImageList1.ListImages("IMG_TABLE").Index
              End If
            End If
          End If
        Next
      End If
    End If


    'This will sort the columns into alphabetical order
    'but allow the calc to be added to the end...
    ListView1.Sorted = True
    ListView1.Sorted = False

    If Val(txtScore.Tag) > 0 Then
      
      'Check if the match score is used
      blnFound = False
      For Each objItem In ListView2.ListItems
        If Left(objItem.Key, 1) = "E" Then
          blnFound = True
          Exit For
        End If
      Next objItem
      Set objItem = Nothing
      
      If Not blnFound Then
        .Add , "E" & txtScore.Tag, txtScore.Text, , ImageList1.ListImages("IMG_CALC").Index
      End If
    End If

    
    If .Count > 0 Then
      .Item(1).Selected = True
    End If

  End With

  
  Set objColumnPrivileges = Nothing
  Set objColumn = Nothing
  Set rsCalculations = Nothing

  UpdateButtonStatus SSTab1.Tab

  'cmdPreferred.Enabled = (cboMatchTables.ListIndex > 0)

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mcolBreakdownCols = Nothing
  Set datData = Nothing
End Sub

Private Sub ListView1_GotFocus()
  cmdAdd.Default = True
End Sub

Private Sub ListView1_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub ListView2_GotFocus()
  cmdRemove.Default = True
End Sub

Private Sub ListView2_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub ListView1_DblClick()
  If Not mblnReadOnly Then
    CopyToSelected False
  End If
End Sub


Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim objItem As ListItem
  
  If Shift = vbCtrlMask And KeyCode = 65 Then
    For Each objItem In ListView1.ListItems
      objItem.Selected = True
    Next objItem
    Set objItem = Nothing
  ElseIf KeyCode = vbKeyReturn Then
    If cmdAdd.Enabled Then
      cmdAdd_Click
      ListView1.SetFocus
    End If
  End If
  
End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim objItem As ListItem
  
  If Shift = vbCtrlMask And KeyCode = 65 Then
    For Each objItem In ListView2.ListItems
      objItem.Selected = True
    Next objItem
    Set objItem = Nothing
    UpdateButtonStatus (Me.SSTab1.Tab)
  ElseIf KeyCode = vbKeyReturn Then
    If cmdRemove.Enabled Then
      cmdRemove_Click
      ListView2.SetFocus
    End If
  End If
  
End Sub

Private Sub ListView2_DblClick()
  If Not mblnReadOnly Then
    CopyToAvailable False
  End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Not mblnReadOnly Then
    If mblnColumnDrag Then
      ListView1.Drag vbCancel
      mblnColumnDrag = False
    End If
  End If

End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
  UpdateButtonStatus (Me.SSTab1.Tab)
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Not mblnReadOnly Then
    If mblnColumnDrag Then
      ListView2.Drag vbCancel
      mblnColumnDrag = False
    End If
  End If
  
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim objItem As ComctlLib.ListItem

  If Not mblnReadOnly Then
    If Button = vbLeftButton Then
      If ListView1.ListItems.Count > 0 Then
        mblnColumnDrag = True
        ListView1.Drag vbBeginDrag
      End If
    End If
  End If

End Sub

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim objItem As ComctlLib.ListItem

  If Not mblnReadOnly Then
    If Button = vbLeftButton Then
      If ListView2.ListItems.Count > 0 Then
        mblnColumnDrag = True
        ListView2.Drag vbBeginDrag
      End If
    End If
  End If

End Sub

Private Sub ListView1_DragDrop(Source As Control, x As Single, y As Single)
  
  If Source Is ListView2 Then
    CopyToAvailable False
    ListView2.Drag vbCancel
  Else
    ListView2.Drag vbCancel
  End If

End Sub

Private Sub ListView2_DragDrop(Source As Control, x As Single, y As Single)
  
  If Source Is ListView1 Then
    If ListView2.HitTest(x, y) Is Nothing Then
      CopyToSelected False
    Else
      CopyToSelected False, ListView2.HitTest(x, y).Index
    End If
    ListView1.Drag vbCancel
  Else
    If ListView2.HitTest(x, y) Is Nothing Then
      ChangeSelectedOrder
    Else
      ChangeSelectedOrder ListView2.HitTest(x, y).Index
    End If
    ListView2.Drag vbCancel
  End If

End Sub

Private Sub Frafieldsavailable_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Source.DragIcon = picNoDrop.Picture
End Sub

Private Sub Frafieldsselected_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Source.DragIcon = picNoDrop.Picture
End Sub

Private Sub ListView2_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument.Picture
  End If
  Set ListView2.DropHighlight = ListView2.HitTest(x, y)
End Sub

Private Sub ListView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  If (Source Is ListView1) Or (Source Is ListView2) Then
    Source.DragIcon = picDocument.Picture
  End If
End Sub

Private Function CopyToSelected(bAll As Boolean, Optional intBeforeIndex As Integer)

  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  Dim fOK As Boolean
  Dim iItemSelectedCount As Integer
  Dim objCalcExpr As clsExprExpression
  Dim objTempItem As ListItem
  Dim objWorkingItem As ListItem
  Dim iSelectedCount As Integer
  Dim iItemToSelect As Integer
  Dim prstTemp As Recordset
  Dim iItemsToDelete() As Variant
  ReDim iItemsToDelete(0)
  Dim intTemp As Integer
  
  Dim sTempTableName As String
  Dim lngColumnID As Long
  Dim lngTableID As Long
  
  Screen.MousePointer = vbHourglass

  'If user has clicked ADD ALL then do this...
  If bAll Then
    For Each objTempItem In ListView1.ListItems
      ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
      fOK = True
      If fOK Then AddToCollection objTempItem
    Next objTempItem
    
    ListView1.ListItems.Clear
    SelectFirst ListView2
    UpdateButtonStatus (Me.SSTab1.Tab)
    'ForceDefinitionToBeHiddenIfNeeded
    Screen.MousePointer = vbNormal
    Changed = True
    Exit Function
  End If
  
  'Get count of how many items we are moving
  For Each objTempItem In ListView1.ListItems
    If objTempItem.Selected = True Then
      iSelectedCount = iSelectedCount + 1
      If iSelectedCount = 1 Then
        Set objWorkingItem = objTempItem
        iItemToSelect = objWorkingItem.Index
      End If
    End If
  Next objTempItem
  
  'If its just one item do this...
  If iSelectedCount = 1 Then
    
    Set objTempItem = objWorkingItem
    
    'If we are not inserting it before existing columns...
    If intBeforeIndex = 0 Then
      ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
      fOK = True
      If fOK Then
        AddToCollection objTempItem
        ListView1.ListItems.Remove iItemToSelect
      End If
    
    Else
      ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
      fOK = True
      If fOK Then
        AddToCollection objTempItem
        ListView1.ListItems.Remove iItemToSelect
      End If
   
    End If

    If ListView1.ListItems.Count > 0 Then
      If iItemToSelect > ListView1.ListItems.Count Then
        iItemToSelect = ListView1.ListItems.Count
      End If
      ListView1.ListItems(iItemToSelect).Selected = True
    End If
    
    If intBeforeIndex = 0 Then
      SelectLast ListView2
    Else
      For Each objTempItem In ListView2.ListItems
        objTempItem.Selected = (objTempItem.Index = intBeforeIndex)
      Next objTempItem
      Set ListView2.DropHighlight = Nothing
    End If
    
    UpdateButtonStatus (Me.SSTab1.Tab)
    'ForceDefinitionToBeHiddenIfNeeded
    Screen.MousePointer = vbNormal
    Changed = True
    Exit Function
  End If

  'There are more than one item selected
  For Each objTempItem In ListView1.ListItems
    
    If objTempItem.Selected Then
      
      If intBeforeIndex = 0 Then
        ListView2.ListItems.Add , objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
        fOK = True
      Else
        ListView2.ListItems.Add intBeforeIndex, objTempItem.Key, objTempItem.Text, objTempItem.Icon, objTempItem.SmallIcon
        fOK = True
        intBeforeIndex = intBeforeIndex + 1
      End If
    
      If fOK = True Then
        AddToCollection objTempItem
        ReDim Preserve iItemsToDelete(UBound(iItemsToDelete) + 1)
        iItemsToDelete(UBound(iItemsToDelete) - 1) = objTempItem.Index
      End If
    
    End If

  Next objTempItem

  ' Remove the selected items from the available listview
  For intTemp = UBound(iItemsToDelete) - 1 To 0 Step -1
    ListView1.ListItems.Remove iItemsToDelete(intTemp)
  Next intTemp
  
  ' Select the top available item in the listview
  If ListView1.ListItems.Count > 0 Then ListView1.ListItems(1).Selected = True
  
  If intBeforeIndex = 0 Then
    SelectLast ListView2
  Else
    For Each objTempItem In ListView2.ListItems
      objTempItem.Selected = (objTempItem.Index = intBeforeIndex)
    Next objTempItem
    Set ListView2.DropHighlight = Nothing
  End If

  UpdateButtonStatus (Me.SSTab1.Tab)
  'ForceDefinitionToBeHiddenIfNeeded
  Screen.MousePointer = vbNormal
  Changed = True
  
End Function


Private Function CopyToAvailable(bAll As Boolean, Optional intBeforeIndex As Integer)

  Dim iLoop As Integer
  Dim iTempItemIndex As Integer
  
  ' Dont add the to the first listview...just remove em and
  ' repopulate the available listview...much quicker
  
  Screen.MousePointer = vbHourglass
  
  For iLoop = ListView2.ListItems.Count To 1 Step -1
    If Not bAll Then
      If ListView2.ListItems(iLoop).Selected Then
        iTempItemIndex = iLoop
        mcolBreakdownCols.Remove ListView2.ListItems(iLoop).Key
        ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
      End If
    Else
      mcolBreakdownCols.Remove ListView2.ListItems(iLoop).Key
      ListView2.ListItems.Remove ListView2.ListItems(iLoop).Key
    
    End If
  Next iLoop
  
  If ListView2.ListItems.Count > 0 Then
    If iTempItemIndex > ListView2.ListItems.Count Then iTempItemIndex = ListView2.ListItems.Count
    If iTempItemIndex > 0 Then ListView2.ListItems(iTempItemIndex).Selected = True
  End If
  
  PopulateAvailable
  
  UpdateButtonStatus (Me.SSTab1.Tab)
  'UpdateOrderButtons
  'ForceDefinitionToBeHiddenIfNeeded

  Changed = True

  Screen.MousePointer = vbNormal

End Function


Private Function UpdateButtonStatus(iTab As Integer)

  On Error Resume Next
  
  Dim tempItem As ListItem, iCount As Integer
  
  'Select Case iTab
  'Case 2:
    ' If there are no items to be selected, disable Add buttons
    If ListView1.ListItems.Count = 0 Then
      cmdAdd.Enabled = False
      cmdAddAll.Enabled = False
    Else
      cmdAdd.Enabled = Not mblnReadOnly
      cmdAddAll.Enabled = Not mblnReadOnly
    End If
    
    ' If there are no items in the 'Selected' Listview then disable move buttons and exit
    If ListView2.ListItems.Count = 0 Then
      cmdMoveUp.Enabled = False
      cmdMoveDown.Enabled = False
      cmdRemove.Enabled = False
      cmdRemoveAll.Enabled = False
      EnableColProperties False
    Else
      cmdRemove.Enabled = Not mblnReadOnly
      cmdRemoveAll.Enabled = Not mblnReadOnly
    
    ' If there are more than 1 items selected then disable the move buttons and exit
    For Each tempItem In ListView2.ListItems
      If tempItem.Selected Then iCount = iCount + 1
    Next tempItem
    
    'Debug.Print Now() & vbTab & " - " & icount
    
    'If iCount > 1 Then
    If iCount > 1 Or iCount < 1 Then
      cmdMoveUp.Enabled = False
      cmdMoveDown.Enabled = False
      EnableColProperties False
    Else
      If ListView2.SelectedItem.Index <> 1 Then cmdMoveUp.Enabled = Not mblnReadOnly Else cmdMoveUp.Enabled = False
      If ListView2.SelectedItem.Index <> ListView2.ListItems.Count Then cmdMoveDown.Enabled = Not mblnReadOnly Else cmdMoveDown.Enabled = False
      EnableColProperties Not mblnReadOnly
    End If
    
   End If

  
'  Case 1:
'    If grdChildren.Rows = 0 Then
'      cmdAddChild.Enabled = fraChild.Enabled
'      cmdEditChild.Enabled = False
'      cmdRemoveChild.Enabled = False
'      cmdRemoveAllChilds.Enabled = False
'    Else
'      If grdChildren.SelBookmarks.Count > 0 Then
'        cmdEditChild.Enabled = Not mblnReadOnly
'        cmdEditChild.SetFocus
'        cmdRemoveChild.Enabled = Not mblnReadOnly
'      Else
'        cmdEditChild.Enabled = False
'        cmdRemoveChild.Enabled = False
'      End If
'      cmdRemoveAllChilds.Enabled = Not mblnReadOnly
'    End If
'
'  End Select

  'TM20020508 Fault 3790
  Call CheckListViewColWidth(ListView1)
  Call CheckListViewColWidth(ListView2)

End Function

Private Function SelectLast(lvwCtl As ListView)

  Dim objItem As ListItem
  
  For Each objItem In lvwCtl.ListItems
    objItem.Selected = IIf(objItem.Index = lvwCtl.ListItems.Count, True, False)
  Next objItem

End Function

Private Function SelectFirst(lvwCtl As ListView)

  Dim objItem As ListItem
  
  For Each objItem In lvwCtl.ListItems
    objItem.Selected = IIf(objItem.Index = 1, True, False)
  Next objItem

End Function


Private Sub cmdAdd_Click()
  CopyToSelected False
End Sub

Private Sub cmdMoveDown_Click()
  ChangeSelectedOrder ListView2.SelectedItem.Index + 2, True
End Sub

Private Sub cmdMoveUp_Click()
  ChangeSelectedOrder ListView2.SelectedItem.Index - 1
End Sub

Private Sub cmdRemove_Click()
  CopyToAvailable False
  If ListView2.ListItems.Count = 0 Then EnableColProperties False
End Sub

Private Sub cmdAddAll_Click()
  CopyToSelected True
End Sub

Private Sub cmdRemoveAll_Click()

  ' Remove All items from the 'Selected' Listview
'  If Me.grdReportOrder.Rows > 0 Then
'    If COAMsgBox("Removing all selected report columns will also clear the report sort order." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, "Custom Reports") = vbYes Then
'      CopyToAvailable True
'      EnableColProperties False
'    End If
'  Else
    If COAMsgBox("Are you sure you wish to remove all breakdown columns?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
      CopyToAvailable True
      EnableColProperties False
    End If
  'End If
  
End Sub

Private Function EnableColProperties(bStatus As Boolean)
  
  Dim objMatchColumn As clsColumn
  Dim lngColumnID As Long
  
  mblnLoading = True
  
  spnSize.Enabled = bStatus
  spnSize.BackColor = IIf(bStatus, vbWindowBackground, vbButtonFace)
  txtHeading.Enabled = bStatus
  txtHeading.BackColor = IIf(bStatus, vbWindowBackground, vbButtonFace)
  spnSize.Value = 0
  spnDec.Value = 0
  txtHeading.Text = vbNullString

  If bStatus = True Then
    If Not ListView2.SelectedItem Is Nothing Then
      
      lngColumnID = Val(Mid(ListView2.SelectedItem.Key, 2))
      
      For Each objMatchColumn In mcolBreakdownCols
        If objMatchColumn.ID = lngColumnID Then
          spnSize.Value = objMatchColumn.Size
          spnDec.Value = objMatchColumn.DecPlaces
          txtHeading.Text = objMatchColumn.Heading
          bStatus = objMatchColumn.IsNumeric
          If bStatus = True Then
            Exit For
          End If
        End If
      Next
    
    End If
  End If
  
  spnDec.Enabled = bStatus
  spnDec.BackColor = IIf(bStatus, vbWindowBackground, vbButtonFace)
  
  mblnLoading = False
  Set objMatchColumn = Nothing

End Function

Private Function AlreadyUsed(strKey As String) As Boolean

  Dim objItem As ListItem
  
  For Each objItem In ListView2.ListItems
    If objItem.Key = strKey Then
      AlreadyUsed = True
      Exit For
    End If
  Next objItem
  
  Set objItem = Nothing
  
End Function


Private Sub CheckListViewColWidth(lstvw As ListView)

  Dim objItem As ListItem
  Dim lngMax As Long
  Dim lngLen As Long
  Dim lngSelectedItem As Long
  
  lngMax = 0
  lngSelectedItem = 0

  If lstvw.ListItems.Count > 0 Then
    
    For Each objItem In lstvw.ListItems
      If lngSelectedItem = 0 And objItem.Selected Then
        objItem.Selected = True
        lngSelectedItem = objItem.Index
      End If
    
      lngLen = Me.TextWidth(objItem.Text)
      If lngMax < lngLen Then
        lngMax = lngLen
      End If
    Next

    If lngSelectedItem = 0 Then
      lstvw.ListItems(1).Selected = True
    End If
  
  End If

  lngMax = lngMax + 60
  lstvw.ColumnHeaders(1).Width = lngMax
  lstvw.Refresh

End Sub


Private Function ChangeSelectedOrder(Optional intBeforeIndex As Integer, Optional mfFromButtons As Boolean)

  ' SUB COMPLETED 28/01/00
  ' This function changes the order of listitems in the selected listview.
  ' At the moment, different arrays are used depending on what information you
  ' need to store...change the array to a type if it would suit the purpose
  ' better
  
  ' Dimension arrays
  Dim iLoop As Integer, Key() As String, Text() As String, Icon() As Variant, SmallIcon() As Variant
  ReDim Key(0), Text(0), Icon(0), SmallIcon(0)
  
  ' Clear the highlight
  Set ListView2.DropHighlight = Nothing
  
  ' If drop point is below all other items, then fix the intbeforeindex var
  If intBeforeIndex = 0 Then intBeforeIndex = ListView2.ListItems.Count + 1
  
  ' First get all the items that are above the drop point that arent selected
  For iLoop = 1 To (intBeforeIndex - 1)
    If ListView2.ListItems(iLoop).Selected = False Then
      ReDim Preserve Key(UBound(Key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
      Key(UBound(Key) - 1) = ListView2.ListItems(iLoop).Key
      Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
    End If
  Next iLoop
  
  ' Now get all the items that are selected
  For iLoop = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(iLoop).Selected = True Then
      ReDim Preserve Key(UBound(Key) + 1)
      ReDim Preserve Text(UBound(Text) + 1)
      ReDim Preserve Icon(UBound(Icon) + 1)
      ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
      Key(UBound(Key) - 1) = ListView2.ListItems(iLoop).Key
      Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
      Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
      SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
    End If
  Next iLoop
  
  ' Now get all the items below the drop point that arent selected
  If intBeforeIndex <> 0 Then
    For iLoop = (intBeforeIndex) To ListView2.ListItems.Count
      If ListView2.ListItems(iLoop).Selected = False Then
        ReDim Preserve Key(UBound(Key) + 1)
        ReDim Preserve Text(UBound(Text) + 1)
        ReDim Preserve Icon(UBound(Icon) + 1)
        ReDim Preserve SmallIcon(UBound(SmallIcon) + 1)
        Key(UBound(Key) - 1) = ListView2.ListItems(iLoop).Key
        Text(UBound(Text) - 1) = ListView2.ListItems(iLoop).Text
        Icon(UBound(Icon) - 1) = ListView2.ListItems(iLoop).Icon
        SmallIcon(UBound(SmallIcon) - 1) = ListView2.ListItems(iLoop).SmallIcon
      End If
    Next iLoop
  End If
  
  ' Clear all items from the listview
  ListView2.ListItems.Clear
  
  ' Add items in the right order from the array
  For iLoop = LBound(Key) To (UBound(Key) - 1)
    ListView2.ListItems.Add , Key(iLoop), Text(iLoop), Icon(iLoop), SmallIcon(iLoop)
  Next iLoop
  
  If mfFromButtons = True Then
    ListView2.ListItems(intBeforeIndex - 1).Selected = True
  Else
    If intBeforeIndex < ListView2.ListItems.Count Then ListView2.ListItems(intBeforeIndex).Selected = True Else ListView2.ListItems(ListView2.ListItems.Count).Selected = True
  End If
  
  mfFromButtons = False
  
  Changed = True
  
  UpdateButtonStatus (Me.SSTab1.Tab)

End Function

Private Function GetTableNameFromColumn(lngColumnID As Long) As String

  Dim rsInfo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysTables.TableName " & _
           "FROM ASRSysColumns JOIN ASRSysTables " & _
           "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           "WHERE ColumnID = " & CStr(lngColumnID)

  Set rsInfo = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableNameFromColumn = rsInfo!TableName
  
  Set rsInfo = Nothing

End Function

Private Function GetTableIDFromColumn(lngColumnID As Long) As String

  Dim rsInfo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysTables.TableID " & _
           "FROM ASRSysColumns JOIN ASRSysTables " & _
           "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           "WHERE ColumnID = " & CStr(lngColumnID)

  Set rsInfo = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableIDFromColumn = rsInfo!TableID
  
  Set rsInfo = Nothing

End Function


Private Sub AddToCollection(objTemp As Object)

  Dim lngColumnID As Long
  Dim rsTemp As Recordset
  Dim objColumnPrivileges As CColumnPrivileges
  Dim objColumn As CColumnPrivilege
  Dim blnNumeric As Boolean
  Dim strType As String
  Dim strHeading As String
  Dim objTempColumn As clsColumn

  If mblnLoading = True Then
    Exit Sub
  End If

  mblnLoading = True
  strType = Left(objTemp.Key, 1)
  lngColumnID = Val(Mid(objTemp.Key, 2))
  blnNumeric = False

  If strType = "C" Then
    Set objColumnPrivileges = GetColumnPrivileges(GetTableNameFromColumn(lngColumnID))
    Set objColumn = objColumnPrivileges.Item(datGeneral.GetColumnName(lngColumnID))
    
    strHeading = GetTableNameFromColumn(objColumn.ColumnID) & "." & objColumn.ColumnName
    txtHeading.Text = strHeading
  
    Select Case objColumn.DataType
    Case sqlNumeric, sqlInteger
      spnSize.Text = objColumn.DisplaySize
      spnDec.Text = objColumn.Decimals
      blnNumeric = True
      
    Case sqlBoolean
      spnSize.Text = 1
      spnDec.Text = 0
    
    Case sqlLongVarChar         ' Working Pattern
      spnSize.Text = 14
      spnDec.Text = 0
    
    Case Else                   ' Dates etc.
      spnSize.Text = objColumn.DisplaySize
      spnDec.Text = 0
  
    End Select

  Else
    strHeading = txtScore.Text
    spnSize.Text = 0
    spnDec.Text = 2
    blnNumeric = True

  End If
  
  
  Set objTempColumn = New clsColumn
  
  objTempColumn.ID = lngColumnID
  objTempColumn.ColType = strType
  objTempColumn.Sequence = 0
  objTempColumn.Heading = strHeading
  objTempColumn.Size = Val(spnSize.Text)
  objTempColumn.DecPlaces = Val(spnDec.Text)
  objTempColumn.IsNumeric = blnNumeric
  
  mcolBreakdownCols.Add objTempColumn, strType & CStr(lngColumnID)
  
  
  Set objTempColumn = Nothing
  Set objColumn = Nothing
  Set objColumnPrivileges = Nothing
  
  mblnLoading = False

End Sub

Private Sub spnDec_Change()
  
  Dim objItem As clsColumn
  
  If Not mblnLoading Then
    Changed = True
    Set objItem = mcolBreakdownCols.Item(ListView2.SelectedItem.Key)
    objItem.DecPlaces = spnDec.Value
    Set objItem = Nothing
  End If

End Sub

Private Sub spnSize_Change()
  
  Dim objItem As clsColumn
  
  If Not mblnLoading Then
    Changed = True
    Set objItem = mcolBreakdownCols.Item(ListView2.SelectedItem.Key)
    objItem.Size = spnSize.Value
    Set objItem = Nothing
  End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

  'fraColumns.Enabled = (SSTab1.Tab = 0)
  'fraExpressions.Enabled = (SSTab1.Tab = 0)
  'fraFieldButtons.Enabled = (SSTab1.Tab = 1)
  'fraFieldsAvailable.Enabled = (SSTab1.Tab = 1)
  'fraFieldsSelected.Enabled = (SSTab1.Tab = 1)

  Dim ctl As Control

  If Not mblnReadOnly Then
    For Each ctl In Me.Controls
      If TypeOf ctl Is VB.Frame Then
        ctl.Enabled = ctl.Left >= 0
      End If
    Next
    UpdateButtonStatus SSTab1.Tab
  End If

End Sub

Private Sub txtHeading_Change()
  
  Dim objItem As clsColumn
  
  If Not mblnLoading Then
    Changed = True
    Set objItem = mcolBreakdownCols.Item(ListView2.SelectedItem.Key)
    objItem.Heading = txtHeading.Text
    Set objItem = Nothing
  End If

End Sub

Private Function LoadChildTableCombo(cboTemp As ComboBox, lngParentTable As Long, lngSelectedID As Long, lngExcludeTables() As Long, blnAllowNone As Boolean) As Long

  Dim objTableView As CTablePrivilege
  Dim lngBaseTableID As Long
  Dim blnAlreadyAdded As Boolean
  Dim lngCount As Long

  mblnLoading = True

  lngBaseTableID = mfrmParent.cboTable1.ItemData(mfrmParent.cboTable1.ListIndex)

  With cboTemp
    .Clear

    If blnAllowNone Then
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End If

    For Each objTableView In gcoTablePrivileges.Collection
      If objTableView.IsTable Then
        If datGeneral.IsAChildOf(objTableView.TableID, lngParentTable) Or objTableView.TableID = lngParentTable Then

          blnAlreadyAdded = False
          For lngCount = 0 To UBound(lngExcludeTables)
            If objTableView.TableID = lngExcludeTables(lngCount) And objTableView.TableID <> lngSelectedID Then
              blnAlreadyAdded = True
              Exit For
            End If
          Next

          If blnAlreadyAdded = False Then
            .AddItem objTableView.TableName
            .ItemData(.NewIndex) = objTableView.TableID
          End If

        End If
      End If
    Next

    If .ListCount > 0 Then
      SetComboItem cboTemp, lngSelectedID
      If .ListIndex < 0 Then
        SetComboItem cboTemp, lngParentTable
      End If
      If .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End If

    LoadChildTableCombo = .ListCount

  End With

  mblnLoading = False

End Function


Private Sub SetIconAndCaption()
  Select Case mfrmParent.MatchReportType
    Case mrtSucession
      SetFormCaption Me, "Succession Planning Table Comparison"
    Case mrtCareer
      SetFormCaption Me, "Career Progression Table Comparison"
  End Select
  Me.HelpContextID = mfrmParent.HelpContextID
End Sub


Public Property Get ChangedName() As Boolean
  ChangedName = mblnChangedName
End Property

