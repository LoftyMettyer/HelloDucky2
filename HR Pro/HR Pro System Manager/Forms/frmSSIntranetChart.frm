VERSION 5.00
Begin VB.Form frmSSIntranetChart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Chart"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   HelpContextID   =   5086
   Icon            =   "frmSSIntranetChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraChartMultiAxes 
      Caption         =   "Multi-axis Data :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   165
      TabIndex        =   15
      Top             =   1725
      Width           =   6870
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
         Index           =   2
         Left            =   5205
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1575
         Width           =   1440
      End
      Begin VB.ComboBox cboParents 
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
         Index           =   1
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboColumns 
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
         Index           =   1
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   750
         Width           =   2670
      End
      Begin VB.ComboBox cboParents 
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
         Index           =   2
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1575
         Width           =   2670
      End
      Begin VB.ComboBox cboColumns 
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
         Index           =   2
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1995
         Width           =   2670
      End
      Begin VB.ComboBox cboSortOrder 
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
         Index           =   1
         Left            =   5205
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   750
         Width           =   1440
      End
      Begin VB.ComboBox cboSortOrder 
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
         Index           =   2
         Left            =   5205
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1995
         Width           =   1440
      End
      Begin VB.Label lblIntersectionHeader 
         AutoSize        =   -1  'True
         Caption         =   "Intersection Data :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   30
         Top             =   1260
         Width           =   1785
      End
      Begin VB.Label lblIntersectionTable 
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
         TabIndex        =   29
         Top             =   1635
         Width           =   600
      End
      Begin VB.Label lblIntersectionColumn 
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
         TabIndex        =   28
         Top             =   2025
         Width           =   795
      End
      Begin VB.Label lblIntersectionAggregate 
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
         Left            =   3945
         TabIndex        =   27
         Top             =   1635
         Width           =   1020
      End
      Begin VB.Label lblIntersectionSortOrder 
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
         Left            =   3945
         TabIndex        =   26
         Top             =   2025
         Width           =   1050
      End
      Begin VB.Label Label4 
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
         TabIndex        =   25
         Top             =   405
         Width           =   600
      End
      Begin VB.Label Label3 
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
         TabIndex        =   24
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Left            =   3945
         TabIndex        =   23
         Top             =   795
         Width           =   1050
      End
   End
   Begin VB.Frame fraChartFilter 
      Caption         =   "Chart Filter :"
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
      Left            =   165
      TabIndex        =   11
      Top             =   4455
      Width           =   6870
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
         Left            =   5235
         TabIndex        =   14
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
         Left            =   1170
         TabIndex        =   13
         Top             =   360
         Width           =   4065
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
         Left            =   5565
         MaskColor       =   &H000000FF&
         TabIndex        =   12
         ToolTipText     =   "Clear Path"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label Label9 
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
         TabIndex        =   31
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
      Left            =   4515
      TabIndex        =   2
      Top             =   5565
      Width           =   1200
   End
   Begin VB.Frame fraSSIChart 
      Caption         =   "Single-axis / Horizontal Data :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   165
      TabIndex        =   3
      Top             =   210
      Width           =   6870
      Begin VB.ComboBox cboParents 
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
         Index           =   0
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   375
         Width           =   2670
      End
      Begin VB.ComboBox cboSortOrder 
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
         Index           =   0
         Left            =   5205
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   750
         Width           =   1440
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
         Index           =   0
         Left            =   5205
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   375
         Width           =   1440
      End
      Begin VB.ComboBox cboColumns 
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
         Index           =   0
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   750
         Width           =   2670
      End
      Begin VB.Label lblSortorder 
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
         Left            =   3975
         TabIndex        =   8
         Top             =   810
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
         Left            =   3975
         TabIndex        =   7
         Top             =   420
         Width           =   1020
      End
      Begin VB.Label lblColumn 
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
         TabIndex        =   6
         Top             =   810
         Width           =   795
      End
      Begin VB.Label lblParents 
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
         TabIndex        =   5
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
      Left            =   5805
      TabIndex        =   4
      Top             =   5565
      Width           =   1200
   End
End
Attribute VB_Name = "frmSSIntranetChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mblnCancelled As Boolean

' Component definition variables.
Private mfChanged As Boolean
Private mfLoading As Boolean
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

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property



Private Sub cboAggregateType_Click(Index As Integer)
  mfChanged = True
  If Not mfLoading Then
  ChartAggregateType = cboAggregateType(Index).ItemData(cboAggregateType(Index).ListIndex)
   RefreshControls
 End If
End Sub

Private Sub cboColumns_Click(Index As Integer)
  Dim piColumnDataType As Integer
  Dim lngColumnID As Long
  Dim piIndex As Integer
  
  mfChanged = True
  
  Select Case Index
    Case 0  ' Horizontal table
      ChartColumnID = cboColumns(0).ItemData(cboColumns(0).ListIndex)
    Case 1
      Chart_ColumnID_2 = cboColumns(1).ItemData(cboColumns(1).ListIndex)
    Case 2
      Chart_ColumnID_3 = cboColumns(2).ItemData(cboColumns(2).ListIndex)
  End Select
  
  If cboColumns(2) <> vbNullString Then
    lngColumnID = cboColumns(2).ItemData(cboColumns(2).ListIndex)
  Else
    lngColumnID = cboColumns(0).ItemData(cboColumns(0).ListIndex)
  End If

  piIndex = IIf(Index = 1, 2, Index)
  fOK = PopulateAggregateCombo(0, piIndex)
  
  ' Disable relevant controls
  ' i.e. if no vertical table/column disable intersection stuff
  ' if intersection selected, disable horizontal aggregate
  
  EnableDisableCombos
  
  RefreshControls
End Sub

Private Sub EnableDisableCombos()
  
  ' Disable 'Intersection' section
  lblIntersectionHeader.Enabled = Not (cboColumns(1) = vbNullString)
  lblIntersectionTable.Enabled = Not (cboColumns(1) = vbNullString)
  cboParents(2).Enabled = Not (cboColumns(1) = vbNullString)
  lblIntersectionAggregate.Enabled = Not (cboColumns(1) = vbNullString)
  cboAggregateType(2).Enabled = Not (cboColumns(1) = vbNullString)
  lblIntersectionColumn.Enabled = Not (cboColumns(1) = vbNullString)
  cboColumns(2).Enabled = Not (cboColumns(1) = vbNullString)
  lblIntersectionSortOrder.Enabled = Not (cboColumns(1) = vbNullString)
  cboSortOrder(2).Enabled = Not (cboColumns(1) = vbNullString)
  
  lblAggregateType.Enabled = (cboColumns(1) = vbNullString)
  cboAggregateType(0).Enabled = (cboColumns(1) = vbNullString)
  
  

End Sub

Private Sub cboParents_Click(Index As Integer)
  
  On Error GoTo ErrorTrap

  If mfLoading Then Exit Sub

  mfChanged = True

  mlngTableID = cboParents(Index).ItemData(cboParents(Index).ListIndex)
  fOK = PopulateColumnsCombo(mlngTableID, Index, 0)
  
  ' Populate the 'Horizontal' combo
  If Index = 0 Then PopulateVerTableCombo (0)
  
  If Index < 2 Then
  ' populate the 'Intersection' combo
    If cboParents(0) = cboParents(1) Then
      ' Intersection combo is any parent or child of parent 1
      PopulateIntersectionCombo (0)
    Else
      ' Intersection combo is only either parent 1 or parent 2
      cboParents(2).Clear
      cboColumns(2).Clear
'      PopulateSortCombo (0)
      If cboParents(0) <> vbNullString And cboParents(1) <> vbNullString Then
        ' Add defaults first
        cboParents(2).AddItem ""
        cboParents(2).ItemData(cboParents(2).NewIndex) = 0
        cboParents(2).AddItem cboParents(0)
        cboParents(2).ItemData(cboParents(2).NewIndex) = cboParents(0).ItemData(cboParents(0).ListIndex)
        cboParents(2).AddItem cboParents(1)
        cboParents(2).ItemData(cboParents(2).NewIndex) = cboParents(1).ItemData(cboParents(1).ListIndex)
'        PopulateSortCombo (0)
      End If
    End If
  End If
  ' Check if the selected expression is for the current table.
  With recExprEdit
    .Index = "idxExprID"
    .Seek "=", txtFilter.Tag, False
    
    If Not .NoMatch Then
      If (!TableID <> mlngTableID) Then
        txtFilter.Tag = 0
        txtFilter.Text = ""
      End If
    Else
      txtFilter.Tag = 0
      txtFilter.Text = ""
      cmdFilterClear.Enabled = False
    End If
  End With
  
  cboSortOrder(Index).ListIndex = 0
  
  EnableDisableCombos
  
  ChartTableID = IIf(cboParents(0) <> vbNullString, cboParents(0).ItemData(cboParents(0).ListIndex), 0)
  Chart_TableID_2 = IIf(cboParents(1) <> vbNullString, cboParents(1).ItemData(cboParents(1).ListIndex), 0)
  Chart_TableID_3 = IIf(cboParents(2) <> vbNullString, cboParents(2).ItemData(cboParents(2).ListIndex), 0)
    
  ' Clear the sort order combo
      
  RefreshControls
  
ErrorTrap:
 
  
End Sub

Public Function PopulateSortCombos(plngSortOrderID As Long)
  ' cheeky
  Dim pstrBinaryString As String
  pstrBinaryString = DecToBin(plngSortOrderID, 3)
  
  For jnCount = 1 To 3
    If Mid(pstrBinaryString, jnCount, 1) = "0" Then
      ' Ascending
      cboSortOrder(jnCount - 1).ListIndex = IIf(cboParents(jnCount - 1) <> vbNullString, 1, -1)
    Else
      ' Descending
      cboSortOrder(jnCount - 1).ListIndex = IIf(cboParents(jnCount - 1) <> vbNullString, 2, -1)
    End If
  Next
      
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

  pstrBinaryString = IIf(cboSortOrder(0) = "Descending", "1", "0")
  pstrBinaryString = pstrBinaryString & IIf(cboSortOrder(1) = "Descending", "1", "0")
  pstrBinaryString = pstrBinaryString & IIf(cboSortOrder(2) = "Descending", "1", "0")
  
  Do
    BinaryToDec = BinaryToDec + (Left(pstrBinaryString, 1) * 2 ^ (Len(pstrBinaryString) - 1))
    pstrBinaryString = Mid(pstrBinaryString, 2)
  Loop Until pstrBinaryString = ""
  
  ConvertCombosToDecimal = BinaryToDec
End Function

Private Sub cboSortOrder_Click(Index As Integer)
  mfChanged = True
  Chart_SortOrderID = ConvertCombosToDecimal
  RefreshControls
End Sub

Private Sub cmdFilterClear_Click()
  txtFilter.Text = vbNullString
  txtFilter.Tag = 0
  mlngChartFilterID = 0
  cmdFilterClear.Enabled = False
  mfChanged = True
  RefreshControls
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
        cmdOK_Click
        Cancel = True
    End Select
  End If
  RefreshControls
End Sub

Private Sub RefreshControls()
  
  cmdOK.Enabled = mfChanged
  
End Sub


'Private Sub optAggregateType_Click(Index As Integer)
'
'  mfChanged = True
'
'  miChartAggregateType = IIf(optAggregateType(0).value = 1, 0, 1)
'End Sub

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
  mlngTableID = IIf(mlngTableID = 0, cboParents(Index).ItemData(cboParents(Index).ListIndex), mlngTableID)
  
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
      
      
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
        cmdFilterClear.Enabled = False
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
                        plngChart_TableID_2, _
                        plngChart_ColumnID_2, _
                        plngChart_TableID_3, _
                        plngChart_ColumnID_3, _
                        plngChart_SortOrderID, _
                        plngChart_SortDirection)


  '  ChartViewID = plngChartViewID
  ChartTableID = miChartTableID
  ChartColumnID = plngChartColumnID
  ChartAggregateType = piChartAggregateType
  ChartFilterID = plngChartFilterID
  Chart_TableID_2 = plngChart_TableID_2
  Chart_ColumnID_2 = plngChart_ColumnID_2
  Chart_TableID_3 = plngChart_TableID_3
  Chart_ColumnID_3 = plngChart_ColumnID_3
  Chart_SortOrderID = plngChart_SortOrderID
  Chart_SortDirection = plngChart_SortDirection
  
  RefreshControls
  
End Sub

Private Sub cmdOK_Click()
  Cancelled = False
  Me.Hide
End Sub

Private Sub Form_Load()
  mfLoading = True
  If Chart_ColumnID_3 > 0 Then
    fOK = PopulateAggregateCombo(miChartAggregateType, 2)
  Else
    fOK = PopulateAggregateCombo(miChartAggregateType, 0)
  End If
  
  PopulateParentsCombo (miChartTableID) ' populate and set default value - Horizontal
  fOK = PopulateColumnsCombo(miChartTableID, 0, ChartColumnID)
  PopulateVerTableCombo (Chart_TableID_2) ' populate and set default value - Vertical
  fOK = PopulateColumnsCombo(Chart_TableID_2, 1, Chart_ColumnID_2)
  PopulateIntersectionCombo (Chart_TableID_3)  ' populate and set default value - Intersection
  fOK = PopulateColumnsCombo(Chart_TableID_3, 2, Chart_ColumnID_3)
  
  txtFilter.Tag = mlngChartFilterID
  txtFilter.Text = GetExpressionName(txtFilter.Tag)
  If txtFilter.Text = "" Then
    cmdFilterClear.Enabled = False
  Else
    cmdFilterClear.Enabled = True
  End If
  txtFilter.Enabled = False
  txtFilter.BackColor = vbButtonFace
  
  For jnCount = 0 To 2
    cboSortOrder(jnCount).Clear
    cboSortOrder(jnCount).AddItem ""
    cboSortOrder(jnCount).AddItem "Ascending"
    cboSortOrder(jnCount).AddItem "Descending"
  Next
  
  PopulateSortCombos (Chart_SortOrderID)
  
  EnableDisableCombos
  
  mfLoading = False
  
  mfChanged = False
  
  RefreshControls
  
End Sub

Private Function PopulateParentsCombo(plngDefaultID As Long) As Boolean
  
  Dim i As Integer
  ' Clear the contents of the combo.
  cboParents(0).Clear

  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If !TableType <> iTabLookup And Not !Deleted Then
        cboParents(0).AddItem !TableName
        cboParents(0).ItemData(cboParents(0).NewIndex) = !TableID
      End If

      .MoveNext
    Loop
  End With

  ' Set the correct item as default
  If plngDefaultID = 0 Then
    cboParents(0).ListIndex = 0
  Else
    For i = 0 To cboParents(0).ListCount - 1
      If cboParents(0).ItemData(i) = plngDefaultID Then
        cboParents(0).ListIndex = i
        Exit For
      End If
    Next
  End If

End Function

Private Function PopulateColumnsCombo(plngTableID As Long, piIndex As Integer, plngDefaultID)

  Dim i As Integer
  
  ' Clear the contents of the combo
  cboColumns(piIndex).Clear

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
          recColEdit!DataType <> dtLONGVARBINARY Then
            cboColumns(piIndex).AddItem recColEdit.Fields("ColumnName")
            cboColumns(piIndex).ItemData(cboColumns(piIndex).NewIndex) = recColEdit.Fields("ColumnID")
        End If
      End If

      recColEdit.MoveNext
    Loop


    ' Set the correct item as default
    If plngDefaultID = 0 Then
      cboColumns(piIndex).ListIndex = 0
    Else
      For i = 0 To cboColumns(piIndex).ListCount - 1
        If cboColumns(piIndex).ItemData(i) = plngDefaultID Then
          cboColumns(piIndex).ListIndex = i
          Exit For
        End If
      Next
    End If
    
    If cboColumns(piIndex).ListIndex < 0 Then cboColumns(piIndex).ListIndex = 0
      
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function PopulateVerTableCombo(plngDefaultID As Long) As Boolean
  Dim iLoop As Integer
  Dim lngBaseTableID As Long
  
  On Error GoTo Error_Trap

  If cboParents(0).ItemData(cboParents(0).ListIndex) = 0 Then Exit Function ' no parent (Horizontal) table set so, nothing to do.

  ' Clear Table Combo
  cboParents(1).Clear

  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  lngBaseTableID = cboParents(0).ItemData(cboParents(0).ListIndex)
  
  ' Add 'empty' to the combo
  cboParents(1).AddItem ""
  cboParents(1).ItemData(cboParents(1).NewIndex) = 0
    
  ' Add the Horizontal table to the combo
  cboParents(1).AddItem cboParents(0)
  cboParents(1).ItemData(cboParents(1).NewIndex) = lngBaseTableID
  
  GetRelatedTables lngBaseTableID, "PARENT"
  GetRelatedTables lngBaseTableID, "CHILD"

  ' add related tables to the combo
  For iLoop = 1 To UBound(mavTables, 2)
      cboParents(1).AddItem mavTables(2, iLoop)
      cboParents(1).ItemData(cboParents(1).NewIndex) = mavTables(1, iLoop)
  Next iLoop
  
  
  ' Set the correct item as default
  If plngDefaultID = 0 Then
    cboParents(1).ListIndex = 0
  Else
    For i = 0 To cboParents(1).ListCount - 1
      If cboParents(1).ItemData(i) = plngDefaultID Then
        cboParents(1).ListIndex = i
        Exit For
      End If
    Next
  End If
    
  cboParents(1).Enabled = (cboParents(1).ListCount > 1)
  cboParents(1).BackColor = IIf(cboParents(1).Enabled, vbWindowBackground, vbButtonFace)
  
TidyUpAndExit:
  Exit Function

Error_Trap:
  MsgBox "Error populating Vertical tables dropdown box.", vbExclamation + vbOKOnly, "Chart Link"
  PopulateVerTableCombo = False
  GoTo TidyUpAndExit
  
End Function

Private Function PopulateIntersectionCombo(plngDefaultID As Long) As Boolean
  Dim iLoop As Integer
  
  Dim plngBaseTableID As Long
  
  plngBaseTableID = cboParents(0).ItemData(cboParents(0).ListIndex)
  
  On Error GoTo Error_Trap

  ' Clear Table Combo
  cboParents(2).Clear

  ' if either the Horizontal, or the Vertical combos are empty, clear the list and
  If cboParents(0) = vbNullString Or cboParents(1) = vbNullString Then Exit Function
  
  ' Get the tables related to the selected base table
  ' Put the table info into an array
  '   Column 1 = table ID
  '   Column 2 = table name
  '   Column 3 = true if this table is an ASCENDENT of the base table
  '            = false if this table is an DESCENDENT of the base table
  ReDim mavTables(3, 0)
  
  ' Add defaults first
  cboParents(2).AddItem ""
  cboParents(2).ItemData(cboParents(2).NewIndex) = 0
  
  ' Add parent 1 as intersection may be same table.
  cboParents(2).AddItem cboParents(0)
  cboParents(2).ItemData(cboParents(2).NewIndex) = plngBaseTableID
    
  GetRelatedTables plngBaseTableID, "PARENT"
  GetRelatedTables plngBaseTableID, "CHILD"

  For iLoop = 1 To UBound(mavTables, 2)
    'If Not AlreadyUsedInReport(CLng(mavTables(1, iLoop)), IIf(mfNew, 0, mlngRelatedTableID)) Then
      cboParents(2).AddItem mavTables(2, iLoop)
      cboParents(2).ItemData(cboParents(2).NewIndex) = mavTables(1, iLoop)
    'End If
  Next iLoop
 
  ' Set the correct item as default
  If plngDefaultID = 0 Then
    cboParents(2).ListIndex = 0
  Else
    For i = 0 To cboParents(2).ListCount - 1
      If cboParents(2).ItemData(i) = plngDefaultID Then
        cboParents(2).ListIndex = i
        Exit For
      End If
    Next
  End If
  
  cboParents(2).Enabled = (cboParents(2).ListCount > 1)
  cboParents(2).BackColor = IIf(cboParents(2).Enabled, vbWindowBackground, vbButtonFace)
  
TidyUpAndExit:
  Exit Function

Error_Trap:
  MsgBox "Error populating Intersection tables dropdown box.", vbExclamation + vbOKOnly, "Chart Link"
  PopulateIntersectionCombo = False
  GoTo TidyUpAndExit
  
End Function

Private Sub GetRelatedTables(plngTableID As Long, psRelationship As String)
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim iLoop As Integer
  Dim fFound As Boolean
  
'  ' If the table being cloned has any parent tables, then remember their IDs
'  ' and their column IDs in the clone register.
'  sSQL = "SELECT tmpRelations.parentID" & _
'    " FROM tmpRelations" & _
'    " WHERE tmpRelations.childID = " & Trim(Str(gLngTableID))
'  Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  
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



Private Function PopulateAggregateCombo(piAggregateType As Integer, Index As Integer) As Boolean
  
  Dim piColumnDataType As Integer
  
  If Index = 1 Then Exit Function ' no aggregates on the horizontal column

  cboAggregateType(Index).Clear
  cboAggregateType(Index).AddItem "Count"
  cboAggregateType(Index).ItemData(cboAggregateType(Index).NewIndex) = 0
    
  piColumnDataType = IIf(Index = 0, GetColumnDataType(mlngChartColumnID), GetColumnDataType(mlngChart_ColumnID_3))
  
  If piColumnDataType = dtinteger Or piColumnDataType = dtnumeric Then
    cboAggregateType(Index).AddItem "Total"
    cboAggregateType(Index).ItemData(cboAggregateType(Index).NewIndex) = 1
    cboAggregateType(Index).AddItem "Average"
    cboAggregateType(Index).ItemData(cboAggregateType(Index).NewIndex) = 2
    cboAggregateType(Index).AddItem "Minimum"
    cboAggregateType(Index).ItemData(cboAggregateType(Index).NewIndex) = 3
    cboAggregateType(Index).AddItem "Maximum"
    cboAggregateType(Index).ItemData(cboAggregateType(Index).NewIndex) = 4
  End If
  
  ' Set the correct item as default
  For i = 0 To cboAggregateType(Index).ListCount - 1
    If cboAggregateType(Index).ItemData(i) = ChartAggregateType Then
      cboAggregateType(Index).ListIndex = i
      Exit For
    End If
  Next

  If cboAggregateType(Index).ListIndex < 0 Then cboAggregateType(Index).ListIndex = 0
  
End Function

Public Property Get ChartTableID() As Long
  ChartTableID = miChartTableID
End Property

Public Property Let ChartTableID(ByVal plngNewValue As Long)
  miChartTableID = plngNewValue
End Property


Public Property Get ChartColumnID() As Long
  ChartColumnID = mlngChartColumnID
End Property

Public Property Let ChartColumnID(ByVal plngNewValue As Long)
  mlngChartColumnID = plngNewValue
End Property

Public Property Get ChartAggregateType() As Integer
  ChartAggregateType = miChartAggregateType
End Property

Public Property Let ChartAggregateType(ByVal piNewValue As Integer)
  miChartAggregateType = piNewValue
End Property

Public Property Get ChartFilterID() As Long
  ChartFilterID = mlngChartFilterID
End Property

Public Property Let ChartFilterID(ByVal plngNewValue As Long)
  mlngChartFilterID = plngNewValue
End Property

Public Property Get Chart_TableID_2() As Long
  Chart_TableID_2 = mlngChart_TableID_2
End Property

Public Property Let Chart_TableID_2(ByVal plngNewValue As Long)
  mlngChart_TableID_2 = plngNewValue
End Property

Public Property Get Chart_ColumnID_2() As Long
  Chart_ColumnID_2 = mlngChart_ColumnID_2
End Property

Public Property Let Chart_ColumnID_2(ByVal plngNewValue As Long)
  mlngChart_ColumnID_2 = plngNewValue
End Property

Public Property Get Chart_TableID_3() As Long
  Chart_TableID_3 = mlngChart_TableID_3
End Property

Public Property Let Chart_TableID_3(ByVal plngNewValue As Long)
  mlngChart_TableID_3 = plngNewValue
End Property

Public Property Get Chart_ColumnID_3() As Long
  Chart_ColumnID_3 = mlngChart_ColumnID_3
End Property

Public Property Let Chart_ColumnID_3(ByVal plngNewValue As Long)
  mlngChart_ColumnID_3 = plngNewValue
End Property

Public Property Get Chart_SortOrderID() As Long
  Chart_SortOrderID = mlngChart_SortOrderID
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


