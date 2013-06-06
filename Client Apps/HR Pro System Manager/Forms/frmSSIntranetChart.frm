VERSION 5.00
Begin VB.Form frmSSIntranetChart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Chart"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   Icon            =   "frmSSIntranetChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDrillDown 
      Caption         =   "HR Pro Report / Utility :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   150
      TabIndex        =   14
      Top             =   2310
      Width           =   5520
      Begin VB.ComboBox cboHRProUtilityType 
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
         Left            =   1380
         List            =   "frmSSIntranetChart.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   330
         Width           =   3930
      End
      Begin VB.ComboBox cboHRProUtility 
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
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   735
         Width           =   3930
      End
      Begin VB.Label lblHRProUtilityType 
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
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   390
         Width           =   645
      End
      Begin VB.Label lblHRProUtility 
         Caption         =   "Name :"
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
         Left            =   195
         TabIndex        =   17
         Top             =   795
         Width           =   780
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
      Left            =   3165
      TabIndex        =   12
      Top             =   3855
      Width           =   1200
   End
   Begin VB.Frame fraSSIChart 
      Caption         =   "Chart Data :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   150
      TabIndex        =   1
      Top             =   195
      Width           =   5520
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
         Left            =   4950
         MaskColor       =   &H000000FF&
         TabIndex        =   10
         ToolTipText     =   "Clear Path"
         Top             =   1095
         UseMaskColor    =   -1  'True
         Width           =   330
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
         Left            =   1380
         TabIndex        =   9
         Top             =   1110
         Width           =   3225
      End
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
         Left            =   4620
         TabIndex        =   8
         Top             =   1095
         Width           =   315
      End
      Begin VB.OptionButton optAggregateType 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3375
         TabIndex        =   7
         Top             =   1560
         Width           =   765
      End
      Begin VB.OptionButton optAggregateType 
         Caption         =   "Count"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2250
         TabIndex        =   6
         Top             =   1560
         Value           =   -1  'True
         Width           =   855
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
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   3930
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
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   3930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aggregate Function :"
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
         Left            =   195
         TabIndex        =   13
         Top             =   1590
         Width           =   1785
      End
      Begin VB.Label lblFilter 
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
         Left            =   195
         TabIndex        =   11
         Top             =   1170
         Width           =   615
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
         Left            =   195
         TabIndex        =   5
         Top             =   735
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
         Left            =   195
         TabIndex        =   4
         Top             =   360
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
      Left            =   4455
      TabIndex        =   0
      Top             =   3855
      Width           =   1200
   End
End
Attribute VB_Name = "frmSSIntranetChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private mblnCancelled As Boolean
'
'' Component definition variables.
'Private mobjComponent As CExprComponent
'Private miComponentType As ExpressionComponentTypes
'Private mavColumns() As Variant
'Private mDataType As Long
'Private mlngTableID As Long
'Private mvar_lngPersonnelTableID As Long
'
'
'Public Property Let Cancelled(ByVal bCancel As Boolean)
'  mblnCancelled = bCancel
'End Property
'
'Public Property Get Cancelled() As Boolean
'  Cancelled = mblnCancelled
'End Property
'
'Private Sub cboParents_Click()
'  mlngTableID = cboParents.ItemData(cboParents.ListIndex)
'  PopulateColumnsCombo (mlngTableID)
'End Sub
'
'Private Sub cmdCancel_Click()
'  Cancelled = True
'  UnLoad Me
'End Sub
'
'Private Sub cmdFilter_Click()
'
'  ' Display the 'Where Clause' expression selection form.
'  'On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim objExpr As CExpression
'
'  fOK = True
'
'  ' Instantiate an expression object.
'  Set objExpr = New CExpression
'
'  With objExpr
'    ' Set the properties of the expression object.
'    .Initialise mlngTableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC
'
'    ' Instruct the expression object to display the
'    ' expression selection form.
'    If .SelectExpression Then
'      txtFilter.Tag = .ExpressionID
'      txtFilter.Text = GetExpressionName(txtFilter.Tag)
'    Else
'      ' Check in case the original expression has been deleted.
'      txtFilter.Text = GetExpressionName(txtFilter.Tag)
'      If txtFilter.Text = vbNullString Then
'        txtFilter.Tag = 0
'      End If
'    End If
'
'  End With
'
'
'TidyUpAndExit:
'  Set objExpr = Nothing
'  If Not fOK Then
'    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
'  End If
'  Exit Sub
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'
'
'
'End Sub
'
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
'  txtFldSelFilter.Text = sExprName
'
'  GetFldSelFilterDetails = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
'
'
'Public Sub Initialize(plngChartViewID As Long, _
'                        plngChartTableID As Long, _
'                        plngChartColumnID As Long, _
'                        plngChartFilterID As Long, _
'                        plngChartAggregateType As Integer)
'
'
'  ChartViewID = plngChartViewID
'  ChartTableID = plngChartTableID
'
'End Sub
'
'Private Sub cmdOk_Click()
'  Cancelled = True
'  Me.Hide
'End Sub
'
'Private Sub Form_Load()
'  PopulateParentsCombo
'  PopulateColumnsCombo (cboParents.ItemData(cboParents.ListIndex))
'  PopulateControls
'End Sub
'
'Private Sub PopulateControls()
'  txtFilter.Tag = 0
'  txtFilter.Enabled = False
'  txtFilter.BackColor = vbButtonFace
'End Sub
'
'
'Private Sub PopulateParentsCombo()
'  ' Clear the contents of the combo.
'  cboParents.Clear
'
'  With recTabEdit
'    .Index = "idxName"
'
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'    End If
'
'    Do While Not .EOF
'      If !TableType <> iTabLookup And Not !Deleted Then
'        cboParents.AddItem !TableName
'        cboParents.ItemData(cboParents.NewIndex) = !TableID
'      End If
'
'      .MoveNext
'    Loop
'  End With
'
''
''  ' Add all the views (and viewIDs) configured for SSI...
''  With frmSSIntranetSetup
''    For iLoop = 0 To .cboButtonLinkView.ListCount - 1
''        cboParents.AddItem .cboButtonLinkView.List(iLoop)
''        cboParents.ItemData(cboParents.NewIndex) = .cboButtonLinkView.ItemData(iLoop)
''    Next iLoop
''  End With
''
''  ' remove personnel_records as a table...
''  mvar_lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
''
''  For iLoop = cboParents.ListCount - 1 To 0 Step -1
''    If cboParents.ItemData(iLoop) = mvar_lngPersonnelTableID Then
''      cboParents.RemoveItem (iLoop)
''    End If
''  Next iLoop
'
'  cboParents.ListIndex = 0
'
'End Sub
'
'Private Function PopulateColumnsCombo(plngTableID As Long) As Boolean
'
'  ' Clear the contents of the combo
'  cboColumns.Clear
'
'  ' Add the table's columns to the view definition in the local database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'
'  recColEdit.Index = "idxTableID"
'  recColEdit.Seek "=", plngTableID
'
'  fOK = Not recColEdit.NoMatch
'
'  If fOK Then
'
'    Do While Not recColEdit.EOF
'
'      ' If no more columns for this table exit loop
'      If recColEdit!TableID <> plngTableID Then
'        Exit Do
'      End If
'
'      ' Don't add deleted or system columns
'      If recColEdit!Deleted <> True And recColEdit!columnType <> giCOLUMNTYPE_SYSTEM Then
'        ' Add the column to the combo
'        cboColumns.AddItem recColEdit.Fields("ColumnName")
'      End If
'
'      recColEdit.MoveNext
'    Loop
'
'
'    cboColumns.ListIndex = 0
'
'  End If
'
'TidyUpAndExit:
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function
'
''Public Property Let ChartViewID(ByVal psNewValue As String)
''  'plngChartViewID = ChartViewID
''End Property
''
''Public Property Get ChartViewID()
''  'ChartViewID = plngChartViewID
''End Property
''
''Public Property Let ChartTableID(ByVal psNewValue As String)
''  'plngChartTableID = ChartTableID
''End Property
''
''Public Property Get ChartTableID()
''  'ChartTableID = cboParents.ItemData(cboParents.ListIndex)
''End Property
''
'
'
'
