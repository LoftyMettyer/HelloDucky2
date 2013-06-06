VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmExportColumns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Column"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1037
   Icon            =   "frmExportColumn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraProperties 
      Height          =   2415
      Left            =   1850
      TabIndex        =   20
      Top             =   1360
      Width           =   4800
      Begin VB.CheckBox chkSuppressNulls 
         Caption         =   "E&xclude if empty"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   2360
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.ComboBox cboConvCase 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   315
         ItemData        =   "frmExportColumn.frx":000C
         Left            =   1635
         List            =   "frmExportColumn.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1500
         Width           =   3000
      End
      Begin VB.TextBox txtHeading 
         Height          =   315
         Left            =   1640
         MaxLength       =   50
         TabIndex        =   22
         Top             =   300
         Width           =   3000
      End
      Begin VB.TextBox txtCMGCode 
         Height          =   315
         Left            =   1640
         TabIndex        =   29
         Top             =   1900
         Visible         =   0   'False
         Width           =   3000
      End
      Begin COASpinner.COA_Spinner txtLength 
         Height          =   300
         Left            =   1635
         TabIndex        =   24
         Top             =   705
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         MaximumValue    =   0
         Text            =   "0"
      End
      Begin COASpinner.COA_Spinner spnDec 
         Height          =   300
         Left            =   1635
         TabIndex        =   26
         Top             =   1080
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         MaximumValue    =   6
         Text            =   "0"
      End
      Begin VB.Label lblConvCase 
         AutoSize        =   -1  'True
         Caption         =   "Convert Case :"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Heading :"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lblCMGCode 
         AutoSize        =   -1  'True
         Caption         =   "CMG Code :"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   1965
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblProp_Decimals 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decimals :"
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   1155
         Width           =   1125
      End
      Begin VB.Label lblLength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size :"
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lblSizeIncreased 
         BackStyle       =   0  'Transparent
         Caption         =   "Size includes one character for decimal point"
         Height          =   720
         Left            =   195
         TabIndex        =   33
         Top             =   1965
         Visible         =   0   'False
         Width           =   3945
      End
   End
   Begin VB.Frame fraType 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1620
      Begin VB.OptionButton optCalculation 
         Caption         =   "C&alculation"
         Height          =   195
         Left            =   200
         TabIndex        =   2
         Top             =   800
         Width           =   1275
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Ot&her"
         Height          =   195
         Left            =   200
         TabIndex        =   4
         Top             =   1600
         Width           =   900
      End
      Begin VB.OptionButton optTable 
         Caption         =   "&Field"
         Height          =   195
         Left            =   200
         TabIndex        =   1
         Top             =   400
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optText 
         Caption         =   "&Value"
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5445
      TabIndex        =   32
      Top             =   3870
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   4185
      TabIndex        =   31
      Top             =   3870
      Width           =   1200
   End
   Begin VB.Frame fraField 
      Height          =   1215
      Left            =   1850
      TabIndex        =   15
      Top             =   120
      Width           =   4800
      Begin VB.ComboBox cboFromColumn 
         Height          =   315
         Left            =   1640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   700
         Width           =   2985
      End
      Begin VB.ComboBox cboFromTable 
         Height          =   315
         Left            =   1640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   2985
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Index           =   1
         Left            =   200
         TabIndex        =   18
         Top             =   760
         Width           =   630
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Index           =   0
         Left            =   200
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraFreeText 
      Height          =   1215
      Left            =   1850
      TabIndex        =   12
      Top             =   120
      Width           =   4800
      Begin VB.TextBox txtOther 
         Height          =   315
         Left            =   1640
         MaxLength       =   256
         TabIndex        =   14
         Top             =   340
         Width           =   2985
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free Text :"
         Height          =   195
         Index           =   3
         Left            =   200
         TabIndex        =   13
         Top             =   400
         Width           =   810
      End
   End
   Begin VB.Frame fraOther 
      Height          =   1215
      Left            =   1850
      TabIndex        =   9
      Top             =   120
      Width           =   4800
      Begin VB.ComboBox cboOther 
         Height          =   315
         ItemData        =   "frmExportColumn.frx":0042
         Left            =   1640
         List            =   "frmExportColumn.frx":0044
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   340
         Width           =   3030
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other :"
         Height          =   195
         Index           =   4
         Left            =   200
         TabIndex        =   10
         Top             =   400
         Width           =   525
      End
   End
   Begin VB.Frame fraCalculation 
      Height          =   1215
      Left            =   1850
      TabIndex        =   5
      Top             =   120
      Width           =   4800
      Begin VB.CommandButton cmdCalculation 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   340
         Width           =   330
      End
      Begin VB.TextBox txtCalculation 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1640
         MaxLength       =   256
         TabIndex        =   7
         Top             =   340
         Width           =   2700
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculation :"
         Height          =   195
         Index           =   2
         Left            =   200
         TabIndex        =   6
         Top             =   400
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmExportColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbCancelled As Boolean                      'Has the operation been cancelled
Private datData As HRProDataMgr.clsDataAccess              'Data Access Class
Private mfrmForm As frmExport
Private mblnNew As Boolean
Private mblnAddAll As Boolean
Private mlngEditingColumnID As Long

Private Function RefreshControls() As Boolean
  
  Dim objExpr As New clsExprExpression
  Dim iTemp As SQLDataType
  Dim bShowDecimals As Boolean
  Dim bEnableLength As Boolean
  Dim bShowConvCase As Boolean
  
  Screen.MousePointer = vbHourglass

  If mblnAddAll Then
    Exit Function
  End If

  If optTable.Value Then

    If cboFromTable.ListCount = 1 Then
      cboFromTable.Enabled = False
      cboFromTable.BackColor = vbButtonFace
      cboFromTable.ListIndex = 0
    Else
      cboFromTable.Enabled = True
      cboFromTable.BackColor = vbWindowBackground
    End If
    
    If cboFromColumn.ListCount > 0 Then

      iTemp = datGeneral.GetDataType(cboFromTable.ItemData(cboFromTable.ListIndex), _
                                      cboFromColumn.ItemData(cboFromColumn.ListIndex))
      'NHRD03062003 Fault 4908
      If cboFromColumn.ListCount < 2 Then
        cboFromColumn.Enabled = False
        cboFromColumn.BackColor = vbButtonFace
      Else
        cboFromColumn.Enabled = True
        cboFromColumn.BackColor = vbWindowBackground
      End If
      
      bShowDecimals = (iTemp = sqlNumeric)
      
      'NPG20071212 Fault 12867
      cboConvCase.Enabled = (iTemp = sqlVarChar Or iTemp = sqlBoolean)
      cboConvCase.BackColor = IIf((iTemp = sqlVarChar Or iTemp = sqlBoolean), vbWindowBackground, vbButtonFace)
      cboConvCase.ForeColor = IIf((iTemp = sqlVarChar Or iTemp = sqlBoolean), vbWindowText, vbGrayText)

    Else
      cboFromColumn.Enabled = False
      cboFromColumn.BackColor = vbButtonFace
      bShowDecimals = False
    End If

    bEnableLength = True
    txtHeading.Text = cboFromTable.Text & "." & cboFromColumn.Text

  ElseIf optCalculation.Value Then

    If Val(txtCalculation.Tag) > 0 Then
      objExpr.ExpressionID = txtCalculation.Tag
      objExpr.ConstructExpression
      
      'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
      ' when all we need is the expression return type (time saving measure).
      objExpr.ValidateExpression True, True
      bShowDecimals = (objExpr.ReturnType = giEXPRVALUE_NUMERIC Or _
                          objExpr.ReturnType = giEXPRVALUE_BYREF_NUMERIC)
      bShowConvCase = (objExpr.ReturnType = giEXPRVALUE_CHARACTER Or _
                          objExpr.ReturnType = giEXPRVALUE_LOGIC)
                          
      Set objExpr = Nothing
    Else
      txtCalculation.Text = ""
      txtCalculation.Tag = 0
      txtLength.Value = 0
      spnDec.Value = 0
      bShowDecimals = False
    End If

    cmdCalculation.Enabled = True
    bEnableLength = True
      
    'NPG20071212 Fault 12867
    cboConvCase.Enabled = (bShowConvCase)
    cboConvCase.BackColor = IIf((bShowConvCase), vbWindowBackground, vbButtonFace)
    cboConvCase.ForeColor = IIf((bShowConvCase), vbWindowText, vbGrayText)

  ElseIf optText.Value Then
    'txtOther.SetFocus
    bEnableLength = False
    bShowDecimals = False
    txtLength.Value = Len(txtOther.Text)
    spnDec.Value = 0

    'NPG20080718 Fault 13276
    cboConvCase.Enabled = False
    cboConvCase.BackColor = vbButtonFace
    cboConvCase.ForeColor = vbGrayText

  ElseIf optOther.Value Then
    Select Case cboOther.ListIndex
    Case -1:
      cboOther.ListIndex = 0
      txtLength.Text = 1
      bEnableLength = True
    Case 0:
      txtLength.Text = 1
      bEnableLength = True
    Case 1
      txtLength.Text = 0
      bEnableLength = False
    Case 2
      'txtLength.Text = 0
      bEnableLength = True
    End Select

    txtHeading.Text = cboOther.Text

    'NPG20080718 Fault 13276
    cboConvCase.Enabled = False
    cboConvCase.BackColor = vbButtonFace
    cboConvCase.ForeColor = vbGrayText

  End If
  
  txtLength.Enabled = bEnableLength
  txtLength.BackColor = IIf(bEnableLength, vbWindowBackground, vbButtonFace)
  lblLength.Enabled = bEnableLength
  
  spnDec.Enabled = bShowDecimals
  spnDec.Enabled = bShowDecimals
  spnDec.BackColor = IIf(bShowDecimals, vbWindowBackground, vbButtonFace)
  lblProp_Decimals.Enabled = bShowDecimals

  fraField.Visible = (optTable.Value = True)
  fraCalculation.Visible = (optCalculation.Value = True)
  fraFreeText.Visible = (optText.Value = True)
  fraOther.Visible = (optOther.Value = True)

  lblConvCase.Enabled = cboConvCase.Enabled

  Screen.MousePointer = vbDefault

End Function

'MH20020712 Fault 4122
'Public Sub Initialise(bNew As Boolean, sType As String, lFromTableID As Integer, lFromColumnID As Long, _
  Optional sFromTable As String, Optional sFromColumn As String, Optional iLength As Integer, _
  Optional frmParentForm As frmExport, Optional iDecimals As Long)
'NPG20071217 Fault 12867
' Public Sub Initialise(bNew As Boolean, sType As String, lFromTableID As Long, lFromColumnID As Long, _
  Optional sFromTable As String, Optional sFromColumn As String, Optional iLength As Long, _
  Optional frmParentForm As frmExport, Optional iDecimals As Long, Optional strHeading As String, _
  Optional bAddAll As Boolean)


Public Sub Initialise(bNew As Boolean, sType As String, lFromTableID As Long, lFromColumnID As Long, _
  Optional sFromTable As String, Optional sFromColumn As String, Optional iLength As Long, _
  Optional frmParentForm As frmExport, Optional iDecimals As Long, Optional strHeading As String, _
  Optional sConvertCase As String, Optional bSuppressNulls As Boolean, Optional bAddAll As Boolean)

  Set datData = New HRProDataMgr.clsDataAccess
  Set mfrmForm = frmParentForm
  
  mblnNew = bNew
  mblnAddAll = bAddAll
  mlngEditingColumnID = lFromColumnID
  
  With cboFromTable
        
    .Clear
        
    'If selected, add the base table
    If frmParentForm.cboBaseTable.Text <> "<None>" Then
      .AddItem frmParentForm.cboBaseTable.Text
      .ItemData(.NewIndex) = frmParentForm.cboBaseTable.ItemData(frmParentForm.cboBaseTable.ListIndex)
    End If
    
    'If selected, add the parent1 table
    If frmParentForm.txtParent1.Text <> "" Then   ' <> "None"
      .AddItem frmParentForm.txtParent1.Text
      .ItemData(.NewIndex) = frmParentForm.txtParent1.Tag
    End If
        
    'If selected, add the parent2 table
    If frmParentForm.txtParent2.Text <> "" Then   ' <> "None"
      .AddItem frmParentForm.txtParent2.Text
      .ItemData(.NewIndex) = frmParentForm.txtParent2.Tag
    End If
        
    'If selected, add the child table
    If frmParentForm.cboChild.Text <> "<None>" Then
      .AddItem frmParentForm.cboChild.Text
      .ItemData(.NewIndex) = frmParentForm.cboChild.ItemData(frmParentForm.cboChild.ListIndex)
    End If

    If bAddAll Then
      ControlsDisableAll fraType
      optTable.Value = True
      ControlsDisableAll fraProperties
      lblTitle(0).Enabled = (.ListCount > 1)
    End If

    If .ListCount >= 0 Then
      SetComboText cboFromTable, mfrmForm.cboBaseTable.Text
      '.ListIndex = 0
    Else
      .ListIndex = -1
    End If
    
    If .ListCount = 1 Then
      .BackColor = vbButtonFace
      .Enabled = False
    Else
      .BackColor = vbWindowBackground
      .Enabled = True
     End If
     
  End With


  With cboOther
    .Clear
    .AddItem "Filler"
    .AddItem "Carriage Return"
    .AddItem "Record Number"
  End With


  'If we are editing an existing grid entry, display the data
  If Not bNew Then
      
    Select Case sType
    
      Case "C"
        optTable.Value = True
        SetComboText cboFromTable, sFromTable
        SetComboText cboFromColumn, sFromColumn
        
        'NPG20071214 Fault 12867 - set case conversion for character fields
        cboConvCase.ListIndex = sConvertCase
        
        'NPG20080617 Suggestion 000816
        chkSuppressNulls.Value = IIf(bSuppressNulls, 1, 0)
    
      Case "X"
        optCalculation.Value = True
        txtCalculation.Tag = lFromColumnID
        txtCalculation.Text = sFromColumn

        'NPG20071214 Fault 12867 - set case conversion for character fields
        cboConvCase.ListIndex = sConvertCase
        
        'NPG20080617 Suggestion 000816
        chkSuppressNulls.Value = IIf(bSuppressNulls, 1, 0)

      Case "T"
        optText.Value = True
        txtOther.Text = sFromTable
      
      Case "F"  'Filler
        optOther.Value = True
        cboOther.ListIndex = 0
      
      Case "R"  'Return
        optOther.Value = True
        cboOther.ListIndex = 1

      Case "N"  'Number
        optOther.Value = True
        cboOther.ListIndex = 2

    End Select

    RefreshControls
    If iLength <> 0 Then txtLength.Text = iLength Else txtLength.Text = 0
    If iDecimals <> 0 Then spnDec.Text = iDecimals Else spnDec.Text = 0

  Else
    RefreshControls
      Dim lngColumnID As Long

    lngColumnID = 0
    If cboFromColumn.ListIndex >= 0 Then
      lngColumnID = cboFromColumn.ItemData(cboFromColumn.ListIndex)
    End If

    If lngColumnID > 0 Then
      'MH20070226 Fault 11953
      'txtLength.Text = GetColumnSize(lngColumnID)
      'spnDec.Text = GetColumnDecimals(lngColumnID)
      spnDec.Text = GetColumnDecimals(lngColumnID)
      txtLength.Text = GetColumnSize(lngColumnID) + IIf(spnDec.Value > 0, 1, 0)
    Else
      txtLength.Text = 0 'vbNullString
      spnDec.Text = vbNullString
    End If

  End If

  If Not mblnNew Then
    txtHeading.Text = strHeading
  End If
  
  Screen.MousePointer = vbDefault
            
End Sub

Private Sub cboFromColumn_Click()
  
''  'If no column selected, exit
''  If cboFromColumn.Text = "" Then Exit Sub
''
''  'Display default column size in fixed length field
''  txtLength.Text = GetColumnSize(cboFromColumn.ItemData(cboFromColumn.ListIndex))
''  'TM20011024 Fault 3019
''  spnDec.Text = GetColumnDecimals(cboFromColumn.ItemData(cboFromColumn.ListIndex))
''
''  'If column is date, suggest 12 as length for fixed length (xx/xx/xxxx + 2 spaces)
''  'If txtLength.Text = "0" Or txtLength.Text = "1" Then txtLength = "12"


  Dim lngColumnID As Long

  lngColumnID = 0
  If cboFromColumn.ListIndex >= 0 Then
    lngColumnID = cboFromColumn.ItemData(cboFromColumn.ListIndex)
  End If

  If lngColumnID > 0 Then
    'MH20070226 Fault 11953
    'txtLength.Text = GetColumnSize(lngColumnID)
    'spnDec.Text = GetColumnDecimals(lngColumnID)
    spnDec.Text = GetColumnDecimals(lngColumnID)
    txtLength.Text = GetColumnSize(lngColumnID) + IIf(spnDec.Value > 0, 1, 0)
    lblSizeIncreased.Visible = (spnDec.Value > 0)
  Else
    txtLength.Text = 0 'vbNullString
    spnDec.Text = vbNullString
  End If

  RefreshControls

End Sub

Private Sub cboFromTable_Click()
    
  'If no table selected, wipe column listbox and exit
  If cboFromTable.Text = "" Then
    cboFromColumn.Clear
    Exit Sub
  End If
  
  Dim sSQL As String
  Dim rsCols As Recordset

  'Get all the columns for the selected table
  Set rsCols = GetColumnDetails(cboFromTable.ItemData(cboFromTable.ListIndex))
   
  With cboFromColumn
    .Clear

    If mblnAddAll Then
      lblTitle(1).Enabled = False
      .Enabled = False
      .BackColor = vbButtonFace
      .Clear
      .AddItem "<All>"
    End If

    Do While Not rsCols.EOF
    
      'MH20030122
      'Don't need to limit a column to be used only once...
      'mfrmForm.grdColumns.Redraw = False
      'If AlreadyUsedInExport(rsCols!ColumnID, IIf(mblnNew = False, mlngEditingColumnID, 0)) = False Then
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
      'End If
      'mfrmForm.grdColumns.Redraw = True
      
      rsCols.MoveNext
    Loop
    
    If .ListCount > 0 Then .ListIndex = 0 Else .ListIndex = -1
  
  End With
    
  rsCols.Close
  Set rsCols = Nothing
  
  If Not mblnAddAll Then
    ' JPD20011107 Fault 3122
    If cboFromColumn.ListCount > 0 Then
    Else
      cboFromColumn.Enabled = False
      cboFromColumn.BackColor = vbButtonFace
    End If
  
    'MH20030120
    txtHeading.Text = cboFromTable.Text & "." & cboFromColumn.Text
  End If

  Screen.MousePointer = vbDefault

End Sub

'Private Function AlreadyUsedInExport(plngColExprID As Long, Optional plngExclusion As Long) As Boolean
'
'  Dim pintOldPosition As Integer
'  Dim pvarBookmark As Variant
'  Dim pintLoop As Integer
'
'  With mfrmForm.grdColumns
'
'    ' Loop thru the export grid, adding data to the combo if they are columns
'    .MoveFirst
'      Do Until pintLoop = .Rows
'        pvarBookmark = .GetBookmark(pintLoop)
'        If .Columns("ColExprID").CellText(pvarBookmark) = plngColExprID Then
'          If plngExclusion = 0 Then
'            AlreadyUsedInExport = True
'              ' RH Fault 2053 - Wrongly states column used in sort order, so reset current row
'              With mfrmForm.grdColumns
'                ' Loop thru the export grid, adding data to the combo if they are columns
'                .MoveFirst
'                  Do Until pintLoop = .Rows
'                    pvarBookmark = .GetBookmark(pintLoop)
'                    If .Columns("ColExprID").CellText(pvarBookmark) = mlngEditingColumnID Then
'                        Exit Function
'                    End If
'                    pintLoop = pintLoop + 1
'                  Loop
'              End With
'            Exit Function
'          Else
'            If .Columns("ColExprID").CellText(pvarBookmark) = plngExclusion Then
'              AlreadyUsedInExport = False
'            Else
'              AlreadyUsedInExport = True
'                ' RH Fault 2053 - Wrongly states column used in sort order, so reset current row
'                With mfrmForm.grdColumns
'                  ' Loop thru the export grid, adding data to the combo if they are columns
'                  .MoveFirst
'                    Do Until pintLoop = .Rows
'                      pvarBookmark = .GetBookmark(pintLoop)
'                      If .Columns("ColExprID").CellText(pvarBookmark) = mlngEditingColumnID Then
'                          Exit Function
'                      End If
'                      pintLoop = pintLoop + 1
'                    Loop
'                End With
'              Exit Function
'            End If
'          End If
'        End If
'        pintLoop = pintLoop + 1
'      Loop
'
'  End With
'
'  ' RH Fault 2053 - Wrongly states column used in sort order, so reset current row
'  With mfrmForm.grdColumns
'    ' Loop thru the export grid, adding data to the combo if they are columns
'    .MoveFirst
'      Do Until pintLoop = .Rows
'        pvarBookmark = .GetBookmark(pintLoop)
'        If .Columns("ColExprID").CellText(pvarBookmark) = mlngEditingColumnID Then
'            Exit Function
'        End If
'        pintLoop = pintLoop + 1
'      Loop
'  End With
'
'  AlreadyUsedInExport = False
'
'End Function

Private Sub cboOther_Click()
'  With txtLength
'    If cboOther.ListIndex = 0 Then
'      .Enabled = True
'      .Text = "1"
'      .BackColor = &H80000005
'    Else
'      .Enabled = False
'      .Text = "0"
'      .BackColor = &H8000000F
'    End If
'  End With
'
  RefreshControls

End Sub

Private Sub cmdCalculation_Click()

  Dim objExpr As New clsExprExpression
  Set objExpr = New clsExprExpression

  With objExpr
    If .Initialise(mfrmForm.cboBaseTable.ItemData(mfrmForm.cboBaseTable.ListIndex), Me.txtCalculation.Tag, giEXPR_RUNTIMECALCULATION, 0) Then
      .SelectExpression True
    End If
'  End With

    If .ExpressionID > 0 Then
      txtCalculation.Text = .Name
      txtCalculation.Tag = .ExpressionID
      txtHeading.Text = .Name
    Else
      txtCalculation.Text = vbNullString
      txtCalculation.Tag = 0
      txtHeading.Text = vbNullString
    End If
    
  End With
  
  Set objExpr = Nothing

  RefreshControls
  
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  Unload Me

End Sub

Public Property Get Cancelled() As Boolean

    Cancelled = mbCancelled

End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)

    mbCancelled = bCancel

End Property

Private Sub cmdOK_Click()

  Dim prstTemp As Recordset
  Dim strErrorMsg As String
  Dim blnFixedLength As Boolean
  Dim blnCarriageReturn As Boolean
  
  blnFixedLength = (mfrmForm.optOutputFormat(fmtFixedLengthFile).Value = True)
  
  'Do some validation to make sure everything required has been entered
  If optTable Then
    If cboFromTable.Text = "" Or cboFromColumn.Text = "" Then
      COAMsgBox "You must select a table and column.", vbExclamation, Me.Caption
      Exit Sub
    End If
  End If
  
  If optCalculation Then
    If txtCalculation.Tag = 0 Then
      COAMsgBox "You must select a calculation.", vbExclamation, Me.Caption
      Exit Sub
    End If

    'MH20001102 Fault 1250
    'Check can still see calculation
    strErrorMsg = IsCalcValid(txtCalculation.Tag)
    If strErrorMsg <> vbNullString Then
      COAMsgBox strErrorMsg, vbExclamation, App.Title
      txtCalculation.Text = vbNullString
      txtCalculation.Tag = 0
      Exit Sub
    End If


    ' If we are not the creator of the definition and this calc is hidden, then
    ' they cant add it to the definition.
    Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & Me.txtCalculation.Tag)
    If prstTemp.Fields("Access") = "HD" And Not mfrmForm.mblnDefinitionCreator Then
      COAMsgBox "Cannot include the '" & Me.txtCalculation.Text & "' calculation." & vbCrLf & _
            " Its hidden and you are not the creator of this definition.", vbInformation + vbOKOnly, "Export"
      Exit Sub
    End If
    Set prstTemp = Nothing
  End If

  If optText Then
    txtLength.Text = Len(txtOther.Text)
    If txtOther.Text = "" Then
      COAMsgBox "You must enter some free text.", vbExclamation, Me.Caption
      Exit Sub
    End If
  End If

  If Not mblnAddAll Then
    If blnFixedLength Then
      If txtLength.Text = 0 Then
        blnCarriageReturn = (optOther.Value = True And cboOther.Text = "Carriage Return")
        If Not blnCarriageReturn Then
          If COAMsgBox("Leaving a size of zero will result in the values not appearing." & vbCrLf & vbCrLf & _
                    "Do you wish to continue ?", vbQuestion + vbYesNo, "Export") = vbNo Then
            Exit Sub
          End If
        End If
      End If
    End If
  
  
    If Trim(txtHeading.Text) = vbNullString Then
      COAMsgBox "You must give this column a heading.", vbExclamation, Me.Caption
      Exit Sub
    End If
  End If

  Cancelled = False
  Me.Hide
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = 192 Then
    KeyCode = 0
  End If

End Sub

Private Sub Form_Load()
  txtLength.MaximumValue = VARCHAR_MAX_Size
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancelled = True
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set datData = Nothing

End Sub


'Private Sub OptFiller_Click()
'
'  txtCalculation.Text = ""
'  txtCalculation.Tag = 0
'  txtCalculation.Enabled = False
'
'  txtOther.Text = ""
'  txtOther.Enabled = False
'  'txtLength.BackColor = &H8000000F
'  cboFromTable.ListIndex = -1
'  cboFromTable.Enabled = False
'  cboFromColumn.ListIndex = -1
'  cboFromColumn.Enabled = False
'  cboFromTable.BackColor = &H8000000F
'  cboFromColumn.BackColor = &H8000000F
'
'  With txtLength
'    .Enabled = True
'    .Text = "1"
'    .BackColor = &H80000005
'  End With
'
'End Sub

Private Sub optCalculation_Click()

  Dim iTemp As SQLDataType
  Dim objExpr As clsExprExpression
  
'  txtOther.Text = ""
'  txtOther.Enabled = False
'  txtOther.BackColor = &H8000000F
'
'  txtCalculation.Text = ""
  txtCalculation.Tag = 0
  txtHeading.Text = vbNullString
'  cmdCalculation.Enabled = True
'
'  cboFromTable.ListIndex = -1
'  cboFromTable.Enabled = False
'  cboFromColumn.ListIndex = -1
'  cboFromColumn.Enabled = False
'  cboFromTable.BackColor = &H8000000F
'  cboFromColumn.BackColor = &H8000000F
'
'  cboOther.Enabled = False
'  cboOther.BackColor = &H8000000F
'  cboOther.ListIndex = -1
'
'  With txtLength
'    .BackColor = vbWindowBackground
'    .Enabled = True
'    .Text = "0"
'  End With
  
  RefreshControls
  
End Sub

Private Sub optOther_Click()

'  cboOther.Enabled = True
'  cboOther.BackColor = &H80000005
  cboOther.ListIndex = 0
  txtHeading.Text = vbNullString
'
'  txtOther.Text = ""
'  txtOther.Enabled = False
'  txtOther.BackColor = &H8000000F
'
'  txtCalculation.Text = ""
'  txtCalculation.Tag = 0
'  cmdCalculation.Enabled = False
'
'  cboFromTable.ListIndex = -1
'  cboFromTable.Enabled = False
'  cboFromColumn.ListIndex = -1
'  cboFromColumn.Enabled = False
'  cboFromTable.BackColor = &H8000000F
'  cboFromColumn.BackColor = &H8000000F

  RefreshControls
  
End Sub

Private Sub optTable_Click()

'  txtOther.Text = ""
'  txtOther.Enabled = False
'  txtOther.BackColor = &H8000000F
'
'  txtCalculation.Text = ""
'  txtCalculation.Tag = 0
'  cmdCalculation.Enabled = False
'
'  cboFromColumn.Enabled = True
'  cboFromColumn.BackColor = &H80000005
'
'  cboOther.Enabled = False
'  cboOther.BackColor = &H8000000F
'  cboOther.ListIndex = -1
'
  With cboFromTable
    .ListIndex = -1
    txtHeading.Text = vbNullString
    If .ListCount <> 1 Then
'      .Enabled = False
'      .ListIndex = 0
'      .BackColor = &H8000000F
'    Else
      SetComboText cboFromTable, mfrmForm.cboBaseTable.Text
''      .ListIndex = 0
'      .Enabled = True
'      .BackColor = &H80000005
    End If
  End With
'
'  With txtLength
'    .BackColor = &H80000005
'    .Enabled = True
'    '.Text = ""
'  End With

  RefreshControls
  
End Sub

Private Sub optText_Click()

'  cboFromTable.ListIndex = -1
'  cboFromTable.Enabled = False
'  cboFromColumn.ListIndex = -1
'  cboFromColumn.Enabled = False
'  cboFromTable.BackColor = &H8000000F
'  cboFromColumn.BackColor = &H8000000F
'
'  txtCalculation.Text = ""
'  txtCalculation.Tag = 0
'  cmdCalculation.Enabled = False
'
'  txtOther.Enabled = True
'  txtOther.BackColor = &H80000005
'
'  cboOther.Enabled = False
'  cboOther.BackColor = &H8000000F
'  cboOther.ListIndex = -1
'
  txtOther.Text = vbNullString
  txtHeading.Text = vbNullString
'  If Me.Visible Then
'    txtOther.SetFocus
'  End If
'
'  With txtLength
'    .Enabled = False
'    .Text = 0
'    .BackColor = &H8000000F
'  End With
    
  RefreshControls
  
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
  
  If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = 0
    Exit Sub
  End If

End Sub

Private Sub spnDec_Change()
'NHRD25092006 Fault 11446 Fault calls for it to mirror Custom Report spinners
'the easiest way to do that seems to be to just remove the coding here and txtlength_change
'  If Me.spnDec.Value >= Me.txtLength.Value Then
'    Me.spnDec.Value = IIf(Me.txtLength.Value < 1, 0, Me.txtLength.Value)
'  End If

  'MH20070226 Fault 11953
  lblSizeIncreased.Visible = (spnDec.Value > 0)

End Sub

Private Sub txtHeading_GotFocus()
  With txtHeading
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLength_Change()
'NHRD25092006 Fault 11446 Fault calls for it to mirror Custom Report spinners
'the easiest way to do that seems to be to just remove the coding here and spnDec_change
'  If Me.spnDec.Value >= Me.txtLength.Value Then
'    Me.spnDec.Value = IIf(Me.txtLength.Value < 1, 0, Me.txtLength.Value)
'  End If
  
End Sub

Private Sub txtOther_Change()
  txtLength.Text = Len(txtOther.Text)
End Sub


Private Function GetColumnDetails(lTableID As Long) As Recordset

    Dim sSQL As String
    Dim rsColumns As Recordset
    
    sSQL = "SELECT ColumnName, ColumnID, Size " & _
       "FROM ASRSysColumns " & _
       "WHERE TableID = " & lTableID & _
       " AND Datatype <> " & sqlVarBinary & _
       " AND Datatype <> " & sqlOle & _
       " AND ColumnType <> " & Trim(Str(colSystem)) & _
       " AND ColumnType <> " & Trim(Str(colLink))

    Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

    Set GetColumnDetails = rsColumns
    
End Function

Public Function GetColumnSize(lColumnID As Long) As Long

  Dim sSQL As String
  Dim rsColumns As Recordset
  
'  sSQL = "Select Size, Datatype From ASRSysColumns Where ColumnID = " & lColumnID
  sSQL = "SELECT [Size] FROM [dbo].[ASRSysColumns] WHERE [ColumnID] = " & lColumnID & ";"
  
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

'  Select Case rsColumns("Datatype")
'    Case sqlDate:
'      GetColumnSize = 12
'    Case sqlBoolean:
'      GetColumnSize = 5
'    Case Else:
'      GetColumnSize = rsColumns(0)
'  End Select
  
  GetColumnSize = rsColumns(0).Value
  
  Set rsColumns = Nothing
    
End Function

Public Function GetColumnDecimals(lColumnID As Long) As Long

  Dim sSQL As String
  Dim rsColumns As Recordset
  
'  sSQL = "Select Size, Datatype From ASRSysColumns Where ColumnID = " & lColumnID
  sSQL = "Select Decimals From ASRSysColumns Where ColumnID = " & lColumnID
  
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  'If rsColumns(0).Value > 0 Then
  '  Stop
  'End If
  
  GetColumnDecimals = rsColumns(0).Value
  
  Set rsColumns = Nothing
    
End Function

Public Sub SetCMGOptions(ByVal sDefaultCMGCode As String)
  
  'NPG20071218 Fault 12867
  ' If non-character, move CMG up to replace the convert case combo
  
  
  'Sets options on this form to enable the CMG Options
  lblCMGCode.Visible = True
  txtCMGCode.Visible = True
  txtCMGCode.Text = sDefaultCMGCode

  With lblSizeIncreased
    .Top = 720
    .Left = 3200
    .Height = 615
    .Width = 1215
  End With

  'NPG20080617 Suggestion S000816
  ' Expand the size of the form and show the Suppress Nulls check box for CMG exports
  Me.Height = 5325
  cmdOK.Top = 4350
  cmdCancel.Top = 4350
  fraType.Height = 4020
  fraProperties.Height = 2775
  
  chkSuppressNulls.Visible = True

End Sub
'Private Function AlreadyUsedInExport(plngColExprID As Long, Optional plngExclusion As Long) As Boolean
''NHRD03062003 Fault 4908 Re-introduced this code for this fault
'  Dim pintOldPosition As Integer
'  Dim pvarbookmark As Variant
'  Dim pintLoop As Integer
'
'  With mfrmForm.grdColumns
'
'    ' Store the old position so we can return it after we have looped thru the grid
'    pintOldPosition = .AddItemRowIndex(.Bookmark)
'
'    ' Loop thru the import grid, adding data to the combo if they are columns
'    .MoveFirst
'      Do Until pintLoop = .Rows
'        pvarbookmark = .GetBookmark(pintLoop)
'        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
'          If plngExclusion = 0 Then
'            AlreadyUsedInExport = True
'            .Bookmark = .GetBookmark(pintOldPosition)
'            .SelBookmarks.Add .Bookmark
'            Exit Function
'          Else
'            If .Columns("ColExprID").CellText(pvarbookmark) = plngExclusion Then
'              AlreadyUsedInExport = False
'            Else
'              AlreadyUsedInExport = True
'              .Bookmark = .GetBookmark(pintOldPosition)
'              .SelBookmarks.Add .Bookmark
'              Exit Function
'            End If
'          End If
'        End If
'        pintLoop = pintLoop + 1
'      Loop
'
'    .Bookmark = .GetBookmark(pintOldPosition)
'    .SelBookmarks.Add .Bookmark
'
'  End With
'
'  AlreadyUsedInExport = False
'
'End Function

Public Sub SetConvertCaseOptions(ByVal sDefaultConvertCase As Integer)
  
  'Sets options on this form to enable the CMG Options
  'lblCMGCode.Visible = True
  'txtCMGCode.Visible = True
  ' sDefaultConvertCase = IIf(IsNull(sDefaultConvertCase, 0), sDefaultConvertCase, 0)
  cboConvCase.ListIndex = sDefaultConvertCase

End Sub


