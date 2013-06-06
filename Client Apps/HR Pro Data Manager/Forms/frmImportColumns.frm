VERSION 5.00
Begin VB.Form frmImportColumns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Column"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1044
   Icon            =   "frmImportColumns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraType 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1260
      Begin VB.OptionButton optTable 
         Caption         =   "Colu&mn"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton optFiller 
         Caption         =   "&Filler"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   760
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2835
      TabIndex        =   12
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4080
      TabIndex        =   13
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Frame fraField 
      Height          =   2295
      Left            =   1500
      TabIndex        =   3
      Top             =   120
      Width           =   3780
      Begin VB.CheckBox chkLookupEntries 
         Caption         =   "C&reate missing lookup table entries"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1875
         Width           =   3465
      End
      Begin VB.TextBox txtLength 
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1100
         Width           =   1545
      End
      Begin VB.CheckBox chkKey 
         Caption         =   "&Key field"
         Height          =   195
         Left            =   200
         TabIndex        =   10
         Top             =   1560
         Width           =   1125
      End
      Begin VB.ComboBox cboTable 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2450
      End
      Begin VB.ComboBox cboColumn 
         Height          =   315
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   700
         Width           =   2450
      End
      Begin VB.Label lblLength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   1155
         Width           =   570
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   765
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmImportColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsData As New HRProDataMgr.clsDataAccess
Private mblnNew As Boolean
Private mblnCancelled As Boolean
Private mlngEditingColumnID As Long
Private mfrmForm As frmImport

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Function Initialise(bNew As Boolean, sType As String, lTableID As Long, lColExprID As Long, fkey As Boolean, iSize As String, _
                           blnInsertLookup As Boolean, Optional frmParentForm As frmImport) As Boolean
        
  Dim prstTemp As Recordset
  
  Set mclsData = New HRProDataMgr.clsDataAccess
  Set mfrmForm = frmParentForm
  
  mblnNew = bNew
  mlngEditingColumnID = lColExprID
  
  With cboTable
        
    .Clear
        
    'Add the base table
    If frmParentForm.cboBaseTable.Text <> "<None>" Then
      .AddItem frmParentForm.cboBaseTable.Text
      .ItemData(.NewIndex) = frmParentForm.cboBaseTable.ItemData(frmParentForm.cboBaseTable.ListIndex)
    End If
    
    'Add its parents
    Set prstTemp = mclsData.OpenRecordset("SELECT AsrSysTables.tablename, AsrSysRelations.ParentID FROM AsrSysTables, AsrSysRelations WHERE asrsystables.tableid = asrsysrelations.parentid and asrsysrelations.childid = " & frmParentForm.cboBaseTable.ItemData(frmParentForm.cboBaseTable.ListIndex), adOpenForwardOnly, adLockReadOnly)
    
    Do Until prstTemp.EOF
      .AddItem prstTemp.Fields(0)
      .ItemData(.NewIndex) = prstTemp.Fields(1)
      prstTemp.MoveNext
    Loop
    
    Set prstTemp = Nothing
    
    .Enabled = (.ListCount > 1)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    
    'MH20001123 Fault 1332 Next line not required !
    'If .ListCount >= 0 Then .ListIndex = 0 Else .ListIndex = -1
        
    If cboColumn.ListCount < 2 Then
      cboColumn.Enabled = False
      cboColumn.BackColor = vbButtonFace
    End If
    
  End With
    
  'If we are editing an existing grid entry, display the data
  If Not bNew Then
      
    Select Case sType
    
      Case "C"
        optTable.Value = True
        EnableColumnControls True
        SetComboText cboTable, datGeneral.GetTableName(lTableID)
        SetComboText cboColumn, datGeneral.GetColumnName(lColExprID)
        'If fkey Then Me.chkKey.Value = 1
        chkKey.Value = IIf(fkey, vbChecked, vbUnchecked)
        
        'MH20030520
        chkLookupEntries.Value = IIf(blnInsertLookup, vbChecked, vbUnchecked)
        
        If iSize <> "" Then txtLength.Text = iSize
'        cboTable.Enabled = (cboTable.ListCount > 1)
'        cboColumn.Enabled = (cboColumn.ListCount > 1)
'        If cboTable.ListCount = 1 Then cboTable.ListIndex = 0
'        cboTable.BackColor = IIf(cboTable.Enabled, vbWindowBackground, vbButtonFace)
'        cboColumn.BackColor = IIf(cboColumn.Enabled, vbWindowBackground, vbButtonFace)
'        txtLength.BackColor = &H8000000F
'        txtLength.Enabled = False
'        'txtLength.Text = "1"
'
'        If mfrmForm.cboFileFormat.ItemData(mfrmForm.cboFileFormat.ListIndex) = 1 Then
'          txtLength.Enabled = True
'          txtLength.BackColor = &H80000005
'         ' txtLength.Text = GetColumnSize(cboColumn.ItemData(cboColumn.ListIndex))
'
'
'          'MH20020226 Don't know who put this comment in (Roy??)
'          'But anyway, spoke to Jed and surely it should be 10 for a date column
'          '(If you want two spaces after it then use a filler....)
'
'          ''If column is date, suggest 12 as length for fixed length (xx/xx/xxxx + 2 spaces)
'          'If txtLength.Text = "0" Then txtLength = "12" ' Or txtLength.Text = "1"
'          If txtLength.Text = "0" Then txtLength = "10"
'
'
'        Else
'          txtLength.Enabled = False
'          txtLength.BackColor = &H8000000F
''            txtLength.Text = ""
'        End If
      
      Case "F"
        
        optFiller.Value = True
        'If iSize <> "" Then txtLength.Text = iSize
        EnableColumnControls False
        txtLength.Text = iSize
'        cboTable.Enabled = False
'        cboColumn.Enabled = False
'        cboTable.BackColor = &H8000000F
'        cboColumn.BackColor = &H8000000F
'        chkKey.Enabled = False
'        chkLookupEntries.Enabled = False
'
'        If mfrmForm.cboFileFormat.ItemData(mfrmForm.cboFileFormat.ListIndex) = 1 Then
'          txtLength.Enabled = True
'          txtLength.BackColor = &H80000005
'          'txtLength.Text = "1"
'        Else
'          txtLength.Enabled = False
'          txtLength.BackColor = &H8000000F
'          'txtLength.Text = "1"
'        End If
    
    End Select
    
    Me.lblLength.Enabled = Me.txtLength.Enabled
    Me.lblTitle(0).Enabled = Me.cboTable.Enabled
    Me.lblTitle(1).Enabled = Me.cboColumn.Enabled
    'EnableColumnControls True

  Else
    'optTable.Value = True
    optTable_Click
  End If
    
  Screen.MousePointer = vbDefault
  Initialise = True
  
End Function


Private Sub cmdCancel_Click()
  Cancelled = True
  Unload Me
End Sub

Private Sub cmdOK_Click()

  'Do some validation to make sure everything required has been entered
  If optTable Then
    If cboTable.Text = "" Or cboColumn.Text = "" Then
      If cboColumn.ListCount > 0 Then
        COAMsgBox "You must select a table and column.", vbExclamation, Me.Caption
      Else
        COAMsgBox "All available columns have been selected.", vbExclamation, Me.Caption
      End If
      
      Exit Sub
    End If

    If Len(Trim(txtLength.Text)) > 0 Then
      If IsNumeric(txtLength.Text) Then
        If Val(txtLength.Text) = "0" Then
          COAMsgBox "The Length value must be greater than 0 characters.", vbExclamation, Me.Caption
          Exit Sub
        End If
        
        ' JIRA - 604 - Don't let the user go nuts with this figure
        If Val(txtLength.Text) > VARCHAR_MAX_Size Then
          COAMsgBox "The Length value cannot exceed " & VARCHAR_MAX_Size & " characters.", vbExclamation, Me.Caption
          txtLength.Text = VARCHAR_MAX_Size
          Exit Sub
        End If
        
      Else
        COAMsgBox "The Length value must be numeric.", vbExclamation, Me.Caption
        Exit Sub
      End If
    End If
  End If
    
  Cancelled = False
  Me.Hide
  
End Sub

Private Sub Form_Load()
  
'  spnSize.Enabled = (mfrmForm.cboFileFormat.ListIndex = 1)
'  spnSize.BackColor = IIf(mfrmForm.cboFileFormat.ListIndex = 1, &H80000005, &H8000000F)

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set mclsData = Nothing

End Sub

Private Sub OptFiller_Click()
  EnableColumnControls False
End Sub

Private Sub optTable_Click()
  
  EnableColumnControls True

  If cboTable.ListCount > 0 Then
    cboTable.ListIndex = 0
  End If

'  'MH20001123 Fault 1332
'  With cboTable
'    If .ListCount > 0 Then
'      With mfrmForm.cboBaseTable
'        If .ListIndex >= 0 Then
'          SetComboItem cboTable, .ItemData(.ListIndex)
'        End If
'      End With
'      If .ListIndex < 0 Then .ListIndex = 0
'      .Enabled = (.ListCount > 1)
'    Else
'      .Enabled = False
'    End If
'  End With
'  cboColumn.Enabled = (cboColumn.ListCount > 1)
'
'  Me.AutoRedraw = True
'  Me.Refresh
'
'  If cboColumn.ListIndex >= 0 Then
'    txtLength.Text = GetColumnSize(cboColumn.ItemData(cboColumn.ListIndex))
'  End If
'
'  If mfrmForm.cboFileFormat.ItemData(mfrmForm.cboFileFormat.ListIndex) = 1 Then
'    txtLength.Enabled = True
'    txtLength.BackColor = vbWindowBackground ' vbButtonFace '&H80000005
'  Else
'    txtLength.Enabled = False
'    txtLength.BackColor = vbButtonFace 'vbWindowBackground ' &H8000000F
'  End If
'
'  If cboTable.Text <> mfrmForm.cboBaseTable.Text Then
'    chkKey.Value = vbChecked
'    chkKey.Enabled = False
'  Else
'    chkKey.Enabled = True
'  End If
'  chkLookupEntries.Enabled = True
'
'
'  lblLength.Enabled = txtLength.Enabled
'  lblTitle(0).Enabled = cboTable.Enabled
'  lblTitle(1).Enabled = cboColumn.Enabled
  
End Sub

Private Sub EnableColumnControls(blnEnabled As Boolean)

  Dim blnSizeEnabled As Boolean

  With mfrmForm.cboFileFormat
    blnSizeEnabled = (blnEnabled Or .ItemData(.ListIndex) = 1)
  End With


  lblTitle(0).Enabled = blnEnabled
  EnableCombo cboTable, blnEnabled

  lblTitle(1).Enabled = blnEnabled
  EnableCombo cboColumn, blnEnabled
  
  If Not blnEnabled Then
    chkKey.Enabled = blnEnabled
    chkKey.Value = vbUnchecked
    chkLookupEntries.Enabled = blnEnabled
    chkLookupEntries.Value = vbUnchecked
  End If
  
  
  lblLength.Enabled = blnSizeEnabled
  txtLength.Enabled = blnSizeEnabled
  txtLength.BackColor = IIf(blnSizeEnabled, vbWindowBackground, vbButtonFace)
  If Not blnEnabled Then
    txtLength.Text = IIf(blnSizeEnabled, "0", vbNullString)
  End If
  'If Not blnSizeEnabled Then
  '  txtLength.Text = vbNullString
  'End If

End Sub


Private Sub txtLength_KeyPress(KeyAscii As Integer)
  
  If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = 0
    Exit Sub
  End If

End Sub

Private Function GetColumnDetails(lTableID As Long) As ADODB.Recordset

    Dim sSQL As String
    Dim rsColumns As Recordset
    
    'sSQL = "Select ColumnName, ColumnID, Size From ASRSysColumns Where TableID = " & lTableID & " AND Datatype > 0 AND Datatype <> 4"
    
    'sSQL = "Select ColumnName, ColumnID, Size From ASRSysColumns Where TableID = " & lTableID & " AND Datatype <> " & sqlVarBinary & " AND Datatype <> " & sqlOle
    
    'sSQL = "SELECT ColumnName, ColumnID, Size " & _
           "FROM ASRSysColumns " & _
           "WHERE TableID = " & lTableID & _
           " AND Datatype <> " & sqlVarBinary & _
           " AND Datatype <> " & sqlOle & _
           " AND ColumnType <> " & Trim(Str(colSystem)) & _
           " AND ColumnType <> " & Trim(Str(colLink)) & _
           " AND ReadOnly = 0"
    
    sSQL = "SELECT ColumnName, ColumnID, Size " & _
           "FROM ASRSysColumns " & _
           "WHERE TableID = " & lTableID & _
           " AND Datatype <> " & sqlVarBinary & _
           " AND Datatype <> " & sqlOle & _
           " AND ColumnType <> " & Trim(Str(colSystem)) & _
           " AND ColumnType <> " & Trim(Str(colLink)) '& _
           " AND ReadOnly = 0"
    
    Set rsColumns = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Set GetColumnDetails = rsColumns
    
End Function

Private Function GetColumnSize(lColumnID As Long) As Long

  Dim rsColumns As Recordset
  Dim sSQL As String
  Dim blnForceKeyed As Boolean
  Dim blnLookupColumn As Boolean
  Dim blnDefaultKeyed As Boolean
  Dim blnAllowKeyed As Boolean
  Dim lngLen As Long

  sSQL = "Select Size, Datatype, ReadOnly, ColumnType From ASRSysColumns Where ColumnID = " & lColumnID
  
  Set rsColumns = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

'  If mfrmForm.cboFileFormat.ItemData(mfrmForm.cboFileFormat.ListIndex) = 1 Then
    Select Case rsColumns("Datatype")
      Case sqlDate:
        txtLength.Text = "10"
      Case sqlBoolean:
        txtLength.Text = "5"
      
      'MH20020426 Fault 3678
      'Working Pattern Size = 14
      Case sqlLongVarChar:
        txtLength.Text = "14"
      
      Case Else:
        lngLen = rsColumns(0)
        txtLength.Text = IIf(lngLen > 999999999, "999999999", CStr(lngLen))
    End Select
'  Else
'    txtLength.Text = vbNullString
'  End If

  blnDefaultKeyed = (cboTable.Text <> mfrmForm.cboBaseTable.Text Or rsColumns!ReadOnly = True)
  blnAllowKeyed = (cboTable.ListCount > 2 Or (cboTable.Text = mfrmForm.cboBaseTable.Text And rsColumns!ReadOnly = False))

  'blnForceKeyed = (cboTable.Text <> mfrmForm.cboBaseTable.Text Or rsColumns!ReadOnly = True)
  'blnForceKeyed = (cboTable.ListCount < 3 And (cboTable.Text <> mfrmForm.cboBaseTable.Text Or rsColumns!ReadOnly = True))

  'chkKey.Enabled = (cboTable.ListCount > 2 Or Not blnForceKeyed)
  'chkKey.Value = IIf(blnForceKeyed, vbChecked, vbUnchecked)
  chkKey.Enabled = blnAllowKeyed
  chkKey.Value = IIf(blnDefaultKeyed, vbChecked, vbUnchecked)


  blnLookupColumn = (rsColumns!ColumnType = colLookup)
  chkLookupEntries.Enabled = blnLookupColumn
  chkLookupEntries.Value = IIf(blnLookupColumn, vbChecked, vbUnchecked)

  rsColumns.Close
  Set rsColumns = Nothing

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    Case KeyCode = 192
        KeyCode = 0
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode <> vbFormCode Then
    Cancelled = True
  End If

End Sub

Private Function AlreadyUsedInImport(plngColExprID As Long, Optional plngExclusion As Long) As Boolean

  Dim pintOldPosition As Integer
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  
  With mfrmForm.grdColumns
    
    ' Store the old position so we can return it after we have looped thru the grid
    pintOldPosition = .AddItemRowIndex(.Bookmark)
    
    ' Loop thru the import grid, adding data to the combo if they are columns
    .MoveFirst
      Do Until pintLoop = .Rows
        pvarbookmark = .GetBookmark(pintLoop)
        If .Columns("ColExprID").CellText(pvarbookmark) = plngColExprID Then
          If plngExclusion = 0 Then
            AlreadyUsedInImport = True
            .Bookmark = .GetBookmark(pintOldPosition)
            .SelBookmarks.Add .Bookmark
            Exit Function
          Else
            If .Columns("ColExprID").CellText(pvarbookmark) = plngExclusion Then
              AlreadyUsedInImport = False
            Else
              AlreadyUsedInImport = True
              .Bookmark = .GetBookmark(pintOldPosition)
              .SelBookmarks.Add .Bookmark
              Exit Function
            End If
          End If
        End If
        pintLoop = pintLoop + 1
      Loop
    
    .Bookmark = .GetBookmark(pintOldPosition)
    .SelBookmarks.Add .Bookmark
    
  End With

  AlreadyUsedInImport = False

End Function

Private Sub cboColumn_Click()
  
  'If no column selected, exit
  If cboColumn.Text = "" Then Exit Sub
  
''  'Display default column size in fixed length field

  'TM20011218 Fault 3039 - Show the column sizes for all file types.
  'If mfrmForm.cboFileFormat.ItemData(mfrmForm.cboFileFormat.ListIndex) = 1 Then
    GetColumnSize cboColumn.ItemData(cboColumn.ListIndex)
    'If txtLength.Text = "0" Or txtLength.Text = "1" Then txtLength = "12"
  'End If
  
End Sub

Private Sub cboTable_Click()
    
  'If no table selected, wipe column listbox and exit. Should never happen.
  If cboTable.Text = "" Then
    cboColumn.Clear
    Exit Sub
  End If
  
  Dim sSQL As String
  Dim rsCols As Recordset
  
  Screen.MousePointer = vbHourglass
  
  'Get all the columns for the selected table
  Set rsCols = GetColumnDetails(cboTable.ItemData(cboTable.ListIndex))
   
  
  'MH20000705 Fault 549
  mfrmForm.grdColumns.Redraw = False
    
  With cboColumn
    .Clear
    Do While Not rsCols.EOF
    
      If AlreadyUsedInImport(rsCols!ColumnID, IIf(mblnNew = False, mlngEditingColumnID, 0)) = False Then
        .AddItem rsCols!ColumnName
        .ItemData(.NewIndex) = rsCols!ColumnID
      End If
      
      rsCols.MoveNext
    Loop

    'If .ListCount > 0 Then .ListIndex = 0 Else .ListIndex = -1
    .Enabled = (.ListCount > 1)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    If .ListCount > 0 Then .ListIndex = 0 Else .ListIndex = -1
    
  End With
    
  mfrmForm.grdColumns.Redraw = True
    
    
  ' If its not the base table selected, then auto check the key field and disable it
  ' otherwise enable the checkbox
  If cboTable.Text <> mfrmForm.cboBaseTable.Text Then
    chkKey.Value = vbChecked
    'chkKey.Enabled = False
    chkKey.Enabled = (Me.cboTable.ListCount > 2)
  Else
    chkKey.Enabled = True
    chkKey.Value = vbUnchecked
  End If
  
  rsCols.Close
  Set rsCols = Nothing
  Screen.MousePointer = vbNormal

End Sub


