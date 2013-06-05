VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmLinkFind 
   Caption         =   "Find"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1045
   Icon            =   "frmLinkFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOrder 
      Caption         =   "Options :"
      Height          =   1200
      Left            =   100
      TabIndex        =   7
      Top             =   100
      Width           =   6000
      Begin VB.ComboBox cmbView 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   4900
      End
      Begin VB.ComboBox cmbOrders 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   4900
      End
      Begin VB.Label lblOrder 
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   255
         Left            =   195
         TabIndex        =   9
         Top             =   765
         Width           =   645
      End
      Begin VB.Label lblView 
         BackStyle       =   0  'Transparent
         Caption         =   "View :"
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   100
      TabIndex        =   6
      Top             =   3100
      Width           =   4000
      Begin VB.CommandButton cmdClearField 
         Cancel          =   -1  'True
         Caption         =   "Clear &Field"
         Height          =   400
         Left            =   1400
         TabIndex        =   4
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   2800
         TabIndex        =   5
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssOleDBGridFindColumns 
      Height          =   1410
      Left            =   100
      TabIndex        =   2
      Top             =   1500
      Width           =   6000
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      GroupHeadLines  =   0
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
      UseExactRowCount=   0   'False
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
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   10583
      _ExtentY        =   2487
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
Attribute VB_Name = "frmLinkFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Link table variables.
Private mobjLinkTable As CTablePrivilege

' Record display option variables.
Private mlngOrderID As Long
Private mlngViewID As Long

' Recordset variables.
Private mrsFindRecords As ADODB.Recordset
Private mlngRecordCount As Long
Private mlngLinkRecordID As Long

' Form handling variables.
Private mfSizing As Boolean
Private mfCancelled As Boolean
Private mfFormattingGrid As Boolean
Private mfFirstColumnsMatch As Boolean
Private mfFirstColumnAscending As Boolean
Private miFirstColumnDataType As Integer


Private mlngLookupColumnID As Long
Private mlngLookupTableID As Long
Private msLookupTableName As String
Private msLookupColumnName As String
Private miLookupDataType As Integer
Private mvLookupSelectedValue As Variant
Private mvLookupSelectedFilterCode As Variant
Private mlngLookupFilterColumnID As Long
Private mlngLookupFilterTableID As Long
Private msLookupFilterTableName As String
Private msLookupFilterColumnName As String
Private miLookupFilterDataType As Integer
Private miLookupColumnGridPosition As Integer

Private msGetRecsSQL As String

Private mavFindColumns() As Variant        ' Find columns details

Private Const dblFINDFORM_MINWIDTH = 5000
Private Const dblFINDFORM_MINHEIGHT = 5000

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property

Public Function Initialise(plngTableID As Long, _
  plngViewID As Long, _
  plngOrderID As Long, _
  Optional pvLookupColumnID As Variant, _
  Optional pvCurrentLookupFilterCode As Variant, _
  Optional plngLookupFilterColumnID As Long) As Boolean
  ' Initialise the link find form.
  Dim sSQL As String
  Dim rsInfo As Recordset
  
  ' Get the link table object.
  Set mobjLinkTable = gcoTablePrivileges.FindTableID(plngTableID)

  ' Get the table's default order ID.
  If plngOrderID > 0 Then
    mlngOrderID = plngOrderID
  Else
    mlngOrderID = mobjLinkTable.DefaultOrderID
  End If
  
  ' Set the find window caption
  'JPD 20031217 Islington changes
  mlngLookupColumnID = 0
  mlngLookupFilterColumnID = 0
  If IsMissing(pvLookupColumnID) Then
    Me.Caption = "Link to " & mobjLinkTable.TableName
    RemoveIcon Me
    cmdClearField.Visible = False
    cmdSelect.Left = cmdClearField.Left
  Else
    Me.Caption = "Lookup to " & mobjLinkTable.TableName
    RemoveIcon Me
    mlngLookupColumnID = CLng(pvLookupColumnID)
    mvLookupSelectedFilterCode = pvCurrentLookupFilterCode
    mlngLookupFilterColumnID = plngLookupFilterColumnID

    sSQL = "SELECT ASRSysTables.tableID," & _
      " ASRSysTables.tableName," & _
      " ASRSysColumns.columnName," & _
      " ASRSysColumns.dataType" & _
      " FROM ASRSysColumns" & _
      " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
      " WHERE ASRSysColumns.columnID = " & CStr(mlngLookupColumnID)
    Set rsInfo = datGeneral.GetRecords(sSQL)
    If Not (rsInfo.EOF And rsInfo.BOF) Then
      mlngLookupTableID = rsInfo!TableID
      msLookupTableName = rsInfo!TableName
      msLookupColumnName = rsInfo!ColumnName
      miLookupDataType = rsInfo!DataType
    End If
    rsInfo.Close
    Set rsInfo = Nothing

    If mlngLookupFilterColumnID > 0 Then
      sSQL = "SELECT ASRSysTables.tableID," & _
        " ASRSysTables.tableName," & _
        " ASRSysColumns.columnName," & _
        " ASRSysColumns.dataType" & _
        " FROM ASRSysColumns" & _
        " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
        " WHERE ASRSysColumns.columnID = " & CStr(mlngLookupFilterColumnID)
      Set rsInfo = datGeneral.GetRecords(sSQL)
      If Not (rsInfo.EOF And rsInfo.BOF) Then
        mlngLookupFilterTableID = rsInfo!TableID
        msLookupFilterTableName = rsInfo!TableName
        msLookupFilterColumnName = rsInfo!ColumnName
        miLookupFilterDataType = rsInfo!DataType
      End If
      rsInfo.Close
      Set rsInfo = Nothing
    End If
  End If
  
  ' Populate the View combo.
  ConfigureViewCombo plngViewID
  If cmbView.ListCount = 0 Then
    COAMsgBox "You do not have 'read' permission on this table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
    mfCancelled = True
    Me.Hide
    Initialise = False
    Exit Function
  End If
  
  'MH20010205 Fault 1793
  'This bit of code was bypassing what had been selected from the combo box
  '(i.e. if there is a default view then display the records in the default view!)
  'If mobjLinkTable.AllowSelect Then
  '  mlngViewID = 0
  'Else
  '  mlngViewID = cmbView.ItemData(cmbView.ListIndex)
  'End If
  mlngViewID = cmbView.ItemData(cmbView.ListIndex)

  ' Populate the Orders combo.
  ConfigureOrdersCombo
  Screen.MousePointer = vbHourglass
  If cmbOrders.ListCount = 0 Then
    COAMsgBox "There are no orders defined for this table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
    mfCancelled = True
    Me.Hide
    Initialise = False
    Exit Function
  End If
      
  ' Get the link records.
  Set mrsFindRecords = New Recordset
  GetRecords
  
  If ssOleDBGridFindColumns.Rows > 0 Then
    ssOleDBGridFindColumns.MoveFirst
    ssOleDBGridFindColumns.SelBookmarks.Add (ssOleDBGridFindColumns.Bookmark)
  End If
    
  cmdSelect.Enabled = (ssOleDBGridFindColumns.Rows <> 0)
  
  Initialise = True
  
End Function

Private Sub ConfigureGrid()
  ' Configure the grid to display the required columns.
  Dim iLoop As Integer
  Dim lngWidth As Long
  Dim dblPreviousColumnWidth As Double

  UI.LockWindow Me.hWnd
   
  lngWidth = 0
  mfFormattingGrid = True
  
  With ssOleDBGridFindColumns
    .Redraw = False
    .Columns.RemoveAll
    
    For iLoop = 0 To (mrsFindRecords.Fields.Count - 1)
      .Columns.Add iLoop
      .Columns(iLoop).Name = mrsFindRecords.Fields(iLoop).Name
      .Columns(iLoop).Visible = (UCase(mrsFindRecords.Fields(iLoop).Name) <> "ID") And _
        (Left(mrsFindRecords.Fields(iLoop).Name, 1) <> "?")
      .Columns(iLoop).Caption = RemoveUnderScores(mrsFindRecords.Fields(iLoop).Name)
      .Columns(iLoop).Alignment = ssCaptionAlignmentLeft
      .Columns(iLoop).CaptionAlignment = ssColCapAlignUseColumnAlignment
    
      ' If the find column is a logic column then set the grid column style to be 'checkbox'.
      If mrsFindRecords.Fields.Item(iLoop).Type = adBoolean Then
        .Columns(iLoop).Style = ssStyleCheckBox
      End If

      'Has the user changed the width of this column
      dblPreviousColumnWidth = GetUserSetting("FindOrder" + LTrim(Str(mlngOrderID)), mrsFindRecords.Fields(iLoop).Name, 0)
      If dblPreviousColumnWidth > 0 Then .Columns(iLoop).Width = dblPreviousColumnWidth

      ' Total the size of the grid columns
      If .Columns(iLoop).Visible Then
        lngWidth = lngWidth + .Columns(iLoop).Width
      End If
    Next iLoop
      
    mfFormattingGrid = False
    .Rebind
    .Rows = mlngRecordCount
    .Redraw = True
  End With
    
  'Adjust size of find window to fit the grid
  lngWidth = lngWidth + (fraOrder.Left * 2) + _
    (((UI.GetSystemMetrics(SM_CXFRAME) * 2) + _
    UI.GetSystemMetrics(SM_CXBORDER)) * Screen.TwipsPerPixelX)
      
  If ssOleDBGridFindColumns.Rows > ssOleDBGridFindColumns.VisibleRows Then
    lngWidth = lngWidth + (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + 20
  End If
    
  Me.Width = lngWidth + 120

  UI.UnlockWindow
  
End Sub
Private Function RecordCount() As Long
  ' Return the number of records in the recordset.
  Dim rsTemp As Recordset
  
  Set rsTemp = datGeneral.GetRecords(msGetRecsSQL)
  If (rsTemp.EOF And rsTemp.BOF) Then
    RecordCount = 0
  Else
    RecordCount = rsTemp(0)
  End If
  rsTemp.Close
  Set rsTemp = Nothing
  
End Function

Private Sub ConfigureOrdersCombo()
  ' Initialise the form to be called from a primary screen.
  Dim fOrderFound As Boolean
  Dim iIndex As Integer
  Dim rsOrder As Recordset

  If mlngViewID > 0 Then
    Set rsOrder = datGeneral.GetViewOrders(mlngViewID, mobjLinkTable.TableID)
  Else
    Set rsOrder = datGeneral.GetOrders(mobjLinkTable.TableID)
  End If
  
  ' Populate the Orders combo.
  With cmbOrders
    .Clear
  
    ' Add the orders to the combo.
    Do While Not rsOrder.EOF
      .AddItem RemoveUnderScores(Trim(rsOrder!Name))
      .ItemData(.NewIndex) = rsOrder!OrderID
      rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
        
    If .ListCount > 0 Then
      ' Select the last used order if possible.
      fOrderFound = False
      For iIndex = 0 To (.ListCount - 1)
        If (.ItemData(iIndex) = mlngOrderID) Then
          fOrderFound = True
          .ListIndex = iIndex
          Exit For
        End If
      Next iIndex
      
      If Not fOrderFound Then
        .ListIndex = 0
      End If
      
      .Enabled = True
    Else
      ' No orders.
      .AddItem "<No order>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
      COAMsgBox "No orders defined for this " & IIf(mlngViewID > 0, "table.", "view."), vbInformation, Me.Caption
      .Enabled = False
    End If
  End With
  
End Sub


Private Sub cmbOrders_Click()
  Dim fOK As Boolean
  
  ' Do nothing if the form is not visible.
  fOK = Me.Visible

  If fOK Then
    Screen.MousePointer = vbHourglass
  End If

  ' Do nothing if there are no records.
  If fOK Then
    fOK = (mlngRecordCount > 0)
  End If
    
  ' Do nothing if there are no orders defined.
  If fOK Then
    fOK = (cmbOrders.ItemData(cmbOrders.ListIndex) > 0)
  End If

  If fOK Then
    ' Set the order ID variable.
    mlngOrderID = cmbOrders.ItemData(cmbOrders.ListIndex)
    
    GetRecords
  
    With ssOleDBGridFindColumns
      cmdSelect.Enabled = (.Rows > 0)
    
      If .Rows > 0 Then
        .MoveFirst
        .SelBookmarks.Add (.Bookmark)
        .SetFocus
      End If
    End With
  End If

  Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmbView_Click()
  Dim fOK As Boolean
  
  ' Do nothing if the form is not visible.
  fOK = Me.Visible
  
  If fOK Then
    Screen.MousePointer = vbHourglass
  End If

  If fOK Then
    ' Set the view ID variable.
    If mobjLinkTable.AllowSelect And _
      (cmbView.ListIndex = 0) Then
      mlngViewID = 0
    Else
      mlngViewID = cmbView.ItemData(cmbView.ListIndex)
    End If
    
    GetRecords
    
    With ssOleDBGridFindColumns
      cmdSelect.Enabled = (.Rows > 0)
    
      If .Rows > 0 Then
        .MoveFirst
        .SelBookmarks.Add (.Bookmark)
        .SetFocus
      End If
    End With
  End If

  Screen.MousePointer = vbDefault
  
End Sub

Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide
  
End Sub

Private Sub cmdClearField_Click()
  mfCancelled = False
  mvLookupSelectedValue = ""
  Me.Hide

End Sub

Private Sub cmdSelect_Click()
  ssOleDBGridFindColumns_DblClick
  
End Sub

Private Sub Form_Activate()

  'Highlight the current row in the grid
  If Not WindowState = vbMinimized Then
    
    ' Show the find records
    ssOleDBGridFindColumns.Visible = True
    ssOleDBGridFindColumns.SetFocus
  
  End If

End Sub

Private Sub Form_Load()
  
  RemoveIcon Me
  Hook Me.hWnd, dblFINDFORM_MINWIDTH, dblFINDFORM_MINHEIGHT
  
  ssOleDBGridFindColumns.RowHeight = 239
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
    Cancel = True
    Me.Hide
  End If
  
End Sub

Private Sub Form_Resize()
  Dim lCount As Long
  Dim lWidth As Long
  Dim iLastColumnIndex As Integer
  Dim iMaxPosition As Integer
    
  Const dblCOORD_XGAP = 200
  Const dblCOORD_YGAP = 200
  
  If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized And Me.ScaleHeight > 0 Then

    'JPD 20030908 Fault 5756
    DisplayApplication
    
    fraOrder.Width = Me.ScaleWidth - (dblCOORD_XGAP * 2)
    cmbOrders.Width = fraOrder.Width - (dblCOORD_XGAP * 6)
    cmbView.Width = fraOrder.Width - (dblCOORD_XGAP * 6)

    ' Size the Find grid.
    With ssOleDBGridFindColumns
      .Width = fraOrder.Width
      .Height = Me.ScaleHeight - .Top - fraButtons.Height - (2 * dblCOORD_YGAP)
    End With
  
    fraButtons.Top = ssOleDBGridFindColumns.Top + ssOleDBGridFindColumns.Height + dblCOORD_YGAP
    fraButtons.Left = Me.ScaleWidth - fraButtons.Width - dblCOORD_XGAP
  
    DoEvents
  
    If ((ssOleDBGridFindColumns.Rows - ssOleDBGridFindColumns.FirstRow + 1) < (ssOleDBGridFindColumns.VisibleRows)) And _
      (ssOleDBGridFindColumns.FirstRow > 1) Then
  
      ssOleDBGridFindColumns.FirstRow = IIf(ssOleDBGridFindColumns.Rows - ssOleDBGridFindColumns.VisibleRows + 1 < 1, _
        1, ssOleDBGridFindColumns.Rows - ssOleDBGridFindColumns.VisibleRows + 1)
    End If
    
    ResizeFindColumns

  End If

End Sub
Public Sub ResizeFindColumns()

  Dim dblCurrentSize As Double
  Dim dblNewSize As Double
  Dim iCount As Integer
  Dim dblResizeFactor As Double
  Dim bNeedScrollBars As Boolean

  With ssOleDBGridFindColumns
  
    'JPD 20050311 Fault 9891
    If .Cols = 0 Then Exit Sub

    .Redraw = False
    bNeedScrollBars = IIf(.Rows > .VisibleRows, True, False)

    ' Calculate the existing size of the find grid
    dblCurrentSize = 0
    For iCount = 0 To (.Cols - 1)
      If .Columns(iCount).Visible Then
        dblCurrentSize = dblCurrentSize + .Columns(iCount).Width
      End If
    Next iCount

    ' Calculate size of resized grid
    dblNewSize = .Width
    If bNeedScrollBars Then
      dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX)
    End If
    dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXFRAME) * 2)
    dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX)

    ' Calculate the ratio that the grid needs to be resized to
    dblResizeFactor = Round(dblNewSize / dblCurrentSize, 2)

  ' Scroll through adjusting each column according to the resize factor
    For iCount = 0 To (.Cols - 1)
      If .Columns(iCount).Visible Then

        'Make the last column nice & snug
        If iCount = (.Cols - 2) Then
          .Columns(iCount).Width = dblNewSize
        Else
          .Columns(iCount).Width = (.Columns(iCount).Width * dblResizeFactor)
          dblNewSize = dblNewSize - .Columns(iCount).Width - 5
        End If

        SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(iCount).Name, .Columns(iCount).Width

      End If
    Next iCount

    .Redraw = True

  End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
  'Tidy things up before unloading
  If Not mrsFindRecords Is Nothing Then
    If mrsFindRecords.State = adStateOpen Then
      mrsFindRecords.Close
    End If
    
    Set mrsFindRecords = Nothing
  End If
  
  Unhook Me.hWnd
End Sub

Public Property Get LinkRecordID() As Long
  LinkRecordID = mlngLinkRecordID
  
End Property
Public Property Get LookupValue() As Variant
  LookupValue = mvLookupSelectedValue
  
End Property

Private Sub GetRecords()
  ' Read the required information about the link table.
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim fNoSelect As Boolean
  Dim iNextIndex As Integer
  Dim lngFirstFindColumnID As Long
  Dim lngFirstSortColumnID As Long
  Dim sSQL As String
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sOrderString As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim objTableView As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String
  Dim fLookupColumnDoneF As Boolean
  
  Dim fOK As Boolean
  Dim lngTableID As Long
  Dim sTableName As String
  Dim lngColumnID As Long
  Dim sColumnName As String
  Dim sType As String
  Dim fAscending As Boolean
  Dim iDataType As Integer
  Dim iCount2 As Integer
  
  fNoSelect = False
  
  sOrderString = ""
  sJoinCode = ""
  sColumnList = ""
  fLookupColumnDoneF = (mlngLookupColumnID <= 0)
  miLookupColumnGridPosition = 0
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)

  mfFirstColumnsMatch = False
  lngFirstFindColumnID = 0
  lngFirstSortColumnID = 0
  mfFirstColumnAscending = True
  miFirstColumnDataType = 0
  
  ReDim mavFindColumns(3, 0)
  
  ' Get the default order items from the database.
  Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)
  
  If rsInfo.EOF And rsInfo.BOF Then
    COAMsgBox "No order defined for this " & IIf(mlngViewID > 0, "view.", "table.") & _
      vbCrLf & "Unable to display records.", vbExclamation, "Security"
    mfCancelled = True
    Me.Hide
  Else
    iCount2 = 0
    
    ' Check the user's privilieges on the order columns.
    Do While (Not rsInfo.EOF) Or (Not fLookupColumnDoneF)
      fOK = True
      
      If (Not rsInfo.EOF) Then
        If rsInfo!ColumnID = mlngLookupColumnID Then
          fLookupColumnDoneF = True
        End If
        
        lngTableID = rsInfo!TableID
        sTableName = rsInfo!TableName
        lngColumnID = rsInfo!ColumnID
        sColumnName = rsInfo!ColumnName
        sType = rsInfo!Type
        fAscending = rsInfo!Ascending
        iDataType = rsInfo!DataType
      Else
        lngTableID = mlngLookupTableID
        sTableName = msLookupTableName
        lngColumnID = mlngLookupColumnID
        sColumnName = msLookupColumnName
        sType = "F"
        fAscending = True
        iDataType = miLookupDataType
        fLookupColumnDoneF = True
      End If
      
      If fOK Then
        If (lngColumnID = mlngLookupColumnID) And (sType = "F") Then
          miLookupColumnGridPosition = iCount2
        End If
        
        ' Get the column privileges collection for the given table.
        If lngTableID = mobjLinkTable.TableID Then
          If mlngViewID = 0 Then
            sSource = mobjLinkTable.TableName
          Else
            sSource = GetViewName(mlngViewID)
          End If
        Else
          sSource = sTableName
        End If
        Set objColumnPrivileges = GetColumnPrivileges(sSource)
        sRealSource = gcoTablePrivileges.Item(sSource).RealSource
        
        fColumnOK = objColumnPrivileges.IsValid(sColumnName)
        
        If fColumnOK Then
          fColumnOK = objColumnPrivileges.Item(sColumnName).AllowSelect
        End If
        Set objColumnPrivileges = Nothing
              
        If fColumnOK Then
          ' The column can be read from the base table/view, or directly from a parent table.
          If sType = "F" Then
            ' Add the column to the column list.
            sColumnList = sColumnList & _
              IIf(Len(sColumnList) > 0, ", ", "") & _
              sRealSource & "." & Trim(sColumnName)
            
            mavFindColumns(0, UBound(mavFindColumns, 2)) = lngColumnID
            mavFindColumns(1, UBound(mavFindColumns, 2)) = datGeneral.GetDataSize(lngColumnID)
            mavFindColumns(2, UBound(mavFindColumns, 2)) = datGeneral.GetDecimalsSize(lngColumnID)
            mavFindColumns(3, UBound(mavFindColumns, 2)) = datGeneral.DoesColumnUseSeparators(lngColumnID)
            ReDim Preserve mavFindColumns(3, UBound(mavFindColumns, 2) + 1)
            
            ' Remember the first Find column.
            If lngFirstFindColumnID = 0 Then
              lngFirstFindColumnID = lngColumnID
            End If
          
            iCount2 = iCount2 + 1
          Else
            ' Add the column to the order string.
            sOrderString = sOrderString & _
              IIf(Len(sOrderString) > 0, ", ", "") & _
              sRealSource & "." & Trim(sColumnName) & _
              IIf(fAscending, "", " DESC")
            
            ' Remember the first Order column.
            If lngFirstSortColumnID = 0 Then
              lngFirstSortColumnID = lngColumnID
              mfFirstColumnAscending = fAscending
              miFirstColumnDataType = iDataType
            End If
          End If
          
          ' If the column comes from a parent table, then add the table to the Join code.
          If lngTableID <> mobjLinkTable.TableID Then
            ' Check if the table has already been added to the join code.
            fFound = False
            For iNextIndex = 1 To UBound(alngTableViews, 2)
              If alngTableViews(1, iNextIndex) = 0 And _
                alngTableViews(2, iNextIndex) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next iNextIndex
            
            If Not fFound Then
              ' The table has not yet been added to the join code, so add it to the array and the join code.
              iNextIndex = UBound(alngTableViews, 2) + 1
              ReDim Preserve alngTableViews(2, iNextIndex)
              alngTableViews(1, iNextIndex) = 0
              alngTableViews(2, iNextIndex) = lngTableID
              
              sJoinCode = sJoinCode & _
                " LEFT OUTER JOIN " & sRealSource & _
                " ON " & CurrentTableViewName & ".ID_" & Trim(Str(lngTableID)) & _
                " = " & sRealSource & ".ID"
            End If
          End If
        Else
          ' The column cannot be read from the base table/view, or directly from a parent table.
          ' If it is a column from a prent table, then try to read it from the views on the parent table.
          If lngTableID <> mobjLinkTable.TableID Then
            ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
            ReDim asViews(0)
            For Each objTableView In gcoTablePrivileges.Collection
              If (Not objTableView.IsTable) And _
                (objTableView.TableID = lngTableID) And _
                (objTableView.AllowSelect) Then
                
                sSource = objTableView.ViewName
                sRealSource = gcoTablePrivileges.Item(sSource).RealSource
  
                ' Get the column permission for the view.
                Set objColumnPrivileges = GetColumnPrivileges(sSource)
  
                If objColumnPrivileges.IsValid(sColumnName) Then
                  If objColumnPrivileges.Item(sColumnName).AllowSelect Then
                    ' Add the view info to an array to be put into the column list or order code below.
                    iNextIndex = UBound(asViews) + 1
                    ReDim Preserve asViews(iNextIndex)
                    asViews(iNextIndex) = objTableView.ViewName
                    
                    ' Add the view to the Join code.
                    ' Check if the view has already been added to the join code.
                    fFound = False
                    For iNextIndex = 1 To UBound(alngTableViews, 2)
                      If alngTableViews(1, iNextIndex) = 1 And _
                        alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                        fFound = True
                        Exit For
                      End If
                    Next iNextIndex
            
                    If Not fFound Then
                      ' The view has not yet been added to the join code, so add it to the array and the join code.
                      iNextIndex = UBound(alngTableViews, 2) + 1
                      ReDim Preserve alngTableViews(2, iNextIndex)
                      alngTableViews(1, iNextIndex) = 1
                      alngTableViews(2, iNextIndex) = objTableView.ViewID
            
                      sJoinCode = sJoinCode & _
                        " LEFT OUTER JOIN " & sRealSource & _
                        " ON " & CurrentTableViewName & ".ID_" & Trim(Str(objTableView.TableID)) & _
                        " = " & sRealSource & ".ID"
                    End If
                  End If
                End If
                Set objColumnPrivileges = Nothing
  
              End If
            Next objTableView
            Set objTableView = Nothing
          
            ' The current user does have permission to 'read' the column through a/some view(s) on the
            ' table.
            If UBound(asViews) = 0 Then
              fNoSelect = True
            Else
              ' Add the column to the column list.
              sColumnCode = ""
              For iNextIndex = 1 To UBound(asViews)
                If iNextIndex = 1 Then
                  sColumnCode = "CASE "
                End If
                  
                sColumnCode = sColumnCode & _
                  " WHEN NOT " & asViews(iNextIndex) & "." & sColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & sColumnName
              Next iNextIndex
                
              If Len(sColumnCode) > 0 Then
                sColumnCode = sColumnCode & _
                  " ELSE NULL" & _
                  " END AS " & _
                  IIf(sType = "F", "", "'?") & _
                  sColumnName & _
                  IIf(sType = "F", "", "'")
                  
                sColumnList = sColumnList & _
                  IIf(Len(sColumnList) > 0, ", ", "") & _
                  sColumnCode
  
                If sType = "F" Then
                  ' Remember the first Find column.
                  If lngFirstFindColumnID = 0 Then
                    lngFirstFindColumnID = lngColumnID
                  End If
                
                  iCount2 = iCount2 + 1
                  
                  ReDim Preserve mavFindColumns(3, UBound(mavFindColumns, 2) + 1)
                  mavFindColumns(0, UBound(mavFindColumns, 2)) = lngColumnID
                  mavFindColumns(1, UBound(mavFindColumns, 2)) = datGeneral.GetDataSize(lngColumnID)
                  mavFindColumns(2, UBound(mavFindColumns, 2)) = datGeneral.GetDecimalsSize(lngColumnID)
                  mavFindColumns(3, UBound(mavFindColumns, 2)) = datGeneral.DoesColumnUseSeparators(lngColumnID)
                Else
                  ' Add the column to the order string.
                  sOrderString = sOrderString & _
                    IIf(Len(sOrderString) > 0, ", ", "") & _
                    "'?" & Trim(sColumnName) & "'" & _
                    IIf(fAscending, "", " DESC")
  
                  ' Remember the first Order column.
                  If lngFirstSortColumnID = 0 Then
                    lngFirstSortColumnID = lngColumnID
                    mfFirstColumnAscending = fAscending
                    miFirstColumnDataType = iDataType
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
      
      If Not rsInfo.EOF Then
        rsInfo.MoveNext
      End If
    Loop

    ' Inform the user if they do not have permission to see the data.
    If fNoSelect Then
      COAMsgBox "You do not have 'read' permission on all of the columns in the selected order." & _
        vbCrLf & "Only permitted columns will be shown.", vbExclamation, "Security"
    End If
    
    mfFirstColumnsMatch = (lngFirstFindColumnID = lngFirstSortColumnID)
  
    ' Create the string for creating the items that will appear in the listbox.
    If Len(sColumnList) > 0 Then
      sSQL = "SELECT " & sColumnList & ", " & CurrentTableViewName & ".id" & _
        " FROM " & CurrentTableViewName & _
        " " & sJoinCode
        
    
      msGetRecsSQL = "SELECT COUNT(id) FROM " & CurrentTableViewName
      
      If Len(mvLookupSelectedFilterCode) > 0 Then
        'sSQL = sSQL & _
            " WHERE " & sRealSource & "." & Trim(msLookupFilterColumnName) & mvLookupSelectedFilterCode
        sSQL = sSQL & _
            " WHERE " & Replace(mvLookupSelectedFilterCode, vbTab, sRealSource & "." & Trim(msLookupFilterColumnName))
            
        'msGetRecsSQL = msGetRecsSQL & _
            " WHERE " & sRealSource & "." & Trim(msLookupFilterColumnName) & mvLookupSelectedFilterCode
        msGetRecsSQL = msGetRecsSQL & _
            " WHERE " & Replace(mvLookupSelectedFilterCode, vbTab, sRealSource & "." & Trim(msLookupFilterColumnName))
      End If

      sSQL = sSQL & _
        IIf(Len(sOrderString) > 0, " ORDER BY " & sOrderString, "")

      
      ' Get the required recordset.
      Set mrsFindRecords = datGeneral.GetPersistentRecords(sSQL, adOpenStatic, adLockReadOnly)
      

      ' Get the recordset's record count.
      mlngRecordCount = RecordCount
    
      ' Configure the grid.
      ConfigureGrid
    Else
      COAMsgBox "You do not have permission to read any of the columns in the selected order for this " & IIf(mlngViewID > 0, "view.", "table.") & _
        vbCrLf & "Unable to display records.", vbExclamation, "Security"

      'JPD 20050311 Fault 9891
      'mfCancelled = True
      'Me.Hide
      With ssOleDBGridFindColumns
        .Redraw = False
        .Columns.RemoveAll
        .Redraw = True
      End With
      
    End If
  End If
    
  rsInfo.Close
  Set rsInfo = Nothing

End Sub
    
Private Sub ssOleDBGridFindColumns_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

  Dim dblSizeAmendment As Double
  Dim iCount As Integer
  Dim iLastColumn As Integer

  With ssOleDBGridFindColumns

    .Redraw = False

    'Find last visible column
    If .Columns(ColIndex + 1).Visible Then
      dblSizeAmendment = .Columns(ColIndex).Width - .ResizeWidth
      .Columns(ColIndex + 1).Width = .Columns(ColIndex + 1).Width + dblSizeAmendment
      SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(ColIndex + 1).Name, .Columns(ColIndex + 1).Width
    End If

    'Save the resized column width
    SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(ColIndex).Name, .ResizeWidth

    .Redraw = True

  End With


End Sub

Private Sub ssOleDBGridFindColumns_DblClick()
  If ssOleDBGridFindColumns.SelBookmarks.Count > 0 Then
    mfCancelled = False
    mlngLinkRecordID = SelectedRecordID
    
    If mlngLinkRecordID > 0 Then
      Me.Hide
    End If
  End If
  
End Sub

Private Function SelectedRecordID() As Long
  ' Return the ID of the selected reocrd in the grid.
  SelectedRecordID = 0
  mvLookupSelectedValue = ""
  
  If ssOleDBGridFindColumns.SelBookmarks.Count > 0 Then
    If ssOleDBGridFindColumns.Columns((ssOleDBGridFindColumns.Cols - 1)).Value <> "" Then
      SelectedRecordID = ssOleDBGridFindColumns.Columns((ssOleDBGridFindColumns.Cols - 1)).Value
      mvLookupSelectedValue = ssOleDBGridFindColumns.Columns(miLookupColumnGridPosition).Value
    End If
  End If

End Function

Private Sub ssOleDBGridFindColumns_KeyPress(KeyAscii As Integer)
  Dim lngThistime As Long
  Static sFind As String
  Static lngLastTime As Long
  
  Select Case KeyAscii
    Case vbKeyReturn
      ssOleDBGridFindColumns_DblClick
    
    ' Otherwise find the record
    Case Else
      ' Only search for alphanumeric characters.
      If (KeyAscii >= 32) And (KeyAscii <= 255) Then
        lngThistime = GetTickCount
        If lngLastTime + 1500 < lngThistime Then
          sFind = Chr(KeyAscii)
        Else
          sFind = sFind & Chr(KeyAscii)
        End If
        lngLastTime = lngThistime
        LocateRecord sFind
      End If
  End Select

End Sub

Private Sub LocateRecord(psSearchString As String)
  Dim fFound As Boolean
  Dim fUseBinarySearch As Boolean
  Dim iIndex As Long
  Dim iComparisonResult As Integer
  Dim lngLoop As Long
  Dim lngUpper As Long
  Dim lngLower As Long
  Dim lngJump As Long
  Dim lngFirstFindColumn As Long
  Dim lngFirstOrderColumn As Long
  Dim varFoundBookmark As Variant
  Dim varOriginalBookmark As Variant
  
  If ssOleDBGridFindColumns.Rows = 0 Then
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  fUseBinarySearch = mfFirstColumnsMatch
  
  If fUseBinarySearch Then
    If (miFirstColumnDataType <> sqlVarChar) And _
     (miFirstColumnDataType <> sqlVarBinary) And _
     (miFirstColumnDataType <> sqlNumeric) And _
     (miFirstColumnDataType <> sqlInteger) Then
    
      fUseBinarySearch = False
    End If
  End If
  
  ' Search the grid for the required string.
  fFound = False
  
  lngLower = 1
  lngUpper = mlngRecordCount
  
  With ssOleDBGridFindColumns
    .Redraw = False
    
    varOriginalBookmark = .Bookmark
    
    If fUseBinarySearch Then
      ' Binary search the grid for the required string.
      Do
        Select Case miFirstColumnDataType
          Case sqlVarChar, sqlVarBinary
            ' JPD String comparison changed from using VB's strComp function to
            ' using our own DictionaryCompareStrings function. VB's strComp
            ' function does not use the same order as that used when SQL orders
            ' by a character column. The DictionaryCompareStrings does.
            'iComparisonResult = StrComp(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare)
            iComparisonResult = datGeneral.DictionaryCompareStrings(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString))

          Case sqlNumeric, sqlInteger
            If Val(ssOleDBGridFindColumns.Columns(0).Text) = Val(psSearchString) Then
              iComparisonResult = 0
            ElseIf Val(ssOleDBGridFindColumns.Columns(0).Text) < Val(psSearchString) Then
              iComparisonResult = -1
            Else
              iComparisonResult = 1
            End If
        End Select
        
        If Not mfFirstColumnAscending Then
          iComparisonResult = iComparisonResult * -1
        End If
        
        Select Case iComparisonResult
          Case 0    ' String found.
            fFound = True
            varFoundBookmark = .Bookmark
            lngUpper = .Bookmark - 1
            lngJump = -((.Bookmark - lngLower) \ 2) - 1
            If lngLower > lngUpper Then Exit Do
  
          Case -1   ' Current record is before the required record.
            lngLower = .Bookmark + 1
            lngJump = ((lngUpper - .Bookmark) \ 2)
            If lngLower > lngUpper Then Exit Do
                   
          Case 1    ' Current record is after the required record.
            lngUpper = .Bookmark - 1
            lngJump = -((.Bookmark - lngLower) \ 2) - 1
            If lngLower > lngUpper Then Exit Do
        End Select
        
        If lngLower = lngUpper Then
          lngJump = lngUpper - .Bookmark
        End If
        
        ' Move to the middle record of the recmaining records to search.
        .MoveRecords lngJump
      Loop
  
      If fFound Then
        .Bookmark = varFoundBookmark
      Else
        .MoveRecords varOriginalBookmark - .Bookmark
      End If
    Else
      ' Sequential search the grid for the required string.
      .MoveFirst
      For lngLoop = lngLower To lngUpper
        ' JPD String comparison changed from using VB's strComp function to
        ' using our own DictionaryCompareStrings function. VB's strComp
        ' function does not use the same order as that used when SQL orders
        ' by a character column. The DictionaryCompareStrings does.
        'If StrComp(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare) = 0 Then
        If datGeneral.DictionaryCompareStrings(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString)) = 0 Then
          Exit For
        End If
        
        If lngLoop < lngUpper Then
          .MoveNext
        Else
          .Bookmark = varOriginalBookmark
        End If
      Next lngLoop
    End If
    
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
  
    .Redraw = True
  End With
  
  Screen.MousePointer = vbDefault

End Sub


Private Sub ssOleDBGridFindColumns_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsFindRecords.MoveLast
    Else
      mrsFindRecords.MoveFirst
    End If
  Else
    mrsFindRecords.Bookmark = StartLocation
  End If
  
  'JPD 20040803 Fault 9013
  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrsFindRecords.Move NumberOfRowsToMove
  NewLocation = mrsFindRecords.Bookmark

End Sub


Private Sub ssOleDBGridFindColumns_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim strFormat As String
  
  iRowsRead = 0
  
  ' Do nothing if we a re just formatting the grid,
  ' ot if there a re no records to display.
  If (mfFormattingGrid) Or (mlngRecordCount = 0) Then Exit Sub
  
  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsFindRecords.EOF Then
        mrsFindRecords.MoveLast
      End If
    Else
      If Not mrsFindRecords.BOF Then
        mrsFindRecords.MoveFirst
      End If
    End If
  Else
    mrsFindRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsFindRecords.MovePrevious
    Else
      mrsFindRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsFindRecords.BOF Or mrsFindRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsFindRecords.Fields.Count - 1)
          Select Case mrsFindRecords.Fields(iFieldIndex).Type
            Case adDBTimeStamp
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsFindRecords(iFieldIndex), DateFormat)
          
            Case adNumeric
              ' Are thousand separators used
              strFormat = "0"
              If mavFindColumns(3, iFieldIndex) Then
                strFormat = "#,0"
              End If
              If mavFindColumns(2, iFieldIndex) > 0 Then
                strFormat = strFormat & "." & String(mavFindColumns(2, iFieldIndex), "0")
              End If
              
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsFindRecords(iFieldIndex), strFormat)
          
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsFindRecords(iFieldIndex)
          
          End Select
          
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsFindRecords.Bookmark
  
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsFindRecords.Bookmark
  
    End Select
    
    If ReadPriorRows Then
      mrsFindRecords.MovePrevious
    Else
      mrsFindRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub


Private Sub ConfigureViewCombo(lngViewID As Long)
  ' Populate the 'views' combo with the views available for the link table.
  Dim fOK As Boolean
  Dim objTableView As CTablePrivilege
  Dim objLookupColumns As CColumnPrivileges
  Dim iViewStart As Integer
  Dim iLoop As Integer
  Dim fAdded As Boolean
  Dim sCaption As String
  
  ' Add the table to the combo if the user has permission to read it.
  fOK = False
  iViewStart = 0
  If mobjLinkTable.AllowSelect Then
    If mlngLookupColumnID > 0 Then
      Set objLookupColumns = GetColumnPrivileges(mobjLinkTable.TableName)
      If objLookupColumns.IsValid(msLookupColumnName) Then
        If objLookupColumns(msLookupColumnName).AllowSelect Then
          If mlngLookupFilterColumnID > 0 Then
            If objLookupColumns.IsValid(msLookupFilterColumnName) Then
              If objLookupColumns(msLookupFilterColumnName).AllowSelect Then
                fOK = True
              End If
            End If
          Else
            fOK = True
          End If
        End If
      End If
      Set objLookupColumns = Nothing
    Else
      fOK = True
    End If
  End If
  
  If fOK Then
    iViewStart = 1
    cmbView.AddItem RemoveUnderScores(mobjLinkTable.TableName)
    
    'MH20010105 This line put the itemdata to the tableid.  If this tableid
    'matched the default view id then this would be selected instead of the
    'the view with the same id.  This itemdata variable for the table is not
    'really required so I have set it to zero for the moment.
    
    'cmbView.ItemData(cmbView.NewIndex) = mobjLinkTable.TableID
    cmbView.ItemData(cmbView.NewIndex) = 0
  End If
  
  ' Add the table's views to the combo if the user has permission to read them.
  For Each objTableView In gcoTablePrivileges.Collection
    If (Not objTableView.IsTable) And _
      (objTableView.TableID = mobjLinkTable.TableID) Then
      
      fOK = False
      If objTableView.AllowSelect Then
        If mlngLookupColumnID > 0 Then
          Set objLookupColumns = GetColumnPrivileges(objTableView.ViewName)
          If objLookupColumns.IsValid(msLookupColumnName) Then
            If objLookupColumns(msLookupColumnName).AllowSelect Then
              fOK = True
            End If
          End If
          Set objLookupColumns = Nothing
        Else
          fOK = True
        End If
      End If
    
      If fOK Then
        'JPD 20040416 Fault 8499
        fAdded = False
        sCaption = "'" & RemoveUnderScores(Trim(objTableView.ViewName)) & "' view"
        For iLoop = iViewStart To (cmbView.ListCount - 1)
          If UCase(sCaption) < UCase(cmbView.List(iLoop)) Then
            cmbView.AddItem sCaption, iLoop
            cmbView.ItemData(cmbView.NewIndex) = objTableView.ViewID
            
            fAdded = True
            
            'JPD 20040524 Fault 8685
            Exit For
          End If
        Next iLoop
        
        If Not fAdded Then
          cmbView.AddItem sCaption
          cmbView.ItemData(cmbView.NewIndex) = objTableView.ViewID
        End If
      End If
    End If
  Next objTableView
  Set objTableView = Nothing
  
  With cmbView
    ' Select the first view.
    If .ListCount > 0 Then
      '.ListIndex = 0
      .Enabled = True
      SetComboItem cmbView, lngViewID
      If cmbView.ListIndex = -1 Then
        cmbView.ListIndex = 0
      End If
    Else
      .Enabled = False
    End If
  End With

End Sub


Private Function GetViewName(plngViewID As Long) As String
  ' Return the name of the given view.
  Dim sName As String
  Dim objTableView As CTablePrivilege

  Set objTableView = gcoTablePrivileges.FindViewID(plngViewID)
  If Not objTableView Is Nothing Then
    sName = objTableView.ViewName
  Else
    sName = ""
  End If
  
  Set objTableView = Nothing
  
  GetViewName = sName
  
End Function

Private Function CurrentTableViewName() As String
  ' Return the name of the table, or view.
  CurrentTableViewName = IIf(mlngViewID = 0, mobjLinkTable.RealSource, GetViewName(mlngViewID))
  
End Function

Public Function SetCurrentRecord(ByVal plngRecordID As Long) As Boolean

  Dim lngRecordNumber As Long
  Dim vData As Variant
  Dim mvarbkCurrentRecord As Variant
  Dim rsInfo As Recordset
  Dim fFound As Boolean
    
  'Only search for existing records.
  If (plngRecordID > 0) And (ssOleDBGridFindColumns.Rows > 0) Then

    ' Only search if the first column is from this table/view
    Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)
    rsInfo.MoveFirst
  
    If rsInfo.Fields("TableName").Value = mobjLinkTable.TableName And Not Left(mobjLinkTable.RealSource, 8) = "ASRSysCV" Then

      ' This puts us near where we need to be, but not exactly as we could have first columns with the same value
      vData = datGeneral.GetOrderValue(plngRecordID, mrsFindRecords.Fields(0).Name, CurrentTableViewName)
      
      If IsEmpty(vData) Then
        SetCurrentRecord = False
        Exit Function
      End If
      
      lngRecordNumber = datGeneral.GetRecordOrderNumber(vData, mrsFindRecords.Fields(0).Name, CurrentTableViewName, mfFirstColumnAscending, mrsFindRecords.Fields(0).Type)

      ' If record number overshoots, maybe there's a filter applied
      If lngRecordNumber > mrsFindRecords.RecordCount Then
        SetCurrentRecord = False
        Exit Function
      End If
      mrsFindRecords.Move lngRecordNumber, 1
    Else
      mrsFindRecords.MoveFirst
    End If
  
    ' Position on current record
    ' JPD20021104 Fault 4698
    fFound = False
    Do While Not mrsFindRecords.EOF
      If Not mrsFindRecords.Fields("ID") = plngRecordID Then
        'Debug.Print mrsFindRecords.Fields("ID")
        mrsFindRecords.MoveNext
      Else
        fFound = True
        Exit Do
      End If
    Loop
  
    If Not fFound Then
      mrsFindRecords.MoveFirst
    End If
    
    ' Find the bookmark for the desired record
    mvarbkCurrentRecord = mrsFindRecords.Bookmark
    ssOleDBGridFindColumns.MoveRecords mvarbkCurrentRecord
    ssOleDBGridFindColumns.Bookmark = mvarbkCurrentRecord
    
    'Highlight the current selection
    ssOleDBGridFindColumns.FirstRow = ssOleDBGridFindColumns.Bookmark
    ssOleDBGridFindColumns.SelBookmarks.RemoveAll
    ssOleDBGridFindColumns.SelBookmarks.Add ssOleDBGridFindColumns.Bookmark
  
  End If

  SetCurrentRecord = True

End Function
Public Sub SetCurrentRecordValue(ByVal pvValue As Variant)

  Dim mvarbkCurrentRecord As Variant
  Dim fFound As Boolean

  'Only search for existing records.
  If (Len(CStr(pvValue)) > 0) And (ssOleDBGridFindColumns.Rows > 0) Then
    mrsFindRecords.MoveFirst

    ' Position on current record
    fFound = False
    Do While Not mrsFindRecords.EOF
      If Not IsNull(mrsFindRecords.Fields(msLookupColumnName)) Then
        If Not CStr(mrsFindRecords.Fields(msLookupColumnName)) = pvValue Then
          mrsFindRecords.MoveNext
        Else
          fFound = True
          Exit Do
        End If
      Else
        mrsFindRecords.MoveNext
      End If
    Loop

    If Not fFound Then
      mrsFindRecords.MoveFirst
    End If

    ' Find the bookmark for the desired record
    mvarbkCurrentRecord = mrsFindRecords.Bookmark
    ssOleDBGridFindColumns.MoveRecords mvarbkCurrentRecord
    ssOleDBGridFindColumns.Bookmark = mvarbkCurrentRecord

    'Highlight the current selection
    ssOleDBGridFindColumns.FirstRow = ssOleDBGridFindColumns.Bookmark
    ssOleDBGridFindColumns.SelBookmarks.RemoveAll
    ssOleDBGridFindColumns.SelBookmarks.Add ssOleDBGridFindColumns.Bookmark
  End If

End Sub


