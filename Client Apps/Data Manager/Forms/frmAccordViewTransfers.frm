VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAccordViewTransfers 
   Caption         =   "Payroll Transfers"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1146
   Icon            =   "frmAccordViewTransfers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4665
      Left            =   11200
      TabIndex        =   11
      Top             =   1110
      Width           =   1245
      Begin VB.CommandButton cmdUnblock 
         Caption         =   "&Unblock"
         Height          =   420
         Left            =   0
         TabIndex        =   15
         Top             =   1575
         Width           =   1245
      End
      Begin VB.CommandButton cmdBlock 
         Caption         =   "&Block"
         Height          =   420
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   420
         Left            =   0
         TabIndex        =   13
         Top             =   4200
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Details..."
         Height          =   420
         Left            =   0
         TabIndex        =   12
         Top             =   90
         Width           =   1245
      End
   End
   Begin VB.Frame fraFilters 
      Caption         =   "Filters :"
      Height          =   990
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   12405
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmAccordViewTransfers.frx":000C
         Left            =   8300
         List            =   "frmAccordViewTransfers.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   495
         Width           =   1200
      End
      Begin VB.ComboBox cboUsers 
         Height          =   315
         Left            =   4000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   500
         Width           =   2280
      End
      Begin VB.ComboBox cboTransfer 
         Height          =   315
         ItemData        =   "frmAccordViewTransfers.frx":0010
         Left            =   150
         List            =   "frmAccordViewTransfers.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   500
         Width           =   2550
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   9700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   500
         Width           =   2500
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   8300
         TabIndex        =   5
         Top             =   250
         Width           =   555
      End
      Begin VB.Label lblUsers 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         Height          =   195
         Left            =   4000
         TabIndex        =   3
         Top             =   250
         Width           =   570
      End
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer :"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   250
         Width           =   810
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Left            =   9700
         TabIndex        =   7
         Top             =   250
         Width           =   705
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdTransferDetails 
      Height          =   4530
      Left            =   90
      TabIndex        =   9
      Top             =   1200
      Width           =   10950
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   0
      ForeColorEven   =   -2147483640
      ForeColorOdd    =   -2147483640
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   2593
      Columns(0).Caption=   "Transfer"
      Columns(0).Name =   "Transfer Type"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2196
      Columns(1).Caption=   "Company Code"
      Columns(1).Name =   "Company Code"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2540
      Columns(2).Caption=   "Employee Code"
      Columns(2).Name =   "Employee Code"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1720
      Columns(3).Caption=   "Type"
      Columns(3).Name =   "Transaction Type"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2963
      Columns(4).Caption=   "Status"
      Columns(4).Name =   "Status"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "TransactionID"
      Columns(5).Name =   "TransactionID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   3
      Columns(5).FieldLen=   256
      Columns(6).Width=   3307
      Columns(6).Caption=   "User"
      Columns(6).Name =   "CreatedBy"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1905
      Columns(7).Caption=   "Date"
      Columns(7).Name =   "Created"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   10
      Columns(8).Width=   1588
      Columns(8).Caption=   "Archived"
      Columns(8).Name =   "Archived"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   7
      TabNavigation   =   1
      _ExtentX        =   19315
      _ExtentY        =   7990
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   5910
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   21696
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar abTransactions 
      Left            =   9765
      Top             =   5445
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmAccordViewTransfers.frx":0014
   End
End
Attribute VB_Name = "frmAccordViewTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbLoading As Boolean
Private miConnectionType As DataMgr.AccordConnection
Private mdToDate As Date
Private mdFromDate As Date
Private miTransferType As Integer
Private mbAllTransferTypes As Boolean
Private datData As clsDataAccess
Private mstrOrderBy As String
Private mstrOrderOrder As String
Private mintSortColumnIndex As Integer
Private mlngCurrentRecordID As Long
Private mrsRecords As Recordset
Private mbVisibleBlocking As Boolean
Private mbEnableBlocking As Boolean
Private msDateFormat As String
Private miViewMode As AccordViewMode

Const MaxTop = 100000

Public Property Let ConnectionType(ByVal piNewValue As DataMgr.AccordConnection)
  miConnectionType = piNewValue
End Property

Public Property Let TransferType(ByVal piNewValue As Integer)
  miTransferType = piNewValue
  mbAllTransferTypes = False
End Property

Public Property Let FromDate(ByVal pdValue As Date)
  mdFromDate = pdValue
End Property

Public Property Let ToDate(ByVal pdValue As Date)
  mdToDate = pdValue
End Property

Private Sub grdTransferDetails_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)

  If IsNull(StartLocation) Then
    StartLocation = 0
  End If
  NewLocation = CLng(StartLocation) + NumberOfRowsToMove
  
'  If IsNull(StartLocation) Then
'    If NumberOfRowsToMove = 0 Then
'      Exit Sub
'    ElseIf NumberOfRowsToMove < 0 Then
'      mrsRecords.MoveLast
'    Else
'      mrsRecords.MoveFirst
'    End If
'  Else
'    mrsRecords.Bookmark = StartLocation
'    NewLocation = mrsRecords.Bookmark
'  End If
'
'  mrsRecords.Move NumberOfRowsToMove
'  NewLocation = mrsRecords.Bookmark
  
End Sub

Private Sub grdTransferDetails_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

 ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer

  iRowsRead = 0
    
  ' This is required as recordset not set when this sub is first run
  If mrsRecords Is Nothing Then Exit Sub
  
  ' Do nothing if there are no records to display
  If mrsRecords.RecordCount < 1 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsRecords.EOF Then
        mrsRecords.MoveLast
      End If
    Else
      If Not mrsRecords.BOF Then
        mrsRecords.MoveFirst
      End If
    End If
  Else
    mrsRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsRecords.BOF Or mrsRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsRecords.Fields.Count - 1)
          Select Case mrsRecords.Fields(iFieldIndex).Name
            Case "ID"
              RowBuf.value(iRowIndex, iFieldIndex) = CStr(mrsRecords.Fields("ID"))
            Case "TransferType"
              RowBuf.value(iRowIndex, iFieldIndex) = GetComboText(cboTransfer, mrsRecords.Fields("TransferType").value)
            Case "Status"
              RowBuf.value(iRowIndex, iFieldIndex) = GetComboText(cboStatus, mrsRecords.Fields("Status").value)
            Case "TransActionType"
              Select Case mrsRecords.Fields("TransActionType").value
              Case 0
                RowBuf.value(iRowIndex, iFieldIndex) = "New"
              Case 1
                RowBuf.value(iRowIndex, iFieldIndex) = "Update"
              Case 2
                RowBuf.value(iRowIndex, iFieldIndex) = "Delete"
              End Select
            Case "CreatedDateTime"
              RowBuf.value(iRowIndex, iFieldIndex) = Format(mrsRecords.Fields("CreatedDateTime").value, msDateFormat)
            Case "Archived"
              RowBuf.value(iRowIndex, iFieldIndex) = IIf(mrsRecords.Fields("Archived").value = True, "Yes", "No")
            Case Else
              RowBuf.value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
    End Select
    
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead
  
End Sub

Private Sub RefreshGrid()

  Dim strSQL As String
  Dim mp As Variant
  
  mp = Screen.MousePointer
  Screen.MousePointer = vbHourglass
  
  strSQL = "SELECT TOP " & MaxTop & " TransferType, CompanyCode, EmployeeCode, TransActionType, Status, TransactionID, CreatedUser, CreatedDateTime, Archived" _
          & " FROM ASRSysAccordTransactions" _

  ' Filter on the transfer
  If cboTransfer.Text <> "<All>" Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "TransferType = " & CStr(cboTransfer.ItemData(cboTransfer.ListIndex))
  End If
  
  ' Filter on the type
  If cboType.Text <> "<All>" Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "TransactionType = " & CStr(cboType.ItemData(cboType.ListIndex))
  End If
  
  ' Filter on the status
  If cboStatus.Text <> "<All>" Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "Status = " & CStr(cboStatus.ItemData(cboStatus.ListIndex))
  End If
  
  ' Filter on the users
  If cboUsers.Text <> "<All>" Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "CreatedUser = '" & Replace(cboUsers.Text, "'", "''") & "'"
  End If
  
  ' Filter out the archived records
  If miViewMode = iARCHIVE_ALL Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "Archived = 1"
  ElseIf miViewMode = iLIVE_ALL Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "Archived = 0"
  End If
  
  ' Filter on the current record
  If miViewMode = iCURRENT_RECORD Then
    strSQL = strSQL & IIf(InStr(strSQL, "WHERE") > 0, " AND ", " WHERE ") & "(HRProRecordID = " & mlngCurrentRecordID & " AND TransferType = " & miTransferType & ")"
  End If
  
  ' Orderificationalize the list (George Bushism #1)
  strSQL = strSQL & " ORDER BY " & mstrOrderBy & " " & mstrOrderOrder
   
  If Not mrsRecords Is Nothing Then
    If mrsRecords.State <> 0 Then
      mrsRecords.Close
    End If
    Set mrsRecords = Nothing
  End If
    
  Set mrsRecords = OpenAccordRecordset(miConnectionType, strSQL, adOpenStatic, adLockReadOnly)
  
   'NHRD13092006 Fault 10065
  If miViewMode <> iCURRENT_RECORD Then
    'Resize the column width to fit into grid space provided
    Me.grdTransferDetails.Columns("Date").Width = (Me.grdTransferDetails.Columns("Date").Width + 900.23)
    'Not Current Record view so remove Archive column
    Me.grdTransferDetails.Columns("Archived").Visible = False
  Else
    SetComboItem cboTransfer, CLng(miTransferType)
    EnableControl cboTransfer, False
  End If

  With grdTransferDetails
    .Redraw = False
    .Rebind
    .Rows = IIf(mrsRecords.RecordCount = -1, 0, mrsRecords.RecordCount)
    .Redraw = True
  End With

  If grdTransferDetails.Rows > 0 Then
    grdTransferDetails.MoveFirst
    grdTransferDetails.SelBookmarks.Add grdTransferDetails.Bookmark
  End If

  RefreshStatusBar
  RefreshButtons
  
  If mrsRecords.RecordCount = MaxTop Then
    Screen.MousePointer = vbDefault
    gobjProgress.CloseProgress
    MsgBox "The search results have been limited to " & Format$(MaxTop, "#,###") & " records.", vbInformation, app.Title
  End If

  Screen.MousePointer = mp

End Sub

Private Sub RefreshStatusBar()

  ' Status bar text
  Dim strStatusBarText As String
  strStatusBarText = ""
  strStatusBarText = strStatusBarText & " " & mrsRecords.RecordCount & " Record" & IIf(mrsRecords.RecordCount > 1 Or mrsRecords.RecordCount = 0, "s", "")
  If mrsRecords.RecordCount > 1 Then
    
    strStatusBarText = strStatusBarText & " sorted by "
    
    Select Case mstrOrderBy
      Case "TransferType"
        strStatusBarText = strStatusBarText & "transfer type"
      Case "CompanyCode"
        strStatusBarText = strStatusBarText & "company code"
      Case "EmployeeCode"
        strStatusBarText = strStatusBarText & "employee code"
      Case "TransactionType"
        strStatusBarText = strStatusBarText & "transaction type"
      Case "Status"
        strStatusBarText = strStatusBarText & "status"
      Case "CreatedUser"
        strStatusBarText = strStatusBarText & "user"
      Case "CreatedDateTime"
        strStatusBarText = strStatusBarText & "created date"
      Case "Archived"
        strStatusBarText = strStatusBarText & "archived"
    End Select
      
    strStatusBarText = strStatusBarText & " in " & IIf(mstrOrderOrder = "ASC", "ascending", "descending")
    strStatusBarText = strStatusBarText & " order"
  
  End If
  
  StatusBar1.SimpleText = strStatusBarText
  
End Sub

Public Function Initialise()
  
  mbLoading = True
   
  mbVisibleBlocking = Not (miViewMode = iARCHIVE_ALL)
  mbEnableBlocking = datGeneral.SystemPermission("ACCORD", "BLOCK")
  msDateFormat = UI.GetSystemDateFormat
  
  Select Case miViewMode
    Case iLIVE_ALL
      Me.Caption = "Payroll Transfers"
    Case iARCHIVE_ALL
      Me.Caption = "Payroll Transfers (Archived)"
    Case iCURRENT_RECORD
      Me.Caption = "Payroll Transfers (Current Record)"
  End Select
   
  With gobjProgress
    '.AviFile = App.Path & "\videos\search.Avi"
    .AVI = dbAccord
    .MainCaption = "Payroll Transfers"
    .Caption = Me.Caption
    .NumberOfBars = 1
    .Bar1Value = 0
    .Bar1MaxValue = 100
    .Bar2Value = 0
    .Bar1Caption = "Loading Transfer Information..."
    .Time = False
    .Cancel = False
    .OpenProgress
  End With
  Screen.MousePointer = vbHourglass
        
  PopulateFilters
  RefreshGrid
  
  Screen.MousePointer = vbDefault
  gobjProgress.CloseProgress
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
  mbLoading = False

End Function

Private Sub cboStatus_Click()
  If Not mbLoading Then
    RefreshGrid
  End If
End Sub

Private Sub cboTransfer_Click()
  If Not mbLoading Then
    RefreshGrid
  End If
End Sub

Private Sub cboType_Click()
  If Not mbLoading Then
    RefreshGrid
  End If
End Sub

Private Sub cboUsers_Click()
  If Not mbLoading Then
    RefreshGrid
  End If
End Sub

Private Sub cmdBlock_Click()
  SetBlockStatus True
End Sub

Private Sub cmdEdit_Click()

  'Dim lngRow As Long
  Dim frmRecord As New DataMgr.frmAccordRecord
  
  grdTransferDetails.Bookmark = grdTransferDetails.SelBookmarks(0)
  'lngRow = grdTransferDetails.AddItemRowIndex(grdTransferDetails.Bookmark)
  
  With frmRecord
    .TransactionID = grdTransferDetails.Columns("TransactionID").Text
    .Initialise
    .Show vbModal
    
     ' Update the grid
    If .Changed Then
     
      With gobjProgress
        '.AviFile = App.Path & "\videos\search.Avi"
        .AVI = dbAccord
        .MainCaption = "Payroll Transfers"
        .Caption = Me.Caption
        .NumberOfBars = 1
        .Bar1Value = 0
        .Bar1MaxValue = 100
        .Bar2Value = 0
        .Bar1Caption = "Refreshing Transfer Information..."
        .Time = False
        .Cancel = False
        .OpenProgress
      End With
      Screen.MousePointer = vbHourglass
        
      RefreshGrid
      
      Screen.MousePointer = vbDefault
      gobjProgress.CloseProgress
      
    End If
    
  End With

End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdUnblock_Click()
  SetBlockStatus False
End Sub

Private Sub Form_Initialize()
  Set datData = New DataMgr.clsDataAccess
  mbAllTransferTypes = True
  mstrOrderBy = "CreatedDateTime"
  mstrOrderOrder = "DESC"
End Sub

Private Sub Form_Activate()
  UI.RemoveClipping
End Sub

Private Sub Form_Load()
  Hook Me.hWnd, 13010, 6720, (Screen.Width - 200), (Screen.Height - 500)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
  
  Set mrsRecords = Nothing
End Sub

Private Sub grdTransferDetails_Click()
  RefreshButtons
End Sub

Public Sub PopulateFilters()

  Dim rstData As ADODB.Recordset

  PopulateAccordTransferTypes cboTransfer, True

  ' The types
  With cboType
    .AddItem "<All>"
    .AddItem "New"
    .ItemData(.NewIndex) = 0
    .AddItem "Update"
    .ItemData(.NewIndex) = 1
    .AddItem "Delete"
    .ItemData(.NewIndex) = 2
    .ListIndex = 0
  End With

  ' Get all of the transfer types (Local only)
  With cboUsers
    .AddItem "<All>"
    Set rstData = OpenAccordRecordset(ACCORD_LOCAL, "SELECT Distinct CreatedUser FROM ASRSysAccordTransactions ORDER BY CreatedUser", adOpenForwardOnly, adLockReadOnly)
    Do Until rstData.EOF
      .AddItem rstData.Fields("CreatedUser").value
      rstData.MoveNext
    Loop
    .ListIndex = 0
  End With

  ' The status's
  With cboStatus
    .AddItem "<All>"
    .ItemData(.NewIndex) = -1
 
    .AddItem "Unknown"
    .ItemData(.NewIndex) = ACCORD_STATUS_UNKNOWN
 
    .AddItem "Pending"
    .ItemData(.NewIndex) = ACCORD_STATUS_PENDING

    .AddItem "Success"
    .ItemData(.NewIndex) = ACCORD_STATUS_SUCCESS

    .AddItem "Success with Warnings"
    .ItemData(.NewIndex) = ACCORD_STATUS_SUCCESS_WARNINGS
    
    .AddItem "Unknown Failure"
    .ItemData(.NewIndex) = ACCORD_STATUS_FAILURE_UNKNOWN
    
    .AddItem "Ignored"
    .ItemData(.NewIndex) = ACCORD_STATUS_IGNORED
    
    .AddItem "Record Already Exists"
    .ItemData(.NewIndex) = ACCORD_STATUS_ALREADY_EXISTS
    
    .AddItem "Record Does Not Exist"
    .ItemData(.NewIndex) = ACCORD_STATUS_DOESNOT_EXIST
    
    .AddItem "More Info Required"
    .ItemData(.NewIndex) = ACCORD_STATUS_MOREINFO_REQUIRED
    
    .AddItem "Blocked"
    .ItemData(.NewIndex) = ACCORD_STATUS_BLOCKED
  
    .AddItem "Void"
    .ItemData(.NewIndex) = ACCORD_STATUS_VOID
  
    If miViewMode = iLIVE_ALL Then
      SetCombo cboStatus, ACCORD_STATUS_PENDING
    Else
      .ListIndex = 0
    End If
  
  End With

  Set rstData = Nothing


End Sub

Private Sub SetBlockStatus(ByVal pbBlock As Boolean)

  Dim strEventIDs  As String
  Dim plngLoop As Long
  Dim strSQL As String
  Dim strSetStatus As String
  Dim strCaption As String
  
  If pbBlock Then
    strSetStatus = Str(ACCORD_STATUS_BLOCKED)
    strCaption = "Blocking"
  Else
    strSetStatus = Str(ACCORD_STATUS_PENDING)
    strCaption = "Unblocking"
  End If

  With gobjProgress
    .NumberOfBars = 1
    .Caption = Me.Caption
    .Bar1Value = 1
    .Time = False
    .Cancel = False
    .Bar1RecordsCaption = strCaption & " transfers"
    .Bar1MaxValue = 100
    .OpenProgress
  End With

  strEventIDs = ""
  For plngLoop = 0 To Me.grdTransferDetails.SelBookmarks.Count - 1
    If Len(strEventIDs) > 0 Then
      strEventIDs = strEventIDs & ","
    End If
    strEventIDs = strEventIDs & grdTransferDetails.Columns("TransactionID").CellValue(Me.grdTransferDetails.SelBookmarks(plngLoop))
  Next plngLoop

  ' Update the pending items
  If Len(strEventIDs) > 0 Then
    strSQL = "UPDATE ASRSysAccordTransactions SET Status = " & strSetStatus & " WHERE TransactionID IN (" & strEventIDs & ")" _
             & " AND (Status = " & ACCORD_STATUS_PENDING & "  OR Status = " & ACCORD_STATUS_BLOCKED & ")"
    ExecuteAccordSql miConnectionType, strSQL

    RefreshGrid
  End If

  gobjProgress.CloseProgress

End Sub

Private Sub grdTransferDetails_DblClick()

  If Not ((grdTransferDetails.SelBookmarks.Count > 1) Or (grdTransferDetails.Rows = 0)) Then
    cmdEdit_Click
  End If
  
End Sub

Private Sub grdTransferDetails_HeadClick(ByVal ColIndex As Integer)

  Select Case ColIndex
    Case 0: mstrOrderBy = "TransferType"
    Case 1: mstrOrderBy = "CompanyCode"
    Case 2: mstrOrderBy = "EmployeeCode"
    Case 3: mstrOrderBy = "TransactionType"
    Case 4: mstrOrderBy = "Status"
    Case 5: mstrOrderBy = mstrOrderBy       ' Do nothing - hidden column
    Case 6: mstrOrderBy = "CreatedUser"
    Case 7: mstrOrderBy = "CreatedDateTime"
  End Select

  If ColIndex = mintSortColumnIndex Then
    If mstrOrderOrder = "ASC" Then mstrOrderOrder = "DESC" Else mstrOrderOrder = "ASC"
  End If
  
  mintSortColumnIndex = ColIndex

  RefreshGrid

End Sub

Private Sub RefreshButtons()

  cmdEdit.Enabled = Not ((grdTransferDetails.SelBookmarks.Count > 1) Or (grdTransferDetails.Rows = 0))
  
  If mbVisibleBlocking Then
    If grdTransferDetails.Rows = 0 Then
      cmdBlock.Enabled = False
      cmdUnblock.Enabled = False
    Else
      If Me.grdTransferDetails.SelBookmarks.Count > 1 Then
        cmdBlock.Enabled = mbEnableBlocking
        cmdUnblock.Enabled = mbEnableBlocking
      Else
        cmdUnblock.Enabled = (grdTransferDetails.Columns("Status").Text = GetComboText(cboStatus, ACCORD_STATUS_BLOCKED)) And mbEnableBlocking
        cmdBlock.Enabled = (grdTransferDetails.Columns("Status").Text = GetComboText(cboStatus, ACCORD_STATUS_PENDING)) And mbEnableBlocking
      End If
    End If
  Else
    cmdBlock.Visible = False
    cmdUnblock.Visible = False
  End If
  
End Sub
'
'Private Sub grdTransferDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' ' Enable/disable the required tools.
' If (Button = vbRightButton) And (Y > Me.grdTransferDetails.RowHeight) Then
'    With Me.abTransactions.Bands("popAccord")
'      .Tools("ID_Block").Enabled = Me.cmdBlock.Enabled
'      .Tools("ID_Unblock").Enabled = Me.cmdUnblock.Enabled
'      .Tools("ID_Details").Enabled = Me.cmdEdit.Enabled
'      .TrackPopup -1, -1
'    End With
'  End If
'
'End Sub

Private Sub grdTransferDetails_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  If grdTransferDetails.SelBookmarks.Count = 1 Then
    grdTransferDetails.SelBookmarks.RemoveAll
    grdTransferDetails.SelBookmarks.Add grdTransferDetails.Bookmark
  End If
  
  grdTransferDetails_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case vbKeyEscape
    Unload Me
  Case KeyCode = vbKeyF5
    RefreshGrid
End Select
 
End Sub

Public Property Let ViewMode(ByVal piNewValue As AccordViewMode)
  miViewMode = piNewValue
End Property

Public Property Let CurrentRecordID(ByVal plngNewValue As Long)
  mlngCurrentRecordID = plngNewValue
End Property

' Gets the value from a combo box without actually setting it.
Public Function GetComboText(cboCombo As ComboBox, lItem As Long) As String

  Dim lCount As Long
  
  With cboCombo
    For lCount = 0 To .ListCount - 1
      If .ItemData(lCount) = lItem Then
        GetComboText = cboCombo.List(lCount)
        Exit For
      End If
    Next
  End With

End Function

Public Function SetCombo(cboCombo As ComboBox, lItem As Long)

  Dim lCount As Long
  
  With cboCombo
    For lCount = 0 To .ListCount - 1
      If .ItemData(lCount) = lItem Then
        cboCombo.ListIndex = lCount
        Exit For
      End If
    Next
  End With

End Function

Private Sub Form_Resize()

  Const COMBO_GAP As Integer = 170
  
  Dim lngComboWidth As Long
  
  Const lngGap As Long = 100
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.Form_Resize()"

  'JPD 20030908 Fault 5756
  DisplayApplication
  
  ' Ensure form does not get too small/big. Also reposition controls as necessary
'  UI.ClipForForm Me, 5550, 11700
'  If Me.Width < 13010 Then Me.Width = 13010
'  If Me.Width > Screen.Width Then Me.Width = (Screen.Width - 200)
'  If Me.Height < 6720 Then Me.Height = 6720
'  If Me.Height > Screen.Height Then Me.Height = (Screen.Height - 500)
  
  fraButtons.Left = Me.ScaleWidth - (fraButtons.Width + lngGap)
  fraButtons.Height = (Me.ScaleHeight - (fraFilters.Height + (lngGap * 2))) - StatusBar1.Height
  cmdOK.Top = fraButtons.Height - cmdOK.Height
  
  fraFilters.Width = Me.ScaleWidth - (lngGap * 2)
  
  cboStatus.Left = fraFilters.Width - (cboStatus.Width + COMBO_GAP)
  lblStatus.Left = cboStatus.Left
  
  cboType.Left = cboStatus.Left - (cboType.Width + COMBO_GAP)
  lblType.Left = cboType.Left
  
  
  lngComboWidth = (cboType.Left - (COMBO_GAP * 3)) / 3
  
  cboTransfer.Move COMBO_GAP, 500, lngComboWidth * 2
  lblTransfer.Left = cboTransfer.Left
  
  cboUsers.Move cboTransfer.Left + cboTransfer.Width + COMBO_GAP, 500, lngComboWidth
  lblUsers.Left = cboUsers.Left
  
  
  grdTransferDetails.Width = fraButtons.Left - (grdTransferDetails.Left + lngGap)
  grdTransferDetails.Height = Me.ScaleHeight - (fraFilters.Height + StatusBar1.Height + (lngGap * 3))
  
  DoColumnSizes
  
  Me.Refresh
  Me.grdTransferDetails.Refresh
   
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub DoColumnSizes()

  Dim lngAvailableWidth As Long
    
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.DoColumnSizes()"

  With grdTransferDetails
  
    If miViewMode = iCURRENT_RECORD Then
      lngAvailableWidth = grdTransferDetails.Width - (.Columns(1).Width + .Columns(3).Width + .Columns(4).Width + .Columns(6).Width + .Columns(7).Width + .Columns(8).Width)
    Else
      lngAvailableWidth = grdTransferDetails.Width - (.Columns(1).Width + .Columns(3).Width + .Columns(4).Width + .Columns(6).Width + .Columns(7).Width)
    End If
    
    .Columns(0).Width = (lngAvailableWidth * 0.68)
    .Columns(2).Width = (lngAvailableWidth * 0.32)
  End With
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


