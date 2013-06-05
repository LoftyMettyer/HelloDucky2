VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmOutlookQueue 
   Caption         =   "Outlook Calendar Queue"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1143
   Icon            =   "frmOutlookQueue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   4845
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   19817
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdOutlookQueue 
      Height          =   3645
      Left            =   90
      TabIndex        =   3
      Top             =   1065
      Width           =   10005
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
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      CaptionAlignment=   0
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Caption=   "Table Name"
      Columns(0).Name =   "Table Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Link Title"
      Columns(1).Name =   "Title"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Subject"
      Columns(2).Name =   "Subject"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Folder Name"
      Columns(3).Name =   "Folder Name"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Folder Location"
      Columns(4).Name =   "Folder Location"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Start Date"
      Columns(5).Name =   "Start Date"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "End Date"
      Columns(6).Name =   "End Date"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "Refresh Date"
      Columns(7).Name =   "Refresh Date"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "Status"
      Columns(8).Name =   "Status"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Error Message"
      Columns(9).Name =   "Error Message"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17648
      _ExtentY        =   6429
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   10260
      TabIndex        =   4
      Top             =   180
      Width           =   1200
   End
   Begin VB.Frame fraFilters 
      Caption         =   "Filters :"
      Height          =   825
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   10020
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmOutlookQueue.frx":000C
         Left            =   8160
         List            =   "frmOutlookQueue.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1650
      End
      Begin VB.ComboBox cboTitle 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   1905
      End
      Begin VB.ComboBox cboFolderLocation 
         Height          =   315
         Left            =   4470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2790
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Left            =   7440
         TabIndex        =   9
         Top             =   375
         Width           =   570
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Link Title :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   375
         Width           =   900
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Name :"
         Height          =   195
         Left            =   3165
         TabIndex        =   6
         Top             =   375
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmOutlookQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare sizing constants
Const BUTTON_GAP = 240
Const BUTTON_WIDTH = 1200
Const GAP_AFTER_LISTVIEW = 1850

' Flag to prevent grid refreshing when combos are being populated and set initially
Private mblnLoading As Boolean

' Flag showing if user is viewing all or just his own entries
Private mblnViewAllEntries As Boolean

' Must be public so the details form can change the bookmark of the recordset
Public mrstHeaders As Recordset

' Data access class
Private mclsData As New clsDataAccess

' Variables to hold the column clicked on, its field and the order to sort the grid
Private pstrOrderField As String
Private pstrOrderDisplay As String
Private pstrOrderOrder As String
Private mintSortColumnIndex As Integer

Private mblnDeleteEnabled As Boolean
Private mblnPurgeEnabled As Boolean

Private mblnSentLoaded As Boolean
Private mblnNotSentLoaded As Boolean

Private Sub cboFolderLocation_Click()
  RefreshGrid
End Sub

Private Sub cboStatus_Click()

  If cboStatus.ListIndex <> 1 Then
    mblnNotSentLoaded = False
  End If
  
  If cboStatus.ListIndex <> 2 Then
    mblnSentLoaded = False
  End If
  
  RefreshGrid
End Sub

Private Sub cboTitle_Click()
  RefreshGrid
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Activate()

  DoColumnSizes

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyEscape Then
    Unload Me
  ElseIf KeyCode = vbKeyF5 Then
    RefreshGrid
  End If

End Sub

Private Sub Form_Load()
  
  Hook Me.hWnd, 11700, 5550
  
  Dim rstUsers As Recordset
  Set mclsData = New clsDataAccess
  
  mblnLoading = True

  'This will update all of the record descriptions in the outlook queue
  'mclsData.ExecuteSql "spASRoutlookQueue"

  'If datGeneral.SystemPermission("OUTLOOK", "REBUILDPURGE") = False Then
  '  cmdPurge.Enabled = False
  '  cmdRebuild.Enabled = False
  'End If


  ' Set height and width to last saved. Form is centred on screen
  Me.Height = GetPCSetting(gsDatabaseName & "\OutlookQueue", "Height", Me.Height)
  Me.Width = GetPCSetting(gsDatabaseName & "\OutlookQueue", "Width", Me.Width)

  ' Set default sort order to be date desc
  pstrOrderField = "RefreshDate"
  pstrOrderDisplay = "'Refresh Date'"
  mintSortColumnIndex = 1
  pstrOrderOrder = "DESC"

  PopulateCombos

  mblnLoading = False

  ' Populate the grid
  RefreshGrid

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  ' Save the window size ready to recall next time user views the event log
  SavePCSetting gsDatabaseName & "\OutlookQueue", "Height", Me.Height
  SavePCSetting gsDatabaseName & "\OutlookQueue", "Width", Me.Width

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

  ' Ensure form does not get too small/big. Also reposition controls as necessary
'  If Me.Width < 11700 Or Me.Width > Screen.Width Then
'    Me.Width = 11700
'  End If
'  If Me.Height < 5550 Or Me.Height > Screen.Height Then
'    Me.Height = 5550
'  End If

  'AE20071005 Fault #12196
  'cmdOK.Left = Me.Width - (BUTTON_GAP + BUTTON_WIDTH)
  cmdOK.Left = Me.ScaleWidth - (BUTTON_WIDTH + (BUTTON_GAP / 2))
  
  'cmdView.Left = cmdOk.Left
  'cmdRebuild.Left = cmdOK.Left
  'cmdPurge.Left = cmdOK.Left

  fraFilters.Width = cmdOK.Left - BUTTON_GAP
  
  cboStatus.Left = fraFilters.Width - cboStatus.Width - BUTTON_GAP
  lblStatus.Left = cboStatus.Left - (BUTTON_GAP * 3)

  cboFolderLocation.Width = (lblStatus.Left - BUTTON_GAP) - cboFolderLocation.Left

  grdOutlookQueue.Width = fraFilters.Width
  'grdOutlookQueue.Height = Me.Height - GAP_AFTER_LISTVIEW
  grdOutlookQueue.Height = Me.Height - (Me.Height - Me.ScaleHeight) - 1500

  DoColumnSizes
  Me.Refresh
  grdOutlookQueue.Refresh

  End Sub

Private Sub DoColumnSizes()

  Dim lngWidth As Long
  
  With grdOutlookQueue

    If .Rows > .VisibleRows Then
      'NHRD - 10042003 - Fault 742
      '.ScrollBars = ssScrollBarsVertical
      .ScrollBars = ssScrollBarsAutomatic
      lngWidth = .Width - 255
    Else
      '.ScrollBars = ssScrollBarsNone
      .ScrollBars = ssScrollBarsAutomatic
      lngWidth = .Width - 15
    End If

    ' All columns except make up 90% of the grid, therefore, split them into the
    ' following percentages of 90% of the grid...
    '.Columns(0).Width = 1100
    '.Columns(1).Width = 1100
    '.Columns(2).Width = 1100
    .Columns(5).Width = 1440
    .Columns(6).Width = 1440
    .Columns(7).Width = 1440
    .Columns(8).Width = 1100
    .Columns(9).Width = 3000

    lngWidth = (lngWidth - 0) / 5

    .Columns(0).Width = lngWidth
    .Columns(1).Width = lngWidth
    .Columns(2).Width = lngWidth
    .Columns(3).Width = lngWidth
    .Columns(4).Width = lngWidth

  End With

End Sub

Private Function RefreshGrid() As Boolean

  ' Populate the grid using filter/sort criteria as set by the user
  Dim pstrSQL As String
  Dim blnWhere As Boolean
  Dim strWhere As String
    
  If mblnLoading = True Then Exit Function

  Screen.MousePointer = vbHourglass
  
  If (mblnNotSentLoaded = False) And (cboStatus.ListIndex = 1) Then
    pstrOrderOrder = "ASC"
    mblnNotSentLoaded = True
  End If
  
  If (mblnSentLoaded = False) And (cboStatus.ListIndex = 2) Then
    pstrOrderOrder = "DESC"
    mblnSentLoaded = True
  End If
  
  pstrSQL = "SELECT ASRSysTables.TableName, " & vbCrLf & _
                   "ASRSysOutlookLinks.Title, " & vbCrLf & _
                   "ASRSysOutlookEvents.Subject, " & vbCrLf & _
                   "ASRSysOutlookfolders.Name as 'FolderName', " & vbCrLf & _
                   "ASRSysOutlookEvents.Folder, " & vbCrLf & _
                   "ASRSysOutlookEvents.StartDate, " & vbCrLf & _
                   "ASRSysOutlookEvents.EndDate, " & vbCrLf & _
                   "ASRSysOutlookEvents.RefreshDate, " & vbCrLf & _
                   "case (ASRSysOutlookEvents.Refresh | ASRSysOutlookEvents.Deleted) when 0 then " & vbCrLf & _
                   "case isnull(ASRSysOutlookEvents.ErrorMessage,'') when '' then " & vbCrLf & _
                   "'Successful' else 'Failed' end else 'Pending' end as 'Status', " & vbCrLf & _
                   "isnull(ASRSysOutlookEvents.ErrorMessage,'') " & vbCrLf & _
                   "FROM ASRSysOutlookEvents " & vbCrLf & _
                   "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysOutlookEvents.TableID " & vbCrLf & _
                   "JOIN ASRSysOutlookLinks ON ASRSysOutlookLinks.LinkID = ASRSysOutlookEvents.LinkID " & vbCrLf & _
                   "JOIN ASRSysOutlookfolders ON ASRSysOutlookfolders.FolderID = ASRSysOutlookEvents.FolderID"

  blnWhere = False
  With cboFolderLocation
    If .ListIndex > 0 Then
      strWhere = strWhere & _
        IIf(blnWhere, " AND ", " WHERE ") & _
        "ASRSysOutlookEvents.Folder = '" & Replace(.Text, "'", "''") & "'"
      blnWhere = True
    End If
  End With

  With cboTitle
    If .ListIndex > 0 Then
      strWhere = strWhere & _
        IIf(blnWhere, " AND ", " WHERE ") & _
        "ASRSysOutlookEvents.LinkID = " & CStr(.ItemData(.ListIndex))
      blnWhere = True
    End If
  End With

  With cboStatus
    If .ListIndex > 0 Then
      strWhere = strWhere & _
        IIf(blnWhere, " AND ", " WHERE ")
      
      Select Case .ItemData(.ListIndex)
      Case 1
        strWhere = strWhere & "Refresh = 1 OR Deleted = 1"
      Case 2
        strWhere = strWhere & "Refresh = 0 AND Deleted = 0 AND isnull(ErrorMessage,'') <> ''"
      Case 3
        strWhere = strWhere & "Refresh = 0 AND Deleted = 0 AND isnull(ErrorMessage,'') = ''"
      End Select

      blnWhere = True
    End If
  End With
  
  pstrSQL = pstrSQL & strWhere
  pstrSQL = pstrSQL & " ORDER BY " & pstrOrderField & " " & pstrOrderOrder
  
  Set mrstHeaders = mclsData.OpenPersistentRecordset(pstrSQL, adOpenKeyset, adLockReadOnly)

  With grdOutlookQueue
    .Redraw = False
    .Rebind
    .Rows = mrstHeaders.RecordCount
    .Redraw = True
  End With

  If grdOutlookQueue.Rows > 0 Then
    grdOutlookQueue.MoveFirst
    grdOutlookQueue.SelBookmarks.Add grdOutlookQueue.Bookmark
  End If

  StatusBar1.SimpleText = " " & mrstHeaders.RecordCount & " Record" & IIf(mrstHeaders.RecordCount <> 1, "s", "") & _
    IIf(mrstHeaders.RecordCount > 1, " Sorted by " & pstrOrderDisplay & " in " & IIf(pstrOrderOrder = "ASC", "Ascending", "Descending") & " order", "")

  DoColumnSizes

  Screen.MousePointer = vbDefault

End Function

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdOutlookQueue_HeadClick(ByVal ColIndex As Integer)

  ' Set the sort criteria depending on the column header clicked and refresh the grid
  Select Case ColIndex
    Case 0: pstrOrderField = "TableName": pstrOrderDisplay = "'Table Name'"
    Case 1: pstrOrderField = "Title": pstrOrderDisplay = "'Title'"
    Case 2: pstrOrderField = "ASRSysOutlookEvents.Subject": pstrOrderDisplay = "'Subject'"
    Case 3: pstrOrderField = "FolderName": pstrOrderDisplay = "'Folder Name'"
    Case 4: pstrOrderField = "ASRSysOutlookEvents.Folder": pstrOrderDisplay = "'Folder Location'"
    Case 5: pstrOrderField = "ASRSysOutlookEvents.StartDate": pstrOrderDisplay = "'Start Date'"
    Case 6: pstrOrderField = "ASRSysOutlookEvents.EndDate": pstrOrderDisplay = "'End Date'"
    Case 7: pstrOrderField = "ASRSysOutlookEvents.RefreshDate": pstrOrderDisplay = "'Refresh Date'"
    Case 8: pstrOrderField = "Status": pstrOrderDisplay = "'Status'"
    Case 9: pstrOrderField = "ASRSysOutlookEvents.ErrorMessage": pstrOrderDisplay = "'Error Message'"
  End Select

  'If pstrOrderOrder = "ASC" Then pstrOrderOrder = "DESC" Else pstrOrderOrder = "ASC"
  If mintSortColumnIndex = ColIndex And pstrOrderOrder = "ASC" Then
    pstrOrderOrder = "DESC"
  Else
    pstrOrderOrder = "ASC"
  End If

  mintSortColumnIndex = ColIndex

  RefreshGrid

End Sub

Private Sub grdOutlookQueue_KeyDown(KeyCode As Integer, Shift As Integer)

'  If KeyCode = 46 And mblnDeleteEnabled = True Then
'    cmdDelete_Click
'  ElseIf KeyCode = 35 And Shift = 2 Then
'    ' ctrl and end pressed
'    'grdOutlookQueue.FirstRow = grdOutlookQueue.Rows - grdOutlookQueue.VisibleRows
'    'grdOutlookQueue.MoveLast
'  ElseIf KeyCode = 36 And Shift = 2 Then
'    ' ctrl and home pressed
'    'grdOutlookQueue.FirstRow = 0
'    'grdOutlookQueue.MoveFirst
'  Else
'    If Shift > 0 Then KeyCode = 0
'  End If

End Sub

Private Sub grdOutlookQueue_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)

  If IsNull(StartLocation) Then
    StartLocation = 0
  End If

  NewLocation = CLng(StartLocation) + NumberOfRowsToMove

End Sub

Private Sub grdOutlookQueue_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String

  sDateFormat = DateFormat

  ' This is required as recordset not set when this sub is first run
  If mrstHeaders Is Nothing Then Exit Sub
  If mrstHeaders.State = adStateClosed Then Exit Sub

  ' Do nothing if we are loading or if there are no records to display
  If mblnLoading = True And mrstHeaders.RecordCount = 0 Then Exit Sub

  If StartLocation < 0 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrstHeaders.EOF Then
        mrstHeaders.MoveLast
      End If
    Else
      If Not mrstHeaders.BOF Then
        mrstHeaders.MoveFirst
      End If
    End If
  Else
    mrstHeaders.Bookmark = StartLocation
    If ReadPriorRows Then
      mrstHeaders.MovePrevious
    Else
      mrstHeaders.MoveNext
    End If
  End If

  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrstHeaders.BOF Or mrstHeaders.EOF Then Exit For

    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrstHeaders.Fields.Count - 1)
          Select Case mrstHeaders.Fields(iFieldIndex).Name
          
          'MH20021105 "ColumnValue is now a string and not a date...
          'Case "ColumnValue", "DateDue"
          Case "StartDate", "EndDate", "RefreshDate"
            RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrstHeaders.Fields(iFieldIndex), sDateFormat & " hh:nn")
          Case Else
            RowBuf.Value(iRowIndex, iFieldIndex) = mrstHeaders(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 2
        RowBuf.Value(iRowIndex, 0) = mrstHeaders(0)
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 3
    End Select

    If ReadPriorRows Then
      mrstHeaders.MovePrevious
    Else
      mrstHeaders.MoveNext
    End If

    iRowsRead = iRowsRead + 1
  Next iRowIndex

  RowBuf.RowCount = iRowsRead

End Sub


Private Function GetRecordDesc(lngExprID As Long, lngRecordID As Long)

  ' Return TRUE if the user has been granted the given permission.
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  If lngExprID < 1 Then
    GetRecordDesc = "Record Description Undefined"
    Exit Function
  End If
  
  
  ' Check if the user can create New instances of the given category.
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.sp_ASRExpr_" & lngExprID
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Result", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngRecordID

    cmADO.Execute

    GetRecordDesc = .Parameters(0).Value
  End With
  Set cmADO = Nothing

Exit Function

LocalErr:
  GetRecordDesc = "Error reading record description" '& vbCr & _
                  "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(mlngRecordDescExprID)
  'fOK = False

End Function


Private Sub PopulateCombos()

  Dim rsTemp As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT DISTINCT " & _
              "ASRSysOutlookLinks.Title as [Title], " & _
              "ASRSysOutlookLinks.LinkID as [LinkID] " & _
           "FROM ASRSysOutlookEvents " & _
           "JOIN ASRSysOutlookLinks ON ASRSysOutlookEvents.LinkID = ASRSysOutlookLinks.LinkID "

  Set rsTemp = mclsData.OpenPersistentRecordset(strSQL, adOpenKeyset, adLockReadOnly)

  With cboTitle
    .Clear
    .AddItem "<All>"
    .ItemData(.NewIndex) = 0
    Do While Not rsTemp.EOF
      .AddItem rsTemp!Title
      .ItemData(.NewIndex) = rsTemp!LinkID
      rsTemp.MoveNext
    Loop
    .ListIndex = 0
  End With


  strSQL = "SELECT DISTINCT Folder FROM ASRSysOutlookEvents " & _
           "WHERE IsNull(Folder,'') <> ''"
  Set rsTemp = mclsData.OpenPersistentRecordset(strSQL, adOpenKeyset, adLockReadOnly)

  With cboFolderLocation
    .Clear
    .AddItem "<All>"
    Do While Not rsTemp.EOF
      .AddItem rsTemp!Folder
      rsTemp.MoveNext
    Loop
    .ListIndex = 0
  End With


  With cboStatus
    .Clear
    .AddItem "<All>"
    .ItemData(.NewIndex) = 0
    .AddItem "Failed"
    .ItemData(.NewIndex) = 2
    .AddItem "Pending"
    .ItemData(.NewIndex) = 1
    .AddItem "Successful"
    .ItemData(.NewIndex) = 3
    .ListIndex = 1  'Default to show failed !
  End With

End Sub

