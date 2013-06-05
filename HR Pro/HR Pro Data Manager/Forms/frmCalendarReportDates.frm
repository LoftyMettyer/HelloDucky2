VERSION 5.00
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmCalendarReportDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar Report Event"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1067
   Icon            =   "frmCalendarReportDates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraEventDesc 
      Caption         =   "Event Description : "
      Height          =   1260
      Left            =   4635
      TabIndex        =   36
      Top             =   3360
      Width           =   4755
      Begin VB.ComboBox cboEventDesc2 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":000C
         Left            =   1680
         List            =   "frmCalendarReportDates.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   705
         Width           =   2955
      End
      Begin VB.ComboBox cboEventDesc1 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":001F
         Left            =   1680
         List            =   "frmCalendarReportDates.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   300
         Width           =   2955
      End
      Begin VB.Label lblEventDesc2 
         Caption         =   "Description 2 : "
         Height          =   255
         Left            =   195
         TabIndex        =   38
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label lblEventDesc1 
         Caption         =   "Description 1 : "
         Height          =   255
         Left            =   195
         TabIndex        =   37
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Frame fraLegend 
      Caption         =   "Key :"
      Height          =   3135
      Left            =   4635
      TabIndex        =   31
      Top             =   120
      Width           =   4755
      Begin VB.OptionButton optCharacter 
         Caption         =   "Charac&ter"
         Height          =   255
         Left            =   200
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optLegendLookup 
         Caption         =   "&Lookup Table"
         Height          =   375
         Left            =   200
         TabIndex        =   13
         Top             =   645
         Width           =   1455
      End
      Begin VB.ComboBox cboLegendTable 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":0030
         Left            =   1680
         List            =   "frmCalendarReportDates.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1800
         Width           =   2955
      End
      Begin VB.TextBox txtCharacter 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         MaxLength       =   2
         TabIndex        =   12
         Top             =   300
         Width           =   510
      End
      Begin VB.ComboBox cboLegendColumn 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":0049
         Left            =   1680
         List            =   "frmCalendarReportDates.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2205
         Width           =   2955
      End
      Begin VB.ComboBox cboLegendCode 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":005A
         Left            =   1680
         List            =   "frmCalendarReportDates.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2610
         Width           =   2955
      End
      Begin VB.ComboBox cboEventType 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":0072
         Left            =   1680
         List            =   "frmCalendarReportDates.frx":0079
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   2955
      End
      Begin COALine.COA_Line ASRLine1 
         Height          =   30
         Left            =   480
         Top             =   1560
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   53
      End
      Begin VB.Label lblLookupColumn 
         Caption         =   "Column : "
         Height          =   255
         Left            =   495
         TabIndex        =   35
         Top             =   2220
         Width           =   1200
      End
      Begin VB.Label lblCalendarCode 
         Caption         =   "Code : "
         Height          =   255
         Left            =   495
         TabIndex        =   34
         Top             =   2625
         Width           =   1200
      End
      Begin VB.Label lblLegendTable 
         Caption         =   "Table : "
         Height          =   255
         Left            =   495
         TabIndex        =   33
         Top             =   1845
         Width           =   1200
      End
      Begin VB.Label lblType 
         Caption         =   "Event Type : "
         Height          =   255
         Left            =   495
         TabIndex        =   32
         Top             =   1125
         Width           =   1140
      End
   End
   Begin VB.Frame fraEvent 
      Caption         =   "Event : "
      Height          =   1575
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   4395
      Begin VB.CommandButton cmdEventFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   3930
         Picture         =   "frmCalendarReportDates.frx":0083
         TabIndex        =   2
         Top             =   1110
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.TextBox txtEventFilter 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1545
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1110
         Width           =   2385
      End
      Begin VB.TextBox txtEventName 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   0
         Top             =   300
         Width           =   2685
      End
      Begin VB.ComboBox cboEventTable 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":00FB
         Left            =   1530
         List            =   "frmCalendarReportDates.frx":0105
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   705
         Width           =   2685
      End
      Begin VB.Label lblEventFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter :"
         Height          =   195
         Left            =   195
         TabIndex        =   40
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label lblEventName 
         Caption         =   "Name :"
         Height          =   255
         Left            =   195
         TabIndex        =   30
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblEventTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Event Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   765
         Width           =   1155
      End
   End
   Begin VB.Frame fraEventEnd 
      Caption         =   "Event End :"
      Height          =   2475
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   4395
      Begin VB.OptionButton optNoEnd 
         Caption         =   "&None"
         Height          =   375
         Left            =   200
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDuration 
         Caption         =   "&Duration "
         Height          =   375
         Left            =   200
         TabIndex        =   9
         Top             =   1875
         Width           =   1110
      End
      Begin VB.OptionButton optEndDate 
         Caption         =   "&End Date"
         Height          =   375
         Left            =   200
         TabIndex        =   6
         Top             =   645
         Width           =   1215
      End
      Begin VB.ComboBox cboDuration 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":011D
         Left            =   1530
         List            =   "frmCalendarReportDates.frx":0124
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1920
         Width           =   2685
      End
      Begin VB.ComboBox cboEndSession 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":0132
         Left            =   1530
         List            =   "frmCalendarReportDates.frx":0139
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1485
         Width           =   2685
      End
      Begin VB.ComboBox cboEndDate 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":0148
         Left            =   1530
         List            =   "frmCalendarReportDates.frx":014F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2685
      End
      Begin VB.Label lblEndSession 
         Caption         =   "Session :"
         Height          =   255
         Left            =   495
         TabIndex        =   27
         Top             =   1545
         Width           =   855
      End
      Begin VB.Label lblEndDate 
         Caption         =   "Date :"
         Height          =   255
         Left            =   495
         TabIndex        =   26
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame fraEventStart 
      Caption         =   "Event Start :"
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   4395
      Begin VB.ComboBox cboStartSession 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":015C
         Left            =   1530
         List            =   "frmCalendarReportDates.frx":0163
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   705
         Width           =   2685
      End
      Begin VB.ComboBox cboStartDate 
         Height          =   315
         ItemData        =   "frmCalendarReportDates.frx":0174
         Left            =   1530
         List            =   "frmCalendarReportDates.frx":017B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   2685
      End
      Begin VB.Label lblStartSession 
         Caption         =   "Start Session :"
         Height          =   255
         Left            =   195
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "Start Date :"
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   320
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6900
      TabIndex        =   20
      Top             =   5200
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8175
      TabIndex        =   21
      Top             =   5200
      Width           =   1200
   End
End
Attribute VB_Name = "frmCalendarReportDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
Private mbCancelled As Boolean
Private mblnLoading As Boolean

Private mlngBaseTableID As Long

'DataAccess Class
Private datData As HRProDataMgr.clsDataAccess

Private mblnNew As Boolean
Private mfrmParent As frmCalendarReport

Private mcolReportEvents As clsCalendarEvents

Private mcolColumnPivilages As CColumnPrivileges

Private mlngEventTableID As Long
Private mlngEventFilterID As Long
Private mlngEventStartDateID As Long
Private mlngEventStartSessionID As Long
Private mlngEventEndDateID As Long
Private mlngEventEndSessionID As Long
Private mlngEventDurationID As Long
Private mintLegendType As Integer
Private mlngLegendType As Long
Private mlngLegendTableID As Long
Private mlngLegendColumnID As Long
Private mlngLegendCodeID As Long
Private mlngLegendEventTypeID As Long
Private mlngEventDesc1ID As Long
Private mlngEventDesc2ID As Long

Private mstrEventKey As String

Private mblnHasStartDate As Boolean
Private mblnHasLookupColumn As Boolean

Private mlngDateColCount As Long

Private Sub GetLookupTableDefaultDetails()

  Dim strSQL As String
  Dim rsColumn As ADODB.Recordset
  
  strSQL = "SELECT * FROM ASRSysColumns WHERE ColumnID = " & CStr(cboEventType.ItemData(cboEventType.ListIndex))
  
  Set rsColumn = datGeneral.GetReadOnlyRecords(strSQL)
  
  With rsColumn
    If .BOF And .EOF Then
      GoTo TidyUpAndExit
      
    Else
      If !ColumnType = ColumnTypes.colLookup Then
        SetComboItem Me.cboLegendTable, CLng(!LookupTableID)
        SetComboItem Me.cboLegendColumn, CLng(!LookupColumnID)
      End If
    End If
  End With
    
TidyUpAndExit:
  rsColumn.Close
  Set rsColumn = Nothing
 
End Sub

Public Function Initialize(pbNew As Boolean, pfrmParentForm As frmCalendarReport, _
                           pcolEvents As clsCalendarEvents, _
                           psEventName As String, plngEventTableID As Long, _
                           plngEventFilterID As Long, psEventFilter As String, _
                           plngStartDateID As Long, plngStartSessionID As Long, _
                           plngEndDateID As Long, plngEndSessionID As Long, _
                           plngDurationID As Long, psCharacterCode As String, _
                           plngLegendTableID As Long, plngLegendColumnID As Long, _
                           plngLegendCodeID As Long, plngEventTypeID As Long, _
                           plngEventDesc1ID As Long, plngEventDesc2ID As Long, _
                           pstrEventKey As String) As Boolean

  On Error GoTo Error_Trap
  
  mblnLoading = True
  
  ' Set references to class modules
  Set datData = New HRProDataMgr.clsDataAccess

  Set mcolReportEvents = pcolEvents
  
  mstrEventKey = pstrEventKey
  
  ' set module level variables
  mblnNew = pbNew
  Set mfrmParent = pfrmParentForm
  mlngBaseTableID = mfrmParent.cboBaseTable.ItemData(mfrmParent.cboBaseTable.ListIndex)
  Me.txtEventName.Text = psEventName
  Me.txtEventFilter.Text = psEventFilter
  Me.txtEventFilter.Tag = plngEventFilterID
  mlngEventTableID = plngEventTableID
  mlngEventFilterID = plngEventFilterID
  mlngEventStartDateID = plngStartDateID
  mlngEventStartSessionID = plngStartSessionID
  mlngEventEndDateID = plngEndDateID
  mlngEventEndSessionID = plngEndSessionID
  mlngEventDurationID = plngDurationID
  Me.txtCharacter.Text = psCharacterCode
  mlngLegendTableID = plngLegendTableID
  mlngLegendColumnID = plngLegendColumnID
  mlngLegendCodeID = plngLegendCodeID
  mlngLegendEventTypeID = plngEventTypeID
  mlngEventDesc1ID = plngEventDesc1ID
  mlngEventDesc2ID = plngEventDesc2ID
  
  If mlngEventEndDateID > 0 Then
    optEndDate.Value = True
  ElseIf mlngEventDurationID > 0 Then
    optDuration.Value = True
  Else
    optNoEnd.Value = True
  End If
  
  If mlngLegendTableID > 0 Then
    optLegendLookup.Value = True
  Else
    optCharacter.Value = True
  End If
  
  If PopulateTables Then
    UpdateEventDependantFields
    RefreshEventFrames
    UpdateLegendDependantFields
    RefreshLegendFrame
  End If
    
  If Me.Cancelled Then
    Initialize = False
    Exit Function
  End If
 
  Me.Changed = False
  mblnLoading = False
  Initialize = True
  
TidyUpAndExit:
  Exit Function
  
Error_Trap:
  COAMsgBox "Error initialising the the Calendar Report Event form.", vbExclamation + vbOKOnly, "Calendar Reports"
  Initialize = False
  GoTo TidyUpAndExit
End Function
Public Property Let Cancelled(bCancelled As Boolean)
  mbCancelled = bCancelled
End Property
Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property
Public Property Get Key() As String
  Key = mstrEventKey
End Property
Public Property Get EventName() As String
  EventName = Me.txtEventName.Text
End Property
Public Property Get EventTableID() As Long
  EventTableID = cboEventTable.ItemData(cboEventTable.ListIndex)
End Property
Public Property Get EventTable() As String
  If cboEventTable.ItemData(cboEventTable.ListIndex) > 0 Then
    EventTable = cboEventTable.List(cboEventTable.ListIndex)
  Else
    EventTable = vbNullString
  End If
End Property
Public Property Get EventFilterID() As Long
  EventFilterID = CLng(txtEventFilter.Tag)
End Property
Public Property Get EventFilterName() As String
  EventFilterName = txtEventFilter.Text
End Property
Public Property Get EventStartDateID() As Long
  EventStartDateID = cboStartDate.ItemData(cboStartDate.ListIndex)
End Property
Public Property Get EventStartSessionID() As Long
  EventStartSessionID = cboStartSession.ItemData(cboStartSession.ListIndex)
End Property
Public Property Get EventEndDateID() As Long
  If optEndDate.Value Then
    EventEndDateID = cboEndDate.ItemData(cboEndDate.ListIndex)
  Else
    EventEndDateID = 0
  End If
End Property
Public Property Get EventEndSessionID() As Long
  If optEndDate.Value Then
    EventEndSessionID = cboEndSession.ItemData(cboEndSession.ListIndex)
  Else
    EventEndSessionID = 0
  End If
End Property
Public Property Get EventDurationID() As Long
  If optDuration.Value Then
    EventDurationID = cboDuration.ItemData(cboDuration.ListIndex)
  Else
    EventDurationID = 0
  End If
End Property
Public Property Get EventLegendRef() As String
  If EventLegendType = 1 Then
    EventLegendRef = cboLegendTable.List(cboLegendTable.ListIndex) & _
                      "." & _
                     cboLegendCode.List(cboLegendCode.ListIndex)
  Else
    EventLegendRef = EventCharacter
  End If
End Property
Public Property Get EventCharacter() As String
  EventCharacter = txtCharacter.Text
End Property

Public Property Get EventLegendTableID() As Long
  If optLegendLookup.Value Then
    EventLegendTableID = cboLegendTable.ItemData(cboLegendTable.ListIndex)
  End If
End Property
Public Property Get EventLegendTable() As String
  If optLegendLookup.Value Then
    If cboLegendTable.ItemData(cboLegendTable.ListIndex) > 0 Then
      EventLegendTable = cboLegendTable.List(cboLegendTable.ListIndex)
    Else
      EventLegendTable = vbNullString
    End If
  End If
End Property
Public Property Get EventLegendType() As Long
  If optLegendLookup.Value Then
    EventLegendType = 1
  Else
    EventLegendType = 0
  End If
End Property
Public Property Get EventLegendColumnID() As Long
  If optLegendLookup.Value Then
    EventLegendColumnID = cboLegendColumn.ItemData(cboLegendColumn.ListIndex)
  Else
    EventLegendColumnID = 0
  End If
End Property
Public Property Get EventLegendCodeID() As Long
  If optLegendLookup.Value Then
    EventLegendCodeID = cboLegendCode.ItemData(cboLegendCode.ListIndex)
  Else
    EventLegendCodeID = 0
  End If
End Property
Public Property Get LegendEventTypeID() As Long
  If optLegendLookup.Value Then
    LegendEventTypeID = cboEventType.ItemData(cboEventType.ListIndex)
  Else
    LegendEventTypeID = 0
  End If
End Property
Public Property Get EventLegendColumn() As String
  If optLegendLookup.Value Then
    EventLegendColumn = cboLegendColumn.List(cboLegendColumn.ListIndex)
  Else
    EventLegendColumn = vbNullString
  End If
End Property
Public Property Get EventLegendCode() As String
  If optLegendLookup.Value Then
    EventLegendCode = cboLegendCode.List(cboLegendCode.ListIndex)
  Else
    EventLegendCode = vbNullString
  End If
End Property
Public Property Get LegendEventType() As String
  If optLegendLookup.Value Then
    LegendEventType = cboEventType.List(cboEventType.ListIndex)
  Else
    LegendEventType = vbNullString
  End If
End Property
Public Property Get EventStartDateColumn() As String
  If cboStartDate.ItemData(cboStartDate.ListIndex) > 0 Then
    EventStartDateColumn = cboStartDate.List(cboStartDate.ListIndex)
  Else
    EventStartDateColumn = vbNullString
  End If
End Property
Public Property Get EventStartSessionColumn() As String
  If cboStartSession.ItemData(cboStartSession.ListIndex) > 0 Then
    EventStartSessionColumn = cboStartSession.List(cboStartSession.ListIndex)
  Else
    EventStartSessionColumn = vbNullString
  End If
End Property
Public Property Get EventEndSessionColumn() As String
  If cboEndSession.ItemData(cboEndSession.ListIndex) > 0 Then
    EventEndSessionColumn = cboEndSession.List(cboEndSession.ListIndex)
  Else
    EventEndSessionColumn = vbNullString
  End If
End Property
Public Property Get EventEndDateColumn() As String
  If cboEndDate.ItemData(cboEndDate.ListIndex) > 0 Then
    EventEndDateColumn = cboEndDate.List(cboEndDate.ListIndex)
  Else
    EventEndDateColumn = vbNullString
  End If
End Property
Public Property Get EventDurationColumn() As String
  If optDuration.Value Then
    EventDurationColumn = cboDuration.List(cboDuration.ListIndex)
  Else
    EventDurationColumn = vbNullString
  End If
End Property
Public Property Get EventDesc1ID() As Long
  EventDesc1ID = cboEventDesc1.ItemData(cboEventDesc1.ListIndex)
End Property
Public Property Get EventDesc1Column() As String
  If cboEventDesc1.ItemData(cboEventDesc1.ListIndex) > 0 Then
    EventDesc1Column = cboEventDesc1.List(cboEventDesc1.ListIndex)
  Else
    EventDesc1Column = vbNullString
  End If
End Property
Public Property Get EventDesc2Column() As String
  If cboEventDesc2.ItemData(cboEventDesc2.ListIndex) > 0 Then
    EventDesc2Column = cboEventDesc2.List(cboEventDesc2.ListIndex)
  Else
    EventDesc2Column = vbNullString
  End If
End Property
Public Property Get EventDesc2ID() As Long
  EventDesc2ID = cboEventDesc2.ItemData(cboEventDesc2.ListIndex)
End Property
Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOK.Enabled = pblnChanged
End Property
Private Sub RefreshLegendFrame()
  
  mblnHasStartDate = (cboStartDate.ListCount > 0)
  mblnHasLookupColumn = (cboEventType.ListCount > 0)
  
  If Me.optCharacter Then
    Me.txtCharacter.Enabled = mblnHasStartDate
    Me.txtCharacter.BackColor = IIf(Me.txtCharacter.Enabled, vbWindowBackground, vbButtonFace)
    
    With cboEventType
      .Enabled = False
      .ListIndex = -1
    End With
    
    With cboLegendTable
      .Enabled = False
      .ListIndex = -1
    End With
    
    With cboLegendColumn
      .Enabled = False
      .ListIndex = -1
    End With
    
    With cboLegendCode
      .Enabled = False
      .ListIndex = -1
    End With
    
  Else
    Me.txtCharacter.Text = ""
    Me.txtCharacter.Enabled = False
    Me.txtCharacter.BackColor = vbButtonFace
    
    With cboLegendTable
      .Enabled = (mblnHasStartDate And mblnHasLookupColumn)
      If (.ListIndex < 0) And (.ListCount > 0) Then
        .ListIndex = 0
      End If
    End With
    
    With cboLegendColumn
      .Enabled = (mblnHasStartDate And mblnHasLookupColumn) And (.ListCount > 0)
      If (.ListIndex < 0) And (.ListCount > 0) Then
        .ListIndex = 0
      End If
    End With
    
    With cboLegendCode
      .Enabled = (mblnHasStartDate And mblnHasLookupColumn) And (.ListCount > 0)
      If (.ListIndex < 0) And (.ListCount > 0) Then
        .ListIndex = 0
      End If
    End With
  
    With cboEventType
      .Enabled = (mblnHasStartDate And mblnHasLookupColumn) And (.ListCount > 0)
      If (.ListIndex < 0) And (.ListCount > 0) Then
        .ListIndex = 0
      End If
    End With
  
  End If
  
  cboLegendTable.BackColor = IIf(Me.cboLegendTable.Enabled, vbWindowBackground, vbButtonFace)
  cboLegendColumn.BackColor = IIf(Me.cboLegendColumn.Enabled, vbWindowBackground, vbButtonFace)
  cboLegendCode.BackColor = IIf(Me.cboLegendCode.Enabled, vbWindowBackground, vbButtonFace)
  cboEventType.BackColor = IIf(Me.cboEventType.Enabled, vbWindowBackground, vbButtonFace)
  
End Sub
Private Sub RefreshEventFrames()
  
  mblnHasStartDate = (cboStartDate.ListCount > 0)
  
  With cboStartDate
    .Enabled = mblnHasStartDate
    If .ListCount > 0 And .ListIndex < 0 Then
      .ListIndex = 0
    End If
  End With
  
  With cboStartSession
    .Enabled = mblnHasStartDate
    If .ListCount > 0 And .ListIndex < 0 Then
      .ListIndex = 0
    End If
  End With
  
  If optNoEnd.Value Then
    With cboEndDate
      .Enabled = False
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
    
    With cboEndSession
      .Enabled = False
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
    
    With cboDuration
      .Enabled = False
    End With
  
  ElseIf optEndDate.Value Then
    With cboEndDate
      .Enabled = mblnHasStartDate
      If .ListCount > 0 And .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End With
    
    With cboEndSession
      .Enabled = mblnHasStartDate
      If .ListCount > 0 And .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End With
    
    With cboDuration
      .Enabled = False
    End With
    
  ElseIf optDuration.Value Then
    With cboEndDate
      .Enabled = False
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
    
    With cboEndSession
      .Enabled = False
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
    
    With cboDuration
      .Enabled = mblnHasStartDate
      If .ListCount > 0 And .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End With
  Else
    With cboEndDate
      .Enabled = False
    End With
    
    With cboEndSession
      .Enabled = False
    End With
    
    With cboDuration
      .Enabled = False
    End With
    
  End If
  
  cboStartDate.BackColor = IIf(Me.cboStartDate.Enabled, vbWindowBackground, vbButtonFace)
  cboStartSession.BackColor = IIf(Me.cboStartSession.Enabled, vbWindowBackground, vbButtonFace)
  cboEndDate.BackColor = IIf(Me.cboEndDate.Enabled, vbWindowBackground, vbButtonFace)
  cboDuration.BackColor = IIf(Me.cboDuration.Enabled, vbWindowBackground, vbButtonFace)
  cboEndSession.BackColor = IIf(Me.cboEndSession.Enabled, vbWindowBackground, vbButtonFace)
  
  With cboEventDesc1
    .Enabled = mblnHasStartDate
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    If .ListCount > 0 And .ListIndex < 0 Then
      .ListIndex = 0
    End If
  End With
  
  With cboEventDesc2
    .Enabled = mblnHasStartDate
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    If .ListCount > 0 And .ListIndex < 0 Then
      .ListIndex = 0
    End If
  End With
  
End Sub
Private Sub UpdateEventDependantFields()

  Dim rsColumns As New ADODB.Recordset
  Dim sSQL As String
  
  ' Clear Start Date combo
  With cboStartDate
    .Clear
  End With

  ' Clear Start Session combo and add <None> entry
  With cboStartSession
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear End Date combo
  With cboEndDate
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear End Session combo and add <None> entry
  With cboEndSession
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear Duration combo
  With cboDuration
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear Event Type combo and add <None> entry
  With cboEventType
    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'    .ListIndex = 0
  End With

  ' Clear Event Description 2 combo and add <None> entry
  With cboEventDesc1
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Clear Event Description 2 combo and add <None> entry
  With cboEventDesc2
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
    .ListIndex = 0
  End With

  ' Get the columns for the selected event table
  sSQL = "SELECT ASRSysColumns.columnID, ASRSysColumns.tableID, ASRSysColumns.columnName, ASRSysColumns.datatype, " & _
    "       ASRSysColumns.columnType, ASRSysColumns.Size, asrsystables.tablename" & _
    " FROM ASRSysColumns INNER JOIN asrsystables on ASRSysColumns.tableid = asrsystables.tableid" & _
    " WHERE (ASRSysColumns.tableID = " & cboEventTable.ItemData(cboEventTable.ListIndex) & _
    "     OR ASRSysColumns.tableID = " & mlngBaseTableID & ")" & _
    " AND ASRSysColumns.columnType <> " & Trim(Str(colSystem)) & _
    " AND ASRSysColumns.columnType <> " & Trim(Str(colLink)) & _
    " AND ASRSysColumns.dataType <> " & Trim(Str(sqlVarBinary)) & _
    " AND ASRSysColumns.dataType <> " & Trim(Str(sqlOle)) & _
    " ORDER BY asrsystables.tablename, ASRSysColumns.columnName"

  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If rsColumns.BOF And rsColumns.EOF Then
    cboStartDate.Enabled = False
    cboStartSession.Enabled = False
    cboEndDate.Enabled = False
    cboEndSession.Enabled = False
    cboDuration.Enabled = False
    cboEventDesc1.Enabled = False
    cboEventDesc2.Enabled = False
    
  Else
    Do While Not rsColumns.EOF
      
      If rsColumns("tableID").Value = cboEventTable.ItemData(cboEventTable.ListIndex) Then
        If rsColumns("DataType") = SQLDataType.sqlDate Then
          'populate the Start Date combo
          cboStartDate.AddItem rsColumns("ColumnName")
          cboStartDate.ItemData(cboStartDate.NewIndex) = rsColumns("ColumnID")
          
          If (mlngEventStartDateID > 0) And (mlngEventStartDateID = rsColumns("ColumnID")) Then
            cboStartDate.ListIndex = cboStartDate.NewIndex
          End If
          
          'populate the End Date combo
          cboEndDate.AddItem rsColumns("ColumnName")
          cboEndDate.ItemData(cboEndDate.NewIndex) = rsColumns("ColumnID")
  
          If (mlngEventEndDateID > 0) And (mlngEventEndDateID = rsColumns("ColumnID")) Then
            cboEndDate.ListIndex = cboEndDate.NewIndex
          End If
        End If
        
        If rsColumns("DataType") = SQLDataType.sqlVarChar And rsColumns("Size") = 2 Then
          'populate the Start Session combo
          cboStartSession.AddItem rsColumns("ColumnName")
          cboStartSession.ItemData(cboStartSession.NewIndex) = rsColumns("ColumnID")
        
          If (mlngEventStartSessionID > 0) And (mlngEventStartSessionID = rsColumns("ColumnID")) Then
            cboStartSession.ListIndex = cboStartSession.NewIndex
          End If
        
          'populate the End Session combo
          cboEndSession.AddItem rsColumns("ColumnName")
          cboEndSession.ItemData(cboEndSession.NewIndex) = rsColumns("ColumnID")
          
          If (mlngEventEndSessionID > 0) And (mlngEventEndSessionID = rsColumns("ColumnID")) Then
            cboEndSession.ListIndex = cboEndSession.NewIndex
          End If
        End If
        
        If (rsColumns("DataType") = SQLDataType.sqlInteger) _
          Or (rsColumns("DataType") = SQLDataType.sqlNumeric) Then
          'populate the Duration combo
          cboDuration.AddItem rsColumns("ColumnName")
          cboDuration.ItemData(cboDuration.NewIndex) = rsColumns("ColumnID")
          
          If (mlngEventDurationID > 0) And (mlngEventDurationID = rsColumns("ColumnID")) Then
            cboDuration.ListIndex = cboDuration.NewIndex
          End If
        End If
  
        'populate the Event Type combo (legend section)
        If rsColumns("columnType") = ColumnTypes.colLookup Then
          cboEventType.AddItem rsColumns("ColumnName")
          cboEventType.ItemData(cboEventType.NewIndex) = rsColumns("ColumnID")
          If (mlngLegendEventTypeID > 0) And (mlngLegendEventTypeID = rsColumns("ColumnID")) Then
            cboEventType.ListIndex = cboEventType.NewIndex
          End If
        End If
        
      End If
      
      'populate the Description 1 combo
      cboEventDesc1.AddItem rsColumns("TableName") & "." & rsColumns("ColumnName")
      cboEventDesc1.ItemData(cboEventDesc1.NewIndex) = rsColumns("ColumnID")
            
      If (mlngEventDesc1ID > 0) And (mlngEventDesc1ID = rsColumns("ColumnID")) Then
        cboEventDesc1.ListIndex = cboEventDesc1.NewIndex
      End If
      
      'populate the Description 2 combo
      cboEventDesc2.AddItem rsColumns("TableName") & "." & rsColumns("ColumnName")
      cboEventDesc2.ItemData(cboEventDesc2.NewIndex) = rsColumns("ColumnID")
     
      If (mlngEventDesc2ID > 0) And (mlngEventDesc2ID = rsColumns("ColumnID")) Then
        cboEventDesc2.ListIndex = cboEventDesc2.NewIndex
      End If
     
      rsColumns.MoveNext
    Loop
    
  End If

  If (cboEventType.ListIndex < 0) And (cboEventType.ListCount > 0) Then
    cboEventType.ListIndex = 0
  ElseIf (cboEventType.ListCount < 1) Then
    optCharacter.Value = True
    'TM22102003 Fault 7347
    'cboLegendTable.Clear
    cboLegendColumn.Clear
    cboLegendCode.Clear
  End If
  
  mblnHasStartDate = (cboStartDate.ListCount > 0)
  mblnHasLookupColumn = (cboEventType.ListCount > 0)
  
  lblEventFilter.Enabled = mblnHasStartDate
  cmdEventFilter.Enabled = mblnHasStartDate
  fraEventStart.Enabled = mblnHasStartDate
  lblStartDate.Enabled = mblnHasStartDate
  lblStartSession.Enabled = mblnHasStartDate
  fraEventEnd.Enabled = mblnHasStartDate
  optNoEnd.Enabled = mblnHasStartDate
  optEndDate.Enabled = mblnHasStartDate
  lblEndDate.Enabled = mblnHasStartDate
  lblEndSession.Enabled = mblnHasStartDate
  optDuration.Enabled = mblnHasStartDate
  fraLegend.Enabled = mblnHasStartDate
  optCharacter.Enabled = mblnHasStartDate
  lblLegendTable.Enabled = (mblnHasStartDate And mblnHasLookupColumn)
  lblLookupColumn.Enabled = (mblnHasStartDate And mblnHasLookupColumn)
  lblCalendarCode.Enabled = (mblnHasStartDate And mblnHasLookupColumn)
  lblType.Enabled = (mblnHasStartDate And mblnHasLookupColumn)
  optLegendLookup.Enabled = (mblnHasStartDate And mblnHasLookupColumn)
  fraEventDesc.Enabled = mblnHasStartDate
  lblEventDesc1.Enabled = mblnHasStartDate
  lblEventDesc2.Enabled = mblnHasStartDate

  rsColumns.Close
  Set rsColumns = Nothing

End Sub
Private Sub UpdateLegendDependantFields()

  Dim rsColumns As New ADODB.Recordset
  Dim sSQL As String
 
'********************************************************************************
' Populate Legend table/column combo boxes
  
  If (cboEventType.ListCount < 1) Or (Not optLegendLookup.Value) Then
    cboEventType.ListIndex = -1
    cboLegendTable.ListIndex = -1
    Exit Sub
  End If
  
  ' Clear legend column combo and add <None> entry
  With cboLegendColumn
    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'    .ListIndex = 0
  End With

  ' Clear legend code combo and add <None> entry
  With cboLegendCode
    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'    .ListIndex = 0
  End With

  ' Get the columns for the selected event table
  sSQL = "SELECT columnID, tableID, columnName, datatype, columnType, Size" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & EventLegendTableID & _
    " AND columnType <> " & Trim(Str(colSystem)) & _
    " AND columnType <> " & Trim(Str(colLink)) & _
    " AND dataType <> " & Trim(Str(sqlVarBinary)) & _
    " AND dataType <> " & Trim(Str(sqlOle)) & _
    " ORDER BY columnName"

  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If rsColumns.BOF And rsColumns.EOF Then
    cboLegendColumn.Enabled = False
    cboLegendCode.Enabled = False
  
  Else
    Do While Not rsColumns.EOF
      If rsColumns("DataType") = SQLDataType.sqlVarChar Then
        'populate the legend column combo
        cboLegendColumn.AddItem rsColumns("ColumnName")
        cboLegendColumn.ItemData(cboLegendColumn.NewIndex) = rsColumns("ColumnID")
        
        If (mlngLegendColumnID > 0) And (mlngLegendColumnID = rsColumns("ColumnID")) Then
          cboLegendColumn.ListIndex = cboLegendColumn.NewIndex
        End If
        
        'populate the legend code combo
        cboLegendCode.AddItem rsColumns("ColumnName")
        cboLegendCode.ItemData(cboLegendCode.NewIndex) = rsColumns("ColumnID")
      
        If (mlngLegendCodeID > 0) And (mlngLegendCodeID = rsColumns("ColumnID")) Then
          cboLegendCode.ListIndex = cboLegendCode.NewIndex
        End If
      End If
      
      rsColumns.MoveNext
    Loop
    
  End If
  rsColumns.Close
  Set rsColumns = Nothing

  If (cboLegendColumn.ListIndex < 0) And (cboLegendColumn.ListCount > 0) Then
    cboLegendColumn.ListIndex = 0
  End If
  If (cboLegendCode.ListIndex < 0) And (cboLegendCode.ListCount > 0) Then
    cboLegendCode.ListIndex = 0
  End If

End Sub
Private Function PopulateTables() As Boolean
  
  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
  
  On Error GoTo Error_Trap
  
  ' Clear Event Table combo
  cboEventTable.Clear

  ' Get the children of the selected base table
  sSQL = "SELECT TableName, TableID " & _
         "FROM ASRSysTables " & _
         "WHERE TableID in " & _
         "  (SELECT ChildID from ASRSysRelations " & _
         "   WHERE ParentID = " & CStr(mlngBaseTableID) & ") OR TableID = " & mlngBaseTableID & _
         "   ORDER BY TableName"
  
  Set rsTables = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  
  If Not rsTables.BOF And Not rsTables.EOF Then
    Do Until rsTables.EOF
'      If AlreadyUsedInReport(rsTables!TableID, IIf(mblnNew, 0, mlngChildTableID)) = False Then
        cboEventTable.AddItem rsTables!TableName
        cboEventTable.ItemData(cboEventTable.NewIndex) = rsTables!TableID
        
        If (Not mblnNew) And (mlngEventTableID > 0) And (mlngEventTableID = rsTables!TableID) Then
          cboEventTable.ListIndex = cboEventTable.NewIndex
        ElseIf (mlngEventTableID < 1) And (mlngBaseTableID = rsTables!TableID) Then
          cboEventTable.ListIndex = cboEventTable.NewIndex
        End If

'      End If
      rsTables.MoveNext
    Loop
    cboEventTable.Enabled = True
    If cboEventTable.ListIndex < 0 And cboEventTable.ListCount > 0 Then
      cboEventTable.ListIndex = 0
    End If
  Else
    cboEventTable.Enabled = False
  End If
  rsTables.Close
  Set rsTables = Nothing
 
  ' Clear Legend Table combo and a <None> item.
  With cboLegendTable
    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'    .ListIndex = 0
  End With
  
  ' Get all the tables for the legend lookup combo
  sSQL = "SELECT TableName, TableID " & _
         "FROM ASRSysTables " & _
         "WHERE TableType = " & TableTypes.tabLookup & _
         "   ORDER BY TableName"
  
  Set rsTables = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  
  If Not rsTables.BOF And Not rsTables.EOF Then
    Do Until rsTables.EOF
      cboLegendTable.AddItem rsTables!TableName
      cboLegendTable.ItemData(cboLegendTable.NewIndex) = rsTables!TableID
      
      If (Not mblnNew) And (mlngLegendTableID > 0) And (mlngLegendTableID = rsTables!TableID) Then
        cboLegendTable.ListIndex = cboLegendTable.NewIndex
      End If
      rsTables.MoveNext
    Loop
    cboLegendTable.Enabled = True
    If cboLegendTable.ListIndex < 0 And cboLegendTable.ListCount > 0 Then
      cboLegendTable.ListIndex = 0
    End If
  Else
    cboLegendTable.Enabled = False
  End If
  rsTables.Close
  Set rsTables = Nothing

  If (cboLegendTable.ListIndex < 0) And (cboLegendTable.ListCount) Then
    cboLegendTable.ListIndex = 0
  End If
  
  PopulateTables = True
  
TidyUpAndExit:
  Set rsTables = Nothing
  Exit Function
  
Error_Trap:
  COAMsgBox "Error populating Event table dropdown box.", vbExclamation + vbOKOnly, "Calendar Reports"
  PopulateTables = False
  GoTo TidyUpAndExit

End Function



Private Sub cboDuration_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboEndDate_click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboEndSession_click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboEventDesc1_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboEventDesc2_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboEventTable_Click()
  
  Dim rsDateColumns As New ADODB.Recordset
  Dim sSQL As String
    
  If Not mblnLoading Then
    
    mlngDateColCount = 0
    
    If mlngEventTableID <> cboEventTable.ItemData(cboEventTable.ListIndex) Then
      sSQL = "SELECT COUNT(ASRSysColumns.ColumnID) AS 'DateColumnCount' " & _
        " FROM ASRSysColumns " & _
        " WHERE ASRSysColumns.tableID = " & cboEventTable.ItemData(cboEventTable.ListIndex) & _
        " AND ASRSysColumns.columnType <> " & Trim(Str(colSystem)) & _
        " AND ASRSysColumns.columnType <> " & Trim(Str(colLink)) & _
        " AND ASRSysColumns.dataType = " & Trim(Str(SQLDataType.sqlDate))
      
      Set rsDateColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
      If rsDateColumns.BOF And rsDateColumns.EOF Then
          COAMsgBox "The selected Event Table has no date columns. Please select an Event Table that contains date columns.", vbOKOnly + vbExclamation, "Calendar Reports"
'          mblnLoading = True
'          SetComboItem Me.cboEventTable, mlngEventTableID
'          mblnLoading = False
'          rsDateColumns.Close
'          Set rsDateColumns = Nothing
'          Exit Sub
      Else
        mlngDateColCount = rsDateColumns.Fields("DateColumnCount").Value
        If mlngDateColCount < 1 Then
          COAMsgBox "The selected Event Table has no date columns. Please select an Event Table that contains date columns.", vbOKOnly + vbExclamation, "Calendar Reports"
'          mblnLoading = True
'          SetComboItem Me.cboEventTable, mlngEventTableID
'          mblnLoading = False
        End If
      End If
      
      rsDateColumns.Close
      Set rsDateColumns = Nothing
    Else
      Set rsDateColumns = Nothing
      Exit Sub
    End If
    
    txtEventFilter.Text = ""
    txtEventFilter.Tag = 0
    txtCharacter.Text = ""
    
    mlngEventTableID = cboEventTable.ItemData(cboEventTable.ListIndex)
 
    UpdateEventDependantFields
    RefreshEventFrames
    RefreshLegendFrame
    
    Changed = True
    
  End If
  
End Sub



Private Sub cboEventType_Click()
  If Not mblnLoading Then
    If cboEventType.ListIndex >= 0 Then
      GetLookupTableDefaultDetails
      Changed = True
    End If
  End If
End Sub

Private Sub cboLegendCode_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboLegendColumn_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub cboLegendTable_Click()
  If Not mblnLoading Then
    UpdateLegendDependantFields
    RefreshLegendFrame
    Changed = True
  End If
End Sub

Private Sub cboStartDate_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub cboStartSession_Click()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

Private Sub cmdCancel_Click()
  Me.Cancelled = True
  Unload Me
End Sub
Private Sub cmdEventFilter_Click()
  GetFilter cboEventTable, txtEventFilter
End Sub
Private Sub GetFilter(ctlSource As Control, ctlTarget As Control)
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    If TypeOf ctlSource Is TextBox Then
      fOK = .Initialise(ctlSource.Tag, Val(ctlTarget.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    ElseIf TypeOf ctlSource Is ComboBox Then
      fOK = .Initialise(ctlSource.ItemData(ctlSource.ListIndex), Val(ctlTarget.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    End If
      
    If fOK Then
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) = True Then
              
        ' Read the selected expression info.
        ctlTarget.Text = IIf(Len(.Name) = 0, "<None>", .Name)
        ctlTarget.Tag = .ExpressionID
        
        Changed = True
      End If
      
    End If
    
    End With
  
  
  Set objExpression = Nothing

  If mfrmParent.DefinitionOwner Then
    'ForceDefinitionToBeHiddenIfNeeded
  End If

End Sub
Private Sub cmdOK_Click()
  If ValidateEventInfo Then
    Cancelled = False
    Me.Hide
  End If
End Sub
Private Function ValidateEventInfo() As Boolean

  ' Check a name has been entered
  If (Trim(txtEventName.Text) = "") Or (Len(Trim(txtEventName.Text)) = 0) Then
    COAMsgBox "You must give this event a name.", vbExclamation, Me.Caption
    txtEventName.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If
 
  ' Check the name is unique
  If Not CheckUniqueEventName(Trim(txtEventName.Text)) Then
    COAMsgBox "An event called '" & Trim(txtEventName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtEventName.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If
  
  ' Check that a valid event table has been selected
  If Not (cboEventTable.ItemData(cboEventTable.ListIndex) > 0) Then
    COAMsgBox "A valid event table has not been selected.", vbExclamation, Me.Caption
    cboEventTable.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If

  ' Check that the event table has date columns
  If (cboStartDate.ListCount < 1) Then
          COAMsgBox "The selected Event Table has no date columns. Please select an Event Table that contains date columns.", vbOKOnly + vbExclamation, Me.Caption
    cboEventTable.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If

  ' Check that a valid start date column has been selected
  If cboStartDate.ListCount > 0 Then
    If Not (cboStartDate.ItemData(cboStartDate.ListIndex) > 0) Then
      COAMsgBox "A valid start date column has not been selected.", vbExclamation, Me.Caption
      cboStartDate.SetFocus
      ValidateEventInfo = False
      Exit Function
    End If
  Else
    COAMsgBox "A valid start date column has not been selected.", vbExclamation, Me.Caption
    cboStartDate.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If
  
  ' Check that either a valid end date or duration column has been selected
  If optDuration.Value And Not (cboDuration.ItemData(cboDuration.ListIndex) > 0) Then
    COAMsgBox "A valid duration column has not been selected.", vbExclamation, Me.Caption
    optDuration.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If
  If optEndDate.Value And Not (cboEndDate.ItemData(cboEndDate.ListIndex) > 0) Then
    COAMsgBox "A valid end date column has not been selected.", vbExclamation, Me.Caption
    optEndDate.SetFocus
    ValidateEventInfo = False
    Exit Function
  End If

  If mblnHasLookupColumn And (cboEventType.ListIndex >= 0) Then
    ' Check that a valid 'set' of lookup tables have been selected
    If optLegendLookup.Value And (cboLegendTable.ListCount < 1) Then
      COAMsgBox "A valid lookup table has not been selected.", vbExclamation, Me.Caption
      ValidateEventInfo = False
      Exit Function
    ElseIf optLegendLookup.Value And Not (cboLegendTable.ItemData(cboLegendTable.ListIndex) > 0) Then
      COAMsgBox "A valid lookup table has not been selected.", vbExclamation, Me.Caption
      cboLegendTable.SetFocus
      ValidateEventInfo = False
      Exit Function
    End If
    
    If optLegendLookup.Value And (cboLegendColumn.ListCount < 1) Then
      COAMsgBox "A valid lookup column has not been selected.", vbExclamation, Me.Caption
      ValidateEventInfo = False
      Exit Function
    ElseIf optLegendLookup.Value And Not (cboLegendColumn.ItemData(cboLegendColumn.ListIndex) > 0) Then
      COAMsgBox "A valid lookup column has not been selected.", vbExclamation, Me.Caption
      cboLegendColumn.SetFocus
      ValidateEventInfo = False
      Exit Function
    End If
    
    If optLegendLookup.Value And (cboLegendCode.ListCount < 1) Then
      COAMsgBox "A valid lookup code has not been selected.", vbExclamation, Me.Caption
      ValidateEventInfo = False
      Exit Function
    ElseIf optLegendLookup.Value And Not (cboLegendCode.ItemData(cboLegendCode.ListIndex) > 0) Then
      COAMsgBox "A valid lookup code has not been selected.", vbExclamation, Me.Caption
      cboLegendCode.SetFocus
      ValidateEventInfo = False
      Exit Function
    End If
    
    If optLegendLookup.Value And (cboLegendCode.ListCount < 1) Then
      COAMsgBox "A valid event type has not been selected.", vbExclamation, Me.Caption
      ValidateEventInfo = False
      Exit Function
    ElseIf optLegendLookup.Value And Not (cboEventType.ItemData(cboEventType.ListIndex) > 0) Then
      COAMsgBox "A valid event type has not been selected.", vbExclamation, Me.Caption
      cboEventType.SetFocus
      ValidateEventInfo = False
      Exit Function
    End If
  End If
  
  ValidateEventInfo = True
End Function
Private Function CheckUniqueEventName(pstrNewEventKey As String) As Boolean

  'checks that the event name does not already exist in the current definition.
  
  Dim i As Integer
  Dim objEvent As clsCalendarEvent
  
  CheckUniqueEventName = True
  
  For Each objEvent In mcolReportEvents.Collection
    If pstrNewEventKey = objEvent.Key Then
      CheckUniqueEventName = False
      Exit Function
    End If
  Next objEvent
  
  CheckUniqueEventName = True
  
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim pintAnswer As Integer
    
    If ((Changed = True) And (UnloadMode <> vbFormCode)) Or ((Changed = True) And (Cancelled = True)) Then
      
      pintAnswer = COAMsgBox("Save changes ?", vbQuestion + vbYesNoCancel, "Calendar Reports")
        
      If pintAnswer = vbYes Then
        cmdOK_Click
        Cancel = True
        Exit Sub
      ElseIf pintAnswer = vbCancel Then
        Cancel = True
        Exit Sub
      Else
        Cancelled = True
      End If
    Else
      If UnloadMode <> vbFormCode Then
        Me.Cancelled = True
      End If

    End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set datData = Nothing
End Sub






Private Sub optCharacter_Click()
  If Not mblnLoading Then
    RefreshLegendFrame
    Changed = True
  End If
End Sub
Private Sub optDuration_Click()
  If Not mblnLoading Then
    RefreshEventFrames
    Changed = True
  End If
End Sub
Private Sub optEndDate_Click()
  If Not mblnLoading Then
    RefreshEventFrames
    Changed = True
  End If
End Sub
Private Sub optLegendLookup_Click()
  If Not mblnLoading Then
    RefreshLegendFrame
    Changed = True
  End If
End Sub
Private Sub optNoEnd_Click()
  If Not mblnLoading Then
    RefreshEventFrames
    Changed = True
  End If
End Sub






Private Sub txtCharacter_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub


Private Sub txtCharacter_Validate(Cancel As Boolean)
  If Trim(txtCharacter.Text) = "." Then
    Cancel = True
    COAMsgBox "The full stop/decimal point is a reserved character in Calendar Reports.", vbOKOnly + vbExclamation, "Calendar Reports"
    txtCharacter.SetFocus
  End If
End Sub


Private Sub txtEventFilter_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub
Private Sub txtEventName_Change()
  If Not mblnLoading Then
    Changed = True
  End If
End Sub

