VERSION 5.00
Begin VB.Form frmDiaryFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Diary Events"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1031
   Icon            =   "frmDiaryFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5550
      Begin VB.CheckBox chkDiaryOnlyMine 
         Caption         =   "&Only show manual events where owner is 'username'"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   1560
         Width           =   5320
      End
      Begin VB.ComboBox cboDateRange 
         Height          =   315
         ItemData        =   "frmDiaryFilter.frx":000C
         Left            =   1965
         List            =   "frmDiaryFilter.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1100
         Width           =   2500
      End
      Begin VB.ComboBox cboAlarmStatus 
         Height          =   315
         ItemData        =   "frmDiaryFilter.frx":004C
         Left            =   1965
         List            =   "frmDiaryFilter.frx":0059
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   700
         Width           =   2500
      End
      Begin VB.ComboBox cboEventType 
         Height          =   315
         ItemData        =   "frmDiaryFilter.frx":0088
         Left            =   1965
         List            =   "frmDiaryFilter.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date range :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   1160
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alarm status :"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   760
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type of event :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4470
      TabIndex        =   9
      Top             =   2175
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3210
      TabIndex        =   8
      Top             =   2175
      Width           =   1200
   End
End
Attribute VB_Name = "frmDiaryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboEventType_Click()
  
  Dim blnSystemOnly As Boolean
  
  With cboEventType
    If .ListIndex <> -1 Then
      blnSystemOnly = (.ItemData(.ListIndex) = 1)
      chkDiaryOnlyMine.Enabled = Not blnSystemOnly
      If blnSystemOnly Then
        chkDiaryOnlyMine.Value = vbUnchecked
      End If
    End If
  End With

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  Dim intCount As Integer
  Dim intCurrentView As Integer
  
  With cboEventType
    If .ListIndex <> -1 Then
      gobjDiary.FilterEventType = .ItemData(.ListIndex)
    End If
  End With
  
  With cboAlarmStatus
    If .ListIndex <> -1 Then
      gobjDiary.FilterAlarmStatus = .ItemData(.ListIndex)
    End If
  End With
  
  With cboDateRange
    If .ListIndex <> -1 Then
      gobjDiary.FilterPastPresent = .ItemData(.ListIndex)
    End If
  End With

  gobjDiary.FilterOnlyMine = (chkDiaryOnlyMine.Value = vbChecked)

  
  'MH20020701 Fault 4088
  'If gobjDiary.GetRecordCount = 0 And _
     (gobjDiary.FilterEventType > 0 Or _
      gobjDiary.FilterAlarmStatus > 0 Or _
      gobjDiary.FilterPastPresent > 0 Or _
      gobjDiary.FilterOnlyMine = True) Then
  If gobjDiary.GetRecordCount = 0 And _
     ((gobjDiary.FilterEventType > 0 And gobjDiary.AllowSystemEvents) Or _
      gobjDiary.FilterAlarmStatus > 0 Or _
      gobjDiary.FilterPastPresent > 0 Or _
      gobjDiary.FilterOnlyMine = True) Then
    
    COAMsgBox "No records match the current filter." & vbCrLf & _
           "No filter is applied.", vbInformation + vbOKOnly, "Diary Filter"
    
    cboEventType.ListIndex = 0
    cboAlarmStatus.ListIndex = 0
    cboDateRange.ListIndex = 0
    chkDiaryOnlyMine.Value = vbUnchecked
    
    gobjDiary.FilterEventType = 0
    gobjDiary.FilterAlarmStatus = 0
    gobjDiary.FilterPastPresent = 0
    gobjDiary.FilterOnlyMine = False
    frmDiary.Caption = gobjDiary.FilterText
    gobjDiary.RefreshDiaryData
    
  Else
    gobjDiary.RefreshDiaryData
    frmDiary.Caption = gobjDiary.FilterText
    Unload Me
  
  End If
  
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

  Call LoadCombos

  With gobjDiary
    SetComboItem cboEventType, .FilterEventType
    SetComboItem cboAlarmStatus, .FilterAlarmStatus
    SetComboItem cboDateRange, .FilterPastPresent
    chkDiaryOnlyMine.Value = IIf(.FilterOnlyMine, vbChecked, vbUnchecked)
  End With

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set frmDiaryFilter = Nothing
End Sub


Private Sub LoadCombos()

  With cboEventType
    .Clear
    
    If gobjDiary.AllowSystemEvents Then
      .AddItem "<All>"
      .ItemData(.NewIndex) = 0
      .AddItem "System Events"
      .ItemData(.NewIndex) = 1
    End If

    .AddItem "Manual Events"
    .ItemData(.NewIndex) = 2
    '.AddItem "Manual Events where owner is '" & gsUserName & "'"
    '.ItemData(.NewIndex) = 3

  End With
  
  With cboAlarmStatus
    .Clear
    .AddItem "<All>"
    .ItemData(.NewIndex) = 0
    .AddItem "Alarmed Events"
    .ItemData(.NewIndex) = 1
    .AddItem "Non-Alarmed Events"
    .ItemData(.NewIndex) = 2
  End With

  With cboDateRange
    .Clear
    .AddItem "<All>"
    .ItemData(.NewIndex) = 0
    .AddItem "Past Events"
    .ItemData(.NewIndex) = 1
    .AddItem "Current and Future Events"
    .ItemData(.NewIndex) = 2
  End With

  chkDiaryOnlyMine.Caption = "Only &show manual events where owner is '" & gsUserName & "'"

End Sub

