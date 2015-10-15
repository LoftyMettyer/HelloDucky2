VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEventLog 
   Caption         =   "Event Log"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1035
   Icon            =   "frmEventLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2650
      Left            =   10200
      TabIndex        =   11
      Top             =   200
      Width           =   1200
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   400
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Height          =   400
         Left            =   0
         TabIndex        =   13
         Top             =   555
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete..."
         Height          =   400
         Left            =   0
         TabIndex        =   14
         Top             =   1125
         Width           =   1200
      End
      Begin VB.CommandButton cmdPurge 
         Caption         =   "&Purge..."
         Height          =   400
         Left            =   0
         TabIndex        =   15
         Top             =   1680
         Width           =   1200
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "&Email..."
         Height          =   400
         Left            =   0
         TabIndex        =   16
         Top             =   2250
         Width           =   1200
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   4845
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   19844
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdEventLog 
      Height          =   3425
      Left            =   120
      TabIndex        =   4
      Top             =   1185
      Width           =   10020
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      stylesets.count =   2
      stylesets(0).Name=   "ssetDormant"
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmEventLog.frx":000C
      stylesets(1).Name=   "ssetActive"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   -2147483646
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmEventLog.frx":0028
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
      StyleSet        =   "ssetDormant"
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ActiveRowStyleSet=   "ssetActive"
      CaptionAlignment=   0
      Columns.Count   =   11
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2990
      Columns(1).Caption=   "Start Time"
      Columns(1).Name =   "StartTime"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2990
      Columns(2).Caption=   "End Time"
      Columns(2).Name =   "EndTime"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1667
      Columns(3).Caption=   "Duration"
      Columns(3).Name =   "Duration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2884
      Columns(4).Caption=   "Type"
      Columns(4).Name =   "Utility Type"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3810
      Columns(5).Caption=   "Name"
      Columns(5).Name =   "Utility Name"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1773
      Columns(6).Caption=   "Status"
      Columns(6).Name =   "Status"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1323
      Columns(7).Caption=   "Mode"
      Columns(7).Name =   "Mode"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1667
      Columns(8).Caption=   "User name"
      Columns(8).Name =   "Username"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "BatchJobID"
      Columns(9).Name =   "BatchJobID"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "BatchRunID"
      Columns(10).Name=   "BatchRunID"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17674
      _ExtentY        =   6041
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
   Begin VB.Frame fraFilters 
      Caption         =   "Filters :"
      Height          =   945
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10020
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   8610
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   500
         Width           =   1305
      End
      Begin VB.ComboBox cboUser 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   500
         Width           =   2010
      End
      Begin VB.ComboBox cboMode 
         Height          =   315
         Left            =   6660
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   500
         Width           =   1050
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmEventLog.frx":0044
         Left            =   3960
         List            =   "frmEventLog.frx":0046
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   500
         Width           =   2200
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Left            =   8610
         TabIndex        =   10
         Top             =   255
         Width           =   675
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User name :"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode :"
         Height          =   195
         Left            =   6660
         TabIndex        =   7
         Top             =   255
         Width           =   585
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   3960
         TabIndex        =   6
         Top             =   255
         Width           =   555
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar abEventLog 
      Left            =   10230
      Top             =   3045
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
      Bands           =   "frmEventLog.frx":0048
   End
End
Attribute VB_Name = "frmEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private pstrOrderOrder As String
Private mintSortColumnIndex As Integer

' Flag showing if user has delete permission
Private mblnDeleteEnabled As Boolean

'MH20030422
Private mstrFilterIDs As String

' Flag showing if user has email permission
Private mblnEmailEnabled As Boolean

' Flag showing if user has purge permission
Private mblnPurgeEnabled As Boolean

Public Function CheckEventExists(pstrEventIDs As String) As Boolean

  Dim rstDetailRecords As ADODB.Recordset
  Dim strSQL As String
  Dim fOK As Boolean
  
  If (Len(pstrEventIDs) > 0) Then
    strSQL = "SELECT [ASRSysEventLog].[ID] FROM [ASRSysEventLog] WHERE [ASRSysEventLog].[ID] IN (" & pstrEventIDs & ")"
  
    Set rstDetailRecords = datGeneral.GetReadOnlyRecords(strSQL)
    
    fOK = Not (rstDetailRecords.BOF And rstDetailRecords.EOF)
    
  Else
    fOK = False
  End If
  
  If Not fOK Then
    COAMsgBox "The selected Event Log record(s) have been deleted by another User.", vbOKOnly + vbExclamation, "Event Log"
  End If
  
  CheckEventExists = fOK
  
End Function

Private Sub abEventLog_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.abEventLog_Click(Tool)", Array(Tool)

  Select Case Tool.Name
  
    Case "View"
      ViewEvent
      
    Case "Delete"
      DeleteEvent
      
    Case "Email"
      EmailEvent
            
  End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub cboMode_Click()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cboMode_Click()"
  
  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cboStatus_Click()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cboStatus_Click()"

  RefreshGrid
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cboUser_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cboUser_Click()"

  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cboType_Click()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cboType_Click()"

  If cboType.Text = "Diary Delete" _
    Or cboType.Text = "Diary Rebuild" _
    Or cboType.Text = "Email Rebuild" _
    Or cboType.Text = "Workflow Rebuild" Then
    
    mblnLoading = True
    cboMode.Enabled = False
    cboMode.BackColor = vbButtonFace
    SetComboText cboMode, "<All>"
    mblnLoading = False
  Else
    cboMode.Enabled = True
    cboMode.BackColor = vbWindowBackground
  End If
  
  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub DeleteEvent()
  
  Dim strEventIDs  As String
  Dim plngLoop As Long
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  Dim arrayBookmarks() As Variant
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.DeleteEvent()"
  
  With frmSelection
    .Source = "Event Log"
    .HelpContextID = Me.HelpContextID
    .Caption = Me.Caption & " Selection"
    .Show vbModal
  End With

  Select Case frmSelection.Answer
  
    Case -1 ' Cancelled
    
      Unload frmSelection
      GoTo TidyUpAndExit
      
    Case 0 ' Delete the currently highlighted entry(ies)
    
      Screen.MousePointer = vbHourglass

      'Workout how many records have been selected
      nTotalSelRows = grdEventLog.SelBookmarks.Count
      
      'Redimension the arrays to the count of the bookmarks
      ReDim arrayBookmarks(nTotalSelRows)
      
      For intCount = 1 To nTotalSelRows
        arrayBookmarks(intCount) = grdEventLog.SelBookmarks.Item(intCount - 1)
      Next intCount
      
      For intCount = 1 To nTotalSelRows
        grdEventLog.Bookmark = arrayBookmarks(intCount)
        
        If Len(strEventIDs) > 0 Then
          strEventIDs = strEventIDs & ","
        End If
        strEventIDs = strEventIDs & grdEventLog.Columns("ID").Value
      Next intCount
      grdEventLog.SelBookmarks.RemoveAll

      If Len(strEventIDs) > 0 Then
        gADOCon.Execute "DELETE FROM AsrSysEventLog WHERE ID IN (" & strEventIDs & ")"
      End If
      
    Case 1 ' Delete entries currently displayed in the grid
    
      Screen.MousePointer = vbHourglass
      
      mrstHeaders.MoveFirst
      Do Until mrstHeaders.EOF
        If Len(strEventIDs) > 0 Then
          strEventIDs = strEventIDs & ","
        End If
        strEventIDs = strEventIDs & mrstHeaders.Fields("ID")
        mrstHeaders.MoveNext
      Loop

      If Len(strEventIDs) > 0 Then
        gADOCon.Execute "DELETE FROM AsrSysEventLog WHERE ID IN (" & strEventIDs & ")"
      End If
      

    Case 2 ' Delete all entries visible to the user
    
      Screen.MousePointer = vbHourglass
    
      If mblnViewAllEntries = True Then
        gADOCon.Execute "DELETE FROM AsrSysEventLog"
      Else
        gADOCon.Execute "DELETE FROM AsrSysEventLog WHERE username = '" & datGeneral.UserNameForSQL & "'"
      End If
      
  End Select
  
'''  If Me.grdEventLog.Rows > 0 Then
'''    Me.grdEventLog.SelBookmarks.RemoveAll
'''    Me.grdEventLog.MoveFirst
'''    Me.grdEventLog.SelBookmarks.Add Me.grdEventLog.Bookmark
'''  End If

  Unload frmSelection
  Screen.MousePointer = vbDefault
  RefreshGrid
  RefreshButtons

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
  
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cmdDelete_Click()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cmdDelete_Click()"

  DeleteEvent
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub cmdEmail_Click()
  'NHRD2608204 Fault 8276 Added Email to right click functionality
  ' Email currently selected events
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cmdEmail_Click()"

  EmailEvent

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
 
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cmdOK_Click()"

  Unload Me

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cmdPurge_Click()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cmdPurge_Click()"

  frmEventLogPurge.Show vbModal
  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub cmdView_Click()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cmdView_Click()"

  ViewEvent

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub Form_Activate()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.Form_Activate()"

  DoColumnSizes
  
  UI.RemoveClipping

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
On Error GoTo ErrorTrap
gobjErrorStack.PushStack "frmEventLog.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    Unload Me
  Case KeyCode = vbKeyF5
    RefreshGrid
  Case KeyCode = vbKeyDelete
    'MH20030516 Fault 4438
    If cmdDelete.Enabled Then
      cmdDelete_Click
    End If
End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub Form_Load()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.Form_Load()"

  Hook Me.hWnd, 11700, 5550
  
  Dim rstUsers As Recordset
  Set mclsData = New clsDataAccess

  mblnLoading = True

  fraButtons.BackColor = Me.BackColor

  'If user does not have delete event log permission, hide the delete button
  If datGeneral.SystemPermission("EVENTLOG", "DELETE") = False Then
    cmdDelete.Enabled = False
    mblnDeleteEnabled = False
  Else
    mblnDeleteEnabled = True
  End If
  
  'If user does not have email event log permission, hide the email button
  If datGeneral.SystemPermission("EVENTLOG", "EMAIL") = False Then
    cmdEmail.Enabled = False
    mblnEmailEnabled = False
  Else
    mblnEmailEnabled = True
  End If
  
  'If user can see all entries, populate and enable the users combo, else
  'populate it with the users name only and disable it
  If datGeneral.SystemPermission("EVENTLOG", "VIEWALL") = True Then
    mblnViewAllEntries = True
    Set rstUsers = mclsData.OpenRecordset("SELECT DISTINCT Username from AsrSysEventLog ORDER BY Username", adOpenForwardOnly, adLockReadOnly)
    cboUser.AddItem "<All>"
    Do Until rstUsers.EOF
      cboUser.AddItem rstUsers.Fields("Username")
      rstUsers.MoveNext
    Loop
    Set rstUsers = Nothing
  Else
    Me.Caption = Me.Caption & " [Viewing own entries only]"
    cboUser.AddItem gsUserName
    cboUser.Enabled = False
    cboUser.BackColor = &H8000000F
  End If
  
  cboUser.ListIndex = 0
  
  'If user does not have purge event log permission, hide the purge button
  If datGeneral.SystemPermission("EVENTLOG", "PURGE") = False Then
    cmdPurge.Enabled = False
    mblnPurgeEnabled = False
  Else
    mblnPurgeEnabled = True
  End If
  
  'Purge the event log before we display it
  gADOCon.Execute "EXEC sp_AsrEventLogPurge"
  
  'Add all available functions to the Type combo
  With cboType
    .AddItem "<All>"
		.AddItem "Cross Tab" : .ItemData(.NewIndex) = eltCrossTab
		.AddItem "9-Box Grid" : .ItemData(.NewIndex) = elt9BoxGrid
    .AddItem "Custom Report": .ItemData(.NewIndex) = eltCustomReport
    .AddItem "Data Transfer": .ItemData(.NewIndex) = eltDataTransfer
    .AddItem "Diary Rebuild": .ItemData(.NewIndex) = eltDiaryRebuild
    .AddItem "Email Rebuild": .ItemData(.NewIndex) = eltEmailRebuild
    .AddItem "Export": .ItemData(.NewIndex) = eltExport
    .AddItem "Global Add": .ItemData(.NewIndex) = eltGlobalAdd
    .AddItem "Global Delete": .ItemData(.NewIndex) = eltGlobalDelete
    .AddItem "Global Update": .ItemData(.NewIndex) = eltGlobalUpdate
    .AddItem "Import": .ItemData(.NewIndex) = eltImport
    .AddItem "Mail Merge": .ItemData(.NewIndex) = eltMailMerge
    .AddItem "Standard Report": .ItemData(.NewIndex) = eltStandardReport
    '.AddItem "Record Editing": .ItemData(.NewIndex) = eltRecordEditing
    .AddItem "System Error": .ItemData(.NewIndex) = eltSystemError
    .AddItem "Match Report": .ItemData(.NewIndex) = eltMatchReport
    .AddItem "Envelopes & Labels": .ItemData(.NewIndex) = eltLabel
'    .AddItem "Label Definition": .ItemData(.NewIndex) = eltLabelType
    .AddItem "Calendar Report": .ItemData(.NewIndex) = eltCalandarReport
    .AddItem "Record Profile": .ItemData(.NewIndex) = eltRecordProfile
    .AddItem "Succession Planning": .ItemData(.NewIndex) = eltSuccessionPlanning
    .AddItem "Career Progression": .ItemData(.NewIndex) = eltCareerProgression
    
    If gbWorkflowEnabled Then
      .AddItem "Workflow Rebuild": .ItemData(.NewIndex) = eltWorkflowRebuild
    End If
    
    'MH20060421 Fault 11077 Two payroll items removed as per PC.
    'If gbAccordEnabled Then .AddItem "Payroll Transfer (In)": .ItemData(.NewIndex) = eltAccordImport
    'If gbAccordEnabled Then .AddItem "Payroll Transfer (Out)": .ItemData(.NewIndex) = eltAccordExport
    
    .ListIndex = 0
  End With
    
  'Add the available Modes to the Mode combo
  With cboMode
    .AddItem "<All>"
    .AddItem "Batch"
    .AddItem "Manual"
    .AddItem "Pack"
    .ListIndex = 0
  End With
  
  'Add all available statuses to the Status combo
  With cboStatus
    .AddItem "<All>"
    .AddItem "Cancelled"
    .AddItem "Error"
    .AddItem "Failed"
    .AddItem "Pending"
    .AddItem "Skipped"
    .AddItem "Successful"
    .ListIndex = 0
  End With
  
  mblnLoading = False
  
  'Set height and width to last saved. Form is centred on screen
  Me.Height = GetPCSetting("EventLog", "Height", Me.Height)
  Me.Width = GetPCSetting("EventLog", "Width", Me.Width)

  'Set default sort order to be date desc
  pstrOrderField = "DateTime"
  mintSortColumnIndex = 1
  pstrOrderOrder = "DESC"
  
  'Populate the grid
  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)

  ' Save the window size ready to recall next time user views the event log
  SavePCSetting "EventLog", "Height", Me.Height
  SavePCSetting "EventLog", "Width", Me.Width
  UI.RemoveClipping

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub Form_Resize()

  Const COMBO_GAP As Integer = 170
  Const lngGap As Long = 120
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.Form_Resize()"

  'JPD 20030908 Fault 5756
  DisplayApplication
  
  ' Ensure form does not get too small/big. Also reposition controls as necessary
'  UI.ClipForForm Me, 5550, 11700
'  If Me.Width < 11700 Then Me.Width = 11700
'  If Me.Width > Screen.Width Then Me.Width = (Screen.Width - 200)
'  If Me.Height < 5550 Then Me.Height = 5550
'  If Me.Height > Screen.Height Then Me.Height = (Screen.Height - 500)
  
  fraButtons.Left = Me.ScaleWidth - (fraButtons.Width + lngGap)
  
  fraFilters.Width = fraButtons.Left - (lngGap * 2)
  
  cboStatus.Left = fraFilters.Width - (cboStatus.Width + COMBO_GAP)
  lblStatus.Left = cboStatus.Left

  cboMode.Left = cboStatus.Left - (cboMode.Width + COMBO_GAP)
  lblMode.Left = cboMode.Left

  cboType.Left = cboMode.Left - (cboType.Width + COMBO_GAP)
  lblType.Left = cboType.Left

  cboUser.Width = cboType.Left - (cboUser.Left + COMBO_GAP)

  grdEventLog.Width = fraFilters.Width
  grdEventLog.Height = Me.ScaleHeight - (fraFilters.Height + StatusBar1.Height + (lngGap * 3))
  
  DoColumnSizes
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
  Me.Refresh
  Me.grdEventLog.Refresh
  
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

  With grdEventLog
    lngAvailableWidth = grdEventLog.Width - (270 + .Columns(1).Width + .Columns(2).Width + .Columns(3).Width + .Columns(4).Width + .Columns(6).Width + .Columns(7).Width)
    
    .Columns(5).Width = (lngAvailableWidth * 0.68)
    .Columns(8).Width = (lngAvailableWidth * 0.32)
  End With
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Function RefreshGrid() As Boolean

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.RefreshGrid()"
  
  ' Populate the grid using filter/sort criteria as set by the user
  Dim pstrSQL As String
  Dim intMode As Integer
  
  If mblnLoading = True Then GoTo TidyUpAndExit
  
  Screen.MousePointer = vbHourglass
  
  pstrSQL = "SELECT [ASRSysEventLog].[ID], "
  pstrSQL = pstrSQL & "[ASRSysEventLog].[DateTime],"
  pstrSQL = pstrSQL & "[ASRSysEventLog].[EndTime],"
  pstrSQL = pstrSQL & "ISNULL([ASRSysEventLog].[Duration],-1) AS 'Duration',"
  pstrSQL = pstrSQL & "[ASRSysEventLog].[Type],"
  pstrSQL = pstrSQL & "[ASRSysEventLog].[Name],"
  pstrSQL = pstrSQL & "CASE [ASRSysEventLog].[Status] WHEN 0 THEN 'Pending' WHEN 1 THEN 'Cancelled' WHEN 2 THEN 'Failed' WHEN 3 THEN 'Successful' WHEN 4 THEN 'Skipped' WHEN 5 THEN 'Error' END AS 'Status',"
  pstrSQL = pstrSQL & "CASE WHEN [ASRSysEventLog].[ReportPack] = 1 THEN 'Pack' WHEN [ASRSysEventLog].[Mode] = 1 THEN 'Batch' WHEN [ASRSysEventLog].[Mode] = 0 THEN 'Manual' END AS 'Mode',"
  pstrSQL = pstrSQL & "[ASRSysEventLog].[Username],"
  pstrSQL = pstrSQL & "[ASRSysEventLog].[BatchJobID],"
  pstrSQL = pstrSQL & "[ASRSysEventLog].[BatchRunID]"
  
  pstrSQL = pstrSQL & "FROM [ASRSysEventLog]"

  pstrSQL = pstrSQL & " WHERE [ASRSysEventLog].[Type] NOT IN (" & eltAccordImport & ", " & eltAccordExport & ")"

  If cboType.ItemData(cboType.ListIndex) > 0 Then
    pstrSQL = pstrSQL & " AND [ASRSysEventLog].[Type] = " & CStr(cboType.ItemData(cboType.ListIndex))
  End If
  
  If cboStatus.Text <> "<All>" Then
    pstrSQL = pstrSQL & IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & "[ASRSysEventLog].[status] = "
    Select Case cboStatus.Text
      Case "Cancelled": pstrSQL = pstrSQL & 1
      Case "Failed": pstrSQL = pstrSQL & 2
      Case "Successful": pstrSQL = pstrSQL & 3
      Case "Skipped": pstrSQL = pstrSQL & 4
      Case "Error": pstrSQL = pstrSQL & 5
      Case Else: pstrSQL = pstrSQL & 0
    End Select
  End If

  If mblnViewAllEntries = False Then
    pstrSQL = pstrSQL & IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & "[ASRSysEventLog].[Username] = '" & datGeneral.UserNameForSQL & "'"
  Else
    If cboUser.Text <> "<All>" Then
      pstrSQL = pstrSQL & IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & "[ASRSysEventLog].[Username] = '" & Replace(cboUser.Text, "'", "''") & "'"
    End If
  End If
  
  ' Put the mode filter in...
  If cboMode.Text <> "<All>" Then
    Select Case cboMode.Text
			Case "Batch" : pstrSQL = pstrSQL & IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & "[ASRSysEventLog].[Mode] = " & 1 & " AND ([ASRSysEventLog].[ReportPack] = " & 0 & " OR [ASRSysEventLog].[ReportPack] IS NULL)"
      Case "Pack": pstrSQL = pstrSQL & IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & "[ASRSysEventLog].[ReportPack] = " & 1
			Case "Manual" : pstrSQL = pstrSQL & IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & "[ASRSysEventLog].[Mode] = " & 0 & " AND ([ASRSysEventLog].[ReportPack] = " & 0 & " OR [ASRSysEventLog].[ReportPack] IS NULL)"
		End Select
  End If
  
  'MH20030422
  If mstrFilterIDs <> vbNullString Then
    pstrSQL = pstrSQL & _
      IIf(InStr(pstrSQL, "WHERE") > 0, " AND ", " WHERE ") & _
      "ID IN (" & mstrFilterIDs & ")"
  End If
  
  ' RH 10/10/00 - BUG 1117 - Mode is either 0 or 1, so when sorting, you have to do the opposite order
  If pstrOrderField = "Mode" Then

      pstrSQL = pstrSQL & " ORDER BY CASE " & _
        " WHEN [ASRSysEventLog].[ReportPack] = 1 THEN 'Pack' WHEN [ASRSysEventLog].[Mode] = 1 THEN 'Batch' WHEN [ASRSysEventLog].[Mode] = 0 THEN 'Manual' END " & _
        IIf(pstrOrderOrder = "ASC", "ASC", "DESC")

  ElseIf pstrOrderField = "Type" Then
      'NHRD15012003 Fault 4617 These items for the type column were
      'being ordered by the type field which is originally numericly based.
      'This code is inserted to order the types by their display names as found
      'under the same fault reference in the Load event of frmEventlog.
      'Any additions to this order-list will have to be duplicated there.
      pstrSQL = pstrSQL & " ORDER BY CASE WHEN [ASRSysEventLog].[Type] = 1 THEN 'Cross Tab'" & _
        " WHEN [ASRSysEventLog].[Type] = 2 THEN 'Custom Report'" & _
        " WHEN [ASRSysEventLog].[Type] = 3 THEN 'Data Transfer'" & _
        " WHEN [ASRSysEventLog].[Type] = 4 THEN 'Export'" & _
        " WHEN [ASRSysEventLog].[Type] = 5 THEN 'Global Add'" & _
        " WHEN [ASRSysEventLog].[Type] = 6 THEN 'Global Delete'" & _
        " WHEN [ASRSysEventLog].[Type] = 7 THEN 'Global Update'" & _
        " WHEN [ASRSysEventLog].[Type] = 8 THEN 'Import'" & _
        " WHEN [ASRSysEventLog].[Type] = 9 THEN 'Mail Merge'" & _
        " WHEN [ASRSysEventLog].[Type] = 10 THEN 'Diary Delete'" & _
        " WHEN [ASRSysEventLog].[Type] = 11 THEN 'Diary Rebuild'" & _
        " WHEN [ASRSysEventLog].[Type] = 12 THEN 'Email Rebuild'" & _
        " WHEN [ASRSysEventLog].[Type] = 13 THEN 'Standard Report'" & _
        " WHEN [ASRSysEventLog].[Type] = 14 THEN 'Record Editing'" & _
        " WHEN [ASRSysEventLog].[Type] = 15 THEN 'System Error'" & _
        " WHEN [ASRSysEventLog].[Type] = 16 THEN 'Match Report'" & _
        " WHEN [ASRSysEventLog].[Type] = 17 THEN 'Calendar Report'" & _
        " WHEN [ASRSysEventLog].[Type] = 18 THEN 'Envelopes & Labels'" & _
        " WHEN [ASRSysEventLog].[Type] = 19 THEN 'Label Definition'" & _
        " WHEN [ASRSysEventLog].[Type] = 20 THEN 'Record Profile'"
        
      pstrSQL = pstrSQL & _
        " WHEN [ASRSysEventLog].[Type] = 21 THEN 'Succession Planning'" & _
        " WHEN [ASRSysEventLog].[Type] = 22 THEN 'Career Progression'" & _
        " WHEN [ASRSysEventLog].[Type] = 23 THEN 'Payroll Transfer'" & _
        " WHEN [ASRSysEventLog].[Type] = 25 THEN 'Workflow Rebuild'" & _
        " ELSE ''" & _
        " END " & IIf(pstrOrderOrder = "ASC", "ASC", "DESC")
        
  ElseIf pstrOrderField = "Status" Then
  
      pstrSQL = pstrSQL & " ORDER BY CASE " & _
        " WHEN [ASRSysEventLog].[Status] = 1 THEN 'Cancelled'" & _
        " WHEN [ASRSysEventLog].[Status] = 2 THEN 'Failed'" & _
        " WHEN [ASRSysEventLog].[Status] = 3 THEN 'Successful'" & _
        " WHEN [ASRSysEventLog].[Status] = 4 THEN 'Skipped'" & _
        " WHEN [ASRSysEventLog].[Status] = 5 THEN 'Error'" & _
        " ELSE 'Pending'" & _
        " END " & IIf(pstrOrderOrder = "ASC", "ASC", "DESC")
  Else
      pstrSQL = pstrSQL & " ORDER BY [ASRSysEventLog].[" & pstrOrderField & "] " & pstrOrderOrder
  End If
  
  Set mrstHeaders = mclsData.OpenPersistentRecordset(pstrSQL, adOpenKeyset, adLockReadOnly)
  
  With grdEventLog
    .Redraw = False
    .Rebind
    .Rows = mrstHeaders.RecordCount
    .Redraw = True
  End With
  
  If Me.grdEventLog.Rows > 0 Then
    Me.grdEventLog.MoveFirst
    Me.grdEventLog.SelBookmarks.Add Me.grdEventLog.Bookmark
  End If
  
  Dim strStatusBarText As String
  
  strStatusBarText = vbNullString
  strStatusBarText = strStatusBarText & " " & mrstHeaders.RecordCount & " Record" & IIf(mrstHeaders.RecordCount > 1 Or mrstHeaders.RecordCount = 0, "s", "")
  If mrstHeaders.RecordCount > 1 Then
    strStatusBarText = strStatusBarText & " Sorted by "
    If pstrOrderField = "DateTime" Then
      strStatusBarText = strStatusBarText & "Start Time "
    ElseIf pstrOrderField = "EndTime" Then
      strStatusBarText = strStatusBarText & "End Time "
    ElseIf pstrOrderField = "Username" Then
      strStatusBarText = strStatusBarText & "User name "
    Else
    strStatusBarText = strStatusBarText & pstrOrderField & " "
    End If
    strStatusBarText = strStatusBarText & "in "
    
    strStatusBarText = strStatusBarText & IIf(pstrOrderOrder = "ASC", "Ascending", "Descending")
    strStatusBarText = strStatusBarText & " order"
  End If
  
  StatusBar1.SimpleText = strStatusBarText
  
  RefreshButtons
  
  DoColumnSizes

  Screen.MousePointer = vbDefault

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError
  
End Function

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdEventLog_Click()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_Click()"

  If (Me.grdEventLog.SelBookmarks.Count > 1) Or (Me.grdEventLog.Rows = 0) Then
    Me.cmdView.Enabled = False
  Else
    Me.cmdView.Enabled = True
  End If
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub grdEventLog_DblClick()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_DblClick()"
  
  If (Me.grdEventLog.Rows > 0) And Me.grdEventLog.SelBookmarks.Count = 1 Then
    ViewEvent
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub RefreshButtons()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.RefreshButtons()"

  cmdView.Enabled = Me.grdEventLog.Rows > 0
  cmdDelete.Enabled = (Me.grdEventLog.Rows > 0) And (mblnDeleteEnabled = True)
  cmdEmail.Enabled = (Me.grdEventLog.Rows > 0) And (mblnEmailEnabled = True)
  

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub grdEventLog_HeadClick(ByVal ColIndex As Integer)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_HeadClick(ColIndex)", Array(ColIndex)

  ' Set the sort criteria depending on the column header clicked and refresh the grid
  Select Case ColIndex
    Case 1: pstrOrderField = "DateTime"
    Case 2: pstrOrderField = "EndTime"
    Case 3: pstrOrderField = "Duration"
    Case 4: pstrOrderField = "Type"
    Case 5: pstrOrderField = "Name"
    Case 6: pstrOrderField = "Status"
    Case 7: pstrOrderField = "Mode"
    Case 8: pstrOrderField = "Username"
  End Select
  
  If ColIndex = mintSortColumnIndex Then
    If pstrOrderOrder = "ASC" Then pstrOrderOrder = "DESC" Else pstrOrderOrder = "ASC"
  End If
  
  mintSortColumnIndex = ColIndex
  
  RefreshGrid
    
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub grdEventLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_MouseUp(Button,Shift,X,Y)", Array(Button, Shift, X, Y)

 If (Button = vbRightButton) And (Y > Me.grdEventLog.RowHeight) Then
    ' Enable/disable the required tools.
    With Me.abEventLog.Bands("bndEventLog")
      .Tools("View").Enabled = Me.cmdView.Enabled
      .Tools("Delete").Enabled = Me.cmdDelete.Enabled
      .Tools("Email").Enabled = Me.cmdEmail.Enabled
      .TrackPopup -1, -1
    End With
    
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub grdEventLog_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    
  If Me.grdEventLog.SelBookmarks.Count > 1 Then
    Me.cmdView.Enabled = False
  ElseIf Me.grdEventLog.SelBookmarks.Count = 1 Then
    Me.grdEventLog.SelBookmarks.RemoveAll
    Me.grdEventLog.SelBookmarks.Add Me.grdEventLog.Bookmark
    Me.cmdView.Enabled = True
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub grdEventLog_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_UnboundPositionData(StartLocation,NumberOfRowsToMove,NewLocation)", Array(StartLocation, NumberOfRowsToMove, NewLocation)

  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrstHeaders.MoveLast
    Else
      mrstHeaders.MoveFirst
    End If
  Else
    mrstHeaders.Bookmark = StartLocation
  End If

  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrstHeaders.Move NumberOfRowsToMove
  NewLocation = mrstHeaders.Bookmark

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub grdEventLog_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.grdEventLog_UnboundReadData(RowBuf,StartLocation,ReadPriorRows)", Array(RowBuf, StartLocation, ReadPriorRows)

  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String

  sDateFormat = DateFormat

  iRowsRead = 0

  ' Do nothing if we a re just formatting the grid,
  ' or if there a re no records to display.
  If mrstHeaders Is Nothing Then GoTo TidyUpAndExit
  If mrstHeaders.State = adStateClosed Then GoTo TidyUpAndExit

  ' Do nothing if we are loading or if there are no records to display
  If mblnLoading = True And mrstHeaders.RecordCount = 0 Then GoTo TidyUpAndExit

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
            Case "ID"
              RowBuf.Value(iRowIndex, iFieldIndex) = CStr(mrstHeaders.Fields("ID"))
            Case "DateTime"
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrstHeaders.Fields("DateTime"), sDateFormat & " hh:nn")
            Case "EndTime"
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrstHeaders.Fields("EndTime"), sDateFormat & " hh:nn")
            Case "Duration"
              RowBuf.Value(iRowIndex, iFieldIndex) = FormatEventDuration(mrstHeaders.Fields("Duration").Value)
            Case "Type"
              RowBuf.Value(iRowIndex, iFieldIndex) = GetUtilityType(mrstHeaders.Fields("Type"))
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrstHeaders(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
    End Select

    If ReadPriorRows Then
      mrstHeaders.MovePrevious
    Else
      mrstHeaders.MoveNext
    End If

    iRowsRead = iRowsRead + 1
  Next iRowIndex

  RowBuf.RowCount = iRowsRead

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

'''Private Sub grdEventLog_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
'''
'''  On Error GoTo ErrorTrap
'''  gobjErrorStack.PushStack "frmEventLog.grdEventLog_UnboundPositionData(StartLocation,NumberOfRowsToMove,NewLocation)", Array(StartLocation, NumberOfRowsToMove, NewLocation)
'''
'''  If IsNull(StartLocation) Then
'''    StartLocation = 0
'''  End If
'''
''''TM20020327 Fault 3058 - Not sure why this is here but it is causing the last row to
''''not be highlighted when first selected. (TM 23/04/02 - This comment is rubbish, please ignore)
'''  NewLocation = CLng(StartLocation) + NumberOfRowsToMove
'''
'''TidyUpAndExit:
'''  gobjErrorStack.PopStack
'''  Exit Sub
'''ErrorTrap:
'''  gobjErrorStack.HandleError
'''
'''End Sub

'''Private Sub grdEventLog_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
'''
'''  On Error GoTo ErrorTrap
'''  gobjErrorStack.PushStack "frmEventLog.grdEventLog_UnboundReadData(RowBuf,StartLocation,ReadPriorRows)", Array(RowBuf, StartLocation, ReadPriorRows)
'''
'''  ' Read the required data from the recordset to the grid.
'''  Dim iRowIndex As Integer
'''  Dim iFieldIndex As Integer
'''  Dim iRowsRead As Integer
'''  Dim sDateFormat As String
'''
'''  sDateFormat = DateFormat
'''
'''  ' This is required as recordset not set when this sub is first run
'''  If mrstHeaders Is Nothing Then GoTo TidyUpAndExit
'''  If mrstHeaders.State = adStateClosed Then GoTo TidyUpAndExit
'''
'''  ' Do nothing if we are loading or if there are no records to display
'''  If mblnLoading = True And mrstHeaders.RecordCount = 0 Then GoTo TidyUpAndExit
'''
'''  If StartLocation < 0 Then GoTo TidyUpAndExit
'''
'''  If IsNull(StartLocation) Or (StartLocation = 0) Then
'''    If ReadPriorRows Then
'''      If Not mrstHeaders.EOF Then
'''        mrstHeaders.MoveLast
'''      End If
'''    Else
'''      If Not mrstHeaders.BOF Then
'''        mrstHeaders.MoveFirst
'''      End If
'''    End If
'''  Else
'''    mrstHeaders.Bookmark = StartLocation
'''    If ReadPriorRows Then
'''      mrstHeaders.MovePrevious
'''    Else
'''      mrstHeaders.MoveNext
'''    End If
'''  End If
'''
'''  ' Read from the row buffer into the grid.
'''  For iRowIndex = 0 To (RowBuf.RowCount - 1)
'''    ' Do nothing if the begining of end of the recordset is Met.
'''    If mrstHeaders.BOF Or mrstHeaders.EOF Then Exit For
'''
'''    ' Optimize the data read based on the ReadType.
'''    Select Case RowBuf.ReadType
'''      Case 0
'''        For iFieldIndex = 0 To (mrstHeaders.Fields.Count - 1)
'''          Select Case mrstHeaders.Fields(iFieldIndex).Name
'''            Case "ID"
'''              RowBuf.Value(iRowIndex, iFieldIndex) = CStr(mrstHeaders.Fields("ID"))
'''            Case "DateTime"
'''              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrstHeaders.Fields("DateTime"), sDateFormat & " hh:nn")
'''            Case "EndTime"
'''              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrstHeaders.Fields("EndTime"), sDateFormat & " hh:nn")
'''            Case "Duration"
'''              RowBuf.Value(iRowIndex, iFieldIndex) = FormatEventDuration(mrstHeaders.Fields("Duration").Value)
'''            Case "Type"
'''              RowBuf.Value(iRowIndex, iFieldIndex) = GetUtilityType(mrstHeaders.Fields("Type"))
'''              'RowBuf.Value(iRowIndex, iFieldIndex) = mrstHeaders.Fields("Type")
'''            Case "Mode"
'''              RowBuf.Value(iRowIndex, iFieldIndex) = IIf(mrstHeaders.Fields("Mode") = False, "Manual", "Batch")
'''            Case Else
'''              RowBuf.Value(iRowIndex, iFieldIndex) = mrstHeaders(iFieldIndex)
'''          End Select
'''        Next iFieldIndex
'''        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
'''      Case 1
'''        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
'''      Case 2
'''        RowBuf.Value(iRowIndex, 0) = mrstHeaders(0)
'''        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
'''      Case 3
'''    End Select
'''
'''    If ReadPriorRows Then
'''      mrstHeaders.MovePrevious
'''    Else
'''      mrstHeaders.MoveNext
'''    End If
'''
'''    iRowsRead = iRowsRead + 1
'''  Next iRowIndex
'''
'''  RowBuf.RowCount = iRowsRead
'''
'''TidyUpAndExit:
'''  gobjErrorStack.PopStack
'''  Exit Sub
'''ErrorTrap:
'''  gobjErrorStack.HandleError
'''
'''End Sub


Private Function EmailEvent()
' Email currently selected events
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.cmdEmail_Click()"

  Dim frmSelection As frmEmailSel
  Dim strEventIDs As String
  Dim intLoop As Integer
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  Dim arrayBookmarks() As Variant
  
  Set frmSelection = New frmEmailSel

  ' Build a string of all the currently selected event entries
  If grdEventLog.SelBookmarks.Count > 0 Then
    
    strEventIDs = vbNullString
    
    'Workout how many records have been selected
    nTotalSelRows = grdEventLog.SelBookmarks.Count

    'Redimension the arrays to the count of the bookmarks
    ReDim arrayBookmarks(nTotalSelRows)
    
    For intCount = 1 To nTotalSelRows
      arrayBookmarks(intCount) = grdEventLog.SelBookmarks.Item(intCount - 1)
    Next intCount

    For intCount = 1 To nTotalSelRows
      grdEventLog.Bookmark = arrayBookmarks(intCount)
      
      If Len(strEventIDs) > 0 Then
        strEventIDs = strEventIDs & ","
      End If
      strEventIDs = strEventIDs & grdEventLog.Columns("ID").Value
    Next intCount
    
    If CheckEventExists(strEventIDs) Then
      frmSelection.SetupEventLogSend strEventIDs
      frmSelection.Caption = Me.Caption & " Selection"
      frmSelection.Show vbModal
    End If
  End If
  
  RefreshGrid
  
TidyUpAndExit:
  Set frmSelection = Nothing
  gobjErrorStack.PopStack
  Exit Function
 
ErrorTrap:
  gobjErrorStack.HandleError
End Function
Private Function ViewEvent()
  
  Dim frmDetails As frmEventLogDetails
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmEventLog.ViewEvent()"

  ' RH 25/09 - Temp fix for bug 988 - need to get rid of blank line somehow!
  If IsNumeric(grdEventLog.Columns("ID").Value) Then
    Set frmDetails = New frmEventLogDetails
    If frmDetails.Initialise(grdEventLog.Columns("ID").Value, _
        IIf((Trim(grdEventLog.Columns("BatchJobID").Value) = ""), 0, grdEventLog.Columns("BatchJobID").Value), _
        IIf((Trim(grdEventLog.Columns("BatchRunID").Value) = ""), 0, grdEventLog.Columns("BatchRunID").Value), Me) = True Then
      
      frmDetails.Caption = Me.Caption & " Details"
      frmDetails.Show vbModal
    Else
      RefreshGrid
      Unload frmDetails
    End If
  End If

TidyUpAndExit:
  Set frmDetails = Nothing
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  gobjErrorStack.HandleError
  
End Function

'MH20030422
Public Sub FilterIDs(ByVal strNewValue As String, utilType As utilityType)
  
  Dim lngIndex As Long
  
  mstrFilterIDs = strNewValue

  If mstrFilterIDs <> vbNullString Then
    
    For lngIndex = cboType.ListCount - 1 To 0 Step -1
      Select Case cboType.List(lngIndex)
      Case "Diary Rebuild", "Email Rebuild", "Workflow Rebuild"
        cboType.RemoveItem lngIndex
      End Select
    Next
    
    EnableCombo cboMode, False
    If utilType = utlReportPack Then
        SetComboText cboMode, "Pack"
    Else
        SetComboText cboMode, "Batch"
    End If
    EnableCombo cboUser, False
    'SetComboText cboUser, gsUserName
    SetComboText cboUser, gsUserName, True
  
    cmdDelete.Visible = False
    cmdPurge.Visible = False
    cmdEmail.Top = cmdDelete.Top
  End If

End Sub

