VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmEventLogDetails 
   Caption         =   "Event Log Details"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1112
   Icon            =   "frmEventLogDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEmail 
      Caption         =   "&Email..."
      Height          =   400
      Left            =   3960
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details :"
      Height          =   4005
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7770
      Begin VB.Frame fraBatch 
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   7335
         Begin VB.ComboBox cboAllJobs 
            Height          =   315
            Left            =   4845
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   60
            Width           =   2490
         End
         Begin COALine.COA_Line ASRLine 
            Height          =   30
            Index           =   0
            Left            =   120
            Top             =   600
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   53
         End
         Begin VB.Label lblBatchJobName 
            BackStyle       =   0  'Transparent
            Caption         =   "BatchJobName"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1695
            TabIndex        =   26
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label lblBatchJobNameLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Job Name :"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label lblAllJobsLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "All Jobs in Batch :"
            Height          =   195
            Left            =   3270
            TabIndex        =   24
            Top             =   120
            Width           =   1665
         End
      End
      Begin VB.Frame fraEventDetails 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   7335
         Begin VB.Label lblMode 
            BackStyle       =   0  'Transparent
            Caption         =   "Mode"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6360
            TabIndex        =   32
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lblModeLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Mode : "
            Height          =   195
            Left            =   5400
            TabIndex        =   31
            Top             =   120
            Width           =   870
         End
         Begin VB.Label lblDuration 
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6360
            TabIndex        =   30
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblDurationLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Duration :"
            Height          =   195
            Left            =   5400
            TabIndex        =   29
            Top             =   480
            Width           =   870
         End
         Begin VB.Label lblEndTime 
            BackStyle       =   0  'Transparent
            Caption         =   "99/99/9999  00:00"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3690
            TabIndex        =   28
            Top             =   480
            Width           =   1605
         End
         Begin VB.Label lblEndTimeLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "End :"
            Height          =   195
            Left            =   2955
            TabIndex        =   27
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "The status"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3690
            TabIndex        =   21
            Top             =   840
            Width           =   1605
         End
         Begin VB.Label lblStatusLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Status :"
            Height          =   195
            Left            =   2955
            TabIndex        =   20
            Top             =   840
            Width           =   705
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "This is the name of the utility that the event is based upon!"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   810
            TabIndex        =   19
            Top             =   120
            Width           =   4530
         End
         Begin VB.Label lblNameLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   645
         End
         Begin VB.Label lblType 
            BackStyle       =   0  'Transparent
            Caption         =   "The type"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   765
            TabIndex        =   17
            Top             =   840
            Width           =   2085
         End
         Begin VB.Label lblTypeLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Type :"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   600
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            Caption         =   "The user"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6360
            TabIndex        =   15
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblUserLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "User name :"
            Height          =   195
            Left            =   5400
            TabIndex        =   14
            Top             =   840
            Width           =   870
         End
         Begin VB.Label lblStartTime 
            BackStyle       =   0  'Transparent
            Caption         =   "99/99/9999  00:00"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   765
            TabIndex        =   13
            Top             =   480
            Width           =   2070
         End
         Begin VB.Label lblStartTimeLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Start :"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   555
         End
      End
      Begin VB.Frame fraRecords 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   7335
         Begin COALine.COA_Line ASRLine 
            Height          =   30
            Index           =   1
            Left            =   120
            Top             =   120
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   53
         End
         Begin VB.Label lblFailLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Records Failed :"
            Height          =   195
            Left            =   2955
            TabIndex        =   10
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblSuccessLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Records Successful :"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblSuccess 
            BackStyle       =   0  'Transparent
            Caption         =   "Count"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1995
            TabIndex        =   8
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lblFailed 
            BackStyle       =   0  'Transparent
            Caption         =   "Count"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   4425
            TabIndex        =   7
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.PictureBox pctBatch 
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   315
         ScaleHeight     =   1020
         ScaleWidth      =   7140
         TabIndex        =   4
         Top             =   225
         Width           =   7140
      End
      Begin SSDataWidgets_B.SSDBGrid grdDetails 
         Height          =   1290
         Left            =   225
         TabIndex        =   0
         Top             =   2520
         Width           =   7305
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         FieldSeparator  =   ";"
         AllowUpdate     =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         CellNavigation  =   1
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns(0).Width=   14076
         Columns(0).Caption=   "Details"
         Columns(0).Name =   "Details"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).VertScrollBar=   -1  'True
         UseDefaults     =   0   'False
         TabNavigation   =   1
         _ExtentX        =   12885
         _ExtentY        =   2275
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print..."
      Height          =   400
      Left            =   5280
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6660
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmEventLogDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Declare sizing/positioning constants
Const BUTTON_GAP = 240
Const BUTTON_WIDTH = 1200
Const BUTTON_HEIGHT = 400
Const GAP_AFTER_FRAME = 360
Const GAP_AFTER_GRID = 405
Const SCROLLBAR_WIDTH = 250
Const LABEL_GAP = 1965
Const LABEL_GAP2 = 1260

' Holds the record level details
Private mrstDetailRecords As Recordset

' Flag to store if we are currently resizing the form
Private mblnSizing As Boolean

' Integer used for control positioning
Private mintCurrentYPosition As Integer

' Flags to store what header record controls to display
Private mblnBatch As Boolean
Private mblnRecords As Boolean
Private mblnEventDetails As Boolean

' Flag to hold if we are loading or not (used when initialising the combo)
Private mblnLoading As Boolean

' Stores the BatchRunID if the selected job is part of a batch job
' (enabled the retrieval of other jobs within the same batch)
Private mlngBatchRunID As Long
Private mlngBatchJobID As Long

Private mlngEventLogID As Long

Private mfrmEventLog As frmEventLog

Private Function GetJobStatus(plngEventID As Long)

  Dim rsJobStatus As ADODB.Recordset

  Set rsJobStatus = datGeneral.GetReadOnlyRecords("SELECT [ASRSysEventLog].[Status] FROM [ASRSysEventLog] WHERE [ASRSysEventLog].[ID] = " & CStr(plngEventID))
  
  With rsJobStatus
    If Not (.BOF And .EOF) Then
      GetJobStatus = GetUtilityStatus(CInt(.Fields("Status")))
    Else
      GetJobStatus = GetUtilityStatus(-1)
    End If
    .Close
  End With
  
  Set rsJobStatus = Nothing
  
End Function

Public Function Initialise(plngKey As Long, _
                          plngBatchJobID As Long, plngBatchRunID As Long, _
                          pfrmEventLog As frmEventLog) As Boolean

  Dim fOK As Boolean
  
  On Error GoTo ErrorTrap
  
  fOK = True
  
  mblnLoading = True
  
  ' Let user know we are doing something, and dont redraw the form until the controls
  ' have all been repositioned
  Screen.MousePointer = vbHourglass
  Me.AutoRedraw = False
  
  mlngEventLogID = plngKey
  mlngBatchJobID = plngBatchJobID
  mlngBatchRunID = plngBatchRunID
  
  Set mfrmEventLog = pfrmEventLog
  
  ' Display the correct header information on the form
  If fOK Then fOK = DoHeaderInfo(mlngEventLogID)
  
  If fOK Then fOK = PopulateDetailsGrid
  
  'If user does not have email event log permission, hide the email button
  If datGeneral.SystemPermission("EVENTLOG", "EMAIL") = False Then
    cmdEmail.Enabled = False
  End If

  ' Needed incase user selects a new job in the batch combo which has different
  ' properties than the previously selected job and controls need repositioning
  If fOK Then Form_Resize
  
  ' Let user know we have finished, and can now redraw the form
  Screen.MousePointer = vbDefault
  Me.AutoRedraw = True
  
  fraBatch.Visible = mblnBatch
  fraEventDetails.Visible = mblnEventDetails
  fraRecords.Visible = mblnRecords
  
'  If cboOtherJobs.Enabled Then SetComboText cboAllJobs, cboOtherJobs.Text
  
  mblnLoading = False
  Initialise = fOK
 
  Dim lngMinFormHeight As Long
  If fraEventDetails.Visible And fraBatch.Visible + fraRecords.Visible Then
    lngMinFormHeight = 5190
  ElseIf fraEventDetails.Visible And fraBatch.Visible Then
    lngMinFormHeight = 4695
  ElseIf fraEventDetails.Visible And fraRecords.Visible Then
    lngMinFormHeight = 4455
  Else
    lngMinFormHeight = 3960
  End If
  Hook Me.hWnd, 9000, lngMinFormHeight

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  Initialise = False
  COAMsgBox "Error retrieving detail entries for this record." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Utility Run Log"
  GoTo TidyUpAndExit

End Function

Private Function PopulateDetailsGrid() As Boolean

  Dim strSQL As String
  
  ' Retrieve the details for the selected header record
  strSQL = "SELECT [AsrSysEventLogDetails].[Notes] " & _
            "FROM [AsrSysEventLogDetails] " & _
            "WHERE [AsrSysEventLogDetails].[EventLogID] = " & mlngEventLogID & " " & _
            "ORDER BY [AsrSysEventLogDetails].[ID]"
            
  Set mrstDetailRecords = datGeneral.GetReadOnlyRecords(strSQL)

  ' Populate the details grid with these records
  With grdDetails
    .RemoveAll
    Do Until mrstDetailRecords.EOF
      .AddItem mrstDetailRecords.Fields("Notes")
      mrstDetailRecords.MoveNext
    Loop
    Set mrstDetailRecords = Nothing
    
    ' Set scrollbar/enabled properties depending on existance of detail records
    If .Rows = 0 Then
      .Columns("Details").Caption = "No details exist for this entry"
      .Enabled = False
      .ScrollBars = ssScrollBarsNone
      '.ScrollBars = ssScrollBarsAutomatic
      .Height = 285
    Else
      .Enabled = True
      .ScrollBars = ssScrollBarsVertical
      '.ScrollBars = ssScrollBarsAutomatic
      .MoveFirst
      .SelBookmarks.Add .Bookmark
      .Columns("Details").Caption = "Details (" & .Rows & " Entries)"
    End If
  End With
  
  ' Update the grid title
  grdDetails_ScrollAfter
  
  PopulateDetailsGrid = True

End Function

Private Sub cboAllJobs_Click()

 ' Reload the details form with the new job details
  Initialise cboAllJobs.ItemData(cboAllJobs.ListIndex), mlngBatchJobID, mlngBatchRunID, mfrmEventLog
    
  ' Update the main event log window so that the job currently displayed on the
  ' details form is highlighted on the main grid
  With mfrmEventLog
    With .mrstHeaders
      .MoveFirst
      .Find "ID = " & cboAllJobs.ItemData(cboAllJobs.ListIndex)
    End With
    
    If Not mfrmEventLog.mrstHeaders.EOF Then
      With .grdEventLog
        .Bookmark = mfrmEventLog.mrstHeaders.Bookmark
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add mfrmEventLog.grdEventLog.Bookmark
      End With
    End If
  End With
  
End Sub


Private Sub cmdEmail_Click()
  'Email this entry
  Dim strEventID As String
  Dim frmSelection As frmEmailSel
  
  Set frmSelection = New frmEmailSel
  strEventID = Trim(Str(mlngEventLogID))

  If mfrmEventLog.CheckEventExists(strEventID) Then
    frmSelection.SetupEventLogSend strEventID
    frmSelection.Show vbModal
  End If
  
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
   
  Dim objPrintDef As clsPrintDef
  Dim rsEvent As ADODB.Recordset
  Dim frmMSG As frmMessageBox
  
  Dim blnPrintAll As Boolean
  Dim blnPrintEnd As Boolean
  
  Dim iTemp As Integer
  Dim intJobCount As Integer
  Dim intEventCount As Integer
  
  Dim plngLoop As Long
  Dim plngPosition1 As Long
  Dim plngPosition2 As Long
  Dim lngReturnCode As Long
  Dim lngSecondaryResponse As Long
  
  Dim strPrompt As String
  Dim strSQL As String
  Dim pstrErrorString As String
  Dim strDateFormat As String
  
  Dim pvarbookmark As Variant
    
  Screen.MousePointer = vbHourglass
  
  strDateFormat = DateFormat
  
  blnPrintAll = False
  
  If mblnBatch Then
AskAgain:
    Set frmMSG = New frmMessageBox
    frmMSG.CustomAddButton "&Selected", 100001
    frmMSG.CustomAddButton "&All", 100002
    frmMSG.CustomAddButton "&Cancel", vbCancel
    
    strPrompt = "Do you want to print the selected event only, or would you like to print all events in this batch?"
    
    lngReturnCode = frmMSG.CustomMessageBox(strPrompt, vbQuestion, "Event Log")
    
    If (lngReturnCode = 100002) Then
      blnPrintAll = True
        lngSecondaryResponse = COAMsgBox("Are you sure you want to print ALL events?", vbYesNo + vbQuestion, "Event Log")
        If lngSecondaryResponse = vbNo Then
          blnPrintAll = False
          GoTo AskAgain
        End If
    ElseIf (lngReturnCode = 100001) Then
      blnPrintAll = False
    Else
      Set frmMSG = Nothing
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    
  End If
  
  Set objPrintDef = New DataMgr.clsPrintDef
    
  If blnPrintAll Then
    'NHRD18012005 Fault 9082 Moved the objPrintDef.IsOK and objPrintDef.PrintStart(False) outside the
    'loop so all batch items are treated as one document.
    If objPrintDef.IsOK Then
      If objPrintDef.PrintStart(False) Then
          
          For intEventCount = 0 To (cboAllJobs.ListCount - 1)
          
            strSQL = vbNullString
            strSQL = strSQL & "SELECT  [E].[ID], "
            strSQL = strSQL & "        [E].[DateTime],"
            strSQL = strSQL & "        [E].[Type],"
            strSQL = strSQL & "        [E].[Name],"
            strSQL = strSQL & "        [E].[Status],"
            strSQL = strSQL & "        [E].[Username],"
            strSQL = strSQL & "        [E].[Mode],"
            strSQL = strSQL & "        [E].[BatchName],"
            strSQL = strSQL & "        [E].[SuccessCount],"
            strSQL = strSQL & "        [E].[FailCOunt],"
            strSQL = strSQL & "        [E].[BatchRunID],"
            strSQL = strSQL & "        [E].[EndTime],"
            strSQL = strSQL & "        [E].[Duration],"
            strSQL = strSQL & "        [E].[BatchJobID],"
            strSQL = strSQL & "        [D].[Notes] "
            strSQL = strSQL & "FROM [ASRSysEventLog] [E]"
            strSQL = strSQL & "      LEFT OUTER JOIN ASRSysEventLogDetails [D]"
            strSQL = strSQL & "      ON [E].[ID] = [D].[EventLogID] "
            strSQL = strSQL & "Where [E].[ID] = " & CStr(cboAllJobs.ItemData(intEventCount))
      
            Set rsEvent = datGeneral.GetReadOnlyRecords(strSQL)
            
            If Not (rsEvent.BOF And rsEvent.EOF) Then
  
              With objPrintDef
                .TabsOnPage = 3
                  .PrintHeader "Event Log : " & GetUtilityType(rsEvent!Type) & " '" & rsEvent!Name & "'"
                
                  .PrintNonBold "Mode :" & vbTab & IIf(rsEvent!Mode, "Batch", "Manual")
                  .PrintNormal
                  .PrintNonBold "Start :" & vbTab & Format(rsEvent!DateTime, strDateFormat & " hh:mm:ss")
                  .PrintNonBold "End :" & vbTab & Format(rsEvent!EndTime, strDateFormat & " hh:mm:ss")
                  .PrintNonBold "Duration :" & vbTab & FormatEventDuration(rsEvent!Duration)
                  .PrintNormal
                  .PrintNonBold "Type :" & vbTab & GetUtilityType(rsEvent!Type)
                  .PrintNonBold "Status :" & vbTab & GetUtilityStatus(rsEvent!Status)
                  .PrintNonBold "User name :" & vbTab & rsEvent!UserName
                  .PrintNormal
                  
                  If fraBatch.Visible = True Then
                    .PrintNonBold "Batch Job Name :" & vbTab & rsEvent!BatchName
                    .PrintNormal
                    .PrintNormal "All Jobs in Batch :"
                    .PrintNormal
                    For iTemp = 0 To cboAllJobs.ListCount - 1
                      .PrintNonBold cboAllJobs.List(iTemp) & " (" & GetJobStatus(cboAllJobs.ItemData(iTemp)) & ")"
                    Next iTemp
                  End If
                  
                  If fraRecords.Visible = True Then
                    .PrintNormal
                    .PrintNormal "Records Successful :" & rsEvent!SuccessCount
                    .PrintNormal "Records Failed :" & rsEvent!FailCount
                  End If
                  
                  .PrintNormal
                  .PrintBold "Details :"
                  .PrintNormal
                  
                  If IsNull(rsEvent!Notes) Then
                    .PrintNonBold "There are no details for this event log entry"
                  Else
                    plngLoop = 0
                    Do Until rsEvent.EOF
                      plngLoop = plngLoop + 1
                      ' Print the detail number header
                      .PrintBold "***  Log entry " & plngLoop & " of " & rsEvent.RecordCount & "  ***"
                      ' Print the actual error / record detail
                      pstrErrorString = rsEvent!Notes
                      pstrErrorString = Replace(pstrErrorString, vbCrLf, vbCr)
                      pstrErrorString = Replace(pstrErrorString, vbTab, "    ")
                      .PrintNonBold pstrErrorString
                      .PrintNormal
                      
                      rsEvent.MoveNext
                    Loop
                  End If
    
                  'If we're at the end of the loop End printing otherwise start new page
                  If intEventCount = (cboAllJobs.ListCount - 1) Then
                     .PrintEnd
                     objPrintDef.PrintConfirm "Event Log Details", "Event Log Details"
                  Else
                    .PrintNewPage
                  End If
              End With
            End If      'Not (rsEvent.BOF And rsEvent.EOF)
            rsEvent.Close
            
            Set rsEvent = Nothing
          Next intEventCount
      End If 'printstart
    End If 'objPrintDef.IsOK
  Else
    If objPrintDef.IsOK Then
    
      With objPrintDef
        .TabsOnPage = 3
        If .PrintStart(False) Then
          .PrintHeader "Event Log : " & lblType.Caption & " '" & Replace(lblName.Caption, "&&", "&") & "'"
        
          .PrintNonBold lblModeLabel.Caption & vbTab & lblMode.Caption
          .PrintNormal
          .PrintNonBold lblStartTimeLabel.Caption & vbTab & lblStartTime.Caption
          .PrintNonBold lblEndTimeLabel.Caption & vbTab & lblEndTime.Caption
          .PrintNonBold lblDurationLabel.Caption & vbTab & lblDuration.Caption
          .PrintNormal
          .PrintNonBold lblTypeLabel.Caption & vbTab & Replace(lblType.Caption, "&&", "&")
          .PrintNonBold lblStatusLabel.Caption & vbTab & lblStatus.Caption
          .PrintNonBold lblUserLabel.Caption & vbTab & lblUser
          .PrintNormal
          
          If fraBatch.Visible = True Then
            .PrintNonBold lblBatchJobNameLabel.Caption & vbTab & Replace(lblBatchJobName.Caption, "&&", "&")
            .PrintNormal
            .PrintNormal lblAllJobsLabel.Caption
            .PrintNormal
            For iTemp = 0 To cboAllJobs.ListCount - 1
              .PrintNonBold cboAllJobs.List(iTemp) & " (" & GetJobStatus(cboAllJobs.ItemData(iTemp)) & ")"
            Next iTemp
          End If
          
          If fraRecords.Visible = True Then
            .PrintNormal
            .PrintNormal lblSuccessLabel.Caption & vbTab & lblSuccess.Caption
            .PrintNormal lblFailLabel.Caption & vbTab & lblFailed.Caption
          End If
          
          .PrintNormal
          .PrintBold "Details : "
          .PrintNormal
          
          If Me.grdDetails.Rows = 0 Then
            .PrintNonBold "There are no details for this event log entry"
          Else
            For plngLoop = 1 To Me.grdDetails.Rows
            
              ' Print the detail number header
              .PrintBold "***  Log entry " & plngLoop & " of " & grdDetails.Rows & "  ***"
              
              ' Print the actual error / record detail
              pvarbookmark = grdDetails.GetBookmark(plngLoop - 1)
              pstrErrorString = grdDetails.Columns("Details").CellText(pvarbookmark)
      
              pstrErrorString = Replace(pstrErrorString, vbCrLf, vbCr)
              pstrErrorString = Replace(pstrErrorString, vbTab, "    ")
              .PrintNonBold pstrErrorString
              
              .PrintNormal
              
            Next plngLoop
          End If
        
          .PrintEnd
          .PrintConfirm "Event Log Details", "Event Log Details"
        End If
      
      End With
      
    End If
  
  End If
  
  Set objPrintDef = Nothing

  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Activate()

  RemoveIcon Me
  
  ' At this point (and not before!) it is possible to set the flags for the picture
  ' boxes. Setting these flags on form_load or initialise would NOT actually set them!
  With fraBatch
    .Visible = mblnBatch
    .Enabled = mblnBatch
  End With
  With fraRecords
    .Visible = mblnRecords
    .Enabled = mblnRecords
  End With
  
End Sub

Private Sub Form_Load()
  
  ' Retrieve the size of the form when last viewed
  Me.Height = GetPCSetting("EventLogDetails", "Height", Me.Height)
  Me.Width = GetPCSetting("EventLogDetails", "Width", Me.Width)

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  ' Store the size of the form for retrieval when next viewed
  SavePCSetting "EventLogDetails", "Height", Me.Height
  SavePCSetting "EventLogDetails", "Width", Me.Width
  
End Sub

Private Sub Form_Resize()
  
  Dim lngOuterFrameWidth As Long
  Dim lngInnerFrameWidth As Long
  Dim lngColumn1Left As Long
  Dim lngColumn2Left As Long
  Dim lngColumn3Left As Long
  
  Dim lngColumn1DataLeft As Long
  Dim lngColumn2DataLeft As Long
  Dim lngColumn3DataLeft As Long
  
  Dim lngColumnWidth As Long
  Dim lngLineWidth As Long
  Dim lngGridTop As Long
  
  Const lngLabelOffset = 50
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  If mblnSizing Then Exit Sub

  mblnSizing = True

'  If fraEventDetails.Visible And fraBatch.Visible + fraRecords.Visible Then
'    If Me.Height < 5190 Then Me.Height = 5190
'  ElseIf fraEventDetails.Visible And fraBatch.Visible Then
'    If Me.Height < 4695 Then Me.Height = 4695
'  ElseIf fraEventDetails.Visible And fraRecords.Visible Then
'    If Me.Height < 4455 Then Me.Height = 4455
'  Else
'    If Me.Height < 3960 Then Me.Height = 3960
'  End If
'
'  If Me.Width < 9000 Then Me.Width = 9000
'
'  If Me.Width > Screen.Width Then Me.Width = (Screen.Width - 200)
'  If Me.Height > Screen.Height Then Me.Height = (Screen.Height - 200)

  lngOuterFrameWidth = Me.Width - (2 * grdDetails.Left)
  lngInnerFrameWidth = lngOuterFrameWidth - 435
  
  lngLineWidth = lngInnerFrameWidth - 135

  cmdOK.Left = Me.ScaleWidth - (BUTTON_WIDTH + (BUTTON_GAP / 2))
  cmdOK.Top = Me.ScaleHeight - (BUTTON_HEIGHT + (BUTTON_GAP / 2))

  cmdPrint.Left = cmdOK.Left - (BUTTON_WIDTH + (BUTTON_GAP / 2))
  cmdPrint.Top = Me.ScaleHeight - (BUTTON_HEIGHT + (BUTTON_GAP / 2))

  cmdEmail.Left = cmdPrint.Left - (BUTTON_WIDTH + (BUTTON_GAP / 2))
  cmdEmail.Top = Me.ScaleHeight - (BUTTON_HEIGHT + (BUTTON_GAP / 2))

  fraDetails.Width = lngOuterFrameWidth
  If cmdPrint.Top - BUTTON_GAP > 0 Then
    fraDetails.Height = cmdPrint.Top - BUTTON_GAP
  End If
  
  lngColumn1Left = 120
  lngColumn1DataLeft = lngColumn1Left + lblNameLabel.Width + 120
  lngColumn2Left = (Me.Width / 3)
  lngColumn2DataLeft = lngColumn2Left + lblEndTimeLabel.Width + 120
  lngColumn3Left = ((Me.Width / 3) * 2)
  lngColumn3DataLeft = lngColumn3Left + lblModeLabel.Width + 120
  
  ' Position the Batch Name and Combo on the form if they are visible
  If fraBatch.Visible Then
    fraBatch.Width = lngInnerFrameWidth
    
    lblBatchJobNameLabel.Left = lngColumn1Left
    lblBatchJobName.Left = lblBatchJobNameLabel.Left + lblBatchJobNameLabel.Width + 120
    lblBatchJobName.Width = lngColumn2Left - lblBatchJobName.Left - 240
    
    lblAllJobsLabel.Left = lngColumn2Left
    cboAllJobs.Left = lblAllJobsLabel.Left + lblAllJobsLabel.Width + 120
    cboAllJobs.Width = fraBatch.Width - cboAllJobs.Left
    
    ASRLine(0).Width = lngLineWidth
  End If

  fraEventDetails.Width = lngInnerFrameWidth
  
  'First Column
  lblNameLabel.Left = lngColumn1Left
  lblName.Left = lngColumn1DataLeft
  lblName.Width = lngColumn3Left - lngColumn1DataLeft - 240
  lblStartTimeLabel.Left = lngColumn1Left
  lblStartTime.Left = lngColumn1DataLeft
  lblStartTime.Width = lngColumn2Left - lngColumn1DataLeft - 240
  lblTypeLabel.Left = lngColumn1Left
  lblType.Left = lngColumn1DataLeft
  lblType.Width = lngColumn2Left - lngColumn1DataLeft - 240
  
  'Second Column
  lblEndTimeLabel.Left = lngColumn2Left
  lblEndTime.Left = lngColumn2DataLeft
  lblEndTime.Width = lngColumn3Left - lngColumn2DataLeft - 240
  lblStatusLabel.Left = lngColumn2Left
  lblStatus.Left = lngColumn2DataLeft
  lblStatus.Width = lngColumn3Left - lngColumn2DataLeft - 240
  
  'Third Column
  lblModeLabel.Left = lngColumn3Left
  lblMode.Left = lngColumn3DataLeft
  lblMode.Width = fraEventDetails.Width - lngColumn3Left - 240
  lblDurationLabel.Left = lngColumn3Left
  lblDuration.Left = lngColumn3DataLeft
  lblDuration.Width = fraEventDetails.Width - lngColumn3Left - 240
  lblUserLabel.Left = lngColumn3Left
  lblUser.Left = lngColumn3DataLeft
  lblUser.Width = fraEventDetails.Width - lngColumn3Left - 240
  
  ' Position the Success/Failed Labels on the form if they are visible
  If fraRecords.Visible Then
    fraRecords.Width = lngInnerFrameWidth
    ASRLine(1).Width = lngLineWidth
    
    lblSuccessLabel.Left = lngColumn1Left
    lblSuccess.Left = lngColumn1Left + lblSuccessLabel.Width + 120
    lblSuccess.Width = lngColumn2Left - lblSuccess.Left - 240
    
    lblFailLabel.Left = lngColumn2Left
    lblFailed.Left = lngColumn2Left + lblFailLabel.Width + 120
    lblFailed.Width = lblSuccess.Width
  End If

  ' Set the grid height and width
  With grdDetails
    .Width = lngOuterFrameWidth - 480
    If (fraDetails.Height - .Top - BUTTON_GAP) > 200 Then
      .Height = fraDetails.Height - .Top - BUTTON_GAP
    End If

    If grdDetails.Enabled Then
      If .Rows > 1 Then
        .Columns("Details").Width = (.Width - SCROLLBAR_WIDTH)
        .ScrollBars = ssScrollBarsVertical
        '.ScrollBars = ssScrollBarsAutomatic
      Else
        .Columns("Details").Width = .Width
        .ScrollBars = ssScrollBarsNone
        '.ScrollBars = ssScrollBarsAutomatic
      End If

    Else
      .Columns("Details").Width = .Width
      
    End If
    .RowHeight = .Height
  End With

  mblnSizing = False

End Sub


Private Function DoHeaderInfo(plngKey As Long) As Boolean

  ' Populate all the relevant header fields and position the controls
  On Error GoTo ErrorTrap

  Dim strSQL As String
  Dim rstTemp As Recordset
  Dim strJobName As String
  Dim strJobType As String
  
  strSQL = "SELECT * FROM ASRSysEventLog WHERE ID = " & plngKey

  Set rstTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  With rstTemp
  
    If .BOF And .EOF Then
      COAMsgBox "This record no longer exists in the event log.", vbExclamation + vbOKOnly, "Event Log"
      DoHeaderInfo = False
      GoTo TidyUpAndExit
    End If
    
    lblName.Caption = Replace(.Fields("Name"), "&", "&&")
    
    lblMode.Caption = IIf(.Fields("Mode").Value, "Batch", "Manual")
    
    lblStartTime.Caption = Format(.Fields("DateTime"), DateFormat & " hh:mm:ss")
    lblEndTime.Caption = Format(.Fields("EndTime"), DateFormat & " hh:mm:ss")
    
    If .Fields("Status") = elsPending Then
      lblDuration.Visible = False
    Else
      lblDuration.Visible = True
    End If
    
    lblDuration.Caption = FormatEventDuration(IIf(IsNull(.Fields("Duration").Value), 0, .Fields("Duration").Value))
    
    lblType.Caption = Replace(GetUtilityType(.Fields("Type")), "&", "&&")
    
    lblStatus.Caption = GetUtilityStatus(.Fields("Status"))
    
    lblUser.Caption = Replace(.Fields("Username"), "&", "&&")
      
    lblSuccess.Caption = IIf(IsNull(.Fields("SuccessCount")), "N/A", .Fields("SuccessCount"))
    lblFailed.Caption = IIf(IsNull(.Fields("FailCount")), "N/A", .Fields("FailCount"))
    
    mlngBatchRunID = IIf(IsNull(.Fields("BatchRunID")), 0, .Fields("BatchRunID"))
        
    If .Fields("Mode") = 0 Then
      mblnBatch = False
    Else
      mblnBatch = True
      lblBatchJobName.Caption = Replace(.Fields("BatchName"), "&", "&&")
    End If
    
    If (IsNull(.Fields("SuccessCount")) Or IsNull(.Fields("FailCount"))) Then
      fraRecords.Visible = False
      mblnRecords = False
    Else
      fraRecords.Visible = True
      mblnRecords = True
    End If
    
    mblnEventDetails = True
    
  End With
  
  If mblnBatch Then
    If cboAllJobs.ListCount = 0 Then
      strSQL = vbNullString
      strSQL = strSQL & " SELECT [ASRSysEventLog].[ID], "
      strSQL = strSQL & "        [ASRSysEventLog].[Name], "
      strSQL = strSQL & "        [ASRSysEventLog].[BatchRunID], "
      strSQL = strSQL & "        [ASRSysEventLog].[BatchJobID], "
      strSQL = strSQL & "        [ASRSysEventLog].[Type] "
      strSQL = strSQL & " FROM [ASRSysEventLog] "
      strSQL = strSQL & " WHERE [ASRSysEventLog].[BatchRunID] = " & mlngBatchRunID
      strSQL = strSQL & " ORDER BY [ASRSysEventLog].[ID]"
      
      Set rstTemp = datGeneral.GetReadOnlyRecords(strSQL)
      Do Until rstTemp.EOF
        strJobName = rstTemp.Fields("Name")
        strJobType = GetUtilityType(rstTemp.Fields("Type"))
        cboAllJobs.AddItem strJobType & " - " & strJobName
        cboAllJobs.ItemData(cboAllJobs.NewIndex) = rstTemp.Fields("ID")
        rstTemp.MoveNext
      Loop
      If cboAllJobs.ListCount <= 1 Then
        If cboAllJobs.ListCount = 0 Then
          cboAllJobs.AddItem "<None>"
        End If
        cboAllJobs.ListIndex = 0
        cboAllJobs.Enabled = False
        cboAllJobs.BackColor = vbButtonFace
      Else
        SetComboItem cboAllJobs, mlngEventLogID
      End If
    End If
  End If

  If mblnEventDetails And mblnBatch And mblnRecords Then
    fraBatch.Top = 240
    fraEventDetails.Top = fraBatch.Top + fraBatch.Height
    fraRecords.Top = fraEventDetails.Top + fraEventDetails.Height
    grdDetails.Top = fraRecords.Top + fraRecords.Height + 240
    
  ElseIf mblnEventDetails And mblnBatch Then
    fraBatch.Top = 240
    fraEventDetails.Top = fraBatch.Top + fraBatch.Height
    fraRecords.Top = fraEventDetails.Top + fraEventDetails.Height
    grdDetails.Top = fraEventDetails.Top + fraEventDetails.Height + 240

  ElseIf mblnEventDetails And mblnRecords Then
    fraBatch.Top = 240
    fraEventDetails.Top = 240
    fraRecords.Top = fraEventDetails.Top + fraEventDetails.Height
    grdDetails.Top = fraRecords.Top + fraRecords.Height + 240
  
  Else
    fraBatch.Top = 240
    fraEventDetails.Top = 240
    fraRecords.Top = fraEventDetails.Top + fraEventDetails.Height
    grdDetails.Top = fraEventDetails.Top + fraEventDetails.Height + 240
  End If

  DoHeaderInfo = True

TidyUpAndExit:
  Set rstTemp = Nothing
  Exit Function
  
ErrorTrap:

  COAMsgBox "Error whilst populating event log detail." & vbCrLf & "(" & Err.Description & ")"
  DoHeaderInfo = False

End Function

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdDetails_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  
  ' Update the grid caption after the user has used keys to view the details
    If grdDetails.Rows = 0 Then
      grdDetails.Columns("Details").Caption = "No details exist for this entry"
    Else
      grdDetails.Columns("Details").Caption = "Details (" & (grdDetails.AddItemRowIndex(grdDetails.FirstRow) + 1) & " Of " & grdDetails.Rows & " Entries)"
    End If

End Sub

Private Sub grdDetails_ScrollAfter()

  ' Update the grid caption after the user has used the scrollbar to view the details
    If grdDetails.Rows = 0 Then
      grdDetails.Columns("Details").Caption = "No details exist for this entry"
    Else
      grdDetails.Columns("Details").Caption = "Details (" & (grdDetails.AddItemRowIndex(grdDetails.FirstRow) + 1) & " Of " & grdDetails.Rows & " Entries)"
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    Unload Me
End Select
  
End Sub

'Private Sub PopulateBatchCombo()
'
'  ' Load all other jobs contained within the same batch as the originally selected
'  ' job. NB, in the order that the jobs are run in the batch
'  Dim prstTemp As Recordset
'
'  Set prstTemp = datGeneral.GetPersistentMainRecordset("SELECT ID,Name FROM ASRSysEventLog WHERE BatchRunID = " & mlngBatchRunID & " ORDER BY DateTime ASC")
'
'  Do Until prstTemp.EOF
'    'TM20011113 Fault 3134 - Take the left 40 chars from the field.
'    'So that a match can be made in SetComboText.
'    cboOtherJobs.AddItem Left(prstTemp.Fields("Name"), 40)
'    cboOtherJobs.ItemData(cboOtherJobs.NewIndex) = prstTemp.Fields("ID")
'    prstTemp.MoveNext
'  Loop
'
'  ' Set the combo position to be the originally selected job
'  SetComboText cboOtherJobs, Me.lblName.Caption
'
'  With cboOtherJobs
'    If .ListCount = 1 Then
'
'      ' RH 19/09/00 - BUG 962
'      '.Clear
'      '.AddItem "<None>"
'      .Enabled = False
'      .BackColor = &H8000000F
'      .ListIndex = 0
'    Else
'      .Enabled = True
'      .BackColor = &H80000005
'    End If
'  End With
'
'
'
'  Set prstTemp = Nothing
'
'End Sub





