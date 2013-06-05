VERSION 5.00
Begin VB.Form frmBatchJobJobSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Job : Job Selection"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1008
   Icon            =   "frmBatchJobJobSelection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4140
      TabIndex        =   5
      Top             =   2295
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2895
      TabIndex        =   4
      Top             =   2295
      Width           =   1200
   End
   Begin VB.Frame fraJobSelection 
      Height          =   2145
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   5235
      Begin VB.CheckBox chkOnlyMine 
         Caption         =   "On&ly show jobs where the owner is"
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   1590
         Width           =   4740
      End
      Begin VB.ComboBox cboJobName 
         Height          =   315
         Left            =   1860
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   705
         Width           =   3200
      End
      Begin VB.ComboBox cboJobType 
         Height          =   315
         ItemData        =   "frmBatchJobJobSelection.frx":000C
         Left            =   1860
         List            =   "frmBatchJobJobSelection.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   3200
      End
      Begin VB.TextBox txtParameter 
         Height          =   315
         Left            =   1860
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1110
         Width           =   3200
      End
      Begin VB.Label lblJobName 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Name :"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label lblJobType 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Type :"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pause Parameter :"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmBatchJobJobSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnCancelled As Boolean                         'Has The Form Been Cancelled
Private mclsData As DataMgr.clsDataAccess
Private mblnDefinitionCreator As Boolean
Private mstrPrevSelectedJob As String
Private msBatchJobHiddenGroups As String
Public Property Get DefinitionCreator() As Boolean
  DefinitionCreator = mblnDefinitionCreator
End Property
Public Property Let DefinitionCreator(ByVal bDefinitionCreator As Boolean)
  mblnDefinitionCreator = bDefinitionCreator
End Property
Public Property Let BatchJobHiddenGroups(ByVal psBatchJobHiddenGroups As String)
  msBatchJobHiddenGroups = UCase(psBatchJobHiddenGroups)
End Property
Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled                             'Expose value of cancelled flag
End Property
Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel                               'Set value of cancelled flag
End Property
Private Sub chkOnlyMine_Click()
  mstrPrevSelectedJob = Me.cboJobName.Text
  cboJobName.Clear                                      'Clear Previous JobNames/IDs
  'ChangeControlStatus True                            'Enable Jobname combo, disable paramter
  GetIndividualJobs
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
  Set mclsData = New DataMgr.clsDataAccess
  ' Screen Item chagnes for Reprot Pack Options
  If gblnReportPackMode Then Me.Caption = "Report Packs : Job Selection"
  Label1.Visible = Not gblnReportPackMode
  txtParameter.Visible = Not gblnReportPackMode
  Me.chkOnlyMine.Caption = Me.chkOnlyMine.Caption + " '" & gsUserName & "'"
  'Add Job Types To Combo -
  'doesnt matter what order you .additem: sorted at runtime if property cboCombo.Sorted is true
  With cboJobType
      .AddItem "Calendar Report"
      .AddItem "Career Progression"
      .AddItem "Cross Tab"
      .AddItem "Custom Report"
      .AddItem "Envelopes & Labels"
      .AddItem "Mail Merge"
      .AddItem "Match Report"
      .AddItem "Record Profile"
      .AddItem "Succession Planning"
      
      If Not gblnReportPackMode Then .AddItem "-- Pause --"
      If Not gblnReportPackMode Then .AddItem "Data Transfer"
      If Not gblnReportPackMode Then .AddItem "Export"
      If Not gblnReportPackMode Then .AddItem "Global Add"
      If Not gblnReportPackMode Then .AddItem "Global Delete"
      If Not gblnReportPackMode Then .AddItem "Global Update"
      If Not gblnReportPackMode Then .AddItem "Import"
      
      If gfAbsenceEnabled And Not gblnReportPackMode Then .AddItem "Absence Breakdown"
      If gfAbsenceEnabled And Not gblnReportPackMode Then .AddItem "Bradford Factor"
      If gfPersonnelEnabled And Not gblnReportPackMode Then .AddItem "Stability Index Report"
      If gfPersonnelEnabled And Not gblnReportPackMode Then .AddItem "Turnover Report"
      'Set combo to first item
      .ListIndex = 0
    End With
End Sub

Public Function Initialise(pstrJobType As String, pstrJobName As String, Optional pstrParameter As String) As Boolean

  SetComboText cboJobType, pstrJobType                  'Set JobType combo
  SetComboText cboJobName, pstrJobName                  'Set JobName combo
  
  If pstrParameter <> "" Then                           'Include Pause Parameter
    txtParameter.Text = pstrParameter
  End If
  
End Function

Private Sub cboJobType_Click()
  
  cboJobName.Clear                                      'Clear Previous JobNames/IDs
  
  'If cboJobType.Text = "-- Pause --" Then
  '  ChangeControlStatus False                           'Disable Jobname combo, enable parameter
  'Else
  '  CheckIfViewOnlyMine
  '  ChangeControlStatus True                            'Enable Jobname combo, disable paramter
  '  GetIndividualJobs
  'End If

  Select Case cboJobType.Text
  Case "-- Pause --"
    ChangeControlStatus False, True
    
  Case "Absence Breakdown", "Bradford Factor", _
         "Stability Index Report", "Turnover Report"
    ChangeControlStatus False, False
    
  Case Else
    CheckIfViewOnlyMine
    ChangeControlStatus True, False
    GetIndividualJobs
    
  End Select

End Sub

Private Sub ChangeControlStatus(blnJobCombo As Boolean, blnPause As Boolean)

  cboJobName.Enabled = blnJobCombo
  cboJobName.BackColor = IIf(blnJobCombo, vbWindowBackground, vbButtonFace)
  chkOnlyMine.Enabled = blnJobCombo
  If chkOnlyMine.Enabled = False Then
    chkOnlyMine.Value = vbUnchecked
  End If
  
  txtParameter.Enabled = blnPause
  txtParameter.BackColor = IIf(blnPause, vbWindowBackground, vbButtonFace)
  If Not blnPause Then
    txtParameter.Text = vbNullString
  End If

  Select Case cboJobType.Text
  Case "-- Pause --"
    
  Case "Absence Breakdown", "Bradford Factor", _
         "Stability Index Report", "Turnover Report"
    txtParameter.Text = "N/A"
    
  Case Else
    txtParameter.Text = "N/A"
    
  End Select
  

End Sub

Private Sub GetIndividualJobs()
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  Dim sBaseTableName As String
  Dim sBaseTableIDColumnName As String
  Dim sAccessTableName As String
  Dim sExtraClause As String
  
  ' We are the definition creator, so show our own hidden jobs
  Select Case UCase(cboJobType.Text)                    'Set the SQL depending on job type
    Case "CALENDAR REPORT"
      sBaseTableName = "ASRSysCalendarReports"
      sBaseTableIDColumnName = "ID"
      sAccessTableName = "ASRSysCalendarReportAccess"
    
    Case "CROSS TAB"
      sBaseTableName = "ASRSysCrossTab"
      sBaseTableIDColumnName = "CrossTabID"
      sAccessTableName = "ASRSysCrossTabAccess"
        
    Case "CUSTOM REPORT"
      sBaseTableName = "ASRSysCustomReportsName"
      sBaseTableIDColumnName = "ID"
      sAccessTableName = "ASRSysCustomReportAccess"
    
    Case "DATA TRANSFER"
      sBaseTableName = "ASRSysDataTransferName"
      sBaseTableIDColumnName = "dataTransferID"
      sAccessTableName = "ASRSysDataTransferAccess"
    
    Case "EXPORT"
      sBaseTableName = "ASRSysExportName"
      sBaseTableIDColumnName = "ID"
      sAccessTableName = "ASRSysExportAccess"
    
    Case "GLOBAL ADD"
      sBaseTableName = "ASRSysGlobalFunctions"
      sBaseTableIDColumnName = "functionID"
      sAccessTableName = "ASRSysGlobalAccess"
      sExtraClause = "(type = 'A')"

    Case "GLOBAL DELETE"
      sBaseTableName = "ASRSysGlobalFunctions"
      sBaseTableIDColumnName = "functionID"
      sAccessTableName = "ASRSysGlobalAccess"
      sExtraClause = "(type = 'D')"
            
    Case "GLOBAL UPDATE"
      sBaseTableName = "ASRSysGlobalFunctions"
      sBaseTableIDColumnName = "functionID"
      sAccessTableName = "ASRSysGlobalAccess"
      sExtraClause = "(type = 'U')"
    
    Case "IMPORT"
      sBaseTableName = "ASRSysImportName"
      sBaseTableIDColumnName = "ID"
      sAccessTableName = "ASRSysImportAccess"
      
    Case "MAIL MERGE"
      sBaseTableName = "ASRSysMailMergeName"
      sBaseTableIDColumnName = "mailMergeID"
      sAccessTableName = "ASRSysMailMergeAccess"
      sExtraClause = "(IsLabel = 0)"
   
    Case "ENVELOPES & LABELS"
      sBaseTableName = "ASRSysMailMergeName"
      sBaseTableIDColumnName = "mailMergeID"
      sAccessTableName = "ASRSysMailMergeAccess"
      sExtraClause = "(IsLabel = 1)"
   
    Case "MATCH REPORT", "SUCCESSION PLANNING", "CAREER PROGRESSION"
      sBaseTableName = "ASRSysMatchReportName"
      sBaseTableIDColumnName = "matchReportID"
      sAccessTableName = "ASRSysMatchReportAccess"
      
      Select Case UCase(cboJobType.Text)
        Case "MATCH REPORT"
          sExtraClause = "MatchReportType = 0"
        Case "SUCCESSION PLANNING"
          sExtraClause = "MatchReportType = 1"
        Case "CAREER PROGRESSION"
          sExtraClause = "MatchReportType = 2"
      End Select

    Case "RECORD PROFILE"
      sBaseTableName = "ASRSysRecordProfileName"
      sBaseTableIDColumnName = "recordProfileID"
      sAccessTableName = "ASRSysRecordProfileAccess"
      
    Case Else
      Exit Sub
  
  End Select
  
  ' Add the bit about only show mine !
  If Len(sBaseTableName) > 0 Then
    sSQL = "SELECT " & sBaseTableName & "." & sBaseTableIDColumnName & " AS id," & _
      sBaseTableName & ".name" & _
      " FROM " & sBaseTableName & _
      " INNER JOIN " & sAccessTableName & " ON " & sBaseTableName & "." & sBaseTableIDColumnName & " = " & sAccessTableName & ".ID" & _
      " INNER JOIN sysusers b ON " & sAccessTableName & ".groupname = b.name" & _
      "   AND b.name = '" & gsUserGroup & "'" & _
      " WHERE ((LOWER(" & sBaseTableName & ".userName) = '" & LCase(datGeneral.UserNameForSQL) & "')" & _
      "   OR (" & sAccessTableName & ".access <> '" & ACCESS_HIDDEN & "'))"

    If Len(sExtraClause) > 0 Then sSQL = sSQL & " AND (" & sExtraClause & ")"
    
    If Me.chkOnlyMine.Value Then sSQL = sSQL & " AND LOWER(" & sBaseTableName & ".userName) = '" & LCase(datGeneral.UserNameForSQL) & "'"
  
    sSQL = sSQL & " ORDER BY " & sBaseTableName & ".name"
  Else
    If Me.chkOnlyMine.Value Then sSQL = sSQL & " AND LOWER(Username) = '" & LCase(datGeneral.UserNameForSQL) & "'"
  
    sSQL = sSQL & " ORDER BY name"
  End If
  
  Set rsTemp = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If rsTemp.BOF And rsTemp.EOF Then
    cboJobName.Enabled = False
    cboJobName.BackColor = &H8000000F '&H80000005
    txtParameter.Enabled = False
    txtParameter.BackColor = &H8000000F '&H80000005
    Set rsTemp = Nothing
    Exit Sub
  End If
  
  cboJobName.Clear
  
  Do Until rsTemp.EOF
    cboJobName.AddItem rsTemp!Name
    cboJobName.ItemData(cboJobName.NewIndex) = rsTemp!ID
    rsTemp.MoveNext
  Loop
  
  If mstrPrevSelectedJob <> "" Then
    SetComboText Me.cboJobName, mstrPrevSelectedJob
    mstrPrevSelectedJob = ""
    
    'JPD 20030818 Fault 6695
    If cboJobName.ListCount > 0 Then
      If cboJobName.ListIndex < 0 Then
        cboJobName.ListIndex = 0
      End If
    End If
  Else
    cboJobName.ListIndex = 0
  End If
  
  Set rsTemp = Nothing

  'JPD 20030815 Fault 6696
  ChangeControlStatus True, False
  
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True                                 'Set Cancelled flag
  Unload Me                                        'Unload the form

End Sub

Private Sub cmdOK_Click()
  
  If Not ValidateCommands Then Exit Sub            'Exit if Type+Name not selected
  
  Cancelled = False                                'Set cancelled flag
  Me.Hide                                          'Hide form from view

End Sub

Private Function ValidateCommands() As Boolean

  Dim blnValid As Boolean

  ValidateCommands = False

  If cboJobType.ListIndex < 0 Then
    COAMsgBox "You must select a job type.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " Validation"
    Exit Function
  End If
  
'  'NHRD27102004 Fault 8240
'  If (cboJobType.Text Like "*Pause*") Then
'    If COAMsgBox("You have included -- Pause -- in the Batch Job this will require user interaction at run time." & vbCrLf & vbCrLf & "Continue?", vbExclamation + vbYesNo, "Batch Job Validation") = vbNo Then
'      Exit Function
'    Else
'      'Continue
'    End If
'  End If
  
  If JobTypeRequiresDef(cboJobType.Text) Then
    If cboJobName.ListIndex < 0 Then
      COAMsgBox "You must select a job name.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " Validation"
      Exit Function
    End If

    If Not OKToSelectThisJob(Me.cboJobType.Text, Me.cboJobName.ItemData(cboJobName.ListIndex)) Then
      Exit Function
    End If
  End If

  ValidateCommands = True

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode <> vbFormCode Then
    Cancelled = True
  End If
  
  Set mclsData = Nothing
  
End Sub

Private Function OKToSelectThisJob(pstrJobType As String, plngID As Long) As Boolean

  On Error GoTo Error_Trap
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  Dim sCurrentUserAccess As String
  Dim fNewAccess As Boolean
  Dim fUtilityIsHidden As Boolean
  Dim rsJobAccess As ADODB.Recordset
  Dim iUtilType As UtilityType
  Dim fWillChangeBatchJobAccess As Boolean
  
  fNewAccess = False
  fWillChangeBatchJobAccess = False
  
  sSQL = "SELECT Access, Username"
  
  Select Case UCase(pstrJobType)
  'Set the SQL depending on job type
  Case "CALENDAR REPORT"
    sSQL = "SELECT userName FROM ASRSysCalendarReports WHERE ID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlCalendarReport, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlCalendarReport, plngID)
    iUtilType = utlCalendarReport
        
  Case "CROSS TAB"
    sSQL = "SELECT userName FROM ASRSysCrossTab WHERE CrossTabID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlCrossTab, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlCrossTab, plngID)
    iUtilType = utlCrossTab
    
  Case "CUSTOM REPORT"
    sSQL = "SELECT userName FROM ASRSysCustomReportsName WHERE ID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlCustomReport, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlCustomReport, plngID)
    iUtilType = utlCustomReport
    
  Case "DATA TRANSFER"
    sSQL = "SELECT userName FROM ASRSysDataTransferName WHERE DataTransferID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlDataTransfer, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlDataTransfer, plngID)
    iUtilType = utlDataTransfer
    
  Case "EXPORT"
    sSQL = "SELECT userName FROM ASRSysExportName WHERE ID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlExport, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlExport, plngID)
    iUtilType = utlExport
    
  Case "GLOBAL ADD"
    sSQL = "SELECT userName FROM ASRSysGlobalFunctions WHERE functionID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(UtlGlobalAdd, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(UtlGlobalAdd, plngID)
    iUtilType = UtlGlobalAdd
    
  Case "GLOBAL DELETE"
    sSQL = "SELECT userName FROM ASRSysGlobalFunctions WHERE functionID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlGlobalDelete, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlGlobalDelete, plngID)
    iUtilType = utlGlobalDelete
    
  Case "GLOBAL UPDATE"
    sSQL = "SELECT userName FROM ASRSysGlobalFunctions WHERE functionID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlGlobalUpdate, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlGlobalUpdate, plngID)
    iUtilType = utlGlobalUpdate
    
  Case "IMPORT"
    sSQL = "SELECT userName FROM ASRSysImportName WHERE ID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlImport, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlImport, plngID)
    iUtilType = utlImport
    
  Case "MAIL MERGE"
    sSQL = "SELECT userName FROM ASRSysMailMergeName WHERE MailMergeID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlMailMerge, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlMailMerge, plngID)
    iUtilType = utlMailMerge
    
  Case "ENVELOPES & LABELS"
    sSQL = "SELECT userName FROM ASRSysMailMergeName WHERE MailMergeID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlLabel, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlLabel, plngID)
    iUtilType = utlLabel
    
  Case "MATCH REPORT"
    sSQL = "SELECT userName FROM ASRSysMatchReportName WHERE matchReportID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlMatchReport, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlMatchReport, plngID)
    iUtilType = utlMatchReport
      
  Case "SUCCESSION PLANNING"
    sSQL = "SELECT userName FROM ASRSysMatchReportName WHERE matchReportID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlSuccession, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlSuccession, plngID)
    iUtilType = utlSuccession
    
  Case "CAREER PROGRESSION"
    sSQL = "SELECT userName FROM ASRSysMatchReportName WHERE matchReportID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlCareer, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlCareer, plngID)
    iUtilType = utlCareer
    
  Case "RECORD PROFILE"
    sSQL = "SELECT userName FROM ASRSysRecordProfileName WHERE recordProfileID = " & plngID
    fNewAccess = True
    sCurrentUserAccess = CurrentUserAccess(utlRecordProfile, plngID)
    fUtilityIsHidden = UtilityIsHiddenToAnyone(utlRecordProfile, plngID)
    iUtilType = utlRecordProfile
  End Select

  Set rsTemp = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "Cannot select this job. It has been deleted by another user.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
    OKToSelectThisJob = False
    Exit Function
  End If
  
  If fNewAccess Then
    If (sCurrentUserAccess = ACCESS_HIDDEN) And _
      (LCase(rsTemp.Fields("Username")) <> LCase(gsUserName)) Then
      
      COAMsgBox "Cannot select this job. It has been made hidden by another user.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
      OKToSelectThisJob = False
      
    ElseIf fUtilityIsHidden And _
      (LCase(rsTemp.Fields("Username")) = LCase(gsUserName)) Then
      
      If mblnDefinitionCreator Then
        OKToSelectThisJob = True
      Else
        'JPD 20060929 Fault 11433
        ' Check if the Batch Job access does really need to be changed.
        Set rsJobAccess = GetUtilityAccessRecords(iUtilType, plngID, False)
        If Not rsJobAccess Is Nothing Then
          ' Add the user groups and their access on this definition to the access grid.
          With rsJobAccess
            Do While Not .EOF
              If !Access = ACCESS_HIDDEN Then
                If InStr(msBatchJobHiddenGroups, vbTab & UCase(!Name) & vbTab) = 0 Then
                  fWillChangeBatchJobAccess = True
                End If
              End If
              
              .MoveNext
            Loop
          
            .Close
          End With
        End If
        Set rsJobAccess = Nothing
        
        If fWillChangeBatchJobAccess Then
          COAMsgBox "Unable to add this job to the " & IIf(gblnReportPackMode, "pack", "batch") & " as it is hidden" & vbCrLf & _
                 "and you are not the owner of the " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
          OKToSelectThisJob = False
        Else
          OKToSelectThisJob = True
        End If
      End If
    Else
      OKToSelectThisJob = True
    End If
  Else
    If rsTemp.Fields("Access") = ACCESS_HIDDEN And LCase(rsTemp.Fields("Username")) <> LCase(gsUserName) Then
      COAMsgBox "Cannot select this job. It has been made hidden by another user.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
      OKToSelectThisJob = False
    ElseIf rsTemp.Fields("Access") = ACCESS_HIDDEN And LCase(rsTemp.Fields("Username")) = LCase(gsUserName) Then
      If mblnDefinitionCreator Then
        OKToSelectThisJob = True
      Else
        'JPD 20060929 Fault 11433
        ' Check if the Batch Job access does really need to be changed.
        Set rsJobAccess = GetUtilityAccessRecords(iUtilType, plngID, False)
        If Not rsJobAccess Is Nothing Then
          ' Add the user groups and their access on this definition to the access grid.
          With rsJobAccess
            Do While Not .EOF
              If !Access = ACCESS_HIDDEN Then
                If InStr(msBatchJobHiddenGroups, vbTab & UCase(!Name) & vbTab) = 0 Then
                  fWillChangeBatchJobAccess = True
                End If
              End If
              
              .MoveNext
            Loop
          
            .Close
          End With
        End If
        Set rsJobAccess = Nothing
        
        If fWillChangeBatchJobAccess Then
          COAMsgBox "Unable to add this job to the " & IIf(gblnReportPackMode, "pack", "batch") & " as it is hidden" & vbCrLf & _
                 "and you are not the owner of the " & IIf(gblnReportPackMode, "report pack", "batch job") & " definition.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Packs", "Batch Jobs")
          OKToSelectThisJob = False
        Else
          OKToSelectThisJob = True
        End If
      End If
    Else
      OKToSelectThisJob = True
    End If
  End If
  
  Set rsTemp = Nothing

  Exit Function
  
Error_Trap:
  
  OKToSelectThisJob = True
  Set rsTemp = Nothing
  
End Function


Private Sub CheckIfViewOnlyMine()

  Dim strType As String
  Dim intOnlyMine As Integer
  
  strType = Replace(UCase(cboJobType.Text), " ", "")

  'Needed to put in this except as this are inconsistant!
  Select Case strType
    Case "CROSSTAB", "CUSTOMREPORT", "CALENDARREPORT", "MATCHREPORT"
      strType = strType & "S"
    'NHRD19122006 Fault 10512
    Case "LABELSDEFINTION"
      strType = "LABELDEFINITION"
    Case "ENVELOPES&LABELS"
      strType = "LABELS"
      
  End Select

  intOnlyMine = GetUserSetting("defsel", "onlymine " & strType, 0)
  chkOnlyMine.Value = IIf(intOnlyMine = 1, vbChecked, vbUnchecked)

End Sub

Private Function JobTypeRequiresDef(strJobType As String) As Boolean

  JobTypeRequiresDef = _
      (strJobType <> "-- Pause --" And _
       strJobType <> "Absence Breakdown" And _
       strJobType <> "Bradford Factor" And _
       strJobType <> "Stability Index Report" And _
       strJobType <> "Turnover Report")

End Function



Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



