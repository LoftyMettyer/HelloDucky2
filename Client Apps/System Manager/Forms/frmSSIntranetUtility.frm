VERSION 5.00
Begin VB.Form frmSSIntranetUtility 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Chart Utility"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSSIntranetUtility.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5250
      TabIndex        =   3
      Top             =   1530
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   400
      Left            =   3975
      TabIndex        =   2
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Frame fraHRProUtilityLink 
      Caption         =   "OpenHR Report / Utility :"
      Height          =   1260
      Left            =   180
      TabIndex        =   4
      Top             =   135
      Width           =   6300
      Begin VB.ComboBox cboHRProUtilityType 
         Height          =   315
         ItemData        =   "frmSSIntranetUtility.frx":000C
         Left            =   1400
         List            =   "frmSSIntranetUtility.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   4700
      End
      Begin VB.ComboBox cboHRProUtility 
         Height          =   315
         Left            =   1400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   4700
      End
      Begin VB.Label lblHRProUtilityType 
         Caption         =   "Type :"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblHRProUtility 
         Caption         =   "Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   765
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSSIntranetUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean
Private mblnRefreshing As Boolean
Private mfChanged As Boolean

Public Sub Initialize(psUtilityType As String, psUtilityID As String)
  
  Dim mfChanged As Boolean
  
  mfChanged = False
  
  GetHRProUtilityTypes
  UtilityType = CStr(utlNineBoxGrid)
  
  UtilityType = psUtilityType
  UtilityID = psUtilityID

  cmdOk.Enabled = (val(psUtilityID) = 0)
  
  'RefreshControls
End Sub



Private Sub RefreshControls()

  Dim sUtilityMessage As String
  
  If mblnRefreshing Then Exit Sub
  
  sUtilityMessage = ""


' For Workflows only, check if the selected Workflow is enabled.
' Display a message as required.
'JPD 20060714 Fault 11226
If cboHRProUtilityType.ListIndex >= 0 Then
  If cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex) = utlWorkflow Then
    If cboHRProUtility.ListCount > 0 Then
      recWorkflowEdit.Index = "idxWorkflowID"
      recWorkflowEdit.Seek "=", cboHRProUtility.ItemData(cboHRProUtility.ListIndex)

      If Not recWorkflowEdit.NoMatch Then
        If (Not recWorkflowEdit.Fields("enabled").value) Then
          sUtilityMessage = "This Workflow is not currently enabled."
        End If
      End If
    End If
  End If
End If

cmdOk.Enabled = True

End Sub




Private Sub GetHRProUtilityTypes()

  ' Populate the Utility Types combo.
  Dim iDefaultItem As Integer
  
  iDefaultItem = 0
  
  With cboHRProUtilityType
    .Clear
      
    If ASRDEVELOPMENT Or Application.NineBoxGridModule Then
      .AddItem "9-Box Grid Report"
      .ItemData(.NewIndex) = utlNineBoxGrid
    End If
    
    .AddItem "Calendar Report"
    .ItemData(.NewIndex) = utlCalendarreport
        
    .AddItem "Custom Report"
    .ItemData(.NewIndex) = utlCustomReport
    
    .AddItem "Mail Merge"
    .ItemData(.NewIndex) = utlMailMerge
    
    .AddItem "Talent Report"
    .ItemData(.NewIndex) = utlTalent
    
    If ASRDEVELOPMENT Or Application.WorkflowModule Then
      .AddItem "Workflow"
      .ItemData(.NewIndex) = utlWorkflow
    End If
    
    .ListIndex = iDefaultItem
  End With
 
End Sub


Private Function ValidateLink() As Boolean
  Dim fValid As Boolean

      If cboHRProUtility.ListIndex < 0 Then
        fValid = False
        MsgBox "No " & cboHRProUtilityType.List(cboHRProUtilityType.ListIndex) & " has been selected.", vbOKOnly + vbExclamation, Application.Name
        cboHRProUtilityType.SetFocus
      Else
        fValid = True
      End If

ValidateLink = fValid

End Function

Private Sub cboHRProUtility_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProUtilityType_Click()
  GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
  mfChanged = True
  RefreshControls
End Sub

Private Sub GetHRProUtilities(pUtilityType As UtilityType)

  ' Populate the utilities combo.
  Dim sSQL As String
  Dim sWhereSQL As String
  Dim rsUtilities As New ADODB.Recordset
  Dim rsLocalUtilities As DAO.Recordset
  Dim sTableName As String
  Dim sIDColumnName As String
  Dim fLocalTable As Boolean
  
  fLocalTable = False
  
  cboHRProUtility.Clear

  Select Case pUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sIDColumnName = "ID"
      
    Case utlCalendarreport
      sTableName = "ASRSysCalendarReports"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sIDColumnName = "CrossTabID"
      sWhereSQL = "ASRSysCrossTab.CrossTabType <> " & ctt9GridBox
      
    Case utlNineBoxGrid
        sTableName = "ASRSysCrossTab"
        sIDColumnName = "CrossTabID"
        sWhereSQL = "ASRSysCrossTab.CrossTabType = " & ctt9GridBox
        
    
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sIDColumnName = "ID"
    
    Case utlDataTransfer
      sTableName = "ASRSysDataTransferName"
      sIDColumnName = "DataTransferID"
      
    Case utlExport
      sTableName = "ASRSysExportName"
      sIDColumnName = "ID"
      
    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTableName = "ASRSysGlobalFunctions"
      sIDColumnName = "functionID"

    Case utlImport
      sTableName = "ASRSysImportName"
      sIDColumnName = "ID"

    Case utlLabel
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
      sWhereSQL = "ASRSysMailMergeName.IsLabel = 1 "
    
    Case utlMailMerge
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
      sWhereSQL = "ASRSysMailMergeName.IsLabel = 0 "

    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sIDColumnName = "recordProfileID"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTableName = "ASRSysMatchReportName"
      sIDColumnName = "matchReportID"
  
    Case utlTalent
      sTableName = "ASRSysTalentReports"
      sIDColumnName = "ID"
  
    Case utlWorkflow
      fLocalTable = True
      sTableName = "tmpWorkflows"
      sIDColumnName = "ID"
      sWhereSQL = "tmpWorkflows.initiationType = " & CStr(WORKFLOWINITIATIONTYPE_MANUAL) & _
        " OR tmpWorkflows.initiationType is null"
  End Select
  
  If Len(sTableName) > 0 Then
    ' Get the available utilities of the given type.
    If fLocalTable Then
      sSQL = "SELECT " & sTableName & "." & sIDColumnName & " AS [ID], " & sTableName & ".name" & _
        " FROM " & sTableName & _
        " WHERE (" & sTableName & ".deleted = FALSE)"
      If Len(sWhereSQL) > 0 Then
        sSQL = sSQL & _
          " AND (" & sWhereSQL & ")"
      End If
      sSQL = sSQL & _
        " ORDER BY " & sTableName & ".name"
      Set rsLocalUtilities = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsLocalUtilities.EOF
        cboHRProUtility.AddItem rsLocalUtilities!Name
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsLocalUtilities!ID

        rsLocalUtilities.MoveNext
      Wend
      rsLocalUtilities.Close
      Set rsLocalUtilities = Nothing
    Else
      sSQL = "SELECT" & _
        "  " & sTableName & ".name," & _
        "  " & sTableName & "." & sIDColumnName & " AS [ID]" & _
        "  FROM " & sTableName
      If Len(sWhereSQL) > 0 Then
        sSQL = sSQL & _
          " WHERE (" & sWhereSQL & ")"
      End If
      sSQL = sSQL & _
        "  ORDER BY " & sTableName & ".name"
      
      rsUtilities.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      While Not rsUtilities.EOF
        cboHRProUtility.AddItem rsUtilities!Name
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsUtilities!ID
  
        rsUtilities.MoveNext
      Wend
      rsUtilities.Close
      Set rsUtilities = Nothing
    End If
  End If
  
  If cboHRProUtility.ListCount > 0 Then
    cboHRProUtility.ListIndex = 0
  End If

End Sub


Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Private Sub cmdCancel_Click()
  Cancelled = True
  UnLoad Me
End Sub

Private Sub cmdOK_Click()
  If ValidateLink Then
    Cancelled = False
    Me.Hide
  End If
End Sub

Public Property Get UtilityType() As String

  If (cboHRProUtility.ListIndex < 0) Then
    UtilityType = ""
  Else
    UtilityType = CStr(cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex))
  End If

End Property

Public Property Get UtilityID() As String

  If (cboHRProUtility.ListIndex < 0) Then
    
    UtilityID = ""
  Else
    UtilityID = CStr(cboHRProUtility.ItemData(cboHRProUtility.ListIndex))
  End If

End Property

Public Property Let UtilityID(ByVal psNewValue As String)

  Dim iLoop As Integer
  
  If (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtility.ListCount - 1
      If cboHRProUtility.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtility.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
End Property

Public Property Let UtilityType(ByVal psNewValue As String)

  Dim iLoop As Integer

  If (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtilityType.ListCount - 1
      If cboHRProUtilityType.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtilityType.ListIndex = iLoop
        Exit For
      End If
    Next iLoop

    GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
  End If
    
End Property


