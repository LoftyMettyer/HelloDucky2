VERSION 5.00
Begin VB.Form frmWorkflowEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3870
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5052
   Icon            =   "frmWorkflowEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPictureClear 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4610
      MaskColor       =   &H000000FF&
      TabIndex        =   9
      ToolTipText     =   "Clear Icon"
      Top             =   2415
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.TextBox txtPicture 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1380
      TabIndex        =   7
      Top             =   2400
      Width           =   2860
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "..."
      Height          =   315
      Left            =   4280
      TabIndex        =   8
      ToolTipText     =   "Select Icon"
      Top             =   2415
      Width           =   315
   End
   Begin VB.TextBox txtURL 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   5
      Top             =   1900
      Width           =   4145
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "&Enabled"
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   2995
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2890
      TabIndex        =   11
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4325
      TabIndex        =   12
      Top             =   3300
      Width           =   1200
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1380
      MaxLength       =   255
      TabIndex        =   1
      Top             =   200
      Width           =   4145
   End
   Begin VB.TextBox txtDescription 
      Height          =   1015
      Left            =   1380
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   700
      Width           =   4145
   End
   Begin VB.Image picPicture 
      Height          =   495
      Left            =   5015
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   510
   End
   Begin VB.Label lblPicture 
      Caption         =   "Picture :"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   2460
      Width           =   750
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "URL :"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   1960
      Width           =   465
   End
   Begin VB.Label lblExpressionName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   260
      Width           =   510
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   760
      Width           =   1125
   End
End
Attribute VB_Name = "frmWorkflowEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mfLoading As Boolean
Private mlngWorkflowID As Long

Private mfChanged As Boolean
Private mbLocked As Boolean

Private mfSaveChanges As Boolean

Private msName As String
Private msDescription As String
Private mlngPictureID As Long
Private mfEnabled As Boolean
Private mfOriginalEnabled As Boolean
Private msURL As String
Private msUserName As String
Private msPassword As String

Private msExternalInitiationQueryString As String
Private miInitiationType As WorkflowInitiationTypes

Private mfrmCallingForm As Form

Private Const MIN_FORM_HEIGHT = 4290
Private Const MIN_FORM_WIDTH = 5885

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
End Property

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property

Public Property Let Changed(ByVal pfNewValue As Boolean)
  If Not mfLoading Then
    mfChanged = pfNewValue
    RefreshScreen
  End If
End Property

Public Property Let Locked(ByVal pbNewValue As Boolean)
  mbLocked = pbNewValue

  
  
End Property

Private Sub FormatScreen()
            
  lblURL.Visible = (miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL)
  txtURL.Visible = (miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL)

  txtDescription.Height = IIf(miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL, txtURL.Top - 185, txtPicture.Top - 200) - txtDescription.Top
   
End Sub

Private Function ValidateWorkflow() As Boolean
  On Error GoTo ErrorTrap
  
  Dim frmWFDes As frmWorkflowDesigner
  Dim fValid As Boolean
  
  fValid = True
  
  If Not mfrmCallingForm Is Nothing Then
    fValid = mfrmCallingForm.ValidateWorkflow(False, True, False)
  End If

TidyUpAndExit:
  ValidateWorkflow = fValid
  Exit Function
  
ErrorTrap:
  fValid = False
  Resume TidyUpAndExit
  
End Function

Public Property Get WorkflowName() As String
  WorkflowName = msName
End Property

Public Property Let WorkflowName(psNewValue As String)
  msName = psNewValue
End Property

Public Property Get WorkflowDescription() As String
  WorkflowDescription = msDescription
End Property

Public Property Let WorkflowDescription(psNewValue As String)
  msDescription = psNewValue
End Property

Public Property Get WorkflowPictureID() As Long
  WorkflowPictureID = mlngPictureID
End Property

Public Property Let WorkflowPictureID(plngNewValue As Long)
  mlngPictureID = plngNewValue
End Property
 
Private Sub RefreshScreen()
  ' Refresh the screen controls.
  Dim fReadOnly As Boolean
  
  fReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)
  
'  txtName.Enabled = (Not fReadOnly)
'  txtDescription.Enabled = (Not fReadOnly)
'  chkEnabled.Enabled = (Not fReadOnly)

  If fReadOnly Then
    ControlsDisableAll Me
  End If

  cmdOK.Enabled = mfChanged And (Not fReadOnly)

End Sub

Public Property Get ReadOnly() As Boolean
  ReadOnly = (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)
End Property

Public Property Let MustSaveChanges(pfNewValue As Boolean)
  mfSaveChanges = pfNewValue
  
End Property

Public Property Get WorkflowID() As Long
  WorkflowID = mlngWorkflowID
  
End Property

Public Property Let WorkflowID(pLngNewID As Long)
  mlngWorkflowID = pLngNewID
  
  mfSaveChanges = True
  RefreshScreen

End Property


Public Property Get Loading() As Boolean
  Loading = mfLoading
  
End Property


Private Function SaveChanges() As Boolean
  ' Save the changes.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim sSQL As String
  Dim rsInfo As DAO.Recordset
  Dim fTransStarted As Boolean
  Dim sQueryString As String
  
  fTransStarted = False

  fOK = True
  
  If (chkEnabled.value = vbChecked) _
    And (mfEnabled <> mfOriginalEnabled) Then
    ' Only allow the workflow to be enabled if its is valid!
    If Not ValidateWorkflow Then
      mfEnabled = False
      chkEnabled.value = vbUnchecked
    
      MsgBox "This workflow cannot be enabled as the definition is invalid.", vbInformation + vbOKOnly, App.ProductName
    End If
  End If
  
  If Not mfSaveChanges Then
    SaveChanges = True
    Exit Function
  End If
    
  If fOK Then
    ' Begin the transaction of data to the local database.
    daoWS.BeginTrans
    fTransStarted = True
  
    With recWorkflowEdit
      If WorkflowID > 0 Then
        ' Modify an existing screen record if it exists.
        .Index = "idxWorkflowID"
        .Seek "=", WorkflowID
  
        If Not .NoMatch Then
        
          Dim perge As Boolean
          
          If Trim(!Name) <> Trim(txtName.Text) Or _
             Trim(!Description) <> Trim(txtDescription.Text) Or _
             !Enabled <> (chkEnabled.value = vbChecked) Then
            perge = True
          End If

          If perge And Not mfrmCallingForm Is Nothing Then
            If mfrmCallingForm.Name = "frmWorkflowOpen" Then
              If WorkflowsWithStatus(WorkflowID, giWFSTATUS_COMPLETE) Or WorkflowsWithStatus(WorkflowID, giWFSTATUS_ERROR) Then

                fOK = (MsgBox("Saving these changes will purge all instances of this workflow from the log." & vbCrLf & _
                              "Do you wish to continue?", vbQuestion + vbYesNo, App.ProductName) = vbYes)
              End If
            End If
          End If
          If Not fOK Then GoTo TidyUpAndExit
          
          .Edit
          !Name = Trim(txtName.Text)
          !Description = Trim(txtDescription.Text)
          !PictureID = IIf(mlngPictureID = 0, Null, mlngPictureID)
          !Enabled = (chkEnabled.value = vbChecked)
          !Changed = True
          !perge = !perge Or perge
          .Update
        End If
        
        If (mfEnabled <> mfOriginalEnabled) _
          And (!InitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) Then
          ' Check if it's used in any links.
          sSQL = "SELECT COUNT(*) AS recCount" & _
            " FROM tmpWorkflowTriggeredLinks" & _
            " WHERE tmpWorkflowTriggeredLinks.deleted = FALSE" & _
            " AND tmpWorkflowTriggeredLinks.workflowID = " & CStr(WorkflowID)
          Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          If rsInfo!reccount > 0 Then
            Application.ChangedWorkflowLink = True
          End If
          rsInfo.Close
          Set rsInfo = Nothing
        End If
      Else
        ' Create and initialise a new Workflow record if required.
        WorkflowID = UniqueColumnValue("tmpWorkflows", "ID")
        .AddNew
        !ID = WorkflowID
        !Name = Trim(txtName.Text)
        !Description = Trim(txtDescription.Text)
        !PictureID = IIf(mlngPictureID = 0, Null, mlngPictureID)
        !GUID = "{" & Mid(CreateGUID(), 2, 36) & "}"
        
        If miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
          sQueryString = GetWorkflowQueryString(WorkflowID * -1, -1)
        End If
        !queryString = sQueryString
        
        !Enabled = (chkEnabled.value = vbChecked)
  
        !New = True
        !Changed = False
        !perge = False
        !Deleted = False
        .Update
      End If
    End With
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    If fTransStarted Then
      daoWS.CommitTrans dbForceOSFlush
    End If
    Application.Changed = True
  Else
    If fTransStarted Then
      daoWS.Rollback
    End If
  End If
  
  SaveChanges = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property


Private Sub chkEnabled_Click()
  Dim fOK As Boolean
  
  fOK = True
  
  If (mfEnabled) _
    And (chkEnabled.value = vbUnchecked) _
    And (miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL) Then
    
    fOK = (MsgBox("The '" & Trim(txtName.Text) & "' workflow may be referenced externally." & vbNewLine & "Are you sure you want to disable it?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes)
  End If
  
  If fOK Then
    mfEnabled = (chkEnabled.value = vbChecked)
    Changed = True
  Else
    chkEnabled.value = vbChecked
  End If
  
End Sub

Private Sub cmdCancel_Click()
  mfCancelled = True
  
  UnLoad Me

End Sub

Private Sub cmdOK_Click()
  ' Validate and save the changes.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  ' Check that a workflow name has been entered.
  fOK = (Len(Trim(txtName.Text)) > 0)
  If Not fOK Then
    MsgBox "Invalid workflow name.", vbOKOnly + vbExclamation, Application.Name
    txtName.SetFocus
  End If

  ' Check that the workflow name entered is unique.
  If fOK Then
    With recWorkflowEdit
      .Index = "idxName"
      .Seek "=", Trim(txtName.Text), False
      If Not .NoMatch Then
        Do While (Not .EOF) And fOK
          If (!Name <> Trim(txtName.Text)) Or (!Deleted) Then
            Exit Do
          End If

          fOK = (!ID = WorkflowID)
          If Not fOK Then
            MsgBox "A workflow named '" & Trim(txtName.Text) & "' already exists!", vbOKOnly + vbExclamation, Application.Name
            txtName.SetFocus
          End If

          .MoveNext
        Loop
      End If
    End With
  End If
  
  If fOK Then
    fOK = SaveChanges
  End If

  If fOK Then
    mfCancelled = False
    UnLoad Me
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub Form_Activate()
  Dim sURL As String
  Dim sQueryString As String
  
  Const DFLT_HEIGHT = 3300
  
  If Loading Then
    
    sURL = GetWorkflowURL
    
    If (Len(msName) > 0) Then
      txtName.Text = msName
      txtDescription.Text = msDescription
      txtURL.Text = IIf((Len(sURL) > 0) And (Len(msExternalInitiationQueryString) > 0), sURL & "?" & msExternalInitiationQueryString, "<undefined>")
      chkEnabled.value = IIf(mfEnabled, vbChecked, vbUnchecked)
      Me.Caption = "Properties - " & msName
    
    ElseIf WorkflowID > 0 Then
      With recWorkflowEdit

        .Index = "idxWorkflowID"
        .Seek "=", WorkflowID

        If Not .NoMatch Then
          txtName.Text = Trim(.Fields("name"))
          txtDescription.Text = Trim(.Fields("description"))
          mlngPictureID = IIf(IsNull(.Fields("PictureID")), 0, .Fields("PictureID"))
          txtURL.Text = IIf((Len(sURL) >= 0) And (Len(Trim(.Fields("queryString"))) > 0), sURL & "?" & Trim(.Fields("queryString")), "<undefined>")
          chkEnabled.value = IIf(.Fields("enabled"), vbChecked, vbUnchecked)
          Me.Caption = "Properties - " & Trim(.Fields("name"))
        End If
      End With
      
    Else
      Me.Caption = "Properties - New Workflow"
      txtURL.Text = "<undefined>"
      chkEnabled.value = vbUnchecked
      chkEnabled.Enabled = False
    End If

    RefreshPictureControls

   'Set focus to column name textbox
    If txtName.Enabled Then
      txtName.SetFocus
    End If

    mfLoading = False
  End If

End Sub
Private Sub Form_Initialize()
  mfLoading = True
  mfCancelled = True

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
  Dim objControl As VB.Control

  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT

  For Each objControl In Me.Controls
    If TypeOf objControl Is TextBox Then
      objControl.Text = vbNullString
    End If
  Next
  Set objControl = Nothing

  cmdOK.Enabled = False
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim iAnswer As Integer
  
  If UnloadMode <> vbFormCode Then

    'Check if any changes have been made.
    If mfChanged Then
      iAnswer = MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If iAnswer = vbYes Then
        Call cmdOK_Click
        If Me.Cancelled Then Cancel = 1
      ElseIf iAnswer = vbNo Then
        Me.Cancelled = True
      ElseIf iAnswer = vbCancel Then
        Cancel = 1
      End If
    Else
      Me.Cancelled = True
    End If
  End If

End Sub

Public Property Set CallingForm(pfrmForm As Form)
  Set mfrmCallingForm = pfrmForm
  
End Property

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

  FormatScreen

End Sub


Private Sub Form_Unload(Cancel As Integer)

  Unhook Me.hWnd

End Sub


Private Sub txtDescription_Change()
  msDescription = Trim(txtDescription.Text)
  Changed = True

End Sub

Private Sub txtDescription_GotFocus()
  ' Select the whole string.
  UI.txtSelText
  cmdOK.Default = False

End Sub


Private Sub txtDescription_LostFocus()
  cmdOK.Default = True

End Sub

Private Sub txtName_Change()
  msName = Trim(txtName.Text)
  Changed = True

End Sub

Private Sub txtName_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub



Public Property Get WorkflowEnabled() As Boolean
  WorkflowEnabled = mfEnabled

End Property

Public Property Let WorkflowEnabled(ByVal pfNewValue As Boolean)
  mfEnabled = pfNewValue
  mfOriginalEnabled = pfNewValue
  
End Property

Public Property Get ExternalInitiationQueryString() As String
  ExternalInitiationQueryString = msExternalInitiationQueryString
  
End Property

Public Property Let ExternalInitiationQueryString(ByVal psNewValue As String)
  msExternalInitiationQueryString = psNewValue
  
End Property

Public Property Get InitiationType() As WorkflowInitiationTypes
  InitiationType = miInitiationType
  
End Property

Public Property Let InitiationType(ByVal piNewValue As WorkflowInitiationTypes)
  miInitiationType = piNewValue

End Property

Private Sub txtURL_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub

Private Sub cmdPicture_Click()

  Dim lngOriginalID As Long
  
  lngOriginalID = mlngPictureID
  
  frmPictSel.SelectedPicture = mlngPictureID
  frmPictSel.ExcludedExtensions = ""
  frmPictSel.Show vbModal
  
  mlngPictureID = frmPictSel.SelectedPicture
  RefreshPictureControls

  If lngOriginalID <> mlngPictureID Then
    Changed = True
  End If
End Sub

Private Sub cmdPictureClear_Click()
  mlngPictureID = 0
  RefreshPictureControls
  Changed = True
End Sub

Private Sub RefreshPictureControls()
  ' Refresh the Picture controls depending on the selected picture.
  Dim sFileName As String

  If mlngPictureID > 0 Then
    With recPictEdit
      .Index = "idxID"
      .Seek "=", mlngPictureID
      
      If Not .NoMatch Then
        txtPicture.Text = !Name
        sFileName = ReadPicture
        picPicture.Picture = LoadPicture(sFileName)
        Kill sFileName
      Else
        mlngPictureID = 0
      End If
    End With
  End If

  If mlngPictureID = 0 Then
    picPicture.Picture = LoadPicture("")
    txtPicture.Text = ""
  End If
  
  cmdPictureClear.Enabled = (mlngPictureID > 0) And (Not ReadOnly)
  
End Sub


