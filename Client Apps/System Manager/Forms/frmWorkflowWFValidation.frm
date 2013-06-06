VERSION 5.00
Begin VB.Form frmWorkflowWFValidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Web Form Validation"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
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
   HelpContextID   =   5080
   Icon            =   "frmWorkflowWFValidation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraValidation 
      Height          =   1850
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   6400
      Begin VB.CommandButton cmdValidationExpression 
         Caption         =   "..."
         Height          =   315
         Left            =   5885
         TabIndex        =   3
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtValidationExpression 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   4590
      End
      Begin VB.TextBox txtMessage 
         Height          =   315
         Left            =   1290
         MaxLength       =   1000
         TabIndex        =   8
         Top             =   1300
         Width           =   4905
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Error"
         Height          =   195
         Index           =   0
         Left            =   1290
         TabIndex        =   5
         Top             =   860
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Warning"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   6
         Top             =   860
         Width           =   1170
      End
      Begin VB.Label lblValidationExpression 
         Caption         =   "Validation :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMessage 
         Caption         =   "Message :"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   1365
         Width           =   915
      End
      Begin VB.Label lblType 
         Caption         =   "Type :"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   855
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5300
      TabIndex        =   10
      Top             =   2100
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4035
      TabIndex        =   9
      Top             =   2100
      Width           =   1200
   End
End
Attribute VB_Name = "frmWorkflowWFValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfReadOnly As Boolean


Private mlngValidationExprID As Long

Private mlngUtilityID As Long
Private mlngBaseTableID As Long
Private miInitiationType As WorkflowInitiationTypes
Private maWFPrecedingElements() As VB.Control
Private maWFAllElements() As VB.Control

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property


Public Property Let Changed(ByVal pfNewValue As Boolean)
  mfChanged = pfNewValue
  RefreshScreen
  
End Property


Private Sub RefreshScreen()
  ' Refresh the screen controls.
  Dim fOKToSave As Boolean
  
  fOKToSave = mfChanged And (Not mfReadOnly)
  
  cmdOK.Enabled = fOKToSave

End Sub


Public Sub Initialise(pfReadOnly As Boolean, _
  plngExprID As Long, _
  piType As WorkflowWebFormValidationTypes, _
  psMessage As String, _
  plngUtilityID As Long, _
  plngUtilityBaseTable As Long, _
  piWorkflowInitiationType As WorkflowInitiationTypes, _
  pavPrecedingWorkflowElements As Variant, _
  pavAllWorkflowElements As Variant)

  mfReadOnly = pfReadOnly
  
  ValidationExprID = plngExprID
  txtValidationExpression.Text = GetExpressionName(mlngValidationExprID)
  
  ValidationType = piType
  
  Message = psMessage
  
  mlngUtilityID = plngUtilityID
  mlngBaseTableID = plngUtilityBaseTable
  miInitiationType = piWorkflowInitiationType
  maWFPrecedingElements = pavPrecedingWorkflowElements
  maWFAllElements = pavAllWorkflowElements
  
  If mfReadOnly Then
    ControlsDisableAll Me
    
    cmdValidationExpression.Enabled = True
  End If
  
  mfChanged = False
  RefreshScreen
  
End Sub

Private Sub cmdCancel_Click()
  If Me.Changed Then
    Select Case MsgBox("You have changed the definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title)
      Case vbYes
        cmdOK_Click
        Exit Sub
      Case vbCancel
        Exit Sub
    End Select
  End If
  
  Cancelled = True
  Me.Hide

End Sub

Private Sub cmdOK_Click()
  Dim fOK As Boolean
  
  fOK = True
  
  If Changed Then
    fOK = ValidateProperties
  End If
  
  If fOK Then
    ' Flag that the change/deletion has been confirmed.
    mfCancelled = False
  
    Me.Hide
  End If

End Sub


Private Function ValidateProperties() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fContinue As Boolean
  Dim frmUsage As frmUsage
  Dim asMessages() As String
  Dim iLoop As Integer

  fContinue = True
  ReDim asMessages(0)

  If ValidationExprID <= 0 Then
    ReDim Preserve asMessages(UBound(asMessages) + 1)
    asMessages(UBound(asMessages)) = "No calculation selected"
  End If
  
  If Len(Message) = 0 Then
    ReDim Preserve asMessages(UBound(asMessages) + 1)
    asMessages(UBound(asMessages)) = "No message"
  End If
  
  ' Display the validity failures to the user.
  fContinue = (UBound(asMessages) = 0)

  If Not fContinue Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList

    For iLoop = 1 To UBound(asMessages)
      frmUsage.AddToList (asMessages(iLoop))
    Next iLoop

    Screen.MousePointer = vbDefault
    frmUsage.ShowMessage "Workflow", "The Web Form validation definition is invalid for the reasons listed below." & _
      vbCrLf & "Do you wish to continue?", UsageCheckObject.Workflow, _
      USAGEBUTTONS_YES + USAGEBUTTONS_NO + USAGEBUTTONS_PRINT, "validation"

    fContinue = (frmUsage.Choice = vbYes)

    UnLoad frmUsage
    Set frmUsage = Nothing
  End If

TidyUpAndExit:
  ValidateProperties = fContinue
  Exit Function
  
ErrorTrap:
  fContinue = True
  Resume TidyUpAndExit
  
End Function


Private Sub cmdValidationExpression_Click()
  Dim objExpr As CExpression
  Dim lngOriginalID As Long

  lngOriginalID = mlngValidationExprID

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise 0, mlngValidationExprID, giEXPR_WORKFLOWCALCULATION, giEXPRVALUE_LOGIC
    .UtilityID = mlngUtilityID
    .UtilityBaseTable = mlngBaseTableID
    .WorkflowInitiationType = miInitiationType
    .PrecedingWorkflowElements = maWFPrecedingElements
    .AllWorkflowElements = maWFAllElements

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression(mfReadOnly) Then
      mlngValidationExprID = .ExpressionID
    Else
      ' Check in case the original expression has been deleted.
      If Not CheckExpression(mlngValidationExprID, 0, False) Then
        mlngValidationExprID = 0
      End If
    End If

    ' Read the selected expression info.
    txtValidationExpression.Text = GetExpressionName(mlngValidationExprID)
  End With

  Set objExpr = Nothing

  If lngOriginalID <> mlngValidationExprID Then
    Changed = True
  End If

End Sub

Private Function CheckExpression(plngExprID As Long, _
  plngTableID As Long, _
  pfCheckTable As Boolean) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If pfCheckTable And (plngTableID <= 0) Then
    fOK = False
  Else
    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", plngExprID, False

      If .NoMatch Then
        fOK = False
      Else
        If pfCheckTable _
          And !TableID <> plngTableID Then
          
          fOK = False
        End If
      End If
    End With
  End If
  
TidyUpAndExit:
  CheckExpression = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
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
  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    Cancel = True
  End If
  
End Sub

Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = mfCancelled
  
End Property

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  mfCancelled = pfNewValue
End Property


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub optType_Click(Index As Integer)
  Changed = True

End Sub

Private Sub txtMessage_Change()
  Changed = True

End Sub


Private Sub txtMessage_GotFocus()
  UI.txtSelText

End Sub



Public Property Get ValidationExprID() As Long
  ValidationExprID = mlngValidationExprID
  
End Property

Public Property Let ValidationExprID(ByVal plngNewValue As Long)
  mlngValidationExprID = plngNewValue
  
End Property

Public Property Get ValidationType() As WorkflowWebFormValidationTypes
  
  If optType(WORKFLOWWFVALIDATIONTYPE_WARNING).value Then
    ValidationType = WORKFLOWWFVALIDATIONTYPE_WARNING
  Else
    ValidationType = WORKFLOWWFVALIDATIONTYPE_ERROR
  End If
  
End Property

Public Property Let ValidationType(ByVal piNewValue As WorkflowWebFormValidationTypes)
  
  optType(WORKFLOWWFVALIDATIONTYPE_ERROR).value = (piNewValue = WORKFLOWWFVALIDATIONTYPE_ERROR)
  optType(WORKFLOWWFVALIDATIONTYPE_WARNING).value = (piNewValue = WORKFLOWWFVALIDATIONTYPE_WARNING)

End Property

Public Property Get Message() As String
  Message = txtMessage.Text
  
End Property

Public Property Let Message(ByVal psNewValue As String)
  txtMessage.Text = psNewValue
  
End Property

