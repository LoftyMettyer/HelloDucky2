VERSION 5.00
Begin VB.Form frmEmailAddr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Address"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5015
   Icon            =   "frmEmailAddr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2445
      TabIndex        =   11
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3720
      TabIndex        =   12
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   4780
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         MaxLength       =   30
         TabIndex        =   13
         Text            =   "Mike"
         Top             =   960
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1600
         MaxLength       =   50
         TabIndex        =   2
         Top             =   300
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   885
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   100
      TabIndex        =   3
      Top             =   1000
      Width           =   4780
      Begin VB.CommandButton cmdCalculated 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4300
         TabIndex        =   10
         Top             =   1100
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.ComboBox cboColumn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   700
         Width           =   3045
      End
      Begin VB.TextBox txtCalculated 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         TabIndex        =   9
         Top             =   1100
         Width           =   2700
      End
      Begin VB.TextBox txtFixed 
         Height          =   315
         Left            =   1600
         MaxLength       =   50
         TabIndex        =   5
         Top             =   300
         Width           =   3045
      End
      Begin VB.OptionButton optCalculated 
         Caption         =   "Calculated"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   1160
         Width           =   1200
      End
      Begin VB.OptionButton optColumn 
         Caption         =   "Column"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   760
         Width           =   1200
      End
      Begin VB.OptionButton optFixed 
         Caption         =   "Fixed"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmEmailAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean
Private mobjEmail As clsEmailAddr
Private mfChanged As Boolean
Private mblnLoading  As Boolean
Private pintAnswer As Integer
Private mblnReadOnly As Boolean


Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property


Public Property Let Changed(pChanged As Boolean)
  mfChanged = pChanged
End Property

Public Property Get Email() As clsEmailAddr
  Set Email = mobjEmail
End Property

Public Property Let Email(ByVal objNewValue As clsEmailAddr)
  Set mobjEmail = objNewValue
End Property

Private Sub cboColumn_Change()

  If Not mblnLoading Then mfChanged = True
  
End Sub

Private Sub cboColumn_Click()

  If Not mblnLoading Then mfChanged = True

End Sub

Private Sub cmdCalculated_Click()
  
  ' Display the Record Description selection form.
  Dim fOK As Boolean
  Dim objExpr As CExpression
  Dim lngExprID As Long
  
  lngExprID = val(txtCalculated.Tag)
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    fOK = .Initialise(mobjEmail.TableID, lngExprID, giEXPR_EMAIL, giEXPRVALUE_CHARACTER)
  
    If fOK Then
      ' Instruct the expression object to display the
      ' expression selection form.
      If .SelectExpression Then
        lngExprID = .ExpressionID
        ' Read the selected expression info.
        txtCalculated.Text = GetExpressionName(lngExprID)
        txtCalculated.Tag = lngExprID
      Else
        ' Check in case the original expression has been deleted.
        With recExprEdit
          .Index = "idxExprID"
          .Seek "=", lngExprID, False
          If Not .NoMatch Then
            txtCalculated.Text = !Name
            txtCalculated.Tag = lngExprID
          Else
            txtCalculated.Text = vbNullString
            txtCalculated.Tag = 0
          End If
        End With
      End If
    End If
  End With
  
  ' Disassociate object variables.
  Set objExpr = Nothing

End Sub

Private Sub cmdCancel_Click()

  mblnCancelled = True

  If mblnCancelled Then
    If mfChanged And cmdOK.Enabled Then
      
      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
        
      If pintAnswer = vbYes Then
        cmdOK_Click
      ElseIf pintAnswer = vbCancel Then
        Exit Sub
      ElseIf pintAnswer = vbNo Then
        UnLoad Me
      End If
    
    Else
      UnLoad Me
    End If
  End If
  
End Sub

Private Sub Form_Activate()

  mblnLoading = False

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

  mblnLoading = True
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)
                  
  If mblnReadOnly Then
    ControlsDisableAll Me
    cmdCalculated.Enabled = True
  Else
    If IsUsedInEmailGroup(mobjEmail.EmailID) Then
      optCalculated.Enabled = False
      optColumn.Enabled = False
      optFixed.Enabled = False
      MsgBox "This email definition is currently used in one or more email groups and therefore cannot be changed from a fixed value.", vbInformation, Me.Caption
    End If
  End If

  mfChanged = False
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = vbFormControlMenu Then
    mblnCancelled = True
    cmdCancel_Click
    If pintAnswer = vbNo Then
      Cancel = False
    ElseIf pintAnswer = vbYes Or pintAnswer = vbCancel Then
      Cancel = True
    End If
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Terminate()
  Set mobjEmail = Nothing
End Sub

Private Sub optCalculated_Click()
  Call OptionClick(2)
End Sub

Private Sub optColumn_Click()
  Call OptionClick(1)
End Sub

Private Sub optFixed_Click()
  Call OptionClick(0)
End Sub


Private Sub OptionClick(intMode As Integer)

  If intMode = 0 Then
    If Not mblnReadOnly Then
      txtFixed.Enabled = True
      txtFixed.BackColor = vbWindowBackground
    End If
  Else
    txtFixed.Text = vbNullString
    txtFixed.Enabled = False
    txtFixed.BackColor = vbButtonFace
  End If

  If intMode = 1 Then
    Call PopulateColumns
  Else
    cboColumn.Clear
    cboColumn.Enabled = False
    cboColumn.BackColor = vbButtonFace
  End If

  If intMode = 2 Then
    'cmdCalculated.Enabled = true
    cmdCalculated.Enabled = (Not mblnReadOnly)
  Else
    txtCalculated = vbNullString
    txtCalculated.Tag = 0
    cmdCalculated.Enabled = False
  End If

  If Not mblnLoading Then mfChanged = True
  
End Sub


Private Sub PopulateColumns()

  Dim rsTemp As Recordset
  Dim strSQL As String

  'ColumnType <> 3  Exclude ID Columns
  'DataType = 12    Only allow string columns
  strSQL = "SELECT tmpColumns.columnID, tmpColumns.columnName" & _
           " FROM tmpColumns" & _
           " WHERE tmpColumns.ColumnType <> 3" & _
           " AND tmpColumns.DataType = 12" & _
           " AND tmpColumns.tableID = " & CStr(mobjEmail.TableID)
  Set rsTemp = daoDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)

  With cboColumn

    .Clear
    Do While Not rsTemp.EOF
      .AddItem rsTemp!ColumnName
      .ItemData(.NewIndex) = rsTemp!ColumnID
      rsTemp.MoveNext
    Loop

    'Maybe put in next line?
    'optColumn.Enabled = (.ListCount > 0)
    If Not mblnReadOnly Then
      .Enabled = (.ListCount > 1) And (mobjEmail.TableID > 0)
      .BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
    End If

    If .ListCount > 0 Then .ListIndex = 0

  End With

  rsTemp.Close
  Set rsTemp = Nothing

End Sub

Private Sub txtCalculated_Change()

  If Not mblnLoading Then mfChanged = True
  
End Sub

Private Sub txtFixed_Change()

  If Not mblnLoading Then mfChanged = True
  
End Sub

Private Sub txtFixed_GotFocus()
  With txtFixed
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtName_Change()

  If Not mblnLoading Then mfChanged = True
  
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub


Public Function Initialise(objEmail As clsEmailAddr)

  Set mobjEmail = New clsEmailAddr
  Set mobjEmail = objEmail
  
'  If mobjEmail.EmailID = 0 Then
'    txtName = vbNullString
'    Call OptionClick(0)
'
'    optFixed.Value = True
'    txtFixed.Text = vbNullString
'
'  Else

  txtName = Trim(mobjEmail.EmailName)
  Call OptionClick(mobjEmail.EmailType)

  Select Case mobjEmail.EmailType
    Case 0
      optFixed.value = True
      txtFixed.Text = Trim(mobjEmail.Fixed)
  
    Case 1
      optColumn.value = True
      SetComboItem cboColumn, mobjEmail.ColumnID
      
    Case 2
      optCalculated.value = True
      txtCalculated.Tag = mobjEmail.ExpressionID
      txtCalculated.Text = GetExpressionName(mobjEmail.ExpressionID)
  End Select

'NPG20071121 Fault 12619

      If Not IsUsedInEmailGroup(mobjEmail.EmailID) Then
        optColumn.Enabled = (mobjEmail.TableID > 0) And (Not mblnReadOnly)
        optCalculated.Enabled = (mobjEmail.TableID > 0) And (Not mblnReadOnly)
      End If
'  End If

End Function

Private Sub cmdOK_Click()
  ' Confirm the Email.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iSequence As Integer
  Dim objNode As ComctlLib.Node
  
  fOK = True
  
  If mfChanged Then
    ' Reset the Cancelled property.
    mblnCancelled = False

    fOK = ValidEmail
    If fOK Then SaveEmail
  
  Else
    mblnCancelled = True
  End If
  
TidyUpAndExit:
  If fOK Then
    ' Unload the form.
    UnLoad Me
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
 End Sub


Private Function ValidEmail() As Boolean

  ValidEmail = False

  If Len(Trim(txtName.Text)) = 0 Then
    MsgBox "Please enter a valid email name.", vbOKOnly + vbExclamation, Application.Name
    txtName.SetFocus
    Exit Function
  End If

  If optFixed.value = True Then
    If Len(Trim(txtFixed.Text)) = 0 Then
      MsgBox "Please enter a fixed value.", vbOKOnly + vbExclamation, Application.Name
      txtFixed.SetFocus
      Exit Function
    End If

'    'MH20030326 Fault 5078
'    If InStr(txtFixed.Text, "'") > 0 Then
'      MsgBox "The email address cannot contain apostrophes.", vbOKOnly + vbExclamation, Application.Name
'      txtFixed.SetFocus
'      Exit Function
'    End If


    'MH20010228 Fault 1933
    'Phil asked us to remove this validation so that internal addresses could be used.
    
    'If Not ValidEmailAddress(txtFixed.Text) Then
    '  MsgBox "Please enter a valid email address.", vbOKOnly + vbExclamation, Application.Name
    '  txtFixed.SetFocus
    '  Exit Function
    'End If

  ElseIf optCalculated.value = True Then
    If val(txtCalculated.Tag) < 1 Then
      MsgBox "Please select an email calculation.", vbOKOnly + vbExclamation, Application.Name
      cmdCalculated.SetFocus
      Exit Function
    End If
  End If
  
  
  ' Check that there are no other Emails for this table with this name.
  With recEmailAddrEdit
    '.Index = "idxTableID"
    '.Seek "=", Email.TableID

    'If Not .NoMatch Then
    If Not (.BOF And .EOF) Then
     .MoveFirst
      Do While Not .EOF
        
        If !TableID = Email.TableID Or !TableID = 0 Then
          If (Trim(!Name) = Trim(txtName.Text)) And _
            (!EmailID <> Email.EmailID) And _
            (!Deleted = False) Then
              MsgBox "An email named '" & Trim(txtName.Text) & "' already exists !", vbOKOnly + vbExclamation, Application.Name
            Exit Function
          End If
        End If
  
        .MoveNext
      Loop
    End If
  End With

  ValidEmail = True

End Function


Private Function SaveEmail() As Boolean
    
  ' Write the changes to the Email object.
  With mobjEmail

    .EmailName = Trim(txtName.Text)

    If optFixed.value = True Then
      .EmailType = 0
      .Fixed = Trim(txtFixed.Text)
      .ColumnID = 0
      .ExpressionID = 0
      .TableID = 0

    ElseIf optColumn.value = True Then
      .EmailType = 1
      .Fixed = vbNullString
      .ColumnID = cboColumn.ItemData(cboColumn.ListIndex)
      .ExpressionID = 0

    Else
      .EmailType = 2
      .Fixed = vbNullString
      .ColumnID = 0
      .ExpressionID = val(txtCalculated.Tag)

    End If

  End With

End Function


Private Function ValidEmailAddress(strEmailAddress) As Boolean
  'Must only contain one @ sign and must be something in front of @ and after @
  
  Dim varTemp As Variant
  
  ValidEmailAddress = False
  
  varTemp = Split(Trim(strEmailAddress), "@")
  
  If UBound(varTemp) = 1 Then
    If Left(strEmailAddress, "1") <> "@" And Right(strEmailAddress, "1") <> "@" Then
      'Check for full stop after @ sign
      ValidEmailAddress = (InStr(varTemp(1), ".") > 0)
    End If
  End If

End Function


Private Function IsUsedInEmailGroup(lngEmailID As Long) As Boolean

  Dim rsEmailGroups As New ADODB.Recordset
  Dim strSQL As String
  Dim strAbsenceType As String
  Dim iListIndex As Integer

  On Error GoTo LocalErr

  IsUsedInEmailGroup = False
  
  strSQL = "SELECT COUNT(EmailGroupID) FROM ASRSysEmailGroupItems WHERE EmailDefID = " & CStr(lngEmailID)
  rsEmailGroups.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  If Not rsEmailGroups.EOF Then
    IsUsedInEmailGroup = (rsEmailGroups.Fields(0).value > 0)
  End If

  rsEmailGroups.Close
  Set rsEmailGroups = Nothing

Exit Function

LocalErr:
  MsgBox Err.Description, vbCritical, Me.Caption

End Function



