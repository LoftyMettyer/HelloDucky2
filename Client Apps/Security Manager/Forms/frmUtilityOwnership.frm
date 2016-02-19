VERSION 5.00
Begin VB.Form frmUtilityOwnership 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utility Ownership Transfer"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8044
   Icon            =   "frmUtilityOwnership.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5580
      TabIndex        =   3
      Top             =   1050
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4245
      TabIndex        =   2
      Top             =   1050
      Width           =   1200
   End
   Begin VB.Frame fraUtilityOwnership 
      Caption         =   "Transfer ownership of all reports, tools and utilities :"
      Height          =   825
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   6690
      Begin VB.ComboBox cboAccess 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7110
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1575
         Width           =   1800
      End
      Begin VB.ComboBox cboTo 
         Height          =   315
         Left            =   3900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2625
      End
      Begin VB.ComboBox cboFrom 
         Height          =   315
         Left            =   765
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   2625
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Transfer ownership of individual tools, reports and utilities :"
         Enabled         =   0   'False
         Height          =   285
         Left            =   195
         TabIndex        =   5
         Top             =   1215
         Width           =   4485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change Ownership..."
         Enabled         =   0   'False
         Height          =   400
         Left            =   7110
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "Toggle Selection"
         Height          =   400
         Left            =   1455
         TabIndex        =   11
         Top             =   5370
         Width           =   1650
      End
      Begin VB.ComboBox cboOwner 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1575
         Width           =   1800
      End
      Begin VB.ListBox lstUtilities 
         Enabled         =   0   'False
         Height          =   3210
         Left            =   1455
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   2025
         Width           =   7455
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1575
         Width           =   1800
      End
      Begin VB.OptionButton optAll 
         Caption         =   "Transfer ownership of all tools, reports and utilities :"
         Height          =   285
         Left            =   200
         TabIndex        =   4
         Top             =   1500
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   195
         Left            =   3525
         TabIndex        =   17
         Top             =   375
         Width           =   285
      End
      Begin VB.Label lblAccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access :"
         Height          =   195
         Left            =   6285
         TabIndex        =   18
         Top             =   1650
         Width           =   615
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   390
         Width           =   465
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Left            =   3465
         TabIndex        =   15
         Top             =   1650
         Width           =   555
      End
      Begin VB.Label lblDefinitions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Definitions :"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   2100
         Width           =   825
      End
      Begin VB.Label lblUtilityType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   495
         TabIndex        =   13
         Top             =   1650
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmUtilityOwnership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnLoading As Boolean
'Private madoCon As adodb.Connection

Public Function Initialise() As Boolean

  On Error GoTo LoadCombos_ERROR

  Initialise = True
  mblnLoading = True

  With gobjProgress
    '.AviFile = ""
    .AVI = dbXferOwnership
    .NumberOfBars = 1
    .Bar1MaxValue = (gObjGroups.Count * 2)
    .Caption = "Loading users, please wait..."
    .MainCaption = "Utility Ownership"
    .Cancel = False
    .Time = False
    .HidePercentages = False
    .OpenProgress
  End With
  
  PopulateUsersFromCombo
  PopulateUsersToCombo cboFrom.Text
  
  gobjProgress.CloseProgress

  'TM20011015 Fault 2187
  'Check if there are users to transfer utilities to.
  If cboTo.ListCount > 0 And cboFrom.ListCount > 2 Then
    cboFrom.ListIndex = 0
    cboTo.ListIndex = 0
  Else
    MsgBox "No users on the system to transfer the utility ownership to.", vbExclamation + vbOKOnly, App.Title
    Initialise = False
  End If
  
TidyUp:

  mblnLoading = False
  Exit Function

LoadCombos_ERROR:

  Initialise = False
  MsgBox "Error whilst initialising the Utility Ownership form." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUp

End Function

'Private Sub cboFrom_Click()
'  'TM20010910 Fault 2807
'  'The Click event is called when then the .ListIndex property is changed in the PopulateUsersCombo
'  'routine when the form is initialised. There for we only need to call the PopulateUsersCombo
'  'routine when the form has fully loaded.
'  If Not mblnLoading Then
'    PopulateUsersCombo cboTo, False, cboFrom.Text
'    cboFrom.Tag = cboFrom.Text
'  End If
'End Sub

Private Sub cboTo_Click()
  cboTo.Tag = cboTo.Text
End Sub

Private Sub cmdCancel_Click()

  Unload Me

End Sub

Private Sub cmdOK_Click()

  'MH20060613 Fault 10805
  'Me.Visible = False

  If ValidateSelection Then
    DoTransfer
  End If

  'MH20060613 Fault 10805
  'Me.Visible = True
  
End Sub

Private Function ValidateSelection() As Boolean

  On Error GoTo ValidateERROR

  If cboFrom.Text = "<All>" Then
    If MsgBox("You have selected to change the ownership of ALL definitions to " & cboTo.Text & vbCrLf & _
                    "This action cannot be undone. Are you sure you wish to continue ?", vbYesNo + vbQuestion, App.Title) = vbNo Then
      ValidateSelection = False
      Me.Visible = True
      Exit Function
    End If
  End If

  If cboFrom.Text = cboTo.Text Then
    MsgBox "You cannot transfer ownership from/to the same person.", vbExclamation + vbOKOnly, App.Title
    ValidateSelection = False
    Me.Visible = True
    Exit Function
  End If

  ValidateSelection = True
  Exit Function

ValidateERROR:

  MsgBox "Error whilst validating selection." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  ValidateSelection = False

End Function

Private Sub Progress(strCaption As String)

  With gobjProgress
    .Bar1Caption = strCaption
    .UpdateProgress False
  End With
  
End Sub

Private Function DoTransfer() As Boolean

  On Error GoTo DoTransferERROR

  Dim strCommand As String
  Dim strFrom As String
  Dim strTo As String
  Dim blnAll As Boolean
  
  With gobjProgress
    '.AviFile = "" 'App.Path & "\videos\table.Avi"
    .AVI = dbXferOwnership
    .Caption = "Ownership Transfer..."
    .MainCaption = "Utility Ownership"
    .NumberOfBars = 1
    .Bar1MaxValue = 15
    .Time = False
    .Cancel = True
    .OpenProgress
  End With
  
  ' Set variables
  Progress "Setting Variables..."
  strFrom = Replace(cboFrom.Text, "'", "''")
  strTo = Replace(cboTo.Text, "'", "''")
  blnAll = (strFrom = "<All>")
  DoEvents
  
  ' Initialise the gADOCon connection
  Progress "Initialising Connection..."
  
'  If gbUseWindowsAuthentication Then
'    Connect ("Driver={SQL Server};Server=" & gsSQLServerName & ";UID=" & gsUserName & ";PWD=" & gsPassword & ";Database=" & gsDatabaseName & ";Integrated Security=SSPI;")
'  Else
'    Connect ("Driver={SQL Server};Server=" & gsSQLServerName & ";UID=" & gsUserName & ";PWD=" & gsPassword & ";Database=" & gsDatabaseName & ";")
'  End If
  
  DoEvents

  ' Batch Jobs
  Progress "Transferring Batch Jobs and Report Packs..."
  strCommand = "UPDATE ASRSysBatchJobName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Calendar Reports
  Progress "Transferring Calendar Reports..."
  strCommand = "UPDATE ASRSysCalendarReports SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Cross Tabs
  Progress "Transferring CrossTabs..."
  strCommand = "UPDATE ASRSysCrossTab SET Username = '" & strTo & "'" & " WHERE CrossTabType <> " & ctt9GridBox
  If Not blnAll Then strCommand = strCommand & " AND Username = '" & strFrom & "'"

  gADOCon.Execute strCommand
  DoEvents
  
  ' 9-Box Grid Reports
  If IsModuleEnabled(modNineBoxGrid) Then
    Progress "Transferring 9-Box Grid Reports..."
    strCommand = "UPDATE ASRSysCrossTab SET Username = '" & strTo & "'" & " WHERE CrossTabType = " & ctt9GridBox
    If Not blnAll Then strCommand = strCommand & " AND Username = '" & strFrom & "'"
  
    gADOCon.Execute strCommand
    DoEvents
  End If
  
  ' Talent Reports
  Progress "Transferring Talent Reports..."
  strCommand = "UPDATE ASRSysTalentReports SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"

  gADOCon.Execute strCommand
  DoEvents
    
  ' Custom Reports
  Progress "Transferring Custom Reports..."
  strCommand = "UPDATE ASRSysCustomReportsName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Data Transfer
  Progress "Transferring Data Transfers..."
  strCommand = "UPDATE ASRSysDataTransferName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Document Types
  Progress "Transferring Document Types..."
  strCommand = "UPDATE ASRSysDocumentManagementTypes SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Diary Events
  Progress "Transferring Diary Events..."
  strCommand = "UPDATE ASRSysDiaryEvents SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Exports
  Progress "Transferring Exports..."
  strCommand = "UPDATE ASRSysExportName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Expressions
  Progress "Transferring Expressions..."
  strCommand = "UPDATE ASRSysExpressions SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Global Functions
  Progress "Transferring Global Functions..."
  strCommand = "UPDATE ASRSysGlobalFunctions SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Import
  Progress "Transferring Imports..."
  strCommand = "UPDATE ASRSysImportName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Mail Merge
  Progress "Transferring Mail Merges..."
  strCommand = "UPDATE ASRSysMailMergeName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Picklists
  Progress "Transferring Picklists..."
  strCommand = "UPDATE ASRSysPicklistName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Record Profiles
  Progress "Transferring Record Profiles..."
  strCommand = "UPDATE ASRSysRecordProfileName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Match Reports/Career Progression/Succession Planning
  Progress "Transferring Match Reports..."
  strCommand = "UPDATE ASRSysMatchReportName SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Envelope & Label Templates
  Progress "Transferring Envelope & Label Templates..."
  strCommand = "UPDATE ASRSysLabelTypes SET Username = '" & strTo & "'"
  If Not blnAll Then strCommand = strCommand & " WHERE Username = '" & strFrom & "'"
  
  gADOCon.Execute strCommand
  DoEvents
  
  ' Close progress bar
  gobjProgress.CloseProgress

  ' Inform user
  MsgBox "Ownership transferred successfully.", vbInformation + vbOKOnly, App.Title
  
  DoTransfer = True
  Exit Function

DoTransferERROR:

  gobjProgress.Visible = False
  gobjProgress.CloseProgress
  MsgBox "Error whilst performing ownership transfer." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  DoTransfer = False

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub PopulateUsersFromCombo()

  Dim rsUsers As ADODB.Recordset

  ' Load the Users combo
  cboFrom.Clear
  cboFrom.AddItem "<All>"
  
  Set rsUsers = New ADODB.Recordset
  rsUsers.Open "spASRGetOwnersForAllUtilities", gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
  
  If Not (rsUsers.BOF And rsUsers.EOF) Then
    Do While Not rsUsers.EOF
      cboFrom.AddItem rsUsers.Fields("UserName").Value
      rsUsers.MoveNext
    Loop
  End If
  
End Sub

Private Sub PopulateUsersToCombo(ByRef strExclude As String)

  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser

  ' Load the Users combo
  cboTo.Clear

  For Each objGroup In gObjGroups
    ' If the collections dont already exist, initialise them
    If Not gObjGroups(objGroup.Name).Users_Initialised Then
      InitialiseUsersCollection gObjGroups(objGroup.Name)
    End If

    gobjProgress.Bar1Caption = "Processing group '" & objGroup.Name & "'"

    ' Now add the users
    For Each objUser In gObjGroups(objGroup.Name).Users
    ' NPG20090206 Fault 11931
    ' If (Not objUser.DeleteUser) And (objUser.MovedUserTo = "") And Not objUser.LoginType = iUSERTYPE_TRUSTEDGROUP then
      If (Not objUser.DeleteUser) _
        And (objUser.MovedUserTo = "") And Not objUser.LoginType = iUSERTYPE_TRUSTEDGROUP _
        And Not objUser.LoginType = iUSERTYPE_ORPHANUSER _
        And Not objUser.LoginType = iUSERTYPE_ORPHANGROUP Then

        If objUser.UserName <> strExclude Then
          cboTo.AddItem objUser.UserName
        End If

      End If
    Next objUser

    If gobjProgress.Visible Then
      gobjProgress.UpdateProgress (False)
    End If

  Next objGroup

  SetComboText cboTo, cboTo.Tag
  If cboTo.ListCount > 0 And cboTo.ListIndex = -1 Then
    cboTo.ListIndex = 0
  End If

  Set objGroup = Nothing
  Set objUser = Nothing

End Sub

