VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmViewCurrentUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Current Users"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5039
   Icon            =   "frmViewCurrentUsers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkKillWebUsers 
      Caption         =   "&Forcibly disconnect all web users"
      Height          =   330
      Left            =   120
      TabIndex        =   18
      Top             =   4590
      Width           =   3315
   End
   Begin VB.CommandButton cmdAutoRetrySave 
      Caption         =   "&Auto Retry Save"
      Height          =   400
      Left            =   7300
      TabIndex        =   17
      Top             =   2895
      Width           =   1710
   End
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "Send &Message"
      Enabled         =   0   'False
      Height          =   400
      Left            =   7300
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   2400
      Width           =   1710
   End
   Begin VB.CheckBox chkASRDevBypass 
      Caption         =   "ASR Dev Bypass"
      Height          =   375
      Left            =   7300
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7300
      TabIndex        =   1
      Top             =   700
      Width           =   1710
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   400
      Left            =   7300
      TabIndex        =   0
      Top             =   200
      Width           =   1710
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Un&lock Database"
      Enabled         =   0   'False
      Height          =   400
      Left            =   7300
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   1900
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Locking :"
      Height          =   1215
      Left            =   105
      TabIndex        =   6
      Top             =   100
      Width           =   7000
      Begin VB.TextBox txtLockDateTime 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4850
         TabIndex        =   14
         Top             =   700
         Width           =   2000
      End
      Begin VB.TextBox txtLockType 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1335
         TabIndex        =   12
         Top             =   700
         Width           =   2000
      End
      Begin VB.TextBox txtLockMachine 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4850
         TabIndex        =   10
         Top             =   300
         Width           =   2000
      End
      Begin VB.TextBox txtLockUser 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1335
         TabIndex        =   8
         Top             =   300
         Width           =   2000
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Date / Time :"
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   765
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lock Type :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Machine :"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   400
      Left            =   7300
      TabIndex        =   2
      Top             =   1400
      Width           =   1710
   End
   Begin VB.Frame fraUsers 
      Height          =   3165
      Left            =   105
      TabIndex        =   15
      Top             =   1350
      Width           =   7000
      Begin SSDataWidgets_B.SSDBGrid grdUsers 
         Height          =   2670
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   6600
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   3
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         BalloonHelp     =   0   'False
         RowNavigation   =   1
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   3572
         Columns(0).Caption=   "User"
         Columns(0).Name =   "User"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2990
         Columns(1).Caption=   "Machine"
         Columns(1).Name =   "Machine"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4604
         Columns(2).Caption=   "Module"
         Columns(2).Name =   "Module"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   11642
         _ExtentY        =   4710
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
   End
End
Attribute VB_Name = "frmViewCurrentUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSaving As Boolean
Private mstrUsersToLogOut As String
Private mblnCancelled As Boolean
Private mintLockType As LockTypes
Private mblnDBLocked As Boolean
Private mbKillWebUsers As Boolean
Private mintWebUserCount As Integer

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "WinMM.dll" () As Long

Private Sub CheckEnableSave()
  cmdSave.Enabled = (grdUsers.Rows = 0 Or _
        (chkASRDevBypass.Visible = True And chkASRDevBypass.value = vbChecked) Or _
        (chkKillWebUsers.value = vbChecked And mintWebUserCount = grdUsers.Rows))
End Sub

Private Sub chkASRDevBypass_Click()
  CheckEnableSave
End Sub

Private Sub chkKillWebUsers_Click()
  CheckEnableSave
End Sub

Private Sub cmdAutoRetrySave_Click()

'Dim objProgress As New HRProProgress.clsHRProProgress
'Dim objProgress As New clsProgress

  If mblnSaving Then
   
    Me.Enabled = False
    
    'Set objProgress = New HRProProgress.clsHRProProgress
    'Set objProgress = New clsProgress
    gobjProgress.NumberOfBars = 1
    gobjProgress.ResetBar1
    gobjProgress.Bar1Caption = "Automatic Retry Save..."
    gobjProgress.Cancel = True
    'gobjProgress.Caption = gobjProgress.Caption
    'gobjProgress.AviFile = gobjProgress.AviFile
    gobjProgress.MainCaption = "Saving Changes"
    gobjProgress.OpenProgress
  
    Do While Not (gobjProgress.Cancelled Or OkayToSave(mstrUsersToLogOut))
      Wait (5000)
    Loop

    If Not gobjProgress.Cancelled Then
      mblnCancelled = False
      Me.Hide
    Else
      gobjProgress.CloseProgress
    End If

    ' AE20080311 Fault #13000
    'objProgress = Nothing
    'Set objProgress = Nothing
  
    Me.Enabled = True
    
  End If

End Sub

Private Sub cmdLock_Click()
  
  'AE20071213 #S000738
  Dim frmMsg As frmLockMessage
    
  If mintLockType <> lckManual Then
    Set frmMsg = New frmLockMessage
  
    frmMsg.Show vbModal
    
    If (Not frmMsg.Cancelled) Then
      Screen.MousePointer = vbHourglass
      
      LockDatabase lckManual

    End If
  
    Set frmMsg = Nothing
  Else
    Screen.MousePointer = vbHourglass
    
    If Application.AccessMode = accFull Or _
      Application.AccessMode = accSupportMode Then
            UnlockDatabase lckManual

    ElseIf gbCurrentUserIsSysSecMgr Then
      If MsgBox("Are you sure that you would like to clear the manual lock?", vbQuestion + vbYesNo) = vbYes Then
        UnlockDatabase lckManual, True
      End If
    
    End If

  End If

  LockCheck
  Screen.MousePointer = vbDefault

End Sub


Private Sub cmdSendMessage_Click()
  ' Broadcast a message to all users.
  Dim frmMsg As frmSendMessage
  Dim cmdSendMessage As ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  Set frmMsg = New frmSendMessage

  frmMsg.Show vbModal

  If (Not frmMsg.Cancelled) And _
    (Len(frmMsg.Message) > 0) Then
    ' Send the message to all current users.
  
    Set cmdSendMessage = New ADODB.Command
    With cmdSendMessage
      .CommandText = "dbo.sp_ASRSendMessage"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
  
      Set pmADO = .CreateParameter("Message", adVarChar, adParamInput, VARCHAR_MAX_Size)
      .Parameters.Append pmADO
      pmADO.value = frmMsg.Message
  
      Set pmADO = .CreateParameter("SPIDs", adVarChar, adParamInput, VARCHAR_MAX_Size)
      .Parameters.Append pmADO
      pmADO.value = vbNullString
  
       .Execute

    End With
    Set cmdSendMessage = Nothing
    
  End If

  Set frmMsg = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    UnLoad Me
End Select
End Sub

Private Sub Form_Load()
  Const GRIDROWHEIGHT = 239
  
  grdUsers.RowHeight = GRIDROWHEIGHT

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  CheckEnableSave

End Sub

Private Sub Form_Activate()
Dim bOK As Boolean

  mblnCancelled = True
  bOK = GetUsers
  
  If bOK Then
    LockCheck
  End If
End Sub

Private Function GetUsers() As Boolean
  
  Dim intTempPointer As Integer
  Dim fOK As Boolean
  
  On Local Error GoTo LocalErr
  
  fOK = True
  
  intTempPointer = Screen.MousePointer
  Screen.MousePointer = vbHourglass

  fOK = CurrentUsersPopulate(grdUsers, "", mintWebUserCount)
  Form_Resize
  cmdSendMessage.Enabled = True

  Screen.MousePointer = intTempPointer

TidyAndExit:
  GetUsers = fOK

Exit Function

LocalErr:
  fOK = False
  Resume TidyAndExit

End Function



Private Sub LockCheck()

  Dim rsTemp As New ADODB.Recordset
    
  'Check that no other user has the database locked...
  rsTemp.Open "sp_ASRLockCheck", gADOCon, adOpenForwardOnly, adLockReadOnly
    
  txtLockUser.Text = vbNullString
  txtLockType.Text = vbNullString
  txtLockMachine.Text = vbNullString
  txtLockDateTime.Text = vbNullString
  mintLockType = lckNone
  
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    If Trim(rsTemp!userName) <> vbNullString Then
      txtLockUser.Text = rsTemp!userName
      txtLockType.Text = rsTemp!Description
      txtLockMachine.Text = rsTemp!HostName
      txtLockDateTime.Text = rsTemp!Lock_Time
      mintLockType = rsTemp!Priority
    End If
  End If

  rsTemp.Close
  Set rsTemp = Nothing
    
  cmdLock.Enabled = True
  If Application.AccessMode = accFull Then
    If mintLockType = lckManual Then
      cmdLock.Caption = "Un&lock Database"
      Me.HelpContextID = 5067
    Else
      cmdLock.Caption = "&Lock Database"
      Me.HelpContextID = 5039
    End If
  ElseIf LCase(gsUserName) = "sa" And mintLockType = lckManual Then
    cmdLock.Caption = "Un&lock Database"
    Me.HelpContextID = 5067
  Else
    cmdLock.Enabled = False
    Me.HelpContextID = 5039
  End If

End Sub


Public Property Get Saving() As Boolean
  Saving = mblnSaving
End Property

Public Property Let Saving(ByVal blnNewValue As Boolean)

  mblnSaving = blnNewValue

  cmdRefresh.Visible = True
  cmdLock.Visible = Not mblnSaving
  cmdSave.Visible = mblnSaving
  cmdCancel.Visible = True
  cmdAutoRetrySave.Visible = mblnSaving
  chkKillWebUsers.Visible = mblnSaving

  chkASRDevBypass.Visible = (ASRDEVELOPMENT And mblnSaving)
  chkASRDevBypass.value = vbChecked

  If Not mblnSaving Then
    cmdCancel.Top = 200
    cmdCancel.Caption = "&OK"
    cmdRefresh.Top = 800
    cmdLock.Top = 1300
    cmdSendMessage.Top = 1800
  End If

End Property


Public Function OkayToSave(Optional strUsersToLogOut As String) As Boolean
  
  Dim fOK As Boolean
  
  'If you can't lock the database or users are still logged in
  'then show the status form (details of lock and users)
  mstrUsersToLogOut = strUsersToLogOut
  mblnDBLocked = False
  
  fOK = GetUsers
  
  ' AE20080325 Fault #12903
'  If fOK Then
'    fOK = (fOK And grdUsers.Rows = 0 And LockDatabase(lckSaving))
'  End If
  
  If fOK Then
    fOK = (fOK And grdUsers.Rows = 0)
  'End If
  ' AE20080527 Fault #13184
  'If fOK Then
    mblnDBLocked = (Not LockDatabase(lckSaveRequest))
    fOK = (fOK And (Not mblnDBLocked))
  End If
  
  OkayToSave = fOK

End Function

'Public Function OkayToSave() As Boolean
'  GetUsers
'  OkayToSave = (grdUsers.Rows = 0 And LockDatabase("Lock Saving"))
'End Function


Private Sub cmdCancel_Click()
  mblnCancelled = True
  Me.Hide
End Sub

Private Sub cmdSave_Click()

  If mblnSaving And Not (chkASRDevBypass.value = vbChecked Or chkKillWebUsers.value = vbChecked) Then
    If Not OkayToSave(mstrUsersToLogOut) Then
    'If Not OkayToSave Then
      Exit Sub
    End If
  End If

  mblnCancelled = False
  Me.Hide

End Sub

Private Sub cmdRefresh_Click()

  Dim bOK As Boolean

  bOK = GetUsers
  
  If bOK Then
    LockCheck
  End If
  
  CheckEnableSave
End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Get SendMessageVisible() As Boolean
  SendMessageVisible = cmdSendMessage.Visible
End Property

Public Property Let SendMessageVisible(ByVal blnNewValue As Boolean)
  cmdSendMessage.Visible = blnNewValue
End Property

Public Property Get Locked() As Variant
  Locked = mblnDBLocked
End Property

Public Sub Wait(ByVal inMilliseconds As Long)
    Dim SleepTime As Long, TimeNow As Long
    Dim SleepTo As Long, SleepEnd As Long

    Const MaxSleep As Long = 100

    TimeNow = timeGetTime()
    SleepTime = inMilliseconds \ 10
    If (SleepTime > MaxSleep) Then SleepTime = MaxSleep
    SleepTo = TimeNow + inMilliseconds

    Do
        DoEvents
        TimeNow = timeGetTime()
        SleepEnd = SleepTo - TimeNow
        If (SleepEnd <= SleepTime) Then Exit Do
        Call Sleep(SleepTime)
    Loop

    If (SleepEnd > 0) Then Call Sleep(SleepEnd)
End Sub

