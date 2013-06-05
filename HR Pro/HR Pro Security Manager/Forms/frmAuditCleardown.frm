VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmAuditCleardown 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule Audit Log Purging"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1005
   Icon            =   "frmAuditCleardown.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4935
      TabIndex        =   13
      Top             =   2070
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   3690
      TabIndex        =   12
      Top             =   2070
      Width           =   1200
   End
   Begin TabDlg.SSTab tabClear 
      Height          =   1830
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   3228
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Data Records"
      TabPicture(0)   =   "frmAuditCleardown.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data &Permissions"
      TabPicture(1)   =   "frmAuditCleardown.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTableColumn"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&User Maintenance"
      TabPicture(2)   =   "frmAuditCleardown.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraUser"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "User &Access"
      TabPicture(3)   =   "frmAuditCleardown.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraAccess"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraAccess 
         Caption         =   "Purge Criteria :"
         Enabled         =   0   'False
         Height          =   1245
         Left            =   -74865
         TabIndex        =   18
         Top             =   405
         Width           =   5820
         Begin VB.ComboBox cboPeriod 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmAuditCleardown.frx":007C
            Left            =   3585
            List            =   "frmAuditCleardown.frx":008C
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   720
            Width           =   1425
         End
         Begin VB.OptionButton optPurge 
            Caption         =   "Purge entries older than :"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   20
            Top             =   750
            Width           =   2550
         End
         Begin VB.OptionButton optNoPurge 
            Caption         =   "Do not automatically purge the User Access records"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   19
            Top             =   345
            Value           =   -1  'True
            Width           =   4860
         End
         Begin COASpinner.COA_Spinner spnDays 
            Height          =   300
            Index           =   3
            Left            =   2790
            TabIndex        =   21
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
      End
      Begin VB.Frame fraData 
         Caption         =   "Purge Criteria :"
         Height          =   1245
         Left            =   135
         TabIndex        =   17
         Top             =   405
         Width           =   5820
         Begin VB.OptionButton optNoPurge 
            Caption         =   "Do not automatically purge the Data records"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   0
            Top             =   345
            Value           =   -1  'True
            Width           =   4245
         End
         Begin VB.OptionButton optPurge 
            Caption         =   "Purge entries older than :"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   750
            Width           =   2595
         End
         Begin VB.ComboBox cboPeriod 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmAuditCleardown.frx":00B4
            Left            =   3585
            List            =   "frmAuditCleardown.frx":00C4
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   1425
         End
         Begin COASpinner.COA_Spinner spnDays 
            Height          =   300
            Index           =   0
            Left            =   2790
            TabIndex        =   2
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
      End
      Begin VB.Frame fraTableColumn 
         Caption         =   "Purge Criteria :"
         Enabled         =   0   'False
         Height          =   1245
         Left            =   -74865
         TabIndex        =   16
         Top             =   405
         Width           =   5820
         Begin VB.OptionButton optNoPurge 
            Caption         =   "Do not automatically purge the Data Permissions records"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   345
            Value           =   -1  'True
            Width           =   5280
         End
         Begin VB.OptionButton optPurge 
            Caption         =   "Purge entries older than :"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   750
            Width           =   2505
         End
         Begin VB.ComboBox cboPeriod 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmAuditCleardown.frx":00EC
            Left            =   3585
            List            =   "frmAuditCleardown.frx":00FC
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   1425
         End
         Begin COASpinner.COA_Spinner spnDays 
            Height          =   300
            Index           =   1
            Left            =   2790
            TabIndex        =   6
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
      End
      Begin VB.Frame fraUser 
         Caption         =   "Purge Criteria :"
         Enabled         =   0   'False
         Height          =   1245
         Left            =   -74865
         TabIndex        =   15
         Top             =   405
         Width           =   5820
         Begin VB.OptionButton optNoPurge 
            Caption         =   "Do not automatically purge the User Maintenance records"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   8
            Top             =   345
            Value           =   -1  'True
            Width           =   5265
         End
         Begin VB.OptionButton optPurge 
            Caption         =   "Purge entries older than :"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   9
            Top             =   750
            Width           =   2505
         End
         Begin VB.ComboBox cboPeriod 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmAuditCleardown.frx":0124
            Left            =   3585
            List            =   "frmAuditCleardown.frx":0134
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   1425
         End
         Begin COASpinner.COA_Spinner spnDays 
            Height          =   300
            Index           =   2
            Left            =   2790
            TabIndex        =   10
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            MaximumValue    =   999
            Text            =   "0"
         End
      End
   End
   Begin VB.Label lblWarning 
      Caption         =   "Warning: Purging Data can affect the results of Audit && Export Functions."
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   105
      TabIndex        =   23
      Top             =   2055
      Visible         =   0   'False
      Width           =   3510
   End
End
Attribute VB_Name = "frmAuditCleardown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miAuditType As Integer
Private mblnReadOnly As Boolean

Public Property Let AuditType(ByVal iNewValue As Integer)
  miAuditType = iNewValue
End Property


Public Sub Initialise()

  ' Initialise the form controls.
  Dim rsData As Recordset
  Dim iIndex As Integer

  'MH20010823 Fault 1917 Only allow "sa" full access to audit log
  'mblnReadOnly = (Application.AccessMode <> accFull)
  mblnReadOnly = (Application.AccessMode <> accFull Or Not gbUserCanManageLogins) ' Or LCase(gsUserName) <> "sa")
  
  If mblnReadOnly Then
    ControlsDisableAll Me
  End If


  Set rsData = GetCleardownData
  
  If rsData.BOF And rsData.EOF Then
    Set rsData = Nothing
    Me.tabClear.Tab = miAuditType - 1
    Exit Sub
  End If
  
  With rsData
    Do While Not .EOF
      Select Case !Type
      Case "Data": iIndex = 0
      Case "Permissions": iIndex = 1
      Case "Users": iIndex = 2
      Case "Access": iIndex = 3
      End Select

      Me.spnDays(iIndex).Value = !frequency
      optPurge(iIndex).Value = True

      Select Case !period
      Case "dd": SetComboText cboPeriod(iIndex), "Day(s)"
      Case "wk": SetComboText cboPeriod(iIndex), "Week(s)"
      Case "mm": SetComboText cboPeriod(iIndex), "Month(s)"
      Case "yy": SetComboText cboPeriod(iIndex), "Year(s)"
      Case Else: cboPeriod(iIndex).ListIndex = -1
      End Select

      rsData.MoveNext
    Loop
  End With

  Set rsData = Nothing
  
  ' Set the current tab to be what grid the user is looking at
  Me.tabClear.Tab = miAuditType - 1

  'JDM - 19/02/02 - Fault 3491 - Activate the warning if the user is purging the data records
  If optPurge(0).Value = True Then
    lblWarning.Visible = True
  End If

End Sub

Private Sub cmdCancel_Click()

  Unload Me
  
End Sub

Private Sub cmdOK_Click()

  If Not Validate Then
    Exit Sub
  End If
  
  SaveCleardown
  
  Unload Me
  
End Sub

Private Sub Form_Load()
  
  ' Select the first tab page as default.
  tabClear.Tab = 0

End Sub

Private Function Validate() As Boolean

  Dim pintLoop As Integer
  
  For pintLoop = 0 To 3
  
    If optPurge(pintLoop).Value Then
      If cboPeriod(pintLoop).Text = "" Then
        MsgBox "You must select a period to purge audit log entries.", vbExclamation + vbOKOnly, "Audit Log"
        tabClear.Tab = pintLoop
        Validate = False
        Exit Function
      End If
    
      ' JDM - 19/02/02 - Fault 3531 - Purge period was just too big...
      If cboPeriod(pintLoop).Text = "Year(s)" And spnDays(pintLoop).Value > 200 Then
        MsgBox "You cannot select a purge period of greater than 200 years.", vbExclamation + vbOKOnly, "Audit Log"
        tabClear.Tab = pintLoop
        Validate = False
        Exit Function
      End If
    
      
    End If
    
  Next pintLoop
 
  Validate = True

End Function

Private Sub SaveCleardown()

  ' Save the Cleardown definition.
  Dim pintLoop As Integer
  Dim pstrType As String
  Dim pstrPeriod As String
  Dim pstrTriggerSQL As String
  Dim pstrTable As String
  
  On Error GoTo ErrorTrap
  
  ' RH 13/10/00 - BUG 1140 - Dont show msg if no auto purge criteria selected
  If optNoPurge(0).Value = False Or _
     optNoPurge(1).Value = False Or _
     optNoPurge(2).Value = False Or _
     optNoPurge(3).Value = False Then
  
    MsgBox "The audit logs will now be purged using the criteria you have just set", vbInformation + vbOKOnly, "Audit Log"
  
  End If
  
  Screen.MousePointer = vbHourglass
  
  DeleteCleardowns

'  ' Remove the trigger (if it exists already)
'  gADOCon.Execute "DROP TRIGGER INS_AsrSysPurgeAuditTrail"
'  gADOCon.Execute "DROP TRIGGER INS_AsrSysPurgeAuditGroup"
'  gADOCon.Execute "DROP TRIGGER INS_AsrSysPurgeAuditPermissions"
'  gADOCon.Execute "DROP TRIGGER INS_AsrSysPurgeAuditAccess"
'
  For pintLoop = 0 To 3

    If optPurge(pintLoop).Value = True Then

      Select Case pintLoop
        Case 0: pstrType = "Data": pstrTable = "AuditTrail"
        Case 1: pstrType = "Permissions": pstrTable = "AuditPermissions"
        Case 2: pstrType = "Users": pstrTable = "AuditGroup"
        Case 3: pstrType = "Access": pstrTable = "AuditAccess"
      End Select

      Select Case cboPeriod(pintLoop).Text
        Case "Day(s)": pstrPeriod = "dd"
        Case "Week(s)": pstrPeriod = "wk"
        Case "Month(s)": pstrPeriod = "mm"
        Case "Year(s)": pstrPeriod = "yy"
      End Select

      InsertCleardown pstrType, spnDays(pintLoop).Value, pstrPeriod

'      ' Now create the trigger for insert on the relevant table
'
'      pstrTriggerSQL = pstrTriggerSQL & "CREATE TRIGGER INS_AsrSysPurge" & pstrTable & " "
'      pstrTriggerSQL = pstrTriggerSQL & "ON AsrSys" & pstrTable & " "
'      pstrTriggerSQL = pstrTriggerSQL & "FOR INSERT AS "
'
'      pstrTriggerSQL = pstrTriggerSQL & "DECLARE @intFrequency int, "
'      pstrTriggerSQL = pstrTriggerSQL & "@strPeriod char(2) "
'
'      pstrTriggerSQL = pstrTriggerSQL & "SELECT @intFrequency = Frequency "
'      pstrTriggerSQL = pstrTriggerSQL & "FROM AsrSysAuditCleardown "
'      pstrTriggerSQL = pstrTriggerSQL & "WHERE Type = '" & pstrType & "' "
'
'      pstrTriggerSQL = pstrTriggerSQL & "SELECT @strPeriod = Period "
'      pstrTriggerSQL = pstrTriggerSQL & "FROM AsrSysAuditCleardown "
'      pstrTriggerSQL = pstrTriggerSQL & "WHERE Type = '" & pstrType & "' "
'
'      pstrTriggerSQL = pstrTriggerSQL & "IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL) "
'
'      pstrTriggerSQL = pstrTriggerSQL & "BEGIN "
'
'      pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'dd' BEGIN DELETE FROM AsrSys" & pstrTable & " WHERE [DateTimeStamp] < DATEADD(dd,-@intFrequency,getdate()) END "
'      pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'wk' BEGIN DELETE FROM AsrSys" & pstrTable & " WHERE [DateTimeStamp] < DATEADD(wk,-@intFrequency,getdate()) END "
'      pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'mm' BEGIN DELETE FROM AsrSys" & pstrTable & " WHERE [DateTimeStamp] < DATEADD(mm,-@intFrequency,getdate()) END "
'      pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'yy' BEGIN DELETE FROM AsrSys" & pstrTable & " WHERE [DateTimeStamp] < DATEADD(yy,-@intFrequency,getdate()) END "
'
'      pstrTriggerSQL = pstrTriggerSQL & "END"
'
'      gADOCon.Execute (pstrTriggerSQL)
'
'      pstrTriggerSQL = ""
'
    End If
  Next pintLoop
  
  gADOCon.Execute "sp_AsrAuditLogPurge"
    
  Screen.MousePointer = vbDefault
    
  Exit Sub
  
ErrorTrap:
  
   Select Case Err.Number
  
    Case -2147217865 ' trigger does not exist
      Resume Next
    
    Case Else
      MsgBox "Error saving purge information" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Audit Log"
      Exit Sub
  End Select
  
  Screen.MousePointer = vbDefault
  
End Sub




Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optNoPurge_Click(Index As Integer)

  Me.cboPeriod(Index).ListIndex = -1
  Me.cboPeriod(Index).BackColor = &H8000000F
  Me.cboPeriod(Index).Enabled = False
  
  Me.spnDays(Index).Value = 0
  Me.spnDays(Index).BackColor = &H8000000F
  Me.spnDays(Index).Enabled = False
  
  'JDM - 19/02/02 - Fault 3491 - Activate the warning if the user is purging the data records
  lblWarning.Visible = (optPurge(0).Value) And (tabClear.Tab = 0)
  
End Sub

Private Sub optPurge_Click(Index As Integer)
  
  If mblnReadOnly Then
    Exit Sub
  End If

  Me.cboPeriod(Index).ListIndex = 0
  Me.cboPeriod(Index).Enabled = True
  Me.cboPeriod(Index).BackColor = &H80000005
  
  Me.spnDays(Index).Enabled = True
  Me.spnDays(Index).BackColor = &H80000005

  'JDM - 19/02/02 - Fault 3491 - Activate the warning if the user is purging the data records
  lblWarning.Visible = optPurge(0).Value And (tabClear.Tab = 0)

End Sub

Private Sub SetComboText(cboCombo As ComboBox, sText As String)

    Dim lCount As Long
    
    With cboCombo
        For lCount = 1 To .ListCount
            If .List(lCount - 1) = sText Then
                .ListIndex = lCount - 1
                Exit For
            End If
        Next
    End With

End Sub

Private Sub tabClear_Click(PreviousTab As Integer)

  If mblnReadOnly Then
    Exit Sub
  End If

  Select Case tabClear.Tab
  Case 0
    fraData.Enabled = True
    fraTableColumn.Enabled = False
    fraUser.Enabled = False
    fraAccess.Enabled = False
    'NHRD24042002 Fault 3781
    'Ensures warning message is displayed when the
    '"Purge entries older than..." option is selected
    lblWarning.Visible = (optPurge(0).Value = True) And tabClear.Tab = 0
    
  Case 1
    fraData.Enabled = False
    fraTableColumn.Enabled = True
    fraUser.Enabled = False
    fraAccess.Enabled = False
    'NHRD24042002 Fault 3781
    'Don't want this message displayed
    'in any conditions for this tab
    lblWarning.Visible = (optPurge(0).Value = True) And tabClear.Tab = 0
    
  Case 2
    fraData.Enabled = False
    fraTableColumn.Enabled = False
    fraUser.Enabled = True
    fraAccess.Enabled = False
    'NHRD24042002 Fault 3781
    'Don't want this message displayed
    'in any conditions for this tab
    lblWarning.Visible = (optPurge(0).Value = True) And tabClear.Tab = 0
  Case 3
    fraData.Enabled = False
    fraTableColumn.Enabled = False
    fraUser.Enabled = False
    fraAccess.Enabled = True
    'NHRD24042002 Fault 3781
    'Don't want this message displayed
    'in any conditions for this tab
    lblWarning.Visible = (optPurge(0).Value = True) And tabClear.Tab = 0
  End Select

End Sub
