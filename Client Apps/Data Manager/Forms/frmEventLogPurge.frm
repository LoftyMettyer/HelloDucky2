VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmEventLogPurge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Event Log Purge"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1036
   Icon            =   "frmEventLogPurge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4800
      TabIndex        =   4
      Top             =   1545
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3555
      TabIndex        =   3
      Top             =   1545
      Width           =   1200
   End
   Begin VB.Frame fraPurge 
      Caption         =   "Purge Criteria :"
      Height          =   1290
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   5925
      Begin VB.ComboBox cboPeriod 
         Height          =   315
         ItemData        =   "frmEventLogPurge.frx":000C
         Left            =   4335
         List            =   "frmEventLogPurge.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   735
         Width           =   1425
      End
      Begin COASpinner.COA_Spinner spnDays 
         Height          =   300
         Left            =   3645
         TabIndex        =   2
         Top             =   735
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaximumValue    =   999
         Text            =   "0"
      End
      Begin VB.OptionButton optPurge 
         Caption         =   "Purge Event Log entries older than :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   765
         Width           =   3465
      End
      Begin VB.OptionButton optNoPurge 
         Caption         =   "Do not automatically purge the Event Log"
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   360
         Width           =   3960
      End
   End
End
Attribute VB_Name = "frmEventLogPurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()

  ' Load the purge information from the database into the controls
  Dim prstTemp As Recordset
  
  Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysEventLogPurge")
  
  If prstTemp.BOF And prstTemp.EOF Then
    optNoPurge.Value = True
  Else
    optPurge.Value = True
    spnDays.Value = prstTemp.Fields("Frequency")
    Select Case UCase(prstTemp.Fields("Period"))
      Case "DD": SetComboText cboPeriod, "Day(s)"
      Case "WK": SetComboText cboPeriod, "Week(s)"
      Case "MM": SetComboText cboPeriod, "Month(s)"
      Case "YY": SetComboText cboPeriod, "Year(s)"
    End Select
  End If
  
  Set prstTemp = Nothing
  
End Sub

Private Sub cmdOK_Click()

  If Not Validate Then
    Exit Sub
  End If
  
  SavePurgeInfo
  Unload Me
  
End Sub

Private Sub cmdCancel_Click()

  Unload Me
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optNoPurge_Click()

  Me.cboPeriod.ListIndex = -1
  Me.cboPeriod.BackColor = &H8000000F
  Me.cboPeriod.Enabled = False
  
  Me.spnDays.Value = 0
  Me.spnDays.BackColor = &H8000000F
  Me.spnDays.Enabled = False

End Sub

Private Sub optPurge_Click()

  Me.cboPeriod.ListIndex = 0
  Me.cboPeriod.Enabled = True
  Me.cboPeriod.BackColor = &H80000005
  
  Me.spnDays.Enabled = True
  Me.spnDays.BackColor = &H80000005

End Sub

Private Function Validate() As Boolean

  If optPurge.Value Then
    If cboPeriod.Text = "" Then
      COAMsgBox "You must select a period to purge event log entries.", vbExclamation + vbOKOnly, "Event Log"
      Validate = False
      Exit Function
    ElseIf (cboPeriod.ListIndex = 3) And (spnDays.Value > 200) Then
      COAMsgBox "You cannot select a purge period of greater than 200 years.", vbExclamation + vbOKOnly, "Event Log"
      Validate = False
      Exit Function
    
    End If
  End If
  
  Validate = True

End Function

Private Sub SavePurgeInfo()

  On Error GoTo ErrorTrap

  Dim pstrTriggerSQL As String
  Dim pstrTemp As String
  
  Select Case cboPeriod.Text
    Case "Day(s)": pstrTemp = "dd"
    Case "Week(s)": pstrTemp = "wk"
    Case "Month(s)": pstrTemp = "mm"
    Case "Year(s)": pstrTemp = "yy"
  End Select
  
  ' Change pointer to hourglass
  Screen.MousePointer = vbHourglass
  
  ' Delete old purge information to the database
  gADOCon.Execute "DELETE FROM AsrSysEventLogPurge"
  
    ' JPD20030206 Fault 5022
'  ' Remove the trigger (if it exists already)
'  gADOCon.Execute "DROP TRIGGER INS_AsrSysPurgeEventLog"

  ' If we are purging, then create a trigger on the INSERT event of the main
  ' Event Log table (AsrSysEventLog) which will delete entries from AsrSysEventLog
  ' and its children in AsrSysEventLogDetails that are older than the purge
  ' criteria
  If optPurge.Value = True Then
  
    gADOCon.Execute "INSERT INTO AsrSysEventLogPurge (Period,Frequency) VALUES ('" & pstrTemp & "'," & spnDays.Value & ")"
  
    ' JPD20030206 Fault 5022
'    pstrTriggerSQL = pstrTriggerSQL & "CREATE TRIGGER INS_AsrSysPurgeEventLog "
'    pstrTriggerSQL = pstrTriggerSQL & "ON AsrSysEventLog "
'    pstrTriggerSQL = pstrTriggerSQL & "FOR INSERT AS "
'
'    pstrTriggerSQL = pstrTriggerSQL & "DECLARE @intFrequency int, "
'    pstrTriggerSQL = pstrTriggerSQL & "@strPeriod char(2) "
'
'    pstrTriggerSQL = pstrTriggerSQL & "SELECT @intFrequency = Frequency "
'    pstrTriggerSQL = pstrTriggerSQL & "FROM AsrSysEventLogPurge "
'
'    pstrTriggerSQL = pstrTriggerSQL & "SELECT @strPeriod = Period "
'    pstrTriggerSQL = pstrTriggerSQL & "FROM AsrSysEventLogPurge "
'
'    pstrTriggerSQL = pstrTriggerSQL & "IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL) "
'
'    pstrTriggerSQL = pstrTriggerSQL & "BEGIN "
'
'    pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'dd' BEGIN DELETE FROM AsrSysEventLog WHERE [DateTime] < DATEADD(dd,-@intfrequency,getdate()) END "
'    pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'wk' BEGIN DELETE FROM AsrSysEventLog WHERE [DateTime] < DATEADD(wk,-@intfrequency,getdate()) END "
'    pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'mm' BEGIN DELETE FROM AsrSysEventLog WHERE [DateTime] < DATEADD(mm,-@intfrequency,getdate()) END "
'    pstrTriggerSQL = pstrTriggerSQL & "IF @strPeriod = 'yy' BEGIN DELETE FROM AsrSysEventLog WHERE [DateTime] < DATEADD(yy,-@intfrequency,getdate()) END "
'
'    pstrTriggerSQL = pstrTriggerSQL & "DELETE FROM AsrSysEventLogDetails WHERE [EventLogID] NOT IN (SELECT ID FROM AsrSysEventLog) "
'
'    pstrTriggerSQL = pstrTriggerSQL & "END"
'
'    gADOCon.Execute (pstrTriggerSQL)
  
  End If
  
  gADOCon.Execute "sp_AsrEventLogPurge"
  
  ' Change pointer back to default
  Screen.MousePointer = vbDefault
  
  If Me.optPurge.Value Then COAMsgBox "Purge completed.", vbInformation + vbOKOnly, "Event Log"
  
  Exit Sub
  
ErrorTrap:
  
  Select Case Err.Number
  
    Case -2147217865 ' Trigger didnt exist in the first place
      Resume Next
    
    Case Else
      Screen.MousePointer = vbDefault
      COAMsgBox "Error saving purge information." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Event Log"
      Exit Sub
  End Select
  
End Sub

