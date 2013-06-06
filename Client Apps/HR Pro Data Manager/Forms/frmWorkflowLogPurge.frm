VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmWorkflowLogPurge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Log Purge"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1141
   Icon            =   "frmWorkflowLogPurge.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPurge 
      Caption         =   "Purge Criteria :"
      Height          =   1290
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   6390
      Begin VB.OptionButton optNoPurge 
         Caption         =   "Do not automatically purge the Workflow Log"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   4275
      End
      Begin VB.OptionButton optPurge 
         Caption         =   "Purge Workflow Log entries older than :"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   765
         Width           =   3780
      End
      Begin VB.ComboBox cboPeriod 
         Height          =   315
         ItemData        =   "frmWorkflowLogPurge.frx":000C
         Left            =   4785
         List            =   "frmWorkflowLogPurge.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   735
         Width           =   1425
      End
      Begin COASpinner.COA_Spinner spnDays 
         Height          =   300
         Left            =   3990
         TabIndex        =   3
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
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4020
      TabIndex        =   5
      Top             =   1550
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5280
      TabIndex        =   6
      Top             =   1550
      Width           =   1200
   End
End
Attribute VB_Name = "frmWorkflowLogPurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SavePurgeInfo()

  On Error GoTo ErrorTrap

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String
  Dim pstrTriggerSQL As String
  
  ' Change pointer to hourglass
  Screen.MousePointer = vbHourglass
  
  
  'MH20080428 Fault 12640
'  ' Delete old purge information to the database
'  gADOCon.Execute "DELETE FROM ASRSYSPurgePeriods WHERE PurgeKey = 'WORKFLOW'"
'
'  If optPurge.Value = True Then
'    gADOCon.Execute "INSERT INTO ASRSYSPurgePeriods (PurgeKey, Period, Unit) VALUES ('WORKFLOW'," & CStr(spnDays.Value) & ",'" & Left(cboPeriod.Text, 1) & "')"
'  End If
  
  If optPurge.Value = True Then
    strSQL = "SET Period = " & CStr(spnDays.Value) & ", Unit = '" & Left(cboPeriod.Text, 1) & "'"
  Else
    strSQL = "SET Period = null, Unit = null"
  End If
  
  strSQL = "UPDATE ASRSYSPurgePeriods " & strSQL & " WHERE PurgeKey = 'WORKFLOW'"
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  
  gADOCon.Execute "spASRWorkflowLogPurge"
  
  ' Change pointer back to default
  Screen.MousePointer = vbDefault
  
  If Me.optPurge.Value Then COAMsgBox "Purge completed.", vbInformation + vbOKOnly, "Workflow Log"
  
  Exit Sub
  
ErrorTrap:
  
  Select Case Err.Number
  
    Case -2147217865 ' Trigger didnt exist in the first place
      Resume Next
    
    Case Else
      Screen.MousePointer = vbDefault
      COAMsgBox "Error saving purge information." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Workflow Log"
      Exit Sub
  End Select
  
End Sub

Private Function Validate() As Boolean

  If optPurge.Value Then
    If cboPeriod.Text = "" Then
      COAMsgBox "You must select a period to purge workflow log entries.", vbExclamation + vbOKOnly, "Workflow Log"
      Validate = False
      Exit Function
    ElseIf (cboPeriod.ListIndex = 3) And (spnDays.Value > 200) Then
      COAMsgBox "You cannot select a purge period of greater than 200 years.", vbExclamation + vbOKOnly, "Workflow Log"
      Validate = False
      Exit Function
    
    End If
  End If
  
  Validate = True

End Function


Private Sub cmdCancel_Click()
  Unload Me

End Sub


Private Sub cmdOK_Click()
  If Not Validate Then
    Exit Sub
  End If
  
  SavePurgeInfo
  Unload Me

End Sub


Private Sub Form_Load()
  ' Load the purge information from the database into the controls
  Dim prstTemp As Recordset
  
  Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM ASRSYSPurgePeriods WHERE PurgeKey = 'WORKFLOW'")

  'MH20080428 Fault 12640
  'If prstTemp.BOF And prstTemp.EOF Then
  If IsNull(prstTemp.Fields("Period")) Then
    optNoPurge.Value = True
  Else
    optPurge.Value = True
    spnDays.Value = prstTemp.Fields("Period")

    Select Case UCase(prstTemp.Fields("Unit"))
      Case "D": SetComboText cboPeriod, "Day(s)"
      Case "W": SetComboText cboPeriod, "Week(s)"
      Case "M": SetComboText cboPeriod, "Month(s)"
      Case "Y": SetComboText cboPeriod, "Year(s)"
    End Select
 End If

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



