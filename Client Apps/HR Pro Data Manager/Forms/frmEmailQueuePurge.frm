VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmEmailQueuePurge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Queue Purge"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1033
   Icon            =   "frmEmailQueuePurge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5040
      TabIndex        =   4
      Top             =   1500
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3795
      TabIndex        =   3
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Frame fraPurge 
      Caption         =   "Purge Criteria :"
      Height          =   1290
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   6165
      Begin VB.ComboBox cboPeriod 
         Height          =   315
         ItemData        =   "frmEmailQueuePurge.frx":000C
         Left            =   4665
         List            =   "frmEmailQueuePurge.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   735
         Width           =   1425
      End
      Begin COASpinner.COA_Spinner spnDays 
         Height          =   300
         Left            =   3885
         TabIndex        =   2
         Top             =   735
         Width           =   705
         _ExtentX        =   1244
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
         Caption         =   "Purge Email Queue entries older than :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   765
         Width           =   3750
      End
      Begin VB.OptionButton optNoPurge 
         Caption         =   "Do not automatically purge the Email Queue"
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   360
         Width           =   4440
      End
   End
End
Attribute VB_Name = "frmEmailQueuePurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub GetPurgeDetails()

  Dim rsTemp As Recordset
  Dim strSQL As String

  
  strSQL = "SELECT * FROM ASRSYSPurgePeriods WHERE PurgeKey = 'EMAIL'"
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  'MH20080428 Fault 12640
  'If rsTemp.BOF And rsTemp.EOF Then
  If IsNull(rsTemp!Period) Then
    optNoPurge = True
    spnDays.Value = 0
    cboPeriod.ListIndex = 0
  Else
    optPurge = True
    spnDays.Value = rsTemp!Period
    cboPeriod.ListIndex = InStr("DWMY", UCase(rsTemp!Unit)) - 1
  
  End If

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()

  Dim rsTemp As Recordset
  Dim strSQL As String

  On Error GoTo LocalErr

  If optPurge = True Then
    
    'Validate new entry
    If Trim(cboPeriod.Text) = vbNullString Then
      MsgBox "Please select a valid unit of time", vbExclamation, Me.Caption
      Exit Sub
    'ElseIf spnDays.Value < 1 Then
    ElseIf spnDays.Value < 0 Then
      MsgBox "Please select a valid number of " & LCase(cboPeriod.Text), vbExclamation, Me.Caption
      Exit Sub
      
    ElseIf (cboPeriod.ListIndex = 3) And (spnDays.Value > 200) Then
      MsgBox "You cannot select a purge period of greater than 200 years.", vbExclamation + vbOKOnly, Me.Caption
      Exit Sub

    End If

    Screen.MousePointer = vbHourglass

    
  'MH20080428 Fault 12640
'    'Delete old entry
'    strSQL = "DELETE FROM ASRSYSPurgePeriods WHERE PurgeKey = 'EMAIL'"
'    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
'
'    'Insert new entry
'    strSQL = "INSERT ASRSYSPurgePeriods (PurgeKey, Period, Unit) " & vbCrLf & _
'             "VALUES('EMAIL'," & CStr(spnDays.Value) & ",'" & Left(cboPeriod.Text, 1) & "')"
'    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

    'Update entry
    strSQL = "UPDATE ASRSYSPurgePeriods " & _
             "SET Period = " & CStr(spnDays.Value) & ", Unit = '" & Left(cboPeriod.Text, 1) & "' " & _
             "WHERE PurgeKey = 'EMAIL'"
    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)



    'Purge queue now !
    strSQL = "EXEC sp_ASRPurgeRecords 'EMAIL', 'ASRSysEmailQueue', 'DateDue'"
    datGeneral.ExecuteSql strSQL, ""
  
  Else
    
  'MH20080428 Fault 12640
    'Delete old entry
    'strSQL = "DELETE FROM ASRSYSPurgePeriods WHERE PurgeKey = 'EMAIL'"
    'Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
    'Update entry
    strSQL = "UPDATE ASRSYSPurgePeriods " & _
             "SET Period = null, Unit = null " & _
             "WHERE PurgeKey = 'EMAIL'"
    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  
  End If

  Screen.MousePointer = vbDefault
  Unload Me

Exit Sub

LocalErr:
  MsgBox "Error saving purge period", vbCritical, Me.Caption

End Sub

Private Sub Form_Load()

  Call GetPurgeDetails

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub optNoPurge_Click()

  cboPeriod.Enabled = False
  cboPeriod.BackColor = vbButtonFace

  spnDays.Enabled = False
  spnDays.BackColor = vbButtonFace

End Sub

Private Sub optPurge_Click()

  cboPeriod.Enabled = True
  cboPeriod.BackColor = vbWindowBackground

  spnDays.Enabled = True
  spnDays.BackColor = vbWindowBackground

End Sub

