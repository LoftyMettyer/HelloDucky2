VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmDiaryDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purge Diary Events"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1030
   Icon            =   "frmDiaryDelete.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3930
      TabIndex        =   9
      Top             =   2055
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5175
      TabIndex        =   10
      Top             =   2055
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1830
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   3228
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&System Events"
      TabPicture(0)   =   "frmDiaryDelete.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Manual Events"
      TabPicture(1)   =   "frmDiaryDelete.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Purge Criteria :"
         Height          =   1290
         Left            =   -74865
         TabIndex        =   4
         Top             =   405
         Width           =   6030
         Begin VB.OptionButton optPurge 
            Caption         =   "Purge manual diary events older than :"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   750
            Width           =   3705
         End
         Begin VB.ComboBox cboPeriod 
            Height          =   315
            Index           =   1
            ItemData        =   "frmDiaryDelete.frx":0044
            Left            =   4620
            List            =   "frmDiaryDelete.frx":0054
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optNoPurge 
            Caption         =   "Do not automatically purge manual diary events"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   345
            Width           =   5340
         End
         Begin COASpinner.COA_Spinner spnDays 
            Height          =   300
            Index           =   1
            Left            =   3885
            TabIndex        =   7
            Top             =   720
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
      Begin VB.Frame Frame1 
         Caption         =   "Purge Criteria :"
         Height          =   1290
         Left            =   135
         TabIndex        =   12
         Top             =   405
         Width           =   6030
         Begin VB.ComboBox cboPeriod 
            Height          =   315
            Index           =   0
            ItemData        =   "frmDiaryDelete.frx":007C
            Left            =   4620
            List            =   "frmDiaryDelete.frx":008C
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optPurge 
            Caption         =   "Purge system diary events older than :"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   750
            Width           =   3660
         End
         Begin COASpinner.COA_Spinner spnDays 
            Height          =   300
            Index           =   0
            Left            =   3885
            TabIndex        =   2
            Top             =   720
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
         Begin VB.OptionButton optNoPurge 
            Caption         =   "Do not automatically purge system diary events"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   0
            Top             =   345
            Width           =   5340
         End
      End
   End
End
Attribute VB_Name = "frmDiaryDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()

  Dim strSQL As String
  Dim strKey As String
  Dim intCount As Integer

  On Error GoTo LocalErr

  For intCount = 0 To 1
  
    strKey = IIf(intCount = 0, "'DIARYSYS'", "'DIARYMAN'")
    
    If optPurge(intCount) = True Then
    
      'Validate new entry
      If Trim(cboPeriod(intCount).Text) = vbNullString Then
        Screen.MousePointer = vbDefault
        COAMsgBox "Please select a valid unit of time", vbExclamation, Me.Caption
        Exit Sub
        
      'ElseIf spnDays(intCount).Value < 1 Then
      ElseIf spnDays(intCount).Value < 0 Then
        Screen.MousePointer = vbDefault
        COAMsgBox "Please select a valid number of " & LCase(cboPeriod(intCount).Text), vbExclamation, Me.Caption
        Exit Sub
      
      ElseIf (cboPeriod(intCount).ListIndex = 3) And (spnDays(intCount).Value > 200) Then
        Screen.MousePointer = vbDefault
        COAMsgBox "You cannot select a purge period of greater than 200 years.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      
      End If

      Screen.MousePointer = vbHourglass
  
      'Delete old entry
      strSQL = "DELETE FROM ASRSYSPurgePeriods WHERE PurgeKey = " & strKey
      datGeneral.ExecuteSql strSQL, ""
  
      'Insert new entry
      strSQL = "INSERT ASRSYSPurgePeriods " & _
                          "(PurgeKey, Period, Unit) " & vbCrLf & _
               "VALUES(" & strKey & "," & _
                           CStr(spnDays(intCount).Value) & ",'" & _
                           Left(cboPeriod(intCount).Text, 1) & "')"
      datGeneral.ExecuteSql strSQL, ""
  
    Else
      
      'Delete old entry
      strSQL = "DELETE FROM ASRSYSPurgePeriods WHERE PurgeKey = " & strKey
      datGeneral.ExecuteSql strSQL, ""
    
    End If

  Next

  datGeneral.ExecuteSql "EXEC sp_ASRDiaryPurge", ""
  Screen.MousePointer = vbDefault

  If optPurge(0) = True Or optPurge(1) = True Then
    COAMsgBox "Diary purge completed.", vbInformation + vbOKOnly, "Diary Delete"
  End If

  Unload Me

Exit Sub

LocalErr:
  COAMsgBox "Error saving purge period", vbCritical, Me.Caption

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyEscape Then
    cmdCancel_Click
  End If
  
End Sub

Private Sub Form_Load()
  Call GetPurgeDetails
End Sub


Private Sub GetPurgeDetails()

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strKey As String
  Dim intCount As Integer

  For intCount = 0 To 1
  
    strKey = IIf(intCount = 0, "'DIARYSYS'", "'DIARYMAN'")
  
    strSQL = "SELECT * FROM ASRSYSPurgePeriods WHERE PurgeKey = " & strKey
    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

    If rsTemp.BOF And rsTemp.EOF Then
      optNoPurge(intCount) = True
      spnDays(intCount).Value = 0
      cboPeriod(intCount).ListIndex = 0
    Else
      optPurge(intCount) = True
      spnDays(intCount).Value = rsTemp!Period
      cboPeriod(intCount).ListIndex = InStr("DWMY", UCase(rsTemp!Unit)) - 1
    End If

  Next

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optNoPurge_Click(Index As Integer)

  cboPeriod(Index).Enabled = False
  cboPeriod(Index).BackColor = vbButtonFace

  spnDays(Index).Enabled = False
  spnDays(Index).BackColor = vbButtonFace

End Sub

Private Sub optPurge_Click(Index As Integer)

  cboPeriod(Index).Enabled = True
  cboPeriod(Index).BackColor = vbWindowBackground

  spnDays(Index).Enabled = True
  spnDays(Index).BackColor = vbWindowBackground

End Sub

