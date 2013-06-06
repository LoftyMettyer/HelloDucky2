VERSION 5.00
Begin VB.Form frmDiaryAlarmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diary Alarms"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1029
   Icon            =   "frmDiaryAlarmSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAlarm 
      Caption         =   "Alarm all events, which you have access to, within current view"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
End
Attribute VB_Name = "frmDiaryAlarmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnUpdate As Boolean

Private Sub chkAlarm_Click()
  mblnUpdate = True
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()

  Dim strSQL As String
  Dim strSetAlarm As String

  If mblnUpdate = True Then  'Check box has been clicked
    strSetAlarm = CStr(chkAlarm.Value)
    Call gobjDiary.SwitchAlarms(strSetAlarm)
  End If
  
  Unload Me

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

  Dim stlSQL As String
  Dim intAlarmedCount As Integer
  Dim intNonAlarmedCount As Integer
  
  intAlarmedCount = gobjDiary.GetAlarmCount(1)
  intNonAlarmedCount = gobjDiary.GetAlarmCount(0)

  If intAlarmedCount > 0 Then
    If intNonAlarmedCount > 0 Then
      chkAlarm = vbGrayed
    Else
      chkAlarm = vbChecked
    End If
  Else
    chkAlarm = vbUnchecked
  End If

  mblnUpdate = False

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set frmDiaryAlarmSet = Nothing
End Sub

